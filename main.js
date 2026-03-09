const PORTFOLIOS_URL = "./data/portfolios.json";
const LIVE_PRICES_URL = "./data/live-prices.json";
const REFRESH_INTERVAL_MS = 5 * 60 * 1000;

const currencyFormatter = new Intl.NumberFormat("ko-KR");
const decimalFormatter = new Intl.NumberFormat("ko-KR", {
  minimumFractionDigits: 0,
  maximumFractionDigits: 1,
});
const percentFormatter = new Intl.NumberFormat("ko-KR", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});
const dateTimeFormatter = new Intl.DateTimeFormat("ko-KR", {
  dateStyle: "medium",
  timeStyle: "short",
  timeZone: "Asia/Seoul",
});

const state = {
  portfolioSource: null,
  livePriceSource: null,
  isRefreshing: false,
  refreshError: null,
  refreshTimerId: null,
};

function formatCurrency(value) {
  return `${currencyFormatter.format(Math.round(value))}원`;
}

function formatSignedCurrency(value) {
  const rounded = Math.round(value);
  const prefix = rounded > 0 ? "+" : "";
  return `${prefix}${currencyFormatter.format(rounded)}원`;
}

function formatPrice(value) {
  return `${currencyFormatter.format(Math.round(value))}원`;
}

function formatAverageCost(value) {
  const isWhole = Math.abs(value - Math.round(value)) < 0.05;
  const displayValue = isWhole ? Math.round(value) : Number(value.toFixed(1));
  return `${decimalFormatter.format(displayValue)}원`;
}

function formatPercent(value) {
  return `${percentFormatter.format(value * 100)}%`;
}

function formatSignedPercent(value) {
  const prefix = value > 0 ? "+" : "";
  return `${prefix}${percentFormatter.format(value * 100)}%`;
}

function formatUpdatedAt(value) {
  if (!value) {
    return "업데이트 시각 없음";
  }

  return dateTimeFormatter.format(new Date(value));
}

function getToneClass(value) {
  if (value > 0) {
    return "positive";
  }

  if (value < 0) {
    return "negative";
  }

  return "";
}

async function fetchJson(url, options = {}) {
  const requestUrl = new URL(url, window.location.href);
  requestUrl.searchParams.set("ts", `${Date.now()}`);

  const response = await fetch(requestUrl, {
    cache: "no-store",
    ...options,
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch ${url}: ${response.status}`);
  }

  return response.json();
}

function buildPortfolioData(source, livePriceSource) {
  const quotes = livePriceSource?.quotes ?? {};
  const hasExternalQuotes =
    Boolean(livePriceSource) && livePriceSource.provider !== "snapshot-fallback";

  const enrichedHoldings = source.holdings.map((holding) => {
    const liveQuote =
      hasExternalQuotes && !holding.manualPriceOnly ? quotes[holding.ticker] : null;
    const currentPrice = Number.isFinite(liveQuote?.price)
      ? liveQuote.price
      : holding.snapshotPrice;
    const marketValue = currentPrice * holding.quantity;
    const profitLoss = marketValue - holding.costBasis;
    const profitRate = holding.costBasis === 0 ? 0 : profitLoss / holding.costBasis;

    return {
      ...holding,
      currentPrice,
      marketValue,
      profitLoss,
      profitRate,
      averageCost: holding.costBasis / holding.quantity,
      priceSource: holding.manualPriceOnly
        ? "manual"
        : liveQuote
          ? "live"
          : "snapshot",
    };
  });

  const totals = enrichedHoldings.reduce(
    (accumulator, holding) => {
      accumulator.quantity += holding.quantity;
      accumulator.marketValue += holding.marketValue;
      accumulator.costBasis += holding.costBasis;
      accumulator.profitLoss += holding.profitLoss;
      return accumulator;
    },
    {
      quantity: 0,
      marketValue: 0,
      costBasis: 0,
      profitLoss: 0,
    },
  );

  const totalAssets = totals.marketValue + source.cash;
  const investedWeight = totalAssets === 0 ? 0 : totals.marketValue / totalAssets;
  const cashWeight = totalAssets === 0 ? 0 : source.cash / totalAssets;
  const totalReturn = totals.costBasis === 0 ? 0 : totals.profitLoss / totals.costBasis;

  const holdingsWithWeight = enrichedHoldings.map((holding) => ({
    ...holding,
    assetWeight: totalAssets === 0 ? 0 : holding.marketValue / totalAssets,
  }));

  return {
    ...source,
    holdings: holdingsWithWeight,
    totals: {
      ...totals,
      totalAssets,
      investedWeight,
      cashWeight,
      totalReturn,
    },
  };
}

function buildBookData(source, livePriceSource) {
  const portfolios = source.accounts.map((account) =>
    buildPortfolioData(account, livePriceSource)
  );

  const overview = portfolios.reduce(
    (accumulator, portfolio) => {
      accumulator.totalAssets += portfolio.totals.totalAssets;
      accumulator.marketValue += portfolio.totals.marketValue;
      accumulator.costBasis += portfolio.totals.costBasis;
      accumulator.profitLoss += portfolio.totals.profitLoss;
      accumulator.cash += portfolio.cash;
      accumulator.holdingsCount += portfolio.holdings.length;
      return accumulator;
    },
    {
      totalAssets: 0,
      marketValue: 0,
      costBasis: 0,
      profitLoss: 0,
      cash: 0,
      holdingsCount: 0,
    },
  );

  overview.totalReturn =
    overview.costBasis === 0 ? 0 : overview.profitLoss / overview.costBasis;

  return {
    meta: source.meta,
    livePriceSource,
    portfolios,
    overview,
  };
}

function buildMarketStatus(book) {
  if (state.isRefreshing) {
    return {
      tone: "status-loading",
      title: "외부 시세 다시 불러오는 중",
      detail: "최신 가격 파일을 확인하고 있습니다.",
    };
  }

  if (state.refreshError) {
    return {
      tone: "status-stale",
      title: "자동 시세 확인 실패",
      detail: "기존 저장 가격으로 계속 계산합니다.",
    };
  }

  if (book.livePriceSource?.provider === "KIS Open API") {
    return {
      tone: "status-live",
      title: "외부 시세 자동 갱신 사용 중",
      detail: `${book.livePriceSource.provider} · 최근 반영 ${formatUpdatedAt(book.livePriceSource.updatedAt)}`,
    };
  }

  return {
    tone: "status-fallback",
    title: "스냅샷 가격으로 계산 중",
    detail: `마지막 저장 시각 ${formatUpdatedAt(book.livePriceSource?.updatedAt)}`,
  };
}

function renderSummaryCards(portfolio) {
  const cards = [
    {
      label: "총자산",
      value: formatCurrency(portfolio.totals.totalAssets),
      detail: `주식 ${formatCurrency(portfolio.totals.marketValue)} + 현금 ${formatCurrency(portfolio.cash)}`,
      tone: "",
    },
    {
      label: "평가손익",
      value: formatSignedCurrency(portfolio.totals.profitLoss),
      detail: `보유 종목 기준 ${formatSignedPercent(portfolio.totals.totalReturn)}`,
      tone: getToneClass(portfolio.totals.profitLoss),
    },
    {
      label: "현금",
      value: formatCurrency(portfolio.cash),
      detail: `총자산 비중 ${formatPercent(portfolio.totals.cashWeight)}`,
      tone: "",
    },
    {
      label: "투자 비중",
      value: formatPercent(portfolio.totals.investedWeight),
      detail: `${portfolio.holdings.length}개 종목 보유`,
      tone: "",
    },
  ];

  return cards
    .map(
      (card) => `
        <article class="summary-card">
          <p class="summary-label">${card.label}</p>
          <p class="summary-value ${card.tone}">${card.value}</p>
          <p class="summary-detail">${card.detail}</p>
        </article>
      `,
    )
    .join("");
}

function renderBookSummaryCards(book) {
  const cards = [
    {
      label: "전체 총자산",
      value: formatCurrency(book.overview.totalAssets),
      detail: `주식 ${formatCurrency(book.overview.marketValue)} + 현금 ${formatCurrency(book.overview.cash)}`,
      tone: "",
    },
    {
      label: "전체 평가손익",
      value: formatSignedCurrency(book.overview.profitLoss),
      detail: `전체 계좌 기준 ${formatSignedPercent(book.overview.totalReturn)}`,
      tone: getToneClass(book.overview.profitLoss),
    },
    {
      label: "전체 현금",
      value: formatCurrency(book.overview.cash),
      detail: `${book.portfolios.length}개 계좌 운영`,
      tone: "",
    },
    {
      label: "총 보유 종목",
      value: `${currencyFormatter.format(book.overview.holdingsCount)}개`,
      detail: "현재가만 자동 갱신, 수량/평단은 원장 기준",
      tone: "",
    },
  ];

  return cards
    .map(
      (card) => `
        <article class="summary-card">
          <p class="summary-label">${card.label}</p>
          <p class="summary-value ${card.tone}">${card.value}</p>
          <p class="summary-detail">${card.detail}</p>
        </article>
      `,
    )
    .join("");
}

function getTopHoldings(portfolio, limit = 3) {
  return [...portfolio.holdings]
    .sort((left, right) => right.marketValue - left.marketValue)
    .slice(0, limit);
}

function renderAccountOverviewCards(book) {
  return book.portfolios
    .map((portfolio) => {
      const topHoldings = getTopHoldings(portfolio);

      return `
        <article class="overview-card">
          <div class="overview-card-head">
            <div>
              <p class="overview-card-label">${portfolio.name}</p>
              <h3 class="overview-card-total">${formatCurrency(portfolio.totals.totalAssets)}</h3>
            </div>
            <a class="overview-card-link" href="#${portfolio.id}">상세 보기</a>
          </div>

          <div class="overview-metric-grid">
            <div class="overview-metric">
              <span>평가손익</span>
              <strong class="${getToneClass(portfolio.totals.profitLoss)}">${formatSignedCurrency(portfolio.totals.profitLoss)}</strong>
            </div>
            <div class="overview-metric">
              <span>수익률</span>
              <strong class="${getToneClass(portfolio.totals.totalReturn)}">${formatSignedPercent(portfolio.totals.totalReturn)}</strong>
            </div>
            <div class="overview-metric">
              <span>현금</span>
              <strong>${formatCurrency(portfolio.cash)}</strong>
            </div>
            <div class="overview-metric">
              <span>보유 종목</span>
              <strong>${currencyFormatter.format(portfolio.holdings.length)}개</strong>
            </div>
          </div>

          <div class="overview-top-list">
            ${topHoldings
              .map(
                (holding) => `
                  <div class="overview-top-item">
                    <span class="overview-top-name">${holding.name}</span>
                    <span class="overview-top-meta">${formatPercent(holding.assetWeight)} · ${formatCurrency(holding.marketValue)}</span>
                  </div>
                `,
              )
              .join("")}
          </div>
        </article>
      `;
    })
    .join("");
}

function renderAccountComparisonRows(book) {
  return book.portfolios
    .map((portfolio) => {
      const primaryHolding = getTopHoldings(portfolio, 1)[0];

      return `
        <tr>
          <td>
            <a class="comparison-account" href="#${portfolio.id}">
              <strong>${portfolio.name}</strong>
              <span>${portfolio.snapshotLabel}</span>
            </a>
          </td>
          <td>${formatCurrency(portfolio.totals.totalAssets)}</td>
          <td class="${getToneClass(portfolio.totals.profitLoss)}">${formatSignedCurrency(portfolio.totals.profitLoss)}</td>
          <td class="${getToneClass(portfolio.totals.totalReturn)}">${formatSignedPercent(portfolio.totals.totalReturn)}</td>
          <td>${formatCurrency(portfolio.cash)}</td>
          <td>${formatPercent(portfolio.totals.investedWeight)}</td>
          <td>
            ${primaryHolding
              ? `
                <div class="comparison-primary">
                  <strong>${primaryHolding.name}</strong>
                  <span>총자산 ${formatPercent(primaryHolding.assetWeight)}</span>
                </div>
              `
              : "-"}
          </td>
        </tr>
      `;
    })
    .join("");
}

function renderTableRows(portfolio) {
  return portfolio.holdings
    .map(
      (holding) => `
        <tr>
          <td>
            <div class="asset-name">${holding.name}</div>
            <div class="asset-meta">
              <span class="tag">${holding.type}</span>
              <span class="tag ${holding.priceSource === "live" ? "tag-live" : "tag-warning"}">
                ${holding.priceSource === "live"
                  ? "자동 시세"
                  : holding.priceSource === "manual"
                    ? "수동 가격"
                    : "스냅샷"}
              </span>
              ${holding.nameInferred ? '<span class="tag tag-warning">이름 추정</span>' : ""}
            </div>
          </td>
          <td>${currencyFormatter.format(holding.quantity)}주</td>
          <td>
            <div class="metric">
              <strong>${formatPrice(holding.currentPrice)}</strong>
              <span>${holding.priceSource === "live"
                ? "외부 시세 반영"
                : holding.priceSource === "manual"
                  ? "수동 기준 가격"
                  : "스크린샷 기준"}</span>
            </div>
          </td>
          <td>${formatAverageCost(holding.averageCost)}</td>
          <td>${formatCurrency(holding.costBasis)}</td>
          <td>${formatCurrency(holding.marketValue)}</td>
          <td class="${getToneClass(holding.profitLoss)}">${formatSignedCurrency(holding.profitLoss)}</td>
          <td class="${getToneClass(holding.profitRate)}">${formatSignedPercent(holding.profitRate)}</td>
          <td>${formatPercent(holding.assetWeight)}</td>
        </tr>
      `,
    )
    .join("");
}

function renderFocusCards(portfolio) {
  return getTopHoldings(portfolio)
    .map(
      (holding, index) => `
        <article class="focus-card">
          <p class="focus-label">상위 보유 ${index + 1}</p>
          <h3 class="focus-name">${holding.name}</h3>
          <p class="focus-meta">${formatCurrency(holding.marketValue)} · 총자산 ${formatPercent(holding.assetWeight)}</p>
          <p class="focus-change ${getToneClass(holding.profitLoss)}">${formatSignedCurrency(holding.profitLoss)} / ${formatSignedPercent(holding.profitRate)}</p>
        </article>
      `,
    )
    .join("");
}

function renderPortfolioSection(portfolio) {
  return `
    <section class="account-stack" id="${portfolio.id}">
      <section class="hero">
        <div class="eyebrow">Account Detail</div>
        <div class="hero-head">
          <div>
            <h2 class="hero-title">${portfolio.name}</h2>
            <p class="hero-copy">
              ${portfolio.snapshotLabel} 기준 원장 데이터를 유지하고, 현재가만 외부 시세 파일로 다시 읽어
              평가금액과 손익을 재계산합니다.
            </p>
          </div>
          <div class="account-pill">총자산 ${formatCurrency(portfolio.totals.totalAssets)}</div>
        </div>
        <div class="summary-grid">
          ${renderSummaryCards(portfolio)}
        </div>
        <div class="focus-strip">
          ${renderFocusCards(portfolio)}
        </div>
      </section>

      <section class="table-card">
        <div class="section-head">
          <div>
            <h3 class="section-title">보유 종목 표</h3>
            <p class="section-copy">
              수량과 매입금액은 고정 원장 데이터입니다. 현재가, 평가금액, 손익, 수익률은 최신 시세 기준으로 다시 계산됩니다.
            </p>
          </div>
          <div class="section-pill">총 ${portfolio.holdings.length}종목</div>
        </div>

        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>종목</th>
                <th>수량</th>
                <th>현재가</th>
                <th>평균단가</th>
                <th>매입금액</th>
                <th>평가금액</th>
                <th>평가손익</th>
                <th>수익률</th>
                <th>총자산 비중</th>
              </tr>
            </thead>
            <tbody>
              ${renderTableRows(portfolio)}
            </tbody>
            <tfoot>
              <tr>
                <td>합계 (주식)</td>
                <td>${currencyFormatter.format(portfolio.totals.quantity)}주</td>
                <td>-</td>
                <td>-</td>
                <td>${formatCurrency(portfolio.totals.costBasis)}</td>
                <td>${formatCurrency(portfolio.totals.marketValue)}</td>
                <td class="${getToneClass(portfolio.totals.profitLoss)}">${formatSignedCurrency(portfolio.totals.profitLoss)}</td>
                <td class="${getToneClass(portfolio.totals.totalReturn)}">${formatSignedPercent(portfolio.totals.totalReturn)}</td>
                <td>${formatPercent(portfolio.totals.investedWeight)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </section>
    </section>
  `;
}

function renderLoading() {
  const app = document.querySelector("#app");

  if (!app) {
    return;
  }

  app.innerHTML = `
    <section class="state-card">
      <p class="state-eyebrow">Loading</p>
      <h1 class="state-title">포트폴리오 데이터를 불러오는 중입니다.</h1>
      <p class="state-copy">계좌 원장과 최신 현재가 파일을 확인하고 있습니다.</p>
    </section>
  `;
}

function renderError(error) {
  const app = document.querySelector("#app");

  if (!app) {
    return;
  }

  app.innerHTML = `
    <section class="state-card">
      <p class="state-eyebrow">Error</p>
      <h1 class="state-title">포트폴리오 데이터를 불러오지 못했습니다.</h1>
      <p class="state-copy">${error.message}</p>
      <button class="ghost-button" data-refresh-quotes>다시 시도</button>
    </section>
  `;
}

function renderApp(book) {
  const app = document.querySelector("#app");

  if (!app) {
    return;
  }

  const marketStatus = buildMarketStatus(book);

  app.innerHTML = `
    <section class="dashboard">
      <section class="hero">
        <div class="eyebrow">Tax-Sheltered Accounts</div>
        <div class="hero-head">
          <div>
            <h1 class="hero-title">절세계좌 포트폴리오</h1>
            <p class="hero-copy">
              연금저축, 퇴직연금, ISA를 같은 기준으로 묶어 총자산, 손익, 현금, 핵심 보유 종목을 한 화면에서 점검할 수 있게 구성했습니다.
              현재가 자동 갱신이 없으면 마지막 스냅샷 가격으로 그대로 계산합니다.
            </p>
          </div>
          <div class="account-pill">기준 통화 ${book.meta.currency}</div>
        </div>

        <div class="hero-toolbar">
          <div class="market-status ${marketStatus.tone}">
            <span class="status-dot"></span>
            <div>
              <strong>${marketStatus.title}</strong>
              <span>${marketStatus.detail}</span>
            </div>
          </div>
          <button class="ghost-button" data-refresh-quotes ${state.isRefreshing ? "disabled" : ""}>
            ${state.isRefreshing ? "업데이트 중..." : "시세 다시 불러오기"}
          </button>
        </div>

        <div class="summary-grid">
          ${renderBookSummaryCards(book)}
        </div>
      </section>

      <section class="overview-section">
        <div class="section-head overview-section-head">
          <div>
            <h2 class="section-title">계좌 한눈에 보기</h2>
            <p class="section-copy">
              각 계좌의 총자산, 손익, 현금, 핵심 보유 종목을 먼저 비교하고 필요한 계좌만 아래 상세 표로 내려가면 됩니다.
            </p>
          </div>
        </div>

        <div class="overview-layout">
          <div class="overview-grid">
            ${renderAccountOverviewCards(book)}
          </div>

          <section class="comparison-card">
            <div class="section-head comparison-head">
              <div>
                <h3 class="section-title">계좌 비교표</h3>
                <p class="section-copy">총자산, 평가손익, 수익률, 투자 비중을 같은 행으로 비교합니다.</p>
              </div>
            </div>

            <div class="comparison-table-wrap">
              <table class="comparison-table">
                <thead>
                  <tr>
                    <th>계좌</th>
                    <th>총자산</th>
                    <th>평가손익</th>
                    <th>수익률</th>
                    <th>현금</th>
                    <th>주식 비중</th>
                    <th>상위 보유</th>
                  </tr>
                </thead>
                <tbody>
                  ${renderAccountComparisonRows(book)}
                </tbody>
              </table>
            </div>
          </section>
        </div>
      </section>

      <section class="detail-section">
        <div class="section-head detail-head">
          <div>
            <h2 class="section-title">계좌 상세</h2>
            <p class="section-copy">
              각 계좌의 원장 기준 수량과 평단, 그리고 최신 현재가 기준 평가금액과 손익을 상세 표로 확인합니다.
            </p>
          </div>
        </div>

        <div class="account-stack-list">
          ${book.portfolios.map(renderPortfolioSection).join("")}
        </div>
      </section>
    </section>
  `;
}

async function refreshLivePrices() {
  if (!state.portfolioSource || state.isRefreshing) {
    return;
  }

  state.isRefreshing = true;
  state.refreshError = null;
  renderApp(buildBookData(state.portfolioSource, state.livePriceSource));

  try {
    state.livePriceSource = await fetchJson(LIVE_PRICES_URL);
  } catch (error) {
    state.refreshError = error;
  } finally {
    state.isRefreshing = false;
    renderApp(buildBookData(state.portfolioSource, state.livePriceSource));
  }
}

function startAutoRefresh() {
  if (state.refreshTimerId) {
    window.clearInterval(state.refreshTimerId);
  }

  state.refreshTimerId = window.setInterval(() => {
    refreshLivePrices();
  }, REFRESH_INTERVAL_MS);
}

async function initializeApp() {
  renderLoading();

  try {
    state.portfolioSource = await fetchJson(PORTFOLIOS_URL);

    try {
      state.livePriceSource = await fetchJson(LIVE_PRICES_URL);
    } catch (error) {
      state.livePriceSource = null;
      state.refreshError = error;
    }

    renderApp(buildBookData(state.portfolioSource, state.livePriceSource));
    startAutoRefresh();
  } catch (error) {
    renderError(error);
  }
}

document.addEventListener("click", (event) => {
  const refreshButton = event.target.closest("[data-refresh-quotes]");

  if (!refreshButton) {
    return;
  }

  if (!state.portfolioSource) {
    initializeApp();
    return;
  }

  refreshLivePrices();
});

initializeApp();
