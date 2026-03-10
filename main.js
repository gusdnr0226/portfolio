const PORTFOLIOS_URL = "./data/portfolios.json";
const LIVE_PRICES_URL = "./data/live-prices.json";
const GITHUB_API_BASE_URL = "https://api.github.com/repos/gusdnr0226/portfolio";
const GITHUB_LIVE_PRICES_URL =
  `${GITHUB_API_BASE_URL}/contents/data/live-prices.json?ref=main`;
const GITHUB_ACTIONS_WORKFLOW_ID = "update-live-prices.yml";
const GITHUB_DEFAULT_BRANCH = "main";
const GITHUB_TOKEN_STORAGE_KEY = "portfolio.refresh.githubToken";
const REFRESH_INTERVAL_MS = 5 * 60 * 1000;
const WORKFLOW_POLL_INTERVAL_MS = 2500;
const WORKFLOW_DISCOVERY_TIMEOUT_MS = 30 * 1000;
const WORKFLOW_COMPLETION_TIMEOUT_MS = 120 * 1000;
const LIVE_PRICE_VISIBILITY_TIMEOUT_MS = 20 * 1000;

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
  refreshPhase: null,
  refreshDetail: null,
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

function getErrorMessage(error) {
  if (error instanceof Error && error.message) {
    return error.message;
  }

  return String(error);
}

function delay(ms) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, ms);
  });
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

function getStoredGitHubToken() {
  try {
    const value = window.localStorage.getItem(GITHUB_TOKEN_STORAGE_KEY);
    return value?.trim() || null;
  } catch {
    return null;
  }
}

function setStoredGitHubToken(token) {
  try {
    window.localStorage.setItem(GITHUB_TOKEN_STORAGE_KEY, token);
  } catch {
    // Ignore storage failures and continue with the active token only.
  }
}

function clearStoredGitHubToken() {
  try {
    window.localStorage.removeItem(GITHUB_TOKEN_STORAGE_KEY);
  } catch {
    // Ignore storage failures.
  }
}

function hasStoredGitHubToken() {
  return Boolean(getStoredGitHubToken());
}

function isAutomaticLivePriceSource(livePriceSource) {
  return livePriceSource?.provider === "KIS Open API";
}

function getUpdatedAtTimestamp(livePriceSource) {
  if (!livePriceSource?.updatedAt) {
    return Number.NEGATIVE_INFINITY;
  }

  const timestamp = Date.parse(livePriceSource.updatedAt);
  return Number.isFinite(timestamp) ? timestamp : Number.NEGATIVE_INFINITY;
}

function hasUsableQuote(quote) {
  return Number.isFinite(quote?.price);
}

function hasAppliedStoredQuotes(book) {
  return book.portfolios.some((portfolio) =>
    portfolio.holdings.some(
      (holding) => holding.priceSource === "live" || holding.priceSource === "stored",
    ),
  );
}

function choosePreferredLivePriceSource(current, candidate) {
  if (!current) {
    return candidate;
  }

  const currentIsAutomatic = isAutomaticLivePriceSource(current);
  const candidateIsAutomatic = isAutomaticLivePriceSource(candidate);

  if (currentIsAutomatic !== candidateIsAutomatic) {
    return candidateIsAutomatic ? candidate : current;
  }

  const currentUpdatedAt = getUpdatedAtTimestamp(current);
  const candidateUpdatedAt = getUpdatedAtTimestamp(candidate);

  if (currentUpdatedAt !== candidateUpdatedAt) {
    return candidateUpdatedAt > currentUpdatedAt ? candidate : current;
  }

  const currentHasQuotes = Object.values(current?.quotes ?? {}).some(hasUsableQuote);
  const candidateHasQuotes = Object.values(candidate?.quotes ?? {}).some(hasUsableQuote);

  if (currentHasQuotes !== candidateHasQuotes) {
    return candidateHasQuotes ? candidate : current;
  }

  return current;
}

function normalizeLivePriceSource(payload, fetchedFrom) {
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    throw new Error("Invalid live price payload");
  }

  return {
    ...payload,
    fetchedFrom,
  };
}

function decodeBase64Json(content) {
  const normalized = content.replace(/\n/g, "");
  const binary = globalThis.atob(normalized);
  const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
  return JSON.parse(new TextDecoder().decode(bytes));
}

function buildGitHubApiUrl(path, params = {}) {
  const url = new URL(`${GITHUB_API_BASE_URL}${path}`);
  url.searchParams.set("ts", `${Date.now()}`);

  Object.entries(params).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== "") {
      url.searchParams.set(key, `${value}`);
    }
  });

  return url;
}

async function fetchGitHubApi(path, { token, method = "GET", body, params } = {}) {
  const headers = {
    accept: "application/vnd.github+json",
    "x-github-api-version": "2022-11-28",
  };

  if (token) {
    headers.authorization = `Bearer ${token}`;
  }

  if (body) {
    headers["content-type"] = "application/json";
  }

  const response = await fetch(buildGitHubApiUrl(path, params), {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined,
  });

  let payload = null;

  if (response.status !== 204) {
    try {
      payload = await response.json();
    } catch {
      payload = null;
    }
  }

  if (!response.ok) {
    if (response.status === 401 || response.status === 403) {
      clearStoredGitHubToken();
    }

    const message = payload?.message ? `: ${payload.message}` : "";
    throw new Error(`GitHub API 요청 실패 (${response.status})${message}`);
  }

  return payload;
}

async function fetchLocalLivePrices() {
  const payload = await fetchJson(LIVE_PRICES_URL);
  return normalizeLivePriceSource(payload, "deployed-file");
}

async function fetchGitHubLivePrices(token) {
  const payload = await fetchJson(GITHUB_LIVE_PRICES_URL, {
    headers: {
      accept: "application/vnd.github+json",
      ...(token ? { authorization: `Bearer ${token}` } : {}),
    },
  });

  if (typeof payload.content !== "string") {
    throw new Error("GitHub live price response did not include content");
  }

  return normalizeLivePriceSource(decodeBase64Json(payload.content), "github-main");
}

async function loadLivePrices({ includeRepo = false, githubToken = null } = {}) {
  let localSource = null;
  let githubSource = null;
  let lastError = null;

  try {
    localSource = await fetchLocalLivePrices();
  } catch (error) {
    lastError = error;
  }

  const shouldCheckGitHub =
    includeRepo || !localSource || localSource.provider === "snapshot-fallback";

  if (shouldCheckGitHub) {
    try {
      githubSource = await fetchGitHubLivePrices(githubToken);
    } catch (error) {
      lastError ??= error;
    }
  }

  const selectedSource = [localSource, githubSource]
    .filter(Boolean)
    .reduce(choosePreferredLivePriceSource, null);

  if (!selectedSource) {
    throw lastError ?? new Error("Failed to load live prices");
  }

  return selectedSource;
}

function renderCurrentBook() {
  if (!state.portfolioSource) {
    return;
  }

  renderApp(buildBookData(state.portfolioSource, state.livePriceSource));
}

function setRefreshProgress({ phase = state.refreshPhase, detail = state.refreshDetail } = {}) {
  state.refreshPhase = phase;
  state.refreshDetail = detail;
  renderCurrentBook();
}

function promptForGitHubToken({ allowEmpty = false } = {}) {
  const existingToken = getStoredGitHubToken() ?? "";
  const promptText = allowEmpty
    ? "GitHub fine-grained token을 입력하세요. 빈 값을 입력하면 저장을 해제합니다."
    : "즉시 시세 불러오기를 쓰려면 GitHub fine-grained token이 필요합니다. Actions 읽기/쓰기 권한이 있는 토큰을 입력하세요.";
  const rawValue = window.prompt(promptText, existingToken);

  if (rawValue === null) {
    return null;
  }

  const token = rawValue.trim();

  if (!token) {
    if (allowEmpty) {
      clearStoredGitHubToken();
    }
    return null;
  }

  setStoredGitHubToken(token);
  return token;
}

async function dispatchLivePriceWorkflow(token) {
  await fetchGitHubApi(`/actions/workflows/${GITHUB_ACTIONS_WORKFLOW_ID}/dispatches`, {
    token,
    method: "POST",
    body: {
      ref: GITHUB_DEFAULT_BRANCH,
    },
  });
}

async function waitForWorkflowRun(token, startedAfterMs) {
  const deadline = Date.now() + WORKFLOW_DISCOVERY_TIMEOUT_MS;

  while (Date.now() < deadline) {
    const payload = await fetchGitHubApi(
      `/actions/workflows/${GITHUB_ACTIONS_WORKFLOW_ID}/runs`,
      {
        token,
        params: {
          event: "workflow_dispatch",
          branch: GITHUB_DEFAULT_BRANCH,
          per_page: 10,
        },
      },
    );

    const run = payload?.workflow_runs?.find(
      (workflowRun) => Date.parse(workflowRun.created_at) >= startedAfterMs,
    );

    if (run) {
      return run;
    }

    await delay(WORKFLOW_POLL_INTERVAL_MS);
  }

  throw new Error("GitHub Actions 실행을 찾지 못했습니다. 토큰 권한을 확인해 주세요.");
}

function formatWorkflowConclusion(conclusion) {
  switch (conclusion) {
    case "success":
      return "성공";
    case "failure":
      return "실패";
    case "cancelled":
      return "취소";
    case "timed_out":
      return "시간 초과";
    case "action_required":
      return "추가 작업 필요";
    default:
      return conclusion ?? "알 수 없음";
  }
}

async function waitForWorkflowCompletion(token, runId) {
  const deadline = Date.now() + WORKFLOW_COMPLETION_TIMEOUT_MS;

  while (Date.now() < deadline) {
    const run = await fetchGitHubApi(`/actions/runs/${runId}`, { token });

    if (run.status === "completed") {
      if (run.conclusion !== "success") {
        throw new Error(
          `시세 갱신 워크플로가 ${formatWorkflowConclusion(run.conclusion)} 상태로 끝났습니다. ${run.html_url}`,
        );
      }

      return run;
    }

    await delay(WORKFLOW_POLL_INTERVAL_MS);
  }

  throw new Error("시세 갱신 워크플로 완료 대기 시간이 초과되었습니다.");
}

async function waitForFreshLivePrices(previousUpdatedAt, githubToken) {
  const deadline = Date.now() + LIVE_PRICE_VISIBILITY_TIMEOUT_MS;
  let lastSource = null;
  let lastError = null;

  while (Date.now() < deadline) {
    try {
      const source = await loadLivePrices({ includeRepo: true, githubToken });
      lastSource = source;

      if (getUpdatedAtTimestamp(source) > previousUpdatedAt) {
        return source;
      }
    } catch (error) {
      lastError = error;
    }

    await delay(WORKFLOW_POLL_INTERVAL_MS);
  }

  if (lastSource) {
    return lastSource;
  }

  throw lastError ?? new Error("갱신된 가격 파일을 아직 읽지 못했습니다.");
}

function buildPortfolioData(source, livePriceSource) {
  const quotes = livePriceSource?.quotes ?? {};
  const isFallbackSource = livePriceSource?.provider === "snapshot-fallback";
  const hasAutomaticQuotes = isAutomaticLivePriceSource(livePriceSource);

  const enrichedHoldings = source.holdings.map((holding) => {
    const storedQuote = !holding.manualPriceOnly ? quotes[holding.ticker] : null;
    const shouldUseStoredQuote =
      hasUsableQuote(storedQuote) &&
      (hasAutomaticQuotes || !isFallbackSource || storedQuote.price !== holding.snapshotPrice);
    const currentPrice = shouldUseStoredQuote ? storedQuote.price : holding.snapshotPrice;
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
        : shouldUseStoredQuote
          ? hasAutomaticQuotes
            ? "live"
            : "stored"
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
  const hasStoredQuotes = hasAppliedStoredQuotes(book);

  if (state.isRefreshing) {
    return {
      tone: "status-loading",
      title:
        state.refreshPhase === "remote"
          ? "즉시 시세 갱신 실행 중"
          : "외부 시세 다시 불러오는 중",
      detail:
        state.refreshDetail ??
        "배포된 가격 파일과 저장소 원본을 순서대로 확인하고 있습니다.",
    };
  }

  if (state.refreshError) {
    return {
      tone: "status-stale",
      title: hasStoredQuotes ? "새 시세 확인 실패" : "자동 시세 확인 실패",
      detail: getErrorMessage(state.refreshError),
    };
  }

  if (isAutomaticLivePriceSource(book.livePriceSource)) {
    return {
      tone: "status-live",
      title: "외부 시세 자동 갱신 사용 중",
      detail: `${book.livePriceSource.provider} · 최근 반영 ${formatUpdatedAt(book.livePriceSource.updatedAt)}`,
    };
  }

  if (hasStoredQuotes) {
    return {
      tone: "status-live",
      title: "저장된 최근 시세 반영 중",
      detail: `${book.livePriceSource?.provider ?? "가격 파일"} · 최근 반영 ${formatUpdatedAt(book.livePriceSource?.updatedAt)}`,
    };
  }

  return {
    tone: "status-fallback",
    title: "스냅샷 가격으로 계산 중",
    detail: book.livePriceSource?.updatedAt
      ? `저장 파일 확인 ${formatUpdatedAt(book.livePriceSource.updatedAt)} · 현재 스냅샷과 동일`
      : "저장된 시세가 없어 마지막 스냅샷으로 계산합니다.",
  };
}

function getPriceSourceLabel(priceSource) {
  switch (priceSource) {
    case "live":
      return "자동 시세";
    case "stored":
      return "저장 시세";
    case "manual":
      return "수동 가격";
    default:
      return "스냅샷";
  }
}

function getPriceSourceDetail(priceSource) {
  switch (priceSource) {
    case "live":
      return "외부 시세 반영";
    case "stored":
      return "최근 저장 시세 반영";
    case "manual":
      return "수동 기준 가격";
    default:
      return "스크린샷 기준";
  }
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
              <span class="tag ${holding.priceSource === "snapshot" || holding.priceSource === "manual" ? "tag-warning" : "tag-live"}">
                ${getPriceSourceLabel(holding.priceSource)}
              </span>
              ${holding.nameInferred ? '<span class="tag tag-warning">이름 추정</span>' : ""}
            </div>
          </td>
          <td>${currencyFormatter.format(holding.quantity)}주</td>
          <td>
            <div class="metric">
              <strong>${formatPrice(holding.currentPrice)}</strong>
              <span>${getPriceSourceDetail(holding.priceSource)}</span>
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
              ${portfolio.snapshotLabel} 기준 원장 데이터를 유지하고, 현재가만 저장된 최신 시세 파일에서 다시 읽어
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
              저장된 최신 가격 파일이 있으면 현재가 계산에 우선 반영하고, 없으면 마지막 스냅샷 가격으로 계산합니다.
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
          <div class="toolbar-actions">
            <button class="ghost-button" data-refresh-quotes ${state.isRefreshing ? "disabled" : ""}>
              ${state.isRefreshing
                ? state.refreshPhase === "remote"
                  ? "즉시 갱신 중..."
                  : "업데이트 중..."
                : "시세 즉시 불러오기"}
            </button>
            <button class="ghost-button ghost-button-muted" data-manage-refresh-token ${state.isRefreshing ? "disabled" : ""}>
              ${hasStoredGitHubToken() ? "토큰 변경" : "토큰 설정"}
            </button>
          </div>
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

async function refreshLivePrices({ triggerRemote = false } = {}) {
  if (!state.portfolioSource || state.isRefreshing) {
    return;
  }

  const previousUpdatedAt = getUpdatedAtTimestamp(state.livePriceSource);
  state.isRefreshing = true;
  state.refreshError = null;
  state.refreshPhase = triggerRemote ? "remote" : "local";
  state.refreshDetail = triggerRemote
    ? "GitHub 토큰과 워크플로 권한을 확인하고 있습니다."
    : "배포된 가격 파일과 저장소 원본을 순서대로 확인하고 있습니다.";
  renderCurrentBook();

  try {
    if (triggerRemote) {
      const token = getStoredGitHubToken() ?? promptForGitHubToken();

      if (!token) {
        return;
      }

      setRefreshProgress({
        phase: "remote",
        detail: "GitHub Actions에 최신 시세 갱신을 요청하고 있습니다.",
      });
      const requestStartedAt = Date.now() - 5000;
      await dispatchLivePriceWorkflow(token);

      setRefreshProgress({
        phase: "remote",
        detail: "GitHub Actions 실행을 기다리고 있습니다.",
      });
      const run = await waitForWorkflowRun(token, requestStartedAt);

      setRefreshProgress({
        phase: "remote",
        detail: "한국투자 Open API에서 최신 시세를 불러오고 있습니다.",
      });
      await waitForWorkflowCompletion(token, run.id);

      setRefreshProgress({
        phase: "remote",
        detail: "갱신된 가격 파일을 다시 읽고 있습니다.",
      });
      state.livePriceSource = await waitForFreshLivePrices(previousUpdatedAt, token);
    } else {
      state.livePriceSource = await loadLivePrices({ includeRepo: true });
    }
  } catch (error) {
    state.refreshError = error;
  } finally {
    state.isRefreshing = false;
    state.refreshPhase = null;
    state.refreshDetail = null;
    renderCurrentBook();
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
      state.livePriceSource = await loadLivePrices();
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
  const manageTokenButton = event.target.closest("[data-manage-refresh-token]");

  if (manageTokenButton) {
    promptForGitHubToken({ allowEmpty: true });
    renderCurrentBook();
    return;
  }

  const refreshButton = event.target.closest("[data-refresh-quotes]");

  if (!refreshButton) {
    return;
  }

  if (!state.portfolioSource) {
    initializeApp();
    return;
  }

  refreshLivePrices({ triggerRemote: true });
});

initializeApp();
