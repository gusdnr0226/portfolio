const account = {
  name: "연금저축1",
  snapshotLabel: "2026년 3월 9일 20:22 스크린샷 기준",
  cash: 94647,
  holdings: [
    {
      name: "KODEX 미국S&P500",
      type: "ETF",
      quantity: 938,
      marketValue: 21100310,
      profitLoss: -234300,
      profitRate: -0.011,
    },
    {
      name: "KODEX 미국성장커버드콜액티브",
      type: "ETF",
      quantity: 1000,
      marketValue: 9320000,
      profitLoss: -607070,
      profitRate: -0.0612,
      nameInferred: true,
    },
    {
      name: "KODEX 미국나스닥100",
      type: "ETF",
      quantity: 259,
      marketValue: 6236720,
      profitLoss: -192550,
      profitRate: -0.0299,
    },
    {
      name: "맥쿼리인프라",
      type: "인프라",
      quantity: 138,
      marketValue: 1551120,
      profitLoss: -11300,
      profitRate: -0.0072,
    },
  ],
};

const currencyFormatter = new Intl.NumberFormat("ko-KR");
const decimalFormatter = new Intl.NumberFormat("ko-KR", {
  minimumFractionDigits: 0,
  maximumFractionDigits: 1,
});
const percentFormatter = new Intl.NumberFormat("ko-KR", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

function formatCurrency(value) {
  return `${currencyFormatter.format(Math.round(value))}원`;
}

function formatPrice(value) {
  const rounded = Math.round(value);
  return `${currencyFormatter.format(rounded)}원`;
}

function formatAverageCost(value) {
  const isWhole = Math.abs(value - Math.round(value)) < 0.05;
  const displayValue = isWhole ? Math.round(value) : Number(value.toFixed(1));
  return `${decimalFormatter.format(displayValue)}원`;
}

function formatPercent(value) {
  return `${percentFormatter.format(value * 100)}%`;
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

function buildPortfolioData(source) {
  const enrichedHoldings = source.holdings.map((holding) => {
    const costBasis = holding.marketValue - holding.profitLoss;
    const averageCost = costBasis / holding.quantity;
    const currentPrice = holding.marketValue / holding.quantity;

    return {
      ...holding,
      costBasis,
      averageCost,
      currentPrice,
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
  const investedWeight = totals.marketValue / totalAssets;
  const cashWeight = source.cash / totalAssets;
  const totalReturn = totals.profitLoss / totals.costBasis;

  const holdingsWithWeight = enrichedHoldings.map((holding) => ({
    ...holding,
    assetWeight: holding.marketValue / totalAssets,
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
      value: formatCurrency(portfolio.totals.profitLoss),
      detail: `보유 종목 기준 ${formatPercent(portfolio.totals.totalReturn)}`,
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

function renderTableRows(portfolio) {
  return portfolio.holdings
    .map(
      (holding) => `
        <tr>
          <td>
            <div class="asset-name">${holding.name}</div>
            <div class="asset-meta">
              <span class="tag">${holding.type}</span>
              ${holding.nameInferred ? '<span class="tag tag-warning">이름 추정</span>' : ""}
            </div>
          </td>
          <td>${currencyFormatter.format(holding.quantity)}주</td>
          <td>
            <div class="metric">
              <strong>${formatPrice(holding.currentPrice)}</strong>
              <span>1주 기준</span>
            </div>
          </td>
          <td>${formatAverageCost(holding.averageCost)}</td>
          <td>${formatCurrency(holding.costBasis)}</td>
          <td>${formatCurrency(holding.marketValue)}</td>
          <td class="${getToneClass(holding.profitLoss)}">${formatCurrency(holding.profitLoss)}</td>
          <td class="${getToneClass(holding.profitRate)}">${formatPercent(holding.profitRate)}</td>
          <td>${formatPercent(holding.assetWeight)}</td>
        </tr>
      `,
    )
    .join("");
}

function renderApp(portfolio) {
  const app = document.querySelector("#app");

  if (!app) {
    return;
  }

  app.innerHTML = `
    <section class="dashboard">
      <section class="hero">
        <div class="eyebrow">Pension Saving Portfolio</div>
        <div class="hero-head">
          <div>
            <h1 class="hero-title">${portfolio.name} 계좌 현황</h1>
            <p class="hero-copy">
              ${portfolio.snapshotLabel} 데이터를 기준으로 종목별 수량, 평균단가, 매입금액,
              평가금액, 평가손익을 한 화면에서 보도록 정리했습니다.
            </p>
          </div>
          <div class="account-pill">기준 통화 KRW</div>
        </div>
        <div class="summary-grid">
          ${renderSummaryCards(portfolio)}
        </div>
      </section>

      <section class="table-card">
        <div class="section-head">
          <div>
            <h2 class="section-title">보유 종목 표</h2>
            <p class="section-copy">
              평균단가와 매입금액은 스크린샷의 수량, 평가금액, 평가손익을 기준으로 역산했습니다.
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
                <td class="${getToneClass(portfolio.totals.profitLoss)}">${formatCurrency(portfolio.totals.profitLoss)}</td>
                <td class="${getToneClass(portfolio.totals.totalReturn)}">${formatPercent(portfolio.totals.totalReturn)}</td>
                <td>${formatPercent(portfolio.totals.investedWeight)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </section>

      <section class="notes-card">
        <h2 class="notes-title">메모</h2>
        <ul class="notes-list">
          <li><strong>현금</strong> 94,647원은 별도 보유 자산으로 요약 카드에 반영했습니다.</li>
          <li><strong>맥쿼리인프라</strong>를 포함한 현재 보유 종목 손익 합계는 -1,045,220원입니다.</li>
          <li><strong>KODEX 미국성장커버드콜액티브</strong> 명칭은 스크린샷에서 축약된 텍스트를 바탕으로 보완한 값입니다.</li>
        </ul>
      </section>
    </section>
  `;
}

renderApp(buildPortfolioData(account));
