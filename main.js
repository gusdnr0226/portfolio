const PORTFOLIOS_URL = "./data/portfolios.json";
const LIVE_PRICES_URL = "./data/live-prices.json";
const DIVIDENDS_URL = "./data/dividends.json";
const DIVIDEND_CALENDAR_URL = "./data/dividend-calendar.json";
const GITHUB_API_BASE_URL = "https://api.github.com/repos/gusdnr0226/portfolio";
const GITHUB_LIVE_PRICES_URL =
  `${GITHUB_API_BASE_URL}/contents/data/live-prices.json?ref=main`;
const LIVE_PRICE_PROXY_BASE_URL = "https://r.jina.ai/http://";
const YAHOO_SPARK_BASE_URL = "https://query1.finance.yahoo.com/v7/finance/spark";
const YAHOO_CHART_BASE_URL = "https://query1.finance.yahoo.com/v8/finance/chart";
const YAHOO_SPARK_BATCH_SIZE = 20;
const REFRESH_INTERVAL_MS = 5 * 60 * 1000;
const AMOUNT_VISIBILITY_STORAGE_KEY = "portfolio.hide-amounts";
const MONTH_LABELS = Array.from({ length: 12 }, (_, index) => `${index + 1}월`);
const DIVIDEND_OVERRIDE_RULES = {
  ULTY: {
    yieldRate: 0.5,
    calendarStrategy: "weekly-even-52",
  },
};
const TREEMAP_LAYOUT_ORDER_BY_SUPER = {
  "해외": ["SCHD", "S&P500", "Nasdaq", "커버드콜", "부동산/리츠", "채권", "USD", "기타"],
  "국내": ["배당ETF", "인프라", "통신", "제조", "금융", "원화", "MM"],
};
const TREEMAP_LAYOUT_ORDER_GLOBAL = [
  "SCHD",
  "S&P500",
  "Nasdaq",
  "커버드콜",
  "부동산/리츠",
  "채권",
  "USD",
  "배당ETF",
  "인프라",
  "통신",
  "제조",
  "금융",
  "원화",
  "MM",
  "기타",
];
const ACCOUNT_SUMMARY_GROUPS = [
  {
    key: "direct-investment",
    label: "직투계좌",
    color: "#6fa9d8",
  },
  {
    key: "isa",
    label: "ISA",
    color: "#5ca893",
  },
  {
    key: "pension-savings-deductible",
    label: "연금저축(세액공제)",
    color: "#d3a15e",
  },
  {
    key: "pension-savings-non-deductible",
    label: "연금저축(세액미공제)",
    color: "#d28797",
  },
  {
    key: "retirement-pension",
    label: "퇴직연금",
    color: "#7e93ae",
  },
];
const PORTFOLIO_DISPLAY_ORDER = [
  "direct-investment",
  "isa",
  "pension-savings-1",
  "pension-savings-ingyeong",
  "pension-savings-2",
  "retirement-pension-hyeonuk",
  "retirement-pension-ingyeong",
];
const portfolioDisplayOrderLookup = new Map(
  PORTFOLIO_DISPLAY_ORDER.map((id, index) => [id, index]),
);
const CATEGORY_TICKER_RULES = {
  "SCHD": [
    "SCHD",
    "TIGER 미국배당다우존스",
    "SOL 미국배당다우존스2호",
    "SOL 미국배당다우존스",
  ],
  "Nasdaq": [
    "QQQ",
    "QLD",
    "KODEX 미국나스닥100",
    "TIGER 미국나스닥100레버리지(합성)",
    "KODEX 미국반도체",
    "KODEX 미국휴머노이드로봇",
  ],
  "S&P500": [
    "VOO",
    "SSO",
    "KODEX 미국S&P500",
    "ACE 미국S&P500미국채혼합50액티브",
  ],
  "커버드콜": [
    "JEPQ",
    "JEPI",
    "TIGER 미국배당다우존스타겟커버드콜2호",
    "KODEX 미국성장커버드콜액티브",
    "KODEX 미국배당커버드콜액티브",
  ],
  "인프라": [
    "KODEX 한국부동산리츠인프라",
  ],
  "부동산/리츠": [
    "AGNC",
    "O",
  ],
  "채권": [
    "IEF",
    "TLT",
  ],
  "배당ETF": [
    "PLUS 고배당주채권혼합",
    "PLUS 자사주매입고배당주",
    "KODEX 200타겟위클리커버드콜",
    "TIGER 코리아배당다우존스",
  ],
  "원화": [
    "자동운용상품(고유계정대)",
    "KRW",
  ],
  "MM": [
    "KODEX 머니마켓액티브",
  ],
  "USD": [
    "USD",
  ],
  "기타": [
    "ULTY",
  ],
};
const CATEGORY_NAME_ALIASES = {
  "유틸리티": "인프라",
  "utility": "인프라",
  "utilities": "인프라",
  "국내 배당": "배당ETF",
  "국내배당": "배당ETF",
  "VOO": "S&P500",
  "QQQ": "Nasdaq",
  "미국 커버드콜": "커버드콜",
  "현금(USD)": "USD",
  "운송물류": "제조",
  "통신사": "통신",
  "미국 국채": "채권",
};
const SUPER_CATEGORY_RULES = {
  "해외": [
    "커버드콜",
    "SCHD",
    "S&P500",
    "Nasdaq",
    "부동산/리츠",
    "채권",
    "USD",
    "기타",
  ],
  "국내": [
    "금융",
    "인프라",
    "배당ETF",
    "제조",
    "통신",
    "원화",
    "MM",
  ],
};
const TREEMAP_SUPER_CATEGORY_ORDER = Object.keys(TREEMAP_LAYOUT_ORDER_BY_SUPER);
const TREEMAP_CATEGORY_COLORS = {
  "SCHD": "#6fa9d8",
  "S&P500": "#7a8eea",
  Nasdaq: "#5d73d6",
  "커버드콜": "#b86f92",
  "부동산/리츠": "#5ca893",
  "채권": "#6f9fb3",
  USD: "#b99b58",
  "배당ETF": "#8ca56c",
  "인프라": "#5f9b80",
  "통신": "#5e81bc",
  "제조": "#b98862",
  "금융": "#8c88b9",
  "원화": "#8e9ab4",
  MM: "#b5a27b",
  "기타": "#7a859d",
};
const PORTFOLIO_CIRCLE_GROUPS = [
  {
    key: "SCHD",
    label: "SCHD",
    color: TREEMAP_CATEGORY_COLORS["SCHD"],
    memberOrder: ["SCHD"],
  },
  {
    key: "S&P500",
    label: "S&P500",
    color: TREEMAP_CATEGORY_COLORS["S&P500"],
    memberOrder: ["S&P500"],
  },
  {
    key: "Nasdaq",
    label: "Nasdaq",
    color: TREEMAP_CATEGORY_COLORS.Nasdaq,
    memberOrder: ["Nasdaq"],
  },
  {
    key: "커버드콜",
    label: "커버드콜",
    color: TREEMAP_CATEGORY_COLORS["커버드콜"],
    memberOrder: ["커버드콜"],
  },
  {
    key: "채권",
    label: "채권",
    color: TREEMAP_CATEGORY_COLORS["채권"],
    memberOrder: ["채권"],
  },
  {
    key: "해외기타",
    label: "해외 기타",
    color: "#5ca893",
    memberOrder: ["부동산/리츠", "기타"],
  },
  {
    key: "국내 주식/ETF",
    label: "국내 주식/ETF",
    color: "#7c9368",
    memberOrder: ["배당ETF", "인프라", "제조", "통신", "금융"],
  },
  {
    key: "현금성자산",
    label: "현금성자산",
    color: "#8e9ab4",
    memberOrder: ["원화", "USD", "MM"],
  },
];
const TREEMAP_CANVAS_ASPECT_DESKTOP = 16 / 10;
const TREEMAP_CANVAS_ASPECT_TABLET = 1.16;
const TREEMAP_CANVAS_ASPECT_MOBILE = 1;

const currencyFormatter = new Intl.NumberFormat("ko-KR");
const compactCurrencyFormatter = new Intl.NumberFormat("ko-KR", {
  notation: "compact",
  maximumFractionDigits: 1,
});
const quantityFormatter = new Intl.NumberFormat("ko-KR", {
  minimumFractionDigits: 0,
  maximumFractionDigits: 1,
});
const decimalFormatter = new Intl.NumberFormat("ko-KR", {
  minimumFractionDigits: 0,
  maximumFractionDigits: 1,
});
const exchangeRateFormatter = new Intl.NumberFormat("ko-KR", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});
const usdCurrencyFormatter = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});
const usdCompactCurrencyFormatter = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  minimumFractionDigits: 1,
  maximumFractionDigits: 1,
});
const usdWholeCurrencyFormatter = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  minimumFractionDigits: 0,
  maximumFractionDigits: 0,
});
const percentFormatter = new Intl.NumberFormat("ko-KR", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});
const weightPercentFormatter = new Intl.NumberFormat("ko-KR", {
  minimumFractionDigits: 1,
  maximumFractionDigits: 1,
});
const dateTimeFormatter = new Intl.DateTimeFormat("ko-KR", {
  dateStyle: "medium",
  timeStyle: "short",
  timeZone: "Asia/Seoul",
});

const state = {
  portfolioSource: null,
  livePriceSource: null,
  liveExchangeRateSource: null,
  dividendSource: null,
  dividendCalendarSource: null,
  dividendCalendarError: null,
  hideAmounts: loadAmountVisibilityPreference(),
  isRefreshing: false,
  refreshError: null,
  refreshPhase: null,
  refreshDetail: null,
  refreshTimerId: null,
};
let treemapTypographyFrameId = 0;
let treemapResizeFrameId = 0;
let sidebarSectionObserver = null;

function normalizeTreemapToken(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

const treemapCategoryLookup = Object.entries(CATEGORY_TICKER_RULES).reduce(
  (lookup, [category, values]) => {
    values.forEach((value) => {
      lookup.set(normalizeTreemapToken(value), category);
    });
    return lookup;
  },
  new Map(),
);
const treemapAliasLookup = new Map(
  Object.entries(CATEGORY_NAME_ALIASES).map(([alias, category]) => [
    normalizeTreemapToken(alias),
    category,
  ]),
);
const treemapSuperCategoryLookup = Object.entries(SUPER_CATEGORY_RULES).reduce(
  (lookup, [superCategory, categories]) => {
    categories.forEach((category) => {
      lookup.set(category, superCategory);
    });
    return lookup;
  },
  new Map(),
);
const treemapCategoryOrderLookup = new Map(
  TREEMAP_LAYOUT_ORDER_GLOBAL.map((category, index) => [category, index]),
);

function loadAmountVisibilityPreference() {
  try {
    return window.localStorage.getItem(AMOUNT_VISIBILITY_STORAGE_KEY) === "true";
  } catch (error) {
    return false;
  }
}

function persistAmountVisibilityPreference() {
  try {
    window.localStorage.setItem(AMOUNT_VISIBILITY_STORAGE_KEY, String(state.hideAmounts));
  } catch (error) {
    // Ignore storage errors and keep the in-memory preference.
  }
}

function shouldHideAmounts() {
  return Boolean(state.hideAmounts);
}

function getMaskedAmountText(value, currency = "KRW", { signed = false, unitless = false } = {}) {
  const prefix = signed ? (value > 0 ? "+" : value < 0 ? "-" : "") : "";

  if (unitless) {
    return `${prefix}****`;
  }

  return currency === "USD" ? `${prefix}$****` : `${prefix}****원`;
}

function formatCurrency(value) {
  if (shouldHideAmounts()) {
    return getMaskedAmountText(value, "KRW");
  }

  return `${currencyFormatter.format(Math.round(value))}원`;
}

function formatCompactCurrency(value) {
  if (shouldHideAmounts()) {
    return getMaskedAmountText(value, "KRW");
  }

  return `${compactCurrencyFormatter.format(Math.round(value))}원`;
}

function formatCompactMoney(value, currency = "KRW") {
  if (shouldHideAmounts()) {
    return getMaskedAmountText(value, currency);
  }

  if (currency === "USD") {
    return usdCompactCurrencyFormatter.format(Number(value.toFixed(1)));
  }

  return formatCompactCurrency(value);
}

function formatSignedCurrency(value) {
  if (shouldHideAmounts()) {
    return getMaskedAmountText(value, "KRW", { signed: true });
  }

  const rounded = Math.round(value);
  const prefix = rounded > 0 ? "+" : "";
  return `${prefix}${currencyFormatter.format(rounded)}원`;
}

function getHoldingCurrency(holding) {
  if (holding?.currency) {
    return holding.currency;
  }

  return /^[0-9]/.test(holding?.ticker ?? "") ? "KRW" : "USD";
}

function getHoldingAverageCost(holding) {
  const quantity = Number(holding?.quantity);
  const costBasis = Number(holding?.costBasis);

  if (!Number.isFinite(quantity) || quantity === 0 || !Number.isFinite(costBasis)) {
    return null;
  }

  return costBasis / quantity;
}

function getHoldingAverageCostKrw(holding) {
  return Number.isFinite(holding?.averageCostKrw) ? holding.averageCostKrw : null;
}

function getCashKrw(account) {
  return Number(account?.cash ?? 0);
}

function getCashUsd(account) {
  return Number(account?.cashUsd ?? 0);
}

function getDividendOverrideRule(ticker) {
  return typeof ticker === "string" ? DIVIDEND_OVERRIDE_RULES[ticker] ?? null : null;
}

function getAnnualDividendPerShare(dividendSource, holding, currentPrice) {
  const overrideRule = getDividendOverrideRule(holding?.ticker);

  if (
    Number.isFinite(overrideRule?.yieldRate) &&
    Number.isFinite(currentPrice) &&
    currentPrice > 0
  ) {
    return currentPrice * overrideRule.yieldRate;
  }

  const annualDividendPerShare = dividendSource?.dividends?.[holding?.ticker]?.annualDividendPerShare;
  return Number.isFinite(annualDividendPerShare) ? annualDividendPerShare : 0;
}

function getDividendCalendarProjectionYear(source) {
  return Number.isFinite(source?.projectionYear) && source.projectionYear > 0
    ? source.projectionYear
    : new Date().getUTCFullYear();
}

function buildWeeklyEvenMonthlyAmounts(year, annualAmount) {
  if (!Number.isFinite(year) || year <= 0 || !Number.isFinite(annualAmount) || annualAmount <= 0) {
    return {};
  }

  const firstFriday = new Date(Date.UTC(year, 0, 1));

  while (firstFriday.getUTCDay() !== 5) {
    firstFriday.setUTCDate(firstFriday.getUTCDate() + 1);
  }

  const weeklyPayouts = Array.from({ length: 52 }, (_, index) => {
    const payout = new Date(firstFriday);
    payout.setUTCDate(firstFriday.getUTCDate() + index * 7);
    return payout;
  });
  const rawMonthlyTotals = Array.from({ length: 12 }, () => 0);
  const weeklyAmount = annualAmount / weeklyPayouts.length;

  weeklyPayouts.forEach((payoutDate) => {
    rawMonthlyTotals[payoutDate.getUTCMonth()] += weeklyAmount;
  });

  let allocatedAmount = 0;

  return rawMonthlyTotals.reduce((monthlyAmounts, rawMonthlyAmount, monthIndex) => {
    const isLastMonth = monthIndex === rawMonthlyTotals.length - 1;
    const amount = isLastMonth
      ? Number((annualAmount - allocatedAmount).toFixed(2))
      : Number(rawMonthlyAmount.toFixed(2));

    if (!isLastMonth) {
      allocatedAmount += amount;
    }

    if (amount !== 0) {
      monthlyAmounts[monthIndex + 1] = amount;
    }

    return monthlyAmounts;
  }, {});
}

function getDividendCalendarMonthlyAmounts(entry, reference, projectionYear) {
  const overrideRule = getDividendOverrideRule(entry?.ticker ?? reference?.ticker);

  if (
    overrideRule?.calendarStrategy === "weekly-even-52" &&
    Number.isFinite(reference?.currentPrice) &&
    Number.isFinite(reference?.quantity) &&
    reference.quantity > 0
  ) {
    return buildWeeklyEvenMonthlyAmounts(
      projectionYear,
      reference.currentPrice * overrideRule.yieldRate * reference.quantity,
    );
  }

  return entry?.monthlyAmounts &&
    typeof entry.monthlyAmounts === "object" &&
    !Array.isArray(entry.monthlyAmounts)
    ? entry.monthlyAmounts
    : {};
}

function formatMoney(value, currency = "KRW") {
  if (shouldHideAmounts()) {
    return getMaskedAmountText(value, currency);
  }

  if (currency === "USD") {
    return usdCurrencyFormatter.format(value);
  }

  return formatCurrency(value);
}

function formatSignedMoney(value, currency = "KRW") {
  if (shouldHideAmounts()) {
    return getMaskedAmountText(value, currency, { signed: true });
  }

  if (currency === "USD") {
    const prefix = value > 0 ? "+" : value < 0 ? "-" : "";
    return `${prefix}${usdCurrencyFormatter.format(Math.abs(value))}`;
  }

  return formatSignedCurrency(value);
}

function formatPrice(value, currency = "KRW") {
  return formatMoney(value, currency);
}

function formatHoldingMoney(value, currency = "KRW") {
  if (shouldHideAmounts()) {
    return getMaskedAmountText(value, currency);
  }

  if (currency === "USD") {
    return usdWholeCurrencyFormatter.format(value);
  }

  return formatCurrency(value);
}

function formatSignedHoldingMoney(value, currency = "KRW") {
  if (shouldHideAmounts()) {
    return getMaskedAmountText(value, currency, { signed: true });
  }

  if (currency === "USD") {
    const prefix = value > 0 ? "+" : value < 0 ? "-" : "";
    return `${prefix}${usdWholeCurrencyFormatter.format(Math.abs(value))}`;
  }

  return formatSignedCurrency(value);
}

function formatOverallMoney(value, currency = "KRW") {
  if (shouldHideAmounts()) {
    return getMaskedAmountText(value, currency);
  }

  if (currency === "USD") {
    return usdCompactCurrencyFormatter.format(Number(value.toFixed(1)));
  }

  return formatCurrency(value);
}

function formatSignedOverallMoney(value, currency = "KRW") {
  if (shouldHideAmounts()) {
    return getMaskedAmountText(value, currency, { signed: true });
  }

  if (currency === "USD") {
    const prefix = value > 0 ? "+" : value < 0 ? "-" : "";
    return `${prefix}${usdCompactCurrencyFormatter.format(Number(Math.abs(value).toFixed(1)))}`;
  }

  return formatSignedCurrency(value);
}

function formatOverallAverageCost(value, currency = "KRW") {
  return formatOverallMoney(value, currency);
}

function formatOverallProfitWithRate(profitLoss, profitRate, currency = "KRW") {
  return `${formatSignedOverallMoney(profitLoss, currency)}(${formatSignedPercentCompact(profitRate)})`;
}

function formatOverallDividendYield(annualDividendPerShare, currentPrice) {
  if (!Number.isFinite(currentPrice) || currentPrice <= 0) {
    return "-";
  }

  return formatPercent(annualDividendPerShare / currentPrice);
}

function formatAverageCost(value, currency = "KRW") {
  if (shouldHideAmounts()) {
    return getMaskedAmountText(value, currency);
  }

  if (currency === "USD") {
    return usdCurrencyFormatter.format(value);
  }

  const isWhole = Math.abs(value - Math.round(value)) < 0.05;
  const displayValue = isWhole ? Math.round(value) : Number(value.toFixed(1));
  return `${decimalFormatter.format(displayValue)}원`;
}

function renderAverageCostCell(averageCost, averageCostCurrency, averageCostKrw, { overall = false } = {}) {
  if (!Number.isFinite(averageCost)) {
    return "-";
  }

  const formatCost = overall ? formatOverallAverageCost : formatAverageCost;
  const primaryLabel = formatCost(averageCost, averageCostCurrency);

  if (!Number.isFinite(averageCostKrw) || averageCostCurrency === "KRW") {
    return primaryLabel;
  }

  return `
    <div class="all-holdings-value-stack">
      <strong>${primaryLabel}</strong>
      <span class="all-holdings-value-sub">${formatCost(averageCostKrw, "KRW")}</span>
    </div>
  `;
}

function formatQuantity(value) {
  const roundedValue = Number(value.toFixed(1));
  return quantityFormatter.format(roundedValue);
}

function formatPercent(value) {
  return `${weightPercentFormatter.format(value * 100)}%`;
}

function formatTreemapPercent(value) {
  return `${weightPercentFormatter.format(value * 100)}%`;
}

function formatTreemapAmount(value) {
  if (shouldHideAmounts()) {
    return getMaskedAmountText(value, "KRW", { unitless: true });
  }

  return currencyFormatter.format(Math.round(value));
}

function formatTreemapCompactAmount(value) {
  if (shouldHideAmounts()) {
    return getMaskedAmountText(value, "KRW", { unitless: true });
  }

  return compactCurrencyFormatter.format(Math.round(value));
}

function formatSignedPercent(value) {
  const prefix = value > 0 ? "+" : "";
  return `${prefix}${percentFormatter.format(value * 100)}%`;
}

function formatSignedPercentCompact(value) {
  const prefix = value > 0 ? "+" : "";
  return `${prefix}${weightPercentFormatter.format(value * 100)}%`;
}

function formatSignedPercentSingleDecimal(value) {
  const prefix = value > 0 ? "+" : "";
  return `${prefix}${weightPercentFormatter.format(value * 100)}%`;
}

function formatProfitWithRateSingleDecimal(profitLoss, costBasis) {
  if (costBasis <= 0) {
    return formatSignedCurrency(profitLoss);
  }

  return `${formatSignedCurrency(profitLoss)} (${formatSignedPercentSingleDecimal(profitLoss / costBasis)})`;
}

function formatProfitWithRate(profitLoss, costBasis) {
  if (costBasis <= 0) {
    return formatSignedCurrency(profitLoss);
  }

  return `${formatSignedCurrency(profitLoss)} (${formatSignedPercent(profitLoss / costBasis)})`;
}

function formatHoldingProfitWithRate(profitLoss, profitRate, currency = "KRW") {
  return `${formatSignedHoldingMoney(profitLoss, currency)}(${formatSignedPercentCompact(profitRate)})`;
}

function formatProfitWithRateCompact(profitLoss, profitRate) {
  return `${formatSignedCurrency(profitLoss)}(${formatSignedPercentCompact(profitRate)})`;
}

function calculateTotalReturnAmount(unrealizedProfitLoss, realizedProfitLoss, cumulativeDividendIncome) {
  return unrealizedProfitLoss + realizedProfitLoss + cumulativeDividendIncome;
}

function calculateTotalReturnRate(totalReturnAmount, costBasis) {
  if (costBasis === 0) {
    return 0;
  }

  return totalReturnAmount / costBasis;
}

function formatCashBreakdown(cashKrw, cashUsd) {
  const parts = [];

  if (cashKrw > 0) {
    parts.push(`원화 ${formatCurrency(cashKrw)}`);
  }

  if (cashUsd > 0) {
    parts.push(`달러 ${formatMoney(cashUsd, "USD")}`);
  }

  return parts.join(" + ");
}

function formatUpdatedAt(value) {
  if (!value) {
    return "업데이트 시각 없음";
  }

  return dateTimeFormatter.format(new Date(value));
}

function formatExchangeRate(value) {
  if (!Number.isFinite(value) || value <= 0) {
    return "-";
  }

  return `${exchangeRateFormatter.format(value)}원`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
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

async function fetchText(url, options = {}) {
  const response = await fetch(url, {
    cache: "no-store",
    ...options,
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch ${url}: ${response.status}`);
  }

  return response.text();
}

function isAutomaticLivePriceSource(livePriceSource) {
  return ["KIS Open API", "Yahoo Finance via r.jina.ai"].includes(
    livePriceSource?.provider,
  );
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

function countUsableQuotes(livePriceSource) {
  return Object.values(livePriceSource?.quotes ?? {}).filter(hasUsableQuote).length;
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

  const currentQuoteCount = countUsableQuotes(current);
  const candidateQuoteCount = countUsableQuotes(candidate);

  if (currentQuoteCount !== candidateQuoteCount) {
    return candidateQuoteCount > currentQuoteCount ? candidate : current;
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

function getExternalQuoteTickers(portfolioSource) {
  if (!portfolioSource?.accounts) {
    return [];
  }

  const tickers = new Set();
  let needsUsdKrwRate = false;

  portfolioSource.accounts.forEach((account) => {
    account.holdings.forEach((holding) => {
      if (!holding.manualPriceOnly) {
        tickers.add(holding.ticker);
      }

      if (getHoldingCurrency(holding) === "USD") {
        needsUsdKrwRate = true;
      }
    });

    if (getCashUsd(account) > 0) {
      needsUsdKrwRate = true;
    }
  });

  if (needsUsdKrwRate) {
    tickers.add("KRW=X");
  }

  return [...tickers];
}

function toYahooSymbol(ticker) {
  if (ticker === "KRW=X") {
    return ticker;
  }

  return /^[0-9]/.test(ticker) ? `${ticker}.KS` : ticker;
}

function toTickerFromYahooSymbol(symbol) {
  return symbol.replace(/\.KS$/i, "");
}

function chunkArray(values, size) {
  const chunks = [];

  for (let index = 0; index < values.length; index += size) {
    chunks.push(values.slice(index, index + size));
  }

  return chunks;
}

function extractJsonPayload(text) {
  const trimmedText = text.trim();

  if (trimmedText.startsWith("{") || trimmedText.startsWith("[")) {
    return JSON.parse(trimmedText);
  }

  const marker = "Markdown Content:";
  const body = trimmedText.includes(marker)
    ? trimmedText.slice(trimmedText.indexOf(marker) + marker.length)
    : trimmedText;
  const firstBraceIndex = body.indexOf("{");
  const lastBraceIndex = body.lastIndexOf("}");

  if (firstBraceIndex === -1 || lastBraceIndex === -1 || lastBraceIndex < firstBraceIndex) {
    throw new Error("외부 시세 응답에서 JSON 본문을 찾지 못했습니다.");
  }

  return JSON.parse(body.slice(firstBraceIndex, lastBraceIndex + 1));
}

function buildYahooSparkProxyUrl(symbols) {
  const yahooUrl = new URL(YAHOO_SPARK_BASE_URL);
  yahooUrl.searchParams.set("symbols", symbols.join(","));
  yahooUrl.searchParams.set("range", "1d");
  yahooUrl.searchParams.set("interval", "1d");
  yahooUrl.searchParams.set("indicators", "close");
  yahooUrl.searchParams.set("_", `${Date.now()}`);

  return `${LIVE_PRICE_PROXY_BASE_URL}${yahooUrl.toString()}`;
}

function buildYahooChartProxyUrl(symbol, { range = "1d", interval = "1m" } = {}) {
  const yahooUrl = new URL(`${YAHOO_CHART_BASE_URL}/${encodeURIComponent(symbol)}`);
  yahooUrl.searchParams.set("range", range);
  yahooUrl.searchParams.set("interval", interval);
  yahooUrl.searchParams.set("includePrePost", "true");
  yahooUrl.searchParams.set("_", `${Date.now()}`);

  return `${LIVE_PRICE_PROXY_BASE_URL}${yahooUrl.toString()}`;
}

function getYahooSparkPrice(result) {
  const response = Array.isArray(result?.response) ? result.response[0] : null;
  const metaPrice = response?.meta?.regularMarketPrice;

  if (Number.isFinite(metaPrice)) {
    return metaPrice;
  }

  const closePrices = response?.indicators?.quote?.[0]?.close;

  if (!Array.isArray(closePrices)) {
    return null;
  }

  for (let index = closePrices.length - 1; index >= 0; index -= 1) {
    if (Number.isFinite(closePrices[index])) {
      return closePrices[index];
    }
  }

  return null;
}

function getYahooChartPrice(payload) {
  const result = Array.isArray(payload?.chart?.result) ? payload.chart.result[0] : null;
  const metaPrice = result?.meta?.regularMarketPrice;

  if (Number.isFinite(metaPrice) && metaPrice > 0) {
    return metaPrice;
  }

  const closePrices = result?.indicators?.quote?.[0]?.close;

  if (!Array.isArray(closePrices)) {
    return null;
  }

  for (let index = closePrices.length - 1; index >= 0; index -= 1) {
    if (Number.isFinite(closePrices[index]) && closePrices[index] > 0) {
      return closePrices[index];
    }
  }

  return null;
}

async function fetchYahooSparkBatch(symbols) {
  const payload = extractJsonPayload(await fetchText(buildYahooSparkProxyUrl(symbols)));
  const results = payload?.spark?.result;

  if (!Array.isArray(results)) {
    throw new Error("Yahoo Finance 응답 형식이 예상과 다릅니다.");
  }

  return results.reduce((quotes, result) => {
    const symbol = result?.symbol ?? result?.response?.[0]?.meta?.symbol;
    const price = getYahooSparkPrice(result);

    if (!symbol || !Number.isFinite(price)) {
      return quotes;
    }

    quotes[toTickerFromYahooSymbol(symbol)] = {
      price,
    };
    return quotes;
  }, {});
}

async function fetchYahooIntradayQuote(ticker, { range = "1d", interval = "1m" } = {}) {
  const symbol = toYahooSymbol(ticker);
  const payload = extractJsonPayload(
    await fetchText(buildYahooChartProxyUrl(symbol, { range, interval })),
  );
  const price = getYahooChartPrice(payload);

  if (!Number.isFinite(price) || price <= 0) {
    throw new Error(`${ticker} 실시간 값을 읽지 못했습니다.`);
  }

  return normalizeLivePriceSource(
    {
      provider: "Yahoo Finance via r.jina.ai",
      updatedAt: new Date().toISOString(),
      quotes: {
        [ticker]: {
          price,
        },
      },
    },
    "yahoo-chart-proxy",
  );
}

async function fetchLocalLivePrices() {
  const payload = await fetchJson(LIVE_PRICES_URL);
  return normalizeLivePriceSource(payload, "deployed-file");
}

async function fetchDividendData() {
  const payload = await fetchJson(DIVIDENDS_URL);

  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    throw new Error("Invalid dividend payload");
  }

  return payload;
}

async function fetchDividendCalendarData() {
  const payload = await fetchJson(DIVIDEND_CALENDAR_URL);

  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    throw new Error("Invalid dividend calendar payload");
  }

  return payload;
}

async function fetchGitHubLivePrices() {
  const payload = await fetchJson(GITHUB_LIVE_PRICES_URL, {
    headers: {
      accept: "application/vnd.github+json",
    },
  });

  if (typeof payload.content !== "string") {
    throw new Error("GitHub live price response did not include content");
  }

  return normalizeLivePriceSource(decodeBase64Json(payload.content), "github-main");
}

async function fetchYahooLivePrices(tickers) {
  const uniqueSymbols = [...new Set(tickers.map(toYahooSymbol))];

  if (!uniqueSymbols.length) {
    throw new Error("외부 시세를 조회할 종목이 없습니다.");
  }

  const batchResults = await Promise.allSettled(
    chunkArray(uniqueSymbols, YAHOO_SPARK_BATCH_SIZE).map((symbols) =>
      fetchYahooSparkBatch(symbols),
    ),
  );
  const quotes = {};
  let lastError = null;

  batchResults.forEach((result) => {
    if (result.status === "fulfilled") {
      Object.assign(quotes, result.value);
      return;
    }

    lastError ??= result.reason;
  });

  if (!countUsableQuotes({ quotes })) {
    throw lastError ?? new Error("Yahoo Finance에서 시세를 읽지 못했습니다.");
  }

  return normalizeLivePriceSource(
    {
      provider: "Yahoo Finance via r.jina.ai",
      updatedAt: new Date().toISOString(),
      quotes,
    },
    "yahoo-spark-proxy",
  );
}

async function loadLiveExchangeRateSource() {
  try {
    return await fetchYahooIntradayQuote("KRW=X");
  } catch (error) {
    return fetchYahooLivePrices(["KRW=X"]);
  }
}

async function loadLivePrices({ includeRepo = false, externalTickers = [] } = {}) {
  const requests = [
    () => fetchLocalLivePrices(),
    ...(externalTickers.length ? [() => fetchYahooLivePrices(externalTickers)] : []),
  ];
  const results = await Promise.allSettled(requests.map((request) => request()));
  const sources = [];
  let lastError = null;

  results.forEach((result) => {
    if (result.status === "fulfilled") {
      sources.push(result.value);
      return;
    }

    lastError ??= result.reason;
  });

  let selectedSource = sources.reduce(choosePreferredLivePriceSource, null);

  if (includeRepo && !isAutomaticLivePriceSource(selectedSource)) {
    try {
      selectedSource = choosePreferredLivePriceSource(
        selectedSource,
        await fetchGitHubLivePrices(),
      );
    } catch (error) {
      lastError ??= error;
    }
  }

  if (!selectedSource) {
    throw lastError ?? new Error("Failed to load live prices");
  }

  return selectedSource;
}

function renderCurrentBook() {
  if (!state.portfolioSource) {
    return;
  }

  renderApp(buildBookData(state.portfolioSource, state.livePriceSource, state.dividendSource));
}

function getUsdKrwRate(source, quotes) {
  const liveRate = quotes["KRW=X"]?.price;

  if (Number.isFinite(liveRate)) {
    return liveRate;
  }

  if (Number.isFinite(source?.snapshotUsdKrw)) {
    return source.snapshotUsdKrw;
  }

  return 1;
}

function convertToBaseCurrency(value, currency, usdKrwRate) {
  return currency === "USD" ? value * usdKrwRate : value;
}

function buildPortfolioData(source, livePriceSource, dividendSource) {
  const quotes = livePriceSource?.quotes ?? {};
  const isFallbackSource = livePriceSource?.provider === "snapshot-fallback";
  const hasAutomaticQuotes = isAutomaticLivePriceSource(livePriceSource);
  const usdKrwRate = getUsdKrwRate(source, quotes);
  const cashKrw = getCashKrw(source);
  const cashUsd = getCashUsd(source);
  const cashBase = cashKrw + convertToBaseCurrency(cashUsd, "USD", usdKrwRate);

  const enrichedHoldings = source.holdings.map((holding) => {
    const currency = getHoldingCurrency(holding);
    const averageCost = getHoldingAverageCost(holding);
    const averageCostCurrency = currency;
    const averageCostKrw = getHoldingAverageCostKrw(holding);
    const storedQuote = !holding.manualPriceOnly ? quotes[holding.ticker] : null;
    const shouldUseStoredQuote =
      hasUsableQuote(storedQuote) &&
      (hasAutomaticQuotes || !isFallbackSource || storedQuote.price !== holding.snapshotPrice);
    const currentPrice = shouldUseStoredQuote ? storedQuote.price : holding.snapshotPrice;
    const marketValue = currentPrice * holding.quantity;
    const profitLoss = marketValue - holding.costBasis;
    const marketValueBase = convertToBaseCurrency(marketValue, currency, usdKrwRate);
    const costBasisBase = Number.isFinite(averageCostKrw)
      ? averageCostKrw * holding.quantity
      : convertToBaseCurrency(holding.costBasis, currency, usdKrwRate);
    const profitLossBase = marketValueBase - costBasisBase;
    const annualDividendPerShare = getAnnualDividendPerShare(dividendSource, holding, currentPrice);
    const annualDividendIncome = annualDividendPerShare * holding.quantity;
    const annualDividendIncomeBase = convertToBaseCurrency(
      annualDividendIncome,
      currency,
      usdKrwRate,
    );
    const realizedProfitLossKrw = Number(holding.realizedProfitLossKrw ?? 0);
    const cumulativeDividendIncomeKrw = Number(holding.cumulativeDividendIncomeKrw ?? 0);
    const profitRate = holding.costBasis === 0 ? 0 : profitLoss / holding.costBasis;
    const profitRateBase = costBasisBase === 0 ? 0 : profitLossBase / costBasisBase;
    const totalReturnAmount = calculateTotalReturnAmount(
      profitLossBase,
      realizedProfitLossKrw,
      cumulativeDividendIncomeKrw,
    );
    const totalReturnRate = calculateTotalReturnRate(totalReturnAmount, costBasisBase);
    const useKrwProfitDisplay = Number.isFinite(averageCostKrw);

    return {
      ...holding,
      currency,
      usdKrwRate,
      currentPrice,
      marketValue,
      marketValueBase,
      costBasisBase,
      profitLoss,
      profitLossBase,
      annualDividendPerShare,
      annualDividendIncome,
      annualDividendIncomeBase,
      realizedProfitLossKrw,
      cumulativeDividendIncomeKrw,
      profitRate,
      profitRateDisplay: useKrwProfitDisplay ? profitRateBase : profitRate,
      profitLossDisplay: useKrwProfitDisplay ? profitLossBase : profitLoss,
      profitLossDisplayCurrency: useKrwProfitDisplay ? "KRW" : currency,
      totalReturnAmount,
      totalReturnRate,
      averageCost,
      averageCostCurrency,
      averageCostKrw,
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
      accumulator.marketValue += holding.marketValueBase;
      accumulator.costBasis += holding.costBasisBase;
      accumulator.profitLoss += holding.profitLossBase;
      accumulator.annualDividendIncome += holding.annualDividendIncomeBase;
      accumulator.realizedProfitLoss += holding.realizedProfitLossKrw;
      accumulator.cumulativeDividendIncome += holding.cumulativeDividendIncomeKrw;
      return accumulator;
    },
    {
      quantity: 0,
      marketValue: 0,
      costBasis: 0,
      profitLoss: 0,
      annualDividendIncome: 0,
      realizedProfitLoss: 0,
      cumulativeDividendIncome: 0,
    },
  );

  const totalAssets = totals.marketValue + cashBase;
  const investedWeight = totalAssets === 0 ? 0 : totals.marketValue / totalAssets;
  const cashWeight = totalAssets === 0 ? 0 : cashBase / totalAssets;
  const totalReturn = totals.costBasis === 0 ? 0 : totals.profitLoss / totals.costBasis;
  const totalReturnAmount = calculateTotalReturnAmount(
    totals.profitLoss,
    totals.realizedProfitLoss,
    totals.cumulativeDividendIncome,
  );
  const totalReturnRate = calculateTotalReturnRate(totalReturnAmount, totals.costBasis);

  const holdingsWithWeight = enrichedHoldings.map((holding) => ({
    ...holding,
    assetWeight: totalAssets === 0 ? 0 : holding.marketValueBase / totalAssets,
  }));
  const holdingCurrencies = [...new Set(holdingsWithWeight.map((holding) => holding.currency))];

  return {
    ...source,
    holdings: holdingsWithWeight,
    cash: cashKrw,
    cashUsd,
    cashBase,
    hasForeignCurrency:
      holdingCurrencies.some((currency) => currency !== "KRW") || cashUsd > 0,
    usdKrwRate,
    totals: {
      ...totals,
      totalAssets,
      investedWeight,
      cashWeight,
      totalReturn,
      totalReturnAmount,
      totalReturnRate,
    },
  };
}

function buildBookData(source, livePriceSource, dividendSource) {
  const portfolios = source.accounts
    .map((account) => buildPortfolioData(account, livePriceSource, dividendSource))
    .sort((left, right) => {
      const leftOrder = portfolioDisplayOrderLookup.get(left.id) ?? Number.MAX_SAFE_INTEGER;
      const rightOrder = portfolioDisplayOrderLookup.get(right.id) ?? Number.MAX_SAFE_INTEGER;

      return leftOrder - rightOrder || left.name.localeCompare(right.name, "ko-KR");
    });

  const overview = portfolios.reduce(
    (accumulator, portfolio) => {
      accumulator.totalAssets += portfolio.totals.totalAssets;
      accumulator.marketValue += portfolio.totals.marketValue;
      accumulator.costBasis += portfolio.totals.costBasis;
      accumulator.profitLoss += portfolio.totals.profitLoss;
      accumulator.cash += portfolio.cashBase;
      accumulator.cashKrw += portfolio.cash;
      accumulator.cashUsd += portfolio.cashUsd;
      accumulator.annualDividendIncome += portfolio.totals.annualDividendIncome;
      accumulator.realizedProfitLoss += portfolio.totals.realizedProfitLoss;
      accumulator.cumulativeDividendIncome += portfolio.totals.cumulativeDividendIncome;
      accumulator.holdingsCount += portfolio.holdings.length;
      return accumulator;
    },
    {
      totalAssets: 0,
      marketValue: 0,
      costBasis: 0,
      profitLoss: 0,
      cash: 0,
      cashKrw: 0,
      cashUsd: 0,
      annualDividendIncome: 0,
      realizedProfitLoss: 0,
      cumulativeDividendIncome: 0,
      holdingsCount: 0,
    },
  );

  overview.totalReturn =
    overview.costBasis === 0 ? 0 : overview.profitLoss / overview.costBasis;
  overview.totalReturnAmount = calculateTotalReturnAmount(
    overview.profitLoss,
    overview.realizedProfitLoss,
    overview.cumulativeDividendIncome,
  );
  overview.totalReturnRate = calculateTotalReturnRate(
    overview.totalReturnAmount,
    overview.costBasis,
  );

  return {
    meta: source.meta,
    livePriceSource,
    hasForeignCurrency: portfolios.some((portfolio) => portfolio.hasForeignCurrency),
    portfolios,
    overview,
  };
}

function getBookUsdKrwRate(book) {
  const liveRate = book.livePriceSource?.quotes?.["KRW=X"]?.price;

  if (Number.isFinite(liveRate) && liveRate > 0) {
    return liveRate;
  }

  const portfolioRate = book.portfolios
    .map((portfolio) => portfolio.usdKrwRate)
    .find((rate) => Number.isFinite(rate) && rate > 0);

  return portfolioRate ?? 0;
}

function getDividendCalendarExchangeRateMeta(book, exchangeRateSource) {
  const liveRate = exchangeRateSource?.quotes?.["KRW=X"]?.price;

  if (Number.isFinite(liveRate) && liveRate > 0) {
    return {
      usdKrwRate: liveRate,
      exchangeRateProvider: exchangeRateSource.provider ?? null,
      exchangeRateUpdatedAt: exchangeRateSource.updatedAt ?? null,
      usesLiveExchangeRate: true,
    };
  }

  return {
    usdKrwRate: getBookUsdKrwRate(book),
    exchangeRateProvider: book.livePriceSource?.provider ?? null,
    exchangeRateUpdatedAt: book.livePriceSource?.updatedAt ?? null,
    usesLiveExchangeRate: isAutomaticLivePriceSource(book.livePriceSource),
  };
}

function buildDividendCalendarData(book, source, error = null, exchangeRateSource = null) {
  const holdingLookup = new Map();

  book.portfolios.forEach((portfolio) => {
    portfolio.holdings.forEach((holding) => {
      holdingLookup.set(`${portfolio.id}:${holding.ticker}`, {
        accountName: portfolio.name,
        ticker: holding.ticker,
        name: holding.name,
        currency: holding.currency,
        quantity: holding.quantity,
        currentPrice: holding.currentPrice,
      });
    });
  });

  const projectionYear = getDividendCalendarProjectionYear(source);
  const months = MONTH_LABELS.map((label, index) => ({
    month: index + 1,
    label,
    totals: {
      KRW: 0,
      USD: 0,
    },
    items: {
      KRW: [],
      USD: [],
    },
  }));
  const entries = Array.isArray(source?.entries) ? source.entries : [];
  let populatedEntryCount = 0;

  entries.forEach((entry) => {
    const reference = holdingLookup.get(`${entry?.accountId ?? ""}:${entry?.ticker ?? ""}`) ?? null;
    const currency =
      entry?.currency === "USD" || reference?.currency === "USD"
        ? "USD"
        : "KRW";
    const accountName =
      typeof entry?.accountName === "string" && entry.accountName.trim()
        ? entry.accountName.trim()
        : reference?.accountName ?? entry?.accountId ?? "계좌 미지정";
    const name =
      typeof entry?.name === "string" && entry.name.trim()
        ? entry.name.trim()
        : reference?.name ?? entry?.ticker ?? "미지정 종목";
    const monthlyAmounts = getDividendCalendarMonthlyAmounts(entry, reference, projectionYear);
    let hasValue = false;

    months.forEach((monthData, index) => {
      const rawAmount = monthlyAmounts[index + 1] ?? monthlyAmounts[String(index + 1)] ?? 0;
      const amount = Number(rawAmount);

      if (!Number.isFinite(amount) || amount === 0) {
        return;
      }

      monthData.totals[currency] += amount;
      monthData.items[currency].push({
        accountName,
        name,
        ticker: entry?.ticker ?? null,
        amount,
      });
      hasValue = true;
    });

    if (hasValue) {
      populatedEntryCount += 1;
    }
  });

  const totals = months.reduce(
    (accumulator, monthData) => {
      accumulator.KRW += monthData.totals.KRW;
      accumulator.USD += monthData.totals.USD;
      return accumulator;
    },
    {
      KRW: 0,
      USD: 0,
    },
  );
  const maxByCurrency = months.reduce(
    (accumulator, monthData) => ({
      KRW: Math.max(accumulator.KRW, monthData.totals.KRW),
      USD: Math.max(accumulator.USD, monthData.totals.USD),
    }),
    {
      KRW: 0,
      USD: 0,
    },
  );
  const exchangeRateMeta = getDividendCalendarExchangeRateMeta(book, exchangeRateSource);
  const usdKrwRate = exchangeRateMeta.usdKrwRate;
  const sharedMaxKrw = months.reduce(
    (maximum, monthData) =>
      Math.max(maximum, monthData.totals.KRW + monthData.totals.USD * usdKrwRate),
    0,
  );
  const sharedMaxUsd = usdKrwRate > 0 ? sharedMaxKrw / usdKrwRate : 0;
  const convertedUsdTotalKrw = totals.USD * usdKrwRate;

  return {
    year: projectionYear,
    updatedAt: source?.updatedAt ?? null,
    methodology: source?.methodology ?? "",
    months,
    totals,
    maxByCurrency,
    usdKrwRate,
    exchangeRateProvider: exchangeRateMeta.exchangeRateProvider,
    exchangeRateUpdatedAt: exchangeRateMeta.exchangeRateUpdatedAt,
    usesLiveExchangeRate: exchangeRateMeta.usesLiveExchangeRate,
    sharedMaxKrw,
    sharedMaxUsd,
    convertedUsdTotalKrw,
    totalBaseKrw: totals.KRW + convertedUsdTotalKrw,
    entryCount: entries.length,
    populatedEntryCount,
    hasData: totals.KRW > 0 || totals.USD > 0,
    error,
  };
}

function resolveAccountSummaryGroup(portfolio) {
  const name = portfolio?.name ?? "";

  if (portfolio?.id === "direct-investment" || name === "직투계좌") {
    return "direct-investment";
  }

  if (portfolio?.id === "isa" || name === "ISA") {
    return "isa";
  }

  if (portfolio?.id === "pension-savings-2" || name === "연금저축(세액미공제)") {
    return "pension-savings-non-deductible";
  }

  if (portfolio?.id?.startsWith("pension-savings") || name.startsWith("연금저축")) {
    return "pension-savings-deductible";
  }

  if (portfolio?.id?.startsWith("retirement-pension") || name.startsWith("퇴직연금")) {
    return "retirement-pension";
  }

  return null;
}

function buildAccountSummaryGroups(book) {
  const groups = ACCOUNT_SUMMARY_GROUPS.map((group) => ({
    ...group,
    totalAssets: 0,
    costBasis: 0,
    investedCostBasis: 0,
    profitLoss: 0,
    totalReturnAmount: 0,
    annualDividendIncome: 0,
    realizedProfitLoss: 0,
    cumulativeDividendIncome: 0,
    accounts: [],
  }));
  const groupLookup = new Map(groups.map((group) => [group.key, group]));

  book.portfolios.forEach((portfolio) => {
    const groupKey = resolveAccountSummaryGroup(portfolio);

    if (!groupKey || !groupLookup.has(groupKey)) {
      return;
    }

    const group = groupLookup.get(groupKey);
    group.totalAssets += portfolio.totals.totalAssets;
    group.costBasis += portfolio.totals.costBasis + portfolio.cashBase;
    group.investedCostBasis += portfolio.totals.costBasis;
    group.profitLoss += portfolio.totals.profitLoss;
    group.totalReturnAmount += portfolio.totals.totalReturnAmount;
    group.annualDividendIncome += portfolio.totals.annualDividendIncome;
    group.realizedProfitLoss += portfolio.totals.realizedProfitLoss;
    group.cumulativeDividendIncome += portfolio.totals.cumulativeDividendIncome;
    group.accounts.push(portfolio.name);
  });

  const maxAssets = groups.reduce(
    (maximum, group) => Math.max(maximum, group.totalAssets),
    0,
  );

  return groups.map((group) => ({
    ...group,
    accountsLabel: group.accounts.join(" + "),
    dividendTax: group.key === "direct-investment" ? group.annualDividendIncome * 0.15 : 0,
    netDividendIncome:
      group.annualDividendIncome -
      (group.key === "direct-investment" ? group.annualDividendIncome * 0.15 : 0),
    dividendYield: group.totalAssets === 0 ? 0 : group.annualDividendIncome / group.totalAssets,
    monthlyNetDividendIncome:
      (group.annualDividendIncome -
        (group.key === "direct-investment" ? group.annualDividendIncome * 0.15 : 0)) / 12,
    profitRate: group.costBasis === 0 ? 0 : group.profitLoss / group.costBasis,
    totalReturnRate:
      group.investedCostBasis === 0 ? 0 : group.totalReturnAmount / group.investedCostBasis,
    weight: book.overview.totalAssets === 0 ? 0 : group.totalAssets / book.overview.totalAssets,
    barScale: maxAssets === 0 ? 0 : group.totalAssets / maxAssets,
  }));
}

function getTreemapCategoryOrder(superCategory, category) {
  const superOrder = TREEMAP_LAYOUT_ORDER_BY_SUPER[superCategory] ?? [];
  const superIndex = superOrder.indexOf(category);

  if (superIndex !== -1) {
    return superIndex;
  }

  if (treemapCategoryOrderLookup.has(category)) {
    return superOrder.length + treemapCategoryOrderLookup.get(category);
  }

  return superOrder.length + TREEMAP_LAYOUT_ORDER_GLOBAL.length;
}

function resolveTreemapCategory(asset) {
  const candidates = [asset.category, asset.name, asset.ticker, asset.type]
    .map(normalizeTreemapToken)
    .filter(Boolean);

  for (const candidate of candidates) {
    if (treemapAliasLookup.has(candidate)) {
      return treemapAliasLookup.get(candidate);
    }

    if (treemapCategoryLookup.has(candidate)) {
      return treemapCategoryLookup.get(candidate);
    }
  }

  const fullText = normalizeTreemapToken(
    [asset.name, asset.type, asset.ticker].filter(Boolean).join(" "),
  );

  if (asset.ticker === "KRW" || fullText.includes("원화") || fullText.includes("고유계정대")) {
    return "원화";
  }

  if (asset.ticker === "USD" || fullText.includes("현금(usd)")) {
    return "USD";
  }

  if (fullText.includes("머니마켓")) {
    return "MM";
  }

  if (fullText.includes("인프라")) {
    return "인프라";
  }

  if (fullText.includes("텔레콤") || fullText.includes("통신")) {
    return "통신";
  }

  if (/(금융|은행|증권|보험)/.test(fullText)) {
    return "금융";
  }

  if (/(현대차|자동차|제조)/.test(fullText)) {
    return "제조";
  }

  return "기타";
}

function resolveTreemapSuperCategory(asset, category) {
  if (asset.superCategory && TREEMAP_LAYOUT_ORDER_BY_SUPER[asset.superCategory]) {
    return asset.superCategory;
  }

  if (treemapSuperCategoryLookup.has(category)) {
    return treemapSuperCategoryLookup.get(category);
  }

  return asset.currency === "USD" ? "해외" : "국내";
}

function buildTreemapData(book) {
  if (!book.overview.totalAssets) {
    return null;
  }

  const assetMap = new Map();

  book.portfolios.forEach((portfolio) => {
    portfolio.holdings.forEach((holding) => {
      const key = `${holding.currency}:${holding.ticker}`;
      const existingAsset = assetMap.get(key) ?? {
        ticker: holding.ticker,
        name: holding.name,
        type: holding.type,
        currency: holding.currency,
        marketValueBase: 0,
        costBasisBase: 0,
        profitLossBase: 0,
        annualDividendIncomeBase: 0,
        realizedProfitLossKrw: 0,
        cumulativeDividendIncomeKrw: 0,
        accountNames: new Set(),
      };

      existingAsset.marketValueBase += holding.marketValueBase;
      existingAsset.costBasisBase += holding.costBasisBase;
      existingAsset.profitLossBase += holding.profitLossBase;
      existingAsset.annualDividendIncomeBase += holding.annualDividendIncomeBase;
      existingAsset.realizedProfitLossKrw += holding.realizedProfitLossKrw;
      existingAsset.cumulativeDividendIncomeKrw += holding.cumulativeDividendIncomeKrw;
      existingAsset.accountNames.add(portfolio.name);
      assetMap.set(key, existingAsset);
    });

    if (portfolio.cash > 0) {
      const key = "KRW:KRW";
      const existingCash = assetMap.get(key) ?? {
        ticker: "KRW",
        name: "원화 현금",
        type: "현금",
        currency: "KRW",
        marketValueBase: 0,
        costBasisBase: 0,
        profitLossBase: 0,
        annualDividendIncomeBase: 0,
        realizedProfitLossKrw: 0,
        cumulativeDividendIncomeKrw: 0,
        accountNames: new Set(),
      };

      existingCash.marketValueBase += portfolio.cash;
      existingCash.costBasisBase += portfolio.cash;
      existingCash.accountNames.add(portfolio.name);
      assetMap.set(key, existingCash);
    }

    if (portfolio.cashUsd > 0) {
      const key = "USD:USD";
      const existingUsdCash = assetMap.get(key) ?? {
        ticker: "USD",
        name: "현금(USD)",
        type: "현금",
        currency: "USD",
        marketValueBase: 0,
        costBasisBase: 0,
        profitLossBase: 0,
        annualDividendIncomeBase: 0,
        realizedProfitLossKrw: 0,
        cumulativeDividendIncomeKrw: 0,
        accountNames: new Set(),
      };

      const usdCashBase = convertToBaseCurrency(
        portfolio.cashUsd,
        "USD",
        portfolio.usdKrwRate,
      );
      existingUsdCash.marketValueBase += usdCashBase;
      existingUsdCash.costBasisBase += usdCashBase;
      existingUsdCash.accountNames.add(portfolio.name);
      assetMap.set(key, existingUsdCash);
    }
  });

  const assets = [...assetMap.values()]
    .filter((asset) => asset.marketValueBase > 0)
    .map((asset) => {
      const category = resolveTreemapCategory(asset);
      const superCategory = resolveTreemapSuperCategory(asset, category);

      return {
        ...asset,
        accountCount: asset.accountNames.size,
        weight: asset.marketValueBase / book.overview.totalAssets,
        category,
        superCategory,
      };
    });

  const superCategoryMap = new Map();

  assets.forEach((asset) => {
    const existingSuperCategory = superCategoryMap.get(asset.superCategory) ?? {
      id: normalizeTreemapToken(asset.superCategory).replace(/\s+/g, "-"),
      name: asset.superCategory,
      marketValueBase: 0,
      costBasisBase: 0,
      profitLossBase: 0,
      annualDividendIncomeBase: 0,
      realizedProfitLossKrw: 0,
      cumulativeDividendIncomeKrw: 0,
      categories: new Map(),
    };
    const existingCategory = existingSuperCategory.categories.get(asset.category) ?? {
      id: `${existingSuperCategory.id}-${normalizeTreemapToken(asset.category).replace(/\s+/g, "-")}`,
      name: asset.category,
      superCategory: asset.superCategory,
      marketValueBase: 0,
      costBasisBase: 0,
      profitLossBase: 0,
      annualDividendIncomeBase: 0,
      realizedProfitLossKrw: 0,
      cumulativeDividendIncomeKrw: 0,
      assets: [],
    };

    existingCategory.marketValueBase += asset.marketValueBase;
    existingCategory.costBasisBase += asset.costBasisBase;
    existingCategory.profitLossBase += asset.profitLossBase;
    existingCategory.annualDividendIncomeBase += asset.annualDividendIncomeBase;
    existingCategory.realizedProfitLossKrw += asset.realizedProfitLossKrw;
    existingCategory.cumulativeDividendIncomeKrw += asset.cumulativeDividendIncomeKrw;
    existingCategory.assets.push(asset);
    existingSuperCategory.marketValueBase += asset.marketValueBase;
    existingSuperCategory.costBasisBase += asset.costBasisBase;
    existingSuperCategory.profitLossBase += asset.profitLossBase;
    existingSuperCategory.annualDividendIncomeBase += asset.annualDividendIncomeBase;
    existingSuperCategory.realizedProfitLossKrw += asset.realizedProfitLossKrw;
    existingSuperCategory.cumulativeDividendIncomeKrw += asset.cumulativeDividendIncomeKrw;
    existingSuperCategory.categories.set(asset.category, existingCategory);
    superCategoryMap.set(asset.superCategory, existingSuperCategory);
  });

  const superCategories = [...superCategoryMap.values()]
    .map((superCategory) => {
      const categories = [...superCategory.categories.values()]
        .map((category) => ({
          ...category,
          weight: category.marketValueBase / book.overview.totalAssets,
          assets: [...category.assets].sort(
            (left, right) =>
              right.marketValueBase - left.marketValueBase ||
              left.name.localeCompare(right.name, "ko-KR"),
          ),
        }))
        .sort(
          (left, right) =>
            getTreemapCategoryOrder(superCategory.name, left.name) -
              getTreemapCategoryOrder(superCategory.name, right.name) ||
            right.marketValueBase - left.marketValueBase,
        );

      return {
        id: superCategory.id,
        name: superCategory.name,
        marketValueBase: superCategory.marketValueBase,
        annualDividendIncomeBase: superCategory.annualDividendIncomeBase,
        weight: superCategory.marketValueBase / book.overview.totalAssets,
        categories,
      };
    })
    .sort((left, right) => {
      const leftIndex = TREEMAP_SUPER_CATEGORY_ORDER.indexOf(left.name);
      const rightIndex = TREEMAP_SUPER_CATEGORY_ORDER.indexOf(right.name);

      if (leftIndex !== -1 || rightIndex !== -1) {
        const normalizedLeftIndex = leftIndex === -1 ? Number.MAX_SAFE_INTEGER : leftIndex;
        const normalizedRightIndex = rightIndex === -1 ? Number.MAX_SAFE_INTEGER : rightIndex;
        return normalizedLeftIndex - normalizedRightIndex;
      }

      return right.marketValueBase - left.marketValueBase;
    });

  return {
    totalAssets: book.overview.totalAssets,
    assetCount: assets.length,
    superCategories,
  };
}

function getTreemapAssetLabel(asset) {
  const ticker = String(asset.ticker ?? "").trim();

  if (asset.type !== "현금" && /^[A-Z]{1,6}$/.test(ticker)) {
    return ticker;
  }

  return asset.name;
}

function getTreemapNodeLabel(node) {
  if (node?.assets || node?.categories) {
    return node.name;
  }

  return getTreemapAssetLabel(node);
}

function getTreemapLabelUnits(label) {
  return [...String(label ?? "")].reduce((sum, character) => {
    if (/[가-힣]/.test(character)) {
      return sum + 1;
    }

    if (/[A-Z]/.test(character)) {
      return sum + 0.66;
    }

    if (/[a-z]/.test(character)) {
      return sum + 0.58;
    }

    if (/[0-9]/.test(character)) {
      return sum + 0.52;
    }

    return sum + 0.34;
  }, 0);
}

function getTreemapGroupTextPenalty(items, rect) {
  if (!items.length) {
    return 0;
  }

  const totalValue = items.reduce((sum, item) => sum + item.value, 0);
  const weightedUnits = items.reduce((sum, item) => {
    const labelUnits = getTreemapLabelUnits(getTreemapNodeLabel(item.node));
    return sum + labelUnits * (item.value / totalValue);
  }, 0);
  const longestUnits = Math.max(
    ...items.map((item) => getTreemapLabelUnits(getTreemapNodeLabel(item.node))),
  );
  const minSide = Math.min(rect.width, rect.height);
  const desiredWidth = Math.min(30, 6 + longestUnits * 1.12);
  const desiredShortSide = Math.min(20, 4.6 + weightedUnits * 0.46);

  return (
    Math.max(0, desiredWidth - rect.width) * 0.11 +
    Math.max(0, desiredShortSide - minSide) * 0.18
  );
}

function getTreemapCanvasAspectRatio() {
  const viewportWidth = window.innerWidth || 1280;

  if (viewportWidth <= 640) {
    return TREEMAP_CANVAS_ASPECT_MOBILE;
  }

  if (viewportWidth <= 960) {
    return TREEMAP_CANVAS_ASPECT_TABLET;
  }

  return TREEMAP_CANVAS_ASPECT_DESKTOP;
}

function getTreemapLayoutRect(width, height) {
  return {
    x: 0,
    y: 0,
    width,
    height,
  };
}

function encodeTreemapVariants(variants) {
  return variants.map((variant) => encodeURIComponent(variant)).join("|");
}

function encodeTreemapHoverValue(value) {
  return encodeURIComponent(String(value ?? ""));
}

function decodeTreemapVariants(value) {
  return String(value ?? "")
    .split("|")
    .map((variant) => {
      try {
        return decodeURIComponent(variant);
      } catch (error) {
        return variant;
      }
    })
    .filter(Boolean);
}

function decodeTreemapHoverValue(value) {
  try {
    return decodeURIComponent(String(value ?? ""));
  } catch (error) {
    return String(value ?? "");
  }
}

function doesTreemapTileCopyFit(tileInner, copy, value) {
  const tileStyle = window.getComputedStyle(tileInner);
  const paddingBlock =
    Number.parseFloat(tileStyle.paddingTop) +
    Number.parseFloat(tileStyle.paddingBottom);
  const availableHeight = tileInner.clientHeight - paddingBlock - 1;
  const copyRect = copy.getBoundingClientRect();
  const heightFits = copyRect.height <= availableHeight;
  const valueFits =
    !value ||
    value.offsetParent === null ||
    value.scrollWidth <= value.clientWidth + 1;

  return heightFits && valueFits;
}

function fitTreemapTileTypography(tileInner) {
  const copy = tileInner.querySelector(".treemap-tile-copy");
  const name = tileInner.querySelector(".treemap-tile-name");
  const value = tileInner.querySelector(".treemap-tile-value");
  const valueVariants = decodeTreemapVariants(value?.dataset.valueVariants).slice(
    0,
    4,
  );

  if (!copy || !name) {
    return;
  }

  const applyScale = (scale) => {
    tileInner.style.setProperty("--treemap-fit-scale", scale.toFixed(3));
  };
  const fullLabel = name.textContent ?? "";
  const applyContent = (valueText = "") => {
    name.textContent = fullLabel;

    if (value) {
      value.textContent = valueText;
    }
  };
  const runBinaryFit = () => {
    let low = 0.38;
    let high = 1;

    for (let iteration = 0; iteration < 10; iteration += 1) {
      const mid = (low + high) / 2;
      applyScale(mid);

      if (doesTreemapTileCopyFit(tileInner, copy, value)) {
        low = mid;
      } else {
        high = mid;
      }
    }

    applyScale(low);
    return {
      fits: doesTreemapTileCopyFit(tileInner, copy, value),
      scale: low,
    };
  };
  const tryVariant = (valueText, hideValue) => {
    tileInner.classList.toggle("treemap-fit-compact", hideValue);
    applyContent(valueText);
    applyScale(1);

    if (doesTreemapTileCopyFit(tileInner, copy, value)) {
      return {
        fits: true,
        scale: 1,
      };
    }

    return runBinaryFit();
  };
  const readableScaleThreshold = 0.68;
  let bestFit = null;

  valueVariants.forEach((valueText, valueIndex) => {
    const result = tryVariant(valueText, false);

    if (!result.fits) {
      return;
    }

    if (result.scale >= readableScaleThreshold) {
      bestFit = {
        valueText,
        hideValue: false,
        scale: result.scale,
      };
      return;
    }

    if (
      !bestFit ||
      result.scale - valueIndex * 0.04 > bestFit.scale - (bestFit.hideValue ? 0.08 : 0)
    ) {
      bestFit = {
        valueText,
        hideValue: false,
        scale: result.scale,
      };
    }
  });

  if (bestFit) {
    tileInner.classList.toggle("treemap-fit-compact", bestFit.hideValue);
    applyContent(bestFit.valueText);
    applyScale(bestFit.scale);
    return;
  }

  const compactResult = tryVariant("", true);

  if (compactResult.fits) {
    tileInner.classList.add("treemap-fit-compact");
    applyContent("");
    applyScale(compactResult.scale);
    return;
  }

  const fallbackValue = valueVariants.at(-1) ?? value?.textContent ?? "";
  const finalFit = bestFit ?? {
    valueText: fallbackValue,
    hideValue: true,
    scale: 0.38,
  };

  tileInner.classList.toggle("treemap-fit-compact", finalFit.hideValue);
  applyContent(finalFit.valueText);
  applyScale(finalFit.scale);
}

function fitTreemapTypography() {
  document.querySelectorAll(".treemap-tile-inner").forEach((tileInner) => {
    fitTreemapTileTypography(tileInner);
  });
}

function scheduleTreemapTypographyFit() {
  if (treemapTypographyFrameId) {
    window.cancelAnimationFrame(treemapTypographyFrameId);
  }

  treemapTypographyFrameId = window.requestAnimationFrame(() => {
    treemapTypographyFrameId = 0;
    fitTreemapTypography();
  });
}

function getTreemapRectAspectRatio(rect) {
  const shortSide = Math.max(Math.min(rect.width, rect.height), 1);
  const longSide = Math.max(rect.width, rect.height);

  return longSide / shortSide;
}

function getTreemapSplitPenalty(
  rect,
  firstRatio,
  splitVertically,
  firstItems,
  secondItems,
) {
  const firstRect = splitVertically
    ? {
        width: rect.width * firstRatio,
        height: rect.height,
      }
    : {
        width: rect.width,
        height: rect.height * firstRatio,
      };
  const secondRect = splitVertically
    ? {
        width: rect.width - firstRect.width,
        height: rect.height,
      }
    : {
        width: rect.width,
        height: rect.height - firstRect.height,
      };
  const worstAspectRatio = Math.max(
    getTreemapRectAspectRatio(firstRect),
    getTreemapRectAspectRatio(secondRect),
  );
  const smallestSide = Math.min(
    firstRect.width,
    firstRect.height,
    secondRect.width,
    secondRect.height,
  );
  const textPenalty =
    getTreemapGroupTextPenalty(firstItems, firstRect) +
    getTreemapGroupTextPenalty(secondItems, secondRect);

  return (
    worstAspectRatio +
    (smallestSide < 7 ? (7 - smallestSide) * 0.45 : 0) +
    textPenalty
  );
}

function findTreemapSplit(items, rect) {
  const totalValue = items.reduce((sum, item) => sum + item.value, 0);
  const preferredOrientationIsVertical = rect.width >= rect.height;
  let runningValue = 0;
  let bestSplit = {
    index: 1,
    splitVertically: preferredOrientationIsVertical,
    score: Number.POSITIVE_INFINITY,
  };

  for (let index = 1; index < items.length; index += 1) {
    runningValue += items[index - 1].value;
    const firstRatio = runningValue / totalValue;
    const balancePenalty = Math.abs(0.5 - firstRatio) * 0.8;

    [true, false].forEach((splitVertically) => {
      const orientationPenalty =
        splitVertically === preferredOrientationIsVertical ? 0 : 0.18;
      const score =
        getTreemapSplitPenalty(
          rect,
          firstRatio,
          splitVertically,
          items.slice(0, index),
          items.slice(index),
        ) +
        balancePenalty +
        orientationPenalty;

      if (score < bestSplit.score) {
        bestSplit = {
          index,
          splitVertically,
          score,
        };
      }
    });
  }

  return bestSplit;
}

function buildTreemapLayouts(
  items,
  rect = { x: 0, y: 0, width: 100, height: 100 },
  space = { width: rect.width, height: rect.height },
) {
  const filteredItems = items.filter((item) => item.value > 0);

  if (!filteredItems.length) {
    return [];
  }

  if (filteredItems.length === 1) {
    return [
      {
        node: filteredItems[0].node,
        x: rect.x,
        y: rect.y,
        width: rect.width,
        height: rect.height,
        spaceWidth: space.width,
        spaceHeight: space.height,
      },
    ];
  }

  const totalValue = filteredItems.reduce((sum, item) => sum + item.value, 0);
  const split = findTreemapSplit(filteredItems, rect);
  const splitIndex = split.index;
  const firstItems = filteredItems.slice(0, splitIndex);
  const secondItems = filteredItems.slice(splitIndex);
  const firstValue = firstItems.reduce((sum, item) => sum + item.value, 0);
  const firstRatio = firstValue / totalValue;
  const { splitVertically } = split;
  const firstRect = splitVertically
    ? {
        x: rect.x,
        y: rect.y,
        width: rect.width * firstRatio,
        height: rect.height,
      }
    : {
        x: rect.x,
        y: rect.y,
        width: rect.width,
        height: rect.height * firstRatio,
      };
  const secondRect = splitVertically
    ? {
        x: rect.x + firstRect.width,
        y: rect.y,
        width: rect.width - firstRect.width,
        height: rect.height,
      }
    : {
        x: rect.x,
        y: rect.y + firstRect.height,
        width: rect.width,
        height: rect.height - firstRect.height,
      };

  return [
    ...buildTreemapLayouts(firstItems, firstRect, space),
    ...buildTreemapLayouts(secondItems, secondRect, space),
  ];
}

function hexToRgba(hex, alpha) {
  const { red, green, blue } = hexToRgb(hex);

  return `rgba(${red}, ${green}, ${blue}, ${alpha})`;
}

function hexToRgb(hex) {
  const normalizedHex = hex.replace("#", "");

  return {
    red: Number.parseInt(normalizedHex.slice(0, 2), 16),
    green: Number.parseInt(normalizedHex.slice(2, 4), 16),
    blue: Number.parseInt(normalizedHex.slice(4, 6), 16),
  };
}

function rgbToHex(red, green, blue) {
  return `#${[red, green, blue]
    .map((channel) => Math.max(0, Math.min(255, Math.round(channel))).toString(16).padStart(2, "0"))
    .join("")}`;
}

function mixHexColors(firstHex, secondHex, firstWeight = 0.5) {
  const normalizedWeight = Math.max(0, Math.min(1, firstWeight));
  const secondWeight = 1 - normalizedWeight;
  const first = hexToRgb(firstHex);
  const second = hexToRgb(secondHex);

  return rgbToHex(
    first.red * normalizedWeight + second.red * secondWeight,
    first.green * normalizedWeight + second.green * secondWeight,
    first.blue * normalizedWeight + second.blue * secondWeight,
  );
}

function getTreemapAccent(category, superCategory) {
  return (
    TREEMAP_CATEGORY_COLORS[category] ??
    (superCategory === "해외" ? "#7a97ca" : "#7da182")
  );
}

function buildTreemapStyle(layout, accent) {
  const fill = mixHexColors(accent, "#d7dbe2", 0.56);
  const fillSoft = mixHexColors(accent, "#e7eaee", 0.28);
  const shell = mixHexColors(accent, "#ccd1d8", 0.12);
  const spaceWidth = layout.spaceWidth ?? 100;
  const spaceHeight = layout.spaceHeight ?? 100;

  return [
    `left:${((layout.x / spaceWidth) * 100).toFixed(4)}%`,
    `top:${((layout.y / spaceHeight) * 100).toFixed(4)}%`,
    `width:${((layout.width / spaceWidth) * 100).toFixed(4)}%`,
    `height:${((layout.height / spaceHeight) * 100).toFixed(4)}%`,
    `--treemap-accent:${accent}`,
    `--treemap-shell:${shell}`,
    `--treemap-fill:${fill}`,
    `--treemap-fill-soft:${fillSoft}`,
    `--treemap-outline:${hexToRgba(accent, 0.18)}`,
    `--treemap-text:#434957`,
    `--treemap-text-soft:#656c7b`,
  ].join(";");
}

function getTreemapTileSizeClass(layout) {
  const area = layout.width * layout.height;
  const minSide = Math.min(layout.width, layout.height);
  const maxSide = Math.max(layout.width, layout.height);
  const aspectRatio = maxSide / Math.max(minSide, 1);
  const classes = [];

  if (area >= 1500 && minSide >= 22) {
    classes.push("tile-xl");
  } else if (area >= 720 && minSide >= 16) {
    classes.push("tile-lg");
  } else if (area >= 260 && minSide >= 11) {
    classes.push("tile-md");
  } else if (area >= 120 && minSide >= 8) {
    classes.push("tile-sm");
  } else {
    classes.push("tile-xs");
  }

  if (aspectRatio >= 3.2 || minSide < 9) {
    classes.push("tile-tight");
  }

  if (aspectRatio >= 5 || minSide < 6 || area < 60) {
    classes.push("tile-micro");
  }

  return classes.join(" ");
}

function hasTreemapSizeClass(sizeClass, className) {
  return sizeClass.split(/\s+/).includes(className);
}

function formatTreemapValueVariants(asset) {
  return [formatTreemapPercent(asset.weight)];
}

function formatTreemapCategoryLabel(category, sizeClass) {
  if (
    hasTreemapSizeClass(sizeClass, "tile-tight") ||
    hasTreemapSizeClass(sizeClass, "tile-micro") ||
    hasTreemapSizeClass(sizeClass, "tile-xs")
  ) {
    return `${category.name} / ${formatPercent(category.weight)}`;
  }

  if (hasTreemapSizeClass(sizeClass, "tile-sm")) {
    return `${category.name} / ${formatTreemapCompactAmount(category.marketValueBase)}`;
  }

  return `${category.name} / ${formatPercent(category.weight)} / ${formatTreemapAmount(category.marketValueBase)}`;
}

function formatTreemapSuperLabel(superCategory, sizeClass) {
  if (
    hasTreemapSizeClass(sizeClass, "tile-tight") ||
    hasTreemapSizeClass(sizeClass, "tile-xs")
  ) {
    return `${superCategory.name} / ${formatPercent(superCategory.weight)}`;
  }

  return `${superCategory.name} / ${formatPercent(superCategory.weight)} / ${formatTreemapAmount(superCategory.marketValueBase)}`;
}

function renderTreemapAssetTile(layout) {
  const asset = layout.node;
  const sizeClass = getTreemapTileSizeClass(layout);
  const accent = getTreemapAccent(asset.category, asset.superCategory);
  const label = getTreemapAssetLabel(asset);
  const valueVariants = formatTreemapValueVariants(asset);
  const hoverName = encodeTreemapHoverValue(asset.name);
  const hoverValue = encodeTreemapHoverValue(formatCurrency(asset.marketValueBase));
  const hoverWeight = encodeTreemapHoverValue(formatTreemapPercent(asset.weight));

  return `
    <article
      class="treemap-tile ${sizeClass}"
      style="${buildTreemapStyle(layout, accent)}"
      data-treemap-hover="asset"
      data-treemap-name="${hoverName}"
      data-treemap-value="${hoverValue}"
      data-treemap-weight="${hoverWeight}"
    >
      <div class="treemap-tile-inner">
        <div class="treemap-tile-copy">
          <strong class="treemap-tile-name">${label}</strong>
          <span class="treemap-tile-value" data-value-variants="${encodeTreemapVariants(valueVariants)}">${valueVariants[0]}</span>
        </div>
      </div>
    </article>
  `;
}

function renderTreemapCategoryTile(layout) {
  const category = layout.node;
  const sizeClass = getTreemapTileSizeClass(layout);
  const accent = getTreemapAccent(category.name, category.superCategory);
  const assetLayouts = buildTreemapLayouts(
    category.assets.map((asset) => ({
      node: asset,
      value: asset.marketValueBase,
    })),
    getTreemapLayoutRect(layout.width, layout.height),
  );

  return `
    <article class="treemap-category ${sizeClass}" style="${buildTreemapStyle(layout, accent)}">
      <div class="treemap-category-head">
        <p class="treemap-category-label">${formatTreemapCategoryLabel(category, sizeClass)}</p>
      </div>
      <div class="treemap-category-body">
        ${assetLayouts.map(renderTreemapAssetTile).join("")}
      </div>
    </article>
  `;
}

function renderTreemapSuperCategoryTile(layout) {
  const superCategory = layout.node;
  const sizeClass = getTreemapTileSizeClass(layout);
  const accent = getTreemapAccent(superCategory.name, superCategory.name);
  const categoryLayouts = buildTreemapLayouts(
    superCategory.categories.map((category) => ({
      node: category,
      value: category.marketValueBase,
    })),
    getTreemapLayoutRect(layout.width, layout.height),
  );

  return `
    <section class="treemap-super ${sizeClass}" style="${buildTreemapStyle(layout, accent)}">
      <div class="treemap-super-head">
        <p class="treemap-super-label">${formatTreemapSuperLabel(superCategory, sizeClass)}</p>
      </div>
      <div class="treemap-super-body">
        ${categoryLayouts.map(renderTreemapCategoryTile).join("")}
      </div>
    </section>
  `;
}

function renderTreemapSection(treemap) {
  if (!treemap || !treemap.superCategories.length) {
    return "";
  }

  const superCategoryLayouts = buildTreemapLayouts(
    treemap.superCategories.map((superCategory) => ({
      node: superCategory,
      value: superCategory.marketValueBase,
    })),
    getTreemapLayoutRect(getTreemapCanvasAspectRatio() * 100, 100),
  );

  return `
    <section
      class="treemap-card sidebar-anchor-section"
      id="dashboard-treemap"
      data-nav-section
      aria-label="Treemap"
    >
      <div class="treemap-wrap">
        <div class="treemap-canvas">
          ${superCategoryLayouts.map(renderTreemapSuperCategoryTile).join("")}
        </div>
        <div class="treemap-hover-card" aria-hidden="true">
          <strong class="treemap-hover-title"></strong>
          <div class="treemap-hover-meta">
            <span class="treemap-hover-value"></span>
            <span class="treemap-hover-weight"></span>
          </div>
        </div>
      </div>
    </section>
  `;
}

function buildMarketStatus(book) {
  const hasStoredQuotes = hasAppliedStoredQuotes(book);
  const unavailableDirectPriceNote = formatUnavailableDirectPriceNote(book);

  if (state.isRefreshing) {
    return {
      tone: "status-loading",
      title: "외부 시세 다시 불러오는 중",
      detail:
        state.refreshDetail ??
        "Yahoo Finance 시세를 직접 확인하고, 실패하면 저장된 가격 파일을 사용합니다.",
      note: unavailableDirectPriceNote,
    };
  }

  if (state.refreshError) {
    return {
      tone: "status-stale",
      title: hasStoredQuotes ? "새 시세 확인 실패" : "자동 시세 확인 실패",
      detail: getErrorMessage(state.refreshError),
      note: unavailableDirectPriceNote,
    };
  }

  if (isAutomaticLivePriceSource(book.livePriceSource)) {
    return {
      tone: "status-live",
      title: "외부 시세 자동 반영 중",
      detail: `${book.livePriceSource.provider} · 최근 반영 ${formatUpdatedAt(book.livePriceSource.updatedAt)}`,
      note: unavailableDirectPriceNote,
    };
  }

  if (hasStoredQuotes) {
    return {
      tone: "status-live",
      title: "저장된 최근 시세 반영 중",
      detail: `${book.livePriceSource?.provider ?? "가격 파일"} · 최근 반영 ${formatUpdatedAt(book.livePriceSource?.updatedAt)}`,
      note: unavailableDirectPriceNote,
    };
  }

  return {
    tone: "status-fallback",
    title: "스냅샷 가격으로 계산 중",
    detail: book.livePriceSource?.updatedAt
      ? `저장 파일 확인 ${formatUpdatedAt(book.livePriceSource.updatedAt)} · 현재 스냅샷과 동일`
      : "저장된 시세가 없어 마지막 스냅샷으로 계산합니다.",
    note: unavailableDirectPriceNote,
  };
}

function formatUnavailableDirectPriceNote(book) {
  const unavailableHoldings = buildAggregatedHoldings(book).filter(
    (holding) => holding.priceSource !== "live",
  );

  if (!unavailableHoldings.length) {
    return "";
  }

  const previewLimit = 3;
  const previewNames = unavailableHoldings
    .slice(0, previewLimit)
    .map((holding) => holding.name)
    .join(", ");
  const remainingCount = unavailableHoldings.length - Math.min(unavailableHoldings.length, previewLimit);
  const remainingText =
    remainingCount > 0 ? ` 외 ${currencyFormatter.format(remainingCount)}종목` : "";

  return `직접 현재가를 못 불러온 ${currencyFormatter.format(unavailableHoldings.length)}종목: ${previewNames}${remainingText}`;
}

function renderSummaryCards(portfolio) {
  const cashBreakdown = formatCashBreakdown(portfolio.cash, portfolio.cashUsd);
  const cards = [
    {
      label: "총자산",
      value: formatCurrency(portfolio.totals.totalAssets),
      detail: `주식 ${formatCurrency(portfolio.totals.marketValue)} + 현금 ${formatCurrency(portfolio.cashBase)}`,
      tone: "",
    },
    {
      label: "Total Return",
      value: formatSignedCurrency(portfolio.totals.totalReturnAmount),
      detail: `미실현손익 + 실현손익 + 누적배당 ${formatSignedPercent(portfolio.totals.totalReturnRate)}`,
      tone: getToneClass(portfolio.totals.totalReturnAmount),
    },
    {
      label: "현금",
      value: formatCurrency(portfolio.cashBase),
      detail: portfolio.cashUsd > 0 && cashBreakdown
        ? `${cashBreakdown} · 비중 ${formatPercent(portfolio.totals.cashWeight)}`
        : `비중 ${formatPercent(portfolio.totals.cashWeight)}`,
      tone: "",
    },
    {
      label: "실현손익",
      value: formatSignedCurrency(portfolio.totals.realizedProfitLoss),
      detail: "계좌 원장 누적 기준",
      tone: getToneClass(portfolio.totals.realizedProfitLoss),
    },
    {
      label: "누적 배당수익",
      value: formatCurrency(portfolio.totals.cumulativeDividendIncome),
      detail: "계좌 원장 누적 기준",
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

function getTopHoldings(portfolio, limit = 3) {
  return [...portfolio.holdings]
    .sort((left, right) => right.marketValueBase - left.marketValueBase)
    .slice(0, limit);
}

function getAggregatePriceSource(priceSources) {
  const uniqueSources = [...new Set(priceSources.filter(Boolean))];

  if (!uniqueSources.length) {
    return "snapshot";
  }

  if (uniqueSources.length === 1) {
    return uniqueSources[0];
  }

  return "mixed";
}

function buildAggregatedHoldings(book) {
  const aggregatedHoldings = new Map();

  const upsertCashHolding = ({
    key,
    ticker,
    name,
    currency,
    amount,
    amountBase,
    accountName,
  }) => {
    if (amount <= 0) {
      return;
    }

    const existingHolding = aggregatedHoldings.get(key) ?? {
      ticker,
      name,
      type: "현금",
      currency,
      quantity: 0,
      marketValue: 0,
      marketValueBase: 0,
      costBasis: 0,
      costBasisBase: 0,
      profitLoss: 0,
      profitLossBase: 0,
      annualDividendPerShare: 0,
      annualDividendIncome: 0,
      annualDividendIncomeBase: 0,
      realizedProfitLossKrw: 0,
      cumulativeDividendIncomeKrw: 0,
      averageCostKrwTotal: 0,
      averageCostKrwQuantity: 0,
      accountNames: new Set(),
      priceSources: [],
      nameInferred: false,
    };

    existingHolding.marketValue += amount;
    existingHolding.marketValueBase += amountBase;
    existingHolding.costBasis += amount;
    existingHolding.costBasisBase += amountBase;
    existingHolding.accountNames.add(accountName);
    aggregatedHoldings.set(key, existingHolding);
  };

  book.portfolios.forEach((portfolio) => {
    portfolio.holdings.forEach((holding) => {
      const key = `${holding.currency}:${holding.ticker}`;
      const existingHolding = aggregatedHoldings.get(key) ?? {
        ticker: holding.ticker,
        name: holding.name,
        type: holding.type,
        currency: holding.currency,
        quantity: 0,
        marketValue: 0,
        marketValueBase: 0,
        costBasis: 0,
        costBasisBase: 0,
        profitLoss: 0,
        profitLossBase: 0,
        annualDividendPerShare: 0,
        annualDividendIncome: 0,
        annualDividendIncomeBase: 0,
        realizedProfitLossKrw: 0,
        cumulativeDividendIncomeKrw: 0,
        averageCostKrwTotal: 0,
        averageCostKrwQuantity: 0,
        accountNames: new Set(),
        priceSources: [],
        nameInferred: false,
      };

      existingHolding.quantity += holding.quantity;
      existingHolding.marketValue += holding.marketValue;
      existingHolding.marketValueBase += holding.marketValueBase;
      existingHolding.costBasis += holding.costBasis;
      existingHolding.costBasisBase += holding.costBasisBase;
      existingHolding.profitLoss += holding.profitLoss;
      existingHolding.profitLossBase += holding.profitLossBase;
      existingHolding.annualDividendPerShare = holding.annualDividendPerShare;
      existingHolding.annualDividendIncome += holding.annualDividendIncome;
      existingHolding.annualDividendIncomeBase += holding.annualDividendIncomeBase;
      existingHolding.realizedProfitLossKrw += Number(holding.realizedProfitLossKrw ?? 0);
      existingHolding.cumulativeDividendIncomeKrw += Number(holding.cumulativeDividendIncomeKrw ?? 0);
      if (Number.isFinite(holding.averageCostKrw)) {
        existingHolding.averageCostKrwTotal += holding.averageCostKrw * holding.quantity;
        existingHolding.averageCostKrwQuantity += holding.quantity;
      }
      existingHolding.accountNames.add(portfolio.name);
      existingHolding.priceSources.push(holding.priceSource);
      existingHolding.nameInferred ||= holding.nameInferred;

      aggregatedHoldings.set(key, existingHolding);
    });

    upsertCashHolding({
      key: "KRW:KRW",
      ticker: "KRW",
      name: "현금(KRW)",
      currency: "KRW",
      amount: portfolio.cash,
      amountBase: portfolio.cash,
      accountName: portfolio.name,
    });

    upsertCashHolding({
      key: "USD:USD",
      ticker: "USD",
      name: "현금(USD)",
      currency: "USD",
      amount: portfolio.cashUsd,
      amountBase: convertToBaseCurrency(portfolio.cashUsd, "USD", portfolio.usdKrwRate),
      accountName: portfolio.name,
    });
  });

  return [...aggregatedHoldings.values()]
    .map((holding) => {
      const hasCompleteAverageCostKrw =
        holding.quantity > 0 &&
        Math.abs(holding.averageCostKrwQuantity - holding.quantity) < 0.000001;
      const averageCostKrw = hasCompleteAverageCostKrw
        ? holding.averageCostKrwTotal / holding.quantity
        : null;
      const profitRateBase =
        holding.type === "현금" || holding.costBasisBase === 0
          ? 0
          : holding.profitLossBase / holding.costBasisBase;
      const profitRate =
        holding.type === "현금" || holding.costBasis === 0
          ? 0
          : holding.profitLoss / holding.costBasis;
      const totalReturnAmount = calculateTotalReturnAmount(
        holding.profitLossBase,
        holding.realizedProfitLossKrw,
        holding.cumulativeDividendIncomeKrw,
      );
      const totalReturnRate =
        holding.type === "현금"
          ? 0
          : calculateTotalReturnRate(totalReturnAmount, holding.costBasisBase);
      const useKrwProfitDisplay = averageCostKrw !== null;

      return {
        ...holding,
        accountNames: [...holding.accountNames],
        accountCount: holding.accountNames.size,
        currentPrice:
          holding.type === "현금" || holding.quantity === 0 ? null : holding.marketValue / holding.quantity,
        averageCost:
          holding.type === "현금" || holding.quantity === 0 ? null : holding.costBasis / holding.quantity,
        averageCostCurrency: holding.currency,
        averageCostKrw,
        profitRate,
        profitRateDisplay: useKrwProfitDisplay ? profitRateBase : profitRate,
        profitLossDisplay: useKrwProfitDisplay ? holding.profitLossBase : holding.profitLoss,
        profitLossDisplayCurrency: useKrwProfitDisplay ? "KRW" : holding.currency,
        totalReturnAmount,
        totalReturnRate,
        assetWeight:
          book.overview.totalAssets === 0 ? 0 : holding.marketValueBase / book.overview.totalAssets,
        priceSource: holding.type === "현금" ? "cash" : getAggregatePriceSource(holding.priceSources),
      };
    })
    .sort(
      (left, right) =>
        right.marketValueBase - left.marketValueBase ||
        left.name.localeCompare(right.name, "ko-KR"),
    );
}

function resolvePortfolioCircleGroup(category) {
  switch (category.name) {
    case "SCHD":
    case "S&P500":
    case "Nasdaq":
    case "커버드콜":
    case "채권":
      return category.name;
    case "원화":
    case "USD":
    case "MM":
      return "현금성자산";
    default:
      return category.superCategory === "국내" ? "국내 주식/ETF" : "해외기타";
  }
}

function buildPortfolioCircleGroups(treemap) {
  if (!treemap) {
    return [];
  }

  const groups = new Map(
    PORTFOLIO_CIRCLE_GROUPS.map((group) => [
      group.key,
      {
        ...group,
        marketValueBase: 0,
        costBasisBase: 0,
        dividendIncomeBase: 0,
        profitLossBase: 0,
        realizedProfitLossKrw: 0,
        cumulativeDividendIncomeKrw: 0,
        members: new Set(),
        assetMap: new Map(),
      },
    ]),
  );

  treemap.superCategories.forEach((superCategory) => {
    superCategory.categories.forEach((category) => {
      const groupKey = resolvePortfolioCircleGroup(category);
      const group = groups.get(groupKey);

      if (!group) {
        return;
      }

      group.marketValueBase += category.marketValueBase;
      group.costBasisBase += category.costBasisBase;
      group.dividendIncomeBase += category.annualDividendIncomeBase;
      group.profitLossBase += category.profitLossBase;
      group.realizedProfitLossKrw += category.realizedProfitLossKrw;
      group.cumulativeDividendIncomeKrw += category.cumulativeDividendIncomeKrw;
      group.members.add(category.name);

      category.assets.forEach((asset) => {
        const assetKey = `${asset.currency}:${asset.ticker}`;
        const existingAsset = group.assetMap.get(assetKey) ?? {
          name: asset.name,
          marketValueBase: 0,
        };

        existingAsset.marketValueBase += asset.marketValueBase;
        group.assetMap.set(assetKey, existingAsset);
      });
    });
  });

  return PORTFOLIO_CIRCLE_GROUPS
    .map((group) => {
      const resolvedGroup = groups.get(group.key);
      const orderedMembers = group.memberOrder.filter((member) => resolvedGroup.members.has(member));
      const remainingMembers = [...resolvedGroup.members]
        .filter((member) => !orderedMembers.includes(member))
        .sort((left, right) => left.localeCompare(right, "ko-KR"));
      const members = [...orderedMembers, ...remainingMembers];
      const topAssets = [...resolvedGroup.assetMap.values()]
        .sort(
          (left, right) =>
            right.marketValueBase - left.marketValueBase ||
            left.name.localeCompare(right.name, "ko-KR"),
        )
        .slice(0, 2)
        .map((asset) => asset.name);
      let detail = topAssets.join(", ");

      if (resolvedGroup.key === "국내 주식/ETF") {
        detail = "배당ETF, 인프라/통신/금융 등";
      }

      if (resolvedGroup.key === "현금성자산") {
        detail = "MM, USD, KRW";
      }

      const totalReturnAmount = calculateTotalReturnAmount(
        resolvedGroup.profitLossBase,
        resolvedGroup.realizedProfitLossKrw,
        resolvedGroup.cumulativeDividendIncomeKrw,
      );
      const totalReturnRate = calculateTotalReturnRate(
        totalReturnAmount,
        resolvedGroup.costBasisBase,
      );

      return {
        ...resolvedGroup,
        members,
        detail,
        totalReturnAmount,
        totalReturnRate,
        weight:
          treemap.totalAssets === 0 ? 0 : resolvedGroup.marketValueBase / treemap.totalAssets,
      };
    })
    .filter((group) => group.marketValueBase > 0);
}

function buildPortfolioCircleSegments(groups) {
  const radius = 36;
  const circumference = 2 * Math.PI * radius;
  let offset = 0;

  return groups.map((group) => {
    const segmentWeight = group.chartWeight ?? group.weight;
    const segmentLength = circumference * segmentWeight;
    const gapLength = Math.min(circumference * 0.008, segmentLength * 0.3);
    const visibleLength = Math.max(segmentLength - gapLength, 0);
    const segment = {
      ...group,
      radius,
      dasharray: `${visibleLength.toFixed(3)} ${(circumference - visibleLength).toFixed(3)}`,
      dashoffset: `${(-offset).toFixed(3)}`,
    };

    offset += segmentLength;
    return segment;
  });
}

function renderPortfolioCirclePlot({
  title,
  chartLabel,
  segments,
  summaryValue,
  summaryMeta = "",
  summaryNote = "",
}) {
  const summaryClassName = `portfolio-circle-summary${
    summaryMeta || summaryNote ? "" : " is-single-line"
  }`;

  return `
    <section class="portfolio-circle-block">
      <p class="portfolio-circle-block-title">${title}</p>
      <div class="portfolio-circle-plot">
        <svg class="portfolio-circle-svg" viewBox="0 0 120 120" role="img" aria-label="${chartLabel}">
          <circle class="portfolio-circle-track" cx="60" cy="60" r="36" fill="none"></circle>
          ${segments
            .map(
              (segment) => `
                <circle
                  class="portfolio-circle-segment"
                  data-circle-group="${segment.key}"
                  cx="60"
                  cy="60"
                  r="${segment.radius}"
                  fill="none"
                  stroke="${segment.color}"
                  stroke-dasharray="${segment.dasharray}"
                  stroke-dashoffset="${segment.dashoffset}"
                ></circle>
              `,
            )
            .join("")}
        </svg>
        <div class="${summaryClassName}">
          <strong>${summaryValue}</strong>
          ${summaryMeta ? `<div class="portfolio-circle-summary-meta">${summaryMeta}</div>` : ""}
          ${summaryNote ? `<div class="portfolio-circle-summary-note">${summaryNote}</div>` : ""}
        </div>
      </div>
    </section>
  `;
}

function renderPortfolioCircleChart(treemap) {
  const groups = buildPortfolioCircleGroups(treemap);

  if (!groups.length) {
    return "";
  }

  const sortedGroups = [...groups].sort(
    (left, right) =>
      right.marketValueBase - left.marketValueBase ||
      left.label.localeCompare(right.label, "ko-KR"),
  );
  const segments = buildPortfolioCircleSegments(sortedGroups);
  const totalCostBasis = sortedGroups.reduce((sum, group) => sum + group.costBasisBase, 0);
  const totalDividendIncome = sortedGroups.reduce((sum, group) => sum + group.dividendIncomeBase, 0);
  const totalDividendYield =
    treemap.totalAssets === 0 ? 0 : totalDividendIncome / treemap.totalAssets;
  const totalProfitLoss = sortedGroups.reduce((sum, group) => sum + group.profitLossBase, 0);
  const totalRealizedProfitLoss =
    sortedGroups.reduce((sum, group) => sum + group.realizedProfitLossKrw, 0);
  const totalCumulativeDividendIncome =
    sortedGroups.reduce((sum, group) => sum + group.cumulativeDividendIncomeKrw, 0);
  const totalReturnAmount = calculateTotalReturnAmount(
    totalProfitLoss,
    totalRealizedProfitLoss,
    totalCumulativeDividendIncome,
  );
  const totalReturnRate = calculateTotalReturnRate(totalReturnAmount, totalCostBasis);
  const valueChartAriaLabel = sortedGroups
    .map((group) => `${group.label} ${formatPercent(group.weight)}`)
    .join(", ");
  const dividendGroups = sortedGroups.map((group) => ({
    ...group,
    chartWeight: totalDividendIncome === 0 ? 0 : group.dividendIncomeBase / totalDividendIncome,
  }));
  const dividendSegments = buildPortfolioCircleSegments(dividendGroups);
  const dividendChartAriaLabel = dividendGroups
    .map((group) => `${group.label} ${formatPercent(group.chartWeight)}`)
    .join(", ");

  return `
    <section
      class="portfolio-circle-card sidebar-anchor-section"
      id="dashboard-overview"
      data-nav-section
      aria-label="Overview"
    >
      <div class="portfolio-circle-layout">
        <div class="portfolio-circle-visual">
          ${renderPortfolioCirclePlot({
            title: "Value",
            chartLabel: `포트폴리오 평가금액 원형 차트: ${valueChartAriaLabel}`,
            segments,
            summaryValue: formatCurrency(treemap.totalAssets),
            summaryNote: `TR ${formatSignedPercentSingleDecimal(totalReturnRate)}`,
          })}
          ${renderPortfolioCirclePlot({
            title: "Dividend",
            chartLabel: `포트폴리오 연간 배당 원형 차트: ${dividendChartAriaLabel}`,
            segments: dividendSegments,
            summaryValue: formatCurrency(totalDividendIncome),
            summaryNote: `배당률 ${formatPercent(totalDividendYield)}`,
          })}
        </div>

        <div class="portfolio-circle-side">
          <div class="portfolio-circle-table">
            <div class="portfolio-circle-table-head">
              <span>카테고리</span>
              <span>평가금액/TR</span>
              <span>연 배당</span>
              <span>비중(금액/배당)</span>
            </div>

            <div class="portfolio-circle-list">
              ${sortedGroups
                .map(
                  (group) => `
                    <div class="portfolio-circle-item" data-circle-group="${group.key}">
                      <div class="portfolio-circle-cell portfolio-circle-label">
                        <div class="portfolio-circle-main">
                          <span class="portfolio-circle-swatch" style="background:${group.color};" aria-hidden="true"></span>
                          <strong>${group.label}</strong>
                        </div>
                        ${group.detail ? `<span class="portfolio-circle-detail">${group.detail}</span>` : ""}
                      </div>
                      <div class="portfolio-circle-cell portfolio-circle-value">
                        <strong>${formatCurrency(group.marketValueBase)}</strong>
                        <span class="portfolio-circle-value-sub ${getToneClass(group.totalReturnAmount)}">${formatProfitWithRateSingleDecimal(group.totalReturnAmount, group.costBasisBase)}</span>
                      </div>
                      <div class="portfolio-circle-cell portfolio-circle-dividend">
                        <strong>${formatCurrency(group.dividendIncomeBase)}</strong>
                        <span class="portfolio-circle-dividend-sub">${formatPercent(group.marketValueBase === 0 ? 0 : group.dividendIncomeBase / group.marketValueBase)}</span>
                      </div>
                      <div class="portfolio-circle-cell portfolio-circle-weight">
                        <strong>${formatPercent(group.weight)} / ${formatPercent(totalDividendIncome === 0 ? 0 : group.dividendIncomeBase / totalDividendIncome)}</strong>
                      </div>
                    </div>
                  `,
                )
                .join("")}
            </div>

            <div class="portfolio-circle-total">
              <div class="portfolio-circle-cell portfolio-circle-label">
                <strong>합계</strong>
              </div>
              <div class="portfolio-circle-cell portfolio-circle-value">
                <strong>${formatCurrency(treemap.totalAssets)}</strong>
                <span class="portfolio-circle-value-sub ${getToneClass(totalReturnAmount)}">${formatProfitWithRateSingleDecimal(totalReturnAmount, totalCostBasis)}</span>
              </div>
              <div class="portfolio-circle-cell portfolio-circle-dividend">
                <strong>${formatCurrency(totalDividendIncome)}</strong>
                <span class="portfolio-circle-dividend-sub">${formatPercent(treemap.totalAssets === 0 ? 0 : totalDividendIncome / treemap.totalAssets)}</span>
              </div>
              <div class="portfolio-circle-cell portfolio-circle-weight">
                <strong>${formatPercent(1)} / ${formatPercent(1)}</strong>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>
  `;
}

function buildDividendCalendarTooltip(monthData, calendar) {
  const lines = [`${monthData.label} 배당 추정`];
  const convertedUsdValue = monthData.totals.USD * calendar.usdKrwRate;

  lines.push(`원화 ${formatMoney(monthData.totals.KRW, "KRW")}`);
  lines.push(
    calendar.usdKrwRate > 0
      ? `외화 ${formatMoney(monthData.totals.USD, "USD")} (${formatCurrency(convertedUsdValue)} 환산)`
      : `외화 ${formatMoney(monthData.totals.USD, "USD")}`,
  );

  const appendItems = (label, items, currency) => {
    if (!items.length) {
      return;
    }

    lines.push(label);
    [...items]
      .sort(
        (left, right) =>
          right.amount - left.amount ||
          left.name.localeCompare(right.name, "ko-KR"),
      )
      .slice(0, 4)
      .forEach((item) => {
        lines.push(`- ${item.accountName} · ${item.name}: ${formatMoney(item.amount, currency)}`);
      });
  };

  appendItems("원화 구성", monthData.items.KRW, "KRW");
  appendItems("외화 구성", monthData.items.USD, "USD");

  return escapeHtml(lines.join("\n"));
}

function renderDividendCalendarSection(book, source, error = null, exchangeRateSource = null) {
  const calendar = buildDividendCalendarData(book, source, error, exchangeRateSource);
  const usdRateCopy =
    calendar.usdKrwRate > 0
      ? `1 USD = ${formatExchangeRate(calendar.usdKrwRate)} 기준으로 외화 구간을 환산해 원화 구간 위에 함께 누적합니다.`
      : "외화 막대 높이는 USD 원금 기준으로 표시합니다.";
  const sectionCopy = error
    ? `배당 캘린더 파일을 읽지 못했습니다: ${error.message}`
    : `월을 바닥 축에 두고, 한 막대 안에 외화 배당을 먼저 쌓고 그 위에 원화 배당을 올립니다. ${usdRateCopy}`;
  const exchangeRateMeta = calendar.exchangeRateUpdatedAt
    ? `환율 ${formatUpdatedAt(calendar.exchangeRateUpdatedAt)}`
    : calendar.usesLiveExchangeRate
      ? `환율 ${calendar.exchangeRateProvider ?? "실시간"}`
      : "환율 포트폴리오 기준";
  const entryMeta = calendar.updatedAt
    ? `엔트리 ${formatUpdatedAt(calendar.updatedAt)}`
    : "엔트리 data/dividend-calendar.json";

  return `
    <section
      class="dividend-calendar-card sidebar-anchor-section"
      id="dividend-calendar"
      data-nav-section
    >
      <div class="section-head">
        <div>
          <h2 class="section-title">배당 캘린더</h2>
          <p class="section-copy">${sectionCopy}</p>
        </div>
        <div class="section-pill">${calendar.year}년 추정</div>
      </div>

      <div class="dividend-calendar-summary">
        <article class="dividend-calendar-stat" data-currency="KRW">
          <span class="dividend-calendar-stat-label">원화 연간 추정</span>
          <strong class="dividend-calendar-stat-value">${formatMoney(calendar.totals.KRW, "KRW")}</strong>
          <span class="dividend-calendar-stat-sub">입력 월 ${currencyFormatter.format(calendar.months.filter((monthData) => monthData.totals.KRW > 0).length)}개</span>
        </article>
        <article class="dividend-calendar-stat" data-currency="USD">
          <span class="dividend-calendar-stat-label">달러 연간 추정</span>
          <strong class="dividend-calendar-stat-value">${formatMoney(calendar.totals.USD, "USD")}</strong>
          <span class="dividend-calendar-stat-sub">입력 월 ${currencyFormatter.format(calendar.months.filter((monthData) => monthData.totals.USD > 0).length)}개</span>
        </article>
        <article class="dividend-calendar-stat" data-currency="BASE">
          <span class="dividend-calendar-stat-label">환산 합계</span>
          <strong class="dividend-calendar-stat-value">${formatCurrency(calendar.totalBaseKrw)}</strong>
          <span class="dividend-calendar-stat-sub">외화 약 ${formatCurrency(calendar.convertedUsdTotalKrw)} 환산 포함</span>
        </article>
        <article class="dividend-calendar-stat">
          <span class="dividend-calendar-stat-label">적용 환율 / 엔트리</span>
          <strong class="dividend-calendar-stat-value">${formatExchangeRate(calendar.usdKrwRate)}</strong>
          <span class="dividend-calendar-stat-sub">${currencyFormatter.format(calendar.populatedEntryCount)} / ${currencyFormatter.format(calendar.entryCount)} · ${entryMeta} · ${exchangeRateMeta}</span>
        </article>
      </div>

      ${calendar.hasData
        ? ""
        : `
          <div class="dividend-calendar-empty">
            data/dividend-calendar.json의 각 종목 monthlyAmounts에 1~12월 금액을 넣으면 월별 막대가 채워집니다.
          </div>
        `}

      <div class="dividend-calendar-combined-wrap">
        <div class="dividend-calendar-combined-meta">
          <span class="dividend-calendar-combined-chip is-usd">막대 하단: 외화 배당</span>
          <span class="dividend-calendar-combined-chip is-krw">막대 상단: 원화 배당</span>
          <span class="dividend-calendar-combined-note">${calendar.usdKrwRate > 0 ? `우측 축은 현재 환율 ${formatExchangeRate(calendar.usdKrwRate)}/USD 환산 기준` : "우측 축 환산 정보 없음"}</span>
        </div>

        <div class="dividend-calendar-combined-chart">
          <div class="dividend-calendar-axis dividend-calendar-axis-left" aria-hidden="true">
            <span class="dividend-calendar-axis-title">KRW</span>
            <span class="dividend-calendar-axis-value is-top">${formatMoney(calendar.sharedMaxKrw, "KRW")}</span>
            <span class="dividend-calendar-axis-zero">0</span>
          </div>

          <div
            class="dividend-calendar-combined-plot"
            role="img"
            aria-label="${calendar.year}년 월별 배당 추정 스택 막대그래프: 각 월 막대의 하단은 외화 배당, 상단은 원화 배당이며 외화는 현재 환율로 환산해 높이를 맞췄습니다."
          >
            ${calendar.months
              .map((monthData) => {
                const usdKrwValue = monthData.totals.USD * calendar.usdKrwRate;
                const totalBaseValue = monthData.totals.KRW + usdKrwValue;
                const totalHeight =
                  calendar.sharedMaxKrw > 0 && totalBaseValue > 0
                    ? (totalBaseValue / calendar.sharedMaxKrw) * 100
                    : 0;
                const usdShare = totalBaseValue > 0 ? (usdKrwValue / totalBaseValue) * 100 : 0;
                const krwShare = totalBaseValue > 0 ? (monthData.totals.KRW / totalBaseValue) * 100 : 0;

                return `
                  <div
                    class="dividend-calendar-month${monthData.totals.KRW > 0 || monthData.totals.USD > 0 ? "" : " is-empty"}"
                    title="${buildDividendCalendarTooltip(monthData, calendar)}"
                  >
                    <span class="dividend-calendar-value is-total">${totalBaseValue > 0 ? formatCompactMoney(totalBaseValue, "KRW") : ""}</span>
                    <div class="dividend-calendar-stack-shell">
                      <div class="dividend-calendar-stack" style="height:${totalHeight.toFixed(1)}%">
                        <span class="dividend-calendar-bar is-usd" style="height:${usdShare.toFixed(1)}%"></span>
                        <span class="dividend-calendar-bar is-krw" style="height:${krwShare.toFixed(1)}%"></span>
                      </div>
                    </div>
                    <span class="dividend-calendar-label">${monthData.label}</span>
                  </div>
                `;
              })
              .join("")}
          </div>

          <div class="dividend-calendar-axis dividend-calendar-axis-right" aria-hidden="true">
            <span class="dividend-calendar-axis-title">USD</span>
            <span class="dividend-calendar-axis-value is-top">${formatMoney(calendar.sharedMaxUsd, "USD")}</span>
            <span class="dividend-calendar-axis-zero">$0</span>
          </div>
        </div>
      </div>
    </section>
  `;
}

function setPortfolioCircleActiveGroup(card, activeGroup) {
  if (!card) {
    return;
  }

  card.querySelectorAll("[data-circle-group]").forEach((element) => {
    const isActive = activeGroup && element.dataset.circleGroup === activeGroup;
    const isDimmed = activeGroup && element.dataset.circleGroup !== activeGroup;

    element.classList.toggle("is-active", Boolean(isActive));
    element.classList.toggle("is-dimmed", Boolean(isDimmed));
  });
}

function hydratePortfolioCircleInteractions() {
  document.querySelectorAll(".portfolio-circle-card").forEach((card) => {
    card.querySelectorAll("[data-circle-group]").forEach((element) => {
      if (element.dataset.circleBound === "true") {
        return;
      }

      element.dataset.circleBound = "true";
      element.addEventListener("mouseenter", () => {
        setPortfolioCircleActiveGroup(card, element.dataset.circleGroup);
      });
      element.addEventListener("mouseleave", () => {
        setPortfolioCircleActiveGroup(card, "");
      });
    });
  });
}

function setTreemapHoverPosition(card, tooltip, event) {
  const wrap = card.querySelector(".treemap-wrap");

  if (!wrap) {
    return;
  }

  const wrapRect = wrap.getBoundingClientRect();
  const tooltipRect = tooltip.getBoundingClientRect();
  let left = event.clientX - wrapRect.left + 16;
  let top = event.clientY - wrapRect.top + 16;

  left = Math.max(12, Math.min(left, wrapRect.width - tooltipRect.width - 12));
  top = Math.max(12, Math.min(top, wrapRect.height - tooltipRect.height - 12));

  tooltip.style.transform = `translate(${left}px, ${top}px)`;
}

function hydrateTreemapInteractions() {
  document.querySelectorAll(".treemap-card").forEach((card) => {
    const tooltip = card.querySelector(".treemap-hover-card");

    if (!tooltip) {
      return;
    }

    const title = tooltip.querySelector(".treemap-hover-title");
    const value = tooltip.querySelector(".treemap-hover-value");
    const weight = tooltip.querySelector(".treemap-hover-weight");
    const clearActive = () => {
      card.querySelectorAll(".treemap-tile.is-hovered").forEach((tile) => {
        tile.classList.remove("is-hovered");
      });
      tooltip.classList.remove("is-visible");
    };

    card.querySelectorAll("[data-treemap-hover]").forEach((tile) => {
      if (tile.dataset.treemapBound === "true") {
        return;
      }

      tile.dataset.treemapBound = "true";
      tile.addEventListener("mouseenter", (event) => {
        title.textContent = decodeTreemapHoverValue(tile.dataset.treemapName);
        value.textContent = decodeTreemapHoverValue(tile.dataset.treemapValue);
        weight.textContent = decodeTreemapHoverValue(tile.dataset.treemapWeight);
        clearActive();
        tile.classList.add("is-hovered");
        tooltip.classList.add("is-visible");
        setTreemapHoverPosition(card, tooltip, event);
      });
      tile.addEventListener("mousemove", (event) => {
        if (!tooltip.classList.contains("is-visible")) {
          return;
        }

        setTreemapHoverPosition(card, tooltip, event);
      });
      tile.addEventListener("mouseleave", () => {
        clearActive();
      });
    });
  });
}

function renderAllHoldingsRows(holdings) {
  return holdings
    .map(
      (holding) => {
        const isCash = holding.type === "현금";

        return `
        <tr>
          <td>
            <div class="asset-name">${holding.name}</div>
          </td>
          <td>${isCash ? "-" : formatQuantity(holding.quantity)}</td>
          <td>
            ${isCash
              ? "-"
              : `
                <div class="metric">
                  <strong>${formatOverallMoney(holding.currentPrice, holding.currency)}</strong>
                </div>
              `}
          </td>
          <td>${isCash ? "-" : renderAverageCostCell(holding.averageCost, holding.averageCostCurrency ?? holding.currency, holding.averageCostKrw, { overall: true })}</td>
          <td>
            <div class="all-holdings-value-stack">
              <strong>${formatOverallMoney(holding.marketValue, holding.currency)}</strong>
              <span class="all-holdings-value-sub ${isCash ? "" : getToneClass(holding.totalReturnAmount)}">${isCash ? "-" : formatOverallProfitWithRate(holding.totalReturnAmount, holding.totalReturnRate, "KRW")}</span>
            </div>
          </td>
          <td>
            ${isCash
              ? "-"
              : `
                <div class="all-holdings-value-stack">
                  <strong>${formatOverallMoney(holding.annualDividendPerShare, holding.currency)}</strong>
                  <span class="all-holdings-value-sub">${formatOverallDividendYield(holding.annualDividendPerShare, holding.currentPrice)}</span>
                </div>
              `}
          </td>
          <td>${isCash ? "-" : formatOverallMoney(holding.annualDividendIncome, holding.currency)}</td>
          <td>
            ${isCash
              ? "-"
              : formatOverallProfitWithRate(
                holding.profitLossDisplay,
                holding.profitRateDisplay,
                holding.profitLossDisplayCurrency,
              )}
          </td>
          <td>${isCash ? "-" : formatSignedCurrency(holding.realizedProfitLossKrw ?? 0)}</td>
          <td>${isCash ? "-" : formatCurrency(holding.cumulativeDividendIncomeKrw ?? 0)}</td>
          <td>${formatPercent(holding.assetWeight)}</td>
        </tr>
      `;
      },
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
          </td>
          <td>${formatQuantity(holding.quantity)}</td>
          <td>
            <div class="metric">
              <strong>${formatPrice(holding.currentPrice, holding.currency)}</strong>
            </div>
          </td>
          <td>${renderAverageCostCell(holding.averageCost, holding.averageCostCurrency ?? holding.currency, holding.averageCostKrw)}</td>
          <td>
            <div class="all-holdings-value-stack">
              <strong>${formatHoldingMoney(holding.marketValue, holding.currency)}</strong>
              <span class="all-holdings-value-sub ${getToneClass(holding.totalReturnAmount)}">${formatHoldingProfitWithRate(holding.totalReturnAmount, holding.totalReturnRate, "KRW")}</span>
            </div>
          </td>
          <td>
            <div class="all-holdings-value-stack">
              <strong>${formatMoney(holding.annualDividendPerShare, holding.currency)}</strong>
              <span class="all-holdings-value-sub">${formatOverallDividendYield(holding.annualDividendPerShare, holding.currentPrice)}</span>
            </div>
          </td>
          <td>${formatHoldingMoney(holding.annualDividendIncome, holding.currency)}</td>
          <td>${formatHoldingProfitWithRate(
            holding.profitLossDisplay,
            holding.profitRateDisplay,
            holding.profitLossDisplayCurrency,
          )}</td>
          <td>${formatSignedCurrency(holding.realizedProfitLossKrw ?? 0)}</td>
          <td>${formatCurrency(holding.cumulativeDividendIncomeKrw ?? 0)}</td>
          <td>${formatPercent(holding.assetWeight)}</td>
        </tr>
      `,
    )
    .join("");
}

function renderAllHoldingsSection(book) {
  const aggregatedHoldings = buildAggregatedHoldings(book);
  const uniqueHoldingsCount = aggregatedHoldings.length;
  const aggregatedCostBasis = aggregatedHoldings.reduce(
    (total, holding) => total + holding.costBasisBase,
    0,
  );
  const aggregatedMarketValue = aggregatedHoldings.reduce(
    (total, holding) => total + holding.marketValueBase,
    0,
  );
  const aggregatedProfitLoss = aggregatedHoldings.reduce(
    (total, holding) => total + holding.profitLossBase,
    0,
  );
  const aggregatedAnnualDividendIncome = aggregatedHoldings.reduce(
    (total, holding) => total + holding.annualDividendIncomeBase,
    0,
  );
  const aggregatedRealizedProfitLoss = aggregatedHoldings.reduce(
    (total, holding) => total + Number(holding.realizedProfitLossKrw ?? 0),
    0,
  );
  const aggregatedCumulativeDividendIncome = aggregatedHoldings.reduce(
    (total, holding) => total + Number(holding.cumulativeDividendIncomeKrw ?? 0),
    0,
  );
  const aggregatedProfitRate =
    aggregatedCostBasis === 0 ? 0 : aggregatedProfitLoss / aggregatedCostBasis;
  const aggregatedTotalReturnAmount = calculateTotalReturnAmount(
    aggregatedProfitLoss,
    aggregatedRealizedProfitLoss,
    aggregatedCumulativeDividendIncome,
  );
  const aggregatedTotalReturnRate = calculateTotalReturnRate(
    aggregatedTotalReturnAmount,
    aggregatedCostBasis,
  );

  return `
    <section
      class="table-card all-holdings-card sidebar-anchor-section"
      id="all-holdings"
      data-nav-section
    >
      <div class="section-head">
        <div>
          <h2 class="section-title">Overall</h2>
          <p class="section-copy">
            같은 종목을 여러 계좌에서 보유 중이면 한 줄로 합산했고, 현금(KRW, USD)도 함께 포함했습니다. 수량과 평가금액은 종목 기준으로 합치고, 평가금액 아래 수익은 Total Return(미실현손익 + 실현손익 + 누적배당) 기준으로 원화 합산 표시합니다.
          </p>
        </div>
        <div class="section-pill">중복 합산 후 ${currencyFormatter.format(uniqueHoldingsCount)}개 자산</div>
      </div>

      <div class="table-wrap">
        <table class="all-holdings-table">
          <thead>
            <tr>
              <th>종목</th>
              <th>수량</th>
              <th>현재가</th>
              <th>평균단가</th>
              <th>평가금액</th>
              <th>배당/1주</th>
              <th>연배당</th>
              <th>미실현손익</th>
              <th>실현손익</th>
              <th>누적배당</th>
              <th>비중</th>
            </tr>
          </thead>
          <tbody>
            ${renderAllHoldingsRows(aggregatedHoldings)}
          </tbody>
          <tfoot>
            <tr>
              <td>합계 (현금 포함, 원화 환산)</td>
              <td>-</td>
              <td>-</td>
              <td>-</td>
              <td>
                <div class="all-holdings-value-stack">
                  <strong>${formatCurrency(aggregatedMarketValue)}</strong>
                  <span class="all-holdings-value-sub ${getToneClass(aggregatedTotalReturnAmount)}">${formatProfitWithRateCompact(aggregatedTotalReturnAmount, aggregatedTotalReturnRate)}</span>
                </div>
              </td>
              <td>-</td>
              <td>${formatCurrency(aggregatedAnnualDividendIncome)}</td>
              <td>${formatProfitWithRateCompact(aggregatedProfitLoss, aggregatedProfitRate)}</td>
              <td>${formatSignedCurrency(aggregatedRealizedProfitLoss)}</td>
              <td>${formatCurrency(aggregatedCumulativeDividendIncome)}</td>
              <td>${formatPercent(book.overview.totalAssets === 0 ? 0 : aggregatedMarketValue / book.overview.totalAssets)}</td>
            </tr>
          </tfoot>
        </table>
      </div>
    </section>
  `;
}

function renderAccountSummarySection(book) {
  const groups = buildAccountSummaryGroups(book);

  return `
    <section
      class="table-card account-summary-card sidebar-anchor-section"
      id="account-summary"
      data-nav-section
    >
      <div class="section-head">
        <div>
          <h2 class="section-title">계좌 요약</h2>
          <p class="section-copy">
            직투계좌, ISA, 연금저축, 퇴직연금을 5개 카테고리로 묶어 총자산과 연배당을 비교합니다. 총자산 옆 괄호는 Total Return(미실현손익 + 실현손익 + 누적배당) 기준이며, 연금저축(세액미공제)는 별도 분류하고 나머지 연금저축 계좌는 세액공제로 묶었습니다. 직투계좌만 배당소득세 15%를 반영했습니다.
          </p>
        </div>
        <div class="section-pill">총자산 ${formatCurrency(book.overview.totalAssets)}</div>
      </div>

      <div class="account-summary-wrap">
        <div class="account-summary-list">
          ${groups
            .map(
              (group) => `
                <div
                  class="account-summary-row"
                  style="--account-summary-color:${group.color}; --account-summary-width:${(group.barScale * 100).toFixed(1)}%;"
                >
                  <div class="account-summary-head">
                    <div class="account-summary-copy">
                      <strong class="account-summary-name">${group.label}</strong>
                      <span class="account-summary-meta">${group.accountsLabel || "해당 계좌 없음"}</span>
                    </div>
                    <div class="account-summary-stats">
                      <div class="account-summary-amount-line">
                        <strong class="account-summary-amount">${formatCurrency(group.totalAssets)}</strong>
                        <span class="account-summary-profit ${getToneClass(group.totalReturnAmount)}">(${formatSignedCurrency(group.totalReturnAmount)}, TR ${formatSignedPercentCompact(group.totalReturnRate)})</span>
                      </div>
                      <span class="account-summary-share">${formatPercent(group.weight)}</span>
                    </div>
                  </div>

                  <div class="account-summary-bar" aria-hidden="true">
                    <span class="account-summary-bar-fill"></span>
                  </div>

                  <div class="account-summary-dividend-line">
                    <span class="account-summary-dividend-item">
                      <strong>연배당</strong>
                      <span>${formatCurrency(group.annualDividendIncome)} (${formatPercent(group.dividendYield)})</span>
                    </span>
                    <span class="account-summary-dividend-item ${group.dividendTax > 0 ? "is-taxed" : ""}">
                      <strong>세금</strong>
                      <span>${group.dividendTax > 0 ? formatSignedCurrency(-group.dividendTax) : "-"}</span>
                    </span>
                    <span class="account-summary-dividend-item">
                      <strong>세후</strong>
                      <span>${formatCurrency(group.netDividendIncome)}</span>
                    </span>
                    <span class="account-summary-dividend-item">
                      <strong>세후 월배당</strong>
                      <span>${formatCurrency(group.monthlyNetDividendIncome)}</span>
                    </span>
                    <span class="account-summary-dividend-item">
                      <strong>실현손익</strong>
                      <span class="${getToneClass(group.realizedProfitLoss)}">${formatSignedCurrency(group.realizedProfitLoss)}</span>
                    </span>
                    <span class="account-summary-dividend-item">
                      <strong>누적배당</strong>
                      <span>${formatCurrency(group.cumulativeDividendIncome)}</span>
                    </span>
                  </div>
                </div>
              `,
            )
            .join("")}
        </div>
      </div>
    </section>
  `;
}

function renderPortfolioSection(portfolio) {
  const heroCopy = portfolio.hasForeignCurrency
    ? `${portfolio.snapshotLabel} 원장 데이터를 유지하고, 현재가와 환율만 외부 시세 또는 저장된 최신 시세 파일에서 다시 읽어 계산합니다. 해외 종목 평균단가는 원래 종목 통화 기준으로 표시하고, 원화 기준 입력값이 있으면 아래에 참고용으로 함께 보여줍니다. 평가손익은 평균단가 기준으로 계산합니다. 계좌 합계와 비중은 KRW 환산으로 계산합니다.`
    : `${portfolio.snapshotLabel} 원장 데이터를 유지하고, 현재가만 외부 시세 또는 저장된 최신 시세 파일에서 다시 읽어 평가금액과 손익을 재계산합니다.`;
  const tableCopy = portfolio.hasForeignCurrency
    ? "수량과 평균단가는 고정 원장 데이터입니다. 해외 종목 평균단가는 종목 통화 기준을 기본으로 보여주고, 원화 기준 입력값이 있으면 참고용으로 함께 표시합니다. 현재가와 평가금액, 배당 관련 값은 종목 통화 기준으로 표기하며, 평가금액 아래 수익은 Total Return(미실현손익 + 실현손익 + 누적배당) 기준으로 원화 환산 계산합니다. 계좌 합계와 비중은 KRW 환산 기준으로 계산합니다."
    : "수량과 평균단가는 고정 원장 데이터입니다. 현재가, 평가금액, 배당/1주, 연배당은 최신 시세와 최근 1년 배당 기준으로 다시 계산됩니다. 평가금액 아래 수익은 Total Return(미실현손익 + 실현손익 + 누적배당) 기준으로 원화 표시합니다.";

  return `
    <section
      class="account-stack sidebar-anchor-section"
      id="${portfolio.id}"
      data-nav-section
    >
      <section class="hero">
        <div class="eyebrow">Account Detail</div>
        <div class="hero-head">
          <div>
            <h2 class="hero-title">${portfolio.name}</h2>
            <p class="hero-copy">${heroCopy}</p>
          </div>
          <div class="account-pill">총자산 ${formatCurrency(portfolio.totals.totalAssets)}</div>
        </div>
        <div class="summary-grid">
          ${renderSummaryCards(portfolio)}
        </div>
      </section>

      <section class="table-card">
        <div class="section-head">
          <div>
            <h3 class="section-title">보유 종목 표</h3>
            <p class="section-copy">${tableCopy}</p>
          </div>
          <div class="section-pill">총 ${portfolio.holdings.length}종목</div>
        </div>

        <div class="table-wrap">
          <table class="all-holdings-table">
            <thead>
              <tr>
                <th>종목</th>
                <th>수량</th>
                <th>현재가</th>
                <th>평균단가</th>
                <th>평가금액</th>
                <th>배당/1주</th>
                <th>연배당</th>
                <th>미실현손익</th>
                <th>실현손익</th>
                <th>누적배당</th>
                <th>비중</th>
              </tr>
            </thead>
            <tbody>
              ${renderTableRows(portfolio)}
            </tbody>
            <tfoot>
              <tr>
                <td>합계 (주식)</td>
                <td>${formatQuantity(portfolio.totals.quantity)}</td>
                <td>-</td>
                <td>-</td>
                <td>
                  <div class="all-holdings-value-stack">
                    <strong>${formatCurrency(portfolio.totals.marketValue)}</strong>
                    <span class="all-holdings-value-sub ${getToneClass(portfolio.totals.totalReturnAmount)}">${formatProfitWithRateCompact(portfolio.totals.totalReturnAmount, portfolio.totals.totalReturnRate)}</span>
                  </div>
                </td>
                <td>-</td>
                <td>${formatCurrency(portfolio.totals.annualDividendIncome)}</td>
                <td>${formatProfitWithRateCompact(portfolio.totals.profitLoss, portfolio.totals.totalReturn)}</td>
                <td>${formatSignedCurrency(portfolio.totals.realizedProfitLoss)}</td>
                <td>${formatCurrency(portfolio.totals.cumulativeDividendIncome)}</td>
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
      <p class="state-copy">계좌 원장과 외부 현재가, 저장된 가격 파일을 순서대로 확인하고 있습니다.</p>
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

function renderSidebarNavigation(book) {
  return `
    <aside class="dashboard-sidebar">
      <div class="sidebar-nav-card">
        <div class="sidebar-nav-brand">
          <div>
            <strong>Portfolio</strong>
          </div>
        </div>

        <nav class="sidebar-nav-groups" aria-label="페이지 메뉴">
          <div class="sidebar-nav-group">
            <p class="sidebar-nav-group-label">Dashboard</p>
            <div class="sidebar-nav-subgroup">
              <a
                class="sidebar-nav-link is-subitem is-active"
                href="#dashboard-overview"
                data-section-link
                data-section-target="dashboard-overview"
              >
                <span>Overview</span>
              </a>
              <a
                class="sidebar-nav-link is-subitem"
                href="#dashboard-treemap"
                data-section-link
                data-section-target="dashboard-treemap"
              >
                <span>Treemap</span>
              </a>
              <a
                class="sidebar-nav-link is-subitem"
                href="#dividend-calendar"
                data-section-link
                data-section-target="dividend-calendar"
              >
                <span>Dividend Calendar</span>
              </a>
              <a
                class="sidebar-nav-link is-subitem"
                href="#all-holdings"
                data-section-link
                data-section-target="all-holdings"
              >
                <span>Overall</span>
              </a>
            </div>
          </div>

          <div class="sidebar-nav-group">
            <p class="sidebar-nav-group-label">Accounts</p>
            <div class="sidebar-nav-subgroup">
              <a
                class="sidebar-nav-link is-subitem"
                href="#account-summary"
                data-section-link
                data-section-target="account-summary"
              >
                <span>Overview</span>
              </a>
              ${book.portfolios
                .map(
                  (portfolio) => `
                    <a
                      class="sidebar-nav-link is-subitem"
                      href="#${portfolio.id}"
                      data-section-link
                      data-section-target="${portfolio.id}"
                    >
                      <span>${portfolio.name}</span>
                    </a>
                  `,
                )
                .join("")}
            </div>
          </div>
        </nav>
      </div>
    </aside>
  `;
}

function setSidebarActiveSection(sectionId) {
  if (!sectionId) {
    return;
  }

  document.querySelectorAll("[data-section-link]").forEach((link) => {
    link.classList.toggle("is-active", link.dataset.sectionTarget === sectionId);
  });
}

function hydrateSidebarNavigation() {
  if (sidebarSectionObserver) {
    sidebarSectionObserver.disconnect();
    sidebarSectionObserver = null;
  }

  const links = [...document.querySelectorAll("[data-section-link]")];

  if (!links.length) {
    return;
  }

  const sectionIds = [...new Set(links.map((link) => link.dataset.sectionTarget).filter(Boolean))];
  const sections = sectionIds.map((id) => document.getElementById(id)).filter(Boolean);

  links.forEach((link) => {
    if (link.dataset.sidebarBound === "true") {
      return;
    }

    link.dataset.sidebarBound = "true";
    link.addEventListener("click", () => {
      setSidebarActiveSection(link.dataset.sectionTarget);
    });
  });

  if (!sections.length) {
    return;
  }

  const visibleSections = new Map();
  sidebarSectionObserver = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          visibleSections.set(entry.target.id, Math.abs(entry.boundingClientRect.top));
          return;
        }

        visibleSections.delete(entry.target.id);
      });

      const nextSectionId = [...visibleSections.entries()].sort((left, right) => left[1] - right[1])[0]?.[0];

      if (nextSectionId) {
        setSidebarActiveSection(nextSectionId);
      }
    },
    {
      rootMargin: "-18% 0px -64% 0px",
      threshold: [0, 0.15, 0.4],
    },
  );

  sections.forEach((section) => {
    sidebarSectionObserver.observe(section);
  });

  setSidebarActiveSection(sections[0].id);
}

function renderApp(book) {
  const app = document.querySelector("#app");

  if (!app) {
    return;
  }

  const marketStatus = buildMarketStatus(book);
  const treemap = buildTreemapData(book);

  app.innerHTML = `
    <section class="dashboard-shell">
      ${renderSidebarNavigation(book)}

      <section class="dashboard-main">
        <section class="dashboard">
          <section class="overview-section">
            <div class="hero-toolbar dashboard-toolbar">
              <div class="market-status ${marketStatus.tone}">
                <span class="status-dot"></span>
                <div>
                  <strong>${marketStatus.title}</strong>
                  <span>${marketStatus.detail}</span>
                  ${marketStatus.note ? `<span>${marketStatus.note}</span>` : ""}
                </div>
              </div>
              <div class="toolbar-actions">
                <button class="ghost-button" data-refresh-quotes ${state.isRefreshing ? "disabled" : ""}>
                  ${state.isRefreshing ? "시세 불러오는 중..." : "시세 다시 불러오기"}
                </button>
                <button
                  class="ghost-button ${state.hideAmounts ? "ghost-button-muted" : ""}"
                  data-toggle-amount-visibility
                  aria-pressed="${state.hideAmounts ? "true" : "false"}"
                >
                  금액표시 ${state.hideAmounts ? "Off" : "On"}
                </button>
              </div>
            </div>

          </section>

          ${renderPortfolioCircleChart(treemap)}

          ${renderTreemapSection(treemap)}

          ${renderDividendCalendarSection(
            book,
            state.dividendCalendarSource,
            state.dividendCalendarError,
            state.liveExchangeRateSource,
          )}

          ${renderAllHoldingsSection(book)}

          ${renderAccountSummarySection(book)}

          <section
            class="detail-section sidebar-anchor-section"
            id="account-details"
            data-nav-section
          >
            <div class="section-head detail-head">
              <div>
                <h2 class="section-title">계좌 상세</h2>
                <p class="section-copy">
                  각 계좌의 원장 기준 수량과 평단, 그리고 최신 현재가와 최근 1년 배당 기준 평가금액, 배당/1주, 연배당을 상세 표로 확인합니다.
                </p>
              </div>
            </div>

            <div class="account-stack-list">
              ${book.portfolios.map(renderPortfolioSection).join("")}
            </div>
          </section>
        </section>
      </section>
    </section>
  `;

  hydrateSidebarNavigation();
  hydratePortfolioCircleInteractions();
  hydrateTreemapInteractions();
  scheduleTreemapTypographyFit();

  if (document.fonts?.ready) {
    document.fonts.ready.then(() => {
      scheduleTreemapTypographyFit();
    });
  }
}

async function refreshLivePrices() {
  if (!state.portfolioSource || state.isRefreshing) {
    return;
  }

  const externalTickers = getExternalQuoteTickers(state.portfolioSource);
  state.isRefreshing = true;
  state.refreshError = null;
  state.refreshPhase = "external";
  state.refreshDetail =
    "Yahoo Finance 시세와 환율을 직접 확인하고, 실패하면 저장된 가격 파일을 확인합니다.";
  renderCurrentBook();

  try {
    const [livePriceResult, exchangeRateResult] = await Promise.allSettled([
      loadLivePrices({
        includeRepo: true,
        externalTickers,
      }),
      loadLiveExchangeRateSource(),
    ]);

    if (exchangeRateResult.status === "fulfilled") {
      state.liveExchangeRateSource = exchangeRateResult.value;
    }

    if (livePriceResult.status === "fulfilled") {
      state.livePriceSource = livePriceResult.value;
    } else {
      throw livePriceResult.reason;
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
    [state.portfolioSource, state.dividendSource] = await Promise.all([
      fetchJson(PORTFOLIOS_URL),
      fetchDividendData(),
    ]);
    try {
      state.dividendCalendarSource = await fetchDividendCalendarData();
      state.dividendCalendarError = null;
    } catch (error) {
      state.dividendCalendarSource = null;
      state.dividendCalendarError = error;
    }
    const externalTickers = getExternalQuoteTickers(state.portfolioSource);

    const [livePriceResult, exchangeRateResult] = await Promise.allSettled([
      loadLivePrices({
        includeRepo: true,
        externalTickers,
      }),
      loadLiveExchangeRateSource(),
    ]);

    if (livePriceResult.status === "fulfilled") {
      state.livePriceSource = livePriceResult.value;
    } else {
      state.livePriceSource = null;
      state.refreshError = livePriceResult.reason;
    }

    if (exchangeRateResult.status === "fulfilled") {
      state.liveExchangeRateSource = exchangeRateResult.value;
    } else {
      state.liveExchangeRateSource = null;
    }

    renderApp(buildBookData(state.portfolioSource, state.livePriceSource, state.dividendSource));
    startAutoRefresh();
  } catch (error) {
    renderError(error);
  }
}

document.addEventListener("click", (event) => {
  const amountVisibilityButton = event.target.closest("[data-toggle-amount-visibility]");

  if (amountVisibilityButton) {
    state.hideAmounts = !state.hideAmounts;
    persistAmountVisibilityPreference();
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

  refreshLivePrices();
});

window.addEventListener("resize", () => {
  if (treemapResizeFrameId) {
    window.cancelAnimationFrame(treemapResizeFrameId);
  }

  treemapResizeFrameId = window.requestAnimationFrame(() => {
    treemapResizeFrameId = 0;

    if (!state.portfolioSource) {
      return;
    }

    renderCurrentBook();
  });
});

initializeApp();
