const PORTFOLIOS_URL = "./data/portfolios.json";
const LIVE_PRICES_URL = "./data/live-prices.json";
const GITHUB_API_BASE_URL = "https://api.github.com/repos/gusdnr0226/portfolio";
const GITHUB_LIVE_PRICES_URL =
  `${GITHUB_API_BASE_URL}/contents/data/live-prices.json?ref=main`;
const LIVE_PRICE_PROXY_BASE_URL = "https://r.jina.ai/http://";
const YAHOO_SPARK_BASE_URL = "https://query1.finance.yahoo.com/v7/finance/spark";
const YAHOO_SPARK_BATCH_SIZE = 20;
const REFRESH_INTERVAL_MS = 5 * 60 * 1000;
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
const CATEGORY_TICKER_RULES = {
  "SCHD": [
    "SCHD",
    "TIGER 미국배당다우존스",
    "SOL 미국배당다우존스2호",
    "SOL 미국배당다우존스",
  ],
  "Nasdaq": [
    "QQQ",
    "KODEX 미국나스닥100",
    "TIGER 미국나스닥100레버리지(합성)",
    "KODEX 미국반도체",
    "KODEX 미국휴머노이드로봇",
  ],
  "S&P500": [
    "VOO",
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
  "SCHD": "#77d7ff",
  "S&P500": "#86abff",
  Nasdaq: "#8e93ff",
  "커버드콜": "#ff8ca1",
  "부동산/리츠": "#5dd4c0",
  "채권": "#77c8d2",
  USD: "#ffd66b",
  "배당ETF": "#8adc78",
  "인프라": "#57d7ae",
  "통신": "#68a8ff",
  "제조": "#ffb067",
  "금융": "#d7a0ff",
  "원화": "#aab8d6",
  MM: "#ffc887",
  "기타": "#98a4c3",
};

const currencyFormatter = new Intl.NumberFormat("ko-KR");
const compactCurrencyFormatter = new Intl.NumberFormat("ko-KR", {
  notation: "compact",
  maximumFractionDigits: 1,
});
const quantityFormatter = new Intl.NumberFormat("ko-KR", {
  minimumFractionDigits: 0,
  maximumFractionDigits: 6,
});
const decimalFormatter = new Intl.NumberFormat("ko-KR", {
  minimumFractionDigits: 0,
  maximumFractionDigits: 1,
});
const usdCurrencyFormatter = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
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

function formatCurrency(value) {
  return `${currencyFormatter.format(Math.round(value))}원`;
}

function formatCompactCurrency(value) {
  return `${compactCurrencyFormatter.format(Math.round(value))}원`;
}

function formatSignedCurrency(value) {
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

function formatMoney(value, currency = "KRW") {
  if (currency === "USD") {
    return usdCurrencyFormatter.format(value);
  }

  return formatCurrency(value);
}

function formatSignedMoney(value, currency = "KRW") {
  if (currency === "USD") {
    const prefix = value > 0 ? "+" : value < 0 ? "-" : "";
    return `${prefix}${usdCurrencyFormatter.format(Math.abs(value))}`;
  }

  return formatSignedCurrency(value);
}

function formatPrice(value, currency = "KRW") {
  return formatMoney(value, currency);
}

function formatAverageCost(value, currency = "KRW") {
  if (currency === "USD") {
    return usdCurrencyFormatter.format(value);
  }

  const isWhole = Math.abs(value - Math.round(value)) < 0.05;
  const displayValue = isWhole ? Math.round(value) : Number(value.toFixed(1));
  return `${decimalFormatter.format(displayValue)}원`;
}

function formatQuantity(value) {
  const roundedValue = Number(value.toFixed(6));
  return `${quantityFormatter.format(roundedValue)}주`;
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

async function fetchLocalLivePrices() {
  const payload = await fetchJson(LIVE_PRICES_URL);
  return normalizeLivePriceSource(payload, "deployed-file");
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

  renderApp(buildBookData(state.portfolioSource, state.livePriceSource));
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

function buildPortfolioData(source, livePriceSource) {
  const quotes = livePriceSource?.quotes ?? {};
  const isFallbackSource = livePriceSource?.provider === "snapshot-fallback";
  const hasAutomaticQuotes = isAutomaticLivePriceSource(livePriceSource);
  const usdKrwRate = getUsdKrwRate(source, quotes);

  const enrichedHoldings = source.holdings.map((holding) => {
    const currency = getHoldingCurrency(holding);
    const storedQuote = !holding.manualPriceOnly ? quotes[holding.ticker] : null;
    const shouldUseStoredQuote =
      hasUsableQuote(storedQuote) &&
      (hasAutomaticQuotes || !isFallbackSource || storedQuote.price !== holding.snapshotPrice);
    const currentPrice = shouldUseStoredQuote ? storedQuote.price : holding.snapshotPrice;
    const marketValue = currentPrice * holding.quantity;
    const profitLoss = marketValue - holding.costBasis;
    const marketValueBase = convertToBaseCurrency(marketValue, currency, usdKrwRate);
    const costBasisBase = convertToBaseCurrency(holding.costBasis, currency, usdKrwRate);
    const profitLossBase = marketValueBase - costBasisBase;
    const profitRate = holding.costBasis === 0 ? 0 : profitLoss / holding.costBasis;

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
      accumulator.marketValue += holding.marketValueBase;
      accumulator.costBasis += holding.costBasisBase;
      accumulator.profitLoss += holding.profitLossBase;
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
    assetWeight: totalAssets === 0 ? 0 : holding.marketValueBase / totalAssets,
  }));
  const holdingCurrencies = [...new Set(holdingsWithWeight.map((holding) => holding.currency))];

  return {
    ...source,
    holdings: holdingsWithWeight,
    hasForeignCurrency: holdingCurrencies.some((currency) => currency !== "KRW"),
    usdKrwRate,
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
    hasForeignCurrency: portfolios.some((portfolio) => portfolio.hasForeignCurrency),
    portfolios,
    overview,
  };
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
        accountNames: new Set(),
      };

      existingAsset.marketValueBase += holding.marketValueBase;
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
        accountNames: new Set(),
      };

      existingCash.marketValueBase += portfolio.cash;
      existingCash.accountNames.add(portfolio.name);
      assetMap.set(key, existingCash);
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
      categories: new Map(),
    };
    const existingCategory = existingSuperCategory.categories.get(asset.category) ?? {
      id: `${existingSuperCategory.id}-${normalizeTreemapToken(asset.category).replace(/\s+/g, "-")}`,
      name: asset.category,
      superCategory: asset.superCategory,
      marketValueBase: 0,
      assets: [],
    };

    existingCategory.marketValueBase += asset.marketValueBase;
    existingCategory.assets.push(asset);
    existingSuperCategory.marketValueBase += asset.marketValueBase;
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

function findTreemapSplitIndex(items) {
  const totalValue = items.reduce((sum, item) => sum + item.value, 0);
  let runningValue = 0;
  let bestIndex = 1;
  let bestDifference = Number.POSITIVE_INFINITY;

  for (let index = 1; index < items.length; index += 1) {
    runningValue += items[index - 1].value;
    const difference = Math.abs(totalValue / 2 - runningValue);

    if (difference < bestDifference) {
      bestDifference = difference;
      bestIndex = index;
    }
  }

  return bestIndex;
}

function buildTreemapLayouts(
  items,
  rect = { x: 0, y: 0, width: 100, height: 100 },
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
      },
    ];
  }

  const totalValue = filteredItems.reduce((sum, item) => sum + item.value, 0);
  const splitIndex = findTreemapSplitIndex(filteredItems);
  const firstItems = filteredItems.slice(0, splitIndex);
  const secondItems = filteredItems.slice(splitIndex);
  const firstValue = firstItems.reduce((sum, item) => sum + item.value, 0);
  const firstRatio = firstValue / totalValue;
  const splitVertically = rect.width >= rect.height;
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
    ...buildTreemapLayouts(firstItems, firstRect),
    ...buildTreemapLayouts(secondItems, secondRect),
  ];
}

function hexToRgba(hex, alpha) {
  const normalizedHex = hex.replace("#", "");
  const red = Number.parseInt(normalizedHex.slice(0, 2), 16);
  const green = Number.parseInt(normalizedHex.slice(2, 4), 16);
  const blue = Number.parseInt(normalizedHex.slice(4, 6), 16);

  return `rgba(${red}, ${green}, ${blue}, ${alpha})`;
}

function getTreemapAccent(category, superCategory) {
  return (
    TREEMAP_CATEGORY_COLORS[category] ??
    (superCategory === "해외" ? "#62a8ff" : "#70d7a5")
  );
}

function buildTreemapStyle(layout, accent) {
  return [
    `left:${layout.x.toFixed(4)}%`,
    `top:${layout.y.toFixed(4)}%`,
    `width:${layout.width.toFixed(4)}%`,
    `height:${layout.height.toFixed(4)}%`,
    `--treemap-accent:${accent}`,
    `--treemap-accent-soft:${hexToRgba(accent, 0.13)}`,
    `--treemap-accent-strong:${hexToRgba(accent, 0.24)}`,
    `--treemap-outline:${hexToRgba(accent, 0.34)}`,
  ].join(";");
}

function getTreemapTileSizeClass(layout) {
  const area = layout.width * layout.height;

  if (area >= 1400) {
    return "tile-xl";
  }

  if (area >= 650) {
    return "tile-lg";
  }

  if (area >= 240) {
    return "tile-md";
  }

  if (area >= 90) {
    return "tile-sm";
  }

  return "tile-xs";
}

function renderTreemapAssetTile(layout) {
  const asset = layout.node;
  const sizeClass = getTreemapTileSizeClass(layout);
  const accent = getTreemapAccent(asset.category, asset.superCategory);
  const positionType =
    asset.type === "현금"
      ? asset.accountCount > 1
        ? `${asset.accountCount}계좌 현금 합산`
        : "현금"
      : asset.accountCount > 1
        ? `${asset.accountCount}계좌 합산`
        : asset.type;

  return `
    <article class="treemap-tile ${sizeClass}" style="${buildTreemapStyle(layout, accent)}">
      <div class="treemap-tile-inner">
        <div class="treemap-tile-head">
          <strong class="treemap-tile-name">${asset.name}</strong>
          ${asset.currency !== "KRW" ? `<span class="treemap-tile-badge">${asset.currency}</span>` : ""}
        </div>
        <div class="treemap-tile-body">
          <span class="treemap-tile-value">${formatCompactCurrency(asset.marketValueBase)}</span>
          <span class="treemap-tile-meta">${formatPercent(asset.weight)}</span>
          <span class="treemap-tile-meta">${positionType}</span>
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
  );

  return `
    <article class="treemap-category ${sizeClass}" style="${buildTreemapStyle(layout, accent)}">
      <div class="treemap-category-head">
        <strong>${category.name}</strong>
        <span>${formatCompactCurrency(category.marketValueBase)} · ${formatPercent(category.weight)}</span>
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
  );

  return `
    <section class="treemap-super ${sizeClass}" style="${buildTreemapStyle(layout, accent)}">
      <div class="treemap-super-head">
        <div>
          <p class="treemap-super-label">${superCategory.name}</p>
          <strong class="treemap-super-value">${formatCompactCurrency(superCategory.marketValueBase)}</strong>
        </div>
        <span class="treemap-super-share">${formatPercent(superCategory.weight)}</span>
      </div>
      <div class="treemap-super-body">
        ${categoryLayouts.map(renderTreemapCategoryTile).join("")}
      </div>
    </section>
  `;
}

function renderTreemapSection(book) {
  const treemap = buildTreemapData(book);

  if (!treemap || !treemap.superCategories.length) {
    return "";
  }

  const superCategoryLayouts = buildTreemapLayouts(
    treemap.superCategories.map((superCategory) => ({
      node: superCategory,
      value: superCategory.marketValueBase,
    })),
  );

  return `
    <section class="treemap-card">
      <div class="section-head">
        <div>
          <h2 class="section-title">전체 종목 트리맵</h2>
          <p class="section-copy">
            면적은 전체 포트폴리오 내 원화 환산 평가금액 비중입니다. 해외와 국내를 먼저 나누고, 그 안에서 주신 분류 규칙 기준으로 카테고리와 종목을 묶었습니다.
            같은 종목을 여러 계좌에서 들고 있으면 한 타일로 합산해서 보여줍니다.
          </p>
        </div>
        <div class="section-pill">총 ${currencyFormatter.format(treemap.assetCount)}자산</div>
      </div>

      <div class="treemap-wrap">
        <div class="treemap-canvas">
          ${superCategoryLayouts.map(renderTreemapSuperCategoryTile).join("")}
        </div>
      </div>

      <div class="treemap-footnote">
        <span>면적 = KRW 환산 평가금액</span>
        <span>중복 보유 = 계좌별 합산</span>
        <span>현금 = 원화 현금 타일로 별도 집계</span>
      </div>
    </section>
  `;
}

function buildMarketStatus(book) {
  const hasStoredQuotes = hasAppliedStoredQuotes(book);

  if (state.isRefreshing) {
    return {
      tone: "status-loading",
      title: "외부 시세 다시 불러오는 중",
      detail:
        state.refreshDetail ??
        "Yahoo Finance 시세를 직접 확인하고, 실패하면 저장된 가격 파일을 사용합니다.",
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
      title: "외부 시세 자동 반영 중",
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
      return "외부 시세 직접 반영";
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
      detail: book.hasForeignCurrency
        ? "현재가와 환율만 자동 갱신, 수량/평단은 원장 기준"
        : "현재가만 외부 시세로 갱신, 수량/평단은 원장 기준",
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

function renderAccountOverviewCards(book) {
  return book.portfolios
    .map((portfolio) => {
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
              ${holding.currency !== "KRW" ? `<span class="tag">${holding.currency}</span>` : ""}
              <span class="tag ${holding.priceSource === "snapshot" || holding.priceSource === "manual" ? "tag-warning" : "tag-live"}">
                ${getPriceSourceLabel(holding.priceSource)}
              </span>
              ${holding.nameInferred ? '<span class="tag tag-warning">이름 추정</span>' : ""}
            </div>
          </td>
          <td>${formatQuantity(holding.quantity)}</td>
          <td>
            <div class="metric">
              <strong>${formatPrice(holding.currentPrice, holding.currency)}</strong>
              <span>${getPriceSourceDetail(holding.priceSource)}</span>
            </div>
          </td>
          <td>${formatAverageCost(holding.averageCost, holding.currency)}</td>
          <td>${formatMoney(holding.costBasis, holding.currency)}</td>
          <td>${formatMoney(holding.marketValue, holding.currency)}</td>
          <td class="${getToneClass(holding.profitLoss)}">${formatSignedMoney(holding.profitLoss, holding.currency)}</td>
          <td class="${getToneClass(holding.profitRate)}">${formatSignedPercent(holding.profitRate)}</td>
          <td>${formatPercent(holding.assetWeight)}</td>
        </tr>
      `,
    )
    .join("");
}

function renderPortfolioSection(portfolio) {
  const heroCopy = portfolio.hasForeignCurrency
    ? `${portfolio.snapshotLabel} 원장 데이터를 유지하고, 현재가와 환율만 외부 시세 또는 저장된 최신 시세 파일에서 다시 읽어 계산합니다. 달러 종목은 USD 기준으로 표기하고 계좌 합계와 비중은 KRW 환산으로 계산합니다.`
    : `${portfolio.snapshotLabel} 원장 데이터를 유지하고, 현재가만 외부 시세 또는 저장된 최신 시세 파일에서 다시 읽어 평가금액과 손익을 재계산합니다.`;
  const tableCopy = portfolio.hasForeignCurrency
    ? "수량과 매입금액은 고정 원장 데이터입니다. 해외 종목 금액은 USD 기준으로 표기하고, 계좌 합계와 비중은 KRW 환산 기준으로 계산합니다."
    : "수량과 매입금액은 고정 원장 데이터입니다. 현재가, 평가금액, 손익, 수익률은 최신 시세 기준으로 다시 계산됩니다.";

  return `
    <section class="account-stack" id="${portfolio.id}">
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
                <td>${formatQuantity(portfolio.totals.quantity)}</td>
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

function renderApp(book) {
  const app = document.querySelector("#app");

  if (!app) {
    return;
  }

  const marketStatus = buildMarketStatus(book);

  app.innerHTML = `
    <section class="dashboard">
      <section class="hero">
        <div class="eyebrow">All Accounts</div>
        <div class="hero-head">
          <div>
            <h1 class="hero-title">전체 포트폴리오</h1>
            <p class="hero-copy">
              절세계좌와 직투계좌를 같은 기준으로 묶어 총자산, 손익, 현금, 계좌별 보유 현황을 한 화면에서 점검할 수 있게 구성했습니다.
              미국 종목은 USD 기준으로 표기하고, 계좌와 전체 합계는 KRW 환산 기준으로 계산합니다.
              브라우저에서 외부 시세를 직접 확인하고, 응답이 없으면 저장된 최신 가격 파일을 폴백으로 사용합니다.
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
              ${state.isRefreshing ? "시세 불러오는 중..." : "시세 다시 불러오기"}
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
              전체 종목 트리맵으로 비중을 먼저 보고, 각 계좌의 총자산, 손익, 현금, 투자 비중은 아래 카드와 비교표에서 이어서 확인하면 됩니다.
            </p>
          </div>
        </div>

        ${renderTreemapSection(book)}

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

  const externalTickers = getExternalQuoteTickers(state.portfolioSource);
  state.isRefreshing = true;
  state.refreshError = null;
  state.refreshPhase = "external";
  state.refreshDetail =
    "Yahoo Finance 시세를 직접 확인하고, 실패하면 저장된 가격 파일을 확인합니다.";
  renderCurrentBook();

  try {
    state.livePriceSource = await loadLivePrices({
      includeRepo: true,
      externalTickers,
    });
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
    const externalTickers = getExternalQuoteTickers(state.portfolioSource);

    try {
      state.livePriceSource = await loadLivePrices({
        includeRepo: true,
        externalTickers,
      });
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
