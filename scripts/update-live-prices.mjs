import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, "..");
const portfolioPath = path.join(rootDir, "data", "portfolios.json");
const livePricesPath = path.join(rootDir, "data", "live-prices.json");

const KIS_BASE_URL =
  process.env.KIS_BASE_URL ?? "https://openapi.koreainvestment.com:9443";

function uniqueHoldings(accounts) {
  const seen = new Map();

  for (const account of accounts) {
    for (const holding of account.holdings) {
      if (!seen.has(holding.ticker)) {
        seen.set(holding.ticker, holding);
      }
    }
  }

  return [...seen.values()];
}

async function readJson(filePath) {
  return JSON.parse(await readFile(filePath, "utf8"));
}

async function fetchAccessToken(appKey, appSecret) {
  const response = await fetch(`${KIS_BASE_URL}/oauth2/tokenP`, {
    method: "POST",
    headers: {
      "content-type": "application/json"
    },
    body: JSON.stringify({
      grant_type: "client_credentials",
      appkey: appKey,
      appsecret: appSecret
    })
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch KIS access token: ${response.status}`);
  }

  const payload = await response.json();

  if (!payload.access_token) {
    throw new Error("KIS access token response did not include access_token");
  }

  return payload.access_token;
}

async function fetchCurrentPrice({ accessToken, appKey, appSecret, ticker }) {
  const url = new URL(
    `${KIS_BASE_URL}/uapi/domestic-stock/v1/quotations/inquire-price`
  );
  url.searchParams.set("fid_cond_mrkt_div_code", "J");
  url.searchParams.set("fid_input_iscd", ticker);

  const response = await fetch(url, {
    headers: {
      authorization: `Bearer ${accessToken}`,
      appkey: appKey,
      appsecret: appSecret,
      tr_id: "FHKST01010100"
    }
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch quote for ${ticker}: ${response.status}`);
  }

  const payload = await response.json();
  const rawPrice = payload?.output?.stck_prpr;

  if (!rawPrice) {
    throw new Error(`KIS quote response for ${ticker} did not include stck_prpr`);
  }

  return Number(rawPrice);
}

function buildFallbackQuotes(holdings, existingQuotes) {
  return Object.fromEntries(
    holdings.map((holding) => [
      holding.ticker,
      {
        price: existingQuotes?.[holding.ticker]?.price ?? holding.snapshotPrice
      }
    ])
  );
}

async function main() {
  const portfolioSource = await readJson(portfolioPath);
  const existingLivePrices = await readJson(livePricesPath);
  const holdings = uniqueHoldings(portfolioSource.accounts);

  const appKey = process.env.KIS_APP_KEY;
  const appSecret = process.env.KIS_APP_SECRET;

  if (!appKey || !appSecret) {
    console.log(
      "KIS_APP_KEY or KIS_APP_SECRET is missing. live-prices.json will remain unchanged."
    );
    return;
  }

  const accessToken = await fetchAccessToken(appKey, appSecret);
  const quotes = {};

  for (const holding of holdings) {
    try {
      const price = await fetchCurrentPrice({
        accessToken,
        appKey,
        appSecret,
        ticker: holding.ticker
      });

      quotes[holding.ticker] = { price };
    } catch (error) {
      console.warn(
        `Falling back to stored price for ${holding.ticker}: ${error.message}`
      );
      quotes[holding.ticker] = {
        price:
          existingLivePrices?.quotes?.[holding.ticker]?.price ?? holding.snapshotPrice
      };
    }
  }

  const payload = {
    provider: "KIS Open API",
    updatedAt: new Date().toISOString(),
    quotes: Object.keys(quotes)
      .sort()
      .reduce((accumulator, ticker) => {
        accumulator[ticker] = quotes[ticker];
        return accumulator;
      }, {})
  };

  await writeFile(livePricesPath, `${JSON.stringify(payload, null, 2)}\n`);
  console.log("Updated data/live-prices.json");
}

main().catch(async (error) => {
  console.error(error);

  const portfolioSource = await readJson(portfolioPath);
  const existingLivePrices = await readJson(livePricesPath);
  const holdings = uniqueHoldings(portfolioSource.accounts);
  const fallbackPayload = {
    provider: existingLivePrices?.provider ?? "snapshot-fallback",
    updatedAt: existingLivePrices?.updatedAt ?? null,
    quotes: buildFallbackQuotes(holdings, existingLivePrices?.quotes)
  };

  await writeFile(livePricesPath, `${JSON.stringify(fallbackPayload, null, 2)}\n`);
  process.exitCode = 1;
});
