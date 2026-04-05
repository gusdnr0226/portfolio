import { createServer } from "node:http";
import { readFile, stat, writeFile } from "node:fs/promises";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

loadDotEnv(path.join(__dirname, ".env"));

const PORT = Number.parseInt(process.env.PORT ?? "8000", 10);
const GITHUB_OWNER = process.env.GITHUB_OWNER ?? "gusdnr0226";
const GITHUB_REPO = process.env.GITHUB_REPO ?? "portfolio";
const GITHUB_BRANCH = process.env.GITHUB_BRANCH ?? "main";
const GITHUB_PORTFOLIO_PATH = process.env.GITHUB_PORTFOLIO_PATH ?? "data/portfolios.json";
const GITHUB_TOKEN = process.env.GITHUB_TOKEN ?? "";
const GITHUB_COMMITTER_NAME = process.env.GITHUB_COMMITTER_NAME ?? "portfolio-sync-bot";
const GITHUB_COMMITTER_EMAIL = process.env.GITHUB_COMMITTER_EMAIL ?? "portfolio-sync-bot@users.noreply.github.com";
const AUTO_COMMIT_ON_UPLOAD = (process.env.AUTO_COMMIT_ON_UPLOAD ?? "true").toLowerCase() !== "false";
const MAX_BODY_BYTES = 2 * 1024 * 1024;
const REPO_FULL_NAME = `${GITHUB_OWNER}/${GITHUB_REPO}`;
const CONTENTS_API_BASE_URL = `https://api.github.com/repos/${REPO_FULL_NAME}/contents/${encodeURIComponentPath(GITHUB_PORTFOLIO_PATH)}`;
const CONTENTS_API_GET_URL = `${CONTENTS_API_BASE_URL}?ref=${encodeURIComponent(GITHUB_BRANCH)}`;
const LOCAL_PORTFOLIO_PATH = path.join(__dirname, GITHUB_PORTFOLIO_PATH);

const CONTENT_TYPES = new Map([
  [".html", "text/html; charset=utf-8"],
  [".js", "application/javascript; charset=utf-8"],
  [".mjs", "application/javascript; charset=utf-8"],
  [".css", "text/css; charset=utf-8"],
  [".json", "application/json; charset=utf-8"],
  [".png", "image/png"],
  [".jpg", "image/jpeg"],
  [".jpeg", "image/jpeg"],
  [".svg", "image/svg+xml"],
  [".ico", "image/x-icon"],
  [".txt", "text/plain; charset=utf-8"],
  [".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"],
  [".xls", "application/vnd.ms-excel"],
]);

function loadDotEnv(filePath) {
  try {
    const source = readFileSync(filePath, "utf8");

    source.split(/\r?\n/).forEach((line) => {
      const trimmed = line.trim();

      if (!trimmed || trimmed.startsWith("#")) {
        return;
      }

      const index = trimmed.indexOf("=");

      if (index < 0) {
        return;
      }

      const key = trimmed.slice(0, index).trim();

      if (!key || process.env[key] !== undefined) {
        return;
      }

      let value = trimmed.slice(index + 1).trim();

      if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
        value = value.slice(1, -1);
      }

      process.env[key] = value;
    });
  } catch {
    // Ignore missing .env in local/dev environments.
  }
}

function encodeURIComponentPath(value) {
  return value
    .split("/")
    .map((segment) => encodeURIComponent(segment))
    .join("/");
}

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
  });
  response.end(JSON.stringify(payload, null, 2));
}

function sendText(response, statusCode, message) {
  response.writeHead(statusCode, {
    "Content-Type": "text/plain; charset=utf-8",
    "Cache-Control": "no-store",
  });
  response.end(message);
}

function getSyncStatus() {
  if (!GITHUB_TOKEN) {
    return {
      enabled: false,
      mode: "download-only",
      reason: "missing-token",
      repoFullName: REPO_FULL_NAME,
      branch: GITHUB_BRANCH,
      path: GITHUB_PORTFOLIO_PATH,
      autoCommitOnUpload: AUTO_COMMIT_ON_UPLOAD,
    };
  }

  return {
    enabled: true,
    mode: "github-commit",
    reason: null,
    repoFullName: REPO_FULL_NAME,
    branch: GITHUB_BRANCH,
    path: GITHUB_PORTFOLIO_PATH,
    autoCommitOnUpload: AUTO_COMMIT_ON_UPLOAD,
  };
}

async function readRequestBody(request) {
  let receivedBytes = 0;
  const chunks = [];

  for await (const chunk of request) {
    receivedBytes += chunk.length;

    if (receivedBytes > MAX_BODY_BYTES) {
      throw new Error("업로드 본문이 너무 큽니다. 2MB 이하 요청만 허용합니다.");
    }

    chunks.push(chunk);
  }

  return Buffer.concat(chunks).toString("utf8");
}

function buildNormalizedJsonText(value) {
  return `${JSON.stringify(value, null, 2)}\n`;
}

function createCommitMessage(payload) {
  const prefix = payload?.sourceFileName
    ? `[portfolio-sync] ${payload.sourceFileName}`
    : "[portfolio-sync] workbook upload";
  const requestedAt = payload?.requestedAt ? ` @ ${payload.requestedAt}` : "";
  return `${prefix}${requestedAt}`;
}

async function fetchGitHubJson(url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${GITHUB_TOKEN}`,
      "User-Agent": "portfolio-sync-server",
      ...(options.headers ?? {}),
    },
  });

  const payload = await response.json().catch(() => null);

  if (!response.ok) {
    const detail = payload?.message ?? `GitHub API 요청 실패 (${response.status})`;
    const error = new Error(detail);
    error.status = response.status;
    throw error;
  }

  return payload;
}

async function loadRemotePortfolioFile() {
  try {
    return await fetchGitHubJson(CONTENTS_API_GET_URL);
  } catch (error) {
    if (error?.status === 404) {
      return null;
    }

    throw error;
  }
}

async function updateRemotePortfolioFile(nextJsonText, payload) {
  const currentFile = await loadRemotePortfolioFile();
  const currentContent = currentFile?.content
    ? Buffer.from(currentFile.content, "base64").toString("utf8")
    : null;

  if (currentContent === nextJsonText) {
    return {
      skipped: true,
      message: "GitHub 반영 생략",
      detail: "저장소의 data/portfolios.json 내용과 동일해서 새 커밋을 만들지 않았습니다.",
      committedAt: new Date().toISOString(),
      repoFullName: REPO_FULL_NAME,
      branch: GITHUB_BRANCH,
      path: GITHUB_PORTFOLIO_PATH,
      commitSha: currentFile?.sha ?? null,
      commitUrl: null,
      localWriteApplied: false,
    };
  }

  const requestBody = {
    message: createCommitMessage(payload),
    content: Buffer.from(nextJsonText, "utf8").toString("base64"),
    branch: GITHUB_BRANCH,
    committer: {
      name: GITHUB_COMMITTER_NAME,
      email: GITHUB_COMMITTER_EMAIL,
    },
  };

  if (currentFile?.sha) {
    requestBody.sha = currentFile.sha;
  }

  const updateResponse = await fetchGitHubJson(CONTENTS_API_BASE_URL, {
    method: "PUT",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(requestBody),
  });

  const commitSha = updateResponse.commit?.sha ?? null;
  const commitUrl = updateResponse.commit?.html_url ?? (
    commitSha ? `https://github.com/${REPO_FULL_NAME}/commit/${commitSha}` : null
  );

  return {
    skipped: false,
    message: "GitHub 반영 완료",
    detail: `${REPO_FULL_NAME}@${GITHUB_BRANCH}:${GITHUB_PORTFOLIO_PATH} 업데이트를 커밋했습니다.`,
    committedAt: new Date().toISOString(),
    repoFullName: REPO_FULL_NAME,
    branch: GITHUB_BRANCH,
    path: GITHUB_PORTFOLIO_PATH,
    commitSha,
    commitUrl,
    localWriteApplied: false,
  };
}

async function tryWriteLocalPortfolioFile(nextJsonText) {
  try {
    await writeFile(LOCAL_PORTFOLIO_PATH, nextJsonText, "utf8");
    return true;
  } catch {
    return false;
  }
}

async function handlePortfolioCommit(request, response) {
  if (!GITHUB_TOKEN) {
    sendJson(response, 503, {
      message: "GITHUB_TOKEN이 없어 자동 GitHub 반영을 사용할 수 없습니다.",
      detail: ".env에 GITHUB_TOKEN을 설정한 뒤 서버를 다시 시작해 주세요.",
    });
    return;
  }

  let bodyText = "";

  try {
    bodyText = await readRequestBody(request);
  } catch (error) {
    sendJson(response, 413, {
      message: error instanceof Error ? error.message : String(error),
    });
    return;
  }

  let payload;

  try {
    payload = JSON.parse(bodyText || "{}");
  } catch {
    sendJson(response, 400, {
      message: "JSON 본문 파싱에 실패했습니다.",
    });
    return;
  }

  const portfolioSource = payload?.portfolioSource;

  if (!portfolioSource || !Array.isArray(portfolioSource.accounts)) {
    sendJson(response, 400, {
      message: "portfolioSource.accounts 배열이 필요합니다.",
    });
    return;
  }

  const nextJsonText = buildNormalizedJsonText(portfolioSource);

  try {
    const result = await updateRemotePortfolioFile(nextJsonText, payload);
    result.localWriteApplied = await tryWriteLocalPortfolioFile(nextJsonText);
    sendJson(response, 200, result);
  } catch (error) {
    sendJson(response, 502, {
      message: error instanceof Error ? error.message : String(error),
      detail: "GitHub Contents API 반영에 실패했습니다.",
    });
  }
}

async function serveStaticFile(requestUrl, response, method) {
  const pathname = requestUrl.pathname === "/" ? "/index.html" : requestUrl.pathname;
  const decodedPath = decodeURIComponent(pathname);
  const filePath = path.resolve(__dirname, `.${decodedPath}`);
  const allowedRoot = `${__dirname}${path.sep}`;

  if (filePath !== __dirname && !filePath.startsWith(allowedRoot)) {
    sendText(response, 403, "Forbidden");
    return;
  }

  try {
    const fileStat = await stat(filePath);

    if (!fileStat.isFile()) {
      sendText(response, 404, "Not found");
      return;
    }

    const contentType = CONTENT_TYPES.get(path.extname(filePath).toLowerCase()) ?? "application/octet-stream";
    const headers = {
      "Content-Type": contentType,
      "Cache-Control": decodedPath.startsWith("/data/") ? "no-store" : "public, max-age=60",
    };

    if (method === "HEAD") {
      response.writeHead(200, headers);
      response.end();
      return;
    }

    const fileBuffer = await readFile(filePath);
    response.writeHead(200, headers);
    response.end(fileBuffer);
  } catch {
    sendText(response, 404, "Not found");
  }
}

const server = createServer(async (request, response) => {
  try {
    const requestUrl = new URL(
      request.url ?? "/",
      `http://${request.headers.host ?? `127.0.0.1:${PORT}`}`,
    );

    if (request.method === "GET" && requestUrl.pathname === "/api/portfolio/sync-status") {
      sendJson(response, 200, getSyncStatus());
      return;
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/portfolio/commit") {
      await handlePortfolioCommit(request, response);
      return;
    }

    if (!["GET", "HEAD"].includes(request.method ?? "GET")) {
      sendText(response, 405, "Method not allowed");
      return;
    }

    await serveStaticFile(requestUrl, response, request.method ?? "GET");
  } catch (error) {
    sendJson(response, 500, {
      message: error instanceof Error ? error.message : String(error),
    });
  }
});

server.listen(PORT, () => {
  console.log(`portfolio server listening on http://127.0.0.1:${PORT}`);
  console.log(`portfolio sync mode: ${getSyncStatus().enabled ? "github-commit" : "download-only"}`);
});

