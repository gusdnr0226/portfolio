# Portfolio

정적 포트폴리오 대시보드입니다. 기본 모드는 HTML/CSS/JS 정적 사이트이고, 선택적으로 `node server.mjs`를 함께 띄우면 엑셀 업로드 결과를 GitHub의 `data/portfolios.json`까지 자동 커밋할 수 있습니다.

## 실행

정적 미리보기만 필요하면:

```bash
python3 -m http.server 8000
```

GitHub 자동 커밋까지 쓰려면:

```bash
cp .env.example .env
node server.mjs
```

접속 주소:

```text
http://127.0.0.1:8000/
```

## 환경 변수

`.env.example` 기준으로 아래 값을 설정합니다.

- `GITHUB_TOKEN`: 필수. GitHub fine-grained token 권한은 `Contents: Read and write`가 필요합니다.
- `GITHUB_OWNER`: 기본값 `gusdnr0226`
- `GITHUB_REPO`: 기본값 `portfolio`
- `GITHUB_BRANCH`: 기본값 `main`
- `GITHUB_PORTFOLIO_PATH`: 기본값 `data/portfolios.json`
- `AUTO_COMMIT_ON_UPLOAD`: 기본값 `true`. `false`면 업로드 후 자동 커밋 없이 버튼으로만 반영합니다.

## 동작 방식

1. 브라우저가 엑셀 파일을 읽고 `accounts`, `holdings`, `trades` 시트를 현재 포트폴리오에 반영합니다.
2. 서버 모드가 활성화되어 있으면 브라우저가 최신 포트폴리오 JSON을 `/api/portfolio/commit`으로 전송합니다.
3. 서버가 GitHub Contents API를 호출해 `data/portfolios.json`을 커밋합니다.
4. 가능하면 로컬 서버의 `data/portfolios.json`도 같이 갱신합니다.

## API

- `GET /api/portfolio/sync-status`: 자동 GitHub 반영 가능 여부 확인
- `POST /api/portfolio/commit`: 최신 포트폴리오 JSON을 GitHub에 반영

## 주의

- 정적 호스팅만 사용하면 브라우저 메모리에서만 바뀝니다. 이 경우 JSON 다운로드 파일을 직접 반영해야 합니다.
- `trades` 시트가 포함된 같은 파일을 반복 업로드하면 거래가 다시 누적됩니다.
- GitHub 커밋 후 GitHub Pages 재배포가 끝나야 새로고침 시 최신 `data/portfolios.json`이 내려옵니다.
