#!/usr/bin/env bash
set -euo pipefail

PROJECT_ID="${PROJECT_ID:-monospace-6}"
REGION="${REGION:-asia-northeast3}"
SERVICE_NAME="${SERVICE_NAME:-portfolio-app}"
GITHUB_OWNER="${GITHUB_OWNER:-gusdnr0226}"
GITHUB_REPO="${GITHUB_REPO:-portfolio}"
GITHUB_BRANCH="${GITHUB_BRANCH:-main}"
GITHUB_PORTFOLIO_PATH="${GITHUB_PORTFOLIO_PATH:-data/portfolios.json}"
AUTO_COMMIT_ON_UPLOAD="${AUTO_COMMIT_ON_UPLOAD:-true}"
SECRET_NAME="${SECRET_NAME:-portfolio-github-token}"

if ! command -v gcloud >/dev/null 2>&1; then
  echo "gcloud CLI가 필요합니다." >&2
  exit 1
fi

gcloud services enable   run.googleapis.com   cloudbuild.googleapis.com   artifactregistry.googleapis.com   secretmanager.googleapis.com   --project="$PROJECT_ID"   --quiet

DEPLOY_ARGS=(
  run deploy "$SERVICE_NAME"
  --project="$PROJECT_ID"
  --region="$REGION"
  --source=.
  --allow-unauthenticated
  --quiet
  --set-env-vars="GITHUB_OWNER=$GITHUB_OWNER,GITHUB_REPO=$GITHUB_REPO,GITHUB_BRANCH=$GITHUB_BRANCH,GITHUB_PORTFOLIO_PATH=$GITHUB_PORTFOLIO_PATH,AUTO_COMMIT_ON_UPLOAD=$AUTO_COMMIT_ON_UPLOAD"
)

if [[ -n "${GITHUB_TOKEN:-}" ]]; then
  if gcloud secrets describe "$SECRET_NAME" --project="$PROJECT_ID" >/dev/null 2>&1; then
    printf '%s' "$GITHUB_TOKEN" | gcloud secrets versions add "$SECRET_NAME"       --project="$PROJECT_ID"       --data-file=-
  else
    printf '%s' "$GITHUB_TOKEN" | gcloud secrets create "$SECRET_NAME"       --project="$PROJECT_ID"       --replication-policy="automatic"       --data-file=-
  fi

  DEPLOY_ARGS+=(--update-secrets="GITHUB_TOKEN=${SECRET_NAME}:latest")
fi

gcloud "${DEPLOY_ARGS[@]}"
