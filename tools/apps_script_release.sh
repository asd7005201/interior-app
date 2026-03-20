#!/usr/bin/env bash
set -euo pipefail

APP_NAME="${1:-}"
DEPLOYMENT_ID="${2:-}"

if [[ -z "$APP_NAME" ]]; then
  echo "usage: tools/apps_script_release.sh <quote|prequote> [deployment_id]" >&2
  exit 1
fi

case "$APP_NAME" in
  quote)
    APP_DIR="/home/mifasol/interior-app/quote_app"
    ;;
  prequote)
    APP_DIR="/home/mifasol/interior-app/prequote_app"
    ;;
  *)
    echo "unknown app: $APP_NAME" >&2
    exit 1
    ;;
esac

cd "$APP_DIR"

echo "[release] push $APP_NAME"
npx clasp push

echo "[release] deploy $APP_NAME"
if [[ -n "$DEPLOYMENT_ID" ]]; then
  npx clasp deploy --deploymentId "$DEPLOYMENT_ID" --description "codex release $(date '+%Y-%m-%d %H:%M:%S')"
else
  npx clasp deploy --description "codex release $(date '+%Y-%m-%d %H:%M:%S')"
fi
