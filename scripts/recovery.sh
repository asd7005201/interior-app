#!/usr/bin/env bash
# ============================================================
# recovery.sh — Google 계정 삭제 후 앱 전체 복구 자동화
# ============================================================
# 사용법: bash scripts/recovery.sh
#
# 사전 준비:
#   1. 새 Google 계정 생성
#   2. 새 계정으로 Google Drive에서 Spreadsheet 2개 생성
#      (prequote_DB, quote_DB — 이름은 자유)
#   3. 이 스크립트를 실행하면 나머지를 안내합니다
# ============================================================
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
DATA_FILE="$REPO_ROOT/data/prequote_static.json"

RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'; NC='\033[0m'
step() { echo -e "\n${GREEN}[$1]${NC} $2"; }
warn() { echo -e "${YELLOW}⚠  $1${NC}"; }
err()  { echo -e "${RED}❌ $1${NC}"; exit 1; }

echo "============================================================"
echo "  인테리어 견적 앱 — 복구 스크립트"
echo "============================================================"

# ── 체크: clasp 설치 ──────────────────────────────────────
step "1/7" "clasp 설치 확인"
if ! command -v npx &>/dev/null; then
  err "Node.js / npx가 설치되어 있지 않습니다. 먼저 설치해주세요."
fi
echo "✅ npx OK"

# ── 체크: 백업 데이터 ──────────────────────────────────────
step "2/7" "백업 데이터 확인"
if [ ! -f "$DATA_FILE" ]; then
  err "백업 파일이 없습니다: $DATA_FILE\n먼저 sync_sheets.sh를 실행해 데이터를 저장해두세요."
fi
echo "✅ $DATA_FILE 존재"

# ── Google 계정 로그인 ─────────────────────────────────────
step "3/7" "clasp Google 계정 로그인"
echo "브라우저가 열리면 새 Google 계정으로 로그인하세요."
cd "$REPO_ROOT/prequote_app"
npx clasp login

# ── 스프레드시트 ID 입력 ────────────────────────────────────
step "4/7" "스프레드시트 ID 입력"
echo ""
echo "Google Drive에서 미리 만들어둔 prequote_DB 스프레드시트를 여세요."
echo "URL에서 ID를 복사하세요: https://docs.google.com/spreadsheets/d/【이_부분】/edit"
echo ""
read -rp "prequote_DB Spreadsheet ID: " PREQUOTE_SS_ID
if [ -z "$PREQUOTE_SS_ID" ]; then
  err "Spreadsheet ID를 입력해주세요."
fi

# ── GAS 프로젝트 새로 생성 ─────────────────────────────────
step "5/7" "GAS 프로젝트 생성 및 코드 업로드"
cd "$REPO_ROOT/prequote_app"

# 새 스크립트 생성
NEW_SCRIPT_ID=$(npx clasp create \
  --title "prequote_app" \
  --type sheets \
  --parentId "$PREQUOTE_SS_ID" \
  2>&1 | grep -o '"scriptId":"[^"]*"' | sed 's/"scriptId":"//;s/"//')

if [ -z "$NEW_SCRIPT_ID" ]; then
  warn "scriptId 자동 감지 실패. .clasp.json을 수동으로 확인하세요."
  NEW_SCRIPT_ID=$(python3 -c "import json; d=json.load(open('.clasp.json')); print(d['scriptId'])" 2>/dev/null || echo "")
fi

echo "✅ 새 Script ID: $NEW_SCRIPT_ID"

# 코드 푸시
npx clasp push --force
echo "✅ 코드 업로드 완료"

# Script Properties 설정 (SPREADSHEET_ID)
echo ""
echo "이제 GAS 에디터에서 initializeAppManual('$PREQUOTE_SS_ID') 을 실행하세요."
echo "그래야 앱이 올바른 스프레드시트와 연결됩니다."
echo ""
read -rp "initializeAppManual 실행 완료 후 Enter: " _

# ── 웹 앱 배포 ────────────────────────────────────────────
step "6/7" "웹 앱 배포"
DEPLOY_OUT=$(npx clasp deploy --description "recovery-$(date '+%Y%m%d')" 2>&1)
echo "$DEPLOY_OUT"
NEW_DEPLOY_ID=$(echo "$DEPLOY_OUT" | grep -o 'AKfycb[^ ]*' | head -1)
if [ -n "$NEW_DEPLOY_ID" ]; then
  NEW_URL="https://script.google.com/macros/s/${NEW_DEPLOY_ID}/exec"
  echo ""
  echo "✅ 새 GAS URL:"
  echo "   $NEW_URL"
  # .env 자동 업데이트
  if [ -f "$REPO_ROOT/prequote-web/.env" ]; then
    sed -i "s|VITE_GAS_API_URL=.*|VITE_GAS_API_URL=$NEW_URL|" "$REPO_ROOT/prequote-web/.env"
    echo "✅ prequote-web/.env 업데이트됨"
  fi
fi

# ── 스프레드시트 데이터 복원 안내 ─────────────────────────
step "7/7" "스프레드시트 데이터 복원"
echo ""
echo "백업 JSON을 스프레드시트에 가져오는 방법:"
echo "  python3 scripts/restore_sheets.py \"$PREQUOTE_SS_ID\""
echo ""
echo "또는 수동으로 data/prequote_static.json의 내용을"
echo "각 시트(SurveyQuestions, EstimateRules 등)에 붙여넣기 하세요."
echo ""

# ── 최종 요약 ──────────────────────────────────────────────
echo "============================================================"
echo "  복구 완료 요약"
echo "============================================================"
echo "GAS Script ID : ${NEW_SCRIPT_ID:-수동 확인 필요}"
echo "GAS 배포 URL  : ${NEW_URL:-clasp deploy 출력 확인}"
echo ""
echo "남은 작업 (수동):"
echo "  1. Cloudflare Pages 환경변수 VITE_GAS_API_URL 업데이트"
echo "  2. GAS Script Properties에 SLACK_WEBHOOK_URL 재입력"
echo "     (어드민 → ⚙ 설정 탭에서 가능)"
echo "  3. 어드민 → ⚙ 설정 탭 → 트리거 설정하기 클릭"
echo "  4. npm run build && git push (CF Pages 재배포)"
echo "============================================================"
