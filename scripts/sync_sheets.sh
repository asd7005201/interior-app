#!/usr/bin/env bash
# ============================================================
# sync_sheets.sh — 가견적 앱 정적 시트 데이터를 GitHub에 백업
# 사용법: bash scripts/sync_sheets.sh
# ============================================================
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
DATA_DIR="$REPO_ROOT/data"
GAS_URL="https://script.google.com/macros/s/AKfycbxpLx1fZSUPX3bhXoJqjyUg5UDo008EA2o1or55Dg3wFu-t_LQuQCtBiXKh5AwVvIHy2A/exec"
ADMIN_PW="${PREQUOTE_ADMIN_PW:-test}"
OUTPUT_FILE="$DATA_DIR/prequote_static.json"

mkdir -p "$DATA_DIR"

echo "⏳ GAS에서 시트 데이터 가져오는 중..."

# GAS API 호출 (2단계 리다이렉트)
LOCATION=$(curl -si -X POST "$GAS_URL" \
  -H "Content-Type: text/plain" \
  -d "{\"action\":\"adminExportStaticData\",\"credential\":\"$ADMIN_PW\"}" \
  2>/dev/null | grep -i "^location:" | tr -d '\r' | sed 's/[Ll]ocation: //')

if [ -z "$LOCATION" ]; then
  echo "❌ GAS 리다이렉트 URL을 가져오지 못했습니다."
  exit 1
fi

RESPONSE=$(curl -sL "$LOCATION")

# ok 확인
OK=$(echo "$RESPONSE" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d.get('ok','false'))" 2>/dev/null || echo "false")
if [ "$OK" != "True" ] && [ "$OK" != "true" ]; then
  echo "❌ API 오류: $RESPONSE"
  exit 1
fi

# data만 추출해서 저장
echo "$RESPONSE" | python3 -c "
import sys, json
d = json.load(sys.stdin)
print(json.dumps(d['data'], ensure_ascii=False, indent=2))
" > "$OUTPUT_FILE"

echo "✅ $OUTPUT_FILE 저장 완료"

# Git 커밋
cd "$REPO_ROOT"
git add "$OUTPUT_FILE"
if git diff --cached --quiet; then
  echo "ℹ 변경 없음 — 커밋 스킵"
else
  git commit -m "data: 가견적 정적 시트 자동 동기화 $(date '+%Y-%m-%d %H:%M')"
  git push origin main
  echo "✅ GitHub 푸시 완료"
fi
