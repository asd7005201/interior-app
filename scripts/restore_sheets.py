#!/usr/bin/env python3
"""
restore_sheets.py — data/prequote_static.json을 Google 스프레드시트에 복원
사용법: python3 scripts/restore_sheets.py <SPREADSHEET_ID>

의존성: pip install gspread oauth2client
인증: ~/.config/gspread/credentials.json 필요 (Google Service Account)
"""
import sys, json, os, time

try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
except ImportError:
    print("pip install gspread oauth2client 설치 후 실행하세요.")
    sys.exit(1)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE  = os.path.join(SCRIPT_DIR, "..", "data", "prequote_static.json")
CRED_FILE  = os.path.expanduser("~/.config/gspread/credentials.json")

def main():
    if len(sys.argv) < 2:
        print("사용법: python3 restore_sheets.py <SPREADSHEET_ID>")
        sys.exit(1)

    ss_id = sys.argv[1].strip()
    print(f"Spreadsheet ID: {ss_id}")

    if not os.path.exists(DATA_FILE):
        print(f"백업 파일 없음: {DATA_FILE}")
        sys.exit(1)

    with open(DATA_FILE, "r", encoding="utf-8") as f:
        backup = json.load(f)

    print(f"백업 날짜: {backup.get('exported_at', '알 수 없음')}")

    # Google Sheets 인증
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    if os.path.exists(CRED_FILE):
        creds = ServiceAccountCredentials.from_json_keyfile_name(CRED_FILE, scope)
        client = gspread.authorize(creds)
    else:
        print("Service Account 인증 파일이 없습니다.")
        print(f"위치: {CRED_FILE}")
        print("Google Cloud Console에서 Service Account JSON을 받아 해당 경로에 저장하세요.")
        sys.exit(1)

    ss = client.open_by_key(ss_id)

    # Settings는 민감 값(SLACK 등) 제외하고 복원
    SKIP_KEYS = {"SLACK_WEBHOOK_URL"}

    for sheet_name, rows in backup["sheets"].items():
        if rows is None:
            print(f"  ⏭ {sheet_name}: 없음 — 스킵")
            continue
        if not rows:
            print(f"  ⏭ {sheet_name}: 빈 시트 — 스킵")
            continue

        print(f"  📋 {sheet_name} 복원 중 ({len(rows)}행)...")

        try:
            worksheet = ss.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            worksheet = ss.add_worksheet(title=sheet_name, rows=len(rows)+10, cols=30)

        # 기존 내용 지우기
        worksheet.clear()
        time.sleep(1)

        # Settings 시트는 민감 키 제외
        if sheet_name == "Settings":
            rows = [r for r in rows if str(r.get("key", "")).strip() not in SKIP_KEYS]

        if not rows:
            continue

        headers = list(rows[0].keys())
        data = [headers]
        for row in rows:
            data.append([str(row.get(h, "") or "") for h in headers])

        worksheet.update("A1", data)
        time.sleep(2)
        print(f"    ✅ {len(rows)}행 완료")

    print("\n✅ 복원 완료!")
    print("다음으로 GAS 에디터에서 initializeAppManual('<SPREADSHEET_ID>') 를 실행하세요.")

if __name__ == "__main__":
    main()
