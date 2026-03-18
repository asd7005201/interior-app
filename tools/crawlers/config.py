"""크롤러 설정"""
import os

# 프로젝트 루트 기준 경로
_PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

# Google 서비스 계정
SERVICE_ACCOUNT_PATH = os.path.join(_PROJECT_ROOT, "tools", ".auth", "service_account.json")

# Google Sheets - quote_DB
SPREADSHEET_ID = "1SOG2d8w_3s1zP0Absc6x97WFzqrBkJpzZbvtGFV22ww"

# Google Drive - Material 루트 폴더
DRIVE_ROOT_FOLDER_ID = "1nn2bY8nqwccQuqZCiC3Z61iT2M-V2di3"

# API 스코프
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# 크롤링 설정
REQUEST_TIMEOUT = 15
REQUEST_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
}
