"""Google Sheets 클라이언트 — CrawlerInbox / CrawlerLogs 읽기쓰기"""
import gspread
from google.oauth2.service_account import Credentials
from . import config


def _get_client():
    creds = Credentials.from_service_account_file(config.SERVICE_ACCOUNT_PATH, scopes=config.SCOPES)
    return gspread.authorize(creds)


def _get_spreadsheet():
    return _get_client().open_by_key(config.SPREADSHEET_ID)


def append_to_inbox(rows: list[dict]):
    """CrawlerInbox 시트에 행 추가. rows는 딕셔너리 리스트."""
    ss = _get_spreadsheet()
    ws = ss.worksheet("CrawlerInbox")
    headers = ws.row_values(1)
    values = []
    for row in rows:
        values.append([str(row.get(h, "")) for h in headers])
    if values:
        ws.append_rows(values, value_input_option="USER_ENTERED")
    return len(values)


def append_to_logs(row: dict):
    """CrawlerLogs 시트에 1행 추가."""
    ss = _get_spreadsheet()
    ws = ss.worksheet("CrawlerLogs")
    headers = ws.row_values(1)
    values = [str(row.get(h, "")) for h in headers]
    ws.append_row(values, value_input_option="USER_ENTERED")
