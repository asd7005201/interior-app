"""Google Sheets 클라이언트 — CrawlerInbox / CrawlerLogs 읽기쓰기

안전 규칙 (CLAUDE.md 준수):
- 싱글톤 클라이언트 (매번 새로 인증하지 않음)
- 헤더 캐싱 (매번 row_values(1) 호출 안 함)
- API 호출 간 최소 3초 대기
- 세션당 호출 카운터 (50회 경고)
- 배치 append 지원
"""
import time
import gspread
from google.oauth2.service_account import Credentials
from . import config

# ── 싱글톤 ──
_client: gspread.Client | None = None
_spreadsheet: gspread.Spreadsheet | None = None
_header_cache: dict[str, list[str]] = {}  # worksheet_name -> headers
_api_call_count = 0
_last_api_call_time = 0.0

API_CALL_LIMIT = 50  # 세션당 경고 상한
API_DELAY = 3.0      # Google API 호출 간 최소 대기(초)


def _rate_limit():
    """API 호출 간 최소 3초 대기 + 카운터."""
    global _api_call_count, _last_api_call_time
    now = time.time()
    elapsed = now - _last_api_call_time
    if elapsed < API_DELAY:
        time.sleep(API_DELAY - elapsed)
    _last_api_call_time = time.time()
    _api_call_count += 1
    if _api_call_count == API_CALL_LIMIT:
        print(f"  ⚠️  Google API 호출 {API_CALL_LIMIT}회 도달 — 주의")
    if _api_call_count % 10 == 0:
        print(f"  📊 Google API 호출 누적: {_api_call_count}회")


def get_api_call_count() -> int:
    return _api_call_count


def _get_client() -> gspread.Client:
    global _client
    if _client is None:
        creds = Credentials.from_service_account_file(
            config.SERVICE_ACCOUNT_PATH, scopes=config.SCOPES
        )
        _client = gspread.authorize(creds)
    return _client


def _get_spreadsheet() -> gspread.Spreadsheet:
    global _spreadsheet
    if _spreadsheet is None:
        _rate_limit()
        _spreadsheet = _get_client().open_by_key(config.SPREADSHEET_ID)
    return _spreadsheet


def _get_headers(worksheet_name: str) -> list[str]:
    """헤더를 캐싱해서 반복 호출 방지."""
    if worksheet_name not in _header_cache:
        ss = _get_spreadsheet()
        _rate_limit()
        ws = ss.worksheet(worksheet_name)
        _rate_limit()
        headers = ws.row_values(1)
        _header_cache[worksheet_name] = headers
    return _header_cache[worksheet_name]


def _get_worksheet(name: str) -> gspread.Worksheet:
    ss = _get_spreadsheet()
    _rate_limit()
    return ss.worksheet(name)


def append_to_inbox(rows: list[dict]) -> int:
    """CrawlerInbox 시트에 행 추가. rows는 딕셔너리 리스트."""
    if not rows:
        return 0
    headers = _get_headers("CrawlerInbox")
    ws = _get_worksheet("CrawlerInbox")
    values = []
    for row in rows:
        values.append([str(row.get(h, "")) for h in headers])
    _rate_limit()
    ws.append_rows(values, value_input_option="USER_ENTERED")
    return len(values)


def append_to_logs(row: dict):
    """CrawlerLogs 시트에 1행 추가."""
    headers = _get_headers("CrawlerLogs")
    ws = _get_worksheet("CrawlerLogs")
    values = [str(row.get(h, "")) for h in headers]
    _rate_limit()
    ws.append_row(values, value_input_option="USER_ENTERED")


def batch_append_to_inbox(all_rows: list[dict]) -> int:
    """대량 배치 저장 — 모든 결과를 한번에 append (API 호출 최소화).

    기존 append_to_inbox는 카테고리별로 호출되지만,
    이 함수는 전체 결과를 모아서 1회 호출.
    gspread append_rows는 한번에 최대 ~40,000셀 권장이므로
    1,000행씩 청크로 나눠서 전송.
    """
    if not all_rows:
        return 0

    headers = _get_headers("CrawlerInbox")
    ws = _get_worksheet("CrawlerInbox")

    values = []
    for row in all_rows:
        values.append([str(row.get(h, "")) for h in headers])

    CHUNK_SIZE = 500  # 안전한 청크 크기
    total = 0
    for i in range(0, len(values), CHUNK_SIZE):
        chunk = values[i:i + CHUNK_SIZE]
        _rate_limit()
        ws.append_rows(chunk, value_input_option="USER_ENTERED")
        total += len(chunk)
        print(f"  → CrawlerInbox 배치 저장: {total}/{len(values)}행")

    return total


def batch_append_to_logs(log_rows: list[dict]) -> int:
    """대량 로그 배치 저장."""
    if not log_rows:
        return 0

    headers = _get_headers("CrawlerLogs")
    ws = _get_worksheet("CrawlerLogs")

    values = []
    for row in log_rows:
        values.append([str(row.get(h, "")) for h in headers])

    _rate_limit()
    ws.append_rows(values, value_input_option="USER_ENTERED")
    return len(values)


def read_inbox_all() -> list[dict]:
    """CrawlerInbox 전체 읽기 (중복 제거 등에 사용)."""
    headers = _get_headers("CrawlerInbox")
    ws = _get_worksheet("CrawlerInbox")
    _rate_limit()
    all_values = ws.get_all_values()
    rows = []
    for row_values in all_values[1:]:  # 헤더 제외
        row_dict = {}
        for i, h in enumerate(headers):
            row_dict[h] = row_values[i] if i < len(row_values) else ""
        rows.append(row_dict)
    return rows


def clear_inbox():
    """CrawlerInbox 데이터 영역만 삭제 (헤더 유지)."""
    ws = _get_worksheet("CrawlerInbox")
    _rate_limit()
    row_count = ws.row_count
    if row_count > 1:
        _rate_limit()
        ws.delete_rows(2, row_count)
    # 헤더 캐시 유지
