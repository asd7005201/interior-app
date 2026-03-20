"""Drive 이미지 전체 삭제 + CrawlerInbox image_file_id 초기화.

Drive Materials 폴더 내 모든 파일/하위폴더 삭제 후,
CrawlerInbox의 image_file_id, image_file_name 컬럼을 비움.
"""
import time
from .drive_manager import clean_folder
from .sheets_client import (
    _get_spreadsheet, _rate_limit, get_api_call_count,
)


def main():
    print("=" * 60)
    print("Drive 이미지 전체 초기화")
    print("=" * 60)

    # 1) Drive 폴더 정리
    print("\n1) Drive Materials 폴더 정리 중...")
    stats = clean_folder()
    print(f"  삭제: 파일 {stats['deleted_files']}개, 폴더 {stats['deleted_folders']}개")
    if stats.get("skipped"):
        print(f"  건너뜀: {stats['skipped']}개")

    # 2) CrawlerInbox image_file_id / image_file_name 초기화
    print("\n2) CrawlerInbox image_file_id 초기화 중...")
    ss = _get_spreadsheet()
    _rate_limit()
    ws = ss.worksheet("CrawlerInbox")
    _rate_limit()
    headers = ws.row_values(1)

    fid_idx = headers.index("image_file_id")
    fname_idx = headers.index("image_file_name")

    def _col_letter(idx):
        result = ""
        while True:
            result = chr(ord('A') + idx % 26) + result
            idx = idx // 26 - 1
            if idx < 0:
                break
        return result

    fid_col = _col_letter(fid_idx)
    fname_col = _col_letter(fname_idx)

    # 전체 행 수 확인
    _rate_limit()
    all_values = ws.get_all_values()
    total_rows = len(all_values) - 1  # 헤더 제외

    if total_rows <= 0:
        print("  데이터 없음")
        return

    print(f"  {total_rows}행의 image_file_id/name 초기화...")

    # 범위 일괄 클리어
    _rate_limit()
    ws.batch_clear([
        f"{fid_col}2:{fid_col}{total_rows + 1}",
        f"{fname_col}2:{fname_col}{total_rows + 1}",
    ])

    print(f"  완료!")
    print(f"\nGoogle API 호출: {get_api_call_count()}회")
    print("=" * 60)


if __name__ == "__main__":
    main()
