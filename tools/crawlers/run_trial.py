"""시범 크롤링 실행 스크립트.

사용법:
    python -m tools.crawlers.run_trial --limit 5
    python -m tools.crawlers.run_trial --clean          # 시트+Drive 전부 정리 후 크롤링
    python -m tools.crawlers.run_trial --clean-only      # 정리만 (크롤링 안 함)
"""
import argparse
import uuid
from datetime import datetime, timezone

from .wallpaper_crawler import WallplanCrawler, BRAND_CATEGORIES


def _clean_all():
    """CrawlerInbox, CrawlerLogs 시트 + Drive Material 폴더 전부 정리."""
    import gspread
    from google.oauth2.service_account import Credentials
    from . import config
    from .drive_manager import clean_folder

    print("기존 데이터 정리 중...")

    # 시트 정리
    creds = Credentials.from_service_account_file(config.SERVICE_ACCOUNT_PATH, scopes=config.SCOPES)
    ss = gspread.authorize(creds).open_by_key(config.SPREADSHEET_ID)
    for name in ["CrawlerInbox", "CrawlerLogs"]:
        ws = ss.worksheet(name)
        if ws.row_count > 1:
            ws.delete_rows(2, ws.row_count)
            print(f"  {name}: 정리 완료")

    # Drive 정리
    stats = clean_folder()
    print(f"  Drive: 파일 {stats['deleted_files']}개, 폴더 {stats['deleted_folders']}개 삭제")
    print()


def main():
    parser = argparse.ArgumentParser(description="벽지 자재 시범 크롤링")
    parser.add_argument("--limit", type=int, default=5, help="브랜드당 크롤링 제품 수")
    parser.add_argument("--brands", nargs="*", default=None, help="크롤링할 브랜드 (기본: 전부)")
    parser.add_argument("--clean", action="store_true", help="기존 데이터 정리 후 크롤링")
    parser.add_argument("--clean-only", action="store_true", help="정리만 (크롤링 안 함)")
    args = parser.parse_args()

    if args.clean or args.clean_only:
        _clean_all()
        if args.clean_only:
            print("정리 완료. 크롤링은 건너뜁니다.")
            return

    brands = args.brands or list(BRAND_CATEGORIES.keys())
    run_id = f"RUN_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6].upper()}"

    print(f"=" * 60)
    print(f"벽지 자재 시범 크롤링")
    print(f"Run ID: {run_id}")
    print(f"브랜드: {', '.join(brands)}")
    print(f"브랜드당 제품 수: {args.limit}")
    print(f"=" * 60)

    total_success = 0
    total_errors = 0

    for brand in brands:
        if brand not in BRAND_CATEGORIES:
            print(f"\n[SKIP] 알 수 없는 브랜드: {brand}")
            continue

        print(f"\n{'─' * 40}")
        crawler = WallplanCrawler(brand)
        crawler.crawl(limit=args.limit)
        crawler.save(run_id)

        total_success += len(crawler.results)
        total_errors += len(crawler.errors)

    # 결과 요약
    print(f"\n{'=' * 60}")
    print(f"크롤링 완료!")
    print(f"  성공: {total_success}개")
    print(f"  실패: {total_errors}개")
    print(f"  Run ID: {run_id}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
