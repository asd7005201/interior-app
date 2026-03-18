"""벽지 자재 크롤링 실행 스크립트.

사용법:
    python -m tools.crawlers.run_trial                   # 전체 크롤링 (2022년 이후)
    python -m tools.crawlers.run_trial --limit 5          # 브랜드당 5개씩만
    python -m tools.crawlers.run_trial --clean            # 기존 정리 후 크롤링
    python -m tools.crawlers.run_trial --clean-only       # 정리만
"""
import argparse
import uuid
from datetime import datetime, timezone

from .wallpaper_crawler import (
    WallplanCrawler, WallplanSearchCrawler,
    BRAND_CATEGORIES, PREMIUM_COLLECTIONS,
)


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
    if stats.get("skipped"):
        print(f"  (권한 없는 파일 {stats['skipped']}개 건너뜀)")
    print()


def main():
    parser = argparse.ArgumentParser(description="벽지 자재 크롤링")
    parser.add_argument("--limit", type=int, default=0, help="브랜드당 크롤링 수 (0=전부)")
    parser.add_argument("--brands", nargs="*", default=None, help="크롤링할 브랜드")
    parser.add_argument("--clean", action="store_true", help="기존 데이터 정리 후 크롤링")
    parser.add_argument("--clean-only", action="store_true", help="정리만")
    parser.add_argument("--skip-premium", action="store_true", help="프리미엄 컬렉션(파사드/디아망/로하스) 건너뛰기")
    args = parser.parse_args()

    if args.clean or args.clean_only:
        _clean_all()
        if args.clean_only:
            print("정리 완료.")
            return

    brands = args.brands or list(BRAND_CATEGORIES.keys())
    run_id = f"RUN_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6].upper()}"
    limit_desc = f"{args.limit}개씩" if args.limit else "전부"

    print(f"{'=' * 60}")
    print(f"벽지 자재 크롤링")
    print(f"Run ID: {run_id}")
    print(f"브랜드: {', '.join(brands)}")
    print(f"수량: {limit_desc} (2022년 이후 필터링)")
    if not args.skip_premium:
        print(f"프리미엄: {', '.join(PREMIUM_COLLECTIONS.keys())}")
    print(f"{'=' * 60}")

    total_success = 0
    total_errors = 0

    # 1) 브랜드별 롤 벽지 크롤링
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

    # 2) 프리미엄 컬렉션 크롤링 (파사드/디아망/로하스)
    if not args.skip_premium:
        # 이미 크롤링된 goodsNo 수집 (중복 방지)
        crawled_goods = set()
        # 간단히 이전 크롤러의 결과에서 goods_no 수집
        # (이미 시트에 저장된 것과는 별개로, 이번 실행 내 중복만 방지)

        for coll_name in PREMIUM_COLLECTIONS:
            print(f"\n{'─' * 40}")
            print(f"[프리미엄] {coll_name}")
            crawler = WallplanSearchCrawler(coll_name)
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
