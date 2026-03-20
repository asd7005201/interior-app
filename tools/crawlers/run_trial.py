"""벽지 자재 크롤링 실행 스크립트 — 배치 저장 모드.

사용법:
    python -m tools.crawlers.run_trial                   # 전체 크롤링 (프리미엄 + 브랜드)
    python -m tools.crawlers.run_trial --limit 5          # 컬렉션당 5개씩만
    python -m tools.crawlers.run_trial --clean            # 기존 정리 후 크롤링
    python -m tools.crawlers.run_trial --clean-only       # 정리만
    python -m tools.crawlers.run_trial --premium-only     # 프리미엄/트렌드 벽지만
"""
import argparse
import uuid
from datetime import datetime, timezone

from .wallpaper_crawler import (
    WallplanCrawler,
    BRAND_CATEGORIES, PREMIUM_CATEGORIES, BRAND_SUBCATEGORIES,
)
from .sheets_client import batch_append_to_inbox, batch_append_to_logs, get_api_call_count


def _clean_all():
    """CrawlerInbox, CrawlerLogs 시트 + Drive Material 폴더 전부 정리."""
    from .sheets_client import clear_inbox, _get_worksheet, _rate_limit
    from .drive_manager import clean_folder

    print("기존 데이터 정리 중...")

    # 시트 정리 (싱글톤 사용)
    clear_inbox()
    print("  CrawlerInbox: 정리 완료")

    ws = _get_worksheet("CrawlerLogs")
    _rate_limit()
    if ws.row_count > 1:
        _rate_limit()
        ws.delete_rows(2, ws.row_count)
    print("  CrawlerLogs: 정리 완료")

    # Drive 정리
    stats = clean_folder()
    print(f"  Drive: 파일 {stats['deleted_files']}개, 폴더 {stats['deleted_folders']}개 삭제")
    if stats.get("skipped"):
        print(f"  (권한 없는 파일 {stats['skipped']}개 건너뜀)")
    print()


def main():
    parser = argparse.ArgumentParser(description="벽지 자재 크롤링")
    parser.add_argument("--limit", type=int, default=0, help="컬렉션당 크롤링 수 (0=전부)")
    parser.add_argument("--clean", action="store_true", help="기존 데이터 정리 후 크롤링")
    parser.add_argument("--clean-only", action="store_true", help="정리만")
    parser.add_argument("--premium-only", action="store_true", help="프리미엄/트렌드 벽지만 크롤링")
    parser.add_argument("--skip-premium", action="store_true", help="프리미엄 건너뛰기")
    parser.add_argument("--skip-subcategories", action="store_true", help="브랜드 세부 컬렉션 건너뛰기")
    args = parser.parse_args()

    if args.clean or args.clean_only:
        _clean_all()
        if args.clean_only:
            print("정리 완료.")
            return

    run_id = f"RUN_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6].upper()}"
    limit_desc = f"{args.limit}개씩" if args.limit else "전부"

    print(f"{'=' * 60}")
    print(f"벽지 자재 크롤링")
    print(f"Run ID: {run_id}")
    print(f"수량: {limit_desc}")
    print(f"{'=' * 60}")

    total_success = 0
    total_errors = 0

    # 배치 저장용 버퍼
    all_inbox_rows: list[dict] = []
    all_log_rows: list[dict] = []

    # 1) ★ 프리미엄/트렌드 벽지 (카테고리 기반, 연도 필터 없음)
    if not args.skip_premium:
        print(f"\n{'━' * 60}")
        print(f"[Phase 1] 프리미엄/트렌드 벽지: {', '.join(PREMIUM_CATEGORIES.keys())}")
        print(f"{'━' * 60}")

        for coll_name, info in PREMIUM_CATEGORIES.items():
            print(f"\n{'─' * 40}")
            print(f"[프리미엄] {coll_name} ({info['brand']})")
            crawler = WallplanCrawler(
                brand=info["brand"],
                cate_cd=info["cate_cd"],
                collection_name=coll_name,
                skip_old_filter=True,
            )
            crawler.crawl(limit=args.limit)
            inbox_rows, log_row = crawler.collect(run_id)
            all_inbox_rows.extend(inbox_rows)
            all_log_rows.append(log_row)
            total_success += len(crawler.results)
            total_errors += len(crawler.errors)

    if args.premium_only:
        _batch_save_and_summary(all_inbox_rows, all_log_rows, total_success, total_errors, run_id)
        return

    # 2) 브랜드별 세부 컬렉션 (로하스, 스케치, 베스티 등)
    if not args.skip_subcategories:
        print(f"\n{'━' * 60}")
        print(f"[Phase 2] 브랜드 세부 컬렉션: {', '.join(BRAND_SUBCATEGORIES.keys())}")
        print(f"{'━' * 60}")

        for coll_name, info in BRAND_SUBCATEGORIES.items():
            print(f"\n{'─' * 40}")
            print(f"[컬렉션] {coll_name} ({info['brand']})")
            crawler = WallplanCrawler(
                brand=info["brand"],
                cate_cd=info["cate_cd"],
                collection_name=coll_name,
                skip_old_filter=True,
            )
            crawler.crawl(limit=args.limit)
            inbox_rows, log_row = crawler.collect(run_id)
            all_inbox_rows.extend(inbox_rows)
            all_log_rows.append(log_row)
            total_success += len(crawler.results)
            total_errors += len(crawler.errors)

    _batch_save_and_summary(all_inbox_rows, all_log_rows, total_success, total_errors, run_id)


def _batch_save_and_summary(inbox_rows, log_rows, success, errors, run_id):
    """배치 저장 + 결과 요약."""
    print(f"\n{'=' * 60}")
    print(f"📤 Google Sheets 배치 저장...")
    print(f"  Inbox: {len(inbox_rows)}행, Logs: {len(log_rows)}행")

    if inbox_rows:
        saved = batch_append_to_inbox(inbox_rows)
        print(f"  ✅ CrawlerInbox: {saved}행 저장")

    if log_rows:
        saved = batch_append_to_logs(log_rows)
        print(f"  ✅ CrawlerLogs: {saved}행 저장")

    print(f"\n{'=' * 60}")
    print(f"크롤링 완료!")
    print(f"  성공: {success}개")
    print(f"  실패: {errors}개")
    print(f"  Google API 호출: {get_api_call_count()}회")
    print(f"  Run ID: {run_id}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
