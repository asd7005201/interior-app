"""전체 자재 크롤링 실행 스크립트 — 배치 저장 모드.

Google Sheets API 호출을 최소화하기 위해:
1. 모든 카테고리를 크롤링하면서 결과를 메모리에 모음
2. 마지막에 한번에 batch_append (API 호출 ~10회 이하)

사용법:
    python -m tools.crawlers.run_all                         # 벽지 제외, 나머지 전체
    python -m tools.crawlers.run_all --include-wallpaper     # 벽지 포함 전체
    python -m tools.crawlers.run_all --only 장판-KCC 페인트-벽면용  # 특정 카테고리만
    python -m tools.crawlers.run_all --limit 3               # 카테고리당 3개씩만 (테스트)
"""
import argparse
import uuid
from datetime import datetime, timezone

from .category_crawler import WallplanCategoryCrawler, MATERIAL_CATEGORIES
from .wallpaper_crawler import (
    WallplanCrawler,
    PREMIUM_CATEGORIES, BRAND_SUBCATEGORIES,
)
from .sheets_client import batch_append_to_inbox, batch_append_to_logs, get_api_call_count


def main():
    parser = argparse.ArgumentParser(description="전체 자재 크롤링")
    parser.add_argument("--limit", type=int, default=0, help="카테고리당 크롤링 수 (0=전부)")
    parser.add_argument("--only", nargs="*", default=None, help="특정 카테고리만 크롤링")
    parser.add_argument("--include-wallpaper", action="store_true", help="벽지도 포함")
    args = parser.parse_args()

    run_id = f"RUN_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6].upper()}"
    limit_desc = f"{args.limit}개씩" if args.limit else "전부"

    total_success = 0
    total_errors = 0

    # 배치 저장용 버퍼
    all_inbox_rows: list[dict] = []
    all_log_rows: list[dict] = []

    # 1) 벽지 (옵션)
    if args.include_wallpaper and not args.only:
        print(f"\n{'=' * 60}")
        print(f"[벽지 - 프리미엄/트렌드]")
        print(f"{'=' * 60}")

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

        print(f"\n{'=' * 60}")
        print(f"[벽지 - 브랜드 세부 컬렉션]")
        print(f"{'=' * 60}")

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

    # 2) 기타 자재 카테고리
    if args.only:
        cat_keys = args.only
    else:
        cat_keys = list(MATERIAL_CATEGORIES.keys())

    print(f"\n{'=' * 60}")
    print(f"자재 카테고리 크롤링")
    print(f"Run ID: {run_id}")
    print(f"카테고리: {len(cat_keys)}개")
    print(f"수량: {limit_desc}")
    print(f"{'=' * 60}")

    for key in cat_keys:
        if key not in MATERIAL_CATEGORIES:
            print(f"\n[SKIP] 알 수 없는 카테고리: {key}")
            continue

        info = MATERIAL_CATEGORIES[key]
        print(f"\n{'─' * 40}")
        print(f"[{key}] trade={info['trade_code']}")

        crawler = WallplanCategoryCrawler(key)
        crawler.crawl(limit=args.limit)
        inbox_rows, log_row = crawler.collect(run_id)
        all_inbox_rows.extend(inbox_rows)
        all_log_rows.append(log_row)

        total_success += len(crawler.results)
        total_errors += len(crawler.errors)

    # 3) 배치 저장 — Google Sheets에 한번에 전송
    print(f"\n{'=' * 60}")
    print(f"📤 Google Sheets 배치 저장 시작...")
    print(f"  Inbox: {len(all_inbox_rows)}행, Logs: {len(all_log_rows)}행")
    print(f"{'=' * 60}")

    if all_inbox_rows:
        saved = batch_append_to_inbox(all_inbox_rows)
        print(f"  ✅ CrawlerInbox 저장 완료: {saved}행")

    if all_log_rows:
        saved = batch_append_to_logs(all_log_rows)
        print(f"  ✅ CrawlerLogs 저장 완료: {saved}행")

    # 결과 요약
    print(f"\n{'=' * 60}")
    print(f"전체 크롤링 완료!")
    print(f"  성공: {total_success}개")
    print(f"  실패: {total_errors}개")
    print(f"  Google API 호출: {get_api_call_count()}회")
    print(f"  Run ID: {run_id}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
