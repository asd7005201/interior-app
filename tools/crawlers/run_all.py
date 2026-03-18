"""전체 자재 크롤링 실행 스크립트.

사용법:
    python -m tools.crawlers.run_all                         # 벽지 제외, 나머지 전체
    python -m tools.crawlers.run_all --include-wallpaper     # 벽지 포함 전체
    python -m tools.crawlers.run_all --only 장판-KCC 페인트   # 특정 카테고리만
    python -m tools.crawlers.run_all --limit 3               # 카테고리당 3개씩만 (테스트)
"""
import argparse
import uuid
from datetime import datetime, timezone

from .category_crawler import WallplanCategoryCrawler, MATERIAL_CATEGORIES
from .wallpaper_crawler import (
    WallplanCrawler, WallplanSearchCrawler,
    BRAND_CATEGORIES, PREMIUM_COLLECTIONS,
)


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

    # 1) 벽지 (옵션)
    if args.include_wallpaper and not args.only:
        print(f"\n{'=' * 60}")
        print(f"[벽지 크롤링]")
        print(f"{'=' * 60}")

        for brand in BRAND_CATEGORIES:
            print(f"\n{'─' * 40}")
            crawler = WallplanCrawler(brand)
            crawler.crawl(limit=args.limit)
            crawler.save(run_id)
            total_success += len(crawler.results)
            total_errors += len(crawler.errors)

        for coll_name in PREMIUM_COLLECTIONS:
            print(f"\n{'─' * 40}")
            crawler = WallplanSearchCrawler(coll_name)
            crawler.crawl(limit=args.limit)
            crawler.save(run_id)
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
        crawler.save(run_id)

        total_success += len(crawler.results)
        total_errors += len(crawler.errors)

    # 결과 요약
    print(f"\n{'=' * 60}")
    print(f"전체 크롤링 완료!")
    print(f"  성공: {total_success}개")
    print(f"  실패: {total_errors}개")
    print(f"  Run ID: {run_id}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
