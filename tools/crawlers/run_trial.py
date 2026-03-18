"""시범 크롤링 실행 스크립트.

사용법:
    python -m tools.crawlers.run_trial --limit 5
"""
import argparse
import uuid
from datetime import datetime, timezone

from .wallpaper_crawler import WallplanCrawler, BRAND_CATEGORIES


def main():
    parser = argparse.ArgumentParser(description="벽지 자재 시범 크롤링")
    parser.add_argument("--limit", type=int, default=5, help="브랜드당 크롤링 제품 수")
    parser.add_argument("--brands", nargs="*", default=None, help="크롤링할 브랜드 (기본: 전부)")
    args = parser.parse_args()

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
