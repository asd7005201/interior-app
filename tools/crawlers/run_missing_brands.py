"""누락 브랜드 크롤링 — FT/DID/코스모스/서울벽지.

시트 정지로 유실된 데이터 복구용 스크립트.
배치 저장 모드로 Google API 최소 사용.
"""
import uuid
from datetime import datetime, timezone

from .wallpaper_crawler import WallplanCrawler, BRAND_SUBCATEGORIES
from .sheets_client import batch_append_to_inbox, batch_append_to_logs, get_api_call_count

MISSING_BRANDS = ("FT벽지", "DID벽지", "코스모스벽지", "서울벽지")


def main():
    run_id = f"RUN_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6].upper()}"

    print(f"{'=' * 60}")
    print(f"누락 브랜드 크롤링: {', '.join(MISSING_BRANDS)}")
    print(f"Run ID: {run_id}")
    print(f"{'=' * 60}")

    all_inbox_rows = []
    all_log_rows = []
    total_success = 0
    total_errors = 0

    for coll_name, info in BRAND_SUBCATEGORIES.items():
        if info["brand"] not in MISSING_BRANDS:
            continue

        print(f"\n{'─' * 40}")
        print(f"[{coll_name}] ({info['brand']})")

        crawler = WallplanCrawler(
            brand=info["brand"],
            cate_cd=info["cate_cd"],
            collection_name=coll_name,
            skip_old_filter=True,
        )
        crawler.crawl(limit=0)
        inbox_rows, log_row = crawler.collect(run_id)
        all_inbox_rows.extend(inbox_rows)
        all_log_rows.append(log_row)
        total_success += len(crawler.results)
        total_errors += len(crawler.errors)

    # 배치 저장
    print(f"\n{'=' * 60}")
    print(f"📤 배치 저장: {len(all_inbox_rows)}행")

    if all_inbox_rows:
        saved = batch_append_to_inbox(all_inbox_rows)
        print(f"  ✅ CrawlerInbox: {saved}행")

    if all_log_rows:
        saved = batch_append_to_logs(all_log_rows)
        print(f"  ✅ CrawlerLogs: {saved}행")

    print(f"\n{'=' * 60}")
    print(f"완료! 성공: {total_success}, 실패: {total_errors}")
    print(f"Google API: {get_api_call_count()}회")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
