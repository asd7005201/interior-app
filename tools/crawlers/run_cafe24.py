"""Cafe24 사이트 크롤링 실행 — bathnmore/maruall/daelimwood.

배치 저장 모드 (Google API 최소화).

사용법:
    python -m tools.crawlers.run_cafe24                          # 전체
    python -m tools.crawlers.run_cafe24 --site maruall            # maruall만
    python -m tools.crawlers.run_cafe24 --site bathnmore          # bathnmore만
    python -m tools.crawlers.run_cafe24 --limit 5                 # 카테고리당 5개
    python -m tools.crawlers.run_cafe24 --dry-run                 # 테스트
"""
import argparse
import uuid
from datetime import datetime, timezone

from .cafe24_crawler import Cafe24Crawler, CAFE24_SITES
from .sheets_client import batch_append_to_inbox, batch_append_to_logs, get_api_call_count


def main():
    parser = argparse.ArgumentParser(description="Cafe24 사이트 크롤링")
    parser.add_argument("--site", type=str, default="", help="특정 사이트만 (maruall/bathnmore/daelimwood)")
    parser.add_argument("--limit", type=int, default=0, help="카테고리당 최대 수 (0=전부)")
    parser.add_argument("--only", nargs="*", default=None, help="특정 site_key만")
    parser.add_argument("--dry-run", action="store_true", help="크롤링만 하고 저장 안 함")
    args = parser.parse_args()

    run_id = f"RUN_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6].upper()}"

    # 대상 필터
    if args.only:
        site_keys = [k for k in args.only if k in CAFE24_SITES]
    elif args.site:
        site_keys = [k for k in CAFE24_SITES if k.startswith(args.site)]
    else:
        site_keys = list(CAFE24_SITES.keys())

    print(f"{'=' * 60}")
    print(f"Cafe24 사이트 크롤링")
    print(f"Run ID: {run_id}")
    print(f"대상: {len(site_keys)}개 카테고리")
    if args.limit:
        print(f"제한: 카테고리당 {args.limit}개")
    print(f"{'=' * 60}")

    all_inbox_rows = []
    all_log_rows = []
    total_success = 0
    total_errors = 0

    for key in site_keys:
        info = CAFE24_SITES[key]
        print(f"\n{'─' * 40}")
        print(f"[{key}] {info['source_site']} — {info['material_type']}")

        crawler = Cafe24Crawler(key)
        crawler.crawl(limit=args.limit)
        inbox_rows, log_row = crawler.collect(run_id)
        all_inbox_rows.extend(inbox_rows)
        all_log_rows.append(log_row)
        total_success += len(crawler.results)
        total_errors += len(crawler.errors)

        print(f"  결과: 성공 {len(crawler.results)}, 실패 {len(crawler.errors)}")

    if args.dry_run:
        print(f"\n(dry-run: 저장 안 함)")
        print(f"총 {total_success}개 수집됨")
        return

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
