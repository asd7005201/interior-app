"""CrawlerInbox 중복 제거 + Materials 시트 sync 스크립트.

사용법:
    python -m tools.crawlers.dedup_and_sync                # 중복 제거 + sync
    python -m tools.crawlers.dedup_and_sync --dedup-only    # 중복 제거만
    python -m tools.crawlers.dedup_and_sync --sync-only     # sync만
    python -m tools.crawlers.dedup_and_sync --dry-run       # 실행 안 하고 분석만

중복 판별 기준: source_url (같은 제품 URL이면 중복)
sync 대상: sync_status="pending"인 행 → Materials 시트에 upsert
"""
import argparse
import uuid
from datetime import datetime, timezone

from .sheets_client import (
    _get_spreadsheet, _get_headers, _rate_limit,
    batch_append_to_inbox, get_api_call_count,
)

# Materials 시트의 헤더 (Code.js의 MATERIAL_MASTER_HEADERS_ 기준)
MATERIALS_HEADERS = [
    "material_id", "name", "brand", "spec", "unit", "unit_price",
    "image_file_id", "image_file_name", "note", "created_at", "updated_at",
    "search_text", "is_active", "is_representative", "material_type", "trade_code",
    "space_type", "sort_order", "expose_to_prequote", "recommendation_score_base",
    "price_band", "tags_summary", "recommendation_note",
]


def dedup_inbox(dry_run: bool = False) -> dict:
    """CrawlerInbox에서 중복 제거.

    Returns: {"total": n, "unique": n, "duplicates": n}
    """
    print("\n📊 CrawlerInbox 중복 분석 중...")

    ss = _get_spreadsheet()
    _rate_limit()
    ws = ss.worksheet("CrawlerInbox")
    _rate_limit()
    all_values = ws.get_all_values()

    if len(all_values) <= 1:
        print("  CrawlerInbox가 비어있습니다.")
        return {"total": 0, "unique": 0, "duplicates": 0}

    headers = all_values[0]
    rows = all_values[1:]

    # source_url 컬럼 인덱스 찾기
    try:
        url_idx = headers.index("source_url")
    except ValueError:
        print("  ⚠️ source_url 컬럼을 찾을 수 없습니다.")
        return {"total": len(rows), "unique": len(rows), "duplicates": 0}

    # 중복 제거 (source_url 기준, 마지막 것 유지)
    seen = {}
    for i, row in enumerate(rows):
        url = row[url_idx] if url_idx < len(row) else ""
        if url:
            seen[url] = i  # 같은 URL이면 마지막 인덱스로 덮어씀

    unique_indices = set(seen.values())
    # URL이 빈 행도 포함
    for i, row in enumerate(rows):
        url = row[url_idx] if url_idx < len(row) else ""
        if not url:
            unique_indices.add(i)

    unique_rows = [rows[i] for i in sorted(unique_indices)]
    duplicates = len(rows) - len(unique_rows)

    # 카테고리별 통계
    trade_idx = headers.index("trade_code") if "trade_code" in headers else None
    if trade_idx is not None:
        trade_counts = {}
        for row in unique_rows:
            tc = row[trade_idx] if trade_idx < len(row) else "unknown"
            trade_counts[tc] = trade_counts.get(tc, 0) + 1
        print(f"\n  📦 카테고리별 제품 수:")
        for tc, count in sorted(trade_counts.items(), key=lambda x: -x[1]):
            print(f"    {tc}: {count}개")

    print(f"\n  전체: {len(rows)}행")
    print(f"  고유: {len(unique_rows)}행")
    print(f"  중복: {duplicates}행")

    if dry_run:
        print("  (dry-run: 실제 삭제 안 함)")
        return {"total": len(rows), "unique": len(unique_rows), "duplicates": duplicates}

    if duplicates > 0:
        print(f"\n  🗑️ 중복 {duplicates}행 제거 중...")
        # 시트 데이터 영역 삭제 후 고유 행만 다시 쓰기
        _rate_limit()
        ws.batch_clear(["A2:ZZ"])  # 데이터 영역만 클리어
        if unique_rows:
            _rate_limit()
            # 500행씩 청크
            for i in range(0, len(unique_rows), 500):
                chunk = unique_rows[i:i + 500]
                _rate_limit()
                ws.append_rows(chunk, value_input_option="USER_ENTERED")
                print(f"    복원: {min(i + 500, len(unique_rows))}/{len(unique_rows)}행")
        print(f"  ✅ 중복 제거 완료: {len(unique_rows)}행 유지")
    else:
        print("  ✅ 중복 없음")

    return {"total": len(rows), "unique": len(unique_rows), "duplicates": duplicates}


def sync_to_materials(dry_run: bool = False) -> dict:
    """CrawlerInbox에서 sync_status=pending인 행을 Materials 시트에 sync.

    Returns: {"synced": n, "skipped": n, "errors": n}
    """
    print("\n📤 CrawlerInbox → Materials sync 시작...")

    ss = _get_spreadsheet()

    # CrawlerInbox 읽기
    _rate_limit()
    inbox_ws = ws_inbox = ss.worksheet("CrawlerInbox")
    _rate_limit()
    inbox_all = inbox_ws.get_all_values()

    if len(inbox_all) <= 1:
        print("  CrawlerInbox가 비어있습니다.")
        return {"synced": 0, "skipped": 0, "errors": 0}

    inbox_headers = inbox_all[0]
    inbox_rows = inbox_all[1:]

    # sync_status 컬럼 찾기
    try:
        sync_idx = inbox_headers.index("sync_status")
    except ValueError:
        print("  ⚠️ sync_status 컬럼을 찾을 수 없습니다.")
        return {"synced": 0, "skipped": 0, "errors": len(inbox_rows)}

    # pending인 행만 필터
    pending_rows = []
    pending_row_numbers = []  # 1-based row number (헤더 제외)
    for i, row in enumerate(inbox_rows):
        status = row[sync_idx] if sync_idx < len(row) else ""
        if status == "pending":
            pending_rows.append(row)
            pending_row_numbers.append(i + 2)  # +2: 헤더(1) + 0-index offset

    print(f"  전체 Inbox: {len(inbox_rows)}행")
    print(f"  pending (sync 대상): {len(pending_rows)}행")

    if not pending_rows:
        print("  ✅ sync할 항목이 없습니다.")
        return {"synced": 0, "skipped": len(inbox_rows), "errors": 0}

    if dry_run:
        print(f"  (dry-run: 실제 sync 안 함)")
        return {"synced": 0, "skipped": len(inbox_rows), "errors": 0}

    # Materials 시트 읽기
    _rate_limit()
    mat_ws = ss.worksheet("Materials")
    _rate_limit()
    mat_all = mat_ws.get_all_values()

    mat_headers = mat_all[0] if mat_all else MATERIALS_HEADERS
    mat_existing = {}
    if len(mat_all) > 1:
        name_idx = mat_headers.index("name") if "name" in mat_headers else 0
        for row in mat_all[1:]:
            name = row[name_idx] if name_idx < len(row) else ""
            if name:
                mat_existing[name] = True

    # inbox → materials 매핑
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    new_materials = []
    synced_count = 0
    skipped_count = 0

    for row in pending_rows:
        row_dict = {}
        for j, h in enumerate(inbox_headers):
            row_dict[h] = row[j] if j < len(row) else ""

        name = row_dict.get("name", "")
        if not name:
            skipped_count += 1
            continue

        # 이미 Materials에 있으면 skip (name 기준)
        if name in mat_existing:
            skipped_count += 1
            continue

        # material_id 생성
        material_id = f"MAT_{uuid.uuid4().hex[:12].upper()}"

        # search_text 생성 (검색용 텍스트)
        search_parts = [
            name, row_dict.get("brand", ""),
            row_dict.get("material_type", ""),
            row_dict.get("spec", ""),
            row_dict.get("trade_code", ""),
            row_dict.get("tags_summary", ""),
        ]
        search_text = " ".join(p for p in search_parts if p)

        mat_row = {
            "material_id": material_id,
            "name": name,
            "brand": row_dict.get("brand", ""),
            "spec": row_dict.get("spec", ""),
            "unit": row_dict.get("unit", "개"),
            "unit_price": row_dict.get("unit_price", ""),
            "image_file_id": row_dict.get("image_file_id", ""),
            "image_file_name": row_dict.get("image_file_name", ""),
            "note": "",
            "created_at": now,
            "updated_at": now,
            "search_text": search_text,
            "is_active": "TRUE",
            "is_representative": row_dict.get("is_representative", ""),
            "material_type": row_dict.get("material_type", ""),
            "trade_code": row_dict.get("trade_code", ""),
            "space_type": row_dict.get("space_type", ""),
            "sort_order": "",
            "expose_to_prequote": row_dict.get("expose_to_prequote", ""),
            "recommendation_score_base": row_dict.get("recommendation_score_base", ""),
            "price_band": row_dict.get("price_band", ""),
            "tags_summary": row_dict.get("tags_summary", ""),
            "recommendation_note": row_dict.get("recommendation_note", ""),
        }
        new_materials.append(mat_row)
        mat_existing[name] = True  # 중복 방지
        synced_count += 1

    # Materials에 배치 추가
    if new_materials:
        values = []
        for m in new_materials:
            values.append([str(m.get(h, "")) for h in mat_headers])

        print(f"\n  📝 Materials에 {len(values)}행 추가 중...")
        for i in range(0, len(values), 500):
            chunk = values[i:i + 500]
            _rate_limit()
            mat_ws.append_rows(chunk, value_input_option="USER_ENTERED")
            print(f"    저장: {min(i + 500, len(values))}/{len(values)}행")

    # CrawlerInbox의 sync_status 업데이트 (pending → synced)
    if pending_row_numbers:
        print(f"\n  📝 Inbox sync_status 업데이트 중...")

        def _col_letter(idx: int) -> str:
            """0-based 인덱스 → 엑셀 컬럼 문자 (A, B, ..., Z, AA, AB, ...)"""
            result = ""
            while True:
                result = chr(ord('A') + idx % 26) + result
                idx = idx // 26 - 1
                if idx < 0:
                    break
            return result

        sync_col = _col_letter(sync_idx)
        try:
            msg_idx = inbox_headers.index("sync_message")
            msg_col = _col_letter(msg_idx)
        except ValueError:
            msg_idx = None

        batch_updates = []
        for row_num in pending_row_numbers:
            batch_updates.append({
                "range": f"{sync_col}{row_num}",
                "values": [["synced"]],
            })
            if msg_idx is not None:
                batch_updates.append({
                    "range": f"{msg_col}{row_num}",
                    "values": [[f"synced at {now}"]],
                })

        # 100개씩 배치
        for i in range(0, len(batch_updates), 100):
            chunk = batch_updates[i:i + 100]
            _rate_limit()
            inbox_ws.batch_update(chunk, value_input_option="USER_ENTERED")
            print(f"    상태 업데이트: {min(i + 100, len(batch_updates))}/{len(batch_updates)}")

    print(f"\n  ✅ Sync 완료:")
    print(f"    신규 추가: {synced_count}개")
    print(f"    건너뜀 (중복/빈값): {skipped_count}개")
    print(f"    Google API 호출 누적: {get_api_call_count()}회")

    return {"synced": synced_count, "skipped": skipped_count, "errors": 0}


def main():
    parser = argparse.ArgumentParser(description="CrawlerInbox 중복 제거 + Materials sync")
    parser.add_argument("--dedup-only", action="store_true", help="중복 제거만")
    parser.add_argument("--sync-only", action="store_true", help="sync만")
    parser.add_argument("--dry-run", action="store_true", help="실행 안 하고 분석만")
    args = parser.parse_args()

    print(f"{'=' * 60}")
    print(f"CrawlerInbox 중복 제거 + Materials Sync")
    print(f"{'=' * 60}")

    if not args.sync_only:
        dedup_inbox(dry_run=args.dry_run)

    if not args.dedup_only:
        sync_to_materials(dry_run=args.dry_run)

    print(f"\n{'=' * 60}")
    print(f"전체 작업 완료! (Google API 호출: {get_api_call_count()}회)")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
