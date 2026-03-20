"""대량 이미지 Drive 업로드 스크립트.

폴더 구조: 공종 > 브랜드 > 컬렉션 (선택)
예: Materials/벽지/LX지인/디아망/product_name.jpg

Google 계정 보호:
- Drive API 호출 간 최소 1초 대기
- 연속 3회 실패 시 중단
- Sheets 배치 업데이트 (50개씩)

사용법:
    python -m tools.crawlers.upload_images                    # 미업로드 전체
    python -m tools.crawlers.upload_images --limit 50         # 50개만
    python -m tools.crawlers.upload_images --trade wallpaper  # 벽지만
    python -m tools.crawlers.upload_images --dry-run          # 분석만
"""
import argparse
import io
import json
import re
import sys
import time

# Windows cp949 인코딩 문제 방지
if sys.stdout.encoding != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import requests
from googleapiclient.http import MediaIoBaseUpload

from .drive_manager import _get_drive
from .sheets_client import (
    _get_spreadsheet, _rate_limit, get_api_call_count,
)
from . import config

# ── trade_code → 공종 폴더명 매핑 ──
TRADE_FOLDERS = {
    "wallpaper": "벽지",
    "flooring": "장판",
    "deco_tile": "데코타일",
    "floor_sheet": "바닥시트지",
    "interior_film": "시트지",
    "paint": "페인트",
    "waterproof": "방수에폭시",
    "tile_paint": "페인트",
    "door_paint": "페인트",
    "molding": "몰딩",
    "lighting": "조명",
    "non_slip": "논슬립",
    "stain": "스테인",
    "primer": "프라이머",
    "tile": "타일",
    "faucet": "수전",
    "toilet": "도기",
    "basin": "도기",
    "sink": "싱크대",
    "door": "도어",
    "repair": "보수제",
}

# Drive 폴더 캐시
_folder_cache: dict[str, str] = {}


def _find_or_create_folder(drive, parent_id: str, name: str) -> str:
    """parent_id 아래에서 name 폴더 찾기/생성. 캐시 활용."""
    cache_key = f"{parent_id}/{name}"
    if cache_key in _folder_cache:
        return _folder_cache[cache_key]

    q = f"'{parent_id}' in parents and name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    result = drive.files().list(q=q, fields="files(id)", pageSize=1).execute()
    files = result.get("files", [])

    if files:
        folder_id = files[0]["id"]
    else:
        metadata = {
            "name": name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_id],
        }
        folder = drive.files().create(body=metadata, fields="id").execute()
        folder_id = folder["id"]
        time.sleep(0.5)

    _folder_cache[cache_key] = folder_id
    return folder_id


# ── 브랜드명 정규화 맵 ──
BRAND_NORMALIZE = {
    "LX지인": "LX하우시스",
    "LX하우시스": "LX하우시스",
    "개나리벽지": "개나리벽지",
    "신한벽지": "신한벽지",
    "현대벽지": "현대벽지",
    "FT벽지": "FT벽지",
    "DID벽지": "DID벽지",
    "코스모스벽지": "코스모스벽지",
    "서울벽지": "서울벽지",
    "대림바스": "대림바스",
    "아메리칸스탠다드": "아메리칸스탠다드",
    "한양": "한양",
    "구정마루": "구정마루",
    "동화자연마루": "동화자연마루",
    "이건마루": "이건마루",
    "산들마루": "산들마루",
    "마에스트로": "마에스트로",
}

# ── 공종별 제품명 브랜드 감지 패턴 ──
# (패턴, 브랜드명, 해당 trade_code 리스트 - 빈 리스트면 전체 적용)
BRAND_DETECT_PATTERNS = [
    # === 벽지 ===
    (r"LX하우시스|LG\d{4,}", "LX하우시스", ["wallpaper"]),
    (r"개나리|GN\d{4,}", "개나리벽지", ["wallpaper"]),
    (r"신한|SH\d{4,}", "신한벽지", ["wallpaper"]),
    (r"현대벽지|큐피트|큐브|큐티에", "현대벽지", ["wallpaper"]),
    (r"FT벽지|이룸|벨루체|더뷰", "FT벽지", ["wallpaper"]),
    (r"DID벽지|ID\d{4,}", "DID벽지", ["wallpaper"]),
    (r"코스모스|C4[56]\d{3}", "코스모스벽지", ["wallpaper"]),
    (r"서울벽지|SW\d{3}", "서울벽지", ["wallpaper"]),
    # === 시트지 (제품명에서 브랜드 추출) ===
    (r"LX하우시스|LX\s*ES\d|벤이프", "LX하우시스", ["interior_film"]),
    (r"3M\s|쓰리엠", "3M", ["interior_film"]),
    (r"현대L&C|보닥|BODAQ", "현대L&C", ["interior_film"]),
    (r"KCC글라스|KCC", "KCC글라스", ["interior_film"]),
    # === 페인트 (제품명에서 브랜드 추출) ===
    (r"삼화페인트|아이생각|아이럭스", "삼화페인트", ["paint", "waterproof", "primer", "stain"]),
    (r"노루페인트|순앤수", "노루페인트", ["paint", "waterproof", "primer", "stain"]),
    (r"KCC|숲으로", "KCC", ["paint", "waterproof", "primer", "stain"]),
    (r"벤자민무어|Benjamin", "벤자민무어", ["paint"]),
    (r"듀럭스|Dulux", "듀럭스", ["paint"]),
    # === 도기/수전 (bathnmore 제품) ===
    (r"대림바스|대림\s*B|DERA|BFB.?\d{3}", "대림바스", ["toilet", "basin", "faucet"]),
    (r"아메리칸스[탠텐][다더]드|FB\d{4}", "아메리칸스탠다드", ["toilet", "basin", "faucet"]),
    (r"JEINIS|제니스|SS-[ABC]", "제니스", ["basin", "faucet"]),
    (r"LIDDEL|리델", "리델", ["basin", "faucet"]),
    (r"루바인|LUBAIN", "루바인", ["basin", "faucet"]),
    (r"시원\s", "시원", ["basin", "faucet"]),
    (r"한양\s|HY-", "한양", ["faucet"]),
    (r"게디|GEDY", "게디", ["basin"]),
    (r"에떼르노|ETERNO", "에떼르노", ["basin", "faucet"]),
    # === 마루/바닥재 ===
    (r"LX하우시스|LX\s*Z", "LX하우시스", ["flooring"]),
    (r"KCC|숲\s", "KCC", ["flooring"]),
    (r"동화자연마루|동화\s*N|greendongwha", "동화자연마루", ["flooring"]),
    (r"구정마루|구정", "구정마루", ["flooring"]),
    (r"이건마루|이건", "이건마루", ["flooring"]),
    (r"산들마루|산들", "산들마루", ["flooring"]),
    (r"마에스트로|MAESTRO", "마에스트로", ["flooring"]),
    (r"한솔|HANSOL", "한솔홈데코", ["flooring"]),
    (r"대진|DAEJIN", "대진", ["flooring"]),
    (r"진양|JINYANG", "진양", ["flooring"]),
    (r"현대\s|현대L", "현대", ["flooring"]),
    # === 도어 (daelimwood 제품) ===
    (r"영림|YOUNGLIM|YPD|YSD", "영림도어", ["door"]),
    (r"KCC도어|KCC\s*D", "KCC도어", ["door"]),
    (r"엔토브|ENTOV", "엔토브", ["door"]),
    (r"제이드|JADE", "제이드", ["door"]),
    (r"캡스톤|CAPSTONE", "캡스톤", ["door"]),
    (r"살라만더|SALAMANDER", "살라만더", ["door"]),
    (r"커널시스텍|KERNELSYSTECH", "커널시스텍", ["door"]),
    # === 범용 (trade_code 무관) ===
    (r"LX하우시스", "LX하우시스", []),
]


def _detect_real_brand(name: str, db_brand: str, trade_code: str = "") -> str:
    """제품명에서 실제 브랜드를 감지. DB brand보다 name이 정확.

    trade_code를 참고하여 해당 공종에 맞는 패턴만 매칭.
    """
    for pattern, brand_name, trades in BRAND_DETECT_PATTERNS:
        # trades가 비어있으면 전체 적용, 있으면 해당 trade만
        if trades and trade_code not in trades:
            continue
        if re.search(pattern, name, re.IGNORECASE):
            return brand_name
    # 감지 실패 시 DB brand 정규화
    return BRAND_NORMALIZE.get(db_brand, db_brand)


# ── 타일 사이즈별 분류 ──
def _classify_tile_by_name(name: str) -> str:
    """타일 제품명에서 사이즈/용도 기반 하위분류."""
    n = name.upper()
    if "모자이크" in name or "MOSAIC" in n:
        return "데코모자이크"
    if "헥사곤" in name or "HEXAGON" in n:
        return "헥사곤"
    if "테라조" in name or "TERRAZZO" in n:
        return "테라조"
    # 사이즈 기반 (벽타일은 보통 세로>가로 직사각)
    m = re.search(r"(\d+)\s*[Xx×]\s*(\d+)", name)
    if m:
        w, h = int(m.group(1)), int(m.group(2))
        if w <= 100 and h <= 300:
            return "소형타일"
        if max(w, h) >= 900:
            return "대형타일"
    return "일반타일"


# ── 도어 종류 분류 ──
def _classify_door_by_name(name: str) -> str:
    """도어 제품명에서 현관문/실내도어 분류."""
    if "현관" in name:
        return "현관문"
    if any(k in name for k in ("ABS", "방문", "실내", "슬라이딩")):
        return "실내도어"
    # 브랜드별 기본: 엔토브/캡스톤/살라만더/제이드 = 현관문
    if any(k in name for k in ("엔토브", "캡스톤", "살라만더", "제이드", "커널시스텍")):
        return "현관문"
    # 영림/KCC = 실내도어가 많음
    if any(k in name for k in ("영림", "YPD", "KCC")):
        return "실내도어"
    return "기타도어"


# ── 시트지 분류 (제품명 기반) ──
def _classify_film_by_name(name: str) -> str:
    """시트지 제품명에서 용도/패턴 분류."""
    if any(k in name for k in ("무늬목", "우드", "WOOD")):
        return "무늬목"
    if any(k in name for k in ("대리석", "마블", "MARBLE")):
        return "대리석"
    if any(k in name for k in ("하이그로시", "글로시", "GLOSSY")):
        return "하이그로시"
    if any(k in name for k in ("메탈", "METAL")):
        return "메탈"
    if any(k in name for k in ("회벽", "벽돌", "BRICK")):
        return "회벽벽돌"
    if any(k in name for k in ("패브릭", "FABRIC")):
        return "패브릭"
    if any(k in name for k in ("가죽", "LEATHER")):
        return "가죽"
    # 단색 판단: 컬러명만 있고 패턴 없으면 단색
    if re.search(r"(ES\d{2,3}|단색|솔리드|SOLID)", name, re.IGNORECASE):
        return "단색"
    return "기타"


def _build_folder_path(drive, root_id: str, row: dict) -> str:
    """행 데이터 → 폴더 경로 생성. 공종/브랜드(or 분류)/컬렉션 3단계.

    전략:
    - 벽지/바닥재: 공종 > 브랜드 > 컬렉션
    - 타일: 공종 > 분류(사이즈/용도) > (브랜드 있으면)
    - 도기/수전: 공종 > 브랜드
    - 시트지: 공종 > 브랜드 or 분류(패턴)
    - 도어: 공종 > 종류(현관/실내) > 브랜드
    - 페인트: 공종 > 브랜드

    Returns: 최종 folder_id
    """
    trade_code = row.get("trade_code", "")
    db_brand = row.get("brand", "")
    name = row.get("name", "")

    # 실제 브랜드 감지 (제품명 + trade_code 기반)
    real_brand = _detect_real_brand(name, db_brand, trade_code)

    # raw_payload_json에서 collection 추출
    collection = ""
    raw = row.get("raw_payload_json", "")
    if raw:
        try:
            payload = json.loads(raw)
            collection = payload.get("collection", "")
        except (json.JSONDecodeError, TypeError):
            pass

    # 1단계: 공종 폴더
    trade_name = TRADE_FOLDERS.get(trade_code, trade_code or "기타")
    current_id = _find_or_create_folder(drive, root_id, trade_name)

    # 2단계+: 공종별 분류 전략
    if trade_code == "tile":
        # 타일: 분류(사이즈/용도) > 브랜드(있으면)
        tile_cat = _classify_tile_by_name(name)
        current_id = _find_or_create_folder(drive, current_id, tile_cat)
        if real_brand:
            current_id = _find_or_create_folder(drive, current_id, real_brand)

    elif trade_code == "door":
        # 도어: 종류(현관/실내) > 브랜드(있으면)
        door_type = _classify_door_by_name(name)
        current_id = _find_or_create_folder(drive, current_id, door_type)
        if real_brand:
            current_id = _find_or_create_folder(drive, current_id, real_brand)

    elif trade_code == "interior_film":
        # 시트지: 브랜드 > 패턴분류, 또는 패턴분류만
        if real_brand:
            current_id = _find_or_create_folder(drive, current_id, real_brand)
        else:
            film_cat = _classify_film_by_name(name)
            current_id = _find_or_create_folder(drive, current_id, film_cat)

    elif trade_code in ("basin", "faucet", "toilet"):
        # 도기/수전: 브랜드
        if real_brand:
            current_id = _find_or_create_folder(drive, current_id, real_brand)
        else:
            current_id = _find_or_create_folder(drive, current_id, "기타")

    else:
        # 벽지/바닥재/페인트 등: 브랜드 > 컬렉션
        if real_brand:
            current_id = _find_or_create_folder(drive, current_id, real_brand)
        elif row.get("material_type"):
            current_id = _find_or_create_folder(drive, current_id, row["material_type"])

        # 컬렉션 하위폴더
        if collection:
            current_id = _find_or_create_folder(drive, current_id, collection)

    return current_id


def _safe_filename(name: str) -> str:
    safe = re.sub(r'[\\/:*?"<>|]', '', name)
    safe = re.sub(r'\s+', ' ', safe).strip()
    if len(safe) > 80:
        safe = safe[:80]
    return safe or "unnamed"


def _guess_ext(url: str) -> str:
    lower = url.lower().split("?")[0]
    for ext in (".png", ".webp", ".gif", ".bmp"):
        if lower.endswith(ext):
            return ext
    return ".jpg"


def _col_letter(col_idx: int) -> str:
    """0-based 인덱스 → 엑셀 컬럼 문자 (A, B, ..., Z, AA, ...)"""
    result = ""
    while True:
        result = chr(ord('A') + col_idx % 26) + result
        col_idx = col_idx // 26 - 1
        if col_idx < 0:
            break
    return result


def upload_all(limit: int = 0, trade_filter: str = "", dry_run: bool = False):
    """CrawlerInbox에서 image_file_id가 비어있는 행의 이미지를 Drive에 업로드."""

    print(f"\n{'=' * 60}")
    print(f"Drive 이미지 대량 업로드")
    print(f"폴더 구조: 공종 > 브랜드 > 컬렉션")
    print(f"{'=' * 60}")

    # 1) CrawlerInbox 읽기
    ss = _get_spreadsheet()
    _rate_limit()
    ws = ss.worksheet("CrawlerInbox")
    _rate_limit()
    all_values = ws.get_all_values()

    if len(all_values) <= 1:
        print("  CrawlerInbox가 비어있습니다.")
        return

    headers = all_values[0]
    rows = all_values[1:]
    idx = {h: i for i, h in enumerate(headers)}

    url_i = idx.get("image_url")
    fid_i = idx.get("image_file_id")
    fname_i = idx.get("image_file_name")

    if url_i is None or fid_i is None:
        print("  필수 컬럼(image_url, image_file_id) 없음")
        return

    # 2) 업로드 대상 필터
    targets = []
    for row_num_0, row in enumerate(rows):
        image_url = row[url_i] if url_i < len(row) else ""
        file_id = row[fid_i] if fid_i < len(row) else ""
        trade = row[idx["trade_code"]] if "trade_code" in idx and idx["trade_code"] < len(row) else ""

        if not image_url or file_id:
            continue

        if trade_filter and trade != trade_filter:
            continue

        row_dict = {h: (row[idx[h]] if idx[h] < len(row) else "") for h in idx}
        targets.append({
            "row_num": row_num_0 + 2,
            "row_dict": row_dict,
        })

        if 0 < limit <= len(targets):
            break

    # 공종별 통계
    trade_counts = {}
    for t in targets:
        tc = t["row_dict"].get("trade_code", "unknown")
        trade_counts[tc] = trade_counts.get(tc, 0) + 1

    print(f"\n  업로드 대상: {len(targets)}개")
    for tc, cnt in sorted(trade_counts.items(), key=lambda x: -x[1]):
        tn = TRADE_FOLDERS.get(tc, tc)
        print(f"    {tn} ({tc}): {cnt}개")

    if dry_run:
        print("\n  (dry-run: 실제 업로드 안 함)")
        return

    if not targets:
        print("  업로드할 항목이 없습니다.")
        return

    # 3) Drive 업로드
    drive = _get_drive()
    root_id = config.DRIVE_ROOT_FOLDER_ID

    session = requests.Session()
    session.headers.update(config.REQUEST_HEADERS)

    uploaded = 0
    failed = 0
    consecutive_fails = 0
    batch_updates = []

    fid_col = _col_letter(fid_i)
    fname_col = _col_letter(fname_i) if fname_i is not None else None

    for i, target in enumerate(targets):
        row_dict = target["row_dict"]
        row_num = target["row_num"]
        image_url = row_dict.get("image_url", "")
        name = row_dict.get("name", "unnamed")

        print(f"  [{i+1}/{len(targets)}] {name[:50]}...", end=" ")

        try:
            # 이미지 다운로드
            resp = session.get(image_url, timeout=15)
            resp.raise_for_status()

            content_type = resp.headers.get("Content-Type", "image/jpeg")
            if ";" in content_type:
                content_type = content_type.split(";")[0].strip()

            # 폴더 경로 생성
            folder_id = _build_folder_path(drive, root_id, row_dict)

            # 파일명
            ext = _guess_ext(image_url)
            filename = f"{_safe_filename(name)}{ext}"

            # Drive 업로드
            metadata = {"name": filename, "parents": [folder_id]}
            media = MediaIoBaseUpload(
                io.BytesIO(resp.content), mimetype=content_type, resumable=True
            )
            result = drive.files().create(
                body=metadata, media_body=media, fields="id,name"
            ).execute()

            file_id = result["id"]
            file_name = result["name"]

            # Sheets 배치 업데이트 준비
            batch_updates.append({
                "range": f"{fid_col}{row_num}",
                "values": [[file_id]],
            })
            if fname_col:
                batch_updates.append({
                    "range": f"{fname_col}{row_num}",
                    "values": [[file_name]],
                })

            uploaded += 1
            consecutive_fails = 0
            print(f"OK")

            # 50개마다 배치 업데이트
            if len(batch_updates) >= 100:
                _flush_batch(ws, batch_updates)
                batch_updates = []

            time.sleep(1)  # Drive API 보호

        except Exception as e:
            failed += 1
            consecutive_fails += 1
            err_msg = str(e)[:60]
            print(f"FAIL ({err_msg})")

            if consecutive_fails >= 3:
                print(f"\n  ⚠️ 연속 3회 실패 — 중단합니다.")
                break

            time.sleep(3)

        # 진행률 (50개마다)
        if (i + 1) % 50 == 0:
            pct = (i + 1) / len(targets) * 100
            print(f"\n  --- {pct:.0f}% ({i+1}/{len(targets)}) 성공:{uploaded} 실패:{failed} ---\n")

    # 남은 배치 업데이트
    if batch_updates:
        _flush_batch(ws, batch_updates)

    print(f"\n{'=' * 60}")
    print(f"업로드 완료!")
    print(f"  성공: {uploaded}개")
    print(f"  실패: {failed}개")
    print(f"  Google API 호출: {get_api_call_count()}회")
    print(f"{'=' * 60}")


def _flush_batch(ws, batch_updates):
    """배치 업데이트를 Sheets에 반영."""
    print(f"    -> Sheets 배치 업데이트: {len(batch_updates)}건...", end=" ")
    _rate_limit()
    ws.batch_update(batch_updates, value_input_option="USER_ENTERED")
    print("OK")


def main():
    parser = argparse.ArgumentParser(description="Drive 이미지 대량 업로드")
    parser.add_argument("--limit", type=int, default=0, help="업로드 수 (0=전부)")
    parser.add_argument("--trade", type=str, default="", help="특정 공종만 (예: wallpaper)")
    parser.add_argument("--dry-run", action="store_true", help="분석만")
    args = parser.parse_args()

    upload_all(limit=args.limit, trade_filter=args.trade, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
