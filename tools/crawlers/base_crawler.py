"""크롤러 베이스 클래스"""
import sys
import io
import uuid
import json
import time
from datetime import datetime, timezone
from abc import ABC, abstractmethod

import requests
from bs4 import BeautifulSoup

# Windows cp949 인코딩 문제 방지
if sys.stdout.encoding != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

from . import config
from .sheets_client import append_to_inbox, append_to_logs
from .drive_manager import ensure_folder_path, upload_image_from_url

# Drive 업로드 활성화 (OAuth 사용자 인증 사용)
ENABLE_DRIVE_UPLOAD = True

# trade_code → Drive 최상위 폴더명 매핑
TRADE_CODE_FOLDERS = {
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
}


class BaseCrawler(ABC):
    """모든 크롤러의 베이스. 서브클래스에서 parse_product_list / parse_product_detail 구현."""

    source_site: str = ""       # 예: "shinhanwall.co.kr"
    brand: str = ""             # 예: "신한벽지"
    trade_code: str = ""        # 예: "wallpaper"

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update(config.REQUEST_HEADERS)
        self.results: list[dict] = []
        self.errors: list[str] = []

    def fetch(self, url: str) -> BeautifulSoup:
        """URL 가져와서 BeautifulSoup 반환. 재시도 1회."""
        for attempt in range(2):
            try:
                resp = self.session.get(url, timeout=config.REQUEST_TIMEOUT)
                resp.raise_for_status()
                resp.encoding = resp.apparent_encoding or "utf-8"
                return BeautifulSoup(resp.text, "html.parser")
            except Exception as e:
                if attempt == 0:
                    time.sleep(1)
                    continue
                raise RuntimeError(f"fetch failed: {url} → {e}") from e

    @abstractmethod
    def get_product_urls(self, limit: int = 5) -> list[dict]:
        """
        제품 URL 리스트 반환.
        각 항목: {"url": ..., "category": "벽지>실크", "material_type": "실크", ...}
        """
        ...

    @abstractmethod
    def parse_product(self, url: str, meta: dict) -> dict:
        """
        제품 페이지를 파싱해서 딕셔너리 반환.
        필수 키: name, spec, unit, unit_price, image_url
        meta에서 category, material_type 등이 전달됨.
        """
        ...

    def crawl(self, limit: int = 5) -> list[dict]:
        """전체 크롤링 플로우. limit = 사이트당 최대 제품 수."""
        print(f"\n[{self.brand}] 제품 목록 수집 중...")
        product_metas = self.get_product_urls(limit=limit)
        print(f"  → {len(product_metas)}개 제품 발견")

        for i, meta in enumerate(product_metas):
            url = meta["url"]
            print(f"  [{i+1}/{len(product_metas)}] {url}")
            try:
                product = self.parse_product(url, meta)
                product = self._enrich(product, meta)
                self.results.append(product)
                print(f"    ✓ {product.get('name', '?')}")
            except Exception as e:
                err_msg = f"[{self.brand}] {url} → {e}"
                self.errors.append(err_msg)
                print(f"    ✗ ERROR: {e}")
            time.sleep(0.5)  # 예의상 딜레이

        return self.results

    def _enrich(self, product: dict, meta: dict) -> dict:
        """공통 필드 채우기 + 이미지 Drive 업로드."""
        now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        ingest_id = f"ING_{uuid.uuid4().hex[:12].upper()}"

        # Drive에 이미지 업로드
        image_file_id = ""
        image_file_name = ""
        image_url = product.get("image_url", "")
        material_type = meta.get("material_type", product.get("material_type", ""))

        if image_url and ENABLE_DRIVE_UPLOAD:
            try:
                # trade_code로 최상위 폴더 결정
                trade_folder = TRADE_CODE_FOLDERS.get(self.trade_code, self.trade_code)
                folder_path = [trade_folder]
                if self.brand:
                    folder_path.append(self.brand)
                if material_type:
                    folder_path.append(material_type)
                folder_id = ensure_folder_path(folder_path)
                ext = _guess_ext(image_url)
                # 상품명으로 파일 저장 (특수문자 제거)
                safe_name = _safe_filename(product.get("name", "") or ingest_id)
                fname = f"{safe_name}{ext}"
                upload_result = upload_image_from_url(image_url, folder_id, fname)
                image_file_id = upload_result["file_id"]
                image_file_name = upload_result["file_name"]
                print(f"    Drive upload: {image_file_name}")
            except Exception as e:
                print(f"    Drive upload skip: {e}")

        tags = f"trade:{self.trade_code}"
        if material_type:
            tags += f", texture:{material_type}"

        return {
            "ingest_id": ingest_id,
            "source_site": self.source_site,
            "source_category": meta.get("category", ""),
            "source_product_id": product.get("product_id", ""),
            "source_url": meta.get("url", ""),
            "name": product.get("name", ""),
            "brand": self.brand,
            "spec": product.get("spec", ""),
            "unit": product.get("unit", "롤"),
            "unit_price": str(product.get("unit_price", "")),
            "image_url": image_url,
            "image_file_id": image_file_id,
            "image_file_name": image_file_name,
            "material_type": material_type,
            "trade_code": self.trade_code,
            "space_type": "",
            "price_band": "",
            "is_representative": "",
            "expose_to_prequote": "",
            "recommendation_score_base": "",
            "recommendation_note": "",
            "tags_summary": tags,
            "tag_payload_json": "",
            "group_payload_json": "",
            "raw_payload_json": json.dumps(product, ensure_ascii=False, default=str),
            "sync_status": "pending",
            "sync_message": "",
            "created_at": now,
            "synced_at": "",
            "material_id": "",
            "run_id": "",  # run_trial에서 채움
        }

    def save(self, run_id: str):
        """결과를 CrawlerInbox에 저장하고 CrawlerLogs에 기록."""
        for r in self.results:
            r["run_id"] = run_id

        if self.results:
            count = append_to_inbox(self.results)
            print(f"  → CrawlerInbox에 {count}행 저장 완료")

        now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        append_to_logs({
            "run_id": run_id,
            "run_at": now,
            "source_site": self.source_site,
            "job_type": "trial_crawl",
            "row_count": str(len(self.results) + len(self.errors)),
            "success_count": str(len(self.results)),
            "error_count": str(len(self.errors)),
            "status": "done" if not self.errors else "partial",
            "message": "; ".join(self.errors[:3]) if self.errors else "OK",
        })


def _safe_filename(name: str) -> str:
    """파일명에 쓸 수 없는 문자 제거."""
    import re
    # 파일시스템 금지 문자 제거
    safe = re.sub(r'[\\/:*?"<>|]', '', name)
    # 연속 공백 정리
    safe = re.sub(r'\s+', ' ', safe).strip()
    # 너무 길면 자르기
    if len(safe) > 100:
        safe = safe[:100]
    return safe or "unnamed"


def _guess_ext(url: str) -> str:
    """URL에서 이미지 확장자 추출."""
    lower = url.lower().split("?")[0]
    for ext in (".png", ".webp", ".gif", ".bmp"):
        if lower.endswith(ext):
            return ext
    return ".jpg"
