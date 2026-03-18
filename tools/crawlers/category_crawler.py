"""범용 카테고리 크롤러 — wallplan.co.kr 기반, 다양한 자재 카테고리 지원"""
import re
import time
from .base_crawler import BaseCrawler


BASE_URL = "https://www.wallplan.co.kr"
LIST_URL = BASE_URL + "/goods/goods_list.php?cateCd={cate_cd}&page={page}"
DETAIL_URL = BASE_URL + "/goods/goods_view.php?goodsNo={goods_no}"

# ── 크롤링 대상 카테고리 정의 ──
MATERIAL_CATEGORIES = {
    # ── 장판 ──
    "장판-KCC": {"cate_cd": "046001001", "trade_code": "flooring", "material_type": "장판", "brand": "KCC"},
    "장판-LG(LX)": {"cate_cd": "046001002", "trade_code": "flooring", "material_type": "장판", "brand": "LX하우시스"},
    "장판-대진": {"cate_cd": "046001007", "trade_code": "flooring", "material_type": "장판", "brand": "대진"},
    "장판-현대": {"cate_cd": "046001006", "trade_code": "flooring", "material_type": "장판", "brand": "현대"},
    "장판-진양": {"cate_cd": "046001003", "trade_code": "flooring", "material_type": "장판", "brand": "진양"},

    # ── 데코타일 (마루 대용) ──
    "데코타일-접착": {"cate_cd": "046002001", "trade_code": "deco_tile", "material_type": "접착식", "brand": ""},
    "바닥시트지": {"cate_cd": "046003", "trade_code": "floor_sheet", "material_type": "바닥시트", "brand": ""},

    # ── 시트지/인테리어필름 ──
    "시트지-단색": {"cate_cd": "044001", "trade_code": "interior_film", "material_type": "단색", "brand": ""},
    "시트지-몰딩샷시": {"cate_cd": "044002", "trade_code": "interior_film", "material_type": "몰딩전용", "brand": ""},
    "시트지-대리석": {"cate_cd": "044011", "trade_code": "interior_film", "material_type": "대리석", "brand": ""},
    "시트지-무늬목": {"cate_cd": "044009", "trade_code": "interior_film", "material_type": "무늬목", "brand": ""},
    "시트지-화벽벽돌": {"cate_cd": "044017", "trade_code": "interior_film", "material_type": "화벽벽돌", "brand": ""},
    "시트지-하이그로시": {"cate_cd": "044012", "trade_code": "interior_film", "material_type": "하이그로시", "brand": ""},
    "시트지-메탈": {"cate_cd": "044015", "trade_code": "interior_film", "material_type": "메탈", "brand": ""},
    "시트지-패브릭": {"cate_cd": "044018", "trade_code": "interior_film", "material_type": "패브릭", "brand": ""},
    "시트지-가죽": {"cate_cd": "044014", "trade_code": "interior_film", "material_type": "가죽", "brand": ""},

    # ── 페인트 (하위 카테고리 분리) ──
    "페인트-벽지벽면용": {"cate_cd": "043008", "trade_code": "paint", "material_type": "벽면용", "brand": ""},
    "페인트-방문가구용": {"cate_cd": "043009", "trade_code": "paint", "material_type": "방문가구용", "brand": ""},
    "페인트-타일욕실": {"cate_cd": "043010", "trade_code": "paint", "material_type": "타일욕실용", "brand": ""},
    "페인트-외부철재": {"cate_cd": "043012", "trade_code": "paint", "material_type": "외부철재", "brand": ""},
    "페인트-칠판자석": {"cate_cd": "043015", "trade_code": "paint", "material_type": "칠판자석", "brand": ""},

    # ── 방수/에폭시 (줄눈/탄성 포함) ──
    "방수에폭시바닥": {"cate_cd": "043014", "trade_code": "waterproof", "material_type": "방수에폭시", "brand": ""},

    # ── 굽도리/걸레받이 (몰딩) ──
    "굽도리걸레받이": {"cate_cd": "046004", "trade_code": "molding", "material_type": "걸레받이", "brand": ""},

    # ── 조명 ──
    "조명-방등": {"cate_cd": "048019005", "trade_code": "lighting", "material_type": "방등", "brand": ""},

    # ── 논슬립 ──
    "논슬립": {"cate_cd": "046009", "trade_code": "non_slip", "material_type": "논슬립", "brand": ""},

    # ── 스테인/바니시 ──
    "스테인바니시": {"cate_cd": "043011", "trade_code": "stain", "material_type": "스테인", "brand": ""},

    # ── 방문/가구용 페인트 ──
    "방문가구페인트": {"cate_cd": "043009", "trade_code": "door_paint", "material_type": "방문가구용", "brand": ""},

    # ── 젯소/프라이머 ──
    "젯소프라이머": {"cate_cd": "043013", "trade_code": "primer", "material_type": "프라이머", "brand": ""},
}


class WallplanCategoryCrawler(BaseCrawler):
    """wallplan.co.kr 범용 카테고리 크롤러."""

    source_site = "wallplan.co.kr"

    def __init__(self, category_key: str):
        super().__init__()
        info = MATERIAL_CATEGORIES[category_key]
        self.cate_cd = info["cate_cd"]
        self.trade_code = info["trade_code"]
        self._material_type = info["material_type"]
        self.brand = info["brand"]
        self.category_key = category_key

    def get_product_urls(self, limit: int = 0) -> list[dict]:
        all_results = []
        page = 1
        MAX_PAGES = 50  # 안전장치: 최대 50페이지 (1000개)

        while page <= MAX_PAGES:
            url = LIST_URL.format(cate_cd=self.cate_cd, page=page)
            soup = self.fetch(url)
            items = soup.select(".item_cont")

            if not items:
                break

            for item in items:
                meta = _parse_list_item(item, self._material_type)
                if meta is None:
                    continue
                all_results.append(meta)
                if 0 < limit <= len(all_results):
                    return all_results

            page += 1
            time.sleep(0.3)

        return all_results

    def parse_product(self, url: str, meta: dict) -> dict:
        soup = self.fetch(url)

        # 제품명
        detail_tit = soup.select_one(".item_detail_tit")
        name = ""
        if detail_tit:
            name = detail_tit.get_text(strip=True)
            if "공유" in name:
                name = name.split("공유")[0].strip()
        if not name:
            name = meta.get("list_name", "")

        # 제품코드
        product_id = ""
        code_match = re.search(r"([A-Z]{1,5}\d{3,6}-?\d*)", name)
        if code_match:
            product_id = code_match.group(1)

        # 가격
        price_el = soup.select_one("dd.price2")
        price_text = price_el.get_text(strip=True) if price_el else ""
        unit_price = re.sub(r"[^\d]", "", price_text)

        # 이미지
        image_url = ""
        og_img = soup.find("meta", property="og:image")
        if og_img and og_img.get("content"):
            image_url = og_img["content"]

        # 규격
        spec = meta.get("list_explain", "")

        return {
            "name": name,
            "product_id": product_id,
            "spec": spec,
            "unit": _detect_unit(spec, self.trade_code),
            "unit_price": unit_price,
            "image_url": image_url,
            "material_type": meta.get("material_type", self._material_type),
        }


def _parse_list_item(item, default_type: str = "") -> dict | None:
    link = item.select_one("a[href*=goods_view]")
    if not link:
        return None

    m = re.search(r"goodsNo=(\d+)", link["href"])
    if not m:
        return None

    goods_no = m.group(1)
    name_el = item.select_one(".item_name")
    name_text = name_el.get_text(strip=True) if name_el else ""
    explain_el = item.select_one(".item_name_explain")
    explain_text = explain_el.get_text(strip=True) if explain_el else ""

    return {
        "url": DETAIL_URL.format(goods_no=goods_no),
        "goods_no": goods_no,
        "category": default_type if default_type else "",
        "material_type": default_type,
        "list_name": name_text,
        "list_explain": explain_text,
    }


def _detect_unit(spec: str, trade_code: str) -> str:
    """규격 텍스트에서 단위 추출."""
    spec_lower = spec.lower()
    if "롤" in spec_lower:
        return "롤"
    if "장" in spec_lower or "매" in spec_lower:
        return "장"
    if "통" in spec_lower or "캔" in spec_lower:
        return "통"
    if trade_code == "paint":
        return "통"
    if trade_code in ("flooring", "floor_sheet"):
        return "m"
    return "개"
