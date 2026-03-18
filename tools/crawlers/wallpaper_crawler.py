"""벽지 크롤러 — wallplan.co.kr (고도몰) 기반, 브랜드별/컬렉션별 크롤링"""
import re
import time
from .base_crawler import BaseCrawler


# ── 카테고리 기반 크롤링 (롤 벽지) ──
BRAND_CATEGORIES = {
    "신한벽지": "042005",
    "LX하우시스": "042003",
    "개나리벽지": "042004",
}

# ── 검색 기반 크롤링 (프리미엄 컬렉션) ──
PREMIUM_COLLECTIONS = {
    "파사드":  {"brand": "신한벽지",  "keyword": "파사드"},
    "디아망":  {"brand": "LX하우시스", "keyword": "디아망"},
    "로하스":  {"brand": "개나리벽지", "keyword": "로하스"},
}

BASE_URL = "https://www.wallplan.co.kr"
LIST_URL = BASE_URL + "/goods/goods_list.php?cateCd={cate_cd}&page={page}"
SEARCH_URL = BASE_URL + "/goods/goods_search.php?keyword={keyword}&page={page}"
DETAIL_URL = BASE_URL + "/goods/goods_view.php?goodsNo={goods_no}"

# 2022년 이후 제품 판별용 — 제품코드 번호가 높을수록 최신
# 각 브랜드별 2022년 이전 코드 상한선 (이 이하는 오래된 제품)
OLD_CODE_THRESHOLDS = {
    "SH": 6800,      # 신한: SH6800 이하 = 아이리스(구) → 제외
    "LG": 49500,     # LX(구 LG): LG49500 이하 = 휘앙세(구) → 제외
}


class WallplanCrawler(BaseCrawler):
    """wallplan.co.kr에서 카테고리 기반 벽지 크롤링 (페이지네이션 지원)."""

    source_site = "wallplan.co.kr"
    trade_code = "wallpaper"

    def __init__(self, brand: str):
        super().__init__()
        self.brand = brand
        self.cate_cd = BRAND_CATEGORIES[brand]

    def get_product_urls(self, limit: int = 0) -> list[dict]:
        """전체 페이지 순회. limit=0이면 전부."""
        all_results = []
        page = 1

        while True:
            url = LIST_URL.format(cate_cd=self.cate_cd, page=page)
            soup = self.fetch(url)
            items = soup.select(".item_cont")

            if not items:
                break

            for item in items:
                meta = _parse_list_item(item)
                if meta is None:
                    continue

                # 오래된 제품 필터링
                if _is_old_product(meta["list_name"]):
                    continue

                all_results.append(meta)
                if 0 < limit <= len(all_results):
                    return all_results

            page += 1
            time.sleep(0.3)

        return all_results

    def parse_product(self, url: str, meta: dict) -> dict:
        return _parse_detail_page(self, url, meta)


class WallplanSearchCrawler(BaseCrawler):
    """wallplan.co.kr에서 키워드 검색 기반 크롤링 (파사드/디아망/로하스 등)."""

    source_site = "wallplan.co.kr"
    trade_code = "wallpaper"

    def __init__(self, collection_name: str):
        super().__init__()
        info = PREMIUM_COLLECTIONS[collection_name]
        self.brand = info["brand"]
        self.keyword = info["keyword"]
        self.collection_name = collection_name

    def get_product_urls(self, limit: int = 0) -> list[dict]:
        all_results = []
        page = 1

        while True:
            url = SEARCH_URL.format(keyword=self.keyword, page=page)
            soup = self.fetch(url)
            items = soup.select(".item_cont")

            if not items:
                break

            for item in items:
                meta = _parse_list_item(item)
                if meta is None:
                    continue

                # 컬렉션 태그 추가
                meta["collection"] = self.collection_name
                all_results.append(meta)
                if 0 < limit <= len(all_results):
                    return all_results

            page += 1
            time.sleep(0.3)

        return all_results

    def parse_product(self, url: str, meta: dict) -> dict:
        return _parse_detail_page(self, url, meta)


# ── 공통 파싱 함수 ──

def _parse_list_item(item) -> dict | None:
    """목록 페이지의 제품 카드 하나를 파싱."""
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

    material_type = _detect_material_type(name_text + " " + explain_text)

    return {
        "url": DETAIL_URL.format(goods_no=goods_no),
        "goods_no": goods_no,
        "category": f"벽지>{material_type}" if material_type else "벽지",
        "material_type": material_type,
        "list_name": name_text,
        "list_explain": explain_text,
    }


def _parse_detail_page(crawler, url: str, meta: dict) -> dict:
    """상세 페이지 파싱."""
    soup = crawler.fetch(url)

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
    code_match = re.search(r"([A-Z]{1,4}\d{3,5}-?\d*)", name)
    if code_match:
        product_id = code_match.group(1)

    # 가격
    price_el = soup.select_one("dd.price2")
    price_text = price_el.get_text(strip=True) if price_el else ""
    unit_price = _parse_price(price_text)

    # 이미지 (og:image)
    image_url = ""
    og_img = soup.find("meta", property="og:image")
    if og_img and og_img.get("content"):
        image_url = og_img["content"]

    # 규격
    spec = meta.get("list_explain", "")

    # 컬렉션 정보
    collection = meta.get("collection", "")
    # explain에서 컬렉션명 추출 (예: [스케치], [파사드], [디아망])
    coll_match = re.search(r"\[([^\]]+)\]", spec)
    if coll_match and not collection:
        collection = coll_match.group(1)

    return {
        "name": name,
        "product_id": product_id,
        "spec": spec,
        "unit": "롤",
        "unit_price": unit_price,
        "image_url": image_url,
        "material_type": meta.get("material_type", ""),
        "collection": collection,
    }


def _is_old_product(name_text: str) -> bool:
    """제품코드 기반으로 2022년 이전 제품 판별."""
    for prefix, threshold in OLD_CODE_THRESHOLDS.items():
        m = re.search(rf"{prefix}(\d+)", name_text)
        if m:
            code_num = int(m.group(1))
            if code_num <= threshold:
                return True
    return False


def _detect_material_type(text: str) -> str:
    """텍스트에서 벽지 소재 타입 감지."""
    if "광폭합지" in text or "광폭 합지" in text:
        return "광폭합지"
    if "광폭실크" in text or "광폭 실크" in text:
        return "광폭실크"
    if "실크" in text:
        return "실크"
    if "합지" in text:
        return "합지"
    if "방염" in text:
        return "방염"
    if "수입" in text:
        return "수입"
    return ""


def _parse_price(text: str) -> str:
    """가격 텍스트에서 숫자만 추출. '38,500원' → '38500'"""
    nums = re.sub(r"[^\d]", "", text)
    return nums if nums else ""
