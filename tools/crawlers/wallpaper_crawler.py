"""벽지 크롤러 — wallplan.co.kr (고도몰) 기반, 브랜드별 카테고리 크롤링"""
import re
from urllib.parse import urljoin
from .base_crawler import BaseCrawler


# wallplan.co.kr 브랜드별 카테고리 코드
BRAND_CATEGORIES = {
    "신한벽지": "042005",
    "LX하우시스": "042003",
    "개나리벽지": "042004",
}

BASE_URL = "https://www.wallplan.co.kr"
LIST_URL = BASE_URL + "/goods/goods_list.php?cateCd={cate_cd}"
DETAIL_URL = BASE_URL + "/goods/goods_view.php?goodsNo={goods_no}"


class WallplanCrawler(BaseCrawler):
    """wallplan.co.kr에서 브랜드별 벽지 크롤링."""

    source_site = "wallplan.co.kr"
    trade_code = "wallpaper"

    def __init__(self, brand: str):
        super().__init__()
        self.brand = brand
        self.cate_cd = BRAND_CATEGORIES[brand]

    def get_product_urls(self, limit: int = 5) -> list[dict]:
        url = LIST_URL.format(cate_cd=self.cate_cd)
        soup = self.fetch(url)

        items = soup.select(".item_cont")
        results = []

        for item in items[:limit]:
            link = item.select_one("a[href*=goods_view]")
            if not link:
                continue

            href = link["href"]
            # goodsNo 추출
            m = re.search(r"goodsNo=(\d+)", href)
            if not m:
                continue
            goods_no = m.group(1)
            full_url = DETAIL_URL.format(goods_no=goods_no)

            # 목록에서 기본 정보 미리 추출
            name_el = item.select_one(".item_name")
            name_text = name_el.get_text(strip=True) if name_el else ""
            explain_el = item.select_one(".item_name_explain")
            explain_text = explain_el.get_text(strip=True) if explain_el else ""

            # 실크/합지/광폭 등 소재 타입 추출
            material_type = _detect_material_type(name_text + " " + explain_text)

            results.append({
                "url": full_url,
                "goods_no": goods_no,
                "category": f"벽지>{material_type}" if material_type else "벽지",
                "material_type": material_type,
                "list_name": name_text,
                "list_explain": explain_text,
            })

        return results

    def parse_product(self, url: str, meta: dict) -> dict:
        soup = self.fetch(url)

        # 제품명 (상세 페이지)
        detail_tit = soup.select_one(".item_detail_tit")
        name = ""
        if detail_tit:
            name = detail_tit.get_text(strip=True)
            # "공유" 이후 텍스트 제거
            if "공유" in name:
                name = name.split("공유")[0].strip()

        # 목록 이름으로 폴백
        if not name:
            name = meta.get("list_name", "")

        # 제품코드 추출 (예: SH15116-7,?"LX..." 등)
        product_id = ""
        code_match = re.search(r"([A-Z]{1,4}\d{3,5}-?\d*)", name)
        if code_match:
            product_id = code_match.group(1)

        # 가격
        price_el = soup.select_one("dd.price2")
        price_text = price_el.get_text(strip=True) if price_el else ""
        unit_price = _parse_price(price_text)

        # 이미지 — og:image 사용 (가장 안정적)
        image_url = ""
        og_img = soup.find("meta", property="og:image")
        if og_img and og_img.get("content"):
            image_url = og_img["content"]

        # 규격 정보
        spec = meta.get("list_explain", "")

        return {
            "name": name,
            "product_id": product_id,
            "spec": spec,
            "unit": "롤",
            "unit_price": unit_price,
            "image_url": image_url,
            "material_type": meta.get("material_type", ""),
        }


def _detect_material_type(text: str) -> str:
    """텍스트에서 벽지 소재 타입 감지."""
    text_lower = text.lower()
    if "광폭합지" in text_lower or "광폭 합지" in text_lower:
        return "광폭합지"
    if "광폭실크" in text_lower or "광폭 실크" in text_lower:
        return "광폭실크"
    if "실크" in text_lower:
        return "실크"
    if "합지" in text_lower:
        return "합지"
    if "방염" in text_lower:
        return "방염"
    if "수입" in text_lower:
        return "수입"
    return ""


def _parse_price(text: str) -> str:
    """가격 텍스트에서 숫자만 추출. '38,500원' → '38500'"""
    nums = re.sub(r"[^\d]", "", text)
    return nums if nums else ""
