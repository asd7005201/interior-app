"""Cafe24 범용 크롤러 — bathnmore / maruall / daelimwood 지원.

Cafe24 표준 구조:
- 목록: ul.prdList > li, 상품 링크 .name a, 이미지 .thumbnail img
- 상세: JSON-LD Product 스키마 (가장 안정적)
- 페이지네이션: ?page=N

사이트별 차이:
- maruall: SSR (requests만으로 OK)
- bathnmore/daelimwood: 일부 JS 렌더링 → 상세 페이지 JSON-LD 의존
"""
import json
import re
import time
from .base_crawler import BaseCrawler

# ── 사이트 + 카테고리 정의 ──

CAFE24_SITES = {
    # ── bathnmore.co.kr: 타일 + 수전/도기 ──
    "bathnmore-타일-바닥": {
        "base_url": "https://bathnmore.co.kr",
        "cate_path": "/category/타일/132/",
        "source_site": "bathnmore.co.kr",
        "trade_code": "tile",
        "material_type": "바닥타일",
        "brand": "",
    },
    "bathnmore-세면기": {
        "base_url": "https://bathnmore.co.kr",
        "cate_path": "/category/세면기/27/",
        "source_site": "bathnmore.co.kr",
        "trade_code": "basin",
        "material_type": "세면기",
        "brand": "",
    },
    "bathnmore-양변기-대림": {
        "base_url": "https://bathnmore.co.kr",
        "cate_path": "/category/대림-양변기/692/",
        "source_site": "bathnmore.co.kr",
        "trade_code": "toilet",
        "material_type": "양변기",
        "brand": "대림바스",
    },
    "bathnmore-양변기-AS": {
        "base_url": "https://bathnmore.co.kr",
        "cate_path": "/category/아메리칸스탠다드-양변기/693/",
        "source_site": "bathnmore.co.kr",
        "trade_code": "toilet",
        "material_type": "양변기",
        "brand": "아메리칸스탠다드",
    },
    "bathnmore-수전-대림": {
        "base_url": "https://bathnmore.co.kr",
        "cate_path": "/category/대림수전/26/",
        "source_site": "bathnmore.co.kr",
        "trade_code": "faucet",
        "material_type": "수전",
        "brand": "대림바스",
    },
    "bathnmore-수전-한양": {
        "base_url": "https://bathnmore.co.kr",
        "cate_path": "/category/한양수전/644/",
        "source_site": "bathnmore.co.kr",
        "trade_code": "faucet",
        "material_type": "수전",
        "brand": "한양",
    },
    "bathnmore-욕조": {
        "base_url": "https://bathnmore.co.kr",
        "cate_path": "/category/욕조/30/",
        "source_site": "bathnmore.co.kr",
        "trade_code": "basin",
        "material_type": "욕조",
        "brand": "",
    },

    # ── maruall.com: 마루 ──
    "maruall-구정": {
        "base_url": "https://maruall.com",
        "cate_path": "/category/구정마루/23/",
        "source_site": "maruall.com",
        "trade_code": "flooring",
        "material_type": "강마루",
        "brand": "구정마루",
    },
    "maruall-동화": {
        "base_url": "https://maruall.com",
        "cate_path": "/category/동화자연마루/67/",
        "source_site": "maruall.com",
        "trade_code": "flooring",
        "material_type": "강마루",
        "brand": "동화자연마루",
    },
    "maruall-LX": {
        "base_url": "https://maruall.com",
        "cate_path": "/category/LX하우시스/85/",
        "source_site": "maruall.com",
        "trade_code": "flooring",
        "material_type": "강마루",
        "brand": "LX하우시스",
    },
    "maruall-산들": {
        "base_url": "https://maruall.com",
        "cate_path": "/category/산들마루/65/",
        "source_site": "maruall.com",
        "trade_code": "flooring",
        "material_type": "강마루",
        "brand": "산들마루",
    },
    "maruall-이건": {
        "base_url": "https://maruall.com",
        "cate_path": "/category/이건마루/66/",
        "source_site": "maruall.com",
        "trade_code": "flooring",
        "material_type": "강마루",
        "brand": "이건마루",
    },
    "maruall-마에스트로": {
        "base_url": "https://maruall.com",
        "cate_path": "/category/마에스트로-마루/87/",
        "source_site": "maruall.com",
        "trade_code": "flooring",
        "material_type": "강마루",
        "brand": "마에스트로",
    },

    # ── daelimwood.com: 도어 ──
    "daelimwood-도어": {
        "base_url": "https://daelimwood.com",
        "cate_path": "/category/실내실외-도어/28/",
        "source_site": "daelimwood.com",
        "trade_code": "door",
        "material_type": "도어",
        "brand": "",
    },
}


class Cafe24Crawler(BaseCrawler):
    """Cafe24 기반 쇼핑몰 범용 크롤러."""

    def __init__(self, site_key: str):
        super().__init__()
        info = CAFE24_SITES[site_key]
        self.base_url = info["base_url"]
        self.cate_path = info["cate_path"]
        self.source_site = info["source_site"]
        self.trade_code = info["trade_code"]
        self._material_type = info["material_type"]
        self.brand = info["brand"]
        self.site_key = site_key

    def get_product_urls(self, limit: int = 0) -> list[dict]:
        """카테고리 페이지 순회하며 상품 URL 수집."""
        all_results = []
        page = 1
        MAX_PAGES = 30

        while page <= MAX_PAGES:
            url = f"{self.base_url}{self.cate_path}?page={page}"
            print(f"    목록 페이지 {page}: {url}")

            try:
                soup = self.fetch(url)
            except Exception as e:
                print(f"    목록 페이지 실패: {e}")
                break

            # 방법 1: SSR 목록 파싱 (maruall 등)
            items = self._parse_list_ssr(soup)

            # 방법 2: JS 변수에서 상품 번호 추출 (bathnmore 등)
            if not items:
                items = self._parse_list_js(soup)

            if not items:
                break

            for item in items:
                all_results.append(item)
                if 0 < limit <= len(all_results):
                    return all_results

            page += 1
            time.sleep(1)  # 외부 사이트이므로 1초 딜레이

        return all_results

    def _parse_list_ssr(self, soup) -> list[dict]:
        """SSR 렌더링된 목록 파싱 (Cafe24 표준)."""
        items = []

        # ul.prdList > li 또는 .prdList li
        product_list = soup.select("ul.prdList > li")
        if not product_list:
            product_list = soup.select(".prdList li")

        for li in product_list:
            # 상품 링크
            link = li.select_one(".name a") or li.select_one(".description a") or li.select_one("a[href*='/product/']")
            if not link:
                continue

            href = link.get("href", "")
            if not href:
                continue

            # 절대 URL
            if href.startswith("/"):
                href = self.base_url + href

            # 상품명
            name = link.get_text(strip=True)

            # 이미지
            img = li.select_one(".thumbnail img") or li.select_one("img")
            img_src = ""
            if img:
                img_src = img.get("src", "") or img.get("data-original", "") or img.get("data-src", "")
                if img_src.startswith("//"):
                    img_src = "https:" + img_src

            # 가격 (ec-data-price attribute 또는 텍스트)
            price = ""
            price_el = li.select_one("[ec-data-price]")
            if price_el:
                price = price_el.get("ec-data-price", "")
            if not price:
                price_el = li.select_one(".price span") or li.select_one(".price")
                if price_el:
                    price_text = price_el.get_text(strip=True)
                    price = re.sub(r"[^\d]", "", price_text)

            # product_no 추출
            product_no = ""
            pno_match = re.search(r"/product/[^/]+/(\d+)/", href)
            if pno_match:
                product_no = pno_match.group(1)
            if not product_no:
                pno_match = re.search(r"product_no=(\d+)", href)
                if pno_match:
                    product_no = pno_match.group(1)
            # li id에서 추출
            if not product_no:
                li_id = li.get("id", "")
                id_match = re.search(r"anchorBoxId_(\d+)", li_id)
                if id_match:
                    product_no = id_match.group(1)

            items.append({
                "url": href,
                "category": self._material_type,
                "material_type": self._material_type,
                "list_name": name,
                "list_explain": "",
                "list_price": price,
                "list_image": img_src,
                "product_no": product_no,
            })

        return items

    def _parse_list_js(self, soup) -> list[dict]:
        """JS 변수에서 상품 번호 추출 (동적 렌더링 사이트)."""
        items = []
        scripts = soup.find_all("script")

        for script in scripts:
            text = script.string or ""
            # aProductPurchaseInfo_{ID} 패턴 검색
            matches = re.findall(r"aProductPurchaseInfo_(\d+)", text)
            for product_no in set(matches):
                detail_url = f"{self.base_url}/product/detail.html?product_no={product_no}"
                items.append({
                    "url": detail_url,
                    "category": self._material_type,
                    "material_type": self._material_type,
                    "list_name": "",
                    "list_explain": "",
                    "list_price": "",
                    "list_image": "",
                    "product_no": product_no,
                })

        return items

    def parse_product(self, url: str, meta: dict) -> dict:
        """상세 페이지 파싱. JSON-LD 우선, 폴백으로 HTML 파싱."""
        soup = self.fetch(url)

        # 1) JSON-LD 파싱 시도
        product_data = self._parse_json_ld(soup)

        # 2) HTML 폴백
        if not product_data.get("name"):
            product_data = self._parse_detail_html(soup, meta)

        # 메타 정보 보충
        if not product_data.get("name") and meta.get("list_name"):
            product_data["name"] = meta["list_name"]
        if not product_data.get("image_url") and meta.get("list_image"):
            product_data["image_url"] = meta["list_image"]
        if not product_data.get("unit_price") and meta.get("list_price"):
            product_data["unit_price"] = meta["list_price"]

        product_data["product_id"] = meta.get("product_no", "")
        product_data["material_type"] = meta.get("material_type", self._material_type)

        return product_data

    def _parse_json_ld(self, soup) -> dict:
        """JSON-LD Product 스키마에서 데이터 추출."""
        scripts = soup.find_all("script", type="application/ld+json")
        for script in scripts:
            try:
                data = json.loads(script.string or "")
                if isinstance(data, list):
                    data = data[0]
                if data.get("@type") == "Product":
                    name = data.get("name", "")
                    image = ""
                    img_field = data.get("image", "")
                    if isinstance(img_field, list) and img_field:
                        image = img_field[0]
                    elif isinstance(img_field, str):
                        image = img_field
                    if image and image.startswith("//"):
                        image = "https:" + image

                    price = ""
                    offers = data.get("offers", {})
                    if isinstance(offers, list) and offers:
                        offers = offers[0]
                    if isinstance(offers, dict):
                        price = str(offers.get("price", ""))

                    brand_name = ""
                    brand = data.get("brand", {})
                    if isinstance(brand, dict):
                        brand_name = brand.get("name", "")

                    spec = data.get("description", "")
                    if len(spec) > 200:
                        spec = spec[:200]

                    return {
                        "name": name,
                        "image_url": image,
                        "unit_price": re.sub(r"[^\d]", "", price),
                        "spec": spec,
                        "unit": "개",
                        "brand_from_page": brand_name,
                    }
            except (json.JSONDecodeError, TypeError, KeyError):
                continue
        return {}

    def _parse_detail_html(self, soup, meta: dict) -> dict:
        """HTML에서 상품 정보 추출 (JSON-LD 실패 시 폴백)."""
        # 상품명
        name = ""
        name_el = soup.select_one(".headingArea h2") or soup.select_one("h2")
        if name_el:
            name = name_el.get_text(strip=True)

        # 가격
        price = ""
        price_el = soup.select_one("#span_product_price_text") or soup.select_one(".price span")
        if price_el:
            price = re.sub(r"[^\d]", "", price_el.get_text(strip=True))

        # 이미지
        image_url = ""
        og_img = soup.find("meta", property="og:image")
        if og_img and og_img.get("content"):
            image_url = og_img["content"]
            if image_url.startswith("//"):
                image_url = "https:" + image_url

        return {
            "name": name,
            "image_url": image_url,
            "unit_price": price,
            "spec": "",
            "unit": "개",
        }
