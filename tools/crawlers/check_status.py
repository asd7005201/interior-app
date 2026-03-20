"""크롤링 + Drive 업로드 현황 리포트"""
import json
from tools.crawlers.sheets_client import read_inbox_all


def main():
    rows = read_inbox_all()

    cats = {}
    for r in rows:
        tc = r.get("trade_code", "unknown")
        has_img = bool(r.get("image_file_id"))
        if tc not in cats:
            cats[tc] = {"total": 0, "uploaded": 0}
        cats[tc]["total"] += 1
        if has_img:
            cats[tc]["uploaded"] += 1

    print("=" * 60)
    print("CrawlerInbox 현황")
    print("=" * 60)
    total_up = 0
    for tc, v in sorted(cats.items(), key=lambda x: -x[1]["total"]):
        pct = int(v["uploaded"] / v["total"] * 100) if v["total"] else 0
        bar = "#" * (pct // 5) + "." * (20 - pct // 5)
        print("  %-20s: %5d개 | Drive: %5d개 [%s] %d%%" % (tc, v["total"], v["uploaded"], bar, pct))
        total_up += v["uploaded"]

    print()
    print("  전체: %d개 / Drive 업로드: %d개" % (len(rows), total_up))
    print("=" * 60)


if __name__ == "__main__":
    main()
