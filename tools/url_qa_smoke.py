#!/usr/bin/env python3
import argparse
import json
import os
import sys
import time

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


DEFAULTS = {
    "quote_base_url": "https://script.google.com/macros/s/AKfycbw138HXXYxd6_qnh98eQpc2rkkdt1CSmlqI3hVWZhiDW5Sg8blfeqCwGf68P2Lv_w-y/exec",
    "prequote_base_url": "https://script.google.com/macros/s/AKfycby0_7tpS0-6ieNzsaauo1WVPpaXVFrAq9fpM4mISpd6bKbb_-MoaD1RNSzGXoqUVbN2Iw/exec",
    "admin_password": "test",
}


def app_frame(page):
    page.wait_for_timeout(3000)
    if "accounts.google.com" in page.url:
        raise RuntimeError("Google sign-in required. Run tools/capture_google_session.py first and pass --storage-state.")
    if len(page.frames) < 3:
        raise RuntimeError(f"expected app frame, got {len(page.frames)} frames")
    return page.frames[2]


def quote_edit_smoke(page, base_url, admin_password):
    frame = app_frame(page)
    out = {"checks": []}

    def check(name, fn):
        try:
            out["checks"].append({"name": name, "status": "PASS", "detail": fn()})
        except Exception as exc:
            out["checks"].append({"name": name, "status": "FAIL", "detail": str(exc)})

    check("render", lambda: frame.locator("#adminPw").wait_for(timeout=15000))
    check("auth_fill", lambda: frame.locator("#adminPw").fill(admin_password))
    check("create_quote", lambda: frame.locator("#btnCreate").click())
    check("editor_visible", lambda: frame.locator("#editor").wait_for(state="visible", timeout=20000))
    check("row_count_before_add", lambda: frame.locator("#itemsBody tr.item-row").count())
    check("add_row", lambda: frame.locator("#btnAddRow").click())
    check("row_count_after_add", lambda: frame.locator("#itemsBody tr.item-row").count())
    check("status", lambda: frame.locator("#status").inner_text())
    return out


def quote_dashboard_smoke(page, base_url, admin_password, query):
    frame = app_frame(page)
    frame.locator("#adminPw").fill(admin_password)
    frame.locator("#q").fill(query)
    frame.locator("#btnSearch").click()
    page.wait_for_timeout(5000)
    return {
        "status": frame.locator("#status").inner_text(),
        "summary": {
            "month_total": frame.locator("#sumMonth").inner_text(),
            "pending": frame.locator("#sumPending").inner_text(),
            "conversion": frame.locator("#statConversion").inner_text(),
        },
        "body_excerpt": frame.locator("body").inner_text()[:1200],
    }


def quote_catalog_smoke(page, base_url, admin_password, query):
    frame = app_frame(page)
    frame.locator("#adminPw").fill(admin_password)
    frame.locator("#matQuery").fill(query)
    frame.locator("#btnSearch").click()
    page.wait_for_timeout(7000)
    return {
        "status": frame.locator("#status").inner_text(),
        "result_meta": frame.locator("#resultMeta").inner_text(),
        "body_excerpt": frame.locator("#catalogBody").inner_text()[:1200],
    }


def prequote_public_smoke(page, base_url):
    frame = app_frame(page)
    landing = frame.locator("body").inner_text()[:1200]
    frame.get_by_text("가견적 시작하기").click()
    page.wait_for_timeout(1000)
    frame.get_by_text("주거").first.click()
    page.wait_for_timeout(1000)
    after_start = frame.locator("body").inner_text()[:1600]
    return {"landing_excerpt": landing, "after_start_excerpt": after_start}


def prequote_admin_smoke(page, base_url, admin_password):
    frame = app_frame(page)
    frame.locator("#loginPassword").fill(admin_password)
    frame.get_by_text("로그인").click()
    try:
        frame.get_by_text("가견적 관리자 워크스페이스").wait_for(timeout=20000)
    except PlaywrightTimeoutError:
        page.wait_for_timeout(12000)
    return {"body_excerpt": frame.locator("body").inner_text()[:2400]}


def run_smoke(args):
    report = {"started_at": time.strftime("%Y-%m-%d %H:%M:%S"), "app": args.app, "results": {}}
    console = []
    page_errors = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.headed)
        context_kwargs = {"viewport": {"width": args.width, "height": args.height}}
        if args.storage_state and os.path.exists(args.storage_state):
            context_kwargs["storage_state"] = args.storage_state
        context = browser.new_context(**context_kwargs)
        page = context.new_page()
        page.on("console", lambda msg: console.append({"type": msg.type, "text": msg.text}))
        page.on("pageerror", lambda err: page_errors.append(str(err)))

        if args.app == "quote":
            edit_url = f"{args.quote_base_url}?page=edit"
            dashboard_url = f"{args.quote_base_url}?page=dashboard"
            catalog_url = f"{args.quote_base_url}?page=catalog"

            page.goto(edit_url, wait_until="networkidle", timeout=args.timeout_ms)
            report["results"]["edit"] = quote_edit_smoke(page, args.quote_base_url, args.admin_password)

            page.goto(dashboard_url, wait_until="networkidle", timeout=args.timeout_ms)
            report["results"]["dashboard"] = quote_dashboard_smoke(
                page, args.quote_base_url, args.admin_password, args.quote_query
            )

            page.goto(catalog_url, wait_until="networkidle", timeout=args.timeout_ms)
            report["results"]["catalog"] = quote_catalog_smoke(
                page, args.quote_base_url, args.admin_password, args.catalog_query
            )
        else:
            page.goto(args.prequote_base_url, wait_until="networkidle", timeout=args.timeout_ms)
            report["results"]["public"] = prequote_public_smoke(page, args.prequote_base_url)

            page.goto(f"{args.prequote_base_url}?page=admin", wait_until="networkidle", timeout=args.timeout_ms)
            report["results"]["admin"] = prequote_admin_smoke(
                page, args.prequote_base_url, args.admin_password
            )

        context.close()
        browser.close()

    report["console"] = console[:50]
    report["page_errors"] = page_errors[:20]
    return report


def build_parser():
    parser = argparse.ArgumentParser(description="Run basic URL smoke QA for quote/prequote apps.")
    parser.add_argument("--app", choices=["quote", "prequote"], required=True)
    parser.add_argument("--quote-base-url", default=DEFAULTS["quote_base_url"])
    parser.add_argument("--prequote-base-url", default=DEFAULTS["prequote_base_url"])
    parser.add_argument("--admin-password", default=DEFAULTS["admin_password"])
    parser.add_argument("--quote-query", default="QA 자동화 테스트")
    parser.add_argument("--catalog-query", default="LG")
    parser.add_argument("--storage-state", default="/home/mifasol/interior-app/tools/.auth/google_quote_admin.json")
    parser.add_argument("--width", type=int, default=1440)
    parser.add_argument("--height", type=int, default=1200)
    parser.add_argument("--timeout-ms", type=int, default=120000)
    parser.add_argument("--headed", action="store_true")
    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()
    report = run_smoke(args)
    json.dump(report, sys.stdout, ensure_ascii=False, indent=2)
    sys.stdout.write("\n")


if __name__ == "__main__":
    main()
