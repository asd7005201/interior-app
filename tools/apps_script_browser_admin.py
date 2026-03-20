#!/usr/bin/env python3
import argparse
import json
import os
import sys

from playwright.sync_api import sync_playwright


DEFAULTS = {
    "quote_base_url": "https://script.google.com/macros/s/AKfycbw138HXXYxd6_qnh98eQpc2rkkdt1CSmlqI3hVWZhiDW5Sg8blfeqCwGf68P2Lv_w-y/exec",
    "prequote_base_url": "https://script.google.com/macros/s/AKfycby0_7tpS0-6ieNzsaauo1WVPpaXVFrAq9fpM4mISpd6bKbb_-MoaD1RNSzGXoqUVbN2Iw/exec",
    "storage_state": "/home/mifasol/interior-app/tools/.auth/google_quote_admin.json",
}


def page_url(app, quote_base_url, prequote_base_url):
    if app == "quote":
        return f"{quote_base_url}?page=edit"
    return f"{prequote_base_url}?page=admin"


def app_frame(page):
    page.wait_for_timeout(3000)
    if "accounts.google.com" in page.url:
      raise RuntimeError("Google sign-in required. Capture a storage state first.")
    if len(page.frames) < 3:
      raise RuntimeError(f"expected app frame, got {len(page.frames)} frames")
    return page.frames[2]


def main():
    parser = argparse.ArgumentParser(description="Call google.script.run admin functions through a signed-in Playwright browser session.")
    parser.add_argument("--app", choices=["quote", "prequote"], required=True)
    parser.add_argument("--function", dest="function_name", required=True)
    parser.add_argument("--params-json", default="[]", help="JSON array passed to google.script.run")
    parser.add_argument("--quote-base-url", default=DEFAULTS["quote_base_url"])
    parser.add_argument("--prequote-base-url", default=DEFAULTS["prequote_base_url"])
    parser.add_argument("--storage-state", default=DEFAULTS["storage_state"])
    parser.add_argument("--headed", action="store_true")
    parser.add_argument("--timeout-ms", type=int, default=120000)
    args = parser.parse_args()

    params = json.loads(args.params_json)
    if not isinstance(params, list):
        raise SystemExit("--params-json must be a JSON array")
    if not os.path.exists(args.storage_state):
        raise SystemExit(f"storage state not found: {args.storage_state}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.headed)
        context = browser.new_context(
            storage_state=args.storage_state,
            viewport={"width": 1440, "height": 1200},
        )
        page = context.new_page()
        page.goto(page_url(args.app, args.quote_base_url, args.prequote_base_url), wait_until="networkidle", timeout=args.timeout_ms)
        frame = app_frame(page)
        result = frame.evaluate(
            """
            ({ functionName, params }) => new Promise((resolve) => {
              var runner = google.script.run
                .withSuccessHandler((res) => resolve({ ok: true, result: res }))
                .withFailureHandler((err) => resolve({ ok: false, error: String(err && err.message ? err.message : err) }));
              runner[functionName].apply(runner, params);
            })
            """,
            {"functionName": args.function_name, "params": params},
        )
        print(json.dumps(result, ensure_ascii=False, indent=2))
        context.close()
        browser.close()


if __name__ == "__main__":
    main()
