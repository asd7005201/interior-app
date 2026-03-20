#!/usr/bin/env python3
import argparse
import os
import time

from playwright.sync_api import sync_playwright


DEFAULTS = {
    "quote_base_url": "https://script.google.com/macros/s/AKfycbw138HXXYxd6_qnh98eQpc2rkkdt1CSmlqI3hVWZhiDW5Sg8blfeqCwGf68P2Lv_w-y/exec",
    "prequote_base_url": "https://script.google.com/macros/s/AKfycby0_7tpS0-6ieNzsaauo1WVPpaXVFrAq9fpM4mISpd6bKbb_-MoaD1RNSzGXoqUVbN2Iw/exec",
}


def target_url(args):
    if args.app == "quote":
        return f"{args.quote_base_url}?page=edit"
    return f"{args.prequote_base_url}?page=admin"


def wait_until_authenticated(page, timeout_ms):
    started = time.time()
    while (time.time() - started) * 1000 < timeout_ms:
        page.wait_for_timeout(1000)
        if "accounts.google.com" in page.url:
            continue
        if "script.google.com" in page.url and len(page.frames) >= 3:
            return True
    return False


def main():
    parser = argparse.ArgumentParser(description="Open Google sign-in once and save Playwright storage state for private Apps Script QA.")
    parser.add_argument("--app", choices=["quote", "prequote"], required=True)
    parser.add_argument("--quote-base-url", default=DEFAULTS["quote_base_url"])
    parser.add_argument("--prequote-base-url", default=DEFAULTS["prequote_base_url"])
    parser.add_argument("--storage-state", default="/home/mifasol/interior-app/tools/.auth/google_quote_admin.json")
    parser.add_argument("--timeout-ms", type=int, default=300000)
    args = parser.parse_args()

    os.makedirs(os.path.dirname(args.storage_state), exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(viewport={"width": 1440, "height": 1200})
        page = context.new_page()
        page.goto(target_url(args), wait_until="load", timeout=args.timeout_ms)
        print("브라우저가 열렸습니다. Google 로그인과 접근 허용을 완료하면 세션을 저장합니다.")
        print("대상 URL:", target_url(args))

        ok = wait_until_authenticated(page, args.timeout_ms)
        if not ok:
            raise SystemExit("로그인 완료를 확인하지 못했습니다. 다시 실행해 주세요.")

        context.storage_state(path=args.storage_state)
        print("세션 저장 완료:", args.storage_state)
        context.close()
        browser.close()


if __name__ == "__main__":
    main()
