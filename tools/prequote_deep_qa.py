#!/usr/bin/env python3
import argparse
import json
import re
import time
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


DEFAULT_BASE_URL = "https://script.google.com/macros/s/AKfycby0_7tpS0-6ieNzsaauo1WVPpaXVFrAq9fpM4mISpd6bKbb_-MoaD1RNSzGXoqUVbN2Iw/exec"
DEFAULT_ADMIN_PASSWORD = "test"


def slugify(value):
    text = re.sub(r"[^0-9A-Za-z._-]+", "-", str(value or "").strip())
    return text.strip("-") or "page"


def screenshot(page, artifact_dir, name):
    path = artifact_dir / f"{slugify(name)}.png"
    page.screenshot(path=str(path), full_page=True)
    return str(path)


def app_frame(page):
    page.wait_for_timeout(2500)
    if "accounts.google.com" in page.url:
        raise RuntimeError("unexpected Google sign-in screen for prequote_app")
    frames = page.frames
    if not frames:
        raise RuntimeError(f"expected app frame, got {len(frames)} frames")
    return frames[-1]


def goto_frame(page, url):
    page.goto(url, wait_until="networkidle", timeout=120000)
    return app_frame(page)


def wait_body_contains(frame, text, timeout_ms=20000):
    frame.get_by_text(text).first.wait_for(timeout=timeout_ms)


def click_option_card(frame, text, timeout_ms=20000):
    card = frame.locator(".opt-card").filter(has_text=text).first
    card.wait_for(timeout=timeout_ms)
    card.click()


def set_contact_answers(frame, payload):
    frame.evaluate(
        """
        (payload) => {
          window.PQ.txt("Q900_NAME", payload.name);
          window.PQ.txt("Q901_PHONE", payload.phone);
          window.PQ.txt("Q902_EMAIL", payload.email);
          window.PQ.txt("Q906_ADDRESS", payload.address);
          window.PQ.txt("Q904_NOTE", payload.note);
          window.PQ.sel("Q903_CONTACT_METHOD", payload.method);
          if (typeof render === "function") render();
        }
        """,
        payload,
    )


def wait_result_frame(page, timeout_ms=30000):
    deadline = time.time() + (timeout_ms / 1000.0)
    while time.time() < deadline:
        for frame in page.frames:
            if "?page=result" in frame.url:
                return frame
        try:
            frame = app_frame(page)
            body_text = frame.locator("body").inner_text()[:800]
            if "가견적 결과" in body_text:
                return frame
        except Exception:
            pass
        page.wait_for_timeout(500)
    return None


def submit_preset_flow(page, base_url, project_label, bundle_label, artifact_dir, tag):
    frame = goto_frame(page, base_url)
    wait_body_contains(frame, "가견적 시작하기")
    landing_excerpt = frame.locator("body").inner_text()[:1200]
    frame.get_by_text("가견적 시작하기").click()
    wait_body_contains(frame, project_label)
    click_option_card(frame, project_label)
    wait_body_contains(frame, "어떤 방식으로 진행할까요?")
    click_option_card(frame, "프리패스 세트")
    frame.page.wait_for_timeout(600)
    if bundle_label not in frame.locator("body").inner_text():
        continue_button = frame.get_by_text("선택 후 계속").last
        if continue_button.count():
            continue_button.click()
    wait_body_contains(frame, bundle_label, timeout_ms=30000)
    click_option_card(frame, bundle_label, timeout_ms=30000)
    wait_body_contains(frame, "원하는 톤만 골라주세요", timeout_ms=30000)
    frame.locator(".opt-card").first.click()
    wait_body_contains(frame, "연락처를 남겨주세요", timeout_ms=30000)

    stamp = time.strftime("%Y%m%d%H%M%S")
    digits = f"{int(time.time() * 1000) % 100000000:08d}"
    contact = {
        "name": f"{tag}-{stamp}",
        "phone": f"010-{digits[:4]}-{digits[4:]}",
        "email": f"{tag.lower()}-{stamp}@example.com",
        "address": f"QA {tag} 테스트 주소",
        "note": "Codex prequote deep QA",
        "method": "KAKAO",
    }
    set_contact_answers(frame, contact)
    frame.get_by_text("결과 확인하기").click()

    result_url = page.url if "?page=result" in page.url else ""
    result_url_frame = None if result_url else wait_result_frame(page, timeout_ms=45000)
    if not result_url and result_url_frame is None:
        submit_frame = app_frame(page)
        readonly_input = submit_frame.locator("input[readonly]").first
        if readonly_input.count():
            fallback_url = readonly_input.input_value().strip()
            if "?page=result" in fallback_url:
                page.goto(fallback_url, wait_until="networkidle", timeout=120000)
                result_url = page.url if "?page=result" in page.url else ""
                result_url_frame = None if result_url else wait_result_frame(page, timeout_ms=15000)
    if not result_url and result_url_frame is None:
        raise RuntimeError(f"result redirect missing for {tag}: {page.url}")

    result_url = result_url or result_url_frame.url
    query = parse_qs(urlparse(result_url).query)
    request_id = (query.get("id") or [""])[0]
    share_token = (query.get("token") or [""])[0]
    result_frame = app_frame(page)
    if not request_id:
        try:
            request_id = result_frame.evaluate("() => new URLSearchParams(window.location.search).get('id') || ''")
        except Exception:
            request_id = ""
    if not share_token:
        try:
            share_token = result_frame.evaluate("() => new URLSearchParams(window.location.search).get('token') || ''")
        except Exception:
            share_token = ""
    wait_body_contains(result_frame, "가견적 결과", timeout_ms=30000)
    result_excerpt = result_frame.locator("body").inner_text()[:2200]
    result_range = result_frame.locator(".hero-range strong").inner_text().strip()

    return {
        "tag": tag,
        "project_label": project_label,
        "bundle_label": bundle_label,
        "contact": contact,
        "request_id": request_id,
        "share_token": share_token,
        "landing_excerpt": landing_excerpt,
        "result_range": result_range,
        "result_excerpt": result_excerpt,
        "screenshots": [
            screenshot(page, artifact_dir, f"{tag}-result"),
        ],
    }


def admin_login(page, base_url, admin_password):
    frame = goto_frame(page, f"{base_url}?page=admin")
    frame.locator("#loginPassword").fill(admin_password)
    frame.locator("#btnLogin").click()
    try:
        frame.get_by_text("가견적 관리자 워크스페이스").wait_for(timeout=20000)
    except PlaywrightTimeoutError:
        page.wait_for_timeout(10000)
    return app_frame(page)


def admin_check_request(frame, request_id, artifact_dir, name):
    frame.locator("#searchInput").fill(request_id)
    frame.page.wait_for_timeout(1000)
    rows = frame.locator("#requestTableBody tr")
    row_count = rows.count()
    row_text = rows.first.inner_text()[:1600] if row_count else ""
    range_text = ""
    detail_excerpt = ""
    if row_count:
        try:
            rows.first.click()
            frame.locator("#detailOverlay, .detail-overlay").first.wait_for(state="visible", timeout=20000)
            frame.page.wait_for_timeout(1200)
            if frame.locator(".detail-money strong").count():
                range_text = frame.locator(".detail-money strong").inner_text().strip()
            detail_excerpt = frame.locator("#detailRoot").inner_text()[:2200]
            close_button = frame.locator("[data-action='detail-close']").first
            if close_button.count():
                close_button.click()
                frame.page.wait_for_timeout(400)
        except Exception:
            pass
    return {
        "request_id": request_id,
        "row_count": row_count,
        "row_text": row_text,
        "detail_range": range_text,
        "detail_excerpt": detail_excerpt,
        "screenshots": [screenshot(frame.page, artifact_dir, name)],
    }


def run(args):
    started_at = time.strftime("%Y-%m-%d %H:%M:%S")
    artifact_dir = Path(args.artifact_dir or f"/home/mifasol/interior-app/tools/.qa_artifacts/prequote_deep_{time.strftime('%Y%m%d_%H%M%S')}")
    artifact_dir.mkdir(parents=True, exist_ok=True)
    report = {
        "started_at": started_at,
        "base_url": args.base_url,
        "artifact_dir": str(artifact_dir),
        "flows": [],
        "admin_checks": [],
        "console": [],
        "page_errors": [],
    }

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.headed)
        context = browser.new_context(viewport={"width": 430, "height": 1100})
        page = context.new_page()
        page.on("console", lambda msg: report["console"].append({"type": msg.type, "text": msg.text}))
        page.on("pageerror", lambda err: report["page_errors"].append(str(err)))

        residential = submit_preset_flow(
            page,
            args.base_url,
            "주거",
            "30평 입주형 밸런스 패키지",
            artifact_dir,
            "RESI_PRESET",
        )
        report["flows"].append(residential)

        commercial = submit_preset_flow(
            page,
            args.base_url,
            "상업",
            "미용실/헤어살롱 오픈 패키지",
            artifact_dir,
            "COMM_PRESET",
        )
        report["flows"].append(commercial)

        frame = admin_login(page, args.base_url, args.admin_password)
        verify_ids = [rid for rid in [residential["request_id"], commercial["request_id"]] if rid]
        for request_id in args.verify_request_id:
            if request_id and request_id not in verify_ids:
                verify_ids.append(request_id)
        for idx, request_id in enumerate(verify_ids, start=1):
            report["admin_checks"].append(
                admin_check_request(frame, request_id, artifact_dir, f"admin-check-{idx}-{request_id}")
            )

        context.close()
        browser.close()

    return report


def build_parser():
    parser = argparse.ArgumentParser(description="Submit preset prequote flows and verify them from admin.")
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL)
    parser.add_argument("--admin-password", default=DEFAULT_ADMIN_PASSWORD)
    parser.add_argument("--artifact-dir", default="")
    parser.add_argument("--verify-request-id", action="append", default=[])
    parser.add_argument("--headed", action="store_true")
    return parser


def main():
    args = build_parser().parse_args()
    report = run(args)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
