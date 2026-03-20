#!/usr/bin/env python3
import argparse
import json
import os
import re
import time
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


DEFAULT_BASE_URL = "https://script.google.com/macros/s/AKfycbw138HXXYxd6_qnh98eQpc2rkkdt1CSmlqI3hVWZhiDW5Sg8blfeqCwGf68P2Lv_w-y/exec"
DEFAULT_STORAGE_STATE = "/home/mifasol/interior-app/tools/.auth/google_quote_admin.json"
DEFAULT_ADMIN_PASSWORD = "test"


def app_frame(page):
    page.wait_for_timeout(2500)
    if "accounts.google.com" in page.url:
        raise RuntimeError("Google sign-in required for quote_app admin flow.")
    frames = page.frames
    if len(frames) < 3:
        raise RuntimeError(f"expected app frame, got {len(frames)} frames")
    return frames[2]


def slugify(value):
    text = re.sub(r"[^0-9A-Za-z._-]+", "-", str(value or "").strip())
    return text.strip("-") or "page"


def screenshot(page, artifact_dir, name):
    path = artifact_dir / f"{slugify(name)}.png"
    page.screenshot(path=str(path), full_page=True)
    return str(path)


def goto_frame(page, url):
    page.goto(url, wait_until="networkidle", timeout=120000)
    return app_frame(page)


def fill_admin(frame, password):
    locator = frame.locator("#adminPw")
    locator.wait_for(timeout=20000)
    locator.fill(password)


def wait_for_status(frame, expected_fragment, timeout_ms=20000):
    deadline = time.time() + (timeout_ms / 1000.0)
    last_text = ""
    while time.time() < deadline:
        try:
            last_text = frame.locator("#status").inner_text(timeout=2000).strip()
        except Exception:
            last_text = ""
        if expected_fragment in last_text:
            return last_text
        frame.page.wait_for_timeout(300)
    raise RuntimeError(f"status did not contain '{expected_fragment}', last='{last_text}'")


def text_or_empty(locator):
    try:
        return locator.inner_text().strip()
    except Exception:
        return ""


def quote_edit_deep(page, context, base_url, password, artifact_dir):
    result = {"page": "edit"}
    frame = goto_frame(page, f"{base_url}?page=edit")
    fill_admin(frame, password)
    frame.locator("#btnCreate").click()
    frame.locator("#editor").wait_for(state="visible", timeout=20000)

    stamp = time.strftime("%Y%m%d-%H%M%S")
    customer_name = f"QA-DEEP-{stamp}"
    frame.locator("#customer_name").fill(customer_name)
    frame.locator("#site_name").fill("상세QA 현장")
    frame.locator("#contact_name").fill("Codex QA")
    frame.locator("#contact_phone").fill("010-0000-0000")
    frame.locator("#memo").fill("상세 QA 자동 생성 데이터")

    row_count_before = frame.locator("#itemsBody tr.item-row").count()
    frame.locator("#btnAddRow").click()
    row = frame.locator("#itemsBody tr.item-row").first
    row.locator(".group_label_input").fill("도장 공정")
    row.locator(".group_code").fill("PAINT")
    row.locator(".location").fill("거실 벽면")
    row.locator(".qty").fill("2")
    row.locator(".unit").fill("식")
    row.locator(".unit_price").fill("50000")
    row.locator(".material").fill("QA 테스트 자재")
    row.locator(".detail").fill("샘플 규격")
    row.locator(".note").fill("자동 QA")
    row.locator(".qty").press("Tab")
    frame.page.wait_for_timeout(1500)

    computed_amount = row.locator(".amount").inner_text().strip()

    frame.locator("#btnSave").click()
    save_status = wait_for_status(frame, "저장")
    share_button_visible = frame.locator("#btnShare").is_visible()
    share_status = ""
    view_url = ""
    print_url = ""
    links_text = text_or_empty(frame.locator("#links"))
    signed_excerpt = ""
    signed_path = ""
    anon_result = {}
    anon_path = ""

    if share_button_visible:
        frame.locator("#btnShare").click()
        share_status = wait_for_status(frame, "공유 링크")
        frame.locator("#links a").first.wait_for(timeout=20000)
        view_url = frame.locator("#links a").first.get_attribute("href") or ""
        print_url = frame.locator("#links a").nth(1).get_attribute("href") or ""
        links_text = text_or_empty(frame.locator("#links"))

        signed_page = context.new_page()
        signed_page.goto(view_url, wait_until="load", timeout=120000)
        signed_page.wait_for_timeout(5000)
        signed_excerpt = app_frame(signed_page).locator("body").inner_text()[:800]
        signed_path = screenshot(signed_page, artifact_dir, "share-view-signed")
        signed_page.close()

        anon_context = context.browser.new_context(viewport={"width": 1440, "height": 1200})
        anon_page = anon_context.new_page()
        anon_page.goto(view_url, wait_until="load", timeout=120000)
        anon_page.wait_for_timeout(2500)
        anon_result = {
            "final_url": anon_page.url,
            "requires_google_login": "accounts.google.com" in anon_page.url,
            "body_excerpt": anon_page.locator("body").inner_text()[:400],
        }
        anon_path = screenshot(anon_page, artifact_dir, "share-view-anon")
        anon_page.close()
        anon_context.close()

    result.update(
        {
            "customer_name": customer_name,
            "row_count_before_add": row_count_before,
            "row_count_after_add": frame.locator("#itemsBody tr.item-row").count(),
            "computed_amount": computed_amount,
            "save_status": save_status,
            "share_button_visible": share_button_visible,
            "share_status": share_status,
            "links_text": links_text,
            "view_url": view_url,
            "print_url": print_url,
            "signed_view_excerpt": signed_excerpt,
            "anonymous_share_check": anon_result,
            "screenshots": [
                screenshot(page, artifact_dir, "edit"),
            ],
        }
    )
    if signed_path:
        result["screenshots"].append(signed_path)
    if anon_path:
        result["screenshots"].append(anon_path)
    return result


def quote_dashboard_deep(page, base_url, password, query, artifact_dir):
    result = {"page": "dashboard"}
    frame = goto_frame(page, f"{base_url}?page=dashboard")
    fill_admin(frame, password)
    frame.locator("#q").fill(query)
    frame.locator("#btnSearch").click()
    frame.page.wait_for_timeout(5000)
    rows = frame.locator("#rows .quote-dashboard-row")
    first_row = rows.first
    result.update(
        {
            "status": text_or_empty(frame.locator("#status")),
            "row_count": rows.count(),
            "summary": {
                "month_total": text_or_empty(frame.locator("#sumMonth")),
                "pending": text_or_empty(frame.locator("#sumPending")),
                "conversion": text_or_empty(frame.locator("#statConversion")),
            },
            "first_row_excerpt": first_row.inner_text()[:700] if rows.count() else "",
            "screenshots": [screenshot(page, artifact_dir, "dashboard")],
        }
    )
    return result


def quote_catalog_deep(page, base_url, password, query, artifact_dir):
    result = {"page": "catalog"}
    frame = goto_frame(page, f"{base_url}?page=catalog")
    fill_admin(frame, password)
    frame.locator("#matQuery").fill(query)
    frame.locator("#btnSearch").click()
    frame.page.wait_for_timeout(7000)
    first_card_text = text_or_empty(frame.locator("#catalogBody").locator("tr").first)
    result.update(
        {
            "status": text_or_empty(frame.locator("#status")),
            "result_meta": text_or_empty(frame.locator("#resultMeta")),
            "first_result_text_length": len(first_card_text),
            "first_result_excerpt": first_card_text[:1000],
            "screenshots": [screenshot(page, artifact_dir, "catalog")],
        }
    )
    return result


def quote_templates_deep(page, base_url, password, artifact_dir):
    result = {"page": "templates"}
    frame = goto_frame(page, f"{base_url}?page=templates")
    fill_admin(frame, password)
    frame.locator("#btnSearch").click()
    frame.page.wait_for_timeout(6000)
    rows = frame.locator("#tplRows tr")
    row_count = rows.count()
    first_row_text = rows.first.inner_text()[:900] if row_count else ""

    loaded_name = ""
    loaded_version_count = 0
    first_open_button_count = rows.first.locator("button:has-text('열기')").count() if row_count else 0
    if row_count and first_open_button_count:
        rows.first.locator("button:has-text('열기')").first.click()
        frame.page.wait_for_timeout(5000)
        loaded_name = frame.locator("#tName").input_value().strip()
        loaded_version_count = frame.locator("#tVersion option").count()

    result.update(
        {
            "status": text_or_empty(frame.locator("#status")),
            "row_count": row_count,
            "first_row_excerpt": first_row_text,
            "loaded_template_name": loaded_name,
            "loaded_version_count": loaded_version_count,
            "screenshots": [screenshot(page, artifact_dir, "templates")],
        }
    )
    return result


def quote_templateslist_deep(page, base_url, password, artifact_dir):
    result = {"page": "templateslist"}
    frame = goto_frame(page, f"{base_url}?page=templateslist")
    fill_admin(frame, password)
    frame.locator("#btnSearch").click()
    frame.page.wait_for_timeout(6000)
    rows = frame.locator("#rows tr")
    row_count = rows.count()
    first_row_text = rows.first.inner_text()[:900] if row_count else ""

    detail_meta = ""
    version_count = 0
    first_version_button_count = rows.first.locator("button:has-text('버전')").count() if row_count else 0
    if row_count and first_version_button_count:
        rows.first.locator("button:has-text('버전')").click()
        frame.locator("#detailModal").wait_for(state="visible", timeout=20000)
        frame.page.wait_for_timeout(3000)
        detail_meta = text_or_empty(frame.locator("#detailMeta"))
        version_count = frame.locator("#versionRows tr").count()

    result.update(
        {
            "status": text_or_empty(frame.locator("#status")),
            "list_meta": text_or_empty(frame.locator("#listMeta")),
            "row_count": row_count,
            "first_row_excerpt": first_row_text,
            "detail_meta": detail_meta,
            "version_row_count": version_count,
            "screenshots": [screenshot(page, artifact_dir, "templateslist")],
        }
    )
    return result


def quote_materialgroups_deep(page, base_url, password, query, artifact_dir):
    result = {"page": "materialgroups"}
    frame = goto_frame(page, f"{base_url}?page=materialgroups")
    fill_admin(frame, password)
    frame.locator("#btnReload").click()
    frame.page.wait_for_timeout(6000)
    rows = frame.locator("#rows tr")
    initial_count = rows.count()
    count_meta = text_or_empty(frame.locator("#countMeta"))
    first_row_text = rows.first.inner_text()[:900] if initial_count else ""

    frame.locator("#btnNew").click()
    frame.page.wait_for_timeout(800)
    after_new_count = frame.locator("#rows tr").count()

    picker_result_meta = ""
    picker_result_count = 0
    if after_new_count:
        frame.locator("#rows tr").first.locator("button:has-text('자재 선택')").click()
        frame.locator("#pickerModal").wait_for(state="visible", timeout=15000)
        frame.locator("#pickerQuery").fill(query)
        frame.locator("#btnPickerSearch").click()
        frame.page.wait_for_timeout(4000)
        picker_result_meta = text_or_empty(frame.locator("#pickerResultMeta"))
        picker_result_count = frame.locator("#pickerResults .list-row").count()

    result.update(
        {
            "status": text_or_empty(frame.locator("#status")),
            "count_meta": count_meta,
            "initial_row_count": initial_count,
            "after_new_row_count": after_new_count,
            "first_row_excerpt": first_row_text,
            "picker_result_meta": picker_result_meta,
            "picker_result_count": picker_result_count,
            "screenshots": [screenshot(page, artifact_dir, "materialgroups")],
        }
    )
    return result


def main():
    parser = argparse.ArgumentParser(description="Run deeper private QA against quote_app.")
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL)
    parser.add_argument("--storage-state", default=DEFAULT_STORAGE_STATE)
    parser.add_argument("--admin-password", default=DEFAULT_ADMIN_PASSWORD)
    parser.add_argument("--catalog-query", default="LG")
    parser.add_argument("--picker-query", default="LG")
    parser.add_argument("--artifact-dir", default="")
    parser.add_argument("--headed", action="store_true")
    args = parser.parse_args()

    artifact_dir = Path(args.artifact_dir or f"/home/mifasol/interior-app/tools/.qa_artifacts/quote_deep_{time.strftime('%Y%m%d_%H%M%S')}")
    artifact_dir.mkdir(parents=True, exist_ok=True)

    report = {
        "started_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "base_url": args.base_url,
        "artifact_dir": str(artifact_dir),
    }
    console = []
    page_errors = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.headed)
        context_kwargs = {"viewport": {"width": 1440, "height": 1200}}
        if args.storage_state and os.path.exists(args.storage_state):
            context_kwargs["storage_state"] = args.storage_state
        context = browser.new_context(**context_kwargs)
        page = context.new_page()
        page.on("console", lambda msg: console.append({"type": msg.type, "text": msg.text}))
        page.on("pageerror", lambda err: page_errors.append(str(err)))

        edit_result = quote_edit_deep(page, context, args.base_url, args.admin_password, artifact_dir)
        dashboard_result = quote_dashboard_deep(
            page,
            args.base_url,
            args.admin_password,
            edit_result["customer_name"],
            artifact_dir,
        )
        catalog_result = quote_catalog_deep(page, args.base_url, args.admin_password, args.catalog_query, artifact_dir)
        templates_result = quote_templates_deep(page, args.base_url, args.admin_password, artifact_dir)
        templateslist_result = quote_templateslist_deep(page, args.base_url, args.admin_password, artifact_dir)
        materialgroups_result = quote_materialgroups_deep(
            page,
            args.base_url,
            args.admin_password,
            args.picker_query,
            artifact_dir,
        )

        context.close()
        browser.close()

    report["results"] = {
        "edit": edit_result,
        "dashboard": dashboard_result,
        "catalog": catalog_result,
        "templates": templates_result,
        "templateslist": templateslist_result,
        "materialgroups": materialgroups_result,
    }
    report["console"] = console[:80]
    report["page_errors"] = page_errors[:30]

    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
