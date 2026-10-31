"""Immediately runnable phone UI with local workbook persistence and real OCR.

No SharePoint installation, OAuth registration or cloud server is used. The
production page, file chooser, IndexedDB, recognizer and workbook writer run
unchanged. A private workbook/photo are supplied at runtime and never added to
Git. Phone layouts are Chromium emulation, not physical devices or Safari.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
import json
from pathlib import Path
import re
import tempfile
import time
from urllib.parse import urlsplit

from playwright.sync_api import sync_playwright

from browser_test import serve_program, download_expression
from connected_file_test import PICKER_SCRIPT, seed_file
from mobile_app_test import DEVICES, SHARE_API, expected_reference
from multi_owner_ocr_test import verify_increment


def arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--program", type=Path, default=Path("/workspace/outputs/Tool-Manager-Fixed.html"))
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--photo", type=Path, required=True)
    parser.add_argument("--serial", required=True)
    parser.add_argument("--last")
    parser.add_argument("--due")
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    parser.add_argument("--diagnostics", type=Path, default=Path("/tmp/tool-direct-runner-test.json"))
    args = parser.parse_args()
    if args.workbook is None:
        args.workbook = next(Path("/workspace/attachments").rglob("*.xlsx"), None)
    for label, file in [("program", args.program), ("workbook", args.workbook), ("photo", args.photo)]:
        if file is None or not file.is_file():
            parser.error("Supply --" + label + " with an existing file")
    return args


def ready(page):
    page.wait_for_function("window.ToolApp&&window.ToolSources&&!ToolApp.state.busy&&!ToolApp.state.restoring", timeout=60000)


def assert_direct_page(page):
    assert page.evaluate("ToolSources.mode()") == "device"
    assert page.locator("#quick-start").is_visible()
    assert page.locator("#quick-workbook").is_visible()
    assert page.locator("#m365-panel").is_hidden()
    assert page.locator("#sharepoint-panel").is_hidden()
    assert not page.locator("#connection-settings").evaluate("el=>el.open")
    measurements = page.evaluate("""()=>{
      const box=document.getElementById('quick-workbook').getBoundingClientRect();
      return {viewport:innerWidth,scrollWidth:document.documentElement.scrollWidth,
              top:box.top,bottom:box.bottom,height:box.height,viewHeight:innerHeight};
    }""")
    assert measurements["scrollWidth"] <= measurements["viewport"], measurements
    assert 0 <= measurements["top"] < measurements["viewHeight"], measurements
    assert measurements["height"] >= 44, measurements
    return measurements


def state_result(page):
    return page.evaluate("""async()=>{
      const photo=ToolApp.state.photos.at(-1),book=await ToolMobileFile.read();
      const hash=async bytes=>Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))).map(n=>n.toString(16).padStart(2,'0')).join('');
      return {saved:!!photo?.saved,approval:photo?.approval,sourceMode:ToolSources.mode(),
        saveCount:ToolApp.state.source.saveCount,sourceError:ToolApp.state.source.error,
        fields:['name','serial','last','due'].map(id=>document.getElementById(id).value),
        photoHash:await hash(photo.image.photoBytes),workbookHash:await hash(book.bytes),
        version:book.version,status:document.getElementById('status').textContent,
        note:document.getElementById('source-save-note').textContent,
        quick:document.getElementById('quick-summary').textContent};
    }""")


def main():
    args = arguments()
    direct_entry = bool(re.search(r'<meta\b[^>]*\bname=["\']tool-manager-direct-entry["\']', args.program.read_text(), re.I))
    reference, expected = expected_reference(args)
    workbook_hash, photo_hash = sha256(args.workbook.read_bytes()).hexdigest(), sha256(args.photo.read_bytes()).hexdigest()
    cases, errors, unexpected, company_requests, published_requests = {}, [], [], [], []
    report = {"program_sha256": sha256(args.program.read_bytes()).hexdigest(),
              "ocr_responses_replaced": False, "expected_values_passed_to_recognizer": False,
              "physical_phone_or_safari_engine_tested": False, "company_cloud_original_changed": False}
    report["direct_entry"] = direct_entry
    with tempfile.TemporaryDirectory(prefix="tool-direct-runner-") as directory, serve_program(args.program) as url, sync_playwright() as playwright:
        temporary = Path(directory)
        host = urlsplit(url).netloc
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])

        def fresh(device, saved_mode=None, configured=False, native_picker=False):
            settings = DEVICES[device] if device in DEVICES else {"viewport": {"width": 1280, "height": 900}}
            context = browser.new_context(**settings, is_mobile=device != "desktop", has_touch=device != "desktop",
                                          locale="ko-KR", accept_downloads=True)
            context.add_init_script("(" + SHARE_API + ")()")
            if native_picker:
                context.add_init_script("(" + PICKER_SCRIPT + ")()")
            if saved_mode:
                context.add_init_script("""(()=>{
                  if(!/^https?:$/.test(location.protocol))return;
                  localStorage.setItem('tool-manager-workbook-source',%s);
                })();""" % json.dumps(saved_mode))
            if configured:
                context.add_init_script("""(()=>{
                  if(!/^https?:$/.test(location.protocol))return;
                  localStorage.setItem('tool-manager-m365-config-v1',JSON.stringify({
                    clientId:'00000000-0000-4000-8000-000000000001',
                    tenantId:'11111111-1111-4111-8111-111111111111',
                    redirectUri:location.origin+location.pathname}));
                })();""")
            page = context.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))

            def route(request_route):
                request = request_route.request
                parsed = urlsplit(request.url)
                if parsed.netloc == host and parsed.path.endswith("/team-settings.json"):
                    published_requests.append(request.url)
                    return request_route.fulfill(status=200, content_type="application/json", body='{"enabled":false}', headers={"cache-control": "no-store"})
                if parsed.scheme in {"http", "https"} and parsed.netloc != host:
                    unexpected.append(request.url)
                    if parsed.hostname in {"graph.microsoft.com", "login.microsoftonline.com"} or "sharepoint" in (parsed.hostname or ""):
                        company_requests.append({"method": request.method, "url": request.url})
                    return request_route.abort()
                return request_route.continue_()

            context.route("**/*", route)
            page.goto(url)
            ready(page)
            return context, page

        for device in ["iphone", "android"]:
            context, page = fresh(device)
            cases[device + "_opens_directly_without_install_or_ids"] = assert_direct_page(page)
            assert not page.evaluate("ToolApp.state.source.connected")
            with page.expect_file_chooser() as event:
                page.locator("#quick-workbook").click()
            event.value.set_files(args.workbook)
            page.wait_for_function("ToolApp.state.source.connected&&ToolApp.state.model&&!ToolApp.state.busy", timeout=60000)
            assert page.evaluate("ToolSources.mode()") == "device"
            assert page.evaluate("ToolMobileFile.read().then(value=>value.file.name)") == args.workbook.name
            assert page.locator("#take-photo").is_enabled() and page.locator("#choose-photos").is_enabled()
            assert page.locator("#quick-clear-workbook").is_visible()
            assert page.locator("#m365-panel").is_hidden() and page.locator("#sharepoint-panel").is_hidden()
            cases[device + "_one_file_picker_import"] = True

            if device == "iphone":
                with page.expect_file_chooser() as event:
                    page.locator("#take-photo").click()
                event.value.set_files([])
                assert page.locator("#camera").get_attribute("capture") == "environment"
                page.reload()
                ready(page)
                assert page.evaluate("ToolApp.state.source.connected&&!!ToolApp.state.model")
                assert page.locator("#take-photo").is_enabled()
                cases["iphone_camera_chooser_and_remembered_workbook"] = True
                before_clear = page.evaluate("ToolMobileFile.read().then(value=>value.version)")
                page.once("dialog", lambda dialog: dialog.dismiss())
                page.locator("#quick-clear-workbook").click()
                assert page.evaluate("ToolApp.state.source.connected")
                assert page.evaluate("ToolMobileFile.read().then(value=>value.version)") == before_clear
                cases["cancel_clear_preserves_device_workbook"] = True
            else:
                before = time.monotonic()
                with page.expect_file_chooser() as event:
                    page.locator("#choose-photos").click()
                event.value.set_files({"name": "tool-camera-photo.jpg", "mimeType": "image/jpeg", "buffer": args.photo.read_bytes()})
                page.wait_for_function("ToolApp.state.photos.length===1&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=180000)
                result = state_result(page)
                assert result["saved"] and result["approval"] == "automatic" and result["saveCount"] == 1 and not result["sourceError"], result
                assert result["fields"] == [expected[key] for key in ["name", "serial", "last", "due"]], result
                assert result["sourceMode"] == "device" and result["workbookHash"] != workbook_hash
                assert "기기" in result["status"] and "회사 공유 원본은 변경하지 않았습니다" in result["status"], result
                updated = temporary / "direct-phone-updated.xlsx"
                download_expression(page, "(await ToolMobileFile.read()).bytes", updated)
                result["verification"] = verify_increment(args.workbook, updated, reference, expected, result["photoHash"])
                backup = temporary / "direct-phone-backup.xlsx"
                download_expression(page, "(await ToolMobileFile.takeBackup()).bytes", backup)
                assert backup.read_bytes() == args.workbook.read_bytes()
                result["original_backup_byte_exact"] = True
                result["seconds"] = round(time.monotonic() - before, 2)
                cases["actual_photo_to_local_calibration_dates_and_photo"] = result
                assert page.locator("#export").is_enabled()
                with page.expect_download() as event:
                    page.locator("#export").click()
                exported = temporary / "direct-phone-exported.xlsx"
                event.value.save_as(exported)
                assert sha256(exported.read_bytes()).hexdigest() == result["workbookHash"]
                cases["user_downloads_exact_updated_workbook"] = True
                page.reload()
                ready(page)
                restored = page.evaluate("""async()=>{
                  const book=await ToolMobileFile.read();
                  return {connected:ToolApp.state.source.connected,model:!!ToolApp.state.model,
                    version:book.version,hash:Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',book.bytes))).map(n=>n.toString(16).padStart(2,'0')).join('')};
                }""")
                assert restored["connected"] and restored["model"] and restored["version"] == result["version"] and restored["hash"] == result["workbookHash"], restored
                assert page.locator("#take-photo").is_enabled()
                cases["updated_local_workbook_restores_without_file_picker"] = restored
                assert page.locator("#quick-download-workbook").is_visible()
                with page.expect_download() as event:
                    page.locator("#quick-download-workbook").click()
                recovered = temporary / "direct-phone-restored-download.xlsx"
                event.value.save_as(recovered)
                assert sha256(recovered.read_bytes()).hexdigest() == result["workbookHash"]
                cases["restored_workbook_downloads_without_another_photo"] = True
            page.once("dialog", lambda dialog: dialog.accept())
            page.locator("#quick-clear-workbook").click()
            page.wait_for_function("!ToolApp.state.source.connected&&!ToolApp.state.model&&!ToolApp.state.busy", timeout=30000)
            assert not page.evaluate("ToolApp.state.source.connected||!!ToolApp.state.model")
            assert page.evaluate("ToolMobileFile.restore().then(value=>value.connected)") is False
            assert page.locator("#quick-workbook").is_visible()
            cases[device + "_clears_only_device_copy"] = True
            if device == "iphone":
                with page.expect_file_chooser() as event:
                    page.locator("#quick-workbook").click()
                assert page.locator("#device-workbook").input_value() == ""
                event.value.set_files(args.workbook)
                page.wait_for_function("ToolApp.state.source.connected&&ToolApp.state.model&&!ToolApp.state.busy")
                assert page.evaluate("ToolMobileFile.read().then(value=>value.file.name)") == args.workbook.name
                cases["same_file_can_be_selected_again_after_clear"] = True
            context.close()

        for previous in ["sharepoint", "microsoft365"]:
            context, page = fresh("iphone", saved_mode=previous)
            assert_direct_page(page)
            assert not page.evaluate("ToolM365Auth.inspect().configured")
            if direct_entry:
                assert page.evaluate("localStorage.getItem('tool-manager-workbook-source')") == previous
                assert page.evaluate("ToolSources.inspect().directEntry") is True
                cases["direct_entry_ignores_and_preserves_old_" + previous + "_choice"] = True
            else:
                cases["unconfigured_" + previous + "_migrates_to_direct_runner"] = True
            context.close()

        context, page = fresh("iphone", saved_mode="microsoft365", configured=True)
        assert page.evaluate("ToolM365Auth.inspect().configured")
        assert not page.evaluate("ToolApp.state.source.connected")
        if direct_entry:
            assert_direct_page(page)
            assert page.evaluate("localStorage.getItem('tool-manager-workbook-source')") == "microsoft365"
            configuration = page.evaluate("JSON.parse(localStorage.getItem('tool-manager-m365-config-v1'))")
            assert configuration["clientId"] == "00000000-0000-4000-8000-000000000001"
            assert configuration["tenantId"] == "11111111-1111-4111-8111-111111111111"
            assert configuration["redirectUri"] == url
            assert page.evaluate("""()=>{
              try{ToolSources.use('microsoft365');return false;}catch(_){return ToolSources.mode()==='device';}
            }""") is True
            # Native option state is inspected directly: this Playwright build
            # reports is_disabled=False for hidden options in a closed details.
            assert page.locator("#source-mode option[value='microsoft365']").evaluate("option=>option.disabled&&option.hidden")
            cases["direct_entry_preserves_root_cloud_settings_without_entering_cloud_mode"] = True
        else:
            assert page.evaluate("ToolSources.mode()") == "microsoft365"
            assert page.locator("#m365-panel").is_visible()
            cases["existing_configured_cloud_source_is_preserved"] = True
        context.close()

        context, page = fresh("desktop", native_picker=True)
        assert page.evaluate("ToolSources.mode()") == "desktop"
        assert page.locator("#quick-start").is_visible() and page.locator("#quick-workbook").is_visible()
        assert seed_file(page, args.workbook) == workbook_hash
        page.locator("#quick-workbook").click()
        page.wait_for_function("ToolApp.state.source.connected&&ToolApp.state.model&&!ToolApp.state.busy", timeout=60000)
        assert page.evaluate("testPickerCalls") == 1
        assert page.evaluate("ToolSources.mode()") == "desktop"
        assert page.evaluate("ToolConnectedFile.read().then(book=>testHash(book.bytes))") == workbook_hash
        assert not page.evaluate("ToolMobileFile.inspect().connected")
        cases["desktop_quick_button_connects_real_file_handle"] = True
        context.close()
        browser.close()

    assert not errors and not unexpected and not company_requests, (errors, unexpected, company_requests)
    if direct_entry:
        assert not published_requests, published_requests
        cases["direct_entry_never_requests_published_cloud_configuration"] = True
    assert sha256(args.workbook.read_bytes()).hexdigest() == workbook_hash
    assert sha256(args.photo.read_bytes()).hexdigest() == photo_hash
    report.update({"cases": cases, "case_count": len(cases), "page_errors": errors,
                   "unexpected_requests": unexpected, "company_requests": company_requests,
                   "published_config_requests": published_requests,
                   "original_photo_and_workbook_unchanged": True})
    args.diagnostics.write_text(json.dumps(report, ensure_ascii=False, indent=2))
    print(json.dumps({"program_sha256": report["program_sha256"], "case_count": len(cases),
                      "cases": list(cases), "page_errors": errors, "company_requests": company_requests,
                      "diagnostics": str(args.diagnostics)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
