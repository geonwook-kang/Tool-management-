"""PC release UI, remembered handle permissions and clean self-downloads.

The production entry, source adapter, IndexedDB, OPFS and OOXML writer run in
Chromium. Only the native picker/permission dialogs and OCR text are controlled;
this checks PC workflow and persistence, not OCR accuracy or OneDrive sync.
All workbook/photo fixtures are synthetic and created in a temporary folder.
"""
from __future__ import annotations

import argparse
from datetime import datetime
from hashlib import sha256
import json
from pathlib import Path
import tempfile
from urllib.parse import urlsplit
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from PIL import Image
from playwright.sync_api import sync_playwright

from browser_test import REPO, serve_program
from connected_file_test import PICKER_SCRIPT, export_shared, seed_file
from download_test import save_download


PRIVATE_MARKER = "PC_UPLOAD_ONLY_739581"
SERIAL = "PC739581246"
PERMISSIONS = """(()=>{
  if(!/^https?:$/.test(location.protocol)||typeof FileSystemFileHandle==='undefined')return;
  const prototype=FileSystemFileHandle.prototype;
  prototype.queryPermission=async function(){return sessionStorage.getItem('pc-test-permission')||'granted';};
  prototype.requestPermission=async function(){
    const calls=JSON.parse(sessionStorage.getItem('pc-test-permission-calls')||'[]');
    calls.push({name:this.name,active:navigator.userActivation.isActive});
    sessionStorage.setItem('pc-test-permission-calls',JSON.stringify(calls));
    sessionStorage.setItem('pc-test-permission','granted');return 'granted';
  };
})();"""
OLD_PREFERENCES = """(()=>{
 if(!/^https?:$/.test(location.protocol))return;
 localStorage.setItem('tool-manager-workbook-source','microsoft365');
 localStorage.setItem('tool-manager-direct-source','device');
 localStorage.setItem('tool-manager-shared-source','shared');
 localStorage.setItem('tool-manager-pc-source-v1','microsoft365');
 localStorage.setItem('tool-manager-m365-config-v1',JSON.stringify({
   clientId:'00000000-0000-4000-8000-000000000001',tenantId:'11111111-1111-4111-8111-111111111111',
   redirectUri:location.origin+location.pathname}));
})();"""


def ready(page):
    page.wait_for_function("window.ToolApp&&window.ToolSources&&!ToolApp.state.busy&&!ToolApp.state.restoring&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=60000)


def fixture(workbook, photo):
    book = Workbook()
    sheet = book.active
    sheet.title = "PCFixture"
    sheet.append(["Cal Next", "Cal Last", "Serial", "Asset #", "Cal Interval", "Cal Cycle", "Label Photo", "SMAX Tool Photo"])
    sheet.append([datetime(2027, 1, 5), datetime(2026, 1, 5), SERIAL, "Fixture_홍길동", 12, "MONTHS"])
    sheet["A2"].number_format = sheet["B2"].number_format = "yyyy-mm-dd"
    sheet["A2"].fill = PatternFill(fgColor="DDEEFF", fill_type="solid")
    sheet["A4"] = PRIVATE_MARKER
    book.save(workbook)
    book.close()
    Image.new("RGB", (180, 120), (71, 150, 190)).save(photo)


def controlled_ocr(page):
    page.evaluate("""()=>{
      const parse=ToolOCR.parse;
      ToolOCR.recognize=async(_file,tools)=>{
        const text='Serial: PC739581246\\nDate of Calibration 2027.01.05\\nDue date of next Calibration 2028.01.05';
        return parse(text,tools,{passes:[{text,source:'original',confidence:96,words:[]}]});
      };
    }""")


def assert_pc(page):
    value = page.evaluate("({mode:ToolSources.mode(),pc:ToolSources.inspect().pcEntry,direct:ToolSources.inspect().directEntry,photos:ToolApp.state.photos.length})")
    assert value == {"mode": "desktop", "pc": True, "direct": True, "photos": 0}, value
    assert page.locator('meta[name="tool-manager-pc-entry"]').count() == 1
    assert page.locator("#quick-workbook").is_visible()
    for selector in ["#shared-panel", "#m365-panel", "#m365-setup-fields", "#sharepoint-panel", "#take-photo"]:
        assert page.locator(selector).is_hidden(), selector
    for value in ["microsoft365", "sharepoint", "shared"]:
        assert page.locator('#source-mode option[value="' + value + '"]').evaluate("element=>element.hidden&&element.disabled"), value
    description = page.locator("h1+p").inner_text()
    assert "시리얼·부분 번호" in description and "엑셀 행" in description


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--program", type=Path, default=Path("/workspace/outputs/Tool-Manager-PC-Final.html"))
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    args = parser.parse_args()
    assert args.program.is_file(), args.program
    cases, errors, external, published = {}, [], [], []
    report = {"program_sha256": sha256(args.program.read_bytes()).hexdigest(), "cases": cases,
              "ocr_responses_controlled": True, "company_cloud_sync_tested": False}
    with tempfile.TemporaryDirectory(prefix="tool-pc-release-") as folder, serve_program(args.program) as url, sync_playwright() as playwright:
        temporary = Path(folder)
        workbook, photo = temporary / "synthetic-PC.xlsx", temporary / (PRIVATE_MARKER + ".png")
        fixture(workbook, photo)
        baseline = sha256(workbook.read_bytes()).hexdigest()
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])

        def context(native=True):
            result = browser.new_context(viewport={"width": 1280, "height": 950}, locale="ko-KR", accept_downloads=True)
            result.add_init_script(OLD_PREFERENCES)
            if native:
                result.add_init_script("if(/^https?:$/.test(location.protocol)){(" + PICKER_SCRIPT + ")();}")
                result.add_init_script(PERMISSIONS)
            else:
                result.add_init_script("Object.defineProperty(window,'showOpenFilePicker',{value:undefined,configurable:true});")
            # Exercise the ordinary anchor download fallback, independently of
            # the browser-native save picker that other storage tests cover.
            result.add_init_script("Object.defineProperty(window,'showSaveFilePicker',{value:undefined,configurable:true});")
            def network(route):
                address = urlsplit(route.request.url)
                if address.path.endswith("team-settings.json"):
                    published.append(route.request.url)
                    return route.abort()
                if address.scheme in {"http", "https"} and address.netloc != urlsplit(url).netloc:
                    external.append(route.request.url)
                    return route.abort()
                return route.continue_()
            result.route("**/*", network)
            result.on("page", lambda page: page.on("pageerror", lambda error: errors.append(str(error))))
            return result

        native = context()
        page = native.new_page()
        page.goto(url)
        ready(page)
        assert_pc(page)
        cases["dedicated_pc_defaults_to_desktop_ignoring_old_cloud_and_device_settings"] = True
        pristine = page.evaluate("ToolDistribution.html()")
        assert seed_file(page, workbook) == baseline
        page.locator("#quick-workbook").click()
        page.wait_for_function("ToolApp.state.source.connected&&ToolApp.state.model&&!ToolApp.state.busy", timeout=60000)
        assert page.evaluate("testPickerCalls") == 1
        assert page.evaluate("ToolSources.mode()") == "desktop"
        controlled_ocr(page)
        page.locator("#photos").set_input_files(photo)
        page.wait_for_function("ToolApp.state.photos[0]?.saved&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=60000)
        assert page.evaluate("ToolApp.state.source.saveCount") == 1
        output = temporary / "updated.xlsx"
        export_shared(page, output)
        book = load_workbook(output)
        sheet = book["PCFixture"]
        assert sheet["A2"].value.strftime("%Y-%m-%d") == "2028-01-05"
        assert sheet["B2"].value.strftime("%Y-%m-%d") == "2027-01-05"
        assert sheet["C2"].value == SERIAL and sheet["A4"].value == PRIVATE_MARKER
        assert len(sheet._images) == 1
        book.close()
        cases["first_quick_connection_saves_into_actual_selected_opfs_workbook"] = True

        html, archive = temporary / "self.html", temporary / "self.zip"
        save_download(page, page.locator("#save-program-html"), html, "Tool-Manager.html")
        save_download(page, page.locator("#save-program-zip"), archive, "Tool-Manager.zip")
        assert html.read_text() == pristine and PRIVATE_MARKER not in html.read_text()
        assert 'name="tool-manager-pc-entry"' in pristine
        with ZipFile(archive) as bundle:
            assert bundle.testzip() is None
            assert bundle.read("Tool-Manager.html").decode() == pristine
            assert bundle.read("사용법.txt").decode() == (REPO / "web/PC-사용법.txt").read_text()
            assert {"licenses/JSZIP-LICENSE.txt", "licenses/TESSERACT-LICENSE.txt"} <= set(bundle.namelist())
        cases["html_and_zip_buttons_export_clean_pc_entry_and_pc_manual"] = True

        page.reload()
        ready(page)
        assert_pc(page)
        assert page.evaluate("ToolApp.state.source.connected&&!!ToolApp.state.model")
        assert page.evaluate("testPickerCalls") == 1
        cases["reload_restores_selected_handle_without_new_picker"] = True

        page.evaluate("sessionStorage.setItem('pc-test-permission','prompt')")
        page.reload()
        ready(page)
        assert page.evaluate("ToolApp.state.source.needsPermission&&!ToolApp.state.source.connected")
        assert "권한" in page.locator("#quick-workbook").inner_text()
        before = page.evaluate("testPickerCalls")
        page.locator("#quick-workbook").click()
        page.wait_for_function("ToolApp.state.source.connected&&!ToolApp.state.busy", timeout=60000)
        assert page.evaluate("testPickerCalls") == before
        calls = page.evaluate("JSON.parse(sessionStorage.getItem('pc-test-permission-calls')||'[]')")
        assert len(calls) == 1 and calls[0]["name"] == "company-shared.xlsx" and calls[0]["active"], calls
        assert not page.evaluate("ToolApp.state.source.needsPermission")
        cases["quick_permission_button_reauthorizes_same_handle_without_file_picker"] = True
        native.close()

        fallback = context(False)
        copy_page = fallback.new_page()
        copy_page.goto(url)
        ready(copy_page)
        assert not copy_page.evaluate("ToolConnectedFile.inspect().supported")
        with copy_page.expect_file_chooser() as chooser:
            copy_page.locator("#quick-workbook").click()
        chooser.value.set_files(workbook)
        copy_page.wait_for_function("ToolSources.mode()==='device'&&ToolApp.state.source.connected&&!ToolApp.state.busy", timeout=60000)
        assert "원본은 변경하지" in copy_page.locator("#quick-save-note").inner_text()
        assert "기기에" in copy_page.locator("#quick-summary").inner_text()
        controlled_ocr(copy_page)
        copy_page.locator("#photos").set_input_files(photo)
        copy_page.wait_for_function("ToolApp.state.photos[0]?.saved&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=60000)
        assert copy_page.evaluate("ToolSources.mode()") == "device"
        assert "회사 공유 원본은 변경하지" in copy_page.locator("#status").inner_text()
        cases["unsupported_picker_explicitly_uses_device_copy_without_original_claim"] = True
        fallback.close()
        browser.close()
        assert sha256(workbook.read_bytes()).hexdigest() == baseline
    assert not errors and not external and not published, (errors, external, published)
    report.update({"case_count": len(cases), "page_errors": errors, "external_requests": external,
                   "published_settings_requests": published, "synthetic_source_workbook_unchanged": True})
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
