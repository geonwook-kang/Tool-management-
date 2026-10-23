"""Engine startup failure must not masquerade as unreadable handwriting.

The browser writes a real OPFS workbook. OCR responses are explicit test doubles
for engine-init failure and recovery; this is not an OCR accuracy test.
"""
from hashlib import sha256
import json
from pathlib import Path
import tempfile

from playwright.sync_api import sync_playwright

from browser_test import arguments, fixture_photo, serve_program, verify_workbook
from connected_file_test import PICKER_SCRIPT, export_shared, seed_file


def settled(page):
    page.wait_for_function("window.ToolApp&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=30000)


def setup_failure(context, page, url, args, photo, error_kind="flag"):
    context.add_init_script("(" + PICKER_SCRIPT + ")()")
    page.goto(url)
    page.wait_for_function("window.ToolApp")
    seed_file(page, args.workbook)
    page.locator("#connect-onedrive").click()
    page.wait_for_function("ToolApp.state.source.connected&&!!ToolApp.state.model")
    settled(page)
    page.evaluate("""kind=>{
      const parse=ToolOCR.parse,apply=ToolWorkbook.apply,save=ToolConnectedFile.save;
      window.testApplyCalls=0;window.testSaveCalls=0;
      window.ToolWorkbook={...ToolWorkbook,apply:async(...args)=>{testApplyCalls++;return apply(...args);}};
      window.ToolConnectedFile={...ToolConnectedFile,save:async(...args)=>{testSaveCalls++;return save(...args);}};
      window.ToolOCR={...ToolOCR,recognize:async(file,tools)=>({...parse('',tools,{confidence:0}),
        failed:true,engineFailure:kind==='flag',errorCode:kind==='code'?'engine-init':'',phase:'initialize',text:'',
        error:'Failed to execute importScripts on WorkerGlobalScope: blob:null/test#core.js failed to load.'})};
    }""", error_kind)
    page.locator("#photos").set_input_files(photo)
    page.wait_for_function("ToolApp.state.photos.length===1&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving")


def assert_engine_failure(page, args, temporary, name):
    assert page.locator("#ocr-engine-error").is_visible()
    assert "인식 엔진 실행 실패" in page.locator("#status").inner_text()
    assert "엔진" in page.locator("#process-summary").inner_text()
    card = page.locator("#photo-list .photo-card .photo-state").inner_text()
    assert "엔진" in card and "확인 필요" not in card, card
    assert "확인 필요 1장" not in page.locator("#process-summary").inner_text()
    assert page.locator("#retry-engine").is_visible()
    assert page.locator("#manual-entry-toggle").is_visible()
    assert page.locator("#name").input_value() == ""
    assert page.locator("#serial").input_value() == ""
    assert not page.locator("#name").is_visible() or page.locator("#name").is_disabled()
    assert not page.locator("#serial").is_visible() or page.locator("#serial").is_disabled()
    assert page.evaluate("testApplyCalls===0&&testSaveCalls===0")
    assert not page.evaluate("!!ToolApp.state.preparedResult||!!ToolApp.state.photos[0].review||!!ToolApp.state.photos[0].saved")
    source = temporary / (name + "-unchanged.xlsx")
    export_shared(page, source)
    assert source.read_bytes() == args.workbook.read_bytes()


def retry_recovers(context, page, url, args, temporary, photo):
    setup_failure(context, page, url, args, photo, error_kind="code")
    assert_engine_failure(page, args, temporary, "retry")
    page.screenshot(path="/tmp/tool-manager-engine-error.png", full_page=True)
    page.evaluate("""()=>{
      const parse=ToolOCR.parse,text='강건욱\\n31550075\\n교정일자 2025.12.29\\n차기교정예정일자 2026.12.29';
      window.ToolOCR={...ToolOCR,recognize:async(file,tools)=>parse(text,tools,{confidence:95,passes:[{text,confidence:95},{text,confidence:95}]})};
    }""")
    # No form fields are filled, selected, or confirmed during retry.
    page.locator("#retry-engine").click()
    page.wait_for_function("ToolApp.state.photos[0].approval==='automatic'&&!!ToolApp.state.photos[0].saved&&!ToolApp.state.source.saving", timeout=30000)
    settled(page)
    assert not page.locator("#ocr-engine-error").is_visible()
    assert page.evaluate("testApplyCalls===1&&testSaveCalls===1")
    record = page.evaluate("""()=>{const p=ToolApp.state.photos[0];return {approval:p.approval,name:p.saved.name,serial:p.saved.serial,row:p.saved.row};}""")
    assert record == {"approval": "automatic", "name": "강건욱", "serial": "31550075", "row": 6}
    updated = temporary / "retry-saved.xlsx"
    export_shared(page, updated)
    workbook = verify_workbook(args.workbook, updated, last="2025-12-29", due="2026-12-29")
    page.screenshot(path="/tmp/tool-manager-engine-recovered.png", full_page=True)
    return {"engine_failure_not_mislabelled": True, "source_unchanged_on_failure": True, "retry_without_typing_saves_automatically": True, "workbook": workbook}


def voluntary_manual(context, page, url, args, temporary, photo):
    setup_failure(context, page, url, args, photo)
    assert_engine_failure(page, args, temporary, "manual")
    toggle = page.locator("#manual-entry-toggle")
    if toggle.get_attribute("type") == "checkbox":
        toggle.check()
    else:
        toggle.click()
    assert page.locator("#name").is_visible() and page.locator("#name").is_enabled()
    assert page.locator("#serial").is_visible() and page.locator("#serial").is_enabled()
    assert page.evaluate("testApplyCalls===0&&testSaveCalls===0")
    assert not page.evaluate("!!ToolApp.state.photos[0].review||!!ToolApp.state.photos[0].saved")
    return {"manual_fields_hidden_until_opt_in": True, "opt_in_itself_does_not_save": True}


def main():
    args = arguments(__doc__)
    initial_hash = sha256(args.workbook.read_bytes()).hexdigest()
    with tempfile.TemporaryDirectory(prefix="tool-engine-failure-") as directory, serve_program(args.program) as url, sync_playwright() as p:
        temporary = Path(directory)
        photo = fixture_photo(args, temporary)
        browser = p.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
        results, errors = {}, []
        for name, case in [("retry", retry_recovers), ("voluntary_manual", voluntary_manual)]:
            context = browser.new_context(accept_downloads=True, locale="ko-KR")
            page = context.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))
            results[name] = case(context, page, url, args, temporary, photo)
            context.close()
        browser.close()
    assert not errors, errors
    assert sha256(args.workbook.read_bytes()).hexdigest() == initial_hash
    results.update({"ocr_results_are_test_doubles": True, "uploaded_original_unchanged": True})
    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
