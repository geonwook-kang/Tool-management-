"""Browser integration tests for automatic upload, manual fallback, and errors.

OCR results are deterministic parser-based stubs here. Real local OCR is tested
separately by offline_ocr_test.py with every HTTP(S) request blocked.
"""
from hashlib import sha256
import json
from pathlib import Path
import tempfile
from time import monotonic

from playwright.sync_api import sync_playwright

from browser_test import arguments, fixture_photo, serve_program, verify_workbook
from offline_ocr_test import synthetic_label

TEXT = "강건욱\n31550075\n교정일자 2026.12.29\n차기교정예정일자 2027.12.29"


def ready(page, count=1):
    page.wait_for_function("n=>ToolApp.state.photos.length===n&&!ToolApp.state.busy&&!ToolApp.state.preparing", arg=count, timeout=30000)


def install_stub(page, confidence=95, error=None):
    page.evaluate("""({text,confidence,error})=>{
      window.testOriginalOCR=window.ToolOCR;
      window.ToolOCR={...window.ToolOCR,recognize:async(file,tools,progress)=>{
        if(error)throw new Error(error);
        if(progress)progress({status:'테스트 문자 인식',progress:1});
        return testOriginalOCR.parse(text,tools,{confidence,passes:[{text,confidence},{text,confidence}]});
      }};
    }""", {"text": TEXT, "confidence": confidence, "error": error})


def load(page, url, workbook):
    page.goto(url)
    page.locator("#workbook").set_input_files(workbook)
    page.wait_for_function("window.ToolApp&&ToolApp.state.model&&!ToolApp.state.busy")
    assert page.locator("#automatic-mode").is_checked()


def auto_success(page, url, args, temporary, photo):
    load(page, url, args.workbook)
    install_stub(page)
    page.locator("#photos").set_input_files(photo)
    ready(page)
    page.wait_for_function("ToolApp.state.preparedResult&&ToolApp.state.preparedResult.bytes.length>0")
    record = page.evaluate("""()=>{const p=ToolApp.state.photos[0];return {approval:p.approval,eligible:p.assessment.eligible,review:!!p.review,toolId:p.draft.toolId,name:p.draft.name,serial:p.draft.serial,last:p.draft.last,due:p.draft.due};}""")
    assert record["approval"] == "automatic" and record["eligible"] and record["review"]
    assert record["name"] == "강건욱" and record["serial"] == "31550075"
    assert record["last"] == "2026-12-29" and record["due"] == "2027-12-29"
    assert page.locator("#automatic-result-download").is_visible()
    with page.expect_download() as event:
        page.locator("#automatic-result-download").click()
    output = temporary / "automatic-result.xlsx"
    event.value.save_as(output)
    result = verify_workbook(args.workbook, output)
    result["approval"] = record["approval"]
    return result


def manual_fallback(page, url, args, temporary, photo):
    load(page, url, args.workbook)
    install_stub(page, confidence=25)
    page.locator("#photos").set_input_files(photo)
    ready(page)
    result = page.evaluate("""()=>{const p=ToolApp.state.photos[0];return {approval:p.approval,eligible:p.assessment.eligible,review:!!p.review,reasons:p.assessment.reasons,prepared:!!ToolApp.state.preparedResult};}""")
    assert not result["eligible"] and not result["review"] and not result["prepared"]
    assert result["reasons"]
    assert not page.locator("#automatic-result-download").is_visible()
    tool_id = page.evaluate("ToolApp.state.model.tools.find(t=>t.sheet==='광주'&&t.row===6).id")
    page.locator("#tool-select").select_option(tool_id)
    page.locator("#last").fill("2026-12-29")
    page.locator("#due").fill("2027-12-29")
    page.locator("#confirm").check()
    page.locator("#save-review").click()
    ready(page)
    assert page.evaluate("ToolApp.state.photos[0].approval") == "manual"
    assert page.evaluate("!!ToolApp.state.photos[0].review")
    return {"low_confidence_requires_review": True, "manual_correction_available": True}


def ocr_failure(page, url, args, photo):
    load(page, url, args.workbook)
    install_stub(page, error="TEST_OCR_FAILURE")
    page.locator("#photos").set_input_files(photo)
    ready(page)
    result = page.evaluate("""()=>{const p=ToolApp.state.photos[0];return {error:p.error,review:!!p.review,prepared:!!ToolApp.state.preparedResult,busy:ToolApp.state.busy};}""")
    assert "TEST_OCR_FAILURE" in str(result["error"])
    assert not result["review"] and not result["prepared"] and not result["busy"]
    assert page.locator("#review-panel").is_visible()
    assert page.locator("#tool-select").is_enabled()
    return {"diagnostic_visible": True, "manual_fallback_available": True}


def generation_failure(page, url, args, temporary, photo):
    load(page, url, args.workbook)
    install_stub(page)
    page.evaluate("""()=>{window.testOriginalWorkbook=window.ToolWorkbook;window.ToolWorkbook={...ToolWorkbook,apply:async()=>{throw new Error('TEST_GENERATION_FAILURE')}};}""")
    page.locator("#photos").set_input_files(photo)
    ready(page)
    assert page.evaluate("!!ToolApp.state.photos[0].review")
    assert not page.evaluate("!!ToolApp.state.preparedResult")
    assert "TEST_GENERATION_FAILURE" in page.locator("#status").inner_text()
    assert not page.locator("#automatic-result-download").is_visible()
    page.evaluate("window.ToolWorkbook=window.testOriginalWorkbook")
    with page.expect_download() as event:
        page.locator("#export").click()
    output = temporary / "retry-result.xlsx"
    event.value.save_as(output)
    ready(page)
    verify_workbook(args.workbook, output)
    return {"generation_error_reported": True, "retry_succeeded": True}


def cancellation_retry(page, url, args, temporary):
    load(page, url, args.workbook)
    label = temporary / "cancel-retry-label.png"
    synthetic_label(label)
    # Stall actual worker initialization so the real ToolOCR cancellation race
    # and the application's cancel button run, rather than mocking recognize.
    page.evaluate("""()=>{
      window.originalWorkerCreation=Tesseract.createWorker;
      Tesseract.createWorker=()=>{window.hungWorkerRequested=true;return new Promise(()=>{});};
    }""")
    page.locator("#photos").set_input_files(label)
    page.wait_for_function("window.hungWorkerRequested&&ToolApp.state.busy")
    began = monotonic()
    page.locator("#cancel-recognition").click()
    page.wait_for_function("!ToolApp.state.busy&&!ToolApp.state.preparing", timeout=5000)
    elapsed = monotonic() - began
    assert "중지" in page.locator("#status").inner_text()
    assert not page.evaluate("!!ToolApp.state.preparedResult")
    assert page.locator("#tool-select").is_enabled()
    page.evaluate("()=>{Tesseract.createWorker=window.originalWorkerCreation;}")
    page.locator("#recognize-one").click()
    page.wait_for_function("!ToolApp.state.busy&&!ToolApp.state.preparing&&!!ToolApp.state.preparedResult", timeout=180000)
    assert page.evaluate("ToolApp.state.photos[0].approval") == "automatic"
    return {"cancelled_initialization_seconds": round(elapsed, 3), "actual_ocr_retry_succeeded": True}


def main():
    args = arguments(__doc__)
    original_hash = sha256(args.workbook.read_bytes()).hexdigest()
    with tempfile.TemporaryDirectory(prefix="tool-automation-test-") as directory:
        temporary = Path(directory)
        photo = fixture_photo(args, temporary)
        with serve_program(args.program) as url, sync_playwright() as p:
            browser = p.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
            errors = []
            result = {}
            cases = [
                ("automatic", lambda page: auto_success(page, url, args, temporary, photo)),
                ("manual_fallback", lambda page: manual_fallback(page, url, args, temporary, photo)),
                ("ocr_error", lambda page: ocr_failure(page, url, args, photo)),
                ("generation_error", lambda page: generation_failure(page, url, args, temporary, photo)),
                ("cancellation", lambda page: cancellation_retry(page, url, args, temporary)),
            ]
            for name, case in cases:
                context = browser.new_context(accept_downloads=True, locale="ko-KR")
                page = context.new_page()
                page.on("dialog", lambda dialog: dialog.accept())
                page.on("pageerror", lambda error: errors.append(str(error)))
                result[name] = case(page)
                context.close()
            assert not errors, errors
            browser.close()
    assert original_hash == sha256(args.workbook.read_bytes()).hexdigest()
    result["original_unchanged"] = True
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
