"""Download a selected photo's reproducible OCR diagnostic ZIP in Chromium.

The photo is generated for this test. OCR responses are controlled so the test
can verify exact source confidence/geometry and engine-failure diagnostics;
this is not a photo recognition accuracy test. The company XLSX is read only.
"""
from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path
import tempfile
from urllib.parse import urlsplit
from zipfile import ZipFile

from PIL import Image, ImageDraw
from playwright.sync_api import sync_playwright

from browser_test import arguments, serve_program


PRIVATE = "diagnostics-private-unrelated-state-9c6ef2"
OTHER = "diagnostics-unselected-photo-7c0b4a"


def jpeg_fixture(path: Path, text: str, color: str):
    image = Image.new("RGB", (640, 840), color)
    draw = ImageDraw.Draw(image)
    draw.rectangle((100, 120, 550, 530), fill="white", outline="black", width=3)
    draw.text((135, 185), text, fill="black", font_size=24)
    exif = Image.Exif()
    exif[270] = text
    exif[305] = "diagnostics-test-original-exif"
    image.save(path, format="JPEG", quality=94, exif=exif)
    return path


def settled(page):
    page.wait_for_function("""window.ToolApp&&!ToolApp.state.busy&&
        !ToolApp.state.preparing&&!ToolApp.state.source.saving""", timeout=30000)


def protect_unrelated_state(page):
    if not page.locator("#sharepoint-url").is_visible():
        page.locator("#sharepoint-url").locator("xpath=ancestor::details[1]").locator("summary").click()
    page.locator("#sharepoint-url").fill("https://example.invalid/" + PRIVATE)
    page.locator("#remember-sharepoint").click()
    page.evaluate("""({privateMarker,otherMarker})=>{
      const state=ToolApp.state;
      state.model={...state.model,diagnosticPrivateMarker:privateMarker};
      state.workbookName=privateMarker+'.xlsx';
      state.source.diagnosticPrivateMarker=privateMarker;
      state.photos[0].draft={...state.photos[0].draft,diagnosticPrivateMarker:privateMarker};
      state.photos[1].result={text:otherMarker,failed:false,diagnosticPrivateMarker:privateMarker};
      state.photos[1].draft={name:otherMarker};
      localStorage.setItem('diagnostics-unrelated-key',privateMarker);
    }""", {"privateMarker": PRIVATE, "otherMarker": OTHER})


def download_report(page, temporary, label, original, other):
    assert page.locator("#download-diagnostics").is_visible()
    assert page.locator("#download-diagnostics").is_enabled()
    with page.expect_download(timeout=30000) as event:
        page.locator("#download-diagnostics").click()
    download = event.value
    assert download.suggested_filename == "Tool-Recognition-Diagnostics.zip"
    assert download.failure() is None
    output = temporary / (label + "-diagnostics.zip")
    download.save_as(output)
    with ZipFile(output) as archive:
        assert archive.testzip() is None
        assert set(archive.namelist()) == {"original.jpg", "recognition.json"}
        photo = archive.read("original.jpg")
        raw = archive.read("recognition.json")
    assert photo == original.read_bytes(), "Original JPEG/EXIF must not be recompressed"
    assert photo != other.read_bytes()
    assert PRIVATE.encode() not in raw and OTHER.encode() not in raw
    assert other.name.encode() not in raw
    report = json.loads(raw)
    assert set(report) == {"schema", "programVersion", "createdAt", "photo", "ocr", "assessment", "processing"}
    assert report["schema"] == 1 and report["programVersion"] == "4.5"
    assert report["photo"]["entry"] == "original.jpg"
    assert report["photo"]["name"] == original.name
    assert report["photo"]["size"] == len(original.read_bytes())
    assert report["photo"]["type"] == "image/jpeg"
    assert report["processing"] == {"approval": None, "saved": False}
    source = page.evaluate("""()=>{
      const p=ToolApp.state.photos[0];return {result:p.result,assessment:p.assessment,
        prepared:!!ToolApp.state.preparedResult,review:!!p.review,saved:!!p.saved,
        applyCalls:window.diagnosticApplyCalls,saveCalls:window.diagnosticSaveCalls};
    }""")
    assert not source["prepared"] and not source["review"] and not source["saved"]
    assert source["applyCalls"] == source["saveCalls"] == 0
    assert report["assessment"] == source["assessment"]
    return report, source["result"]


def install_success_response(page):
    page.evaluate("""()=>{
      const serial='0716805829',text=serial+'\\n2026.06.22\\n2026.12.22';
      const words=[
        {text:serial,confidence:91.625,bbox:{x0:70,y0:30,x1:330,y1:62},lineBox:{x0:65,y0:25,x1:335,y1:67}},
        {text:'2026.06.22',confidence:86.25,bbox:{x0:100,y0:140,x1:350,y1:176}},
        {text:'2026.12.22',confidence:93.75,bbox:{x0:100,y0:205,x1:350,y1:241}}];
      window.ToolOCR={...ToolOCR,recognize:async(file,tools)=>{
        const result=ToolOCR.parse(text,tools,{passes:[
          {text,confidence:88.5,source:'original',words},
          {text:'2026.06.22\\n2026.12.22',confidence:90.5,source:'date-crop',words:words.slice(1)}]});
        result.internalUnrelatedState='diagnostics-private-unrelated-state-9c6ef2';
        return result;
      }};
    }""")


def install_engine_failure(page):
    page.evaluate("""()=>{
      window.ToolOCR={...ToolOCR,recognize:async(file,tools)=>({...ToolOCR.parse('',tools),
        failed:true,engineFailure:true,errorCode:'engine-init',code:'engine-init',
        error:'Diagnostic fixture: embedded OCR initialization failed.'})};
    }""")


def main():
    args = arguments(__doc__)
    original_hash = sha256(args.workbook.read_bytes()).hexdigest()
    errors, requests = [], []
    with tempfile.TemporaryDirectory(prefix="tool-diagnostics-") as directory:
        temporary = Path(directory)
        original = jpeg_fixture(temporary / "selected-original.jpg", "ORIGINAL-JPEG-SOURCE-EXIF", "silver")
        other = jpeg_fixture(temporary / "unselected-other.jpg", OTHER, "yellow")
        with serve_program(args.program) as url, sync_playwright() as playwright:
            browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
            context = browser.new_context(accept_downloads=True, locale="ko-KR")
            page = context.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))
            page.on("dialog", lambda dialog: dialog.accept())
            page.goto(url)
            settled(page)
            assert page.locator("#download-diagnostics").count() == 1, "Build v4.5 before running diagnostics test"
            origin = urlsplit(url).netloc
            context.route("**/*", lambda route: (requests.append(route.request.url), route.abort())
                          if urlsplit(route.request.url).scheme in {"http", "https"}
                          and urlsplit(route.request.url).netloc != origin else route.continue_())
            context.set_offline(True)
            page.locator("#automatic-mode").uncheck()
            page.locator("#workbook").set_input_files(args.workbook)
            page.wait_for_function("ToolApp.state.model&&!ToolApp.state.busy")
            page.evaluate("""()=>{
              diagnosticApplyCalls=0;diagnosticSaveCalls=0;
              const apply=ToolWorkbook.apply,save=ToolConnectedFile.save;
              window.ToolWorkbook={...ToolWorkbook,apply:async(...input)=>{diagnosticApplyCalls++;return apply(...input);}};
              window.ToolConnectedFile={...ToolConnectedFile,save:async(...input)=>{diagnosticSaveCalls++;return save(...input);}};
            }""")
            page.locator("#photos").set_input_files([original, other])
            settled(page)
            assert page.evaluate("ToolApp.state.photos.length") == 2
            protect_unrelated_state(page)
            empty, result = download_report(page, temporary, "before-recognition", original, other)
            assert result is None
            assert empty["ocr"]["text"] == "" and empty["ocr"]["confidence"] is None
            assert empty["ocr"]["passes"] == [] and empty["ocr"]["identityEvidence"] is None

            install_success_response(page)
            page.locator("#recognize-one").click()
            settled(page)
            protect_unrelated_state(page)
            recognized, result = download_report(page, temporary, "recognized", original, other)
            assert not recognized["ocr"]["failed"]
            assert recognized["ocr"]["text"] == result["text"]
            assert recognized["ocr"]["passes"] == result["passes"]
            assert recognized["ocr"]["identityEvidence"] == result["identityEvidence"]
            assert recognized["ocr"]["dateEvidence"] == result["dateEvidence"]
            words = recognized["ocr"]["passes"][0]["words"]
            assert words[0]["confidence"] == 91.625
            assert words[0]["bbox"] == {"x0": 70, "y0": 30, "x1": 330, "y1": 62}
            assert words[0]["lineBox"] == {"x0": 65, "y0": 25, "x1": 335, "y1": 67}
            assert recognized["ocr"]["passes"][1]["source"] == "date-crop"
            serial = next(proof for proof in recognized["ocr"]["identityEvidence"]["observedSerialEvidence"]
                          if proof["serial"] == "0716805829" and proof["kind"] == "workbook")
            assert {"source": "original", "confidence": 91.625} in serial["sourceReadings"]

            install_engine_failure(page)
            page.locator("#recognize-one").click()
            settled(page)
            assert page.locator("#ocr-engine-error").is_visible()
            assert not page.locator("#review-form").is_visible()
            assert page.evaluate("!ToolApp.state.photos[0].manualEntry")
            protect_unrelated_state(page)
            failed, result = download_report(page, temporary, "engine-failure", original, other)
            assert failed["ocr"]["failed"] and failed["ocr"]["engineFailure"]
            assert failed["ocr"]["code"] == "engine-init"
            assert failed["ocr"]["error"] == result["error"]
            assert failed["ocr"]["text"] == ""
            assert failed["ocr"]["passes"] == result["passes"]
            assert all(not item["text"] for item in failed["ocr"]["passes"])
            assert not errors, errors
            assert not requests, requests
            browser.close()
    assert sha256(args.workbook.read_bytes()).hexdigest() == original_hash
    print(json.dumps({"program": str(args.program), "program_sha256": sha256(args.program.read_bytes()).hexdigest(),
                      "states": ["before recognition", "recognized with source confidence/geometry", "engine initialization failure"],
                      "original_jpeg_and_exif_byte_identical": True,
                      "unrelated_photo_workbook_model_private_url_and_storage_excluded": True,
                      "no_excel_generation_or_write_side_effect": True,
                      "original_company_workbook_unchanged": True,
                      "page_errors": errors, "external_network_requests": requests,
                      "ocr_accuracy": "not measured: controlled OCR responses"}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
