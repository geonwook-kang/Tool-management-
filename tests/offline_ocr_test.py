"""Real Korean/English OCR to XLSX with network disabled after HTML loads.

This uses a generated clear printed label, not the user's handwritten photo.
No recognizer, parser, worker, or response is mocked or intercepted for success.
"""
from hashlib import sha256
import json
from pathlib import Path
import tempfile

from PIL import Image, ImageDraw, ImageFont
from playwright.sync_api import sync_playwright

from browser_test import arguments, serve_program, verify_workbook


def synthetic_label(output):
    image = Image.new("RGB", (1800, 1050), "white")
    draw = ImageDraw.Draw(image)
    font_path = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
    name_font = ImageFont.truetype(font_path, 94)
    serial_font = ImageFont.truetype(font_path, 82)
    font = ImageFont.truetype(font_path, 58)
    draw.text((95, 75), "강건욱", font=name_font, fill="black")
    draw.text((95, 250), "Serial: 31550075", font=serial_font, fill="black")
    draw.text((95, 455), "Date of Calibration", font=font, fill="black")
    draw.text((95, 555), "2026.12.29", font=serial_font, fill="black")
    draw.text((95, 745), "Due date of next Calibration", font=font, fill="black")
    draw.text((95, 845), "2027.12.29", font=serial_font, fill="black")
    image.save(output)


def main():
    args = arguments(__doc__)
    original_hash = sha256(args.workbook.read_bytes()).hexdigest()
    with tempfile.TemporaryDirectory(prefix="tool-offline-ocr-") as directory:
        temporary = Path(directory)
        label = temporary / "printed-korean-label.png"
        synthetic_label(label)
        with serve_program(args.program) as url, sync_playwright() as p:
            browser = p.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
            context = browser.new_context(accept_downloads=True, locale="ko-KR")
            page = context.new_page()
            page.set_default_timeout(20000)
            errors, requests = [], []
            page.on("pageerror", lambda error: errors.append(str(error)))
            page.goto(url)
            page.wait_for_function("window.ToolApp&&window.ToolOCR&&window.ToolRuntime")
            context.route("**/*", lambda route: (requests.append(route.request.url), route.abort()) if route.request.url.startswith(("http:", "https:")) else route.continue_())
            context.set_offline(True)
            page.locator("#workbook").set_input_files(args.workbook)
            page.wait_for_function("ToolApp.state.model&&!ToolApp.state.busy")
            assert page.locator("#automatic-mode").is_checked()
            page.locator("#photos").set_input_files(label)
            page.wait_for_function("ToolApp.state.photos.length===1&&!ToolApp.state.busy&&!ToolApp.state.preparing", timeout=180000)
            result = page.evaluate("""()=>{
              const p=ToolApp.state.photos[0],r=p.result||{};
              return {approval:p.approval,error:p.error,assessment:p.assessment,text:r.text,identity:r.identityEvidence,dates:r.dateEvidence,prepared:!!ToolApp.state.preparedResult};
            }""")
            print(json.dumps({"ocr_diagnostics": result}, ensure_ascii=False, indent=2), flush=True)
            assert result["prepared"], "Real offline OCR failed to prepare XLSX; see printed diagnostics"
            assert result["approval"] == "automatic" and result["assessment"]["eligible"]
            assert result["assessment"]["name"] == "강건욱"
            assert result["assessment"]["serial"] == "31550075"
            assert result["assessment"]["last"] == "2026-12-29"
            assert result["assessment"]["due"] == "2027-12-29"
            with page.expect_download() as event:
                page.locator("#automatic-result-download").click()
            updated = temporary / "offline-real-ocr-result.xlsx"
            event.value.save_as(updated)
            verification = verify_workbook(args.workbook, updated)
            assert not requests, f"HTTP(S) attempted while offline: {requests}"
            assert not errors, errors
            browser.close()
    assert original_hash == sha256(args.workbook.read_bytes()).hexdigest()
    print(json.dumps({"offline_real_ocr": "passed", "printed_label_only": True, "http_requests": requests, "workbook": verification, "original_unchanged": True}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
