"""Run real offline OCR in a sandboxed frame with an opaque (null) origin.

Windows file:// pages can create blob:null workers. An HTTP-only smoke test
does not cover that worker origin. This test uses sandbox="allow-scripts
allow-downloads" without allow-same-origin to reproduce it without weakening
browser security, TLS verification, or the managed file:// URL policy.

By default a restrictive script policy permits the initial Blob worker and
embedded WebAssembly, but prevents that worker importing another script. This
recreates the failed secondary Blob import without pretending the sandbox is
identical to Windows file:// policy. Use --allow-script-imports to test the
opaque origin alone; some Chromium versions permit both Blob imports there.

The image is a generated clear printed label, not the user's handwritten photo.
No OCR response, worker, browser origin, or workbook operation is mocked.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import tempfile
import time
from zipfile import ZipFile

from playwright.sync_api import sync_playwright
from PIL import Image, ImageChops, ImageStat

from browser_test import REPO, serve_program, verify_workbook
from offline_ocr_test import synthetic_label


def arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--program", type=Path, default=REPO / "web/dist/Tool-Manager.html")
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    parser.add_argument("--expect-import-error", action="store_true", help="Verify the old blob:null importScripts failure instead of expecting a successful workbook.")
    parser.add_argument("--allow-script-imports", action="store_true", help="Test the null origin alone without restricting secondary worker script imports.")
    args = parser.parse_args()
    if args.workbook is None:
        args.workbook = next(Path("/workspace/attachments").rglob("*.xlsx"), None)
    if args.workbook is None or not args.workbook.is_file():
        parser.error("Supply --workbook /absolute/path/to/the/attached/company.xlsx")
    if not args.program.is_file():
        parser.error("Supply --program /absolute/path/to/the/standalone-program.html")
    return args


def wait_for(frame, expression, timeout=20):
    # Playwright's wait_for_function internally calls eval in the page. Poll
    # through the browser protocol so the restrictive script policy can remain
    # intact instead of adding unsafe-eval merely for the test runner.
    until = time.monotonic() + timeout
    while time.monotonic() < until:
        if frame.evaluate("() => Boolean(" + expression + ")"):
            return
        time.sleep(0.05)
    raise AssertionError("Timed out waiting for " + expression)


def main():
    args = arguments()
    original_hash = sha256(args.workbook.read_bytes()).hexdigest()
    html = args.program.read_text(encoding="utf-8")
    if not args.allow_script_imports:
        policy = '<meta http-equiv="Content-Security-Policy" content="script-src &apos;unsafe-inline&apos; &apos;wasm-unsafe-eval&apos;; worker-src blob:; connect-src &apos;none&apos;">'
        assert "<head>" in html, "The standalone program must have a head element"
        html = html.replace("<head>", "<head>" + policy, 1)
    verification = None
    with tempfile.TemporaryDirectory(prefix="tool-opaque-ocr-") as directory:
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
            # The parent is ordinary local HTTP; only the program's frame gets
            # an opaque origin. No browser security flag changes that origin.
            page.goto(url.rsplit("/", 1)[0] + "/")
            page.evaluate("""source => {
              const frame=document.createElement('iframe');
              frame.id='opaque-program';
              frame.setAttribute('sandbox','allow-scripts allow-downloads');
              frame.style.cssText='width:1280px;height:1600px;border:0';
              frame.srcdoc=source;
              document.body.replaceChildren(frame);
            }""", html)
            frame = page.locator("#opaque-program").element_handle().content_frame()
            assert frame is not None
            wait_for(frame, "window.ToolApp&&window.ToolOCR&&window.ToolRuntime")
            origin = frame.evaluate("""() => {
              const url=URL.createObjectURL(new Blob(['test'],{type:'text/plain'}));
              const result={origin:self.origin,url:location.href,blobURL:url};
              URL.revokeObjectURL(url);return result;
            }""")
            assert origin["origin"] == "null" and origin["blobURL"].startswith("blob:null/"), origin

            def no_http(route):
                if route.request.url.startswith(("http:", "https:")):
                    requests.append(route.request.url)
                    route.abort()
                else:
                    route.continue_()

            context.route("**/*", no_http)
            context.set_offline(True)
            frame.locator("#workbook").set_input_files(args.workbook)
            wait_for(frame, "ToolApp.state.model&&!ToolApp.state.busy")
            assert frame.locator("#automatic-mode").is_checked()
            frame.locator("#photos").set_input_files(label)
            wait_for(frame, "ToolApp.state.photos.length===1&&!ToolApp.state.busy&&!ToolApp.state.preparing", timeout=180)
            result = frame.evaluate("""() => {
              const p=ToolApp.state.photos[0],r=p.result||{};
              return {approval:p.approval,error:p.error,failed:r.failed,
                ocrError:r.error,assessment:p.assessment,text:r.text,
                identity:r.identityEvidence,dates:r.dateEvidence,
                prepared:!!ToolApp.state.preparedResult};
            }""")
            print(json.dumps({"origin": origin, "ocr_diagnostics": result, "page_errors": errors}, ensure_ascii=False, indent=2), flush=True)
            assert not requests, f"HTTP(S) requested while offline: {requests}"

            if args.expect_import_error:
                diagnosis = "\n".join([str(result.get("ocrError", "")), str(result.get("error", "")), *errors])
                assert not result["prepared"] and result.get("failed"), result
                assert "importScripts" in diagnosis and "blob:null/" in diagnosis and "core.js" in diagnosis, diagnosis
            else:
                assert result["prepared"], "Opaque-origin OCR failed to prepare XLSX; see diagnostics above"
                assert result["approval"] == "automatic" and result["assessment"]["eligible"], result
                assert result["assessment"]["name"] == "강건욱"
                assert result["assessment"]["serial"] == "31550075"
                assert result["assessment"]["last"] == "2026-12-29"
                assert result["assessment"]["due"] == "2027-12-29"
                # Require the actual OCR text as well as the workbook match.
                assert "강건욱" in result["text"].replace(" ", "")
                assert "31550075" in result["text"].replace(" ", "")
                with page.expect_download() as event:
                    frame.locator("#automatic-result-download").click()
                updated = temporary / "opaque-origin-ocr-result.xlsx"
                event.value.save_as(updated)
                verification = verify_workbook(args.workbook, updated)
                with ZipFile(args.workbook) as before, ZipFile(updated) as after:
                    new_media = [name for name in after.namelist() if name not in before.namelist() and name.startswith("xl/media/") and not name.endswith("/")]
                    assert len(new_media) == 1
                    # The application normalizes uploaded pictures to JPEG.
                    # Check content instead of comparing the input PNG bytes.
                    with Image.open(BytesIO(after.read(new_media[0]))) as inserted, Image.open(label) as source:
                        assert inserted.size == source.size
                        difference = ImageStat.Stat(ImageChops.difference(inserted.convert("RGB"), source.convert("RGB")))
                        assert max(difference.mean) < 3, difference.mean
                        verification["inserted_photo_matches"] = True
                assert not errors, errors
            browser.close()
    assert original_hash == sha256(args.workbook.read_bytes()).hexdigest(), "Original workbook changed"
    print(json.dumps({"opaque_origin": "null", "secondary_script_imports_restricted": not args.allow_script_imports, "old_import_error_reproduced": args.expect_import_error, "offline_real_ocr": "not expected" if args.expect_import_error else "passed", "printed_label_only": True, "http_requests": requests, "workbook": verification, "original_unchanged": True}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
