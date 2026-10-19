"""Exercise HTML/ZIP self-downloads and restricted-preview source fallback.

Run after the standalone build. Requires Playwright, Chromium, openpyxl, and
Pillow. Fixtures contain synthetic data only. A localhost server is used
because the managed cloud browser can block file:// navigation. Passing these
checks does not establish that ChatGPT's separate attachment service works.
"""
from __future__ import annotations

import argparse
from contextlib import contextmanager
from datetime import datetime
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
import json
from pathlib import Path
import tempfile
from threading import Thread
from urllib.parse import quote
from zipfile import ZipFile

from openpyxl import Workbook
from PIL import Image
from playwright.sync_api import sync_playwright


REPO = Path(__file__).resolve().parents[1]
MARKER = "PrivateUploadOnly739581"


@contextmanager
def serve_program(program: Path):
    class QuietHandler(SimpleHTTPRequestHandler):
        def log_message(self, format, *args):
            pass

    server = ThreadingHTTPServer(
        ("127.0.0.1", 0), partial(QuietHandler, directory=str(program.resolve().parent))
    )
    thread = Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield f"http://127.0.0.1:{server.server_port}/{quote(program.name)}"
    finally:
        server.shutdown()
        thread.join()
        server.server_close()


def save_download(page, button, path: Path, expected_filename: str) -> None:
    with page.expect_download(timeout=90_000) as event:
        button.click()
    download = event.value
    assert download.suggested_filename == expected_filename, download.suggested_filename
    assert download.failure() is None
    download.save_as(path)


def check_zip(path: Path, expected_html: bytes) -> dict:
    with ZipFile(path) as archive:
        files = {name for name in archive.namelist() if not name.endswith("/")}
        assert "Tool-Manager.html" in files, files
        assert "사용법.txt" in files, files
        assert {"licenses/JSZIP-LICENSE.txt", "licenses/TESSERACT-LICENSE.txt"} <= files, files
        assert all(name in {"Tool-Manager.html", "사용법.txt"} or name.startswith("licenses/") for name in files), files
        assert archive.read("Tool-Manager.html") == expected_html
        manual = REPO / "web/사용법.txt"
        if manual.exists():
            assert archive.read("사용법.txt").decode("utf-8") == manual.read_text()
        for name in ("JSZIP-LICENSE.txt", "TESSERACT-LICENSE.txt"):
            local_license = REPO / "web/vendor" / name
            if local_license.exists():
                assert archive.read("licenses/" + name) == local_license.read_bytes()
        return {"files": sorted(files), "html_bytes": len(expected_html)}


def populate_private_state(page, temporary: Path) -> None:
    automatic = page.locator("#automatic-mode")
    if automatic.count():
        automatic.uncheck()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = MARKER
    sheet.append(["Cal Next", "Cal Last", "Serial", "Asset #", "Cal Interval", "Cal Cycle", "Label Photo"])
    sheet.append([datetime(2027, 1, 2), datetime(2026, 1, 2), "987654321", "South_Fixture_" + MARKER, 12, "MONTHS"])
    for key in ("A2", "B2"):
        sheet[key].number_format = "yyyy-mm-dd"
    workbook_path = temporary / (MARKER + ".xlsx")
    workbook.save(workbook_path)
    workbook.close()
    photo_path = temporary / (MARKER + ".png")
    Image.new("RGB", (100, 80), color=(222, 10, 200)).save(photo_path)
    page.locator("#workbook").set_input_files(workbook_path)
    page.wait_for_function("window.ToolApp && ToolApp.state.model && !ToolApp.state.busy")
    page.locator("#photos").set_input_files(photo_path)
    page.wait_for_function("ToolApp.state.photos.length === 1")
    assert MARKER in page.locator("#workbook-summary").inner_text()
    assert MARKER in page.locator("#photo-list").inner_text()
    page.locator("#tool-select").select_option(page.evaluate("ToolApp.state.model.tools[0].id"))
    assert MARKER in page.locator("#name").input_value()
    page.locator("#last").fill("2026-03-04")
    page.locator("#due").fill("2027-03-04")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--program", type=Path, default=Path("/workspace/outputs/Tool-Manager-Fixed.html"))
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    args = parser.parse_args()
    if not args.program.exists():
        fallback = REPO / "web/dist/Tool-Manager.html"
        if fallback.exists():
            args.program = fallback
    assert args.program.is_file(), args.program
    errors = []
    external_requests = []
    with tempfile.TemporaryDirectory(prefix="tool-download-test-") as directory:
        temporary = Path(directory)
        with sync_playwright() as playwright, serve_program(args.program) as url:
            browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
            context = browser.new_context(accept_downloads=True, locale="ko-KR")
            context.set_default_timeout(90_000)
            # Exercise Blob URLs in browsers without the optional File System
            # Access API; this also avoids requiring interactive permission.
            context.add_init_script("window.showSaveFilePicker = undefined;")

            def allow_local_only(route):
                if route.request.url.startswith("http://127.0.0.1:"):
                    route.continue_()
                else:
                    external_requests.append(route.request.url)
                    route.abort()

            context.route("**/*", allow_local_only)
            context.on("page", lambda page: page.on("pageerror", lambda error: errors.append(str(error))))
            page = context.new_page()
            page.goto(url, wait_until="load", timeout=90_000)
            page.wait_for_function("window.ToolDistribution && window.ToolApp")
            # The initial DOM is frozen before application startup. Parsing can
            # normalize source formatting, so use this pre-upload immutable
            # snapshot instead of comparing browser output to raw build bytes.
            pristine = page.evaluate("ToolDistribution.html()").encode("utf-8")
            assert page.locator("#save-program-html").is_visible()
            assert page.locator("#save-program-zip").is_visible()
            assert page.locator('script[src]').count() == 0, "Standalone HTML must embed its script dependencies"
            runtime = page.evaluate("""async () => {
                const data = await ToolRuntime.ready();
                return {worker_source_bytes:data.workerSource.length,
                    core_source_bytes:data.coreSource.length,
                    wasm_embedded:data.coreSource.includes('data:application/octet-stream;base64,AGFzbQ'),
                    korean_model_bytes:data.languages.kor.length,
                    english_model_bytes:data.languages.eng.length,
                    korean_gzip:Array.from(data.languages.kor.slice(0,2)),
                    english_gzip:Array.from(data.languages.eng.slice(0,2))};
            }""")
            assert runtime["worker_source_bytes"] > 1_000, runtime
            assert runtime["core_source_bytes"] > 100_000 and runtime["wasm_embedded"], runtime
            assert runtime["korean_model_bytes"] > 100_000 and runtime["english_model_bytes"] > 100_000, runtime
            assert runtime["korean_gzip"] == runtime["english_gzip"] == [0x1F, 0x8B], runtime
            populate_private_state(page, temporary)
            html_path = temporary / "Tool-Manager.html"
            zip_path = temporary / "Tool-Manager.zip"
            save_download(page, page.locator("#save-program-html"), html_path, "Tool-Manager.html")
            assert html_path.read_bytes() == pristine, "Runtime state leaked into HTML or source regeneration changed it"
            assert MARKER not in html_path.read_text()
            save_download(page, page.locator("#save-program-zip"), zip_path, "Tool-Manager.zip")
            package = check_zip(zip_path, pristine)
            page.close()

            # The self-downloaded HTML must remain standalone and save itself
            # again without retaining the prior person's uploaded records.
            with serve_program(html_path) as downloaded_url:
                reopened = context.new_page()
                reopened.goto(downloaded_url, wait_until="load", timeout=90_000)
                reopened.wait_for_function("window.ToolDistribution && window.ToolApp")
                assert reopened.evaluate("ToolDistribution.html()").encode("utf-8") == pristine, "Reopened self-source was not stable"
                assert reopened.evaluate("ToolApp.state.model === null && ToolApp.state.photos.length === 0")
                assert reopened.locator("#workbook").input_value() == ""
                assert MARKER not in reopened.locator("body").inner_text()
                second = temporary / "second"
                second.mkdir()
                save_download(reopened, reopened.locator("#save-program-html"), second / "Tool-Manager.html", "Tool-Manager.html")
                assert (second / "Tool-Manager.html").read_bytes() == pristine
                save_download(reopened, reopened.locator("#save-program-zip"), second / "Tool-Manager.zip", "Tool-Manager.zip")
                check_zip(second / "Tool-Manager.zip", pristine)
                reopened.close()

            preview = context.new_page()
            preview.set_content(f'<iframe id="preview" sandbox="allow-scripts allow-downloads" src="{url}"></iframe>')
            frame = preview.frame_locator("#preview")
            frame.locator("#save-program-html").wait_for()
            iframe_dir = temporary / "iframe"
            iframe_dir.mkdir()
            save_download(preview, frame.locator("#save-program-html"), iframe_dir / "Tool-Manager.html", "Tool-Manager.html")
            assert (iframe_dir / "Tool-Manager.html").read_bytes() == pristine
            save_download(preview, frame.locator("#save-program-zip"), iframe_dir / "Tool-Manager.zip", "Tool-Manager.zip")
            check_zip(iframe_dir / "Tool-Manager.zip", pristine)
            preview.close()

            # The platform may forbid browser saves by omitting allow-downloads.
            # Source reveal + select-all must remain usable in that preview.
            blocked = context.new_page()
            blocked.set_content(f'<iframe id="restricted" sandbox="allow-scripts" src="{url}"></iframe>')
            restricted = blocked.frame_locator("#restricted")
            restricted.locator("#show-save-source").click()
            source = restricted.locator("#program-source-text")
            assert source.is_visible()
            assert source.get_attribute("readonly") is not None
            assert source.input_value().encode("utf-8") == pristine
            restricted.locator("#copy-program-source").click()
            assert source.evaluate("el => el.selectionStart === 0 && el.selectionEnd === el.value.length"), "Copy fallback did not select all source"
            assert not errors, errors
            assert not external_requests, external_requests
            result = {
                "program": str(args.program),
                "top_level_html_and_zip_downloads": "passed",
                "downloaded_program_reopens_and_downloads_again": "passed",
                "uploaded_workbook_photo_and_edits_excluded": "passed (byte-identical pristine HTML)",
                "sandbox_allow_downloads_html_and_zip": "passed",
                "sandbox_without_downloads_source_selection_fallback": "passed",
                "package": package,
                "embedded_offline_runtime": runtime,
                "page_errors": errors,
                "external_network_requests": external_requests,
                "offline_ocr_accuracy": "not exercised by this download test",
                "platform_attachment_delivery": "outside browser test scope",
            }
            browser.close()
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
