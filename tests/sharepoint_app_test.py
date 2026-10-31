"""SharePoint-session phone uploads through the real app in an iframe srcdoc.

The parent page, session cookie and same-origin SharePoint REST service are
controlled fixtures. The production app, adapter, IndexedDB backups, OOXML and
OCR run unchanged. Private input files are read only and remain outside Git.
Chromium phone emulation does not assert actual tenant access, physical phone
installation or Safari engine support.
"""
from __future__ import annotations

import argparse
from contextlib import contextmanager
from hashlib import sha256
from html import escape
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
from pathlib import Path
import subprocess
import tempfile
from threading import Thread
import time
from urllib.parse import parse_qs, unquote, urlsplit

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

from multi_owner_ocr_test import target_reference, verify_increment
from team_app_test import workbook_bytes


SESSION = "TEST_SHAREPOINT_SESSION_NOT_AN_APP_ID"
DIGEST = "TEST_SHAREPOINT_FORM_DIGEST_NOT_A_BEARER"
UNIQUE = "11111111-1111-4111-8111-111111111111"
SITE = "/sites/Tools"
FILE = SITE + "/Shared Documents/shared-tools.xlsx"
MOBILE = {
    "viewport": {"width": 390, "height": 844},
    "user_agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Mobile/15E148 Safari/604.1",
    "is_mobile": True, "has_touch": True, "locale": "ko-KR", "accept_downloads": True,
}
ANDROID = {**MOBILE, "viewport": {"width": 412, "height": 915},
           "user_agent": "Mozilla/5.0 (Linux; Android 14; Mobile) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Mobile Safari/537.36"}
REPO = Path(__file__).resolve().parents[1]


class SharePoint:
    """A session-authenticated file with server checkout and draft publication."""

    def __init__(self, original):
        self.original, self.data = original, original
        self.revision = 1
        self.mode = "normal"
        self.calls, self.puts = [], []
        self.site_url = ""
        self.checked_out = False
        self.checkout_snapshot = None
        self.published_data = original
        self.checkout_attempts, self.checkouts, self.checkins, self.undos = [], [], [], []
        self.mutations = []
        self.content_verified_during_checkout = False
        self.checkout_put_start = 0

    @property
    def tag(self):
        return '"' + UNIQUE + "," + str(self.revision) + '"'

    def metadata(self):
        return {"Name": "shared-tools.xlsx", "Length": str(len(self.data)),
                "UniqueId": UNIQUE, "ServerRelativeUrl": FILE,
                "ETag": self.tag, "Exists": True, "CheckOutType": 0 if self.checked_out else 2}

    def route(self, route):
        request = route.request
        parsed = urlsplit(request.url)
        request_headers = request.all_headers()
        assert parsed.path.startswith(SITE + "/_api/"), request.url
        assert request_headers.get("authorization") is None, "The adapter must not request a Graph bearer"
        assert "sp-session=" + SESSION in request_headers.get("cookie", ""), "Existing same-origin session cookie was not sent"
        call = {"path": parsed.path, "method": request.method,
                "cookie_present": True, "if_match": request_headers.get("if-match"),
                "digest_present": request_headers.get("x-requestdigest") == DIGEST}
        self.calls.append(call)
        headers = {"cache-control": "no-store", "content-type": "application/json;odata=nometadata"}
        if self.mode in {"read401", "read403"}:
            return route.fulfill(status=int(self.mode[-3:]), headers=headers, body='{"error":{"message":"Read access denied"}}')
        if parsed.path.endswith("/_api/contextinfo"):
            assert request.method == "POST"
            return route.fulfill(status=200, headers=headers,
                                 body=json.dumps({"FormDigestValue": DIGEST,
                                                  "FormDigestTimeoutSeconds": 1800,
                                                  "WebFullUrl": self.site_url}))
        assert "GetFileByServerRelativePath(decodedUrl=@file)" in parsed.path, request.url
        values = parse_qs(parsed.query)
        assert values.get("@file") == ["'" + FILE + "'"], values
        operation = unquote(parsed.path).rsplit("/", 1)[-1]
        if request.method == "POST" and operation == "CheckOut()":
            assert request_headers.get("x-requestdigest") == DIGEST
            self.checkout_attempts.append({"tag": self.tag, "already_checked_out": self.checked_out})
            self.mutations.append("checkout-attempt")
            if self.mode == "checkout409" or self.checked_out:
                # Another session wins after the caller's last metadata read.
                # The failed caller must never undo that other checkout.
                self.checked_out = True
                return route.fulfill(status=409, headers=headers, body='{"error":{"message":"File already checked out"}}')
            before_tag = self.tag
            self.checkout_snapshot = (self.data, self.revision)
            self.checked_out = True
            self.content_verified_during_checkout = False
            self.checkout_put_start = len(self.puts)
            self.revision += 1  # A checkout can itself change a file's ETag.
            self.checkouts.append({"before_tag": before_tag, "locked_tag": self.tag})
            self.mutations.append("checkout-acquired")
            return route.fulfill(status=204, headers={"cache-control": "no-store", "etag": self.tag})
        if request.method == "POST" and operation == "UndoCheckOut()":
            assert request_headers.get("x-requestdigest") == DIGEST
            assert self.checked_out and self.checkout_snapshot is not None, "Must not undo an existing or another session's checkout"
            self.data, self.revision = self.checkout_snapshot
            self.checkout_snapshot = None
            self.checked_out = False
            self.undos.append({"restored_tag": self.tag})
            self.mutations.append("undo-own-checkout")
            return route.fulfill(status=204, headers={"cache-control": "no-store", "etag": self.tag})
        if request.method == "POST" and operation.startswith("CheckIn("):
            assert "checkintype=1" in operation and "Tool Manager calibration update" in operation, operation
            assert request_headers.get("x-requestdigest") == DIGEST
            assert self.checked_out and self.checkout_snapshot is not None
            assert self.content_verified_during_checkout, "The written bytes must be verified while the server checkout is still held"
            state = request.frame.evaluate("({saved:ToolApp.state.photos.some(p=>p.saved),saveCount:ToolApp.state.source.saveCount})")
            assert not state["saved"], "The UI must not mark a photo saved before server check-in"
            self.checkins.append({"ui_before_checkin": state, "draft_hash": sha256(self.data).hexdigest(),
                                  "published_hash_before_checkin": sha256(self.published_data).hexdigest()})
            self.published_data = self.data
            self.checked_out = False
            self.checkout_snapshot = None
            self.revision += 1
            self.mutations.append("checkin-published")
            return route.fulfill(status=204, headers={"cache-control": "no-store", "etag": self.tag})
        if request.method == "GET" and not parsed.path.endswith("/$value"):
            return route.fulfill(status=200, headers=headers, body=json.dumps(self.metadata()))
        if request.method == "GET":
            if self.checked_out and len(self.puts) > self.checkout_put_start and self.puts[-1].get("committed"):
                self.content_verified_during_checkout = True
            return route.fulfill(status=200, headers={"cache-control": "no-store",
                                 "content-type": "application/octet-stream", "etag": self.tag}, body=self.data)
        assert request.method == "POST" and parsed.path.endswith("/$value"), request.url
        assert request_headers.get("x-http-method", "").upper() == "PUT"
        assert request_headers.get("x-requestdigest") == DIGEST
        assert self.checked_out and self.checkout_snapshot is not None, "Binary replacement requires a newly acquired server checkout"
        incoming = request.post_data_buffer
        self.puts.append({"if_match": request_headers.get("if-match"),
                          "sha256": sha256(incoming).hexdigest(), "digest_present": True,
                          "locked_tag": self.tag, "committed": False})
        self.mutations.append("binary-put")
        self.content_verified_during_checkout = False
        # Do not model stream If-Match as an undocumented concurrency guarantee.
        # The real exclusion boundary is CheckOut, held through verification.
        if self.mode == "write412":
            return route.fulfill(status=412, headers=headers, body='{"error":{"message":"Version conflict"}}')
        if self.mode == "write403":
            return route.fulfill(status=403, headers=headers, body='{"error":{"message":"Write denied"}}')
        self.data = incoming
        self.revision += 1
        self.puts[-1]["committed"] = True
        if self.mode == "commit-abort":
            # The server accepted the draft, but the response never reaches
            # the browser. The app must keep its checkout journal and must not
            # claim success, retry the upload or undo possibly committed data.
            return route.abort("failed")
        return route.fulfill(status=204, headers={"cache-control": "no-store", "etag": self.tag})


@contextmanager
def serve_parent(program):
    """Actual srcdoc origin inheritance, matching the SPFx host boundary."""
    source = program.read_text(encoding="utf-8")
    documents = {}

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass

        def do_GET(self):
            parsed = urlsplit(self.path)
            if parsed.path == "/standalone.html":
                data = source.encode()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Cache-Control", "no-store")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return
            if parsed.path != SITE + "/SitePages/ToolManager.aspx":
                self.send_error(404)
                return
            wrong_origin = parse_qs(parsed.query).get("wrong-origin") == ["1"]
            data = documents[wrong_origin]
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    origin = "http://127.0.0.1:" + str(server.server_port)
    settings = [{"siteUrl": origin + SITE, "fileUrl": origin + FILE,
                 "applicationSha256": sha256(program.read_bytes()).hexdigest()},
                {"siteUrl": "https://other.example.invalid/sites/Tools",
                 "fileUrl": "https://other.example.invalid" + FILE,
                 "applicationSha256": sha256(program.read_bytes()).hexdigest()}]
    # Invoke the exact checked-in SPFx helper, not a mirror of its injection.
    node = r"""
      const fs=require('fs'),ts=require('./sharepoint/node_modules/typescript');
      const source=fs.readFileSync('./sharepoint/src/hostFrame.ts','utf8');
      const js=ts.transpileModule(source,{compilerOptions:{module:ts.ModuleKind.CommonJS,target:ts.ScriptTarget.ES2019}}).outputText;
      const exports={};new Function('exports',js)(exports);
      const input=JSON.parse(fs.readFileSync(0,'utf8'));
      process.stdout.write(JSON.stringify(input.settings.map(context=>exports.hostedDocument(input.html,context))));
    """
    children = json.loads(subprocess.run(["node", "-e", node], cwd=REPO,
                          input=json.dumps({"html": source, "settings": settings}),
                          text=True, capture_output=True, check=True).stdout)
    listener = '<script>window.frameMessages=[];addEventListener("message",event=>{const frame=document.getElementById("tool-frame");if(!frame||event.source!==frame.contentWindow||event.origin!==location.origin)return;const data=event.data||{};if(data.type!=="TOOL_SHAREPOINT_HEIGHT"&&data.type!=="TOOL_SHAREPOINT_READY")return;frameMessages.push(data);if(data.type==="TOOL_SHAREPOINT_HEIGHT"&&typeof data.height==="number"&&Number.isFinite(data.height))frame.style.height=Math.round(Math.max(650,Math.min(30000,data.height)))+"px";});</script>'
    for wrong_origin, child in zip([False, True], children):
        parent = '<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"></head><body style="margin:0">' + listener + '<iframe title="Tool Manager" id="tool-frame" sandbox="allow-scripts allow-same-origin allow-forms allow-downloads allow-modals" allow="camera; clipboard-write; web-share" style="width:100%;height:2600px;border:0;display:block" srcdoc="' + escape(child, quote=True) + '"></iframe></body></html>'
        documents[wrong_origin] = parent.encode()
    thread = Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield origin, origin + SITE + "/SitePages/ToolManager.aspx"
    finally:
        server.shutdown()
        thread.join()
        server.server_close()


def ready(frame):
    frame.wait_for_function("window.ToolApp&&window.ToolSharePointFile&&window.ToolSources&&!ToolApp.state.busy&&!ToolApp.state.restoring", timeout=60000)


def download_bytes(page, frame, expression, destination):
    with page.expect_download() as event:
        frame.evaluate("""async expression=>{
          const value=await eval('(async()=>('+expression+'))()');
          const link=document.createElement('a');link.href=URL.createObjectURL(new Blob([value]));
          link.download='shared-workbook.xlsx';document.body.appendChild(link);link.click();link.remove();
        }""", expression)
    event.value.save_as(destination)


def expected_reference(args):
    reference = target_reference(args.workbook, args.serial)
    book = load_workbook(args.workbook)
    sheet = book[reference["sheet"]]
    asset = str(sheet.cell(reference["row"], reference["headers"]["Asset #"]).value)
    expected = {"name": asset.rsplit("_", 1)[-1].strip(), "serial": args.serial,
                "last": args.last or sheet[reference["last_cell"]].value.strftime("%Y-%m-%d"),
                "due": args.due or sheet[reference["due_cell"]].value.strftime("%Y-%m-%d")}
    book.close()
    return reference, expected


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--program", type=Path, default=Path("/workspace/outputs/Tool-Manager-SharePoint.html"))
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--photo", type=Path, required=True)
    parser.add_argument("--serial", required=True)
    parser.add_argument("--last")
    parser.add_argument("--due")
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    parser.add_argument("--diagnostics", type=Path, default=Path("/tmp/tool-sharepoint-app-test.json"))
    args = parser.parse_args()
    if args.workbook is None:
        args.workbook = next(Path("/workspace/attachments").rglob("*.xlsx"), None)
    for label, file in [("program", args.program), ("workbook", args.workbook), ("photo", args.photo)]:
        if file is None or not file.is_file():
            parser.error("Supply --" + label + " with an existing file")
    reference, expected = expected_reference(args)
    original, photograph = args.workbook.read_bytes(), args.photo.read_bytes()
    before = {"workbook": sha256(original).hexdigest(), "photo": sha256(photograph).hexdigest()}
    report = {"program_sha256": sha256(args.program.read_bytes()).hexdigest(), "cases": {},
              "real_company_tenant_used": False, "physical_phone_or_safari_engine_tested": False,
              "ocr_responses_replaced": False, "source_adapter_replaced": False,
              "authentication_method": "Actual inherited-origin browser cookie; controlled SharePoint REST",
              "app_registration_or_bearer_provider_used": False, "expected_values_passed_to_recognizer": False}
    errors, unexpected = [], []
    with tempfile.TemporaryDirectory(prefix="tool-sharepoint-app-") as directory, serve_parent(args.program) as (origin, url), sync_playwright() as playwright:
        temporary = Path(directory)
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])

        def case(data, mode="normal", wrong_origin=False, device=MOBILE):
            cloud = SharePoint(data)
            cloud.site_url = origin + SITE
            cloud.mode = mode
            context = browser.new_context(**device)
            context.add_cookies([{"name": "sp-session", "value": SESSION, "url": origin, "httpOnly": True, "sameSite": "Lax"}])

            def route(request_route):
                parsed = urlsplit(request_route.request.url)
                if parsed.netloc == urlsplit(origin).netloc and "/_api/" in parsed.path:
                    return cloud.route(request_route)
                if parsed.netloc != urlsplit(origin).netloc:
                    unexpected.append(request_route.request.url)
                    return request_route.abort()
                return request_route.continue_()

            context.route("**/*", route)
            page = context.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))
            page.goto(url + ("?wrong-origin=1" if wrong_origin else ""))
            frame = page.frame_locator("#tool-frame")
            page.wait_for_function("document.getElementById('tool-frame').contentWindow.ToolApp")
            child = page.frames[1]
            ready(child)
            return context, page, child, cloud

        context, page, child, cloud = case(workbook_bytes())
        assert child.evaluate("location.href") == "about:srcdoc"
        assert child.evaluate("window.origin") == origin
        assert child.evaluate("!Object.getOwnPropertyDescriptor(window,'ToolSharePointContext').writable&&!Object.getOwnPropertyDescriptor(window,'ToolSharePointContext').configurable")
        assert child.evaluate("ToolSources.mode()") == "sharepoint"
        assert child.evaluate("ToolApp.state.source.connected&&!!ToolApp.state.model")
        assert child.locator("#m365-client-id").is_hidden() and child.locator("#m365-tenant-id").is_hidden()
        assert child.locator("#m365-signin").is_hidden()
        assert not child.locator("#take-photo").is_disabled()
        assert child.evaluate("document.documentElement.scrollWidth<=innerWidth")
        assert "shared-tools.xlsx" in child.locator("#connection-status").inner_text()
        assert cloud.calls and all(call["cookie_present"] for call in cloud.calls) and not cloud.puts
        page.wait_for_function("frameMessages.some(message=>message.type==='TOOL_SHAREPOINT_READY')&&frameMessages.some(message=>message.type==='TOOL_SHAREPOINT_HEIGHT')")
        page.wait_for_function("document.querySelector('#tool-frame').clientHeight>=document.querySelector('#tool-frame').contentDocument.documentElement.scrollHeight")
        report["cases"]["srcdoc_same_origin_cookie_auto_connect_no_ids_or_manual_file_selection"] = True
        report["cases"]["phone_layout_camera_controls_without_horizontal_overflow"] = True
        report["cases"]["production_readonly_host_injection_and_ready_height_protocol"] = True
        with page.expect_download() as exported:
            child.evaluate("()=>{document.getElementById('save-program-html').click()}")
        export_path = temporary / "pristine-export.html"
        exported.value.save_as(export_path)
        text = export_path.read_text()
        assert "Object.defineProperty(window,\"ToolSharePointContext\"" not in text
        assert origin not in text and FILE not in text and SESSION not in text and DIGEST not in text
        assert "data-sharepoint-shell" not in text.split("<script id=\"runtime-assets\"", 1)[0]
        report["cases"]["exported_program_excludes_injected_company_context_and_session"] = True
        context.close()

        context, page, child, cloud = case(workbook_bytes(), device=ANDROID)
        assert child.evaluate("ToolSources.mode()==='sharepoint'&&ToolApp.state.source.connected")
        assert child.evaluate("document.documentElement.scrollWidth<=innerWidth")
        assert not child.locator("#choose-photos").is_disabled()
        assert not cloud.puts
        report["cases"]["android_gallery_layout_and_no_input_connection"] = True
        context.close()

        context, page, child, cloud = case(workbook_bytes(), "read401")
        assert not child.evaluate("ToolApp.state.source.connected")
        assert cloud.calls and not cloud.puts
        message = child.locator("#status").inner_text() + child.locator("#connection-status").inner_text()
        assert "로그인" in message, message
        assert "관리자 승인" not in message and "IT 승인" not in message
        report["cases"]["actual_401_requests_sharepoint_login_without_oauth_app_setup"] = True
        context.close()

        context, page, child, cloud = case(workbook_bytes(), "read403")
        assert not child.evaluate("ToolApp.state.source.connected")
        assert cloud.calls and not cloud.puts
        message = child.locator("#status").inner_text() + child.locator("#connection-status").inner_text()
        assert "403" in message or "권한" in message, message
        assert "관리자 승인" not in message and "IT 승인" not in message
        report["cases"]["access_result_uses_actual_403_without_assumed_approval_gate"] = True
        context.close()

        context, page, child, cloud = case(workbook_bytes(), wrong_origin=True)
        assert not child.evaluate("ToolApp.state.source.connected")
        assert not cloud.calls and not cloud.puts
        assert "SharePoint" in child.locator("#status").inner_text()
        report["cases"]["external_origin_cannot_reuse_sharepoint_session"] = True
        context.close()

        # A public mobile page remains useful for explicit local-copy work.
        # It cannot claim a SharePoint connection merely because an XLSX was
        # manually imported, but capture/gallery must become usable.
        context = browser.new_context(**MOBILE)

        def public_route(route):
            parsed = urlsplit(route.request.url)
            if parsed.netloc != urlsplit(origin).netloc or "/_api/" in parsed.path:
                unexpected.append(route.request.url)
                return route.abort()
            return route.continue_()

        context.route("**/*", public_route)
        page = context.new_page()
        page.on("pageerror", lambda error: errors.append(str(error)))
        page.goto(origin + "/standalone.html")
        ready(page)
        assert page.evaluate("ToolSources.mode()==='sharepoint'&&!ToolSources.inspect().hosted&&!ToolApp.state.model")
        assert page.locator("#take-photo").is_disabled()
        page.locator("#workbook").set_input_files({"name": "manual-copy.xlsx", "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "buffer": workbook_bytes()})
        page.wait_for_function("ToolApp.state.model&&!ToolApp.state.busy")
        assert page.evaluate("ToolApp.state.model.tools[0].serial==='ABC12345'&&!ToolApp.state.source.connected&&!ToolSharePointFile.inspect().connected")
        assert not page.locator("#take-photo").is_disabled() and not page.locator("#choose-photos").is_disabled()
        report["cases"]["public_mobile_manual_import_enables_photos_without_claiming_cloud_save"] = True
        context.close()

        # Actual supplied camera JPEG, unchanged production OCR, target matching,
        # date cells and photo insertion, followed by conditional SharePoint save.
        context, page, child, cloud = case(original)
        assert child.evaluate("ToolApp.state.source.connected")
        started = time.monotonic()
        with page.expect_file_chooser() as chooser:
            child.locator("#take-photo").click()
        chooser.value.set_files({"name": "camera-picture.jpg", "mimeType": "image/jpeg", "buffer": photograph})
        child.wait_for_function("ToolApp.state.photos.length===1&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=180000)
        result = child.evaluate("""async()=>{
          const photo=ToolApp.state.photos[0],assessment=photo.assessment;
          const hash=async bytes=>Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))).map(x=>x.toString(16).padStart(2,'0')).join('');
          return {assessment,saved:!!photo.saved,approval:photo.approval,sourceMode:ToolSources.mode(),saveCount:ToolApp.state.source.saveCount,
            sourceError:ToolApp.state.source.error,inputs:['name','serial','last','due'].map(id=>document.getElementById(id).value),
            photoHash:await hash(photo.image.photoBytes),preparedHash:ToolApp.state.preparedResult?await hash(ToolApp.state.preparedResult.bytes):null,
            status:document.getElementById('status').textContent};
        }""")
        result["seconds"] = round(time.monotonic() - started, 2)
        assert result["saved"] and result["approval"] == "automatic" and result["saveCount"] == 1 and not result["sourceError"], result
        assert result["inputs"] == [expected[key] for key in ["name", "serial", "last", "due"]], result
        assert result["sourceMode"] == "sharepoint"
        assert len(cloud.puts) == 1 and cloud.puts[0]["if_match"] == cloud.checkouts[0]["locked_tag"]
        assert cloud.puts[0]["digest_present"] and sha256(cloud.data).hexdigest() == result["preparedHash"]
        assert len(cloud.checkouts) == len(cloud.checkins) == 1 and not cloud.checked_out
        assert cloud.published_data == cloud.data
        assert cloud.checkins[0]["ui_before_checkin"] == {"saved": False, "saveCount": 0}
        assert cloud.checkins[0]["published_hash_before_checkin"] == before["workbook"]
        assert cloud.checkins[0]["draft_hash"] == result["preparedHash"]
        assert cloud.mutations == ["checkout-attempt", "checkout-acquired", "binary-put", "checkin-published"]
        result["server_checkout_held_until_verified_and_checked_in"] = True
        updated = temporary / "actual-photo-updated.xlsx"
        updated.write_bytes(cloud.data)
        result["workbook_verification"] = verify_increment(args.workbook, updated, reference, expected, result["photoHash"])
        backup = temporary / "actual-original-backup.xlsx"
        download_bytes(page, child, "(await ToolSharePointFile.takeBackup()).bytes", backup)
        assert backup.read_bytes() == original
        result["original_backup_byte_exact"] = True
        report["cases"]["actual_photo_ocr_to_conditional_sharepoint_excel_save"] = result

        saved_before_conflict = cloud.data
        # A pre-existing checkout with unchanged file bytes AND ETag must be
        # read freshly even if the previous workbook body is cached.
        cloud.checked_out = True
        attempts_before, puts_before = len(cloud.checkout_attempts), len(cloud.puts)
        child.evaluate("async()=>{ToolApp.state.photos[0].saved=null;await ToolApp.prepareResult();}")
        held = child.evaluate("({saved:!!ToolApp.state.photos[0].saved,error:ToolApp.state.source.error})")
        assert held["error"] and not held["saved"] and cloud.data == saved_before_conflict, held
        assert len(cloud.checkout_attempts) == attempts_before and len(cloud.puts) == puts_before
        assert cloud.checked_out and not cloud.undos
        report["cases"]["preexisting_checkout_with_unchanged_etag_blocks_binary_write"] = held
        cloud.checked_out = False  # The other user's checkout is released externally.

        cloud.mode = "checkout409"
        child.evaluate("async()=>{await ToolApp.prepareResult();}")
        concurrent = child.evaluate("({saved:!!ToolApp.state.photos[0].saved,error:ToolApp.state.source.error})")
        assert concurrent["error"] and not concurrent["saved"] and cloud.data == saved_before_conflict, concurrent
        assert len(cloud.checkout_attempts) == attempts_before + 1 and len(cloud.puts) == puts_before
        assert cloud.checked_out and not cloud.undos, "Rejected checkout must not release another session's lock"
        report["cases"]["concurrent_checkout_rejection_preserves_file_and_other_lock"] = concurrent
        cloud.checked_out = False

        cloud.mode = "write412"
        child.evaluate("async()=>{await ToolApp.prepareResult();}")
        conflict = child.evaluate("({saved:!!ToolApp.state.photos[0].saved,error:ToolApp.state.source.error,status:document.getElementById('status').textContent})")
        assert conflict["error"] and not conflict["saved"] and cloud.data == saved_before_conflict, conflict
        assert len(cloud.puts) == 2 and cloud.puts[1]["if_match"] == cloud.checkouts[1]["locked_tag"]
        assert len(cloud.undos) == 1 and not cloud.checked_out and cloud.published_data == saved_before_conflict
        report["cases"]["412_preserves_server_bytes_and_is_never_reported_as_saved"] = conflict
        cloud.mode = "write403"
        child.evaluate("async()=>{await ToolApp.prepareResult();}")
        denied = child.evaluate("({saved:!!ToolApp.state.photos[0].saved,error:ToolApp.state.source.error,status:document.getElementById('status').textContent})")
        assert denied["error"] and not denied["saved"] and cloud.data == saved_before_conflict, denied
        assert "권한" in denied["error"] and "관리자 승인" not in denied["error"] and "IT 승인" not in denied["error"]
        assert len(cloud.puts) == 3 and cloud.puts[-1]["if_match"] == cloud.checkouts[-1]["locked_tag"]
        assert len(cloud.undos) == 2 and not cloud.checked_out and cloud.published_data == saved_before_conflict
        report["cases"]["actual_write_403_retains_photo_and_never_reports_saved"] = denied

        # Use the real visible reconnect action and original photo again.
        # The next stream replacement commits, then loses its HTTP response.
        # This creates the production durable journal rather than seeding IDB.
        cloud.mode = "commit-abort"
        child.locator("#sharepoint-retry").click()
        child.wait_for_function("ToolApp.state.photos.length===1&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=180000)
        uncertain = child.evaluate("({photoRetained:ToolApp.state.photos.length===1,saved:!!ToolApp.state.photos[0].saved,error:ToolApp.state.source.error,pending:ToolSharePointFile.inspect().pendingCheckout})")
        assert uncertain["photoRetained"] and uncertain["pending"] and not uncertain["saved"] and uncertain["error"], uncertain
        assert cloud.checked_out and len(cloud.puts) == 4 and len(cloud.undos) == 2 and len(cloud.checkins) == 1
        assert cloud.published_data == saved_before_conflict
        assert child.locator("#take-photo").is_disabled() and child.locator("#choose-photos").is_disabled()
        operation_counts = (len(cloud.puts), len(cloud.undos), len(cloud.checkins), len(cloud.checkout_attempts))
        dialogs = []
        page.on("dialog", lambda dialog: (dialogs.append(dialog.type), dialog.accept()))
        cloud.mode = "normal"
        page.reload()
        page.wait_for_function("document.getElementById('tool-frame').contentWindow.ToolApp")
        child = page.frames[1]
        ready(child)
        recovery = child.evaluate("({connected:ToolApp.state.source.connected,error:ToolApp.state.source.error,pending:ToolSharePointFile.inspect().pendingCheckout,status:document.getElementById('status').textContent,note:document.getElementById('sharepoint-host-note').textContent})")
        assert recovery["connected"] and recovery["pending"] and recovery["error"], recovery
        assert "확인" in recovery["status"] and "사진만" not in recovery["status"], recovery
        assert "사진만" not in recovery["note"] and "확인" in recovery["note"], recovery
        assert child.locator("#take-photo").is_disabled() and child.locator("#choose-photos").is_disabled()
        assert (len(cloud.puts), len(cloud.undos), len(cloud.checkins), len(cloud.checkout_attempts)) == operation_counts
        assert child.locator("#retry-connected-save").is_visible()
        child.locator("#retry-connected-save").click()
        child.wait_for_function("!ToolApp.state.busy&&!ToolApp.state.preparing")
        assert child.evaluate("ToolSharePointFile.inspect().pendingCheckout&&!!ToolApp.state.source.error")
        assert (len(cloud.puts), len(cloud.undos), len(cloud.checkins), len(cloud.checkout_attempts)) == operation_counts

        # Model the owner checking in this exact draft using SharePoint itself.
        # The app may clear recovery only after a fresh server read sees None2.
        cloud.published_data = cloud.data
        cloud.checked_out = False
        cloud.checkout_snapshot = None
        cloud.revision += 1
        cloud.mutations.append("manual-checkin-by-owner")
        child.locator("#retry-connected-save").click()
        child.wait_for_function("!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolSharePointFile.inspect().pendingCheckout&&!ToolApp.state.source.error")
        assert not child.locator("#take-photo").is_disabled() and not child.locator("#choose-photos").is_disabled()
        assert (len(cloud.puts), len(cloud.undos), len(cloud.checkins), len(cloud.checkout_attempts)) == operation_counts
        uncertain.update({"reloadPendingStatus": recovery, "beforeunloadDialogs": dialogs,
                          "noAutomaticWriteUndoOrCheckinOnReload": True,
                          "explicitServerResolutionThenReadClearsRecovery": True})
        report["cases"]["uncertain_committed_draft_survives_reload_and_blocks_until_server_resolution"] = uncertain
        storage = child.evaluate("JSON.stringify({local:{...localStorage},session:{...sessionStorage}})")
        assert SESSION not in storage and DIGEST not in storage
        assert child.evaluate("document.cookie") == "", "HttpOnly session must never be read by app scripts"
        context.close()
        browser.close()
    assert not errors and not unexpected, (errors, unexpected)
    assert sha256(args.workbook.read_bytes()).hexdigest() == before["workbook"]
    assert sha256(args.photo.read_bytes()).hexdigest() == before["photo"]
    report.update({"page_errors": errors, "unexpected_requests": unexpected, "private_input_files_unchanged": True})
    args.diagnostics.write_text(json.dumps(report, ensure_ascii=False, indent=2))
    print(json.dumps({"program_sha256": report["program_sha256"], "cases": list(report["cases"]),
                      "page_errors": errors, "unexpected_requests": unexpected,
                      "diagnostics": str(args.diagnostics)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
