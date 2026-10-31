"""Mobile layout, camera/gallery, real OCR and Graph-backed Excel UI tests.

The HTTP Graph service and its bearer provider are controlled fixtures. The
production Graph adapter, workbook writer, backup storage, mobile UI and OCR
run unchanged. Camera input is a file chooser; iPhone and Android viewport/UA
emulation is not a physical phone or Safari engine test. Private inputs are
provided at runtime and remain outside Git.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
import json
from pathlib import Path
import tempfile
import time
from urllib.parse import urlsplit, parse_qs, urlencode

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

from browser_test import REPO, serve_program, download_expression
from graph_file_test import Cloud, SHARE, TOKEN
from multi_owner_ocr_test import target_reference, verify_increment
from m365_auth_test import jwt


CLIENT = "00000000-0000-4000-8000-000000000001"
TENANT = "11111111-1111-4111-8111-111111111111"
DEVICES = {
    "iphone": {"viewport": {"width": 390, "height": 844},
               "user_agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Mobile/15E148 Safari/604.1"},
    "android": {"viewport": {"width": 412, "height": 915},
                "user_agent": "Mozilla/5.0 (Linux; Android 14; Mobile) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Mobile Safari/537.36"}
}
SHARE_API = r"""() => {
  window.shareCalls=[];
  Object.defineProperty(navigator,'canShare',{configurable:true,value:options=>Array.isArray(options.files)&&options.files.length===1});
  Object.defineProperty(navigator,'share',{configurable:true,value:async options=>{
    const entry={active:navigator.userActivation.isActive,names:options.files.map(f=>f.name),types:options.files.map(f=>f.type)};
    shareCalls.push(entry);
    entry.hashes=await Promise.all(options.files.map(async file=>Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',await file.arrayBuffer()))).map(n=>n.toString(16).padStart(2,'0')).join('')));
  }});
}"""


def arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--program", type=Path, default=Path("/tmp/tool-mobile-integration.html"))
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--photo", type=Path, required=True)
    parser.add_argument("--serial", required=True, help="Expected original workbook serial; used only by Python assertions")
    parser.add_argument("--last", help="Expected ISO photo date; defaults to the selected workbook row")
    parser.add_argument("--due", help="Expected ISO photo due date; defaults to the selected workbook row")
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    parser.add_argument("--diagnostics", type=Path, default=Path("/tmp/tool-mobile-app-test.json"))
    parser.add_argument("--oauth-only", action="store_true", help="Run only the full PKCE redirect/queued-photo resume case")
    args = parser.parse_args()
    if args.workbook is None:
        args.workbook = next(Path("/workspace/attachments").rglob("*.xlsx"), None)
    for label, file in [("program", args.program), ("workbook", args.workbook), ("photo", args.photo)]:
        if file is None or not file.is_file():
            parser.error("Supply --" + label + " with an existing file")
    return args


def expected_reference(args):
    reference = target_reference(args.workbook, args.serial)
    book = load_workbook(args.workbook)
    sheet = book[reference["sheet"]]
    asset = str(sheet.cell(reference["row"], reference["headers"]["Asset #"]).value)
    value = {"name": asset.rsplit("_", 1)[-1].strip(), "serial": args.serial,
             "last": args.last or sheet[reference["last_cell"]].value.strftime("%Y-%m-%d"),
             "due": args.due or sheet[reference["due_cell"]].value.strftime("%Y-%m-%d")}
    book.close()
    return reference, value


def ready(page):
    page.wait_for_function("window.ToolApp&&window.ToolSources&&!ToolApp.state.busy&&!ToolApp.state.restoring")


def reveal_settings(page, selector):
    """Open the visible settings disclosure before changing its controls."""
    for item in page.locator(selector).evaluate("el=>{const ids=[];for(let p=el.parentElement;p;p=p.parentElement)if(p.tagName==='DETAILS'&&p.id)ids.push(p.id);return ids.reverse();}"):
        details = page.locator("#" + item)
        if not details.evaluate("el=>el.open"):
            details.locator(":scope > summary").click()


def layout(page):
    value = page.evaluate("""()=>({
      viewport:innerWidth,document:document.documentElement.scrollWidth,
      mode:document.querySelector('#source-mode').value,
      capture:document.querySelector('#camera').getAttribute('capture'),
      cameraAccept:document.querySelector('#camera').accept,
      fontSize:parseFloat(getComputedStyle(document.querySelector('#m365-client-id')).fontSize),
      controls:['take-photo','choose-photos'].map(id=>{let r=document.getElementById(id).getBoundingClientRect();return {width:r.width,height:r.height};})
    })""")
    assert value["document"] <= value["viewport"], value
    assert value["mode"] == "microsoft365", value
    assert value["capture"] == "environment" and value["cameraAccept"] == "image/*"
    assert value["fontSize"] >= 16 and all(control["height"] >= 44 and control["width"] >= 44 for control in value["controls"])
    return value


def main():
    args = arguments()
    reference, expected = expected_reference(args)
    original = args.workbook.read_bytes()
    original_hash, photo_hash = sha256(original).hexdigest(), sha256(args.photo.read_bytes()).hexdigest()
    report = {"program_sha256": sha256(args.program.read_bytes()).hexdigest(), "cases": {},
              "real_microsoft_tenant_used": False, "authentication_stub": "getAccessToken in basic UI cases; full redirect case uses the real auth module with controlled issuer HTTP",
              "ocr_responses_replaced": False, "graph_adapter_replaced": False,
              "physical_phone_or_safari_engine_tested": False, "expected_values_passed_to_recognizer": False}
    errors, unexpected = [], []
    with tempfile.TemporaryDirectory(prefix="tool-mobile-app-") as directory, serve_program(args.program) as url, sync_playwright() as playwright:
        temporary = Path(directory)
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])

        def context_for(device, cloud, oauth=None):
            context = browser.new_context(**DEVICES[device], is_mobile=True, has_touch=True, device_scale_factor=3, locale="ko-KR", accept_downloads=True)
            context.add_init_script("(" + SHARE_API + ")()")
            page = context.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))

            def route(request_route):
                parsed = urlsplit(request_route.request.url)
                if oauth is not None and parsed.hostname == "login.microsoftonline.com":
                    if parsed.path.endswith("/authorize"):
                        values = parse_qs(parsed.query)
                        oauth["authorize"].append(values)
                        assert values["redirect_uri"] == [url]
                        assert values["code_challenge_method"] == ["S256"]
                        target = url + "?" + urlencode({"code": "controlled-ui-authorization-code", "state": values["state"][0]})
                        return request_route.fulfill(status=302, headers={"location": target, "cache-control": "no-store"})
                    if parsed.path.endswith("/token"):
                        post = parse_qs(request_route.request.post_data or "")
                        oauth["tokens"].append(post)
                        values = oauth["authorize"][-1]
                        import base64
                        challenge = base64.urlsafe_b64encode(sha256(post["code_verifier"][0].encode()).digest()).decode().rstrip("=")
                        assert values["code_challenge"] == [challenge]
                        assert "client_secret" not in post and "offline_access" not in post["scope"][0]
                        claims = {"aud": CLIENT, "tid": TENANT, "iss": "https://login.microsoftonline.com/" + TENANT + "/v2.0",
                                  "nonce": values["nonce"][0], "exp": int(time.time()) + 3600,
                                  "name": "Mobile Test Engineer", "preferred_username": "engineer@example.invalid"}
                        data = {"token_type": "Bearer", "expires_in": 3600, "access_token": TOKEN,
                                "scope": "openid profile Files.ReadWrite.All", "id_token": jwt(claims)}
                        return request_route.fulfill(status=200, content_type="application/json", headers={"access-control-allow-origin": "*", "cache-control": "no-store"}, body=json.dumps(data))
                    unexpected.append(request_route.request.url)
                    return request_route.abort()
                if parsed.hostname in {"graph.microsoft.com", "download.example.test"}:
                    return cloud.route(request_route)
                if parsed.netloc == urlsplit(url).netloc and parsed.path.endswith("/team-settings.json"):
                    return request_route.fulfill(status=200, content_type="application/json", headers={"cache-control": "no-store"}, body='{"enabled":false}')
                if parsed.netloc != urlsplit(url).netloc:
                    unexpected.append(request_route.request.url)
                    return request_route.abort()
                return request_route.continue_()

            context.route("**/*", route)
            page.goto(url)
            ready(page)
            # These cases exercise the optional external Microsoft 365 source.
            # The default unconfigured phone now opens the local workbook runner.
            reveal_settings(page, "#source-mode")
            page.locator("#source-mode").select_option("microsoft365")
            ready(page)
            return context, page

        def connect_cloud(page):
            # This is the only authentication seam. All Graph HTTP, source
            # metadata, download/PUT, concurrency and backup logic stay real.
            page.evaluate("token=>{window.ToolM365Auth={...ToolM365Auth,getAccessToken:async()=>token};}", TOKEN)
            reveal_settings(page, "#m365-client-id")
            page.locator("#m365-client-id").fill(CLIENT)
            page.locator("#m365-tenant-id").fill(TENANT)
            page.locator("#cloud-workbook-url").fill(SHARE)
            page.locator("#connect-onedrive").click()
            page.wait_for_function("ToolApp.state.source.connected&&ToolApp.state.model&&!ToolApp.state.busy")
            assert page.evaluate("ToolSources.mode()") == "microsoft365"
            assert "회사 공유 엑셀" in page.locator("#connection-status").inner_text()

        def photo_result(page):
            return page.evaluate("""async()=>{
              const photo=ToolApp.state.photos.at(-1),a=photo.assessment;
              return {assessment:a,saved:!!photo.saved,approval:photo.approval,
                rawResult:{serial:photo.result?.serial,last:photo.result?.last,due:photo.result?.due,failed:photo.result?.failed},
                sourceMode:ToolSources.mode(),saveCount:ToolApp.state.source.saveCount,sourceError:ToolApp.state.source.error,
                inputs:['name','serial','last','due'].map(id=>document.getElementById(id).value),
                photoHash:Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',photo.image.photoBytes))).map(n=>n.toString(16).padStart(2,'0')).join(''),
                preparedHash:Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',ToolApp.state.preparedResult.bytes))).map(n=>n.toString(16).padStart(2,'0')).join(''),
                status:document.querySelector('#status').textContent};
            }""")

        def assert_values(result):
            assert result["saved"] and result["approval"] == "automatic" and result["saveCount"] == 1 and not result["sourceError"], result
            assert result["inputs"] == [expected[key] for key in ["name", "serial", "last", "due"]], result
            assert [result["assessment"].get(key) for key in ["name", "serial", "last", "due"]] == result["inputs"], result
            assert not result["rawResult"]["failed"], result

        for device in ([] if args.oauth_only else ["iphone", "android"]):
            cloud = Cloud(original)
            context, page = context_for(device, cloud)
            report["cases"][device + "_layout"] = layout(page)
            connect_cloud(page)
            if device == "iphone":
                page.evaluate("()=>{window.bitmapFallbacks=0;window.createImageBitmap=async()=>{bitmapFallbacks++;throw new Error('controlled bitmap decoder unavailable')};}")
            before = time.monotonic()
            with page.expect_file_chooser() as picker:
                page.locator("#take-photo" if device == "iphone" else "#choose-photos").click()
            selection = {"name": "camera-capture" if device == "iphone" else "gallery-picture.jpg", "mimeType": "image/jpeg", "buffer": args.photo.read_bytes()}
            picker.value.set_files(selection)
            page.wait_for_function("ToolApp.state.photos.length===1&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=180000)
            result = photo_result(page)
            result["seconds"] = round(time.monotonic() - before, 2)
            assert_values(result)
            assert result["sourceMode"] == "microsoft365" and "회사 공유 원본" in result["status"], result
            assert len(cloud.puts) == 1 and cloud.puts[0]["if_match"] == '"version-1"'
            assert sha256(cloud.data).hexdigest() == result["preparedHash"]
            saved = temporary / (device + "-cloud-updated.xlsx")
            saved.write_bytes(cloud.data)
            result["workbook_verification"] = verify_increment(args.workbook, saved, reference, expected, result["photoHash"])
            backup = temporary / (device + "-cloud-backup.xlsx")
            download_expression(page, "(await ToolGraphFile.takeBackup()).bytes", backup)
            assert backup.read_bytes() == original
            result["original_backup_byte_exact"] = True
            if device == "iphone":
                assert page.evaluate("bitmapFallbacks") >= 2, "Both preparation and actual OCR should use real Image fallback"
                assert page.evaluate("ToolApp.state.photos[0].file.name") == "camera-capture"
                result["extensionless_jpeg_and_decoder_fallback"] = True
            page.locator("#share-workbook").click()
            page.wait_for_function("shareCalls.length===1&&shareCalls[0].hashes")
            shared = page.evaluate("shareCalls[0]")
            assert shared["active"] and shared["hashes"] == [result["preparedHash"]]
            assert shared["names"][0].endswith(".xlsx") and shared["types"] == ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]
            result["native_share_user_gesture_and_exact_bytes"] = True
            assert page.evaluate("document.documentElement.scrollWidth<=innerWidth")
            report["cases"][device + "_actual_ocr_cloud_save"] = result

            if device == "android":
                prior = cloud.data
                cloud.mode = "put412"
                page.evaluate("async()=>{ToolApp.state.photos[0].saved=null;await ToolApp.prepareResult()}")
                held = page.evaluate("({error:ToolApp.state.source.error,saved:!!ToolApp.state.photos[0].saved,status:document.querySelector('#status').textContent,connection:document.querySelector('#connection-status').textContent})")
                assert held["error"] and not held["saved"] and cloud.data == prior, held
                assert "확인" in held["connection"] and "다른 사용자" in held["error"], held
                assert len(cloud.puts) == 2 and cloud.puts[-1]["if_match"] == '"version-2"'
                report["cases"]["cloud_412_not_reported_as_saved"] = held
                count = page.evaluate("ToolApp.state.photos.length")
                page.locator("#camera").set_input_files({"name": "unsupported.heic", "mimeType": "image/heic", "buffer": b"not-a-decodable-heic-image"})
                page.wait_for_function("!ToolApp.state.busy")
                assert page.evaluate("ToolApp.state.photos.length") == count
                message = page.locator("#status").inner_text()
                assert "JPEG" in message and "열지 못" in message, message
                report["cases"]["unsupported_image_has_actionable_error"] = True
            context.close()

        if not args.oauth_only:
            cloud = Cloud(original)
            context, page = context_for("android", cloud)
            reveal_settings(page, "#source-mode")
            page.locator("#source-mode").select_option("device")
            page.wait_for_function("ToolSources.mode()==='device'&&!ToolApp.state.busy&&!ToolApp.state.restoring")
            with page.expect_file_chooser() as picker:
                page.locator("#connect-onedrive").click()
            picker.value.set_files(args.workbook)
            page.wait_for_function("ToolApp.state.source.connected&&ToolApp.state.model&&!ToolApp.state.busy")
            assert "이 기기" in page.locator("#connection-status").inner_text()
            with page.expect_file_chooser() as picker:
                page.locator("#choose-photos").click()
            picker.value.set_files(args.photo)
            page.wait_for_function("ToolApp.state.photos.length===1&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=180000)
            result = photo_result(page)
            assert_values(result)
            assert result["sourceMode"] == "device" and "회사 공유 원본은 변경하지 않았습니다" in result["status"], result
            assert not cloud.calls and not cloud.puts and cloud.data == original
            saved = temporary / "android-device-updated.xlsx"
            download_expression(page, "(await ToolMobileFile.read()).bytes", saved)
            result["workbook_verification"] = verify_increment(args.workbook, saved, reference, expected, result["photoHash"])
            result["no_graph_requests_or_cloud_changes"] = True
            report["cases"]["device_mode_local_only_actual_ocr"] = result
            context.close()

        # Start with a queued camera photo before any workbook or login. The
        # real auth module navigates through PKCE and returns to this document;
        # root initialization must reconnect, consume its durable queue, read
        # the photo with the real worker and save the same controlled cloud item.
        cloud, oauth = Cloud(original), {"authorize": [], "tokens": []}
        context, page = context_for("iphone", cloud, oauth)
        dialogs = []
        page.on("dialog", lambda dialog: (dialogs.append(dialog.type), dialog.accept()))
        with page.expect_file_chooser() as picker:
            page.locator("#take-photo").click()
        picker.value.set_files({"name": "before-login.jpg", "mimeType": "image/jpeg", "buffer": args.photo.read_bytes()})
        page.wait_for_function("ToolApp.state.photos.length===1&&!ToolApp.state.busy")
        assert not page.evaluate("!!ToolApp.state.model") and not cloud.calls
        reveal_settings(page, "#m365-client-id")
        page.locator("#m365-client-id").fill(CLIENT)
        page.locator("#m365-tenant-id").fill(TENANT)
        page.locator("#cloud-workbook-url").fill(SHARE)
        with page.expect_request("https://login.microsoftonline.com/**/authorize?*"):
            page.locator("#m365-signin").click()
        page.wait_for_function("window.ToolApp&&ToolM365Auth.inspect().signedIn&&ToolApp.state.photos.length===1&&ToolApp.state.source.saveCount===1&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=180000)
        result = photo_result(page)
        assert_values(result)
        assert len(oauth["authorize"]) == len(oauth["tokens"]) == len(cloud.puts) == 1
        assert "code=" not in page.url and "state=" not in page.url
        assert "로그인됨" in page.locator("#m365-status").inner_text()
        assert page.evaluate("sessionStorage.getItem('tool-manager-m365-pkce-v1')") is None
        assert page.evaluate("ToolResume.peek()") is None
        storage = page.evaluate("JSON.stringify({local:{...localStorage},session:{...sessionStorage}})")
        assert TOKEN not in storage and "controlled-signature" not in storage
        saved = temporary / "oauth-resumed-cloud-updated.xlsx"
        saved.write_bytes(cloud.data)
        result["workbook_verification"] = verify_increment(args.workbook, saved, reference, expected, result["photoHash"])
        backup = temporary / "oauth-resumed-original-backup.xlsx"
        download_expression(page, "(await ToolGraphFile.takeBackup()).bytes", backup)
        assert backup.read_bytes() == original
        result.update({"actual_auth_module_replaced": False, "real_pkce_s256_verified": True,
                       "callback_removed_from_url": True, "transient_queue_and_verifier_cleared": True,
                       "tokens_not_persisted": True, "beforeunload_dialogs": dialogs,
                       "company_tenant_used": False})
        report["cases"]["full_pkce_redirect_queued_photo_resume_to_cloud"] = result
        context.close()
        browser.close()
    assert not errors and not unexpected, (errors, unexpected)
    assert sha256(args.workbook.read_bytes()).hexdigest() == original_hash
    assert sha256(args.photo.read_bytes()).hexdigest() == photo_hash
    report.update({"page_errors": errors, "unexpected_requests": unexpected, "original_photo_and_workbook_unchanged": True})
    args.diagnostics.write_text(json.dumps(report, ensure_ascii=False, indent=2))
    print(json.dumps({"program_sha256": report["program_sha256"], "cases": list(report["cases"]),
                      "actual_ocr_cloud_saves": 1 if args.oauth_only else 3, "actual_ocr_device_saves": 0 if args.oauth_only else 1,
                      "page_errors": errors, "unexpected_requests": unexpected, "diagnostics": str(args.diagnostics)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
