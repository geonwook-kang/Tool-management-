"""No-input shared-workbook onboarding through the production mobile app.

Only Microsoft OAuth/Graph and the public same-origin settings response are
controlled HTTP fixtures. The app, PKCE, local storage, IndexedDB, workbook
parser, source adapter and queue guards run unchanged. No credentials, company
workbooks or private photos are used; this is Chromium phone emulation, not a
physical iPhone/Android or actual tenant consent test.
"""
from __future__ import annotations

import argparse
import base64
from datetime import datetime
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import tempfile
import time
from urllib.parse import parse_qs, urlencode, urlsplit

from openpyxl import Workbook
from PIL import Image
from playwright.sync_api import sync_playwright

from browser_test import serve_program
from graph_file_test import Cloud, TOKEN, TICKET
from m365_auth_test import CLIENT, TENANT, jwt


MOBILE = {
    "viewport": {"width": 390, "height": 844},
    "user_agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Mobile/15E148 Safari/604.1",
    "is_mobile": True, "has_touch": True, "locale": "ko-KR", "accept_downloads": True,
}


def workbook_bytes():
    memory = BytesIO()
    book = Workbook()
    sheet = book.active
    sheet.title = "Tools"
    sheet.append(["Cal Next", "Cal Last", "Serial", "Asset #", "Cal Interval", "Cal Cycle", "Label Photo", "SMAX Tool Photo"])
    sheet.append([datetime(2027, 3, 17), datetime(2026, 3, 17), "ABC12345", "South_Test_홍길동", 12, "MONTHS"])
    sheet["A2"].number_format = sheet["B2"].number_format = "yyyy-mm-dd"
    book.save(memory)
    return memory.getvalue()


class TargetCloud(Cloud):
    def __init__(self, original, suffix):
        super().__init__(original)
        self.target = {"driveId": "drive-" + suffix, "itemId": "item-" + suffix, "name": "team-" + suffix + ".xlsx"}

    def item(self):
        value = super().item()
        value.update({"id": self.target["itemId"], "name": self.target["name"],
                      "parentReference": {"driveId": self.target["driveId"]},
                      "@microsoft.graph.downloadUrl": "https://download.example.test/document/" + self.target["itemId"] + "?ticket=" + TICKET})
        return value


def config(cloud):
    return {"v": 1, "clientId": CLIENT, "tenantId": TENANT, "name": "직원 공용 툴 대장", "target": cloud.target.copy()}


def invite(url, value):
    encoded = base64.urlsafe_b64encode(json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode()).decode().rstrip("=")
    return url + "#team=" + encoded


def ready(page):
    page.wait_for_function("window.ToolApp&&window.ToolTeamConfig&&!ToolApp.state.busy&&!ToolApp.state.restoring", timeout=60000)


def reveal(page, selector):
    # Exercise the actual buttons while opening only their enclosing settings.
    for summary in page.locator(selector).evaluate("el=>{const ids=[];for(let p=el.parentElement;p;p=p.parentElement)if(p.tagName==='DETAILS'){if(!p.id)p.id='test-details-'+ids.length;ids.push(p.id);}return ids.reverse();}"):
        node = page.locator("#" + summary)
        if not node.evaluate("el=>el.open"):
            node.locator(":scope > summary").click()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--program", type=Path, default=Path("/workspace/outputs/Tool-Manager-Fixed.html"))
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    args = parser.parse_args()
    original = workbook_bytes()
    cases, errors, unexpected = {}, [], []
    with tempfile.TemporaryDirectory(prefix="tool-team-app-") as directory, serve_program(args.program) as url, sync_playwright() as playwright:
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
        local_host = urlsplit(url).netloc

        def new_case(published=None):
            clouds = {suffix: TargetCloud(original, suffix) for suffix in ["a", "b"]}
            traffic = {"authorize": [], "tokens": [], "settings": [], "dialogs": [], "published": published}
            context = browser.new_context(**MOBILE)
            context.add_init_script("""(()=>{
              window.teamCopies=[];window.teamShares=[];
              Object.defineProperty(navigator,'clipboard',{configurable:true,value:{writeText:async value=>{teamCopies.push({value,active:navigator.userActivation.isActive});}}});
              Object.defineProperty(navigator,'share',{configurable:true,value:async value=>{teamShares.push({...value,active:navigator.userActivation.isActive});}});
            })();""")

            def route(request_route):
                request = request_route.request
                parsed = urlsplit(request.url)
                if parsed.hostname == "login.microsoftonline.com":
                    if parsed.path.endswith("/authorize"):
                        values = parse_qs(parsed.query)
                        traffic["authorize"].append(values)
                        assert values["redirect_uri"] == [url]
                        assert values["client_id"] == [CLIENT] and values["code_challenge_method"] == ["S256"]
                        target = url + "?" + urlencode({"code": "team-controlled-code", "state": values["state"][0]})
                        return request_route.fulfill(status=302, headers={"location": target, "cache-control": "no-store"})
                    if parsed.path.endswith("/token"):
                        post = parse_qs(request.post_data or "")
                        traffic["tokens"].append(post)
                        values = traffic["authorize"][-1]
                        expected = base64.urlsafe_b64encode(sha256(post["code_verifier"][0].encode()).digest()).decode().rstrip("=")
                        assert values["code_challenge"] == [expected]
                        assert "client_secret" not in post and "offline_access" not in post["scope"][0]
                        claims = {"aud": CLIENT, "tid": TENANT, "iss": "https://login.microsoftonline.com/" + TENANT + "/v2.0",
                                  "nonce": values["nonce"][0], "exp": int(time.time()) + 3600,
                                  "name": "Test Staff", "preferred_username": "staff@example.invalid"}
                        data = {"token_type": "Bearer", "expires_in": 3600, "access_token": TOKEN,
                                "scope": "openid profile Files.ReadWrite.All", "id_token": jwt(claims)}
                        return request_route.fulfill(status=200, content_type="application/json", headers={"access-control-allow-origin": "*", "cache-control": "no-store"}, body=json.dumps(data))
                    unexpected.append(request.url)
                    return request_route.abort()
                if parsed.hostname in {"graph.microsoft.com", "download.example.test"}:
                    # Team invitations identify an exact file. They must never
                    # need to resolve a stored sharing URL or guess a target.
                    assert "/shares/" not in parsed.path, request.url
                    for cloud in clouds.values():
                        target = cloud.target
                        if parsed.hostname == "graph.microsoft.com" and parsed.path.startswith("/v1.0/drives/" + target["driveId"] + "/items/" + target["itemId"]) or parsed.hostname == "download.example.test" and parsed.path == "/document/" + target["itemId"]:
                            return cloud.route(request_route)
                    unexpected.append(request.url)
                    return request_route.abort()
                if parsed.netloc == local_host and parsed.path.endswith("/team-settings.json"):
                    traffic["settings"].append({"url": request.url, "headers": request.headers})
                    value = traffic["published"] or {"enabled": False}
                    return request_route.fulfill(status=200, content_type="application/json", headers={"cache-control": "no-store"}, body=json.dumps(value))
                if parsed.netloc != local_host:
                    unexpected.append(request.url)
                    return request_route.abort()
                return request_route.continue_()

            context.route("**/*", route)
            page = context.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))
            page.on("dialog", lambda dialog: (traffic["dialogs"].append(dialog.message), dialog.accept()))
            return context, page, clouds, traffic

        def assert_no_login_or_file_access(clouds, traffic):
            assert not traffic["authorize"] and not traffic["tokens"]
            assert all(not cloud.calls and not cloud.puts for cloud in clouds.values())

        def login(page, traffic, suffix="a"):
            before = len(traffic["tokens"])
            page.locator("#m365-signin").click()
            page.wait_for_function("expected=>window.ToolM365Auth&&window.ToolApp&&window.ToolGraphFile&&ToolM365Auth.inspect().signedIn&&ToolApp.state.source.connected&&ToolApp.state.model&&!ToolApp.state.busy&&ToolGraphFile.inspect().target?.itemId===expected", arg="item-" + suffix, timeout=60000)
            assert len(traffic["tokens"]) == before + 1
            assert page.evaluate("ToolSources.mode()") == "microsoft365"
            assert page.evaluate("ToolApp.state.model.tools[0].serial") == "ABC12345"
            assert not page.locator("#connection-settings").evaluate("el=>el.open")
            assert "team-" + suffix + ".xlsx" in page.locator("#team-summary").inner_text()
            assert "code=" not in page.url and "state=" not in page.url and "#team=" not in page.url

        # A new employee follows the manager's link and only clicks company
        # login. No app ID, tenant, path or file picker is filled by the test.
        context, page, clouds, traffic = new_case()
        selected = config(clouds["a"])
        page.goto(invite(url, selected))
        ready(page)
        assert_no_login_or_file_access(clouds, traffic)
        assert page.evaluate("ToolTeamConfig.inspect().configuration") == selected
        assert page.locator("#m365-client-id").input_value() == CLIENT
        assert page.locator("#m365-tenant-id").input_value() == TENANT
        assert page.locator("#cloud-workbook-url").input_value() == ""
        login(page, traffic)
        assert not clouds["a"].puts and not clouds["b"].calls
        cases["fresh_invitation_explicit_pkce_login_no_manual_settings"] = True
        print("Team app: invited employee connected without manual settings", flush=True)

        # Share/copy retain user activation; exported deployment settings
        # contain public app/file identifiers, no bearer or private share URL.
        reveal(page, "#copy-team-link")
        page.locator("#copy-team-link").click()
        page.wait_for_function("teamCopies.length===1")
        copied = page.evaluate("teamCopies[0]")
        assert copied["active"] and copied["value"].startswith(url + "#team=")
        raw = copied["value"].split("#team=", 1)[1]
        decoded = json.loads(base64.urlsafe_b64decode(raw + "=" * (-len(raw) % 4)))
        assert decoded["target"] == selected["target"] and decoded["clientId"] == CLIENT
        reveal(page, "#share-team-link")
        page.locator("#share-team-link").click()
        page.wait_for_function("teamShares.length===1")
        assert page.evaluate("teamShares[0].active")
        assert page.evaluate("teamShares[0].url") == copied["value"]
        reveal(page, "#download-team-settings")
        with page.expect_download() as event:
            page.locator("#download-team-settings").click()
        downloaded = Path(directory) / "team-settings.json"
        event.value.save_as(downloaded)
        published = json.loads(downloaded.read_text())
        assert published["enabled"] is True and published["target"] == selected["target"]
        assert published["clientId"] == CLIENT and published["tenantId"] == TENANT
        serialized = json.dumps(published)
        assert TOKEN not in serialized and TICKET not in serialized and "shareUrl" not in serialized and "redirectUri" not in serialized
        cases["manager_copy_share_and_target_only_deployment_export"] = True
        print("Team app: manager share/copy/export verified", flush=True)

        # Browser restart loses in-memory tokens, remembers the exact team,
        # and never redirects or accesses a company file before explicit login.
        before_reads = len(clouds["a"].calls)
        before_login = len(traffic["authorize"])
        page.reload()
        ready(page)
        assert not page.evaluate("ToolM365Auth.inspect().signedIn")
        assert page.evaluate("ToolTeamConfig.inspect().configuration.target.itemId") == "item-a"
        assert len(clouds["a"].calls) == before_reads and len(traffic["authorize"]) == before_login
        login(page, traffic)
        page.locator("#m365-signout").click()
        page.wait_for_function("!ToolM365Auth.inspect().signedIn&&!ToolApp.state.source.connected")
        assert page.evaluate("ToolTeamConfig.inspect().configuration.target.itemId") == "item-a"
        login(page, traffic)
        cases["restart_and_logout_keep_team_but_require_explicit_login"] = True
        print("Team app: restart/logout retains team settings", flush=True)
        stored = page.evaluate("JSON.stringify({local:{...localStorage},session:{...sessionStorage}})")
        assert TOKEN not in stored and "controlled-signature" not in stored
        assert all(not cloud.puts for cloud in clouds.values())
        context.close()

        # The ordinary app URL on a completely fresh phone receives the
        # published target. This covers PWA/Safari storage separation without
        # assuming browser settings migrate into the installed application.
        context, page, clouds, traffic = new_case(published)
        page.goto(url)
        ready(page)
        assert_no_login_or_file_access(clouds, traffic)
        assert page.evaluate("ToolTeamConfig.inspect().configuration.target.itemId") == "item-a"
        assert traffic["settings"]
        login(page, traffic)
        assert not clouds["b"].calls
        cases["ordinary_app_url_cold_phone_uses_published_default"] = True
        print("Team app: fresh phone uses published target at ordinary URL", flush=True)
        context.close()

        # The invitation is consumed before an OAuth round trip. A published
        # default A must not take precedence over the explicitly selected B
        # when the callback reconstructs its (even empty) pending context.
        context, page, clouds, traffic = new_case(published)
        page.goto(invite(url, config(clouds["b"])))
        ready(page)
        assert_no_login_or_file_access(clouds, traffic)
        assert page.evaluate("ToolTeamConfig.inspect().configuration.target.itemId") == "item-b"
        login(page, traffic, "b")
        assert not clouds["a"].calls and clouds["b"].calls
        assert page.evaluate("ToolResume.peek()") is None
        assert all(not cloud.puts for cloud in clouds.values())
        cases["explicit_invitation_survives_oauth_over_different_published_default"] = True
        print("Team app: explicit invitation survives OAuth over other published default", flush=True)
        context.close()

        # A different invitation cannot reuse stale target A or consume A's
        # pending original. Blocked sign-in is acceptable while a mismatched
        # queue is present; any connection must target B and leave A untouched.
        context, page, clouds, traffic = new_case()
        page.goto(invite(url, config(clouds["a"])))
        ready(page)
        login(page, traffic)
        queued = page.evaluate("""async()=>{
          await ToolResume.save({context:{mode:'microsoft365',driveId:'drive-a',itemId:'item-a',shareUrl:'',uncertain:false},photos:[{id:'pending-a',file:new File([new Uint8Array([255,216,255,217])],'pending-a.jpg',{type:'image/jpeg'}),draft:{},saved:null}]});
          return ToolResume.peek();
        }""")
        before_a = len(clouds["a"].calls)
        page.evaluate("""()=>{
          window.staleRestoreCalls=0;
          window.ToolGraphFile={...ToolGraphFile,restore:async()=>{staleRestoreCalls++;throw Error('A shared selected-target slot must not override an explicit team target');}};
        }""")
        page.goto(invite(url, config(clouds["b"])))
        page.wait_for_function("window.ToolTeamConfig&&ToolTeamConfig.inspect().configuration?.target?.itemId==='item-b'", timeout=60000)
        ready(page)
        assert len(clouds["a"].calls) == before_a
        assert page.evaluate("ToolTeamConfig.inspect().configuration.target.itemId") == "item-b"
        if not page.evaluate("ToolM365Auth.inspect().signedIn"):
            page.locator("#m365-signin").click()
            ready(page)
        pending = page.evaluate("ToolResume.peek()")
        assert pending and pending["count"] == 1 and pending["context"]["itemId"] == "item-a" and pending["createdAt"] == queued["createdAt"]
        target = page.evaluate("ToolGraphFile.inspect().target")
        assert not target or target["itemId"] == "item-b"
        assert page.evaluate("staleRestoreCalls") == 0
        if page.evaluate("ToolM365Auth.inspect().signedIn"):
            assert page.evaluate("ToolApp.state.source.connected") and target["itemId"] == "item-b"
        assert len(clouds["a"].calls) == before_a and all(not cloud.puts for cloud in clouds.values())
        assert not page.evaluate("ToolApp.state.photos.length")
        cases["different_invitation_does_not_read_stale_target_or_consume_other_queue"] = True
        cases["explicit_team_never_restores_shared_selected_target_slot"] = True
        print("Team app: different invitation preserves previous target queue", flush=True)
        context.close()

        # Opening another invitation in an existing app must not discard a
        # freshly selected photo or silently change its destination. This
        # checks the workflow guard, so OCR is explicitly disabled by its UI.
        context, page, clouds, traffic = new_case()
        page.goto(invite(url, config(clouds["a"])))
        ready(page)
        login(page, traffic)
        page.locator("#automatic-mode").uncheck()
        image_bytes = BytesIO()
        Image.new("RGB", (32, 24), (130, 160, 190)).save(image_bytes, format="JPEG")
        page.locator("#photos").set_input_files({"name": "unsaved-original.jpg", "mimeType": "image/jpeg", "buffer": image_bytes.getvalue()})
        page.wait_for_function("ToolApp.state.photos.length===1&&!ToolApp.state.busy")
        before_reads = len(clouds["a"].calls)
        page.goto(invite(url, config(clouds["b"])))
        page.wait_for_function("!location.hash.includes('team=')")
        ready(page)
        assert page.evaluate("ToolTeamConfig.inspect().configuration.target.itemId") == "item-a"
        assert page.evaluate("ToolGraphFile.inspect().target.itemId") == "item-a"
        assert page.evaluate("ToolApp.state.photos.length===1&&ToolApp.state.photos[0].file.name==='unsaved-original.jpg'&&!ToolApp.state.photos[0].saved")
        assert len(clouds["a"].calls) == before_reads and not clouds["b"].calls
        assert all(not cloud.puts for cloud in clouds.values())
        assert "저장 대상은 변경하지 않았습니다" in page.locator("#status").inner_text()
        cases["unsaved_photo_blocks_same_document_invitation_without_data_loss"] = True
        print("Team app: unsaved photo prevents silent target change", flush=True)
        context.close()

        # Invalid invitations cannot mutate the remembered team or initiate a
        # company login/read. A fresh invalid invitation is handled visibly.
        context, page, clouds, traffic = new_case()
        page.goto(url + "#team=%%%")
        ready(page)
        assert "team=" not in page.url
        assert not page.evaluate("ToolTeamConfig.inspect().configured")
        assert_no_login_or_file_access(clouds, traffic)
        assert page.locator("#status").inner_text()
        cases["malformed_invitation_scrubbed_no_auth_or_file_mutation"] = True
        context.close()
        browser.close()
    assert not errors and not unexpected, {"page_errors": errors, "unexpected_requests": unexpected}
    print(json.dumps({"program_sha256": sha256(args.program.read_bytes()).hexdigest(), "cases": cases, "case_count": len(cases),
                      "page_errors": errors, "unexpected_requests": unexpected, "company_data_used": False,
                      "real_microsoft_tenant_used": False, "physical_phone_tested": False}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
