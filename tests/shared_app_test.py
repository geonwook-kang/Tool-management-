"""Owner setup and stable staff links through the actual shared HTML app.

The production PKCE/authentication, workspace, source adapter, merge helper,
OOXML writer and UI run unchanged. A controlled HTTP issuer and Graph service
stand in for Microsoft. Only the concurrency case replaces OCR responses with
fixture text; --photo additionally exercises the unmodified real OCR engine.
Chromium phone emulation does not verify physical phones, Safari or a real
company tenant. Private workbook/photo inputs stay outside the repository.
"""
from __future__ import annotations

import argparse
import base64
from datetime import datetime
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import re
import tempfile
import time
from urllib.parse import parse_qs, unquote, urlencode, urlsplit
from zipfile import ZipFile

from playwright.sync_api import sync_playwright

from browser_test import serve_program
from m365_auth_test import CLIENT, TENANT, jwt
from multi_owner_ocr_test import target_reference, verify_increment
from shared_merge_test import fixture, check_combined
from smax_workbook_test import drawing_info


OWNER = "22222222-2222-4222-8222-222222222222"
MEMBER_A = "33333333-3333-4333-8333-333333333333"
MEMBER_B = "44444444-4444-4444-8444-444444444444"
SHARE_A = "https://company.example.invalid/sites/tools/original-a.xlsx?web=1"
SHARE_B = "https://company.example.invalid/sites/tools/original-b.xlsx?web=1"
CONFIG_SHARE = "https://company.example.invalid/:u:/settings-read-only"
DOWNLOAD_HOST = "download.shared.example.test"
PHONE = {"viewport": {"width": 390, "height": 844}, "is_mobile": True,
         "has_touch": True, "locale": "ko-KR", "accept_downloads": True,
         "user_agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Mobile/15E148 Safari/604.1"}


class Cases(dict):
    def __setitem__(self, key, value):
        super().__setitem__(key, value)
        print("PASS " + key, flush=True)


class SharedCloud:
    """One central store shared by every independently authenticated browser."""

    def __init__(self, original):
        self.files = {
            key: {"data": original, "revision": 1, "driveId": "master-drive-" + key,
                  "itemId": "master-item-" + key, "name": "master-" + key + ".xlsx"}
            for key in ["a", "b"]
        }
        self.settings = None
        self.calls, self.puts, self.config_puts, self.conflicts = [], [], [], []
        self.block_first_pair = False
        self.pending = []
        self.held_conflict = None
        self.first_receipt_downloaded = False
        self.shared_to = set()

    def metadata(self, key):
        file = self.settings if key == "settings" else self.files[key]
        assert file is not None
        return {"id": file["itemId"], "name": file["name"], "size": len(file["data"]),
                "eTag": '"' + key + "-" + str(file["revision"]) + '"',
                "file": {"mimeType": "application/json" if key == "settings" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
                "parentReference": {"driveId": file["driveId"]},
                "@microsoft.graph.downloadUrl": "https://" + DOWNLOAD_HOST + "/" + key + "?ticket=controlled-proof"}

    @staticmethod
    def respond(route, value, status=200):
        return route.fulfill(status=status, content_type="application/json",
                             headers={"access-control-allow-origin": "*", "cache-control": "no-store"},
                             body=json.dumps(value))

    def commit_workbook(self, route, key, actor):
        request = route.request
        before = self.metadata(key)
        if request.headers.get("if-match") != before["eTag"]:
            self.conflicts.append({"actor": actor, "key": key, "if_match": request.headers.get("if-match")})
            return self.respond(route, {"error": {"code": "preconditionFailed"}}, 412)
        data = request.post_data_buffer
        assert data is not None and data[:4] == b"PK\x03\x04"
        self.puts.append({"actor": actor, "key": key, "if_match": before["eTag"], "hash": sha256(data).hexdigest()})
        self.files[key]["data"] = data
        self.files[key]["revision"] += 1
        return self.respond(route, self.metadata(key))

    def route(self, route, actor):
        request, parsed = route.request, urlsplit(route.request.url)
        if request.method == "OPTIONS":
            return route.fulfill(status=204, headers={"access-control-allow-origin": "*",
                "access-control-allow-methods": "GET,PUT,POST,OPTIONS", "access-control-allow-headers": "authorization,content-type,if-match,if-none-match,prefer"})
        if parsed.hostname == DOWNLOAD_HOST:
            assert "authorization" not in request.headers and "cookie" not in request.headers
            key = parsed.path.strip("/")
            file = self.settings if key == "settings" else self.files[key]
            if self.held_conflict and key == self.held_conflict[1] and actor == self.held_conflict[3]:
                self.first_receipt_downloaded = True
            return route.fulfill(status=200, headers={"access-control-allow-origin": "*", "cache-control": "no-store"}, body=file["data"])
        assert parsed.hostname == "graph.microsoft.com"
        assert request.headers.get("authorization") == "Bearer shared-controlled-" + actor
        path = unquote(parsed.path.removeprefix("/v1.0"))
        self.calls.append({"actor": actor, "method": request.method, "path": path})
        if path == "/me/drive":
            return self.respond(route, {"id": "owner-drive" if actor == OWNER else "personal-" + actor,
                                       "owner": {"user": {"id": actor}}, "driveType": "business"})
        if path.startswith("/me/drive/root:/ToolManagerShared-") and path.endswith(".json:/content"):
            assert actor == OWNER and request.method == "PUT" and request.headers.get("if-none-match") == "*"
            value = json.loads(request.post_data)
            assert set(value) == {"v", "workspaceId", "ownerObjectId", "workbook"} and value["ownerObjectId"] == OWNER
            self.settings = {"data": request.post_data_buffer, "revision": 1, "driveId": "owner-drive",
                             "itemId": "workspace-settings", "name": path.split("root:/", 1)[1].removesuffix(":/content")}
            return self.respond(route, self.metadata("settings"), 201)
        if path.startswith("/shares/u!") and "/driveItem" in path:
            code = path.split("/shares/u!", 1)[1].split("/driveItem", 1)[0]
            share = base64.urlsafe_b64decode(code + "=" * (-len(code) % 4)).decode()
            key = {SHARE_A: "a", SHARE_B: "b", CONFIG_SHARE: "settings"}[share]
            if key == "settings":
                self.shared_to.add(actor)
            return self.respond(route, self.metadata(key))
        for key in ["settings", "a", "b"]:
            file = self.settings if key == "settings" else self.files[key]
            if file is None:
                continue
            root = "/drives/" + file["driveId"] + "/items/" + file["itemId"]
            if not path.startswith(root):
                continue
            tail = path[len(root):]
            if key == "settings" and actor != OWNER and actor not in self.shared_to:
                return self.respond(route, {"error": {"code": "accessDenied"}}, 403)
            if tail == "/createLink":
                assert actor == OWNER and json.loads(request.post_data) == {"type": "view", "scope": "organization"}
                return self.respond(route, {"roles": ["read"], "link": {"type": "view", "scope": "organization", "webUrl": CONFIG_SHARE}})
            if tail == "/permissions":
                return self.respond(route, {"value": [{"roles": ["owner"], "grantedToV2": {"user": {"id": OWNER}}},
                                                          {"roles": ["read"], "link": {"type": "view", "scope": "organization"}}]})
            if tail == "" and request.method == "GET":
                result = self.respond(route, self.metadata(key))
                # Release the rejected concurrent write after the first
                # writer's byte verification, so this test isolates a real
                # 412 retry rather than an uncertain post-commit read race.
                if self.held_conflict and self.first_receipt_downloaded and key == self.held_conflict[1] and actor == self.held_conflict[3]:
                    held, held_key, held_actor, _ = self.held_conflict
                    self.held_conflict = None
                    self.commit_workbook(held, held_key, held_actor)
                return result
            if tail == "/content" and request.method == "PUT":
                if key == "settings":
                    self.config_puts.append({"actor": actor, "if_match": request.headers.get("if-match")})
                    if actor != OWNER:
                        return self.respond(route, {"error": {"code": "accessDenied"}}, 403)
                    if request.headers.get("if-match") != self.metadata(key)["eTag"]:
                        return self.respond(route, {"error": {"code": "preconditionFailed"}}, 412)
                    file["data"], file["revision"] = request.post_data_buffer, file["revision"] + 1
                    return self.respond(route, self.metadata(key))
                if self.block_first_pair:
                    self.pending.append((route, key, actor))
                    if len(self.pending) < 2:
                        return
                    self.block_first_pair = False
                    first, second = self.pending
                    self.pending = []
                    self.commit_workbook(*first)
                    self.held_conflict = (*second, first[2])
                    self.first_receipt_downloaded = False
                    return
                return self.commit_workbook(route, key, actor)
        raise AssertionError((request.method, path))


def reveal(page, selector):
    for identity in page.locator(selector).evaluate("el=>{const ids=[];for(let p=el.parentElement;p;p=p.parentElement)if(p.tagName==='DETAILS'){if(!p.id)p.id='test-detail-'+ids.length;ids.push(p.id);}return ids.reverse();}"):
        details = page.locator("#" + identity)
        if not details.evaluate("el=>el.open"):
            details.locator(":scope > summary").click()


def ready(page):
    page.wait_for_function("window.ToolApp&&window.ToolSharedWorkspace&&window.ToolSharedUI&&!ToolApp.state.busy&&!ToolApp.state.restoring", timeout=60000)


def connected(page, key="a"):
    try:
        page.wait_for_function("key=>window.ToolApp&&ToolApp.state.source.connected&&ToolApp.state.model&&!ToolApp.state.busy&&!ToolApp.state.restoring&&ToolSharedFile.inspect().target?.itemId==='master-item-'+key", arg=key, timeout=60000)
    except Exception as error:
        value = page.evaluate("""()=>({url:location.href,source:window.ToolApp?.state.source,
          auth:window.ToolM365Auth?.inspect(),workspace:window.ToolSharedWorkspace?.inspect(),
          status:document.querySelector('#status')?.textContent,
          sharedStatus:document.querySelector('#shared-auth-status')?.textContent})""")
        raise AssertionError({"expected_target": key, "diagnostic": value}) from error
    assert page.evaluate("ToolSources.mode()") == "shared"


def photo_result(page):
    return page.evaluate("""async()=>{
      const photo=ToolApp.state.photos.at(-1);
      return {saved:!!photo?.saved,approval:photo?.approval,assessment:photo?.assessment,
        saveCount:ToolApp.state.source.saveCount,error:ToolApp.state.source.error,
        target:ToolSharedFile.inspect().target,modelTools:ToolApp.state.model.tools.map(t=>({serial:t.serial,last:t.last,due:t.due})),
        hash:photo?Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',photo.image.photoBytes))).map(n=>n.toString(16).padStart(2,'0')).join(''):'',
        preparedHash:ToolApp.state.preparedResult?Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',ToolApp.state.preparedResult.bytes))).map(n=>n.toString(16).padStart(2,'0')).join(''):''};
    }""")


def controlled_ocr(page, name, serial, last="2026.02.15", due="2027.02.15"):
    page.evaluate("""({name,serial,last,due})=>{const parse=ToolOCR.parse.bind(ToolOCR);const text=name+'\\n'+serial+'\\n교정일자 '+last+'\\n차기교정예정일자 '+due;window.ToolOCR={...ToolOCR,recognize:async()=>parse(text,ToolApp.state.model.tools,{confidence:98,passes:[{text,confidence:98},{text,confidence:98}]})};}""",
                  {"name": name, "serial": serial, "last": last, "due": due})


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--program", type=Path, default=Path("/workspace/outputs/Tool-Manager-Shared.html"))
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    parser.add_argument("--diagnostics", type=Path, default=Path("/tmp/tool-shared-app-test.json"))
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--photo", type=Path)
    parser.add_argument("--serial")
    parser.add_argument("--last")
    parser.add_argument("--due")
    args = parser.parse_args()
    if not args.program.is_file():
        parser.error("Supply --program with the built dedicated shared HTML")
    if args.photo and not all([args.workbook, args.serial, args.last, args.due]):
        parser.error("The real OCR case needs --workbook --photo --serial --last --due")
    for file in [args.workbook, args.photo]:
        if file is not None and not file.is_file():
            parser.error("Input file missing: " + str(file))
    assert b'tool-manager-shared-entry' in args.program.read_bytes()
    cases, errors, unexpected, contexts = Cases(), [], [], []
    report = {"program_sha256": sha256(args.program.read_bytes()).hexdigest(), "cases": cases,
              "real_microsoft_tenant_used": False, "physical_phone_or_safari_engine_tested": False,
              "authentication_api_replaced": False, "workspace_api_replaced": False,
              "graph_adapter_replaced": False, "concurrency_ocr_responses_replaced": True,
              "real_photo_ocr_responses_replaced": False if args.photo else None}
    with tempfile.TemporaryDirectory(prefix="tool-shared-app-") as name, serve_program(args.program) as url, sync_playwright() as playwright:
        temporary = Path(name)
        original = fixture(temporary)
        original_hash = sha256(original.read_bytes()).hexdigest()
        cloud = SharedCloud(original.read_bytes())
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
        local_host = urlsplit(url).netloc

        def new_device(actor, mobile=True, silent_error=False):
            context = browser.new_context(**(PHONE if mobile else {"viewport": {"width": 1360, "height": 1000}, "locale": "ko-KR", "accept_downloads": True}))
            contexts.append(context)
            traffic = {"authorize": [], "tokens": [], "dialogs": [], "silent_error": silent_error}
            context.add_init_script("""(()=>{
              window.sharedCopies=[];window.sharedShares=[];
              Object.defineProperty(navigator,'clipboard',{configurable:true,value:{writeText:async value=>sharedCopies.push({value,active:navigator.userActivation.isActive})}});
              Object.defineProperty(navigator,'share',{configurable:true,value:async value=>sharedShares.push({...value,active:navigator.userActivation.isActive})});
            })();""")

            def route(request_route):
                request, parsed = request_route.request, urlsplit(request_route.request.url)
                if parsed.hostname == "login.microsoftonline.com":
                    if parsed.path.endswith("/authorize"):
                        values = parse_qs(parsed.query)
                        traffic["authorize"].append(values)
                        assert values["redirect_uri"] == [url] and values["client_id"] == [CLIENT]
                        assert values["code_challenge_method"] == ["S256"]
                        query = {"code": "controlled-shared-code-" + actor, "state": values["state"][0]}
                        if traffic["silent_error"] and values.get("prompt") == ["none"]:
                            query = {"error": "interaction_required", "error_description": "Controlled company sign-in required", "state": values["state"][0]}
                        return request_route.fulfill(status=302, headers={"location": url + "?" + urlencode(query), "cache-control": "no-store"})
                    if parsed.path.endswith("/token"):
                        post = parse_qs(request.post_data or "")
                        traffic["tokens"].append(post)
                        values = traffic["authorize"][-1]
                        challenge = base64.urlsafe_b64encode(sha256(post["code_verifier"][0].encode()).digest()).decode().rstrip("=")
                        assert values["code_challenge"] == [challenge] and "client_secret" not in post
                        assert "offline_access" not in post["scope"][0]
                        claims = {"aud": CLIENT, "tid": TENANT, "iss": "https://login.microsoftonline.com/" + TENANT + "/v2.0",
                                  "nonce": values["nonce"][0], "exp": int(time.time()) + 3600, "oid": actor,
                                  "name": "Shared Fixture Engineer", "preferred_username": actor + "@example.invalid"}
                        return SharedCloud.respond(request_route, {"token_type": "Bearer", "expires_in": 3600,
                            "access_token": "shared-controlled-" + actor, "scope": "openid profile Files.ReadWrite.All", "id_token": jwt(claims)})
                if parsed.hostname in {"graph.microsoft.com", DOWNLOAD_HOST}:
                    return cloud.route(request_route, actor)
                if parsed.netloc == local_host and parsed.path.endswith("/team-settings.json"):
                    return SharedCloud.respond(request_route, {"enabled": False})
                if parsed.netloc != local_host:
                    unexpected.append(request.url)
                    return request_route.abort()
                return request_route.continue_()

            context.route("**/*", route)
            page = context.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))
            page.on("dialog", lambda dialog: (traffic["dialogs"].append(dialog.message), dialog.accept()))
            return context, page, traffic

        # An unconfigured URL must not pretend a local copy is the company file.
        _, empty, empty_traffic = new_device(MEMBER_A)
        empty.goto(url)
        ready(empty)
        assert not empty.evaluate("ToolApp.state.source.connected")
        assert not empty_traffic["authorize"] and not cloud.calls
        assert empty.locator("#shared-setup-link").is_visible()
        assert empty.locator("#workbook").is_hidden()
        cases["generic_entry_shows_setup_without_file_picker_or_false_connection"] = True

        # Owner enters public application ID once on their laptop. The actual
        # redirect round trip deliberately drops ?setup=1, as Entra does.
        _, owner, owner_traffic = new_device(OWNER, mobile=False)
        owner.goto(url + "?setup=1")
        ready(owner)
        reveal(owner, "#shared-client-id")
        owner.locator("#shared-client-id").fill(CLIENT)
        reveal(owner, "#shared-tenant-id")
        owner.locator("#shared-tenant-id").fill(TENANT)
        owner.locator("#shared-save-client").click()
        owner.locator("#shared-signin").click()
        owner.wait_for_function("window.ToolM365Auth&&window.ToolApp&&ToolM365Auth.inspect().signedIn&&!ToolApp.state.busy&&!ToolApp.state.restoring", timeout=60000)
        assert len(owner_traffic["tokens"]) == 1
        reveal(owner, "#shared-workbook-url")
        owner.locator("#shared-workbook-url").fill(SHARE_A)
        owner.locator("#shared-create").click()
        connected(owner)
        assert owner.evaluate("ToolSharedWorkspace.inspect().isOwner")
        join = owner.evaluate("ToolSharedWorkspace.invitationLink()")
        assert join.startswith(url + "#workspace=")
        encoded = parse_qs(urlsplit(join).fragment)["workspace"][0]
        bootstrap = json.loads(base64.urlsafe_b64decode(encoded + "=" * (-len(encoded) % 4)))
        assert set(bootstrap) == {"v", "workspaceId", "clientId", "tenantId", "settings"}
        assert "workbook" not in bootstrap and bootstrap["settings"]["itemId"] == "workspace-settings"
        assert all(value not in join for value in ["shared-controlled-", SHARE_A, "master-item-a", "master-drive-a"])
        assert set(json.loads(cloud.settings["data"])) == {"v", "workspaceId", "ownerObjectId", "workbook"}
        owner.locator("#shared-copy-link").click()
        assert owner.evaluate("sharedCopies.at(-1).value") == join
        assert owner.evaluate("sharedCopies.at(-1).active")
        cases["owner_setup_real_pkce_creates_private_settings_org_read_only_link"] = {"settings_revision": cloud.settings["revision"], "workbook": "a"}

        # Two separate phones have distinct memory tokens and empty storage.
        _, alice, alice_traffic = new_device(MEMBER_A)
        _, bob, bob_traffic = new_device(MEMBER_B)
        for page, traffic in [(alice, alice_traffic), (bob, bob_traffic)]:
            page.goto(join)
            connected(page)
            assert len(traffic["authorize"]) == len(traffic["tokens"]) == 1
            assert traffic["authorize"][0].get("prompt") == ["none"]
            assert page.locator("#shared-client-id").is_hidden() and page.locator("#shared-workbook-url").is_hidden()
            assert page.locator("#workbook").is_hidden()
            assert page.evaluate("document.documentElement.scrollWidth<=innerWidth")
        assert alice.evaluate("ToolM365Auth.inspect().account.objectId") == MEMBER_A
        assert bob.evaluate("ToolM365Auth.inspect().account.objectId") == MEMBER_B
        cases["independent_staff_phones_auto_sso_same_original_without_inputs"] = True

        # A successful login must not suppress silent SSO on an ordinary
        # subsequent launch. The bearer lives only in memory; the workspace
        # pointer survives and supplies public client/file configuration.
        alice.goto(url)
        connected(alice)
        assert len(alice_traffic["authorize"]) == len(alice_traffic["tokens"]) == 2
        assert alice_traffic["authorize"][-1].get("prompt") == ["none"]
        assert alice.url == join
        cases["ordinary_reopen_uses_remembered_workspace_and_silent_sso"] = True

        # Non-owner code cannot redirect coworkers to another workbook, and
        # Graph's actual config ACL rejects a direct conditional PUT too.
        old_settings = cloud.settings["data"]
        error = alice.evaluate("""async share=>{try{await ToolSharedWorkspace.changeWorkbook({workbookShareUrl:share});return null;}catch(e){return {code:e.code,message:e.message};}}""", SHARE_B)
        assert error and error["code"] == "workspace-owner", error
        denied = alice.evaluate("""async()=>{const s=ToolSharedWorkspace.inspect().bootstrap.settings;return (await fetch('https://graph.microsoft.com/v1.0/drives/'+s.driveId+'/items/'+s.itemId+'/content',{method:'PUT',headers:{Authorization:'Bearer '+await ToolM365Auth.getAccessToken(),'Content-Type':'application/json','If-Match':'\"settings-1\"'},body:'{}'})).status;}""")
        assert denied == 403 and cloud.settings["data"] == old_settings
        # A rejected administrative action clears proof; an ordinary refresh
        # reacquires it and keeps the fixed authoritative destination.
        alice.locator("#shared-refresh").click()
        connected(alice)
        cases["members_cannot_change_owner_configuration"] = True

        # Deliberately overlap both genuine workbook PUTs, from the same
        # version. Only OCR text is controlled here; both clients' parser,
        # automatic eligibility, photo bytes, writer and transport are real.
        for page, name_, serial in [(alice, "홍길동", "TEST0001"), (bob, "김가온", "TEST0002")]:
            controlled_ocr(page, name_, serial)
        cloud.block_first_pair = True
        alice.locator("#photos").set_input_files(temporary / "new-a.jpg")
        bob.locator("#photos").set_input_files(temporary / "new-b.png")
        for page in [alice, bob]:
            page.wait_for_function("ToolApp.state.photos.length===1&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=90000)
            result = photo_result(page)
            assert result["saved"] and result["approval"] == "automatic" and not result["error"], result
        assert len(cloud.puts) == 2 and len(cloud.conflicts) == 1, (cloud.puts, cloud.conflicts)
        assert {value["actor"] for value in cloud.puts} == {MEMBER_A, MEMBER_B}
        merged = temporary / "merged.xlsx"
        merged.write_bytes(cloud.files["a"]["data"])
        checked = check_combined(original, merged, {"A2": datetime(2027, 2, 15), "B2": datetime(2026, 2, 15), "A3": datetime(2027, 2, 15), "B3": datetime(2026, 2, 15)})
        with ZipFile(original) as before_archive, ZipFile(merged) as archive:
            _, _, before_pictures = drawing_info(before_archive, "Tools")
            _, _, pictures = drawing_info(archive, "Tools")
            assert set(pictures) == set(before_pictures)
            assert sha256(pictures[(2, 7)][0]["bytes"]).hexdigest() == photo_result(alice)["hash"]
            assert sha256(pictures[(3, 7)][0]["bytes"]).hexdigest() == photo_result(bob)["hash"]
            for slot in set(before_pictures) - {(2, 7), (3, 7)}:
                assert pictures[slot] == before_pictures[slot], slot
        for page in [alice, bob]:
            result = photo_result(page)
            assert result["target"]["itemId"] == "master-item-a"
        # The second writer must adopt the verified merged bytes, rather than
        # retaining its original unmerged candidate as the displayed result.
        second_actor = cloud.puts[-1]["actor"]
        second_page = alice if second_actor == MEMBER_A else bob
        second = photo_result(second_page)
        assert second["preparedHash"] == sha256(cloud.files["a"]["data"]).hexdigest(), second
        assert all(tool["last"] == "2026-02-15" for tool in second["modelTools"] if tool["serial"] in {"TEST0001", "TEST0002"})
        cases["two_authenticated_users_412_rebase_preserves_both_photos_dates_and_styles"] = {"conditional_writes": len(cloud.puts), "actual_412": len(cloud.conflicts), "cells_and_styles": checked}

        # A different photo in the same slot is not a disjoint-row merge.
        # Hold that upload through a subsequent refresh and an unrelated
        # automatic save, so it cannot silently overwrite the winning photo.
        _, contender_a, _ = new_device(MEMBER_A)
        _, contender_b, _ = new_device(MEMBER_B)
        for page in [contender_a, contender_b]:
            page.goto(join)
            connected(page)
            controlled_ocr(page, "이해솔", "TEST0003")
        before_count = len(cloud.puts)
        before_conflicts = len(cloud.conflicts)
        cloud.block_first_pair = True
        contender_a.locator("#photos").set_input_files(temporary / "new-c.png")
        contender_b.locator("#photos").set_input_files(temporary / "new-b.png")
        for page in [contender_a, contender_b]:
            page.wait_for_function("ToolApp.state.photos.length===1&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=90000)
        assert len(cloud.puts) == before_count + 1 and len(cloud.conflicts) == before_conflicts + 1
        winning_actor = cloud.puts[-1]["actor"]
        held = contender_b if winning_actor == MEMBER_A else contender_a
        winner = contender_a if winning_actor == MEMBER_A else contender_b
        winner_hash = photo_result(winner)["hash"]
        held_value = held.evaluate("({saved:!!ToolApp.state.photos[0].saved,review:!!ToolApp.state.photos[0].review,approval:ToolApp.state.photos[0].approval,edited:ToolApp.state.photos[0].edited,sharedConflict:ToolApp.state.photos[0].sharedConflict,error:ToolApp.state.source.error})")
        assert held_value["sharedConflict"] and held_value["edited"] and not held_value["saved"] and not held_value["review"] and not held_value["approval"] and held_value["error"], held_value
        # The transport read the winning cloud revision while attempting its
        # rebase. The app still displays its older model. A premature manual
        # confirmation must preserve the hold until that model is refreshed.
        assert held.evaluate("ToolApp.state.source.version!==ToolSharedFile.inspect().version")
        puts_before_manual = len(cloud.puts)
        held.locator("#confirm").check()
        held.locator("#save-review").click()
        held.wait_for_function("!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving")
        held_value = held.evaluate("({saved:!!ToolApp.state.photos[0].saved,review:!!ToolApp.state.photos[0].review,approval:ToolApp.state.photos[0].approval,sharedConflict:ToolApp.state.photos[0].sharedConflict,status:document.querySelector('#status').textContent})")
        assert held_value["sharedConflict"] and not held_value["saved"] and not held_value["review"] and not held_value["approval"], held_value
        assert "최신" in held_value["status"] and len(cloud.puts) == puts_before_manual, held_value
        cases["premature_manual_confirmation_keeps_conflict_hold_until_latest_read"] = True
        held.locator("#shared-refresh").click()
        connected(held)
        controlled_ocr(held, "홍길동", "TEST0001", "2027.02.15", "2028.02.15")
        held.locator("#photos").set_input_files(temporary / "new-a.jpg")
        held.wait_for_function("ToolApp.state.photos.length===2&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=60000)
        value = held.evaluate("ToolApp.state.photos.map(p=>({saved:!!p.saved,review:!!p.review,sharedConflict:!!p.sharedConflict,approval:p.approval}))")
        assert value[0] == {"saved": False, "review": False, "sharedConflict": True, "approval": None}, value
        assert value[1]["saved"] and value[1]["approval"] == "automatic", value
        assert len(cloud.puts) == before_count + 2
        with ZipFile(BytesIO(cloud.files["a"]["data"])) as archive:
            _, _, pictures = drawing_info(archive, "Tools")
            assert sha256(pictures[(4, 7)][0]["bytes"]).hexdigest() == winner_hash
        cases["same_slot_conflict_stays_held_after_refresh_and_unrelated_auto_save"] = True

        # Manager changes the central target after all devices are already
        # open. Their existing staff URL stays exactly the same.
        saved_a = cloud.files["a"]["data"]
        reveal(owner, "#shared-workbook-url")
        owner.locator("#shared-workbook-url").fill(SHARE_B)
        owner.locator("#shared-change-target").click()
        connected(owner, "b")
        assert owner.evaluate("ToolSharedWorkspace.invitationLink()") == join
        assert json.loads(cloud.settings["data"])["workbook"]["itemId"] == "master-item-b"
        for page in [alice, bob]:
            page.locator("#shared-refresh").click()
            connected(page, "b")
        _, fresh, fresh_traffic = new_device(MEMBER_A)
        fresh.goto(join)
        connected(fresh, "b")
        assert len(fresh_traffic["authorize"]) == 1
        assert fresh.locator("#shared-client-id").is_hidden() and fresh.locator("#shared-workbook-url").is_hidden()
        assert cloud.files["a"]["data"] == saved_a and cloud.files["b"]["data"] == original.read_bytes()
        assert [p for p in cloud.config_puts if p["actor"] == OWNER] == [{"actor": OWNER, "if_match": '"settings-1"'}]
        cases["owner_target_change_updates_open_and_new_devices_same_staff_link"] = True

        # Silent SSO failures present one normal login action, with no repeated
        # redirects, no path prompts and no writes to a local workbook copy.
        _, interaction, interaction_traffic = new_device(MEMBER_B, silent_error=True)
        interaction.goto(join)
        ready(interaction)
        interaction.wait_for_function("window.ToolM365Auth&&!ToolM365Auth.inspect().signedIn&&document.querySelector('#shared-signin')&&!document.querySelector('#shared-signin').disabled")
        assert len(interaction_traffic["authorize"]) == 1 and not interaction_traffic["tokens"]
        interaction.wait_for_timeout(300)
        assert len(interaction_traffic["authorize"]) == 1
        interaction.locator("#shared-signin").click()
        connected(interaction, "b")
        assert len(interaction_traffic["authorize"]) == 2 and interaction_traffic["authorize"][-1].get("prompt") != ["none"]
        cases["silent_sso_interaction_required_offers_single_normal_login_no_loop"] = True

        if args.photo:
            # Use a fresh browser after an external Excel update, preserving
            # the same central config/link. Expectations never enter JS/OCR.
            original_actual = args.workbook.read_bytes()
            cloud.files["b"]["data"] = original_actual
            cloud.files["b"]["revision"] += 1
            _, actual, _ = new_device(MEMBER_A)
            actual.goto(join)
            connected(actual, "b")
            began = time.monotonic()
            with actual.expect_file_chooser() as chooser:
                actual.locator("#take-photo").click()
            chooser.value.set_files({"name": "fresh-camera-photo", "mimeType": "image/jpeg", "buffer": args.photo.read_bytes()})
            actual.wait_for_function("ToolApp.state.photos.length===1&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=240000)
            result = photo_result(actual)
            assert result["saved"] and result["approval"] == "automatic" and not result["error"], result
            assert [result["assessment"][key] for key in ["serial", "last", "due"]] == [args.serial, args.last, args.due], result
            updated = temporary / "actual-updated.xlsx"
            updated.write_bytes(cloud.files["b"]["data"])
            expected = {"serial": args.serial, "last": args.last, "due": args.due}
            verification = verify_increment(args.workbook, updated, target_reference(args.workbook, args.serial), expected, result["hash"])
            assert result["preparedHash"] == sha256(updated.read_bytes()).hexdigest()
            assert args.workbook.read_bytes() == original_actual
            cases["actual_camera_photo_real_ocr_shared_original_autosave"] = {"seconds": round(time.monotonic() - began, 2), **verification}

        # Public/persistent client metadata may be remembered. Bearer tokens,
        # OAuth codes, verifier and ID tokens must not be stored or shared.
        for context in contexts:
            for page in context.pages:
                stored = page.evaluate("JSON.stringify({local:Object.entries(localStorage),session:Object.entries(sessionStorage)})")
                assert "shared-controlled-" not in stored and "controlled-shared-code" not in stored
                assert not re.search(r'"(?:access_token|id_token|code_verifier|refresh_token)"', stored)
                assert "shared-controlled-" not in page.url
        assert sha256(original.read_bytes()).hexdigest() == original_hash
        assert not errors and not unexpected, {"errors": errors, "unexpected_requests": unexpected}
        cases["no_bearer_tokens_in_persistent_storage_or_staff_urls"] = True
        report["page_errors"], report["unexpected_requests"] = errors, unexpected
        report["graph_conditional_writes"] = cloud.puts
        for context in contexts:
            context.close()
        browser.close()
    args.diagnostics.write_text(json.dumps(report, ensure_ascii=False, indent=2))
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
