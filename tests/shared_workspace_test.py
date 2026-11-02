"""Owner-controlled cloud workspace settings through real browser fetch/storage.

Graph, the organization sharing ACL, and token issuance are controlled HTTP
fixtures. This checks protocol and state handling, not a real company tenant.
No private workbook, photo, Microsoft account, or credentials are used.
"""
from __future__ import annotations

import argparse
import base64
import copy
import json
from pathlib import Path
import tempfile
from urllib.parse import parse_qs, urlsplit

from playwright.sync_api import sync_playwright

from browser_test import REPO, serve_program


CLIENT = "11111111-1111-4111-8111-111111111111"
OWNER = "22222222-2222-4222-8222-222222222222"
STAFF = "33333333-3333-4333-8333-333333333333"
TOKEN = "TEST_WORKSPACE_BEARER_NEVER_PERSISTED"
TICKET = "TEST_TEMP_DOWNLOAD_PROOF_NEVER_PERSISTED"
BOOK_A = "https://tenant.example.test/:x:/r/sites/Tools/book-a.xlsx?web=1"
BOOK_B = "https://tenant.example.test/:x:/r/sites/Tools/book-b.xlsx?web=1"
ORG_LINK = "https://tenant.example.test/:t:/s/Tools/settings-org-view"


class Cloud:
    def __init__(self):
        self.document = None
        self.revision = 1
        self.file_name = ""
        self.calls = []
        self.puts = []
        self.actor = "owner"
        self.mode = "normal"
        self.redeemed = False
        self.metadata_reads = 0
        self.paused = None
        self.pause_permissions = False

    @property
    def etag(self):
        return '"settings-' + str(self.revision) + '"'

    def data(self):
        if self.mode == "invalid-json":
            return b"{invalid"
        value = copy.deepcopy(self.document)
        if self.mode == "wrong-workspace":
            value["workspaceId"] = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"
        if self.mode == "credential-field":
            value["accessToken"] = TOKEN
        return json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode()

    def settings(self):
        return {"id": "settings-1", "name": self.file_name, "size": len(self.data()), "eTag": self.etag,
                "file": {"mimeType": "application/json"}, "parentReference": {"driveId": "owner-drive"},
                "@microsoft.graph.downloadUrl": "https://download.example.test/config?ticket=" + TICKET}

    def book(self, key):
        return {"id": "book-" + key, "name": "관리 엑셀 " + key + ".xlsx", "size": 6000,
                "eTag": '"workbook-1"', "file": {"mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
                "parentReference": {"driveId": "company-drive"}}

    def permissions(self):
        values = [{"id": "owner-permission", "roles": ["owner"], "grantedToV2": {"user": {"id": OWNER}}},
                  {"id": "org-view", "roles": ["read"], "link": {"scope": "organization", "type": "view", "webUrl": ORG_LINK}}]
        if self.mode == "other-writer":
            values.append({"id": "bad-grant", "roles": ["write"], "grantedToV2": {"user": {"id": STAFF}}})
        response = {"value": values}
        if self.mode == "partial-permissions":
            response["@odata.nextLink"] = "https://graph.microsoft.com/v1.0/permissions?$skiptoken=remaining"
        return response

    def route(self, route):
        request = route.request
        parsed = urlsplit(request.url)
        headers = {"access-control-allow-origin": "*", "access-control-allow-methods": "GET,POST,PUT,OPTIONS",
                   "access-control-allow-headers": "Authorization,Content-Type,If-Match,If-None-Match,Prefer", "cache-control": "no-store"}
        if request.method == "OPTIONS":
            return route.fulfill(status=204, headers=headers)
        auth = request.headers.get("authorization")
        self.calls.append({"host": parsed.hostname, "path": parsed.path, "method": request.method,
                           "authorization": auth, "if_match": request.headers.get("if-match"),
                           "if_none_match": request.headers.get("if-none-match"), "prefer": request.headers.get("prefer")})
        if parsed.hostname == "download.example.test":
            assert auth is None, "Graph bearer must not leave graph.microsoft.com"
            return route.fulfill(status=200, headers={**headers, "content-type": "application/json"}, body=self.data())
        assert parsed.hostname == "graph.microsoft.com" and parsed.path.startswith("/v1.0/"), request.url
        assert auth == "Bearer " + TOKEN
        path = parsed.path[5:]

        def answer(value, status=200):
            return route.fulfill(status=status, headers={**headers, "content-type": "application/json"}, body=json.dumps(value, ensure_ascii=False))

        if path == "/me/drive":
            actor = OWNER if self.actor == "owner" else STAFF
            return answer({"id": "owner-drive" if self.actor == "owner" else "staff-drive", "owner": {"user": {"id": actor}}, "driveType": "business"})
        if path.startswith("/me/drive/root:/") and path.endswith(":/content"):
            assert request.method == "PUT" and request.headers.get("if-none-match") == "*"
            self.document = json.loads(request.post_data)
            from urllib.parse import unquote
            self.file_name = unquote(path[len("/me/drive/root:/"):-len(":/content")])
            assert self.file_name.startswith("ToolManagerShared-") and self.file_name.endswith(".json")
            self.puts.append({"kind": "create", "document": copy.deepcopy(self.document)})
            return answer(self.settings(), 201)
        if path.startswith("/shares/u!"):
            encoded = path[len("/shares/u!"):].split("/", 1)[0]
            link = base64.urlsafe_b64decode(encoded + "=" * (-len(encoded) % 4)).decode()
            assert request.headers.get("prefer") == "redeemSharingLinkIfNecessary"
            if link == ORG_LINK:
                self.redeemed = True
                value = self.settings()
                if self.mode == "wrong-redemption":
                    value["id"] = "other-settings"
                return answer(value)
            return answer(self.book("b" if link == BOOK_B else "a"))
        if path.startswith("/drives/company-drive/items/book-"):
            return answer(self.book("b" if "book-b" in path else "a"))
        assert path.startswith("/drives/owner-drive/items/settings-1"), request.url
        if path.endswith("/createLink"):
            assert request.method == "POST"
            assert json.loads(request.post_data) == {"type": "view", "scope": "organization"}
            link = {"scope": "organization", "type": "view", "webUrl": ORG_LINK}
            roles = ["read"]
            if self.mode == "anonymous-link":
                link["scope"] = "anonymous"
            if self.mode == "edit-link":
                link["type"] = "edit"
                roles = ["write"]
            return answer({"id": "org-view", "link": link, "roles": roles})
        if path.endswith("/permissions"):
            if self.pause_permissions:
                self.pause_permissions = False
                self.paused = (route, headers)
                return
            return answer(self.permissions())
        if path.endswith("/content"):
            assert request.method == "PUT"
            self.puts.append({"kind": "change", "if_match": request.headers.get("if-match"), "document": json.loads(request.post_data)})
            if self.actor != "owner" or self.mode == "write403":
                return answer({}, 403)
            if self.mode == "write412" or request.headers.get("if-match") != self.etag:
                return answer({}, 412)
            self.document = json.loads(request.post_data)
            self.revision += 1
            if self.mode == "commit-abort":
                return route.abort()
            return answer(self.settings())
        if self.mode in {"need-redemption", "wrong-redemption"} and not self.redeemed:
            return answer({}, 403)
        if self.mode == "settings403":
            return answer({}, 403)
        self.metadata_reads += 1
        value = self.settings()
        if self.mode == "changing-read" and self.metadata_reads % 2 == 0:
            value["eTag"] = '"raced-version"'
        return answer(value)


def outcome(page, expression):
    return page.evaluate("async expression=>{try{return {ok:true,value:await eval(expression)}}catch(error){return {ok:false,code:error.code,message:error.message}}}", expression)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    args = parser.parse_args()
    source = (REPO / "web/shared-workspace.js").read_text()
    auth = {"configured": True, "supported": True, "signedIn": True, "clientId": CLIENT, "tenantId": "organizations",
            "account": {"objectId": OWNER, "username": "owner@example.invalid"}}
    cases, errors, unexpected = {}, [], []
    with tempfile.TemporaryDirectory(prefix="tool-shared-workspace-") as directory, sync_playwright() as playwright:
        program = Path(directory) / "shared.html"
        program.write_text('<!doctype html><html><head><meta charset="utf-8"></head><body><script>window.authState=' + json.dumps(auth) + ';window.ToolM365Auth={inspect:()=>JSON.parse(JSON.stringify(authState)),getAccessToken:async()=>{if(!authState.signedIn)throw new Error("login required");return ' + json.dumps(TOKEN) + ';}};</script><script>' + source + '</script></body></html>')
        with serve_program(program) as url:
            browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])

            def fresh():
                cloud = Cloud()
                context = browser.new_context()
                def route(request_route):
                    parsed = urlsplit(request_route.request.url)
                    if parsed.hostname in {"graph.microsoft.com", "download.example.test"}:
                        return cloud.route(request_route)
                    if parsed.netloc != urlsplit(url).netloc:
                        unexpected.append(request_route.request.url)
                        return request_route.abort()
                    return request_route.continue_()
                context.route("**/*", route)
                page = context.new_page()
                page.on("pageerror", lambda error: errors.append(str(error)))
                page.goto(url)
                page.wait_for_function("window.ToolSharedWorkspace")
                return context, page, cloud

            context, page, cloud = fresh()
            created = outcome(page, "ToolSharedWorkspace.create({workbookShareUrl:" + json.dumps(BOOK_A) + "})")
            assert created["ok"], created
            pointer = created["value"]
            assert pointer["settings"]["shareUrl"] == ORG_LINK
            assert page.evaluate("ToolSharedWorkspace.inspect().isOwner")
            original_link = page.evaluate("ToolSharedWorkspace.invitationLink()")
            assert original_link.startswith(url + "#workspace=") and "?" not in original_link
            raw = original_link.split("#workspace=", 1)[1]
            decoded = json.loads(base64.urlsafe_b64decode(raw + "=" * (-len(raw) % 4)))
            assert decoded == pointer and "workbook" not in decoded
            stored = page.evaluate("JSON.stringify({...localStorage})")
            assert TOKEN not in stored and TICKET not in stored and "access_token" not in original_link
            assert cloud.document["workbook"]["itemId"] == "book-a" and len(cloud.puts) == 1
            assert cloud.puts[0]["kind"] == "create"
            cases["owner_creates_private_settings_and_org_read_only_invitation"] = True

            page.goto(original_link)
            assert page.evaluate("ToolSharedWorkspace.takeInvitation()") == pointer
            assert page.evaluate("ToolSharedWorkspace.takeInvitation()") is None
            assert page.url == original_link, "Keep the workspace bookmark for a fresh home-screen browser"
            refreshed = outcome(page, "ToolSharedWorkspace.refresh()")
            assert refreshed["ok"] and refreshed["value"]["workbook"]["itemId"] == "book-a"
            assert refreshed["value"]["configVersion"] == cloud.etag
            cases["invitation_survives_bookmark_and_reload_without_credentials"] = True

            cloud.actor = "staff"
            page.evaluate("id=>{authState.account={objectId:id,username:'staff@example.invalid'}}", STAFF)
            result = outcome(page, "ToolSharedWorkspace.refresh()")
            assert result["ok"] and not page.evaluate("ToolSharedWorkspace.inspect().isOwner")
            before = len(cloud.puts)
            denied = outcome(page, "ToolSharedWorkspace.changeWorkbook({workbookShareUrl:" + json.dumps(BOOK_B) + "})")
            assert not denied["ok"] and denied["code"] == "workspace-owner" and len(cloud.puts) == before, denied
            cases["staff_reads_settings_but_cannot_change_shared_destination"] = True

            cloud.actor = "owner"
            page.evaluate("id=>{authState.account={objectId:id,username:'owner@example.invalid'}}", OWNER)
            result = outcome(page, "ToolSharedWorkspace.refresh()")
            original_version = result["value"]["configVersion"]
            changed = outcome(page, "ToolSharedWorkspace.changeWorkbook({workbookShareUrl:" + json.dumps(BOOK_B) + ",expectedVersion:" + json.dumps(original_version) + "})")
            assert changed["ok"] and changed["value"]["workbook"]["itemId"] == "book-b", changed
            assert cloud.puts[-1]["if_match"] == original_version
            assert page.evaluate("ToolSharedWorkspace.invitationLink()") == original_link
            cloud.actor = "staff"
            page.evaluate("id=>{authState.account={objectId:id,username:'staff@example.invalid'}}", STAFF)
            staff_latest = outcome(page, "ToolSharedWorkspace.refresh()")
            assert staff_latest["ok"] and staff_latest["value"]["workbook"]["itemId"] == "book-b"
            assert page.evaluate("ToolSharedWorkspace.invitationLink()") == original_link
            cases["owner_changes_central_pointer_same_employee_link_gets_new_master"] = True
            context.close()

            for mode, expected in [("anonymous-link", "workspace-sharing"), ("edit-link", "workspace-sharing"), ("other-writer", "workspace-owner"), ("partial-permissions", "workspace-owner")]:
                context, page, cloud = fresh(); cloud.mode = mode
                value = outcome(page, "ToolSharedWorkspace.create({workbookShareUrl:" + json.dumps(BOOK_A) + "})")
                assert not value["ok"] and value["code"] == expected and not page.evaluate("ToolSharedWorkspace.inspect().configured"), value
                cases[mode + "_not_published_as_owner_only"] = True
                context.close()

            for mode, expected in [("invalid-json", "workspace-settings"), ("wrong-workspace", "workspace-settings"), ("credential-field", "workspace-settings"), ("changing-read", "workspace-conflict"), ("settings403", "workspace-denied"), ("wrong-redemption", "workspace-settings")]:
                context, page, cloud = fresh()
                value = outcome(page, "ToolSharedWorkspace.create({workbookShareUrl:" + json.dumps(BOOK_A) + "})")
                assert value["ok"], value
                cloud.mode = mode; cloud.redeemed = False; cloud.metadata_reads = 0
                held = outcome(page, "ToolSharedWorkspace.refresh()")
                assert not held["ok"] and held["code"] == expected, held
                state = page.evaluate("ToolSharedWorkspace.inspect()")
                assert state["configured"] and state["workbook"] is None and not state["configVersion"]
                cases[mode + "_no_stale_target_fallback"] = True
                context.close()

            context, page, cloud = fresh()
            assert outcome(page, "ToolSharedWorkspace.create({workbookShareUrl:" + json.dumps(BOOK_A) + "})")["ok"]
            cloud.mode = "need-redemption"; cloud.redeemed = False
            redeemed = outcome(page, "ToolSharedWorkspace.refresh()")
            assert redeemed["ok"] and cloud.redeemed
            assert any(call["path"].startswith("/v1.0/shares/") and call["prefer"] for call in cloud.calls)
            cases["organization_view_link_redeemed_only_on_actual_denial"] = True
            cloud.mode = "write412"
            rejected = outcome(page, "ToolSharedWorkspace.changeWorkbook({workbookShareUrl:" + json.dumps(BOOK_B) + "})")
            assert not rejected["ok"] and rejected["code"] == "workspace-conflict" and cloud.document["workbook"]["itemId"] == "book-a", rejected
            cases["conditional_owner_change_does_not_overwrite_other_change"] = True
            cloud.mode = "normal"
            assert outcome(page, "ToolSharedWorkspace.refresh()")["ok"]
            cloud.mode = "commit-abort"
            value = outcome(page, "ToolSharedWorkspace.changeWorkbook({workbookShareUrl:" + json.dumps(BOOK_B) + "})")
            assert not value["ok"] and cloud.document["workbook"]["itemId"] == "book-b"
            count = len(cloud.puts)
            cloud.mode = "normal"
            resolved = outcome(page, "ToolSharedWorkspace.refresh()")
            assert resolved["ok"] and resolved["value"]["workbook"]["itemId"] == "book-b" and len(cloud.puts) == count
            cases["unknown_commit_requires_fresh_read_never_auto_retries"] = True
            context.close()

            context, page, cloud = fresh()
            assert outcome(page, "ToolSharedWorkspace.create({workbookShareUrl:" + json.dumps(BOOK_A) + "})")["ok"]
            cloud.pause_permissions = True
            page.evaluate("url=>{window.changeResult=ToolSharedWorkspace.changeWorkbook({workbookShareUrl:url}).then(value=>({ok:true,value})).catch(error=>({ok:false,code:error.code}));}", BOOK_B)
            for _ in range(100):
                if cloud.paused:
                    break
                page.wait_for_timeout(10)
            assert cloud.paused
            page.evaluate("()=>{authState.signedIn=false;authState.account=null}")
            route, headers = cloud.paused
            route.fulfill(status=200, headers={**headers, "content-type": "application/json"}, body=json.dumps(cloud.permissions()))
            result = page.evaluate("changeResult")
            assert not result["ok"] and result["code"] == "workspace-changed" and len(cloud.puts) == 1, result
            assert not page.evaluate("ToolSharedWorkspace.inspect().isOwner")
            cases["logout_cannot_continue_with_previous_owner_token"] = True
            context.close()

            context, page, cloud = fresh()
            assert outcome(page, "ToolSharedWorkspace.create({workbookShareUrl:" + json.dumps(BOOK_A) + "})")["ok"]
            before = len(cloud.calls)
            page.evaluate("()=>{authState.signedIn=false;authState.account=null}")
            value = outcome(page, "ToolSharedWorkspace.refresh()")
            assert not value["ok"] and len(cloud.calls) == before
            cases["expired_login_keeps_pointer_without_using_stale_settings"] = True
            context.close()

            context, page, cloud = fresh()
            value = outcome(page, "ToolSharedWorkspace.create({workbookShareUrl:" + json.dumps(BOOK_A) + "})")
            assert value["ok"]
            before = len(cloud.calls)
            page.evaluate("""()=>{
              ToolM365Auth.getAccessToken=()=>new Promise(resolve=>{window.finishToken=()=>resolve('TEST_WORKSPACE_BEARER_NEVER_PERSISTED')});
              window.delayedRead=ToolSharedWorkspace.refresh().then(value=>({ok:true,value})).catch(error=>({ok:false,code:error.code}));
            }""")
            page.wait_for_function("typeof finishToken==='function'")
            page.evaluate("""()=>{
              const next=ToolSharedWorkspace.inspect().bootstrap;
              next.workspaceId='aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
              ToolSharedWorkspace.configure(next);finishToken();
            }""")
            raced = page.evaluate("delayedRead")
            assert not raced["ok"] and raced["code"] == "workspace-changed" and len(cloud.calls) == before, raced
            assert page.evaluate("ToolSharedWorkspace.inspect().bootstrap.workspaceId") == "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"
            assert page.evaluate("ToolSharedWorkspace.inspect().workspace") is None
            cases["workspace_switch_while_token_pending_cannot_read_old_pointer"] = True
            context.close()
            browser.close()
    assert not errors and not unexpected, (errors, unexpected)
    print(json.dumps({"cases": cases, "case_count": len(cases), "page_errors": errors, "unexpected_requests": unexpected,
                      "real_company_account_used": False,
                      "scope": "Real browser fetch/localStorage/URL handling with controlled Graph sharing and files"}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
