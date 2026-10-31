"""Real-browser checks for private-fragment team onboarding, without company data."""
from __future__ import annotations

import argparse
import base64
import json
from pathlib import Path
import tempfile
from urllib.parse import urlsplit

from playwright.sync_api import sync_playwright

from browser_test import REPO, serve_program


CLIENT = "00000000-0000-4000-8000-000000000001"
TENANT = "11111111-1111-4111-8111-111111111111"
SHARE = "https://example.sharepoint.com/:x:/s/team/company.xlsx?e=sharing-code"
CONFIG = {"v": 1, "clientId": CLIENT, "tenantId": TENANT, "shareUrl": SHARE, "name": "공용 툴 대장"}
TARGET_CONFIG = {"v": 1, "clientId": CLIENT, "tenantId": TENANT, "name": "공용 툴 대장",
                 "target": {"driveId": "b!registered-drive", "itemId": "registered-item", "name": "공용.xlsx"}}
KEY = "tool-manager-team-config-v1"


def encode(value):
    return base64.urlsafe_b64encode(json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode()).decode().rstrip("=")


def caught(page, expression):
    return page.evaluate("""expression=>{try{return {value:eval(expression)}}catch(e){return {code:e.code,message:e.message}}}""", expression)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    args = parser.parse_args()
    source = (REPO / "web/team-config.js").read_text()
    cases = {}
    with tempfile.TemporaryDirectory(prefix="tool-team-config-") as directory:
        program = Path(directory) / "index.html"
        program.write_text('<!doctype html><meta charset="utf-8"><script>' + source.replace("</script", "<\\/script") + "</script>")
        with serve_program(program) as url, sync_playwright() as playwright:
            browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
            context = browser.new_context()
            page = context.new_page()
            errors, external, published_requests = [], [], []
            published = {"status": 200, "body": {"enabled": False}, "headers": {}}
            page.on("pageerror", lambda error: errors.append(str(error)))

            def route(request):
                if urlsplit(request.request.url).path.endswith("/team-settings.json"):
                    published_requests.append({"url": request.request.url, "headers": request.request.headers})
                    if published.get("abort"):
                        return request.abort()
                    body = published["body"]
                    return request.fulfill(status=published["status"], content_type="application/json", headers=published["headers"],
                                           body=json.dumps(body) if not isinstance(body, str) else body)
                if urlsplit(request.request.url).netloc != urlsplit(url).netloc:
                    external.append(request.request.url)
                    return request.abort()
                return request.continue_()

            context.route("**/*", route)

            def navigate(target):
                page.goto("about:blank")
                page.goto(target)
            navigate(url + "?source=homescreen&token=should-not-copy#section")
            supplied = dict(CONFIG, clientId=CLIENT.upper(), accessToken="never-save-this", password="never-save-this", redirectUri="https://attacker.invalid/")
            generated = page.evaluate("value=>ToolTeamConfig.create(value)", supplied)
            assert generated.startswith(url + "#team=")
            assert json.loads(base64.urlsafe_b64decode(generated.split("#team=")[1] + "=" * (-len(generated.split("#team=")[1]) % 4))) == CONFIG
            assert "token=" not in generated and "never-save-this" not in generated
            assert page.evaluate("ToolTeamConfig.inspect().configured") is False
            assert page.evaluate("localStorage.length") == 0
            cases["clean_allowlisted_same_origin_link_without_storage_side_effect"] = True

            navigate(generated + "&section&other=x%20y")
            assert page.url == url + "#section&other=x%20y", page.url
            assert page.evaluate("ToolTeamConfig.inspect().configured") is False
            assert page.evaluate("ToolTeamConfig.takeInvitation()") == CONFIG
            assert page.evaluate("ToolTeamConfig.takeInvitation()") is None
            assert page.evaluate("localStorage.length") == 0
            cases["unicode_invitation_scrubbed_before_use_take_once_no_silent_apply"] = True

            new_config = {**CONFIG, "name": "다른 공용 대장"}
            page.evaluate("fragment=>{location.hash=fragment}", "team=" + encode(new_config) + "&keep=1")
            assert "#team=" in page.url
            assert page.evaluate("ToolTeamConfig.takeInvitation()") == new_config
            assert page.url == url + "#keep=1"
            assert page.evaluate("ToolTeamConfig.takeInvitation()") is None
            assert page.evaluate("localStorage.length") == 0
            page.evaluate("location.hash='team=invalid!' ")
            assert caught(page, "ToolTeamConfig.takeInvitation()").get("code") == "team-link"
            assert page.url == url
            assert page.evaluate("ToolTeamConfig.takeInvitation()") is None
            cases["same_document_new_hash_invitation_and_invalid_link_consumed_once"] = True

            saved = page.evaluate("value=>ToolTeamConfig.remember(value)", supplied)
            assert saved == CONFIG
            assert json.loads(page.evaluate("key=>localStorage.getItem(key)", KEY)) == CONFIG
            assert page.evaluate("()=>{let x=ToolTeamConfig.inspect();x.configuration.name='mutated';return ToolTeamConfig.inspect().configuration.name}") == CONFIG["name"]
            page.reload()
            assert page.evaluate("ToolTeamConfig.inspect().configuration") == CONFIG
            assert "never-save-this" not in page.evaluate("JSON.stringify(localStorage)")
            cases["persistent_allowlist_and_return_value_isolation"] = True

            page.evaluate("ToolTeamConfig.clear()")
            assert page.evaluate("ToolTeamConfig.inspect().configured") is False
            assert page.evaluate("key=>localStorage.getItem(key)", KEY) is None
            cases["explicit_clear"] = True

            bad = [
                {**CONFIG, "v": 2}, {**CONFIG, "clientId": "invalid"}, {**CONFIG, "tenantId": "common"},
                {**CONFIG, "tenantId": "foo.onmicrosoft.com.attacker.invalid"},
                {**CONFIG, "shareUrl": "javascript:alert(1)"}, {**CONFIG, "shareUrl": "http://example.invalid/book.xlsx"},
                {**CONFIG, "shareUrl": "https://user:password@example.invalid/book.xlsx"},
                {**CONFIG, "shareUrl": "https://example.invalid/" + "x" * 12000},
                {**CONFIG, "name": "x" * 121}, {**CONFIG, "name": "bad\nname"}, [], None,
            ]
            for value in bad:
                assert page.evaluate("value=>{try{ToolTeamConfig.create(value);return false}catch(e){return e.code==='team-config'}}", value), value
                assert page.evaluate("value=>{try{ToolTeamConfig.remember(value);return false}catch(e){return e.code==='team-config'}}", value), value
            for tenant in ["organizations", "company.onmicrosoft.com", TENANT.upper()]:
                assert page.evaluate("value=>!!ToolTeamConfig.create(value)", {**CONFIG, "tenantId": tenant})
            cases["strict_schema_tenant_https_credentials_name_length_boundaries"] = True

            target_link = page.evaluate("value=>ToolTeamConfig.create(value)", {**TARGET_CONFIG, "access_token": "never-save-this"})
            navigate(target_link)
            assert page.evaluate("ToolTeamConfig.takeInvitation()") == TARGET_CONFIG
            page.evaluate("value=>ToolTeamConfig.remember(value)", TARGET_CONFIG)
            assert page.evaluate("()=>{let c=ToolTeamConfig.inspect().configuration;c.target.itemId='mutated';return ToolTeamConfig.inspect().configuration.target.itemId}") == TARGET_CONFIG["target"]["itemId"]
            exported = json.loads(page.evaluate("value=>ToolTeamConfig.exportSettings(value)", {**TARGET_CONFIG, "password": "never-save-this", "redirectUri": "https://attacker.invalid/"}))
            assert exported == {"enabled": True, **TARGET_CONFIG}
            assert "shareUrl" not in exported
            assert page.evaluate("value=>{try{ToolTeamConfig.exportSettings(value);return false}catch(e){return e.code==='team-config'}}", CONFIG)
            invalid_targets = [
                {**TARGET_CONFIG, "shareUrl": SHARE}, {**TARGET_CONFIG, "target": None},
                *[{**TARGET_CONFIG, "target": {**TARGET_CONFIG["target"], "driveId": value}} for value in ["", "x" * 201, "a/b", "a\\b", "a\nb", "a%2Fb"]],
                *[{**TARGET_CONFIG, "target": {**TARGET_CONFIG["target"], "name": value}} for value in ["", "x.txt", "path/book.xlsx", "bad\n.xlsx"]],
            ]
            for value in invalid_targets:
                assert page.evaluate("value=>{try{ToolTeamConfig.create(value);return false}catch(e){return e.code==='team-config'}}", value), value
            cases["resolved_target_invitation_deep_clone_public_export_and_bounded_ids"] = True

            published.update(body={"enabled": True, **TARGET_CONFIG, "access_token": "never-save-this"})
            loaded = page.evaluate("ToolTeamConfig.loadPublished()")
            assert loaded == TARGET_CONFIG
            assert page.evaluate("ToolTeamConfig.inspect().configuration") == TARGET_CONFIG
            assert published_requests[-1]["url"] == url.rsplit("/", 1)[0] + "/team-settings.json"
            assert "authorization" not in published_requests[-1]["headers"]
            for response in [{"status": 404, "body": {}}, {"status": 200, "body": {"enabled": False}}]:
                published.update(response)
                assert page.evaluate("ToolTeamConfig.loadPublished()") is None
            cases["published_same_origin_target_load_disabled_missing_and_no_auth"] = True

            for response in [
                {"status": 200, "body": {"enabled": True, **CONFIG}},
                {"status": 200, "body": {**TARGET_CONFIG}},
                {"status": 200, "body": "{broken json"},
                {"status": 200, "body": {"enabled": True, **TARGET_CONFIG, "target": {"driveId": "invalid/drive", "itemId": "x", "name": "x.xlsx"}}},
                {"status": 500, "body": {}},
                {"status": 302, "body": {}, "headers": {"location": "https://attacker.invalid/settings.json"}},
                {"status": 200, "body": " " * 17000},
            ]:
                published.update(status=200, headers={}, abort=False)
                published.update(response)
                result = page.evaluate("async()=>{try{return {value:await ToolTeamConfig.loadPublished()}}catch(e){return {code:e.code}}}")
                allowed_codes = {"team-unavailable"} if response["status"] == 302 else {"team-config", "team-published"}
                assert result.get("code") in allowed_codes, response
                assert page.evaluate("ToolTeamConfig.inspect().configuration") == TARGET_CONFIG
            cases["published_invalid_redirect_error_oversize_never_overwrite_current"] = True

            stream_result = page.evaluate("""async()=>{
                const original=window.fetch;window.fetch=async()=>new Response(new ReadableStream({start(c){c.enqueue(new Uint8Array(17000));c.close()}}));
                try{await ToolTeamConfig.loadPublished();return false}catch(e){return e.code==='team-published'}finally{window.fetch=original}
            }""")
            assert stream_result
            observed = page.evaluate("""async()=>{
                const original=window.fetch;let observed;
                window.fetch=async(url,options)=>{observed={url,cache:options.cache,credentials:options.credentials,redirect:options.redirect,signal:!!options.signal};return new Response(JSON.stringify({enabled:false}))};
                try{await ToolTeamConfig.loadPublished();return observed}finally{window.fetch=original}
            }""")
            assert observed == {"url": url.rsplit("/", 1)[0] + "/team-settings.json", "cache": "no-store", "credentials": "omit", "redirect": "error", "signal": True}
            cases["published_stream_limit_and_fetch_privacy_options"] = True

            unavailable = page.evaluate("""async()=>{
                const original=window.fetch;window.fetch=async()=>{throw new TypeError('Network error')};
                try{await ToolTeamConfig.loadPublished();return false}catch(e){return e.code==='team-unavailable'}finally{window.fetch=original}
            }""")
            assert unavailable
            timeout = page.evaluate("""async()=>{
                const original=window.fetch;window.fetch=async(url,options)=>new Promise((resolve,reject)=>options.signal.addEventListener('abort',()=>reject(new DOMException('Aborted','AbortError'))));
                try{await ToolTeamConfig.loadPublished();return false}catch(e){return e.code==='team-unavailable'}finally{window.fetch=original}
            }""")
            assert timeout
            offline = page.evaluate("""async()=>{
                Object.defineProperty(navigator,'onLine',{configurable:true,value:false});const original=window.fetch;let calls=0;window.fetch=async()=>{calls++;throw new Error('unexpected')};
                try{return {result:await ToolTeamConfig.loadPublished(),calls}}finally{window.fetch=original;delete navigator.onLine}
            }""")
            assert offline == {"result": None, "calls": 0}
            cases["network_timeout_unavailable_distinguished_and_offline_skips_network"] = True

            payload = encode(CONFIG)
            malformed = ["", "!", "YWJj", encode({**CONFIG, "v": 2}), encode({key: value for key, value in CONFIG.items() if key != "v"}),
                         "A" * 18001, base64.urlsafe_b64encode(b"\xff\xfe").decode().rstrip("=")]
            for encoded in malformed:
                navigate(url + "#team=" + encoded + "&keep=1")
                assert page.url == url + "#keep=1", (encoded[:80], page.url[:150])
                assert caught(page, "ToolTeamConfig.takeInvitation()").get("code") == "team-link", encoded[:80]
                assert page.evaluate("ToolTeamConfig.takeInvitation()") is None
            for fragment in ["team=" + payload + "&team=" + payload, "%74eam=" + payload + "&team=" + payload]:
                navigate(url + "#" + fragment + "&keep=1")
                assert page.url == url + "#keep=1"
                assert caught(page, "ToolTeamConfig.takeInvitation()").get("code") == "team-link"
            cases["malformed_duplicate_encoded_duplicate_oversize_utf8_rejected_and_scrubbed"] = True

            for suffix in ["#ordinary-section", "#keep=x%20y&section", "?team=" + payload + "#untouched"]:
                navigate(url + suffix)
                assert page.url == url + suffix
                assert page.evaluate("ToolTeamConfig.takeInvitation()") is None
            cases["unrelated_hash_and_query_preserved_without_import"] = True

            page.evaluate("key=>localStorage.setItem(key,'{invalid json')", KEY)
            page.reload()
            assert page.evaluate("ToolTeamConfig.inspect().configured") is False
            page.evaluate("key=>localStorage.removeItem(key)", KEY)
            page.reload()
            page.evaluate("value=>ToolTeamConfig.remember(value)", CONFIG)
            page.evaluate("()=>{Storage.prototype.setItem=function(){throw new DOMException('quota','QuotaExceededError')}}")
            assert page.evaluate("value=>{try{ToolTeamConfig.remember(value);return false}catch(e){return e.code==='team-storage'}}", {**CONFIG, "name": "replacement"})
            assert page.evaluate("ToolTeamConfig.inspect().configuration") == CONFIG
            assert page.evaluate("ToolTeamConfig.inspect().persistent") is False
            cases["invalid_stored_data_and_quota_fail_without_replacing_current_config"] = True

            page.evaluate("source=>{let f=document.createElement('iframe');f.sandbox='allow-scripts';f.srcdoc='<script>'+source.replace(/<\\/script/gi,'<\\\\/script')+'</script>';document.body.append(f)}", source)
            page.wait_for_function("frames.length===1")
            frame = page.frames[-1]
            frame.wait_for_function("window.ToolTeamConfig")
            assert frame.evaluate("ToolTeamConfig.inspect().supported") is False
            assert frame.evaluate("value=>{try{ToolTeamConfig.create(value);return false}catch(e){return e.code==='team-unsupported'}}", CONFIG)
            assert frame.evaluate("ToolTeamConfig.loadPublished()") is None
            cases["opaque_origin_cannot_provision_and_does_not_fetch_published_settings"] = True
            assert not errors, errors
            assert not external, external
            cases["no_uncaught_errors_or_network_requests_to_workbook_or_other_origins"] = True
            context.close()
            browser.close()
    print(json.dumps({"cases": cases, "real_company_credentials_used": False}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
