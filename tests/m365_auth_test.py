"""Browser tests for Microsoft SPA PKCE boundaries, using a controlled issuer.

No real Microsoft account, credentials or company files are used. The browser
performs actual Web Crypto, navigation, storage and CORS fetch operations; only
Microsoft endpoint responses are intercepted by the test.
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
from pathlib import Path
import tempfile
import time
from urllib.parse import parse_qs, urlencode, urlsplit

from playwright.sync_api import sync_playwright

from browser_test import REPO, serve_program

CLIENT = "00000000-0000-4000-8000-000000000001"
TENANT = "11111111-1111-4111-8111-111111111111"
PENDING = "tool-manager-m365-pkce-v1"
CONFIG = "tool-manager-m365-config-v1"
SHARED_PENDING = "tool-manager-shared-pkce-v1"
SHARED_CONFIG = "tool-manager-shared-auth-config-v1"
ACCESS = "controlled-access-token-not-a-real-credential"


def jwt(claims):
    part = lambda value: base64.urlsafe_b64encode(json.dumps(value, separators=(",", ":")).encode()).decode().rstrip("=")
    return part({"alg": "RS256", "typ": "JWT"}) + "." + part(claims) + ".controlled-signature"


def error(page, expression):
    return page.evaluate("""async expression => {
      try { await eval(expression); return {unexpectedSuccess:true}; }
      catch(error) { return {code:error.code,message:error.message,oAuthError:error.oAuthError}; }
    }""", expression)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    args = parser.parse_args()
    source = (REPO / "web/m365-auth.js").read_text()
    cases = {}
    with tempfile.TemporaryDirectory(prefix="m365-auth-") as directory:
        app = Path(directory) / "index.html"
        app.write_text('<!doctype html><meta charset="utf-8"><title>Auth boundary test</title><script>' + source.replace("</script", "<\\/script") + "</script>")
        (Path(directory) / "shared.html").write_text('<!doctype html><meta charset="utf-8"><meta name="tool-manager-shared-entry" content="true"><title>Shared auth test</title><script>' + source.replace("</script", "<\\/script") + "</script>")
        with serve_program(app) as url, sync_playwright() as playwright:
            browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])

            def open_case(shared=False, legacy_state=None):
                context = browser.new_context()
                if legacy_state is not None:
                    context.add_init_script("""(({config,pending})=>{
                      if(!sessionStorage.getItem('auth-legacy-fixture-seeded')){
                        localStorage.setItem('tool-manager-m365-config-v1',config);
                        sessionStorage.setItem('tool-manager-m365-pkce-v1',pending);
                        sessionStorage.setItem('auth-legacy-fixture-seeded','true');
                      }
                    })(""" + json.dumps(legacy_state) + ");")
                page = context.new_page()
                case_url = url.replace("index.html", "shared.html") if shared else url
                traffic = {"authorize": [], "tokens": [], "unexpected": [], "claims": {}, "status": 200, "network_fail": False,
                           "app_url": case_url, "pending_key": SHARED_PENDING if shared else PENDING, "shared": shared}

                def endpoint(route):
                    request = route.request
                    if request.url.startswith("https://login.microsoftonline.com/") and request.url.split("?", 1)[0].endswith("/authorize"):
                        traffic["authorize"].append(request.url)
                        return route.fulfill(status=200, content_type="text/html", body="<!doctype html><title>Controlled issuer</title>")
                    if request.url.startswith("https://login.microsoftonline.com/") and request.url.endswith("/token"):
                        traffic["tokens"].append({"url": request.url, "body": parse_qs(request.post_data or ""), "headers": request.headers})
                        if traffic["network_fail"]:
                            return route.abort()
                        authorization = parse_qs(urlsplit(traffic["authorize"][-1]).query)
                        claims = {"aud": CLIENT, "tid": TENANT, "iss": "https://login.microsoftonline.com/" + TENANT + "/v2.0",
                                  "nonce": authorization["nonce"][0], "exp": int(time.time()) + 3600,
                                  "name": "Test Engineer", "preferred_username": "engineer@example.invalid"}
                        claims.update(traffic["claims"])
                        data = {"token_type": "Bearer", "expires_in": 3600, "access_token": ACCESS,
                                "scope": traffic.get("scope", "openid profile Files.ReadWrite.All"), "id_token": jwt(claims)}
                        data.update(traffic.get("data", {}))
                        return route.fulfill(status=traffic["status"], content_type="application/json", headers={"Access-Control-Allow-Origin": "*"}, body=json.dumps(data))
                    if urlsplit(request.url).netloc != urlsplit(url).netloc:
                        traffic["unexpected"].append(request.url)
                        return route.abort()
                    return route.continue_()

                context.route("**/*", endpoint)
                page.goto(case_url)
                page.evaluate("config=>ToolM365Auth.configure(config)", {"clientId": CLIENT, "tenantId": TENANT})
                return context, page, traffic

            def begin(page, traffic, options=None):
                with page.expect_request("https://login.microsoftonline.com/**/authorize?*"):
                    if options is None:
                        page.evaluate("ToolM365Auth.signIn()")
                    else:
                        page.evaluate("options=>ToolM365Auth.signIn(options)", options)
                page.wait_for_url("https://login.microsoftonline.com/**")
                page.goto(traffic["app_url"])
                pending = page.evaluate("key=>JSON.parse(sessionStorage.getItem(key))", traffic["pending_key"])
                authorization = parse_qs(urlsplit(traffic["authorize"][-1]).query)
                assert len(pending["verifier"]) == 43
                assert authorization["code_challenge_method"] == ["S256"]
                expected = base64.urlsafe_b64encode(hashlib.sha256(pending["verifier"].encode()).digest()).decode().rstrip("=")
                assert authorization["code_challenge"] == [expected]
                assert authorization["redirect_uri"] == [traffic["app_url"]]
                assert "offline_access" not in authorization["scope"][0]
                assert "client_secret" not in authorization
                assert authorization["state"] == [pending["state"]]
                assert authorization["nonce"] == [pending["nonce"]]
                if options is None:
                    assert ("prompt" not in authorization) if traffic["shared"] else authorization["prompt"] == ["select_account"]
                elif options.get("prompt"):
                    assert authorization["prompt"] == [options["prompt"]]
                else:
                    assert "prompt" not in authorization
                return pending

            def callback(page, pending, query=None, fragment=""):
                page.goto(pending["configuration"]["redirectUri"] + "?" + urlencode(query or {"code": "controlled-authorization-code", "state": pending["state"]}) + fragment)
                assert "code=" not in page.url and "state=" not in page.url and "access_token=" not in page.url

            context, page, traffic = open_case()
            pending = begin(page, traffic)
            callback(page, pending)
            result = page.evaluate("ToolM365Auth.initialize()")
            assert result["signedIn"] and result["account"]["username"] == "engineer@example.invalid"
            assert page.evaluate("ToolM365Auth.getAccessToken()") == ACCESS
            assert page.evaluate("key=>sessionStorage.getItem(key)", PENDING) is None
            stored = page.evaluate("JSON.stringify({local:{...localStorage},session:{...sessionStorage}})")
            assert ACCESS not in stored and "controlled-signature" not in stored and pending["verifier"] not in stored
            body = traffic["tokens"][0]["body"]
            assert body["code_verifier"] == [pending["verifier"]] and body["client_id"] == [CLIENT]
            assert "client_secret" not in body and "offline_access" not in body["scope"][0]
            assert traffic["tokens"][0]["headers"].get("content-type") == "application/x-www-form-urlencoded"
            page.evaluate("ToolM365Auth.initialize()")
            assert len(traffic["tokens"]) == 1, "Callback must be consumed once"
            page.evaluate("ToolM365Auth.signOut()")
            assert error(page, "ToolM365Auth.getAccessToken()")["code"] == "auth-required"
            assert not traffic["unexpected"]
            cases["pkce_success_memory_only_logout_one_time"] = True
            context.close()

            for label, updates, expected in [
                ("wrong_state", {"query": {"code": "test", "state": "wrong-state"}}, "auth-state"),
                ("expired_pending", {"age": 11 * 60 * 1000}, "auth-state"),
                ("future_pending", {"future": 60000}, "auth-state"),
                ("duplicate_code", {"duplicate": True}, "auth-state"),
                ("provider_denied", {"query": {"error": "access_denied", "error_description": "Untrusted <script>secret</script>"}}, "auth-denied"),
                ("wrong_nonce", {"claims": {"nonce": "wrong"}}, "auth-response"),
                ("wrong_audience", {"claims": {"aud": "another-app"}}, "auth-response"),
                ("wrong_issuer", {"claims": {"iss": "https://untrusted.invalid/tenant/v2.0"}}, "auth-response"),
                ("wrong_tenant", {"claims": {"tid": "22222222-2222-4222-8222-222222222222", "iss": "https://login.microsoftonline.com/22222222-2222-4222-8222-222222222222/v2.0"}}, "auth-response"),
                ("missing_file_permission", {"scope": "openid profile User.Read"}, "auth-permission"),
                ("bad_token_type", {"data": {"token_type": "MAC"}}, "auth-response"),
                ("bad_expiry", {"data": {"expires_in": "NaN"}}, "auth-response"),
                ("provider_failure", {"status": 400}, "auth-consent"),
                ("network_failure", {"network_fail": True}, "auth-network")
            ]:
                context, page, traffic = open_case()
                pending = begin(page, traffic)
                if "age" in updates or "future" in updates:
                    page.evaluate("({key,age,future})=>{let p=JSON.parse(sessionStorage.getItem(key));p.createdAt=Date.now()-age+future;sessionStorage.setItem(key,JSON.stringify(p))}", {"key": PENDING, "age": updates.get("age", 0), "future": updates.get("future", 0)})
                traffic.update({key: value for key, value in updates.items() if key in {"claims", "scope", "data", "status", "network_fail"}})
                query = updates.get("query", {"code": "test"}).copy()
                query.setdefault("state", pending["state"])
                if updates.get("duplicate"):
                    page.goto(url + "?code=one&code=two&state=" + pending["state"])
                    assert "code=" not in page.url
                else:
                    callback(page, pending, query)
                failure = error(page, "ToolM365Auth.initialize()")
                assert failure.get("code") == expected, (label, failure)
                assert "Untrusted" not in failure["message"] and "<script>" not in failure["message"]
                assert not page.evaluate("ToolM365Auth.inspect().signedIn")
                assert page.evaluate("key=>sessionStorage.getItem(key)", PENDING) is None
                if expected in {"auth-state", "auth-denied"}:
                    assert not traffic["tokens"], (label, "An invalid callback reached the token endpoint")
                assert not traffic["unexpected"]
                cases[label] = True
                context.close()

            context, page, traffic = open_case()
            pending = begin(page, traffic)
            page.goto(url + "?fragment-return=1#access_token=untrusted-token&state=" + pending["state"])
            assert "access_token" not in page.url
            assert error(page, "ToolM365Auth.initialize()")["code"] == "auth-state"
            assert not traffic["tokens"]
            cases["implicit_fragment_rejected_and_scrubbed"] = True
            context.close()

            context, page, traffic = open_case()
            for configuration in [
                {"clientId": "not-a-guid", "tenantId": TENANT},
                {"clientId": CLIENT, "tenantId": "untrusted.invalid/path"},
                {"clientId": CLIENT, "tenantId": TENANT, "redirectUri": "https://untrusted.invalid/index.html"},
                {"clientId": CLIENT, "tenantId": TENANT, "redirectUri": url + "?token=x"},
                {"clientId": CLIENT, "tenantId": TENANT, "redirectUri": url.replace("index.html", "other.html")}
            ]:
                denied = page.evaluate("value=>{try{ToolM365Auth.configure(value);return false}catch(e){return !!e.code}}", configuration)
                assert denied
            for tenant in ["organizations", "example.onmicrosoft.com", TENANT]:
                assert page.evaluate("value=>ToolM365Auth.configure(value).configured", {"clientId": CLIENT, "tenantId": tenant})
            cases["authority_and_same_path_redirect_validation"] = True
            context.close()

            context, page, traffic = open_case()
            pending = begin(page, traffic)
            traffic["data"] = {"expires_in": 10}
            callback(page, pending)
            page.evaluate("ToolM365Auth.initialize()")
            assert error(page, "ToolM365Auth.getAccessToken()")["code"] == "auth-expired"
            assert not page.evaluate("ToolM365Auth.inspect().signedIn")
            cases["near_expiry_requires_explicit_login"] = True
            context.close()

            # The common-workbook entry owns separate configuration and PKCE
            # records. Existing working desktop/individual app settings must
            # remain byte-identical through the shared app's login and logout.
            legacy_state = {
                "config": json.dumps({"clientId": "99999999-9999-4999-8999-999999999999", "tenantId": TENANT,
                                      "redirectUri": url}, separators=(",", ":")),
                "pending": json.dumps({"legacyFlow": "do-not-replace", "createdAt": 123}, separators=(",", ":"))
            }
            context, page, traffic = open_case(shared=True, legacy_state=legacy_state)
            assert page.evaluate("key=>localStorage.getItem(key)", CONFIG) == legacy_state["config"]
            assert page.evaluate("key=>sessionStorage.getItem(key)", PENDING) == legacy_state["pending"]
            assert page.evaluate("key=>JSON.parse(localStorage.getItem(key)).redirectUri", SHARED_CONFIG) == traffic["app_url"]
            pending = begin(page, traffic)
            assert "prompt" not in parse_qs(urlsplit(traffic["authorize"][-1]).query)
            assert page.evaluate("key=>localStorage.getItem(key)", CONFIG) == legacy_state["config"]
            assert page.evaluate("key=>sessionStorage.getItem(key)", PENDING) == legacy_state["pending"]
            traffic["claims"] = {"oid": "ABCDEF01-1234-4567-89AB-CDEF01234567"}
            callback(page, pending)
            result = page.evaluate("ToolM365Auth.initialize()")
            assert result["signedIn"] and result["account"]["objectId"] == "abcdef01-1234-4567-89ab-cdef01234567"
            assert page.evaluate("ToolM365Auth.getAccessToken()") == ACCESS
            assert page.evaluate("key=>sessionStorage.getItem(key)", SHARED_PENDING) is None
            page.evaluate("ToolM365Auth.signOut()")
            assert page.evaluate("key=>localStorage.getItem(key)", CONFIG) == legacy_state["config"]
            assert page.evaluate("key=>sessionStorage.getItem(key)", PENDING) == legacy_state["pending"]
            assert page.evaluate("key=>!!localStorage.getItem(key)", SHARED_CONFIG)
            stored = page.evaluate("JSON.stringify({local:{...localStorage},session:{...sessionStorage}})")
            assert ACCESS not in stored and "controlled-signature" not in stored and pending["verifier"] not in stored
            assert not traffic["unexpected"]
            cases["shared_scope_preserves_root_config_and_pending_sso_oid_logout"] = True
            context.close()

            # A silent SSO response can legitimately ask for interaction. It
            # is not an application-path/administrator setup failure, and an
            # automatic attempt must not redeem an error at the token endpoint.
            for oauth_error in ["interaction_required", "login_required", "consent_required", "account_selection_required"]:
                context, page, traffic = open_case(shared=True)
                pending = begin(page, traffic, {"prompt": "none"})
                callback(page, pending, {"error": oauth_error, "error_description": "Untrusted <script>secret</script>", "state": pending["state"]})
                failure = error(page, "ToolM365Auth.initialize()")
                assert failure["code"] == "auth-interaction-required" and failure["oAuthError"] == oauth_error, failure
                assert failure["message"] == "회사 계정으로 계속하려면 로그인을 확인해 주세요."
                assert "Untrusted" not in failure["message"] and "error=" not in page.url and "state=" not in page.url
                assert page.evaluate("key=>sessionStorage.getItem(key)", SHARED_PENDING) is None
                assert not traffic["tokens"] and not traffic["unexpected"]
                cases["silent_" + oauth_error + "_requires_interaction_without_token_exchange"] = True
                context.close()

            context, page, traffic = open_case(shared=True)
            pending = begin(page, traffic, {"prompt": "none"})
            authorization_count = len(traffic["authorize"])
            pending_bytes = page.evaluate("key=>sessionStorage.getItem(key)", SHARED_PENDING)
            for options in [None, "none", [], {"prompt": "none select_account"}, {"prompt": "unknown"}, {"prompt": False}]:
                denied = page.evaluate("""async options=>{
                  try{await ToolM365Auth.signIn(options);return {unexpectedSuccess:true};}
                  catch(error){return {code:error.code};}
                }""", options)
                assert denied.get("code") == "auth-prompt", denied
                assert page.evaluate("key=>sessionStorage.getItem(key)", SHARED_PENDING) == pending_bytes
                assert page.url == traffic["app_url"] and len(traffic["authorize"]) == authorization_count
            assert not traffic["tokens"] and not traffic["unexpected"]
            cases["invalid_prompt_keeps_existing_pkce_and_does_not_navigate"] = True
            context.close()

            for shared in [False, True]:
                for prompt in ["", "none", "select_account", "login", "consent"]:
                    context, page, traffic = open_case(shared=shared)
                    pending = begin(page, traffic, {"prompt": prompt})
                    callback(page, pending)
                    assert page.evaluate("ToolM365Auth.initialize().then(value=>value.signedIn)")
                    assert not traffic["unexpected"]
                    context.close()
            cases["explicit_prompt_allowlist_retains_full_pkce_success"] = True

            context, page, traffic = open_case(shared=True)
            pending = begin(page, traffic)
            traffic["claims"] = {"oid": "not-a-directory-object-id"}
            callback(page, pending)
            assert error(page, "ToolM365Auth.initialize()")["code"] == "auth-response"
            assert not page.evaluate("ToolM365Auth.inspect().signedIn")
            assert not traffic["unexpected"]
            cases["invalid_present_directory_object_id_rejected"] = True
            context.close()

            context, page, traffic = open_case(shared=True)
            pending = begin(page, traffic)
            callback(page, pending, {"error": "invalid_client", "state": pending["state"]})
            failure = error(page, "ToolM365Auth.initialize()")
            assert failure["code"] == "auth-response" and failure["oAuthError"] == "invalid_client"
            assert not traffic["tokens"] and not traffic["unexpected"]
            cases["other_oauth_error_keeps_safe_provider_code"] = True
            context.close()
            browser.close()
    print(json.dumps({"cases": cases, "case_count": len(cases), "real_microsoft_account_used": False,
                      "security_scope": "Actual browser crypto/navigation/storage/CORS; controlled issuer responses"}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
