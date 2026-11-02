"""Owner-configured shared workbook transport with real Graph/IDB protocols.

The settings provider and HTTP Graph service are controlled fixtures. The
production shared transport, Graph transport, backups, client isolation and
OOXML byte operations run unchanged. No company account or data is used.
"""
from __future__ import annotations

import argparse
import base64
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import tempfile
from urllib.parse import urlsplit

from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright

from browser_test import REPO, serve_program
from graph_file_test import Cloud, SETUP, TOKEN, TICKET


WORKSPACE = "11111111-1111-4111-8111-111111111111"
OWNER = "22222222-2222-4222-8222-222222222222"


def workbook_bytes():
    memory = BytesIO()
    book = Workbook()
    book.active.append(["Serial", "Cal Last", "Cal Next"])
    book.active.append(["00123456", "2026-01-01", "2027-01-01"])
    book.active.append(["00234567", "2026-02-01", "2027-02-01"])
    book.save(memory)
    return memory.getvalue()


class TargetCloud(Cloud):
    def __init__(self, original, suffix):
        super().__init__(original)
        self.target = {"driveId": "drive-" + suffix, "itemId": "item-" + suffix, "name": "shared-" + suffix + ".xlsx"}
        self.forced_412 = 0

    def item(self):
        result = super().item()
        result.update({"id": self.target["itemId"], "name": self.target["name"],
                       "parentReference": {"driveId": self.target["driveId"]},
                       "@microsoft.graph.downloadUrl": "https://download.example.test/document/" + self.target["itemId"] + "?ticket=" + TICKET})
        return result

    def route(self, route):
        if route.request.method == "PUT" and self.forced_412:
            self.forced_412 -= 1
            self.puts.append({"if_match": route.request.headers.get("if-match"), "forced_conflict": True})
            self.revision += 1
            return route.fulfill(status=412, content_type="application/json", headers={"access-control-allow-origin": "*"}, body="{}")
        return super().route(route)


def workspace(target, revision='"settings-1"'):
    return {"v": 1, "workspaceId": WORKSPACE, "ownerObjectId": OWNER, "workbook": target.copy(), "configVersion": revision}


HELPERS = """()=>{
 window.testWorkspace=null;window.testSettingsReads=0;window.testSettingsFailure='';window.testSettingsFailureStatus=0;
 window.ToolSharedWorkspace={refresh:async()=>{
   testSettingsReads++;
   if(testSettingsFailure){const error=new Error('TEST_SETTINGS_FAILURE');error.code=testSettingsFailure;if(testSettingsFailureStatus)error.httpStatus=testSettingsFailureStatus;throw error;}
   return JSON.parse(JSON.stringify(testWorkspace));
 }};
 window.setCell=async(bytes,reference,value)=>{
   const zip=await JSZip.loadAsync(bytes),entry='xl/worksheets/sheet1.xml';
   const doc=new DOMParser().parseFromString(await zip.file(entry).async('string'),'application/xml');
   const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
   const cell=Array.from(doc.getElementsByTagNameNS('*','c')).find(c=>c.getAttribute('r')===reference);
   if(!cell)throw Error('missing fixture cell');
   while(cell.firstChild)cell.removeChild(cell.firstChild);cell.setAttribute('t','inlineStr');
   const item=doc.createElementNS(ns,'is'),text=doc.createElementNS(ns,'t');text.textContent=value;item.append(text);cell.append(item);
   zip.file(entry,new XMLSerializer().serializeToString(doc));
   return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
 };
 window.tryShared=async operation=>{try{return {ok:true,value:await operation()};}catch(error){return {ok:false,code:error.code,httpStatus:error.httpStatus,message:error.message,uncertain:ToolSharedFile.inspect().uncertain};}};
}"""


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    args = parser.parse_args()
    original = workbook_bytes()
    cases, errors, unexpected = {}, [], []
    with tempfile.TemporaryDirectory(prefix="tool-shared-file-") as directory:
        shell = Path(directory) / "index.html"
        shell.write_text('<!doctype html><html><head><meta charset="utf-8"></head><body></body></html>')
        with serve_program(shell) as url, sync_playwright() as playwright:
            browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
            clouds = {key: TargetCloud(original, key) for key in ["a", "b"]}

            def context():
                value = browser.new_context()

                def route(request_route):
                    request = request_route.request
                    parsed = urlsplit(request.url)
                    for cloud in clouds.values():
                        target = cloud.target
                        if parsed.hostname == "graph.microsoft.com" and parsed.path.startswith("/v1.0/drives/" + target["driveId"] + "/items/" + target["itemId"]) or parsed.hostname == "download.example.test" and parsed.path == "/document/" + target["itemId"]:
                            return cloud.route(request_route)
                    if parsed.hostname in {"graph.microsoft.com", "download.example.test"}:
                        unexpected.append(request.url)
                        return request_route.abort()
                    return request_route.continue_()

                value.route("**/*", route)
                return value

            def page_for(ctx, target="a"):
                page = ctx.new_page()
                page.on("pageerror", lambda error: errors.append(str(error)))
                page.goto(url)
                page.evaluate(SETUP)
                page.add_script_tag(path=REPO / "web/vendor/jszip.min.js")
                page.add_script_tag(path=REPO / "web/graph-file.js")
                page.evaluate(HELPERS)
                page.evaluate("value=>testWorkspace=value", workspace(clouds[target].target))
                page.add_script_tag(path=REPO / "web/shared-file.js")
                page.evaluate("encoded=>window.originalBytes=Uint8Array.from(atob(encoded),c=>c.charCodeAt(0))", base64.b64encode(original).decode())
                return page

            def connect(page):
                return page.evaluate("""async()=>{
                  const result=await ToolSharedFile.connect({target:{driveId:'ignored',itemId:'ignored'},shareUrl:'https://ignored.invalid/workbook.xlsx',validate:async file=>{await JSZip.loadAsync(await file.arrayBuffer());}});
                  window.before=await ToolSharedFile.read();return {connected:result.connected,inspect:ToolSharedFile.inspect()};
                }""")

            def no_extra_write_error(page, expression, expected, cloud="a"):
                prior = len(clouds[cloud].puts)
                result = page.evaluate("async()=>await tryShared(" + expression + ")")
                assert not result["ok"] and result.get("code") == expected and len(clouds[cloud].puts) == prior, {
                    key: result.get(key) for key in ["ok", "code", "httpStatus", "message", "uncertain"]
                }
                return result

            # A persisted local selection is ignored: the cloud settings alone
            # choose which exact workbook is read by connect and restore.
            ctx = context(); page = page_for(ctx)
            page.evaluate("async target=>{await ToolGraphFile.connect({target});window.graphRestores=0;ToolGraphFile=Object.freeze({...ToolGraphFile,restore:async()=>{graphRestores++;throw Error('stale restore must not run')}})}", clouds["b"].target)
            b_reads = len(clouds["b"].calls)
            result = connect(page)
            assert result["connected"] and result["inspect"]["target"] == clouds["a"].target
            assert result["inspect"]["sourceMode"] == "shared" and result["inspect"]["mode"] == "shared"
            page.evaluate("async()=>await ToolSharedFile.restore()")
            assert page.evaluate("graphRestores") == 0 and len(clouds["b"].calls) == b_reads
            cases["owner_target_overrides_local_and_caller_selection"] = True

            # Returned target structures and receipt bytes cannot mutate the
            # internal destination or immutable Graph snapshot.
            receipt = page.evaluate("""async()=>{
              const exposed=ToolSharedFile.inspect();exposed.target.itemId='forged';
              const bytes=await setCell(before.bytes,'C2','2028-01-01');
              const result=await ToolSharedFile.save({bytes,expectedVersion:before.version});
              const savedHash=await testHash(result.bytes);result.bytes[0]=0;result.target.itemId='forged';
              const current=await ToolSharedFile.read();
              return {savedHash,currentHash:await testHash(current.bytes),target:ToolSharedFile.inspect().target,configVersion:result.configVersion,rebaseCount:result.rebaseCount};
            }""")
            assert receipt["savedHash"] == receipt["currentHash"] == sha256(clouds["a"].data).hexdigest()
            assert receipt["target"] == clouds["a"].target and receipt["rebaseCount"] == 0
            cases["verified_receipt_bytes_and_target_are_immutable_copies"] = True

            # Owner changes are fetched on every read, including after an
            # application reload; no stale selected-file slot is restored.
            page.evaluate("value=>testWorkspace=value", workspace(clouds["b"].target, '"settings-2"'))
            switched = page.evaluate("async()=>{const current=await ToolSharedFile.read();return {changed:current.bindingChanged,target:current.target,inspect:ToolSharedFile.inspect()}}")
            assert switched["changed"] and switched["target"] == clouds["b"].target
            page.close(); page = page_for(ctx, "b")
            assert connect(page)["inspect"]["target"] == clouds["b"].target
            cases["owner_path_change_applies_on_next_read_and_restart"] = True
            ctx.close()

            # Two devices read the same version. A second safe rebase merges
            # another row into the already updated shared workbook.
            clouds["a"] = TargetCloud(original, "a")
            first_ctx, second_ctx = context(), context()
            first, second = page_for(first_ctx), page_for(second_ctx)
            connect(first); connect(second)
            first.evaluate("async()=>{const bytes=await setCell(before.bytes,'C2','2028-01-01');return ToolSharedFile.save({bytes,expectedVersion:before.version});}")
            second_result = second.evaluate("""async()=>{
              window.rebases=0;const bytes=await setCell(before.bytes,'C3','2028-02-01');
              const saved=await ToolSharedFile.save({bytes,expectedVersion:before.version,rebase:async current=>{rebases++;return {bytes:await setCell(current.bytes,'C3','2028-02-01')};}});
              return {rebases,receiptRebases:saved.rebaseCount,hash:await testHash(saved.bytes)};
            }""")
            assert second_result["rebases"] == second_result["receiptRebases"] == 1
            assert second_result["hash"] == sha256(clouds["a"].data).hexdigest()
            book = load_workbook(BytesIO(clouds["a"].data))
            assert book.active["C2"].value == "2028-01-01" and book.active["C3"].value == "2028-02-01"
            book.close()
            cases["two_device_disjoint_updates_are_merged_into_one_workbook"] = True
            first_ctx.close(); second_ctx.close()

            def scenario():
                clouds["a"] = TargetCloud(original, "a")
                ctx = context(); page = page_for(ctx); connect(page)
                page.evaluate("async()=>window.updateBytes=await setCell(before.bytes,'C2','2028-01-01')")
                return ctx, page, clouds["a"]

            ctx, page, cloud = scenario()
            page.evaluate("value=>testWorkspace=value", workspace(clouds["b"].target, '"settings-2"'))
            no_extra_write_error(page, "()=>ToolSharedFile.save({bytes:updateBytes,expectedVersion:before.version})", "shared-destination-changed")
            cases["changed_destination_blocks_old_file_write"] = True
            ctx.close()

            ctx, page, cloud = scenario()
            page.evaluate("()=>testWorkspace.configVersion='\"settings-2\"'")
            no_extra_write_error(page, "()=>ToolSharedFile.save({bytes:updateBytes,expectedVersion:before.version})", "shared-destination-changed")
            cases["changed_configuration_revision_blocks_stale_write"] = True
            ctx.close()

            # A late owner change after backup, before PUT, must be checked by
            # the production Graph adapter's asynchronous destination hook.
            ctx, page, cloud = scenario()
            page.evaluate("""()=>{
              const originalSave=ToolGraphFile.save;
              ToolGraphFile=Object.freeze({...ToolGraphFile,save:options=>originalSave({...options,verifyDestination:async()=>{testWorkspace.configVersion='\"settings-late\"';return options.verifyDestination();}})});
            }""")
            no_extra_write_error(page, "()=>ToolSharedFile.save({bytes:updateBytes,expectedVersion:before.version})", "shared-destination-changed")
            cases["destination_rechecked_after_backup_before_upload"] = True
            ctx.close()

            # The owner can change a separate settings item after the final
            # preflight and while the workbook PUT is in flight. Keep the
            # verified receipt for A, flag the new binding, and never copy it
            # into B or report the already-committed upload as a failure.
            ctx, page, cloud = scenario()
            b_puts = len(clouds["b"].puts)
            page.evaluate("""next=>{
              const originalFetch=window.fetch;
              window.fetch=async(input,options)=>{
                const pending=originalFetch(input,options);
                if(options&&options.method==='PUT')testWorkspace=next;
                return pending;
              };
            }""", workspace(clouds["b"].target, '"settings-inflight"'))
            inflight = page.evaluate("""async()=>{
              const saved=await ToolSharedFile.save({bytes:updateBytes,expectedVersion:before.version,rebase:()=>{throw Error('must not repeat a confirmed commit')}});
              return {target:saved.target,hash:await testHash(saved.bytes),changed:saved.destinationChanged,verified:saved.configurationVerified,inspect:ToolSharedFile.inspect()};
            }""")
            assert inflight["changed"] and inflight["verified"] and inflight["target"] == cloud.target
            assert inflight["hash"] == sha256(cloud.data).hexdigest() and len(cloud.puts) == 1 and len(clouds["b"].puts) == b_puts
            assert not inflight["inspect"]["connected"] and inflight["inspect"]["needsReconnect"] and inflight["inspect"]["target"] == cloud.target
            cases["inflight_owner_change_retains_actual_saved_target"] = True
            ctx.close()

            # A post-commit settings 401 cannot turn a verified workbook write
            # into a retryable upload error or clear the actual receipt.
            ctx, page, cloud = scenario()
            page.evaluate("""()=>{
              const originalFetch=window.fetch;
              window.fetch=async(input,options)=>{
                const pending=originalFetch(input,options);
                if(options&&options.method==='PUT'){testSettingsFailure='auth-expired';testSettingsFailureStatus=401;}
                return pending;
              };
            }""")
            post_auth = page.evaluate("""async()=>{
              const saved=await ToolSharedFile.save({bytes:updateBytes,expectedVersion:before.version,rebase:()=>{throw Error('must not repeat a confirmed commit')}});
              return {target:saved.target,hash:await testHash(saved.bytes),changed:saved.destinationChanged,verified:saved.configurationVerified,requiresSignIn:saved.requiresSignIn,inspect:ToolSharedFile.inspect()};
            }""")
            assert not post_auth["verified"] and not post_auth["changed"] and post_auth["requiresSignIn"]
            assert post_auth["target"] == cloud.target and post_auth["hash"] == sha256(cloud.data).hexdigest() and len(cloud.puts) == 1
            assert not post_auth["inspect"]["connected"] and post_auth["inspect"]["needsReconnect"] and post_auth["inspect"]["requiresSignIn"]
            no_extra_write_error(page, "()=>ToolSharedFile.save({bytes:updateBytes,expectedVersion:ToolSharedFile.inspect().version})", "auth-expired")
            cases["post_commit_settings_401_keeps_receipt_without_retry"] = True
            ctx.close()

            ctx, page, cloud = scenario()
            cloud.forced_412 = 1
            retried = page.evaluate("""async()=>{
              const saved=await ToolSharedFile.save({bytes:updateBytes,expectedVersion:before.version,rebase:current=>setCell(current.bytes,'C2','2028-01-01')});
              return {rebaseCount:saved.rebaseCount,hash:await testHash(saved.bytes)};
            }""")
            assert retried["rebaseCount"] == 1 and len(cloud.puts) == 2 and retried["hash"] == sha256(cloud.data).hexdigest()
            cases["http_412_rebases_only_after_fresh_shared_snapshot"] = True
            ctx.close()

            ctx, page, cloud = scenario()
            cloud.forced_412 = 9
            exhausted = page.evaluate("""async()=>{
              window.rebases=0;const result=await tryShared(()=>ToolSharedFile.save({bytes:updateBytes,expectedVersion:before.version,rebase:async current=>{rebases++;return setCell(current.bytes,'C2','2028-01-01');}}));
              return {...result,rebases};
            }""")
            assert exhausted["code"] == "conflict" and exhausted["rebases"] == 2 and len(cloud.puts) == 3
            cases["conflict_retries_are_bounded_to_two"] = True
            ctx.close()

            ctx, page, cloud = scenario()
            cloud.forced_412 = 1
            denied = page.evaluate("""async()=>await tryShared(()=>ToolSharedFile.save({bytes:updateBytes,expectedVersion:before.version,rebase:()=>{const e=Error('Another user changed the same tool');e.code='shared-row-conflict';throw e;}}))""")
            assert denied["code"] == "shared-row-conflict" and len(cloud.puts) == 1 and cloud.data == original
            cases["same_row_conflict_callback_stops_without_overwrite"] = True
            ctx.close()

            for status in [401, 403, 409, 423, 429]:
                ctx, page, cloud = scenario(); cloud.mode = "put" + str(status)
                result = page.evaluate("""async()=>{window.rebases=0;const value=await tryShared(()=>ToolSharedFile.save({bytes:updateBytes,expectedVersion:before.version,rebase:()=>{rebases++;return updateBytes;}}));return {...value,rebases};}""")
                assert not result["ok"] and result["rebases"] == 0 and len(cloud.puts) == 1 and cloud.data == original
                cases["non_conflict_" + str(status) + "_not_retried"] = True
                ctx.close()

            ctx, page, cloud = scenario(); cloud.mode = "commit-then-network-error"
            first_error = page.evaluate("async()=>await tryShared(()=>ToolSharedFile.save({bytes:updateBytes,expectedVersion:before.version,rebase:()=>{throw Error('must not rebase unknown commit')}}))")
            assert first_error["uncertain"] and len(cloud.puts) == 1
            no_extra_write_error(page, "()=>ToolSharedFile.save({bytes:updateBytes,expectedVersion:before.version})", "requires-read")
            cloud.mode = "normal"
            restored = page.evaluate("async()=>{await ToolSharedFile.read();return ToolSharedFile.inspect()}")
            assert not restored["uncertain"]
            cases["uncertain_commit_held_until_verified_read"] = True
            ctx.close()

            ctx, page, cloud = scenario()
            for action in ["read", "save"]:
                page.evaluate("()=>testSettingsFailure='workspace-offline'")
                before_calls = len(cloud.calls)
                result = page.evaluate("async action=>await tryShared(()=>action==='read'?ToolSharedFile.read():ToolSharedFile.save({bytes:updateBytes,expectedVersion:before.version}))", action)
                assert result["code"] == "workspace-offline" and len(cloud.calls) == before_calls and not cloud.puts
            assert not page.evaluate("ToolSharedFile.inspect().connected")
            cases["settings_failure_never_falls_back_to_remembered_file"] = True
            ctx.close()

            ctx, page, cloud = scenario()
            stale = page.evaluate("""async()=>{
              const result=await tryShared(()=>ToolSharedFile.save({bytes:updateBytes,expectedVersion:before.version,verifyCurrent:()=>false,rebase:()=>{throw Error('must not rebase stale UI')}}));
              return result;
            }""")
            assert stale["code"] == "stale" and not cloud.puts
            cases["changed_photo_revision_never_rebased"] = True
            ctx.close()
            browser.close()
    assert not errors and not unexpected, (errors, unexpected)
    print(json.dumps({"checks": cases, "case_count": len(cases), "page_errors": errors, "unexpected_requests": unexpected,
                      "scope": "Real Chromium shared/Graph adapters, IndexedDB and OOXML; controlled settings/auth/HTTP, no real company account."}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
