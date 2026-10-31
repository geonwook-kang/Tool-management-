"""Real browser IDB/backups and controlled Graph HTTP workbook protocol.

The supplied XLSX remains read only. Microsoft Graph, a short-lived download
URL and authentication are controlled HTTP fixtures; this does not assert that
a company tenant has granted access. No tenant credentials are used or stored.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import tempfile
from urllib.parse import urlsplit
from zipfile import ZipFile, ZIP_DEFLATED

from playwright.sync_api import sync_playwright
from openpyxl import Workbook

from browser_test import REPO, serve_program


TOKEN = "TEST_GRAPH_BEARER_NOT_PERSISTED"
TICKET = "TEST_DOWNLOAD_PROOF_NOT_PERSISTED"
SHARE = "https://example.sharepoint.com/:x:/r/sites/Tools/shared-tools.xlsx?web=1"
NAME = "shared-tools.xlsx"


class Cloud:
    def __init__(self, original):
        self.original = original
        self.data = original
        self.revision = 1
        self.calls = []
        self.puts = []
        self.mode = "normal"
        self.metadata_reads = 0
        self.race_at_get = None
        self.item_overrides = {}

    @property
    def tag(self):
        return f'"version-{self.revision}"'

    def item(self):
        result = {"id": "item-123", "name": NAME, "eTag": self.tag, "size": len(self.data),
                  "file": {"mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
                  "parentReference": {"driveId": "drive-123"},
                  "@microsoft.graph.downloadUrl": "https://download.example.test/document?ticket=" + TICKET}
        result.update(self.item_overrides)
        return result

    def route(self, route):
        request = route.request
        headers = {"access-control-allow-origin": "*", "access-control-allow-methods": "GET,PUT,OPTIONS",
                   "access-control-allow-headers": "Authorization,Content-Type,If-Match", "cache-control": "no-store"}
        if request.method == "OPTIONS":
            route.fulfill(status=204, headers=headers)
            return
        parsed = urlsplit(request.url)
        auth = request.headers.get("authorization")
        self.calls.append({"host": parsed.hostname, "path": parsed.path, "method": request.method,
                           "authorization": auth, "if_match": request.headers.get("if-match")})
        if parsed.hostname == "download.example.test":
            assert auth is None, "Graph bearer must never leave graph.microsoft.com"
            data = self.data
            if self.mode == "bad-verification" and self.puts:
                altered = bytearray(data)
                altered[-1] ^= 1
                data = bytes(altered)
            route.fulfill(status=200, body=data, headers={**headers, "content-type": "application/octet-stream"})
            return
        assert parsed.hostname == "graph.microsoft.com" and parsed.path.startswith("/v1.0/")
        assert auth == "Bearer " + TOKEN
        if self.mode in {"read401", "read403"}:
            route.fulfill(status=int(self.mode[-3:]), body="{}", headers=headers)
            return
        if request.method == "GET":
            self.metadata_reads += 1
            if self.race_at_get == self.metadata_reads:
                self.revision += 1
            route.fulfill(status=200, body=json.dumps(self.item()), headers={**headers, "content-type": "application/json"})
            return
        assert request.method == "PUT" and parsed.path.endswith("/content")
        incoming = request.post_data_buffer
        self.puts.append({"if_match": request.headers.get("if-match"), "sha256": sha256(incoming).hexdigest()})
        if self.mode.startswith("put"):
            route.fulfill(status=int(self.mode[3:]), body="{}", headers=headers)
            return
        if request.headers.get("if-match") != self.tag:
            route.fulfill(status=412, body="{}", headers=headers)
            return
        self.data = incoming
        self.revision += 1
        if self.mode == "commit-then-network-error":
            route.abort("failed")
            return
        result = self.item()
        if self.mode == "missing-write-etag":
            del result["eTag"]
        route.fulfill(status=200, body=json.dumps(result), headers={**headers, "content-type": "application/json"})


def arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    args = parser.parse_args()
    if args.workbook is None:
        args.workbook = next(Path("/workspace/attachments").rglob("*.xlsx"), None)
    if args.workbook is None or not args.workbook.is_file():
        parser.error("Supply --workbook with a read-only source XLSX")
    return args


SETUP = """()=>{
 window.testAuthEnabled=true;
 window.ToolM365Auth={getAccessToken:async()=>{
  if(!testAuthEnabled){const e=new Error('TEST_AUTH_REQUIRED');e.code='auth-required';throw e;}
  return 'TEST_GRAPH_BEARER_NOT_PERSISTED';
 }};
 window.testHash=async b=>Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',b)),n=>n.toString(16).padStart(2,'0')).join('');
 window.tryGraph=async fn=>{try{return {ok:true,value:await fn()};}catch(e){return {ok:false,code:e.code,uncertain:ToolGraphFile.inspect().uncertain};}};
}"""


def load_module(page, url, original, updated, source=True):
    page.goto(url)
    page.evaluate(SETUP)
    page.add_script_tag(path=REPO / "web/graph-file.js")
    page.evaluate("()=>{for(const id of ['original','updated']){let e=document.createElement('input');e.type='file';e.id=id;document.body.append(e);}}")
    page.locator("#original").set_input_files({"name": NAME, "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "buffer": original})
    page.locator("#updated").set_input_files({"name": NAME, "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "buffer": updated})
    page.evaluate("async()=>{window.originalBytes=new Uint8Array(await document.querySelector('#original').files[0].arrayBuffer());window.updatedBytes=new Uint8Array(await document.querySelector('#updated').files[0].arrayBuffer());}")


def connect(page):
    return page.evaluate("async shareUrl=>{const r=await ToolGraphFile.connect({shareUrl,validate:async f=>{if(!(await f.arrayBuffer()).byteLength)throw Error('empty')}});return {connected:r.connected,persistent:r.persistent,name:r.name,version:r.version}}", SHARE)


def new_context(browser, cloud):
    context = browser.new_context()
    context.route("https://graph.microsoft.com/**", cloud.route)
    context.route("https://download.example.test/**", cloud.route)
    return context


def main():
    args = arguments()
    original = args.workbook.read_bytes()
    source_hash = sha256(original).hexdigest()
    memory = BytesIO(original)
    with ZipFile(memory, "a", compression=ZIP_DEFLATED) as archive:
        archive.writestr("graph-protocol-fixture.txt", "Protocol-only workbook change; source cells and photos are preserved.")
    updated = memory.getvalue()
    with ZipFile(BytesIO(updated)) as archive:
        assert archive.testzip() is None
    # One full-size supplied workbook exercises binary fidelity. Negative
    # protocol branches use a small valid OOXML fixture to avoid repeatedly
    # transferring the same private photo-heavy workbook through test IPC.
    small_memory = BytesIO()
    small_book = Workbook()
    small_book.active.append(["Serial", "Cal Last", "Cal Next"])
    small_book.active.append(["0012345", "2027-03-17", "2027-09-17"])
    small_book.save(small_memory)
    small_original = small_memory.getvalue()
    small_update_memory = BytesIO(small_original)
    with ZipFile(small_update_memory, "a", compression=ZIP_DEFLATED) as archive:
        archive.writestr("graph-protocol-fixture.txt", "Conditional write protocol fixture")
    small_updated = small_update_memory.getvalue()
    outcomes = {}
    errors = []
    with tempfile.TemporaryDirectory(prefix="tool-graph-protocol-") as directory:
        shell = Path(directory) / "index.html"
        shell.write_text("<!doctype html><meta charset='utf-8'><title>Graph workbook protocol test</title>")
        with serve_program(shell) as url, sync_playwright() as playwright:
            browser = playwright.chromium.launch(executable_path=args.chromium, args=["--no-sandbox"])

            # Real browser fetches and IDB transactions; only the HTTP service
            # and the public client's authentication result are controlled.
            cloud = Cloud(original)
            context = new_context(browser, cloud)
            page = context.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))
            load_module(page, url, original, updated)
            print("Graph protocol: initial connection", flush=True)
            initial = connect(page)
            assert initial["connected"] and initial["persistent"]
            print("Graph protocol: verified save", flush=True)
            initial_downloads = sum(call["host"] == "download.example.test" for call in cloud.calls)
            cache_read = page.evaluate("""async()=>{
              const first=await ToolGraphFile.read();first.bytes[10]^=1;
              const second=await ToolGraphFile.read();return {hash:await testHash(second.bytes),version:second.version};
            }""")
            assert cache_read["hash"] == source_hash
            assert sum(call["host"] == "download.example.test" for call in cloud.calls) == initial_downloads
            outcomes["unchanged_etag_reuses_immutable_memory_snapshot"] = True
            saved = page.evaluate("""async()=>{
              const before=await ToolGraphFile.read();
              const saved=await ToolGraphFile.save({bytes:updatedBytes,expectedVersion:before.version});
              const current=await ToolGraphFile.read(),backup=await ToolGraphFile.takeBackup();
              return {saved,hash:await testHash(current.bytes),backupHash:await testHash(backup.bytes),inspect:ToolGraphFile.inspect()};
            }""")
            assert saved["hash"] == sha256(updated).hexdigest() and saved["backupHash"] == source_hash
            assert not saved["inspect"]["uncertain"] and len(cloud.puts) == 1
            assert cloud.puts[0]["if_match"] == '"version-1"'
            assert cloud.data == updated
            assert sum(call["host"] == "download.example.test" for call in cloud.calls) == initial_downloads + 1, "PUT must force a fresh verification download"
            outcomes["verified_save_and_exact_backup"] = True

            print("Graph protocol: persistence privacy", flush=True)
            privacy = page.evaluate("""async()=>{
              const db=await new Promise((resolve,reject)=>{let r=indexedDB.open('tool-management-graph-file-v1');r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error)});
              const get=store=>new Promise((resolve,reject)=>{let t=db.transaction(store),r=t.objectStore(store).getAll();r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error)});
              const targets=await get('targets'),backups=(await get('backups')).map(b=>({...b,bytes:'binary-omitted'}));
              return {targets,backups,local:JSON.stringify(localStorage),session:JSON.stringify(sessionStorage),caches:await caches.keys()};
            }""")
            raw = json.dumps(privacy)
            assert TOKEN not in raw and TICKET not in raw and "download.example.test" not in raw and SHARE not in raw
            assert set(privacy["targets"][0]) == {"driveId", "itemId", "name"}
            assert not privacy["caches"]
            outcomes["no_token_or_download_url_persistence"] = True

            # App restart remembers the chosen document but needs a fresh
            # company login; it never persists the bearer to reconnect itself.
            prior_requests = len(cloud.calls)
            page.reload()
            page.evaluate(SETUP)
            page.evaluate("testAuthEnabled=false")
            page.add_script_tag(path=REPO / "web/graph-file.js")
            restored = page.evaluate("async()=>{await ToolGraphFile.restore();return ToolGraphFile.inspect()}")
            assert restored["requiresSignIn"] and not restored["connected"] and len(cloud.calls) == prior_requests
            page.evaluate("testAuthEnabled=true")
            restored = page.evaluate("async()=>{await ToolGraphFile.restore();return ToolGraphFile.inspect()}")
            assert restored["connected"]
            outcomes["restore_requires_real_authentication"] = True
            context.close()

            # A staff launch link identifies one exact drive/item without
            # requiring a private SharePoint sharing URL or redeeming a link.
            cloud = Cloud(small_original)
            context = new_context(browser, cloud)
            page = context.new_page()
            load_module(page, url, small_original, small_updated)
            direct = page.evaluate("""async()=>{
              const connected=await ToolGraphFile.connect({target:{driveId:'drive-123',itemId:'item-123'},validate:async file=>{
                if(file.name!=='shared-tools.xlsx'||await testHash(await file.arrayBuffer())!==await testHash(originalBytes))throw Error('wrong file');
              }});
              const before=await ToolGraphFile.read();
              await ToolGraphFile.save({expectedVersion:before.version,bytes:updatedBytes});
              return {connected:connected.connected,target:ToolGraphFile.inspect().target,hash:await testHash((await ToolGraphFile.read()).bytes)};
            }""")
            assert direct["connected"] and direct["target"] == {"driveId": "drive-123", "itemId": "item-123", "name": NAME}
            assert direct["hash"] == sha256(small_updated).hexdigest() and len(cloud.puts) == 1
            assert not any("/shares/" in call["path"] for call in cloud.calls)
            outcomes["exact_target_connects_and_saves_without_share_link"] = True
            page.reload()
            page.evaluate(SETUP)
            page.add_script_tag(path=REPO / "web/graph-file.js")
            assert page.evaluate("async()=>{const s=await ToolGraphFile.restore();return s.connected&&s.target.itemId==='item-123'}")
            outcomes["verified_exact_target_is_persisted"] = True
            context.close()

            negative_targets = [
                ("target_and_share_link", {"shareUrl": SHARE, "target": {"driveId": "drive-123", "itemId": "item-123"}}, {}, "normal", None),
                ("invalid_target_id", {"target": {"driveId": "drive-123", "itemId": "item-123\n"}}, {}, "normal", None),
                ("invalid_target_name", {"target": {"driveId": "drive-123", "itemId": "item-123", "name": "program.html"}}, {}, "normal", None),
                ("target_permission_denied", {"target": {"driveId": "drive-123", "itemId": "item-123"}}, {}, "read403", None),
                ("wrong_item_metadata", {"target": {"driveId": "drive-123", "itemId": "item-123"}}, {"id": "other-item"}, "normal", None),
                ("wrong_drive_metadata", {"target": {"driveId": "drive-123", "itemId": "item-123"}}, {"parentReference": {"driveId": "other-drive"}}, "normal", None),
                ("target_is_folder", {"target": {"driveId": "drive-123", "itemId": "item-123"}}, {"folder": {}}, "normal", None),
                ("target_not_xlsx", {"target": {"driveId": "drive-123", "itemId": "item-123"}}, {"name": "wrong.pdf"}, "normal", None),
                ("target_too_large", {"target": {"driveId": "drive-123", "itemId": "item-123"}}, {"size": 100 * 1024 * 1024 + 1}, "normal", None),
                ("target_corrupt_bytes", {"target": {"driveId": "drive-123", "itemId": "item-123"}}, {}, "normal", "bytes"),
                ("target_validation_rejects", {"target": {"driveId": "drive-123", "itemId": "item-123"}}, {}, "normal", "validate"),
            ]
            for label, options, overrides, mode, reject in negative_targets:
                print("Graph protocol:", label, flush=True)
                cloud = Cloud(small_original)
                cloud.item_overrides = overrides
                cloud.mode = mode
                if reject == "bytes":
                    cloud.data = b"NO" + small_original[2:]
                context = new_context(browser, cloud)
                page = context.new_page()
                load_module(page, url, small_original, small_updated)
                result = page.evaluate("""async({options,reject})=>{
                  if(reject==='validate')options.validate=async()=>{throw Error('TEST_WORKBOOK_VALIDATION_REJECTED');};
                  const result=await tryGraph(()=>ToolGraphFile.connect(options));
                  const after=await ToolGraphFile.restore();
                  return {result,connected:after.connected,target:after.target};
                }""", {"options": options, "reject": reject})
                assert not result["result"]["ok"] and not result["connected"] and result["target"] is None, (label, result)
                assert not cloud.puts and not any("/shares/" in call["path"] for call in cloud.calls)
                if label in {"target_and_share_link", "invalid_target_id", "invalid_target_name"}:
                    assert not cloud.calls
                outcomes[label + "_not_persisted"] = result["result"]["code"]
                context.close()

            def scenario(label, mode="normal"):
                print("Graph protocol:", label, flush=True)
                cloud = Cloud(small_original)
                context = new_context(browser, cloud)
                page = context.new_page()
                page.on("pageerror", lambda error: errors.append(str(error)))
                load_module(page, url, small_original, small_updated)
                connect(page)
                page.evaluate("async()=>{window.before=await ToolGraphFile.read();return true}")
                cloud.mode = mode
                return cloud, context, page

            cloud, context, page = scenario("changed_document")
            cloud.revision += 1
            result = page.evaluate("async()=>await tryGraph(()=>ToolGraphFile.save({bytes:updatedBytes,expectedVersion:before.version}))")
            assert result["code"] == "conflict" and not cloud.puts and cloud.data == cloud.original
            outcomes["changed_before_save"] = result["code"]
            before_downloads = sum(call["host"] == "download.example.test" for call in cloud.calls)
            cloud.revision += 1
            page.evaluate("async()=>{await ToolGraphFile.read();return true}")
            assert sum(call["host"] == "download.example.test" for call in cloud.calls) == before_downloads + 1
            outcomes["changed_etag_forces_fresh_download"] = True
            context.close()

            cloud, context, page = scenario("changed_during_backup")
            cloud.race_at_get = cloud.metadata_reads + 2
            result = page.evaluate("async()=>await tryGraph(()=>ToolGraphFile.save({bytes:updatedBytes,expectedVersion:before.version}))")
            assert result["code"] == "conflict" and not cloud.puts and cloud.data == cloud.original
            outcomes["changed_during_backup"] = result["code"]
            context.close()

            cloud, context, page = scenario("backup_failure")
            result = page.evaluate("""async()=>{
              const put=IDBObjectStore.prototype.put;IDBObjectStore.prototype.put=function(...a){if(this.name==='backups')throw new DOMException('TEST_QUOTA','QuotaExceededError');return put.apply(this,a)};
              try{return await tryGraph(()=>ToolGraphFile.save({bytes:updatedBytes,expectedVersion:before.version}));}finally{IDBObjectStore.prototype.put=put;}
            }""")
            assert result["code"] == "backup-failed" and not cloud.puts and cloud.data == cloud.original
            outcomes["backup_failure_blocks_upload"] = True
            context.close()

            cloud, context, page = scenario("stale_ui")
            result = page.evaluate("""async()=>{let checks=0;const r=await tryGraph(()=>ToolGraphFile.save({bytes:updatedBytes,expectedVersion:before.version,verifyCurrent:()=>++checks<4}));return {...r,checks};}""")
            assert result["code"] == "stale" and result["checks"] == 4 and not cloud.puts
            outcomes["changed_photo_selection_blocks_upload"] = True
            context.close()

            for http, expected in [(401,"auth-required"),(403,"permission-denied"),(409,"file-locked"),(412,"conflict"),(423,"file-locked"),(429,"throttled")]:
                cloud, context, page = scenario(f"put_{http}", "put" + str(http))
                result = page.evaluate("async()=>await tryGraph(()=>ToolGraphFile.save({bytes:updatedBytes,expectedVersion:before.version}))")
                assert result["code"] == expected and not result["uncertain"] and cloud.data == cloud.original
                outcomes["conditional_put_" + str(http)] = expected
                context.close()

            for mode in ["commit-then-network-error", "bad-verification", "missing-write-etag", "put503"]:
                cloud, context, page = scenario(mode, mode)
                result = page.evaluate("async()=>await tryGraph(()=>ToolGraphFile.save({bytes:updatedBytes,expectedVersion:before.version}))")
                assert not result["ok"] and result["uncertain"] and len(cloud.puts) == 1, (mode,result)
                again = page.evaluate("async()=>await tryGraph(()=>ToolGraphFile.save({bytes:updatedBytes,expectedVersion:before.version}))")
                assert again["code"] == "requires-read" and len(cloud.puts) == 1
                cloud.mode = "normal"
                downloads = sum(call["host"] == "download.example.test" for call in cloud.calls)
                page.evaluate("async()=>{await ToolGraphFile.read();return true}")
                assert sum(call["host"] == "download.example.test" for call in cloud.calls) == downloads + 1, "Unknown write must discard cache and force a fresh download"
                assert not page.evaluate("ToolGraphFile.inspect().uncertain")
                assert page.evaluate("async()=>await testHash((await ToolGraphFile.takeBackup()).bytes)") == sha256(cloud.original).hexdigest()
                outcomes[mode] = "held_until_fresh_read"
                context.close()

            cloud, context, page = scenario("concurrent_clients")
            second = context.new_page()
            load_module(second, url, small_original, small_updated)
            connect(second)
            second.evaluate("async()=>{window.before=await ToolGraphFile.read();return true}")
            page.evaluate("()=>window.pending=tryGraph(()=>ToolGraphFile.save({bytes:updatedBytes,expectedVersion:before.version}))")
            second.evaluate("()=>window.pending=tryGraph(()=>ToolGraphFile.save({bytes:updatedBytes,expectedVersion:before.version}))")
            first_result = page.evaluate("pending")
            second_result = second.evaluate("pending")
            assert sorted([first_result["ok"],second_result["ok"]]) == [False,True]
            assert len(cloud.puts) == 1 and cloud.data == small_updated
            outcomes["concurrent_tabs_one_writer"] = True
            context.close()
            browser.close()
    assert not errors, errors
    assert sha256(args.workbook.read_bytes()).hexdigest() == source_hash
    print(json.dumps({"checks":outcomes,"original_xlsx_unchanged":True,"page_errors":errors,
                      "scope":"Real browser fetch/IndexedDB/conditional write protocol with controlled service; live company authentication is not tested."},ensure_ascii=False,indent=2))


if __name__ == "__main__":
    main()
