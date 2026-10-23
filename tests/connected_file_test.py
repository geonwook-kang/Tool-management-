"""Real Chromium OPFS/IndexedDB tests for the connected workbook workflow.

The picker returns a real origin-private file handle. Read/write/abort, hashing,
backup persistence, and handle restoration are real browser operations. OCR is
a deterministic parser-based stub; actual OneDrive/SharePoint sync is untested.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
import json
from pathlib import Path
import tempfile
from urllib.parse import urlsplit

from playwright.sync_api import sync_playwright

from browser_test import REPO, download_expression, fixture_photo, serve_program, verify_workbook


NAME = "company-shared.xlsx"
PICKER_SCRIPT = """()=>{
  window.testPickerCalls=Number(sessionStorage.getItem('testPickerCalls')||0);
  window.showOpenFilePicker=async()=>{
    window.testPickerCalls++;sessionStorage.setItem('testPickerCalls',String(testPickerCalls));
    return [await(await navigator.storage.getDirectory()).getFileHandle('company-shared.xlsx')];
  };
}"""


def arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--program", type=Path, default=REPO / "web/dist/Tool-Manager.html")
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    parser.add_argument("--module-only", action="store_true")
    args = parser.parse_args()
    if args.workbook is None:
        args.workbook = next(Path("/workspace/attachments").rglob("*.xlsx"), None)
    if args.workbook is None:
        parser.error("Pass --workbook with the attached company workbook")
    args.photo = None
    return args


def seed_file(page, workbook):
    page.evaluate("""()=>{if(!document.getElementById('test-fixture')){
      const input=document.createElement('input');input.type='file';input.id='test-fixture';input.hidden=true;document.body.append(input);
    }}""")
    page.locator("#test-fixture").set_input_files(workbook)
    return page.evaluate("""async()=>{
      const h=await(await navigator.storage.getDirectory()).getFileHandle('company-shared.xlsx',{create:true});
      const w=await h.createWritable();await w.write(await document.querySelector('#test-fixture').files[0].arrayBuffer());await w.close();
      window.testHandle=h;
      window.testHash=async bytes=>Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))).map(n=>n.toString(16).padStart(2,'0')).join('');
      return testHash(await(await h.getFile()).arrayBuffer());
    }""")


def get_handle_script():
    return "await(await navigator.storage.getDirectory()).getFileHandle('company-shared.xlsx')"


def export_shared(page, output):
    download_expression(page, f"new Uint8Array(await(await({get_handle_script()})).getFile().then(f=>f.arrayBuffer()))", output)


def module_page(page, url):
    origin = f"{urlsplit(url).scheme}://{urlsplit(url).netloc}"
    page.route("**/module-fixture.html", lambda route: route.fulfill(status=200, content_type="text/html", body="<!doctype html><title>Connected-file module test</title>"))
    page.goto(origin + "/module-fixture.html")
    page.add_script_tag(path=REPO / "web/connected-file.js")


def module_checks(context, page, url, args):
    context.add_init_script("(" + PICKER_SCRIPT + ")()")
    module_page(page, url)
    initial_hash = seed_file(page, args.workbook)
    connection = page.evaluate("async()=>await ToolConnectedFile.connect()")
    assert connection["connected"] and connection["name"] == NAME
    assert page.evaluate("testPickerCalls") == 1
    result = page.evaluate("""async()=>{
      const api=ToolConnectedFile;
      const first=await api.read(),hash=async()=>testHash(await(await testHandle.getFile()).arrayBuffer());
      const outcome={};
      try{await api.save({bytes:new Uint8Array([1,2,3]),expectedVersion:first.version});outcome.invalid='accepted';}
      catch(e){outcome.invalid=e.code;}
      outcome.invalidUnchanged=(await hash())===await testHash(first.bytes);
      await api.save({bytes:first.bytes,expectedVersion:first.version});
      outcome.noChangeUnchanged=(await hash())===await testHash(first.bytes);
      const fresh=await api.read();
      let checks=0;
      try{await api.save({bytes:new Uint8Array([80,75,3,4,55]),expectedVersion:fresh.version,verifyCurrent:()=>++checks<4});outcome.stale='accepted';}
      catch(e){outcome.stale=e.code;}
      outcome.staleChecks=checks;outcome.staleUnchanged=(await hash())===await testHash(first.bytes);
      const originalPut=IDBObjectStore.prototype.put;
      IDBObjectStore.prototype.put=function(...args){if(this.name==='backups')throw new DOMException('TEST_BACKUP_FAILURE','UnknownError');return originalPut.apply(this,args);};
      try{await api.save({bytes:new Uint8Array([80,75,3,4,66]),expectedVersion:(await api.read()).version});outcome.backupFailure='accepted';}
      catch(e){outcome.backupFailure=e.code;}
      finally{IDBObjectStore.prototype.put=originalPut;}
      outcome.backupFailureUnchanged=(await hash())===await testHash(first.bytes);
      const before=await api.read(),next=new Uint8Array([80,75,3,4,77]);
      await api.save({bytes:next,expectedVersion:before.version});
      const backup=await api.takeBackup();
      outcome.backup=backup&&{hash:await testHash(backup.bytes),name:backup.name};
      outcome.savedBytes=Array.from(new Uint8Array(await(await testHandle.getFile()).arrayBuffer()));
      return outcome;
    }""")
    assert result["invalid"] == "invalid-file" and result["invalidUnchanged"]
    assert result["noChangeUnchanged"]
    assert result["stale"] == "stale" and result["staleChecks"] == 4 and result["staleUnchanged"]
    assert result["backupFailure"] == "backup-failed" and result["backupFailureUnchanged"]
    assert result["backup"]["hash"] == initial_hash
    assert result["savedBytes"] == [80, 75, 3, 4, 77]

    # A queued writer must hash again after acquiring the real browser lock.
    page.evaluate("""async()=>{
      window.beforeConflict=await ToolConnectedFile.read();
      window.heldLock=navigator.locks.request('tool-management-connected-workbook',()=>new Promise(resolve=>window.releaseTestLock=resolve));
    }""")
    page.wait_for_function("!!window.releaseTestLock")
    page.evaluate("""()=>{window.pendingSave=ToolConnectedFile.save({bytes:new Uint8Array([80,75,3,4,88]),expectedVersion:beforeConflict.version}).then(v=>({ok:true}),e=>({ok:false,code:e.code}));}""")
    conflict = page.evaluate("""async()=>{
      const w=await testHandle.createWritable();await w.write(new Uint8Array([80,75,3,4,99]));await w.close();releaseTestLock();
      const result=await pendingSave;result.bytes=Array.from(new Uint8Array(await(await testHandle.getFile()).arrayBuffer()));return result;
    }""")
    assert conflict == {"ok": False, "code": "conflict", "bytes": [80, 75, 3, 4, 99]}, conflict

    # Reload reconstructs both the handle and original pre-write backup via IDB.
    page.reload()
    page.add_script_tag(path=REPO / "web/connected-file.js")
    restored = page.evaluate("""async()=>{
      const state=await ToolConnectedFile.restore(),backup=await ToolConnectedFile.takeBackup();
      return {connected:state.connected,name:state.name,pickers:testPickerCalls,backupBytes:backup&&backup.bytes.length};
    }""")
    assert restored["connected"] and restored["name"] == NAME and restored["pickers"] == 1
    assert restored["backupBytes"] > 9_000_000
    permission = page.evaluate("""async()=>{
      const q=FileSystemHandle.prototype.queryPermission,r=FileSystemHandle.prototype.requestPermission;
      FileSystemHandle.prototype.queryPermission=async()=> 'prompt';FileSystemHandle.prototype.requestPermission=async()=> 'denied';
      try{const restore=await ToolConnectedFile.restore(),access=await ToolConnectedFile.requestAccess();return {restore,access};}
      finally{FileSystemHandle.prototype.queryPermission=q;FileSystemHandle.prototype.requestPermission=r;}
    }""")
    assert not permission["restore"]["connected"] and permission["restore"]["needsPermission"]
    assert not permission["access"]["connected"] and permission["access"]["needsPermission"]
    verification = page.evaluate("""async()=>{
      await ToolConnectedFile.requestAccess();
      const before=await ToolConnectedFile.read(),native=FileSystemFileHandle.prototype.createWritable;
      FileSystemFileHandle.prototype.createWritable=async function(...args){
        const handle=this,stream=await native.apply(handle,args),close=stream.close.bind(stream);
        stream.close=async()=>{await close();const concurrent=await native.call(handle);await concurrent.write(new Uint8Array([80,75,3,4,111]));await concurrent.close();};
        return stream;
      };
      let code;
      try{await ToolConnectedFile.save({bytes:new Uint8Array([80,75,3,4,112]),expectedVersion:before.version});}
      catch(e){code=e.code;}
      finally{FileSystemFileHandle.prototype.createWritable=native;}
      const uncertain=ToolConnectedFile.inspect().uncertain;
      let retryCode;try{await ToolConnectedFile.save({bytes:before.bytes,expectedVersion:before.version});}catch(e){retryCode=e.code;}
      const backup=await ToolConnectedFile.takeBackup(),h=await(await navigator.storage.getDirectory()).getFileHandle('company-shared.xlsx');
      const bytes=Array.from(new Uint8Array(await(await h.getFile()).arrayBuffer()));
      await ToolConnectedFile.read();
      return {code,uncertain,retryCode,bytes,backup:Array.from(backup.bytes),cleared:!ToolConnectedFile.inspect().uncertain};
    }""")
    assert verification == {"code": "verification-failed", "uncertain": True, "retryCode": "requires-read", "bytes": [80, 75, 3, 4, 111], "backup": [80, 75, 3, 4, 99], "cleared": True}, verification
    return {"real_opfs_write": True, "invalid_stale_backup_failure_no_write": True, "lock_conflict_preserved_external_bytes": True, "idb_handle_backup_restored": True, "permission_denial_reported": True, "readback_failure_requires_reread_preserves_backup": True}


def install_ocr(page, last, due):
    page.evaluate("""({last,due})=>{
      const parse=ToolOCR.parse,text='강건욱\\n31550075\\n교정일자 '+last.replaceAll('-','.')+'\\n차기교정예정일자 '+due.replaceAll('-','.');
      window.ToolOCR={...ToolOCR,recognize:async(file,tools)=>parse(text,tools,{confidence:95,passes:[{text,confidence:95},{text,confidence:95}]})};
    }""", {"last": last, "due": due})


def ready(page):
    page.wait_for_function("window.ToolApp&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=30000)


def external_edit(page, value):
    page.evaluate("""async value=>{
      const h=await(await navigator.storage.getDirectory()).getFileHandle('company-shared.xlsx');
      const zip=await JSZip.loadAsync(await(await h.getFile()).arrayBuffer());
      const doc=new DOMParser().parseFromString(await zip.file('xl/worksheets/sheet1.xml').async('string'),'application/xml');
      const cell=Array.from(doc.getElementsByTagNameNS('*','c')).find(c=>c.getAttribute('r')==='C2');
      while(cell.firstChild)cell.removeChild(cell.firstChild);cell.setAttribute('t','inlineStr');
      const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main',is=doc.createElementNS(ns,'is'),t=doc.createElementNS(ns,'t');
      t.textContent=value;is.append(t);cell.append(is);zip.file('xl/worksheets/sheet1.xml',new XMLSerializer().serializeToString(doc));
      const bytes=await zip.generateAsync({type:'uint8array',compression:'DEFLATE'}),w=await h.createWritable();await w.write(bytes);await w.close();
    }""", value)


def ui_checks(context, page, url, args, temporary):
    context.add_init_script("(" + PICKER_SCRIPT + ")()")
    page.goto(url)
    page.wait_for_function("window.ToolApp")
    generic_link = "https://example.sharepoint.com/sites/Calibration/Shared%20Documents/Tools.xlsx"
    page.locator("#sharepoint-url").evaluate("el=>{el.closest('details').open=true;}")
    page.locator("#sharepoint-url").fill(generic_link)
    page.locator("#remember-sharepoint").click()
    assert page.locator("#open-sharepoint").get_attribute("href") == generic_link
    assert not page.evaluate("link=>ToolDistribution.html().includes(link)", generic_link)
    page.locator("#sharepoint-url").evaluate("el=>{el.closest('details').open=false;}")
    seed_file(page, args.workbook)
    page.locator("#connect-onedrive").click()
    page.wait_for_function("ToolApp.state.source.connected&&!!ToolApp.state.model")
    ready(page)
    assert page.evaluate("testPickerCalls") == 1
    photo = fixture_photo(args, temporary)
    install_ocr(page, "2026-12-29", "2027-12-29")
    page.locator("#photos").set_input_files(photo)
    page.wait_for_function("ToolApp.state.photos.length>0&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=30000)
    page.wait_for_function("ToolApp.state.photos.some(p=>!!p.saved)&&!ToolApp.state.source.saving")
    assert "이 PC" in page.locator("#status").inner_text() and "OneDrive" in page.locator("#status").inner_text()
    first = temporary / "first-shared-save.xlsx"
    export_shared(page, first)
    first_result = verify_workbook(args.workbook, first)
    assert page.evaluate("testPickerCalls") == 1
    with page.expect_download() as event:
        page.locator("#download-backup").click()
    backup = temporary / "pre-write-backup.xlsx"
    event.value.save_as(backup)
    assert backup.read_bytes() == args.workbook.read_bytes()
    # Another person's edit between batches is adopted before the next update.
    external_edit(page, "CONCURRENT MANUFACTURER")
    concurrent = temporary / "concurrent-source.xlsx"
    export_shared(page, concurrent)
    photo2 = temporary / "next-year-photo.png"
    photo2.write_bytes(photo.read_bytes())
    install_ocr(page, "2027-12-29", "2028-12-29")
    page.locator("#photos").set_input_files(photo2)
    page.wait_for_function("ToolApp.state.photos.filter(p=>!!p.saved).length===2&&!ToolApp.state.source.saving", timeout=30000)
    ready(page)
    second = temporary / "second-shared-save.xlsx"
    export_shared(page, second)
    second_result = verify_workbook(concurrent, second, last="2027-12-29", due="2028-12-29")
    assert not page.evaluate("ToolApp.state.source.error")
    stages = page.evaluate("ToolApp.state.photos.map(p=>({last:p.saved.last,serial:p.saved.serial}))")
    assert stages == [{"last": "2026-12-29", "serial": "31550075"}, {"last": "2027-12-29", "serial": "31550075"}]
    assert page.locator("#photo-list .photo-card").count() == 2
    with page.expect_download() as event:
        page.locator("#download-backup").click()
    second_backup = temporary / "second-pre-write-backup.xlsx"
    event.value.save_as(second_backup)
    assert second_backup.read_bytes() == concurrent.read_bytes()
    second_id = page.evaluate("ToolApp.state.photos[1].id")
    page.locator(f'.photo-card[data-id="{second_id}"]').click()
    page.screenshot(path="/tmp/tool-manager-connected.png", full_page=True)

    page.reload()
    page.wait_for_function("window.ToolApp&&ToolApp.state.source.connected&&!!ToolApp.state.model")
    ready(page)
    assert page.evaluate("testPickerCalls") == 1
    assert page.locator("#sharepoint-url").input_value() == generic_link
    assert not page.evaluate("link=>ToolDistribution.html().includes(link)", generic_link)
    assert page.evaluate("ToolApp.state.model.tools.find(t=>t.sheet==='광주'&&t.row===6).last") == "2027-12-29"
    after_restore = temporary / "restored-source.xlsx"
    export_shared(page, after_restore)
    assert after_restore.read_bytes() == second.read_bytes()
    with page.expect_download() as event:
        page.locator("#download-backup").click()
    restored_backup = temporary / "restored-backup.xlsx"
    event.value.save_as(restored_backup)
    assert restored_backup.read_bytes() == concurrent.read_bytes()

    context.add_init_script("FileSystemHandle.prototype.queryPermission=async()=> 'prompt';FileSystemHandle.prototype.requestPermission=async()=> 'denied';")
    page.reload()
    page.wait_for_function("window.ToolApp&&ToolConnectedFile.inspect().needsPermission")
    assert not page.evaluate("ToolApp.state.source.connected")
    assert page.locator("#reconnect-onedrive").is_visible()
    page.locator("#reconnect-onedrive").click()
    ready(page)
    assert not page.evaluate("ToolApp.state.source.connected")
    assert "권한" in page.locator("#connection-status").inner_text()
    assert page.locator("#photos").is_disabled()
    denied_source = temporary / "permission-denied-source.xlsx"
    export_shared(page, denied_source)
    assert denied_source.read_bytes() == second.read_bytes()
    return {"photos_only_auto_save": first_result, "next_batch_preserves_external_edit": second_result, "saved_gallery_retained_not_reapplied": True, "picker_calls": 1, "exact_durable_backups": True, "restore_no_write_no_picker": True, "permission_denial_not_saved": True, "browser_only_link_persists_not_exported": True}


def conflict_ui(context, page, url, args, temporary):
    context.add_init_script("(" + PICKER_SCRIPT + ")()")
    page.goto(url)
    page.wait_for_function("window.ToolApp")
    seed_file(page, args.workbook)
    page.locator("#connect-onedrive").click()
    page.wait_for_function("ToolApp.state.source.connected&&!!ToolApp.state.model")
    ready(page)
    install_ocr(page, "2026-12-29", "2027-12-29")
    page.evaluate("""()=>{
      const apply=ToolWorkbook.apply;
      window.testOriginalApply=apply;
      window.ToolWorkbook={...ToolWorkbook,apply:async(...args)=>{window.generationStarted=true;await new Promise(resolve=>window.releaseGeneration=resolve);return apply(...args);}};
    }""")
    photo = fixture_photo(args, temporary)
    page.locator("#photos").set_input_files(photo)
    page.wait_for_function("!!window.generationStarted")
    external_edit(page, "EDIT DURING GENERATION")
    external = temporary / "generation-external.xlsx"
    export_shared(page, external)
    page.evaluate("()=>{window.releaseGeneration();}")
    page.wait_for_function("!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving&&!!ToolApp.state.source.error")
    assert not page.evaluate("!!ToolApp.state.photos[0].saved")
    assert page.locator("#automatic-result-download").is_visible()
    after = temporary / "after-conflict.xlsx"
    export_shared(page, after)
    assert after.read_bytes() == external.read_bytes()
    with page.expect_download() as event:
        page.locator("#automatic-result-download").click()
    independent = temporary / "conflict-download.xlsx"
    event.value.save_as(independent)
    verify_workbook(args.workbook, independent)
    assert page.locator("#status").get_attribute("class") == "error"
    assert page.evaluate("!!ToolApp.state.source.error&&!ToolApp.state.source.savedAt")
    # Choosing a normal input workbook explicitly exits the connected source,
    # so subsequent results cannot silently write to the old company file.
    page.evaluate("()=>{window.ToolWorkbook={...ToolWorkbook,apply:window.testOriginalApply};}")
    page.locator("#workbook").set_input_files(args.workbook)
    page.wait_for_function("!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.connected&&!!ToolApp.state.preparedResult")
    assert not page.evaluate("ToolConnectedFile.inspect().connected")
    after_manual = temporary / "source-after-manual-mode.xlsx"
    export_shared(page, after_manual)
    assert after_manual.read_bytes() == external.read_bytes()
    return {"external_edit_not_overwritten": True, "saved_status_not_faked": True, "generated_result_still_downloadable": True, "manual_workbook_disconnects_shared_source": True}


def main():
    args = arguments()
    original_hash = sha256(args.workbook.read_bytes()).hexdigest()
    with tempfile.TemporaryDirectory(prefix="tool-connected-test-") as directory, serve_program(args.program) as url, sync_playwright() as p:
        browser = p.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
        remote_requests = []
        def watch_requests(context):
            context.on("request", lambda request: remote_requests.append(request.url) if urlsplit(request.url).scheme in {"http", "https"} and urlsplit(request.url).netloc != urlsplit(url).netloc else None)
        context = browser.new_context(accept_downloads=True, locale="ko-KR")
        watch_requests(context)
        result = {"module": module_checks(context, context.new_page(), url, args)}
        context.close()
        if not args.module_only:
            context = browser.new_context(accept_downloads=True, locale="ko-KR")
            watch_requests(context)
            page = context.new_page()
            page.on("dialog", lambda dialog: dialog.accept())
            result["ui"] = ui_checks(context, page, url, args, Path(directory))
            context.close()
            context = browser.new_context(accept_downloads=True, locale="ko-KR")
            watch_requests(context)
            page = context.new_page()
            page.on("dialog", lambda dialog: dialog.accept())
            result["conflict_ui"] = conflict_ui(context, page, url, args, Path(directory))
            context.close()
        browser.close()
    assert not remote_requests, remote_requests
    assert sha256(args.workbook.read_bytes()).hexdigest() == original_hash
    result["uploaded_original_unchanged"] = True
    result["onedrive_cloud_sync_tested"] = False
    result["remote_requests"] = remote_requests
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
