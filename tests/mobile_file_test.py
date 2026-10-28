"""Device-workbook persistence with real IndexedDB, hashing and XLSX changes.

No File System Access API or cloud service is used. An unchanged company XLSX
is supplied at runtime. Generated workbooks remain in /tmp. This tests device
storage, real photo/date writes, backup/atomic failure/concurrency semantics;
it does not measure OCR accuracy or OneDrive server synchronization.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
import json
from pathlib import Path
import tempfile
from urllib.parse import urlsplit

from PIL import Image
from playwright.sync_api import sync_playwright

from browser_test import REPO, serve_program, fixture_photo, verify_workbook, download_expression


def arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--program", type=Path, default=Path('/workspace/outputs/Tool-Manager-Fixed.html'))
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--chromium", default='/usr/bin/chromium')
    args = parser.parse_args()
    if args.workbook is None:
        args.workbook = next(Path('/workspace/attachments').rglob('*.xlsx'), None)
    if args.workbook is None or not args.workbook.is_file():
        parser.error('Supply --workbook with the attached company workbook')
    args.photo = None
    return args


def module_page(page, url):
    origin = f'{urlsplit(url).scheme}://{urlsplit(url).netloc}'
    page.route('**/device-fixture.html', lambda route: route.fulfill(status=200, content_type='text/html', body='<!doctype html><input id="book" type="file"><input id="picture" type="file">'))
    page.goto(origin + '/device-fixture.html')
    page.add_script_tag(path=REPO / 'web/vendor/jszip.min.js')
    page.add_script_tag(path=REPO / 'web/workbook.js')
    page.add_script_tag(path=REPO / 'web/mobile-file.js')
    page.evaluate("""()=>{window.testHash=async bytes=>Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))).map(n=>n.toString(16).padStart(2,'0')).join('');}""")


def main():
    args = arguments()
    original_hash = sha256(args.workbook.read_bytes()).hexdigest()
    results, errors, external = {}, [], []
    with tempfile.TemporaryDirectory(prefix='tool-device-storage-') as directory, serve_program(args.program) as url, sync_playwright() as playwright:
        temporary = Path(directory)
        photo = fixture_photo(args, temporary)
        with Image.open(photo) as image:
            dimensions = list(image.size)
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=['--no-sandbox'])
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        page.on('pageerror', lambda error: errors.append(str(error)))
        context.on('request', lambda request: external.append(request.url) if urlsplit(request.url).scheme in {'http', 'https'} and urlsplit(request.url).netloc != urlsplit(url).netloc else None)
        module_page(page, url)
        page.locator('#book').set_input_files(args.workbook)
        page.locator('#picture').set_input_files(photo)
        imported = page.evaluate("""async()=>{
          const api=ToolMobileFile;
          const empty=await api.restore();
          const file=document.querySelector('#book').files[0];
          const connected=await api.connect(file,{validate:async f=>ToolWorkbook.load(await f.arrayBuffer())});
          const snapshot=await api.read();window.initialVersion=snapshot.version;
          return {empty:empty.connected,mode:connected.mode,connected:connected.connected,persistent:connected.persistent,
            name:snapshot.file.name,hash:await testHash(snapshot.bytes),backup:await api.takeBackup(),supported:api.supported()};
        }""")
        assert imported['empty'] is False and imported['connected'] and imported['persistent'] and imported['supported']
        assert imported['mode'] == 'device-workbook' and imported['hash'] == original_hash and imported['backup'] is None
        results['import'] = imported
        invalid = page.evaluate("""async()=>{
          const before=await ToolMobileFile.read(),out={};
          for(const [name,bytes] of [['bad.xlsx',[1,2,3]],['bad.xlsm',[80,75,3,4]]]){
            try{await ToolMobileFile.connect(new File([new Uint8Array(bytes)],name));out[name]='accepted';}catch(e){out[name]=e.code;}
          }
          try{await ToolMobileFile.connect(document.querySelector('#book').files[0],{validate:async()=>{throw new Error('TEST_VALIDATION_REJECTION');}});out.validation='accepted';}
          catch(e){out.validation=e.message;}
          out.unchanged=(await ToolMobileFile.read()).version===before.version;return out;
        }""")
        assert invalid == {'bad.xlsx': 'invalid-file', 'bad.xlsm': 'invalid-file', 'validation': 'TEST_VALIDATION_REJECTION', 'unchanged': True}
        results['invalid_imports'] = invalid
        update = page.evaluate("""async dims=>{
          const before=await ToolMobileFile.read(),model=await ToolWorkbook.load(before.bytes),tool=model.tools.find(t=>t.serial==='31550075');
          const update={toolId:tool.id,name:tool.name,serial:tool.serial,last:'2026-12-29',due:'2027-12-29',
            photoBytes:new Uint8Array(await document.querySelector('#picture').files[0].arrayBuffer()),photoExtension:'png',photoWidth:dims[0],photoHeight:dims[1]};
          const bytes=await ToolWorkbook.apply(model,[update]);window.testUpdated=bytes;
          const saved=await ToolMobileFile.save({bytes,expectedVersion:before.version});
          const after=await ToolMobileFile.read(),backup=await ToolMobileFile.takeBackup();
          return {connected:saved.connected,mode:saved.mode,backupHash:await testHash(backup.bytes),hash:await testHash(after.bytes),version:saved.version};
        }""", dimensions)
        assert update['connected'] and update['mode'] == 'device-workbook' and update['backupHash'] == original_hash and update['hash'] != original_hash
        output = temporary / 'device-updated.xlsx'
        download_expression(page, '(await ToolMobileFile.read()).bytes', output)
        results['workbook'] = verify_workbook(args.workbook, output)
        results['save'] = update
        failures = page.evaluate("""async()=>{
          const api=ToolMobileFile,before=await api.read(),backup=await api.takeBackup(),out={};
          const assertState=async()=>{const a=await api.read(),b=await api.takeBackup();return a.version===before.version&&await testHash(a.bytes)===await testHash(before.bytes)&&await testHash(b.bytes)===await testHash(backup.bytes);};
          try{await api.save({bytes:before.bytes,expectedVersion:initialVersion});out.stale='accepted';}catch(e){out.stale=e.code;}
          out.staleUnchanged=await assertState();
          const put=IDBObjectStore.prototype.put;
          for(const key of ['backup','workbook']){
            IDBObjectStore.prototype.put=function(value,k){if(this.name==='workspace'&&k===key)throw new DOMException('TEST_'+key,'UnknownError');return put.apply(this,arguments);};
            try{await api.save({bytes:before.bytes,expectedVersion:before.version});out[key]='accepted';}catch(e){out[key]=e.code;}
            finally{IDBObjectStore.prototype.put=put;}
            out[key+'Unchanged']=await assertState();
          }
          IDBObjectStore.prototype.put=function(value,key){if(this.name==='workspace'&&key==='workbook')throw new DOMException('TEST_FULL','QuotaExceededError');return put.apply(this,arguments);};
          try{await api.save({bytes:before.bytes,expectedVersion:before.version});out.quota='accepted';}catch(e){out.quota=e.code;}
          finally{IDBObjectStore.prototype.put=put;}
          out.quotaUnchanged=await assertState();
          let checks=0;try{await api.save({bytes:before.bytes,expectedVersion:before.version,verifyCurrent:()=>++checks<3});out.revision='accepted';}catch(e){out.revision=e.code;}
          out.revisionChecks=checks;out.revisionUnchanged=await assertState();
          return out;
        }""")
        assert failures['stale'] == 'conflict' and failures['staleUnchanged']
        assert failures['backup'] == 'backup-failed' and failures['backupUnchanged']
        assert failures['workbook'] == 'write-failed' and failures['workbookUnchanged']
        assert failures['quota'] == 'storage-full' and failures['quotaUnchanged']
        assert failures['revision'] == 'stale' and failures['revisionChecks'] == 3 and failures['revisionUnchanged']
        results['atomic_failures'] = failures
        # A second document restores from durable storage without re-selecting the file.
        other = context.new_page()
        module_page(other, url)
        restored = other.evaluate("""async()=>{const state=await ToolMobileFile.restore(),r=await ToolMobileFile.read(),b=await ToolMobileFile.takeBackup();return {connected:state.connected,name:r.file.name,hash:await testHash(r.bytes),backupHash:await testHash(b.bytes),version:r.version};}""")
        assert restored['connected'] and restored['hash'] == update['hash'] and restored['backupHash'] == original_hash
        results['new_document_restore'] = restored
        # Distinct module instances have independent JS queues but one real IDB
        # transaction domain. Both saves use the same old version and bytes.
        page.evaluate('window.firstDeviceAPI=ToolMobileFile')
        page.add_script_tag(path=REPO / 'web/mobile-file.js')
        race = page.evaluate("""async()=>{
          await ToolMobileFile.restore();const before=await firstDeviceAPI.read();
          const results=await Promise.all([firstDeviceAPI,ToolMobileFile].map(api=>api.save({bytes:before.bytes,expectedVersion:before.version}).then(value=>({ok:true,version:value.version}),e=>({ok:false,code:e.code}))));
          return {results,after:(await ToolMobileFile.read()).version,before:before.version};
        }""")
        assert sum(item['ok'] for item in race['results']) == 1
        assert [item['code'] for item in race['results'] if not item['ok']] == ['conflict']
        assert race['after'] != race['before']
        results['concurrent_saves'] = race
        stale_tab = other.evaluate("""async version=>{try{await ToolMobileFile.save({bytes:(await ToolMobileFile.read()).bytes,expectedVersion:version});return 'accepted';}catch(e){return e.code;}}""", restored['version'])
        assert stale_tab == 'conflict'
        results['stale_other_tab'] = stale_tab
        clearing = page.evaluate("""async()=>{
          const api=ToolMobileFile,before=await api.read();const oldClear=IDBObjectStore.prototype.clear;
          IDBObjectStore.prototype.clear=function(){throw new DOMException('TEST_CLEAR','UnknownError');};
          let failure;try{await api.disconnect();failure='accepted';}catch(e){failure=e.code;}finally{IDBObjectStore.prototype.clear=oldClear;}
          const unchanged=(await api.read()).version===before.version;await api.disconnect();
          const after=await api.restore(),backup=await api.takeBackup();return {failure,unchanged,connected:after.connected,backup,mode:after.mode};
        }""")
        assert clearing['failure'] == 'storage-failed' and clearing['unchanged'] and not clearing['connected'] and clearing['backup'] is None
        results['clear'] = clearing
        missing = other.evaluate("""async()=>{try{await ToolMobileFile.read();return {code:'accepted'};}catch(e){return {code:e.code,state:ToolMobileFile.inspect()};}}""")
        assert missing['code'] == 'not-connected' and not missing['state']['connected'] and missing['state']['needsReconnect']
        results['other_tab_after_clear'] = {'code': missing['code'], 'connected': missing['state']['connected']}
        no_storage_context = browser.new_context()
        no_storage_page = no_storage_context.new_page()
        module_page(no_storage_page, url)
        unavailable = no_storage_page.evaluate("""async()=>{Object.defineProperty(window,'indexedDB',{configurable:true,value:null});try{await ToolMobileFile.restore();return 'accepted';}catch(e){return {code:e.code,connected:ToolMobileFile.inspect().connected,supported:ToolMobileFile.supported()};}}""")
        assert unavailable == {'code': 'unsupported', 'connected': False, 'supported': False}
        results['unavailable_storage'] = unavailable
        no_storage_context.close()
        browser.close()
    assert not errors and not external, (errors, external)
    assert sha256(args.workbook.read_bytes()).hexdigest() == original_hash
    results.update({'company_workbook_unchanged': True, 'page_errors': errors, 'external_requests': external,
                    'storage_scope': 'Device IndexedDB only; cloud original and OCR accuracy not tested'})
    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
