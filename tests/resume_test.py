"""Real IndexedDB auth-redirect photo checkpoints, not OCR accuracy/auth tests.

Photographs are local fixtures, never transmitted. Real File bytes/EXIF, JPEG
metadata, processed-image bytes, raw OCR proof and saved-state markers survive
reload. Token/config/url/handle objects are excluded by explicit field lists.
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

from browser_test import REPO, serve_program


def arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--program', type=Path, default=Path('/workspace/outputs/Tool-Manager-Fixed.html'))
    parser.add_argument('--chromium', default='/usr/bin/chromium')
    return parser.parse_args()


def module_page(page, url):
    origin = f'{urlsplit(url).scheme}://{urlsplit(url).netloc}'
    page.route('**/resume-fixture.html', lambda route: route.fulfill(status=200, content_type='text/html', body='<!doctype html><input id="photo" type="file">'))
    page.goto(origin + '/resume-fixture.html')
    for path in ['web/smax.js', 'web/ocr.js', 'web/automation.js', 'web/resume.js']:
        page.add_script_tag(path=REPO / path)
    page.evaluate("""()=>{
      window.testHash=async bytes=>Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))).map(n=>n.toString(16).padStart(2,'0')).join('');
      window.testContext={mode:'microsoft365',driveId:'drive-a',itemId:'item-a',shareUrl:'https://tenant.sharepoint.com/:x:/r/sites/team/book.xlsx?e=demo',uncertain:true};
      window.testTools=[{id:'tool:1',name:'홍길동',serial:'ABC12345',months:12,last:'2026-03-17',due:'2027-03-17',photoCount:1}];
      window.makePhoto=async()=>{
        const file=document.querySelector('#photo').files[0],bytes=new Uint8Array(await file.arrayBuffer());
        const text='이름: 홍길동\\nSerial: ABC12345\\n교정일자 2027.03.17\\n차기교정예정일자 2028.03.17';
        const result=ToolOCR.parse(text,testTools,{passes:[{text,source:'original',confidence:96,words:[]}]});
        // Actual normalized raw words and SMAX field crop provenance are also
        // preserved; this extra pass is tested separately from normal replay.
        const image={photoBytes:bytes,photoExtension:'jpg',photoWidth:40,photoHeight:30,url:'SECRET_DO_NOT_STORE'};
        return {id:'photo:1',file,url:URL.createObjectURL(file),image,draft:{toolId:'tool:1',name:'홍길동',serial:'ABC12345',last:'2027-03-17',due:'2028-03-17',photoKind:'calibration',clientSecret:'SECRET_DO_NOT_STORE'},
          result:{...result,access_token:'SECRET_DO_NOT_STORE',configuration:{password:'SECRET_DO_NOT_STORE'}},
          assessment:{eligible:true,toolId:'tool:1',name:'홍길동',serial:'ABC12345',last:'2027-03-17',due:'2028-03-17',reasons:[],accessToken:'SECRET_DO_NOT_STORE'},
          review:{toolId:'tool:1',name:'홍길동',serial:'ABC12345',last:'2027-03-17',due:'2028-03-17',photoKind:'calibration',...image},approval:'automatic',edited:false,manualEntry:false,error:'',
          saved:{id:'tool:1',name:'홍길동',serial:'ABC12345',last:'2027-03-17',due:'2028-03-17',photoKind:'calibration',version:'etag:already-saved',bindingIdentity:'["workspace-a","owner-a","drive-a","item-a"]',sourceName:'original-shared-a.xlsx',authorization:'SECRET_DO_NOT_STORE'},
          token:'SECRET_DO_NOT_STORE',handle:{kind:'file',secret:'SECRET_DO_NOT_STORE'},callback:()=>{},authConfig:{secret:'SECRET_DO_NOT_STORE'}};
      };
      window.rawQueue=()=>new Promise((resolve,reject)=>{const r=indexedDB.open('tool-management-auth-resume-v1',1);r.onsuccess=()=>{const db=r.result,tx=db.transaction('pending','readonly'),q=tx.objectStore('pending').get('queue');q.onsuccess=()=>resolve(q.result);tx.oncomplete=()=>db.close();tx.onerror=()=>reject(tx.error);};r.onerror=()=>reject(r.error);});
    }""")


def main():
    args = arguments()
    outcome, errors, external = {}, [], []
    with tempfile.TemporaryDirectory(prefix='tool-auth-resume-') as directory, serve_program(args.program) as url, sync_playwright() as playwright:
        photo = Path(directory) / 'pending-original.jpg'
        image = Image.new('RGB', (40, 30), (160, 190, 215))
        exif = Image.Exif(); exif[274] = 6; exif[315] = 'resume-byte-preservation-fixture'
        image.save(photo, quality=93, exif=exif)
        original = photo.read_bytes(); original_hash = sha256(original).hexdigest()
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=['--no-sandbox'])
        context = browser.new_context()
        page = context.new_page()
        page.on('pageerror', lambda error: errors.append(str(error)))
        context.on('request', lambda request: external.append(request.url) if urlsplit(request.url).scheme in {'http','https'} and urlsplit(request.url).netloc != urlsplit(url).netloc else None)
        module_page(page, url)
        page.locator('#photo').set_input_files(photo)
        saved = page.evaluate("""async()=>{
          const empty=await ToolResume.inspect(),p=await makePhoto();
          const smax={...p,id:'photo:2',result:{text:'ABC12345',passes:[{text:'ABC12345',source:'smax-field-serial-7-0',confidence:94,smaxRole:'serial',smaxLabelVerified:true,smaxLabelText:'Serial/Asset Number',smaxLabelBBox:{x0:1,y0:2,x1:3,y1:4},smaxValueBBox:{x0:5,y0:6,x1:7,y1:8},words:[{text:'ABC12345',confidence:94,bbox:{x0:10,y0:11,x1:12,y1:13},lineBox:{x0:9,y0:10,x1:14,y1:15},credentials:'SECRET_DO_NOT_STORE'}],accessToken:'SECRET_DO_NOT_STORE'}]},review:null,saved:null,approval:null,edited:true,sharedConflict:true,error:'Another user changed this photo slot; explicit confirmation required.'};
          const result=await ToolResume.save({context:{...testContext,accessToken:'SECRET_DO_NOT_STORE'},photos:[p,smax]});
          const stored=await rawQueue(),json=JSON.stringify(stored,(key,value)=>value instanceof Blob?{blob:true,name:value.name}:ArrayBuffer.isView(value)?{typedArray:true,bytes:value.byteLength}:value);
          return {empty,result,peek:await ToolResume.peek(),inspect:await ToolResume.inspect(),json,fileHash:await testHash(await stored.photos[0].file.arrayBuffer()),imageHash:await testHash(stored.photos[0].image.photoBytes),fileModified:stored.photos[0].file.lastModified,alias:stored.photos[0].review.photoBytes===stored.photos[0].image.photoBytes};
        }""")
        assert saved['empty'] == {'pending':False,'count':0}
        assert saved['result']['count'] == 2 and saved['result']['expiresAt'] - saved['result']['createdAt'] == 600000
        assert saved['peek']['context']['uncertain'] is True and saved['peek']['context']['mode'] == 'microsoft365'
        assert 'context' not in saved['inspect'] and 'photos' not in saved['inspect']
        assert 'SECRET_DO_NOT_STORE' not in saved['json'] and 'blob:' not in saved['json'] and 'callback' not in saved['json']
        assert saved['fileHash'] == saved['imageHash'] == original_hash and saved['alias']
        outcome['saved'] = {key:saved[key] for key in ['empty','result','inspect','fileHash','imageHash','alias']}
        mismatch = page.evaluate("""async()=>({take:await ToolResume.take({...testContext,itemId:'different-item'}),pending:(await ToolResume.inspect()).pending})""")
        assert mismatch == {'take':None,'pending':True}
        outcome['wrong_target_preserves_pending'] = mismatch
        # A second tab cannot replace another tab's pending explicit-login queue.
        other = context.new_page(); module_page(other,url); other.locator('#photo').set_input_files(photo)
        refused = other.evaluate("""async()=>{try{await ToolResume.save({context:testContext,photos:[await makePhoto()]});return 'accepted';}catch(e){return e.code;}}""")
        assert refused == 'resume-other-tab'
        outcome['other_tab_overwrite'] = refused
        # Matching workbook IDs do not let another tab claim or consume the
        # queue while its owner is returning from a company-login redirect.
        cross_tab_claim = other.evaluate("""async()=>{
          const before=await ToolResume.peek(),taken=await ToolResume.take(testContext),after=await ToolResume.peek();
          return {sameTab:before.sameTab,taken:!!taken,pending:!!after,count:after&&after.count,unchanged:after&&after.createdAt===before.createdAt};
        }""")
        assert cross_tab_claim == {'sameTab':False,'taken':False,'pending':True,'count':2,'unchanged':True}
        assert page.evaluate("async()=>{const p=await ToolResume.peek();return p.sameTab&&p.count===2}")
        outcome['other_tab_take_preserves_owner_queue'] = cross_tab_claim
        # Reload recreates all module memory, as an authorization redirect does.
        module_page(page,url)
        restored = page.evaluate("""async()=>{
          const peek=await ToolResume.peek(),r=await ToolResume.take(testContext),p=r.photos[0],q=r.photos[1];
          const parsed=ToolOCR.parse(p.result.text,testTools,{confidence:p.result.confidence,passes:p.result.passes});
          return {peek,r:{context:r.context,count:r.photos.length,fileHash:await testHash(await p.file.arrayBuffer()),imageHash:await testHash(p.image.photoBytes),name:p.file.name,type:p.file.type,modified:p.file.lastModified,
            hasUrl:'url'in p,hasHandle:'handle'in p,saved:p.saved,approval:p.approval,review:p.review&&{serial:p.review.serial,hasImage:p.review.photoBytes===p.image.photoBytes},
            assessment:ToolAutomation.assess({tools:testTools},parsed),metadata:q.result.passes[0],held:{sharedConflict:q.sharedConflict,edited:q.edited,review:q.review,approval:q.approval,saved:q.saved,error:q.error}},after:await ToolResume.inspect(),again:await ToolResume.take(testContext)};
        }""")
        r = restored['r']
        assert restored['peek']['sameTab'] and r['count'] == 2 and r['fileHash'] == r['imageHash'] == original_hash
        assert r['modified'] == saved['fileModified'] and r['name'] == photo.name and r['type'] == 'image/jpeg'
        assert not r['hasUrl'] and not r['hasHandle'] and r['review']['hasImage']
        assert r['saved']['version'] == 'etag:already-saved' and r['approval'] == 'automatic'
        assert r['saved']['bindingIdentity'] == '["workspace-a","owner-a","drive-a","item-a"]' and r['saved']['sourceName'] == 'original-shared-a.xlsx'
        assert r['held'] == {'sharedConflict':True,'edited':True,'review':None,'approval':None,'saved':None,'error':'Another user changed this photo slot; explicit confirmation required.'}
        assert r['assessment']['eligible'] and r['assessment']['serial'] == 'ABC12345'
        assert r['metadata']['smaxLabelVerified'] and r['metadata']['smaxRole'] == 'serial' and r['metadata']['words'][0]['lineBox']['x1'] == 14
        assert restored['after'] == {'pending':False,'count':0} and restored['again'] is None
        outcome['redirect_roundtrip'] = restored
        # Atomic one-time claims, even from independently queued module instances.
        page.locator('#photo').set_input_files(photo)
        page.evaluate("""async()=>{await ToolResume.save({context:testContext,photos:[await makePhoto()]});window.firstResume=ToolResume;}""")
        page.add_script_tag(path=REPO/'web/resume.js')
        claims = page.evaluate("""async()=>{const r=await Promise.all([firstResume.take(testContext),ToolResume.take(testContext)]);return {claims:r.filter(Boolean).length,remaining:await ToolResume.inspect()};}""")
        assert claims == {'claims':1,'remaining':{'pending':False,'count':0}}
        outcome['one_time_claim'] = claims
        page.evaluate("""async()=>{await ToolResume.save({context:testContext,photos:[await makePhoto()]});}""")
        failures = page.evaluate("""async()=>{
          const before=await ToolResume.peek(),put=IDBObjectStore.prototype.put;let redirected=false,code;
          IDBObjectStore.prototype.put=function(){if(this.name==='pending')throw new DOMException('TEST_FULL','QuotaExceededError');return put.apply(this,arguments);};
          try{await ToolResume.save({context:testContext,photos:[await makePhoto()]});redirected=true;}catch(e){code=e.code;}finally{IDBObjectStore.prototype.put=put;}
          const after=await ToolResume.peek();
          let empty,wrong,invalid;try{await ToolResume.save({context:testContext,photos:[]});empty='accepted';}catch(e){empty=e.code;}
          try{await ToolResume.save({context:{...testContext,itemId:'different-item'},photos:[await makePhoto()]});wrong='accepted';}catch(e){wrong=e.code;}
          try{await ToolResume.save({context:{mode:'microsoft365',shareUrl:'javascript:alert(1)'},photos:[await makePhoto()]});invalid='accepted';}catch(e){invalid=e.code;}
          return {redirected,code,unchanged:before.createdAt===after.createdAt&&before.count===after.count,empty,wrong,invalid};
        }""")
        assert failures == {'redirected':False,'code':'resume-storage-full','unchanged':True,'empty':'resume-pending-not-empty','wrong':'resume-context-conflict','invalid':'resume-context-invalid'}
        outcome['failed_checkpoint_aborts_redirect'] = failures
        expiry = page.evaluate("""async()=>{
          const originalNow=Date.now,peek=await ToolResume.peek();Date.now=()=>peek.expiresAt;
          let result;try{result={peek:await ToolResume.peek(),take:await ToolResume.take(testContext),stored:!!(await rawQueue())};}finally{Date.now=originalNow;}
          return result;
        }""")
        assert expiry == {'peek':None,'take':None,'stored':False}
        outcome['ten_minute_expiry_purged'] = expiry
        cleared = page.evaluate("""async()=>{await ToolResume.save({context:testContext,photos:[await makePhoto()]});await ToolResume.clear();return {peek:await ToolResume.peek(),stored:!!(await rawQueue())};}""")
        assert cleared == {'peek':None,'stored':False}
        outcome['explicit_clear'] = cleared
        new_context = browser.new_context(); unavailable_page = new_context.new_page(); module_page(unavailable_page,url); unavailable_page.locator('#photo').set_input_files(photo)
        unavailable = unavailable_page.evaluate("""async()=>{Object.defineProperty(window,'indexedDB',{value:null,configurable:true});let redirected=false,code;try{await ToolResume.save({context:testContext,photos:[await makePhoto()]});redirected=true;}catch(e){code=e.code;}return {redirected,code,fileRemains:document.querySelector('#photo').files.length===1};}""")
        assert unavailable == {'redirected':False,'code':'resume-storage-failed','fileRemains':True}
        outcome['unavailable_storage'] = unavailable
        new_context.close(); browser.close()
        assert photo.read_bytes() == original
    assert not errors and not external, (errors,external)
    outcome.update({'original_photo_unchanged':True,'page_errors':errors,'external_requests':external,'scope':'Device checkpoint only; real OAuth and camera accuracy not exercised'})
    print(json.dumps(outcome,ensure_ascii=False,indent=2))


if __name__ == '__main__':
    main()
