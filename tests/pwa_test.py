"""Real Chromium service-worker installation, offline use and safe updates.

No OCR response is replaced and no company data is needed for these shell
checks. Camera/OCR/workbook integration is covered by the application tests.
"""
from __future__ import annotations

import argparse
from contextlib import contextmanager
from functools import partial
from hashlib import sha256
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile
from threading import Thread
from urllib.parse import urlsplit

from playwright.sync_api import Error as PlaywrightError, sync_playwright

ROOT = Path(__file__).resolve().parents[1]


@contextmanager
def server(directory):
    requests = []

    class Handler(SimpleHTTPRequestHandler):
        def log_message(self, format, *args):
            pass

        def end_headers(self):
            self.send_header('Cache-Control', 'no-cache')
            super().end_headers()

        def do_GET(self):
            requests.append(('GET', self.path))
            if urlsplit(self.path).path.endswith('/private/company.xlsx'):
                data = b'PK-private-workbook-fixture'
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Length', str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            else:
                super().do_GET()

        def do_POST(self):
            requests.append(('POST', self.path))
            self.rfile.read(int(self.headers.get('Content-Length', '0')))
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b'private upload acknowledged')

    http = ThreadingHTTPServer(('127.0.0.1', 0), partial(Handler, directory=str(directory)))
    thread = Thread(target=http.serve_forever, daemon=True)
    thread.start()
    try:
        yield f'http://127.0.0.1:{http.server_port}/nested/Tool-management-/', requests
    finally:
        http.shutdown()
        thread.join()
        http.server_close()


CACHE_STATE = """async()=>{
 const values=[];
 for(const name of await caches.keys()) {
  if(name.startsWith('tool-manager-shell:'))values.push({name,keys:(await(await caches.open(name)).keys()).map(r=>r.url)});
 }
 return values;
}"""
UPDATE = """async()=>{
 const reg=await navigator.serviceWorker.getRegistration();window.lastInstallState='';
 reg.addEventListener('updatefound',()=>{const worker=reg.installing;if(worker){window.lastInstallState=worker.state;worker.addEventListener('statechange',()=>window.lastInstallState=worker.state);}});
 await reg.update();
 return !!reg;
}"""


def variant(folder, originals, version, fail_hash=False):
    old = json.loads(originals['build-info.json'])['version']
    assets = {entry['path'][2:]: originals[entry['path'][2:]] for entry in json.loads(originals['build-info.json'])['shell']}
    assets['index.html'] = assets['index.html'].replace(old.encode(), version.encode())
    assets['pwa.js'] = assets['pwa.js'].replace(old.encode(), version.encode())
    entries = [{'path': './' + name, 'sha256': sha256(data).hexdigest(), 'bytes': len(data)} for name, data in sorted(assets.items())]
    if fail_hash:
        next(entry for entry in entries if entry['path'] == './index.html')['sha256'] = '0' * 64
    worker = originals['service-worker.js'].decode().replace(old, version)
    worker = re.sub(r'const APP_ASSETS = .*?;\nconst CACHE_PREFIX', 'const APP_ASSETS = ' + json.dumps(entries, separators=(',', ':')) + ';\nconst CACHE_PREFIX', worker, count=1, flags=re.S)
    assert worker.count('const APP_ASSETS = ') == 1
    for name, data in assets.items():
        (folder / name).write_bytes(data)
    (folder / 'service-worker.js').write_text(worker)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--standalone', type=Path, default=Path('/workspace/outputs/Tool-Manager-Fixed.html'))
    parser.add_argument('--chromium', default='/usr/bin/chromium')
    args = parser.parse_args()
    outcomes = {}
    with tempfile.TemporaryDirectory(prefix='tool-pwa-test-') as temporary:
        root = Path(temporary)
        output = root / 'nested/Tool-management-'
        subprocess.run([sys.executable, str(ROOT / 'web/build_pwa.py'), '--standalone', str(args.standalone), '--output-dir', str(output)], check=True)
        originals = {str(path.relative_to(output)): path.read_bytes() for path in output.rglob('*') if path.is_file()}
        info = json.loads(originals['build-info.json'])
        manifest = json.loads(originals['manifest.webmanifest'])
        assert manifest['scope'] == './' and manifest['start_url'] == './index.html?source=homescreen' and 'share_target' not in manifest
        assert manifest['short_name'] == '툴 매니저'
        assert not re.search(r'/\*__[A-Z_]+__\*/', originals['index.html'].decode())
        assert all(entry['path'].startswith('./') for entry in info['shell'])
        assert info['private_data_included'] is False
        with server(root) as (url, requests), sync_playwright() as playwright:
            browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=['--no-sandbox'])
            context = browser.new_context(accept_downloads=True, locale='ko-KR')
            page = context.new_page()
            errors, external = [], []
            page.on('pageerror', lambda error: errors.append(str(error)))
            context.on('request', lambda request: external.append(request.url) if urlsplit(request.url).scheme in {'http','https'} and urlsplit(request.url).netloc != urlsplit(url).netloc else None)
            page.goto(url)
            page.wait_for_function('window.ToolPWA&&ToolPWA.inspect().offlineReady&&navigator.serviceWorker.controller', timeout=60000)
            page.wait_for_function('window.ToolApp&&!ToolApp.state.busy')
            assert page.url == url + 'index.html'
            outcomes['root_path_canonicalized_without_navigation'] = True
            assert page.evaluate('ToolPWA.inspect().version') == info['version']
            initial = page.evaluate(CACHE_STATE)
            assert len(initial) == 1 and len(initial[0]['keys']) == len(info['shell'])
            outcomes['verified_complete_precache'] = True
            # Exercise the real immutable self-export after user-visible
            # configuration and a decoded photo have entered the live app.
            exported = page.evaluate("""async()=>{
              const entries={'m365-client-id':'11111111-2222-4333-8444-555555555555',
                'm365-tenant-id':'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee',
                'cloud-workbook-url':'https://tenant-fixture.sharepoint.com/private-workbook?access=private-canary'};
              for(const [id,value] of Object.entries(entries)){const field=document.getElementById(id);if(field){field.value=value;field.setAttribute('value',value);}}
              const picture=document.createElement('canvas');picture.width=picture.height=3;picture.getContext('2d').fillRect(0,0,3,3);
              const blob=await new Promise(resolve=>picture.toBlob(resolve,'image/png'));
              await ToolApp.addPhotos([new File([blob],'private-photo-canary.png',{type:'image/png',lastModified:1})]);
              if(ToolApp.state.photos.length!==1)throw new Error('Actual photograph preparation did not run');
              const source=ToolDistribution.html(),doc=new DOMParser().parseFromString(source,'text/html');
              const result={shellNodes:doc.querySelectorAll('[data-pwa-shell]').length,
                externalScripts:Array.from(doc.scripts).filter(script=>script.hasAttribute('src')).map(script=>script.getAttribute('src')),
                manifestOrIcons:Array.from(doc.querySelectorAll('link')).filter(link=>/manifest|icon/.test(link.rel)).map(link=>link.getAttribute('href')),
                workerRegistration:Array.from(doc.scripts).some(script=>script.textContent.includes('navigator.serviceWorker.register(')),
                privateValues:Object.values(entries).concat(['private-photo-canary.png']).filter(value=>source.includes(value)),
                corePresent:!!doc.getElementById('runtime-assets')&&!!doc.getElementById('distribution-meta')};
              for(const photo of ToolApp.state.photos)URL.revokeObjectURL(photo.url);ToolApp.state.photos=[];ToolApp.state.selectedId=null;
              return result;
            }""")
            assert exported == {'shellNodes':0,'externalScripts':[],'manifestOrIcons':[],
                                'workerRegistration':False,'privateValues':[],'corePresent':True}, exported
            outcomes['standalone_self_export_omits_pwa_and_entered_private_data'] = True
            assert page.locator('#pwa-install').is_visible()
            page.locator('#pwa-install').click()
            assert page.locator('#pwa-install-help').is_visible()
            outcomes['install_menu_fallback'] = True
            await_private = page.evaluate("""async()=>{
              const a=await fetch('./private/company.xlsx?access_token=private-fixture');
              const b=await fetch('./upload',{method:'POST',body:'private-photo-fixture'});
              return [await a.text(),await b.text()];
            }""")
            assert await_private == ['PK-private-workbook-fixture','private upload acknowledged']
            assert page.evaluate(CACHE_STATE) == initial
            assert any(method == 'POST' for method, _ in requests)
            outcomes['workbook_photo_and_query_not_cached'] = True
            context.set_offline(True)
            page.reload()
            page.wait_for_function('window.ToolPWA&&ToolPWA.inspect().offlineReady&&window.ToolApp', timeout=60000)
            assert page.locator('h1').is_visible()
            assert page.evaluate('ToolPWA.inspect().version') == info['version']
            assert page.evaluate('!!window.ToolOCR&&!!window.ToolRuntime')
            outcomes['offline_reload_under_nested_scope'] = True
            page.goto(url + 'index.html?callback-fixture=preserved#fragment')
            page.wait_for_function('window.ToolPWA&&ToolPWA.inspect().offlineReady')
            assert page.url == url + 'index.html?callback-fixture=preserved#fragment'
            assert page.evaluate(CACHE_STATE) == initial
            outcomes['offline_index_callback_query_not_added_to_cache'] = True
            context.set_offline(False)
            # A new worker with an invalid required asset must leave the old
            # active shell and its complete cache untouched.
            variant(output, originals, 'failed-build-fixture', fail_hash=True)
            page.evaluate(UPDATE)
            page.wait_for_function("window.lastInstallState==='redundant'", timeout=60000)
            assert page.evaluate(CACHE_STATE) == initial
            context.set_offline(True)
            page.reload()
            page.wait_for_function('window.ToolPWA&&ToolPWA.inspect().offlineReady', timeout=60000)
            assert page.evaluate('ToolPWA.inspect().version') == info['version']
            outcomes['failed_update_retains_previous_offline_app'] = True
            context.set_offline(False)
            variant(output, originals, 'updated-build-fixture')
            page.evaluate(UPDATE)
            page.wait_for_function("navigator.serviceWorker.getRegistration().then(r=>!!r.waiting)", timeout=60000)
            page.wait_for_function('ToolPWA.inspect().updateAvailable')
            page.evaluate('ToolApp.state.busy=true')
            assert page.evaluate('ToolPWA.applyUpdate()') is False
            page.evaluate('ToolApp.state.busy=false')
            outcomes['busy_current_tab_prevents_update'] = True
            second = context.new_page()
            second.goto(url)
            second.wait_for_function('window.ToolPWA&&window.ToolApp&&!ToolApp.state.busy')
            second.evaluate("ToolApp.state.busy=true;window.unsavedCanary='retained'")
            assert page.evaluate('ToolPWA.applyUpdate()') is False
            assert page.evaluate('navigator.serviceWorker.getRegistration().then(r=>!!r.waiting)')
            outcomes['busy_other_tab_prevents_update'] = True
            second.evaluate('ToolApp.state.busy=false')
            with page.expect_navigation(wait_until='domcontentloaded', timeout=30000):
                try:
                    page.evaluate('ToolPWA.applyUpdate()')
                except PlaywrightError as error:
                    assert 'Execution context was destroyed' in str(error), str(error)
            page.wait_for_function("window.ToolPWA&&ToolPWA.inspect().offlineReady&&ToolPWA.inspect().version==='updated-build-fixture'", timeout=60000)
            assert second.evaluate('window.unsavedCanary') == 'retained'
            latest = page.evaluate(CACHE_STATE)
            assert len(latest) == 1 and 'updated-build-fixture' in latest[0]['name']
            outcomes['explicit_idle_update_and_other_tab_state_preserved'] = True
            assert not errors and not external, (errors, external)
            context.close()
            # iOS installation help is distinct from a claimed installation.
            ios = browser.new_context(user_agent='Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 Version/18.0 Mobile/15E148 Safari/604.1')
            iphone = ios.new_page()
            iphone.goto(url)
            iphone.wait_for_function('window.ToolPWA')
            iphone.locator('#pwa-install').click()
            assert 'Safari' in iphone.locator('#pwa-install-help').inner_text()
            assert iphone.evaluate('ToolPWA.inspect().installed') is False
            outcomes['ios_instructions_without_false_install_claim'] = True
            ios.close()
            browser.close()
        # Deployment directories with unrelated data are refused, never wiped.
        extra = output / 'private-company.xlsx'
        extra.write_bytes(b'private fixture')
        attempted = subprocess.run([sys.executable, str(ROOT/'web/build_pwa.py'), '--standalone', str(args.standalone), '--output-dir', str(output)], capture_output=True, text=True)
        assert attempted.returncode != 0 and extra.read_bytes() == b'private fixture'
        outcomes['unrelated_output_data_rejected_without_deleting_it'] = True
    print(json.dumps({'cases':outcomes,'page_errors':errors,'external_requests':external,
        'scope':'Real Chromium SW/cache and UI; no claim of physical iPhone/Android installation or OCR accuracy.'}, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
