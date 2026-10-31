"""Same-origin SharePoint REST adapter tests in real Chromium.

The cookie session, SharePoint HTTP responses and synthetic XLSX files are
controlled fixtures. This does not assert access to an actual company tenant.
The production adapter, IndexedDB backups, hashes, concurrency and srcdoc
origin inheritance run unchanged. No credentials or company data are used.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import tempfile
from urllib.parse import parse_qs, quote, urlsplit

from openpyxl import Workbook
from playwright.sync_api import sync_playwright

from browser_test import REPO, serve_program

UNIQUE = '11111111-1111-4111-8111-111111111111'
FILE = "/sites/Tools/Shared Documents/team's 측정 #50%.xlsx"
ENCODED_FILE = quote(FILE, safe='/')
DIGEST = 'FIXTURE_FORM_DIGEST_NOT_PERSISTED'
SPY = """window.restFetchCalls=[];(()=>{const actual=window.fetch;window.fetch=function(url,options){restFetchCalls.push({url:String(url),credentials:options?.credentials,mode:options?.mode,cache:options?.cache,redirect:options?.redirect});return actual.apply(this,arguments)}})();"""


def workbook(value):
    output = BytesIO(); book = Workbook(); book.active['A1'] = value
    book.save(output); return output.getvalue()


class Rest:
    def __init__(self, original, origin):
        self.data = original; self.origin = origin; self.version = 1
        self.site = origin + '/sites/Tools'; self.calls = []; self.writes = []
        self.mode = 'normal'; self.metadata_reads = 0; self.race_at = None
        self.check_out_type = 2; self.checkouts = 0; self.checkins = 0; self.undos = 0; self.locked_original = None

    @property
    def tag(self):
        return '"' + UNIQUE + ',' + str(self.version) + '"'

    def route(self, route):
        request = route.request; parsed = urlsplit(request.url)
        headers = request.all_headers()
        assert headers.get('authorization') is None
        assert headers.get('cookie') == 'company-session=fixture-cookie'
        entry = {'method': request.method, 'path': parsed.path, 'query': parsed.query,
                 'if_match': request.headers.get('if-match'), 'digest': request.headers.get('x-requestdigest')}
        self.calls.append(entry)
        if self.mode in ['read401', 'read403']:
            return route.fulfill(status=int(self.mode[-3:]), body='{}')
        if parsed.path.endswith('/_api/contextinfo'):
            assert request.method == 'POST' and not request.headers.get('x-requestdigest')
            data = {'FormDigestValue': DIGEST if self.mode != 'bad-digest' else '', 'WebFullUrl': self.site}
            return route.fulfill(status=200, content_type='application/json', body=json.dumps(data))
        assert parsed.path.startswith('/sites/Tools/_api/web/GetFileByServerRelativePath(decodedUrl=@file)')
        alias = parse_qs(parsed.query)['@file'][0]
        assert alias[0] == alias[-1] == "'"
        assert alias[1:-1].replace("''", "'") == FILE
        if parsed.path.endswith('/CheckOut()'):
            assert request.method == 'POST' and request.headers.get('x-requestdigest') == DIGEST
            self.checkouts += 1
            if self.mode == 'checkout409' or self.check_out_type != 2:
                if self.mode == 'checkout409': self.check_out_type = 0
                return route.fulfill(status=409, body='{}')
            if self.mode == 'changed-at-checkout': self.data = workbook('someone changed before lock')
            self.locked_original = self.data; self.check_out_type = 0; self.version += 1
            if self.mode == 'checkout-disconnect': return route.abort('failed')
            return route.fulfill(status=204)
        if '/CheckIn(' in parsed.path:
            assert request.method == 'POST' and request.headers.get('x-requestdigest') == DIGEST
            assert self.check_out_type == 0
            self.checkins += 1
            if self.mode == 'checkin-disconnect': return route.abort('failed')
            self.check_out_type = 2; self.version += 1; self.locked_original = None
            return route.fulfill(status=204)
        if parsed.path.endswith('/UndoCheckOut()'):
            assert request.method == 'POST' and request.headers.get('x-requestdigest') == DIGEST
            self.undos += 1; assert self.check_out_type == 0
            self.data = self.locked_original; self.locked_original = None; self.check_out_type = 2; self.version += 1
            return route.fulfill(status=204)
        if parsed.path.endswith('/$value'):
            if request.method == 'GET':
                return route.fulfill(status=200, content_type='application/octet-stream', body=self.data)
            assert request.method == 'POST' and request.headers.get('x-http-method') == 'PUT'
            assert request.headers.get('x-requestdigest') == DIGEST
            assert self.check_out_type == 0
            incoming = request.post_data_buffer
            self.writes.append({'bytes': incoming, 'if_match': request.headers.get('if-match')})
            if self.mode.startswith('put'):
                return route.fulfill(status=int(self.mode[-3:]), body='{}')
            assert request.headers.get('if-match') == self.tag
            self.data = incoming; self.version += 1
            if self.mode == 'committed-disconnect':
                return route.abort('failed')
            if self.mode == 'wrong-verification':
                self.data = workbook('intervening different content')
            return route.fulfill(status=204, headers={'ETag': self.tag})
        assert request.method == 'GET'
        self.metadata_reads += 1
        if self.race_at == self.metadata_reads:
            self.version += 1
        data = {'Name': FILE.rsplit('/', 1)[-1], 'Length': str(len(self.data)), 'UniqueId': UNIQUE,
                'ServerRelativeUrl': FILE, 'ETag': self.tag, 'Exists': True, 'CheckOutType': self.check_out_type}
        if self.mode == 'wrong-id': data['UniqueId'] = '22222222-2222-4222-8222-222222222222'
        if self.mode == 'wrong-path': data['ServerRelativeUrl'] = '/sites/Tools/other.xlsx'
        if self.mode == 'bad-etag': data['ETag'] = '*'
        if self.mode == 'bad-size': data['Length'] = 101 * 1024 * 1024
        return route.fulfill(status=200, content_type='application/json', body=json.dumps(data))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--chromium', default='/usr/bin/chromium')
    args = parser.parse_args()
    source = (REPO / 'web/sharepoint-file.js').read_text()
    original, updated = workbook('original'), workbook('updated')
    cases = {}; errors = []; unexpected = []
    with tempfile.TemporaryDirectory(prefix='tool-sp-file-') as temporary:
        program = Path(temporary) / 'adapter.html'; program.write_text('<!doctype html><html><body>SharePoint test</body></html>')
        with serve_program(program) as url, sync_playwright() as playwright:
            origin = urlsplit(url).scheme + '://' + urlsplit(url).netloc
            browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=['--no-sandbox'])

            def new_case(context_file=ENCODED_FILE, framed=False):
                rest = Rest(original, origin); context = browser.new_context()
                context.add_cookies([{'name': 'company-session', 'value': 'fixture-cookie', 'url': origin}])
                page = context.new_page(); page.on('pageerror', lambda error: errors.append(str(error)))
                def route(request_route):
                    parsed = urlsplit(request_route.request.url)
                    if parsed.netloc != urlsplit(url).netloc:
                        unexpected.append(request_route.request.url); return request_route.abort()
                    if '/_api/' in parsed.path: return rest.route(request_route)
                    return request_route.continue_()
                context.route('**/*', route); page.goto(url)
                config = {'siteUrl': rest.site, 'fileUrl': context_file}
                if framed:
                    markup = '<!doctype html><script>window.ToolSharePointContext=' + json.dumps(config) + ';</script><script>' + SPY + source.replace('</script', '<\\/script') + '</script>'
                    page.evaluate('markup=>{const f=document.createElement("iframe");f.srcdoc=markup;document.body.appendChild(f)}', markup)
                    frame = page.frames[-1]; frame.wait_for_function('window.ToolSharePointFile')
                    return context, page, frame, rest
                page.evaluate('value=>{window.ToolSharePointContext=value}', config)
                page.add_script_tag(content=SPY + source)
                return context, page, page, rest

            def connect(frame, options=None):
                return frame.evaluate('async options=>{const value=await ToolSharePointFile.connect(options||{});window.snapshot=value;return {name:value.file.name,version:value.version,target:ToolSharePointFile.inspect().target}}', options)

            def save(frame, version=None):
                return frame.evaluate('async value=>{try{return {ok:true,value:await ToolSharePointFile.save({bytes:new Uint8Array(value.bytes),expectedVersion:value.version||snapshot.version,verifyCurrent:()=>true})}}catch(e){return {ok:false,code:e.code,message:e.message,status:ToolSharePointFile.inspect()}}}', {'bytes': list(updated), 'version': version})

            context, page, frame, rest = new_case(framed=True)
            assert frame.evaluate('location.origin') == 'null' and frame.evaluate('window.origin') == origin
            result = connect(frame)
            assert result['target']['uniqueId'] == UNIQUE and result['name'] == FILE.rsplit('/', 1)[-1]
            assert frame.evaluate('ToolSharePointFile.inspect().supported')
            for request in frame.evaluate('restFetchCalls'):
                assert request['credentials'] == request['mode'] == 'same-origin' and request['cache'] == 'no-store' and request['redirect'] == 'error'
                assert request['url'].startswith(origin + '/')
            cases['real_srcdoc_inherits_company_origin_and_session_cookie'] = True
            context.close()

            context, page, frame, rest = new_case(context_file='https://vanity.example.test/:x:/r' + ENCODED_FILE + '?d=private-share-proof&e=transient-ticket')
            result = connect(frame)
            assert result['target']['serverRelativeUrl'] == FILE
            assert all('vanity.example.test' not in call['url'] and 'private-share-proof' not in call['url'] and 'transient-ticket' not in call['url'] for call in frame.evaluate('restFetchCalls'))
            assert 'private-share-proof' not in json.dumps(result['target'])
            cases['vanity_share_link_extracts_only_exact_local_path_without_forwarding_proof'] = True
            context.close()

            context, page, frame, rest = new_case()
            result = connect(frame, {'target': {'siteUrl': rest.site, 'serverRelativeUrl': FILE}})
            assert result['target']['serverRelativeUrl'] == FILE
            cases['explicit_decoded_target_preserves_hash_and_percent_filename'] = True
            context.close()

            context, page, frame, rest = new_case()
            result = frame.evaluate('''async()=>{try{await ToolSharePointFile.connect({validate:async()=>{throw new Error('fixture invalid workbook')}});return null}catch(e){return e.message}}''')
            assert result and not frame.evaluate('ToolSharePointFile.inspect().connected')
            assert not frame.evaluate('(async()=>{const value=await ToolSharePointFile.restore();return value.connected})()')
            assert not rest.writes
            cases['workbook_validation_failure_never_persists_a_connection'] = True
            context.close()

            context, page, frame, rest = new_case()
            result = connect(frame)
            assert result['target']['serverRelativeUrl'] == FILE
            assert frame.evaluate('ToolSharePointFile.inspect().connected')
            count = len([x for x in rest.calls if x['path'].endswith('/$value')])
            frame.evaluate('async()=>{snapshot.bytes[0]=0;const value=await ToolSharePointFile.read();if(value.bytes[0]!==80)throw Error("cached bytes mutated");window.snapshot=value}')
            assert len([x for x in rest.calls if x['path'].endswith('/$value')]) == count
            cases['exact_path_unicode_apostrophe_metadata_and_immutable_memory_cache'] = True
            frame.evaluate('async()=>{const value=await ToolSharePointFile.read();value.target.serverRelativeUrl="/sites/Tools/wrong.xlsx";window.snapshot=await ToolSharePointFile.read()}')
            assert frame.evaluate('ToolSharePointFile.inspect().target.serverRelativeUrl') == FILE
            cases['returned_target_cannot_mutate_internal_selection'] = True
            saved = save(frame)
            assert saved['ok'] and rest.data == updated and len(rest.writes) == 1
            assert rest.writes[0]['if_match'].endswith(',2"')
            assert rest.checkouts == rest.checkins == 1 and rest.check_out_type == 2
            backup = frame.evaluate('async()=>Array.from((await ToolSharePointFile.takeBackup()).bytes)')
            assert bytes(backup) == original
            cases['conditional_write_fresh_digest_backup_and_verified_sha'] = True
            old = result['version']; count = len(rest.writes)
            refused = save(frame, old)
            assert not refused['ok'] and refused['code'] == 'conflict' and len(rest.writes) == count
            cases['stale_expected_version_never_overwrites'] = True
            page.reload(); page.evaluate('config=>{window.ToolSharePointContext=config}', {'siteUrl': rest.site, 'fileUrl': ENCODED_FILE}); page.add_script_tag(content=source)
            restored = page.evaluate('async()=>{const r=await ToolSharePointFile.restore();window.snapshot=r;return r.connected}')
            assert restored and page.evaluate('ToolSharePointFile.inspect().target.uniqueId') == UNIQUE
            assert bytes(page.evaluate('async()=>Array.from((await ToolSharePointFile.takeBackup()).bytes)')) == original
            stored = page.evaluate('''async()=>{const db=await new Promise((res,rej)=>{const q=indexedDB.open('tool-management-sharepoint-file-v1',2);q.onsuccess=()=>res(q.result);q.onerror=()=>rej(q.error)});const rows={};for(const name of ['targets','backups','locks'])rows[name]=await new Promise((res,rej)=>{const q=db.transaction(name).objectStore(name).getAll();q.onsuccess=()=>res(q.result);q.onerror=()=>rej(q.error)});db.close();return JSON.stringify(rows)}''')
            assert DIGEST not in stored and 'fixture-cookie' not in stored
            cases['real_idb_restore_and_backup_without_session_or_digest_storage'] = True
            page.reload(); page.evaluate('config=>{window.ToolSharePointContext=config}', {'siteUrl': rest.site, 'fileUrl': '/sites/Tools/Shared Documents/other.xlsx'}); page.add_script_tag(content=source)
            before = len(rest.calls); state = page.evaluate('ToolSharePointFile.restore()')
            assert not state['connected'] and state['needsReconnect'] and len(rest.calls) == before
            cases['changed_webpart_file_never_restores_old_destination'] = True
            context.close()

            context, page, frame, rest = new_case(); connect(frame)
            rest.check_out_type = 0  # metadata changes without changing ETag
            result = save(frame)
            assert result['code'] == 'file-locked' and not rest.writes and rest.checkouts == rest.checkins == rest.undos == 0
            assert not any(call['path'].endswith('/contextinfo') for call in rest.calls)
            cases['preexisting_checkout_even_same_etag_is_never_reused_or_undone'] = True; context.close()

            context, page, frame, rest = new_case(); connect(frame); rest.mode = 'checkout409'
            result = save(frame)
            assert result['code'] == 'file-locked' and not rest.writes and rest.checkouts == 1 and rest.undos == 0 and rest.check_out_type == 0
            assert not result['status']['uncertain'] and not result['status']['pendingCheckout']
            cases['concurrent_checkout_rejection_never_undoes_other_writer_lock'] = True; context.close()

            context, page, frame, rest = new_case(); connect(frame); rest.mode = 'changed-at-checkout'
            result = save(frame)
            assert result['code'] == 'conflict' and not rest.writes and rest.checkouts == rest.undos == 1 and rest.check_out_type == 2
            assert rest.data != original and not result['status']['uncertain']
            cases['post_checkout_original_sha_detects_intervening_change_and_preserves_it'] = True; context.close()

            context, page, frame, rest = new_case(); connect(frame); rest.mode = 'checkout-disconnect'
            result = save(frame)
            assert not result['ok'] and result['status']['uncertain'] and not rest.writes and rest.undos == 0 and rest.check_out_type == 0
            rest.mode = 'normal'
            page.reload(); page.evaluate('config=>{window.ToolSharePointContext=config}', {'siteUrl': rest.site, 'fileUrl': ENCODED_FILE}); page.add_script_tag(content=source)
            state = page.evaluate('async()=>{const value=await ToolSharePointFile.restore();window.snapshot=value;return ToolSharePointFile.inspect()}')
            assert state['uncertain'] and state['pendingCheckout'] and save(page)['code'] == 'requires-read' and not rest.writes and rest.undos == 0
            rest.check_out_type = 2; rest.version += 1
            page.evaluate('async()=>{window.snapshot=await ToolSharePointFile.read()}')
            assert not page.evaluate('ToolSharePointFile.inspect().pendingCheckout')
            cases['unknown_checkout_acquisition_journal_survives_reload_without_unsafe_undo'] = True; context.close()

            for mode, expected in [('put412', 'conflict'), ('put409', 'file-locked'), ('put423', 'file-locked'), ('put403', 'permission-denied'), ('put401', 'auth-required')]:
                context, page, frame, rest = new_case(); connect(frame); rest.mode = mode
                result = save(frame)
                assert not result['ok'] and result['code'] == expected and rest.data == original
                assert not result['status']['uncertain']
                assert rest.undos == 1 and rest.check_out_type == 2
                cases[mode + '_no_false_save_or_blind_retry'] = True; context.close()

            for mode in ['committed-disconnect', 'wrong-verification', 'checkin-disconnect']:
                context, page, frame, rest = new_case(); connect(frame); rest.mode = mode
                result = save(frame)
                assert not result['ok'] and result['status']['uncertain'] and len(rest.writes) == 1
                count = len(rest.calls); held = save(frame)
                assert held['code'] == 'requires-read' and len(rest.calls) == count and len(rest.writes) == 1
                rest.mode = 'normal'; frame.evaluate('async()=>{window.snapshot=await ToolSharePointFile.read()}')
                assert frame.evaluate('ToolSharePointFile.inspect().uncertain') and rest.undos == 0
                assert save(frame)['code'] == 'requires-read'
                # Manual resolution in SharePoint, not an automatic Undo that
                # could discard a successfully written draft.
                rest.check_out_type = 2; rest.version += 1
                frame.evaluate('async()=>{window.snapshot=await ToolSharePointFile.read()}')
                assert not frame.evaluate('ToolSharePointFile.inspect().uncertain')
                cases[mode + '_uncertain_until_explicit_fresh_read'] = True; context.close()

            context, page, frame, rest = new_case(); connect(frame)
            frame.evaluate("(()=>{const put=IDBObjectStore.prototype.put;IDBObjectStore.prototype.put=function(...args){if(this.name==='backups')throw new DOMException('full','QuotaExceededError');return put.apply(this,args)}})()")
            result = save(frame)
            assert not result['ok'] and result['code'] == 'backup-failed' and not rest.writes
            assert not any(call['path'].endswith('/contextinfo') for call in rest.calls)
            cases['backup_quota_failure_before_digest_or_write_preserves_original'] = True; context.close()

            context, page, frame, rest = new_case(); connect(frame); before = len(rest.calls)
            result = frame.evaluate('async bytes=>{try{await ToolSharePointFile.save({bytes:new Uint8Array(bytes),expectedVersion:snapshot.version,verifyCurrent:()=>false});return null}catch(e){return e.code}}', list(updated))
            assert result == 'stale' and len(rest.calls) == before and not rest.writes
            cases['changed_ui_revision_stops_before_network'] = True; context.close()

            for mode, expected in [('wrong-path','invalid-target'),('bad-etag','invalid-version'),('bad-size','invalid-target')]:
                context, page, frame, rest = new_case(); rest.mode = mode
                result = frame.evaluate('async()=>{try{await ToolSharePointFile.connect({});return null}catch(e){return e.code}}')
                assert result == expected and not frame.evaluate('ToolSharePointFile.inspect().connected')
                cases[mode + '_metadata_rejected'] = True; context.close()

            context, page, frame, rest = new_case(); connect(frame); rest.mode = 'wrong-id'
            result = save(frame); assert not result['ok'] and result['code'] == 'invalid-target' and not rest.writes
            cases['replaced_file_unique_id_refuses_write'] = True; context.close()

            context, page, frame, rest = new_case(); connect(frame); rest.mode = 'bad-digest'
            result = save(frame); assert not result['ok'] and result['code'] == 'invalid-digest' and not rest.writes
            cases['invalid_form_digest_no_write'] = True; context.close()

            context, page, frame, rest = new_case(); rest.race_at = 2
            result = frame.evaluate('async()=>{try{await ToolSharePointFile.connect({});return null}catch(e){return e.code}}')
            assert result == 'conflict' and not frame.evaluate('ToolSharePointFile.inspect().connected')
            cases['changing_read_etag_never_adopts_inconsistent_snapshot'] = True; context.close()

            context, page, frame, rest = new_case(); connect(frame)
            results = frame.evaluate('''async bytes=>Promise.all([1,2].map(async()=>{try{await ToolSharePointFile.save({bytes:new Uint8Array(bytes),expectedVersion:snapshot.version});return 'saved'}catch(e){return e.code}}))''', list(updated))
            assert results == ['saved', 'conflict'] and len(rest.writes) == 1
            cases['concurrent_same_version_writes_serialize_without_lost_update'] = True; context.close()

            context, page, frame, rest = new_case()
            for value in [origin + '/other/file.xlsx', 'https://other.example.test/sites/Tools/file.xlsx', '/sites/Tools/file.txt', '/sites/Tools/bad\\name.xlsx']:
                code = frame.evaluate('async fileUrl=>{try{await ToolSharePointFile.connect({fileUrl});return null}catch(e){return e.code}}', value)
                assert code in ['invalid-file', 'wrong-origin']
            assert not rest.calls
            page.evaluate("ToolSharePointContext.siteUrl='https://company.example.test/sites/Tools'")
            assert not page.evaluate('ToolSharePointFile.inspect().supported')
            code = frame.evaluate('async()=>{try{await ToolSharePointFile.connect({});return null}catch(e){return e.code}}')
            assert code == 'wrong-origin' and not rest.calls
            cases['public_origin_cannot_reuse_company_context_or_cookie'] = True
            context.close(); browser.close()
    assert not errors and not unexpected, (errors, unexpected)
    print(json.dumps({'cases': cases, 'count': len(cases), 'page_errors': errors, 'unexpected_requests': unexpected,
                      'real_company_session_used': False, 'original_sha256': sha256(original).hexdigest()}, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
