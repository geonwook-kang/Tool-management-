"""Real PC OCR/autosave for runtime photographs and explicitly erased variants.

A private JSON manifest supplies original image paths, expected workbook rows,
photo dates and optional normalized erase rectangles. Expected values stay in
Python assertions; the recognizer sees only image bytes and the real workbook.
Masking removes existing pixels with a neutral patch and never draws text.
This tests the same real offline worker, connected-file OPFS writer, backup,
photo conversion and OOXML preservation used by the PC program. Actual company
OneDrive synchronization is outside this test. Company inputs remain outside Git.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
import json
from pathlib import Path
import tempfile
import time

from PIL import Image, ImageDraw, ImageOps
from playwright.sync_api import sync_playwright

from browser_test import REPO, download_expression, serve_program
from connected_file_test import PICKER_SCRIPT, export_shared, ready, seed_file
from multi_owner_ocr_test import OBSERVE, target_reference, verify_increment


def arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--program', type=Path, default=Path('/workspace/outputs/Tool-Manager-PC-Final.html'))
    parser.add_argument('--workbook', type=Path)
    parser.add_argument('--manifest', type=Path, required=True)
    parser.add_argument('--diagnostics', type=Path, default=Path('/tmp/tool-pc-final-ocr.json'))
    parser.add_argument('--chromium', default='/usr/bin/chromium')
    parser.add_argument('--diagnose-only', action='store_true')
    parser.add_argument('--baseline', type=Path)
    parser.add_argument('--maximum-seconds', type=float)
    args = parser.parse_args()
    args.workbook = args.workbook or next(Path('/workspace/attachments').rglob('*.xlsx'), None)
    for label, path in [('program', args.program), ('workbook', args.workbook), ('manifest', args.manifest)]:
        if path is None or not path.is_file():
            parser.error('Supply --' + label + ' with an existing file')
    args.cases = json.loads(args.manifest.read_text())
    if not isinstance(args.cases, list) or not args.cases:
        parser.error('Manifest must be a nonempty list')
    for case in args.cases:
        if not all(isinstance(case.get(k), str) and case[k] for k in ['label', 'photo', 'name', 'serial', 'last', 'due']):
            parser.error('Every case needs label, photo, name, serial, last and due')
        if not Path(case['photo']).is_file():
            parser.error('Photo missing: ' + case['photo'])
        for rectangle in case.get('erase_rectangles', []):
            if len(rectangle) != 4 or any(not isinstance(v, (int, float)) or not 0 <= v <= 1 for v in rectangle) or rectangle[0] >= rectangle[2] or rectangle[1] >= rectangle[3]:
                parser.error('Erase rectangles are normalized x0,y0,x1,y1')
    return args


def prepare_variant(case, directory, index):
    source = Path(case['photo'])
    rectangles = case.get('erase_rectangles', [])
    if not rectangles:
        return source, {'kind': 'original', 'pixels_altered': False, 'source_sha256': sha256(source.read_bytes()).hexdigest()}
    with Image.open(source) as original:
        image = ImageOps.exif_transpose(original).convert('RGB')
    draw = ImageDraw.Draw(image)
    for rectangle in rectangles:
        x0, y0, x1, y1 = [round(v * (image.width if n % 2 == 0 else image.height)) for n, v in enumerate(rectangle)]
        # A neutral patch removes text; no identity or date is synthesized.
        draw.rectangle((x0, y0, x1, y1), fill=tuple(case.get('erase_color', [239, 239, 232])))
    output = directory / f'{index}-erased.jpg'
    image.save(output, format='JPEG', quality=97, subsampling=0)
    return output, {'kind': 'explicit-pixel-erasure', 'pixels_altered': True,
                    'source_sha256': sha256(source.read_bytes()).hexdigest(),
                    'tested_dimensions': list(image.size), 'normalized_rectangles': rectangles,
                    'jpeg_quality': 97}


def main():
    args = arguments()
    original_hash = sha256(args.workbook.read_bytes()).hexdigest()
    source_hashes = {case['photo']: sha256(Path(case['photo']).read_bytes()).hexdigest() for case in args.cases}
    report = {'program_sha256': sha256(args.program.read_bytes()).hexdigest(),
              'worker_results_replaced': False, 'expected_values_passed_to_recognizer': False,
              'actual_company_sync_tested': False, 'cases': [], 'failures': []}
    errors, requests = [], []
    with tempfile.TemporaryDirectory(prefix='tool-pc-final-ocr-') as name, serve_program(args.program) as url, sync_playwright() as pw:
        temporary = Path(name)
        prior_file = temporary / 'prior.xlsx'
        prior_file.write_bytes(args.workbook.read_bytes())
        browser = pw.chromium.launch(executable_path=args.chromium, headless=True, args=['--no-sandbox'])
        context = browser.new_context(accept_downloads=True, locale='ko-KR')
        context.add_init_script('(' + PICKER_SCRIPT + ')()')
        page = context.new_page()
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.goto(url)
        page.wait_for_function('window.ToolApp&&window.ToolOCR&&window.ToolRuntime')
        page.evaluate(OBSERVE)
        context.route('**/*', lambda route: (requests.append(route.request.url), route.abort()) if route.request.url.startswith(('http:', 'https:')) else route.continue_())
        context.set_offline(True)
        seed_file(page, args.workbook)
        page.evaluate("ToolSources.use('desktop')")
        page.evaluate('ToolApp.connectSource(false)')
        page.wait_for_function('ToolApp.state.source.connected&&ToolApp.state.model')
        ready(page)
        page.locator('#remember-matches').uncheck()
        for index, case in enumerate(args.cases):
            photo, variant = prepare_variant(case, temporary, index)
            reference = target_reference(prior_file, case['serial'])
            before = page.evaluate('({reads:ocrTimings.reads.length,inits:ocrTimings.initializations.length,saves:ToolApp.state.source.saveCount})')
            began = time.monotonic()
            page.locator('#photos').set_input_files(photo)
            page.wait_for_function('ToolApp.state.photos.length===1&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving', timeout=240000)
            evidence = page.evaluate('''async before=>{
                const p=ToolApp.state.photos[0],handle=await(await navigator.storage.getDirectory()).getFileHandle('company-shared.xlsx'),backup=await ToolConnectedFile.takeBackup();
                return {result:p.result,assessment:p.assessment,saved:!!p.saved,approval:p.approval,
                    sourceError:ToolApp.state.source.error,saveDelta:ToolApp.state.source.saveCount-before.saves,
                    preparedPhotoHash:await testHash(p.image.photoBytes),sourceHash:await testHash(await(await handle.getFile()).arrayBuffer()),backupHash:backup&&await testHash(backup.bytes),
                    inputs:Object.fromEntries(['name','serial','last','due'].map(k=>[k,document.getElementById(k).value])),
                    reads:ocrTimings.reads.slice(before.reads),initializations:ocrTimings.initializations.slice(before.inits)};
            }''', before)
            result = evidence['result'] or {}
            assessment = evidence['assessment'] or {}
            entry = {'label': case['label'], 'seconds': round(time.monotonic() - began, 3),
                     'tested_photo_sha256': sha256(photo.read_bytes()).hexdigest(), 'variant': variant,
                     'cold_worker': index == 0, **evidence}
            entry['ocr_calls'] = len(entry['reads'])
            entry['ocr_milliseconds'] = sum(read.get('milliseconds') or 0 for read in entry['reads'])
            failures = []
            expected = tuple(case[k] for k in ['name', 'serial', 'last', 'due'])
            for label, actual in [('assessment', tuple(assessment.get(k) for k in ['name', 'serial', 'last', 'due'])), ('visible inputs', tuple(evidence['inputs'].get(k) for k in ['name', 'serial', 'last', 'due']))]:
                if actual != expected:
                    failures.append(label + ' did not match expected row/dates')
            if result.get('failed') or result.get('identityRecoveryIncomplete'):
                failures.append('OCR failed or identity recovery incomplete')
            if not assessment.get('eligible') or not evidence['saved'] or evidence['saveDelta'] != 1 or evidence['approval'] != 'automatic':
                failures.append('Photo did not automatically save exactly once')
            identity = result.get('identityEvidence') or {}
            if case.get('expect_name_absent') and any(proof.get('name') == case['name'] and not proof.get('learned') for proof in identity.get('nameMatches', [])):
                failures.append('Erased owner unexpectedly had name evidence')
            if case.get('expect_serial_partial') and any(proof.get('serial') == case['serial'] and not proof.get('learned') for proof in identity.get('serialMatches', [])):
                failures.append('Erased serial unexpectedly had exact serial evidence')
            if case.get('expect_serial_partial') and assessment.get('matchMethod') not in ['partial-serial', 'gapped-serial', 'partial-serial-groups', 'serial-fragments', 'serial-subsequence']:
                failures.append('Partial serial did not use a partial matching method')
            if args.maximum_seconds is not None and entry['seconds'] > args.maximum_seconds:
                failures.append('Exceeded machine-specific time bound')
            prior_hash = sha256(prior_file.read_bytes()).hexdigest()
            if evidence['saved']:
                assert evidence['backupHash'] == prior_hash and evidence['sourceHash'] != prior_hash
                updated = temporary / f'saved-{index}.xlsx'
                export_shared(page, updated)
                entry['workbook_verification'] = verify_increment(prior_file, updated, reference, case, evidence['preparedPhotoHash'])
                backup = temporary / f'backup-{index}.xlsx'
                download_expression(page, '(await ToolConnectedFile.takeBackup()).bytes', backup)
                assert backup.read_bytes() == prior_file.read_bytes()
                entry['backup_byte_exact'] = True
                prior_file.write_bytes(updated.read_bytes())
            else:
                assert evidence['sourceHash'] == prior_hash and evidence['saveDelta'] == 0
            entry['failures'] = failures
            report['cases'].append(entry)
            report['failures'].extend(case['label'] + ': ' + failure for failure in failures)
            args.diagnostics.write_text(json.dumps(report, ensure_ascii=False, indent=2))
            print(json.dumps({'label': case['label'], 'seconds': entry['seconds'], 'ocr_calls': entry['ocr_calls'], 'saved': evidence['saved'], 'assessment': assessment, 'failures': failures}, ensure_ascii=False), flush=True)
            page.locator('#remove-photo').click()
            page.wait_for_function('ToolApp.state.photos.length===0&&!ToolApp.state.preparing&&!ToolApp.state.source.saving')
        report.update({'worker_initializations': page.evaluate('ocrTimings.initializations.length'), 'page_errors': errors, 'http_requests': requests})
        browser.close()
    assert not errors and not requests, (errors, requests)
    assert sha256(args.workbook.read_bytes()).hexdigest() == original_hash
    assert all(sha256(Path(path).read_bytes()).hexdigest() == digest for path, digest in source_hashes.items())
    report.update({'original_company_workbook_unchanged': True, 'original_source_photos_unchanged': True})
    if args.baseline:
        baseline = json.loads(args.baseline.read_text())
        assert [case['tested_photo_sha256'] for case in baseline['cases']] == [case['tested_photo_sha256'] for case in report['cases']]
        previous = sum(case['seconds'] for case in baseline['cases'])
        current = sum(case['seconds'] for case in report['cases'])
        report['speed_comparison'] = {'same_tested_image_bytes_and_order': True,
                                      'baseline_program_sha256': baseline['program_sha256'],
                                      'baseline_seconds': round(previous, 3), 'current_seconds': round(current, 3),
                                      'speedup': round(previous / current, 3),
                                      'scope': 'This browser and PC; not a guarantee for all PCs or photographs'}
    args.diagnostics.write_text(json.dumps(report, ensure_ascii=False, indent=2))
    if not args.diagnose_only:
        assert not report['failures'], report['failures']


if __name__ == '__main__':
    main()
