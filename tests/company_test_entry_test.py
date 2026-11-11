"""Large synthetic XLSX through the isolated company-copy testing entry.

The fixture exceeds 100 MiB because an unrelated archive sheet contains a
valid, incompressible PNG, not padding or a malformed archive. The real file
reader, matching policy, OOXML writer and download run in Chromium. OCR replies
are controlled; this does not claim actual photo accuracy, compatibility with
the inaccessible company-wide workbook, or company server synchronization.
"""
from __future__ import annotations

import argparse
from datetime import datetime
from hashlib import sha256
import json
import os
from pathlib import Path
import tempfile
import time
from urllib.parse import urlsplit
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as WorkbookImage
from openpyxl.styles import PatternFill
from PIL import Image
from playwright.sync_api import sync_playwright

from browser_test import serve_program
from smax_workbook_test import drawing_info, rel_path


SERIAL = 'TEST739581246'
HEADERS = ['Cal Next', 'Cal Last', 'Serial', 'Asset #', 'Cal Interval',
           'Cal Cycle', 'Label Photo', 'SMAX Tool Photo']
OLD_PROFILES = '{"synthetic_untouched_profile":"old-profile-739581"}'
GUARDS = r"""(() => {
 window.testUnsafe={nativePicker:0,nativeWriter:0,idb:[]};
 localStorage.setItem('tool-manager-pc-source-v1','desktop');
 localStorage.setItem('tool-manager-workbook-source','microsoft365');
 localStorage.setItem('tool-management-profiles-v1',
   '{"synthetic_untouched_profile":"old-profile-739581"}');
 Object.defineProperty(window,'showOpenFilePicker',{configurable:true,value:async()=>{
   testUnsafe.nativePicker++;throw new Error('Company test must not use a native original picker');
 }});
 if(typeof FileSystemFileHandle!=='undefined')FileSystemFileHandle.prototype.createWritable=async()=>{
   testUnsafe.nativeWriter++;throw new Error('Company test must not write a native original');
 };
 const open=indexedDB.open.bind(indexedDB);
 indexedDB.open=(name,...args)=>{
   testUnsafe.idb.push(name);
   throw new Error('Company test must not access persisted workbook stores: '+name);
 };
})();"""


def file_hash(path):
    digest = sha256()
    with Path(path).open('rb') as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def entry_hash(archive, name):
    digest = sha256()
    with archive.open(name) as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def fixture(path, photo, archive_photo=None, duplicate=False, extra_rows_per_sheet=0):
    book = Workbook()
    for index, name in enumerate(['Synthetic South', 'Synthetic North']):
        sheet = book.active if index == 0 else book.create_sheet()
        sheet.title = name
        sheet.append(HEADERS)
        serial = SERIAL if index == 0 or duplicate else 'OTHER4821759'
        owner = '홍길동' if index == 0 else '김하늘'
        sheet.append([datetime(2027, 1, 5), datetime(2026, 1, 5), serial,
                      'Fixture_' + owner, 12, 'MONTHS'])
        sheet['A2'].number_format = sheet['B2'].number_format = 'yyyy-mm-dd'
        sheet['A2'].fill = PatternFill(fgColor='DDEEFF', fill_type='solid')
        sheet.merge_cells('A4:D4')
        sheet['A4'] = 'Unselected merged-cell record'
        sheet['E4'] = '=SUM(1,2)'
        sheet.add_image(WorkbookImage(photo), 'G2')
        sheet.add_image(WorkbookImage(photo), 'H2')
        for number in range(extra_rows_per_sheet):
            row = number + 10
            values = [datetime(2027, 1, 5), datetime(2026, 1, 5),
                      f'REG{index}{number:06d}', 'Fixture_보관담당', 12, 'MONTHS']
            for column, value in enumerate(values, 1):
                sheet.cell(row, column, value)
            sheet.cell(row, 1).number_format = sheet.cell(row, 2).number_format = 'yyyy-mm-dd'
    archive = book.create_sheet('Archive only')
    archive['A1'] = 'This sheet has no supported tool headers and must be preserved'
    archive['B3'] = '=SUM(5,7)'
    archive.add_image(WorkbookImage(archive_photo or photo), 'C5')
    book.save(path)
    book.close()


def install_ocr(page):
    page.evaluate(r"""serial=>{
      const parse=ToolOCR.parse;
      ToolOCR={...ToolOCR,recognize:async(_file,tools)=>{
        const text='Serial: '+serial+'\n교정일자 2027.01.05\n차기교정예정일자 2028.01.05';
        return parse(text,tools,{confidence:96,passes:[{text,source:'original',confidence:96,words:[]}]});
      }};
    }""", SERIAL)


def ready(page):
    page.wait_for_function('window.ToolApp&&window.ToolCompanyTest&&document.getElementById("company-test-panel")&&'
                           '!ToolApp.state.busy&&!ToolApp.state.restoring&&'
                           '!ToolApp.state.preparing&&!ToolApp.state.source.saving', timeout=360000)


def assert_isolated(page):
    unsafe = page.evaluate('testUnsafe')
    assert unsafe == {'nativePicker': 0, 'nativeWriter': 0, 'idb': []}, unsafe
    assert page.evaluate("localStorage.getItem('tool-manager-pc-source-v1')") == 'desktop'
    assert page.evaluate("localStorage.getItem('tool-management-profiles-v1')") == OLD_PROFILES
    return unsafe


def verify_output(original, updated, expected_photo_hash):
    before, after = load_workbook(original), load_workbook(updated)
    assert before.sheetnames == after.sheetnames
    permitted = {('Synthetic South', 'A2'): datetime(2028, 1, 5),
                 ('Synthetic South', 'B2'): datetime(2027, 1, 5)}
    checked = 0
    for old_sheet in before:
        new_sheet = after[old_sheet.title]
        assert (old_sheet.max_row, old_sheet.max_column) == (new_sheet.max_row, new_sheet.max_column)
        assert list(old_sheet.merged_cells.ranges) == list(new_sheet.merged_cells.ranges)
        for row in old_sheet:
            for old in row:
                new = new_sheet[old.coordinate]
                assert new.value == permitted.get((old_sheet.title, old.coordinate), old.value), (old_sheet.title, old.coordinate)
                assert new._style == old._style
                if (old_sheet.title, old.coordinate) not in permitted:
                    assert new.data_type == old.data_type
                checked += 1
    assert sum(len(sheet._images) for sheet in before) == sum(len(sheet._images) for sheet in after) == 5
    before.close(); after.close()
    with ZipFile(original) as old, ZipFile(updated) as new:
        assert new.testzip() is None
        old_names, new_names = set(old.namelist()), set(new.namelist())
        assert old_names <= new_names
        changed = {name for name in old_names if entry_hash(old, name) != entry_hash(new, name)}
        sheet, drawing, old_pictures = drawing_info(old, 'Synthetic South')
        _, _, new_pictures = drawing_info(new, 'Synthetic South')
        assert changed <= {sheet, drawing, rel_path(drawing), '[Content_Types].xml'}, changed
        assert set(old_pictures) == set(new_pictures)
        assert sha256(new_pictures[(2, 7)][0]['bytes']).hexdigest() == expected_photo_hash
        assert old_pictures[(2, 8)] == new_pictures[(2, 8)]
        added = {name for name in new_names - old_names if not name.endswith('/')}
        assert len(added) == 1 and next(iter(added)).startswith('xl/media/')
        largest = max(old.infolist(), key=lambda entry: entry.file_size)
        assert largest.filename not in changed and largest.file_size > 100 * 1024 * 1024
    return {'cells_and_styles_checked': checked, 'pictures': 5,
            'unrelated_large_photo_bytes_preserved': True,
            'unsupported_sheet_formulas_and_images_preserved': True,
            'changed_parts': sorted(changed)}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--program', type=Path, default=Path('/workspace/outputs/Tool-Manager-Company-Test.html'))
    parser.add_argument('--chromium', default='/usr/bin/chromium')
    parser.add_argument('--diagnostics', type=Path, default=Path('/tmp/company-test-entry-report.json'))
    args = parser.parse_args()
    assert args.program.is_file(), args.program
    errors, external = [], []
    report = {'actual_company_workbook_tested': False, 'actual_company_sync_tested': False,
              'ocr_responses_controlled': True, 'cases': {}}
    with tempfile.TemporaryDirectory(prefix='company-test-volume-') as directory:
        temp = Path(directory)
        small_photo, huge_photo = temp/'small.png', temp/'unrelated-large-photo.png'
        Image.new('RGB', (160, 100), (73, 148, 191)).save(small_photo)
        # Valid 108MB RGB image data keeps ZIP compression honest: the large
        # file must exercise real image-bearing XLSX preservation end to end.
        print('Creating valid >100MiB image-bearing synthetic workbook', flush=True)
        image = Image.frombytes('RGB', (6000, 6000), os.urandom(6000 * 6000 * 3))
        image.save(huge_photo, format='PNG', compress_level=0)
        image.close()
        original, duplicate = temp/'synthetic-company.xlsx', temp/'duplicate-company.xlsx'
        fixture(original, small_photo, huge_photo, extra_rows_per_sheet=500)
        fixture(duplicate, small_photo, duplicate=True)
        shared_source = temp/'shared-drawing-source.xlsx'
        fixture(shared_source, small_photo)
        shared = temp/'shared-drawing-company.xlsx'
        # The unsupported archive sheet also owns drawing1. Supported-sheet
        # metadata deliberately omits it, so the audit must inspect every
        # retained worksheet relationship before allowing drawing1 to change.
        with ZipFile(shared_source) as source, ZipFile(shared, 'w') as output:
            for member in source.infolist():
                content = source.read(member.filename)
                if member.filename == 'xl/worksheets/_rels/sheet3.xml.rels':
                    relations = ET.fromstring(content)
                    for relation in relations:
                        if relation.attrib.get('Type', '').endswith('/drawing'):
                            relation.set('Target', '/xl/drawings/drawing1.xml')
                    content = ET.tostring(relations, encoding='utf-8', xml_declaration=True)
                output.writestr(member, content)
        corrupt = temp/'corrupt-company.xlsx'
        data = bytearray(duplicate.read_bytes())
        with ZipFile(duplicate) as archive:
            member = archive.getinfo('xl/media/image1.png')
            offset = member.header_offset + 30 + len(member.filename.encode()) + len(member.extra)
            data[offset + member.compress_size // 2] ^= 1
        corrupt.write_bytes(data)
        assert 100*1024*1024 < original.stat().st_size < 250*1024*1024
        original_hash = file_hash(original)
        report['fixture_bytes'] = original.stat().st_size
        print('Synthetic fixture bytes:', original.stat().st_size, flush=True)
        with serve_program(args.program) as url, sync_playwright() as pw:
            browser = pw.chromium.launch(executable_path=args.chromium, headless=True, args=['--no-sandbox'])
            context = browser.new_context(locale='ko-KR', accept_downloads=True)
            context.add_init_script(GUARDS)
            def network(route):
                if urlsplit(route.request.url).scheme in {'http', 'https'} and urlsplit(route.request.url).netloc != urlsplit(url).netloc:
                    external.append(route.request.url); return route.abort()
                return route.continue_()
            context.route('**/*', network)
            page = context.new_page()
            page.on('pageerror', lambda error: errors.append(str(error)))
            page.goto(url)
            ready(page)
            assert page.evaluate('ToolSources.mode()') == 'device'
            assert not page.evaluate('ToolSources.inspect().connected')
            denied = page.evaluate("['desktop','microsoft365','sharepoint','shared'].map(mode=>{try{ToolSources.use(mode);return false;}catch(_){return true;}})")
            assert all(denied)
            assert not page.locator('#remember-matches').is_checked()
            install_ocr(page)
            began = time.monotonic()
            with page.expect_file_chooser() as event:
                page.locator('#quick-workbook').click()
            chooser = event.value
            assert chooser.element.get_attribute('id') == 'workbook'
            chooser.set_files(original)
            page.wait_for_function('ToolApp.state.model&&!ToolApp.state.busy', timeout=360000)
            report['input_load_seconds'] = round(time.monotonic()-began, 3)
            metadata = page.evaluate('ToolCompanyTest.report()')
            assert metadata['input']['bytes'] > 100*1024*1024, metadata
            assert metadata['input']['toolCount'] == 1002 and len(metadata['input']['allSheetNames']) == 3, metadata
            assert metadata['input']['mediaCount'] == 5 and metadata['input']['duplicateSerialCount'] == 0, metadata
            assert not page.evaluate('ToolApp.state.source.connected')
            copy = temp/'byte-identical-test-copy.xlsx'
            with page.expect_download(timeout=360000) as event:
                page.locator('#company-test-copy').click()
            assert '_전사테스트' in event.value.suggested_filename
            event.value.save_as(copy)
            assert file_hash(copy) == original_hash
            report['cases']['separate_test_copy_is_byte_identical_to_input'] = True
            began = time.monotonic()
            page.locator('#photos').set_input_files(small_photo)
            page.wait_for_function('ToolApp.state.preparedResult&&!ToolApp.state.busy&&!ToolApp.state.preparing', timeout=360000)
            report['conversion_seconds'] = round(time.monotonic()-began, 3)
            result = page.evaluate(r"""async()=>{
              const p=ToolApp.state.photos[0];
              return {approval:p.approval,assessment:p.assessment,
                photoHash:Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',p.image.photoBytes))).map(n=>n.toString(16).padStart(2,'0')).join(''),
                filename:ToolApp.state.preparedResult.name,source:ToolApp.state.source,report:ToolCompanyTest.report()};
            }""")
            assert result['approval'] == 'automatic' and result['assessment']['eligible'], result
            assert result['assessment']['serial'] == SERIAL and result['assessment']['name'] == '홍길동', result
            assert '_전사테스트' in result['filename'], result['filename']
            assert result['source']['saveCount'] == 0 and not result['source']['connected'], result['source']
            assert result['report']['conversion']['verified'], result['report']
            with page.expect_download(timeout=30000) as event:
                page.locator('#company-test-report').click()
            audit_file = temp/'downloaded-audit.json'
            event.value.save_as(audit_file)
            downloaded_audit = json.loads(audit_file.read_text())
            assert downloaded_audit['conversion']['verified']
            output = temp/'company-test-result.xlsx'
            with page.expect_download(timeout=360000) as event:
                page.locator('#automatic-result-download').click()
            assert '_전사테스트' in event.value.suggested_filename
            event.value.save_as(output)
            report['cases']['over_100MiB_real_XLSX_pipeline'] = verify_output(original, output, result['photoHash'])
            report['audit'] = result['report']
            assert file_hash(original) == original_hash
            report['cases']['original_file_unchanged'] = True
            report['large_fixture_isolation'] = assert_isolated(page)
            page.goto(url)
            ready(page)
            assert not page.evaluate('ToolApp.state.model'), 'Test workbook unexpectedly restored'
            install_ocr(page)
            page.locator('#workbook').set_input_files(duplicate)
            page.wait_for_function('ToolApp.state.model&&!ToolApp.state.busy', timeout=60000)
            page.locator('#photos').set_input_files(small_photo)
            ready(page)
            held = page.evaluate('({eligible:ToolApp.state.photos[0].assessment.eligible,prepared:!!ToolApp.state.preparedResult,report:ToolCompanyTest.report()})')
            assert not held['eligible'] and not held['prepared'], held
            assert held['report']['input']['duplicateSerialCount'] > 0, held
            report['cases']['duplicate_registered_serial_stays_held'] = True
            report['duplicate_fixture_isolation'] = assert_isolated(page)
            page.goto(url)
            ready(page)
            install_ocr(page)
            page.locator('#workbook').set_input_files(shared)
            page.wait_for_function('ToolApp.state.model&&!ToolApp.state.busy', timeout=60000)
            owner_metadata = page.evaluate('ToolApp.state.model.companyTest.sheetParts')
            assert len(owner_metadata) == 2 and all(sheet['name'] != 'Archive only' for sheet in owner_metadata), owner_metadata
            page.locator('#photos').set_input_files(small_photo)
            ready(page)
            blocked = page.evaluate('({eligible:ToolApp.state.photos[0].assessment.eligible,prepared:!!ToolApp.state.preparedResult,error:ToolApp.state.preparationError,report:ToolCompanyTest.report()})')
            assert blocked['eligible'] and not blocked['prepared'] and blocked['error'], blocked
            assert not blocked['report']['conversion']['verified'], blocked
            assert 'xl/drawings/drawing1.xml' in blocked['report']['conversion']['unexpectedParts'], blocked
            report['cases']['drawing_shared_with_unsupported_sheet_blocks_output'] = True
            report['shared_drawing_fixture_isolation'] = assert_isolated(page)
            page.goto(url)
            ready(page)
            page.locator('#workbook').set_input_files(corrupt)
            ready(page)
            assert not page.evaluate('ToolApp.state.model'), 'Corrupt workbook was accepted'
            assert not page.evaluate('ToolApp.state.preparedResult')
            assert '실패' in page.locator('#status').inner_text(), page.locator('#status').inner_text()
            report['cases']['CRC_corruption_stays_rejected'] = True
            unsafe = assert_isolated(page)
            report['cases']['no_native_or_persisted_original_access'] = unsafe
            report['cases']['old_PC_preferences_and_profiles_unchanged'] = True
            browser.close()
    assert not errors and not external, (errors, external)
    report['page_errors'], report['external_requests'] = errors, external
    report['case_count'] = len(report['cases'])
    args.diagnostics.write_text(json.dumps(report, ensure_ascii=False, indent=2))
    print(json.dumps({'case_count': report['case_count'], 'fixture_bytes': report['fixture_bytes'],
                      'load_seconds': report['input_load_seconds'], 'conversion_seconds': report['conversion_seconds'],
                      'original_unchanged': True, 'page_errors': errors, 'external_requests': external,
                      'report': str(args.diagnostics)}, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
