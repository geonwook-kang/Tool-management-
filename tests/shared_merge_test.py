"""Real OOXML rebase checks for concurrent shared-workbook photo uploads.

The workbook and pictures are generated fixtures. The production workbook
writer, ZIP/XML parser and rebase helper run unchanged in real Chromium.
There are no Microsoft/network-service stubs because this helper performs no
network requests and never writes an input workbook.
"""
from __future__ import annotations

from datetime import datetime
from hashlib import sha256
import json
from pathlib import Path
import tempfile
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as WorkbookImage
from openpyxl.styles import Font, PatternFill
from PIL import Image
from playwright.sync_api import sync_playwright

from browser_test import REPO, serve_program, download_expression
from smax_workbook_test import drawing_info


def fixture(directory):
    book = Workbook()
    sheet = book.active
    sheet.title = "Tools"
    sheet.append(["Cal Next", "Cal Last", "Serial", "Asset #", "Cal Interval", "Cal Cycle", "Label Photo", "SMAX Tool Photo", "Notes"])
    for index, owner in enumerate(["홍길동", "김가온", "이해솔"], start=1):
        sheet.append([datetime(2026, 2, 15), datetime(2025, 2, 15), "TEST000" + str(index), "Region_Fixture_" + owner, 12, "MONTHS", None, None, "keep-" + str(index)])
        row = index + 1
        for col in [1, 2]:
            sheet.cell(row, col).number_format = "yyyy-mm-dd"
            sheet.cell(row, col).font = Font(name="Arial", color="123456")
            sheet.cell(row, col).fill = PatternFill("solid", fgColor="EAF2FF")
        for slot, color in [("G", (180, index * 40, 70)), ("H", (70, 180, index * 40))]:
            file = directory / (slot + str(index) + ".png")
            Image.new("RGB", (80, 50), color).save(file)
            image = WorkbookImage(file)
            image.width, image.height = 120, 75
            sheet.add_image(image, slot + str(row))
    sheet["I2"] = "=1+2"
    sheet.column_dimensions["G"].width = 24
    sheet.column_dimensions["H"].width = 24
    file = directory / "base.xlsx"
    book.save(file)
    book.close()
    for name, kind, color in [("new-a.jpg", "JPEG", (245, 120, 60)), ("new-b.png", "PNG", (60, 120, 245)), ("new-c.png", "PNG", (180, 60, 180))]:
        Image.new("RGB", (80, 50), color).save(directory / name, format=kind)
    return file


def check_combined(original, updated, expected):
    before, after = load_workbook(original), load_workbook(updated)
    sheet = before["Tools"]
    count = 0
    for row in sheet:
        for old in row:
            new = after["Tools"][old.coordinate]
            assert new.value == expected.get(old.coordinate, old.value), (old.coordinate, new.value)
            assert old._style == new._style, old.coordinate
            count += 1
    assert len(before["Tools"]._images) == len(after["Tools"]._images) == 6
    assert before["Tools"].column_dimensions["G"].width == after["Tools"].column_dimensions["G"].width
    before.close()
    after.close()
    return count


def main():
    cases, errors, unexpected = {}, [], []
    with tempfile.TemporaryDirectory(prefix="tool-shared-merge-") as name:
        directory = Path(name)
        original = fixture(directory)
        original_hash = sha256(original.read_bytes()).hexdigest()
        program = directory / "index.html"
        program.write_text('<!doctype html><html><body><input id="book" type="file"><input id="pictures" type="file" multiple></body></html>')
        with serve_program(program) as url, sync_playwright() as playwright:
            browser = playwright.chromium.launch(executable_path="/usr/bin/chromium", headless=True, args=["--no-sandbox"])
            page = browser.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))
            page.on("request", lambda request: unexpected.append(request.url) if not request.url.startswith(url.rsplit("/", 1)[0]) else None)
            page.goto(url)
            for source in ["vendor/jszip.min.js", "workbook.js", "shared-merge.js"]:
                page.add_script_tag(path=REPO / "web" / source)
            page.locator("#book").set_input_files(original)
            page.locator("#pictures").set_input_files([directory / item for item in ["new-a.jpg", "new-b.png", "new-c.png"]])
            page.evaluate("""async()=>{
              window.originalBytes=new Uint8Array(await document.getElementById('book').files[0].arrayBuffer());
              window.baselineModel=await ToolWorkbook.load(originalBytes);
              window.photoBytes=await Promise.all(Array.from(document.getElementById('pictures').files).map(async file=>new Uint8Array(await file.arrayBuffer())));
              window.makeUpdate=(serial,index,kind='calibration',options={})=>{
                const tool=baselineModel.tools.find(tool=>tool.serial===serial);
                return {toolId:tool.id,name:tool.name,serial:tool.serial,photoKind:kind,last:'2026-02-15',due:'2027-02-15',
                  photoBytes:photoBytes[index].slice(),photoExtension:index===0?'jpg':'png',photoWidth:80,photoHeight:50,...options};
              };
              window.reportError=async action=>{try{await action();return null;}catch(error){return {code:error.code||'',message:error.message};}};
            }""")
            # Two engineers start from exactly the same original. The second
            # rebase must retain the first upload in another row and all styles.
            value = page.evaluate("""async()=>{
              window.first=makeUpdate('TEST0001',0);window.second=makeUpdate('TEST0002',1);
              window.firstBytes=await ToolWorkbook.apply(baselineModel,[first]);
              window.combined=await ToolSharedMerge.rebase({baselineModel,baselineBytes:originalBytes,currentBytes:firstBytes,updates:[second]});
              return {tools:combined.model.tools.map(tool=>({serial:tool.serial,last:tool.last,due:tool.due})),
                inputToolId:second.toolId,outputToolId:combined.updates[0].toolId,
                callerPhotoUnchanged:second.photoBytes.every((value,index)=>value===photoBytes[1][index])};
            }""")
            combined = directory / "combined.xlsx"
            download_expression(page, "combined.bytes", combined)
            assert value["callerPhotoUnchanged"]
            checked = check_combined(original, combined, {"A2": datetime(2027, 2, 15), "B2": datetime(2026, 2, 15), "A3": datetime(2027, 2, 15), "B3": datetime(2026, 2, 15)})
            with ZipFile(original) as old, ZipFile(combined) as new:
                _, _, a = drawing_info(old, "Tools")
                _, _, b = drawing_info(new, "Tools")
                assert set(a) == set(b)
                assert b[(2, 7)][0]["bytes"] == (directory / "new-a.jpg").read_bytes()
                assert b[(3, 7)][0]["bytes"] == (directory / "new-b.png").read_bytes()
                for slot in set(a) - {(2, 7), (3, 7)}:
                    assert a[slot] == b[slot], slot
            cases["two_clients_different_rows_merge_photos_dates_and_preserve_other_content"] = {"cells_and_styles": checked, "pictures": 6}

            # A label and SMAX upload occupy independent slots on the same tool.
            result = page.evaluate("""async()=>{
              const update=makeUpdate('TEST0001',1,'smax',{updateDue:true,last:'2025-02-15'});
              window.paired=await ToolSharedMerge.rebase({baselineModel,baselineBytes:originalBytes,currentBytes:firstBytes,updates:[update]});
              return {last:paired.model.tools[0].last,due:paired.model.tools[0].due,reboundLast:paired.updates[0].last};
            }""")
            assert result == {"last": "2026-02-15", "due": "2027-02-15", "reboundLast": "2026-02-15"}, result
            paired = directory / "paired.xlsx"
            download_expression(page, "paired.bytes", paired)
            with ZipFile(paired) as archive:
                _, _, pictures = drawing_info(archive, "Tools")
                assert pictures[(2, 7)][0]["bytes"] == (directory / "new-a.jpg").read_bytes()
                assert pictures[(2, 8)][0]["bytes"] == (directory / "new-b.png").read_bytes()
            cases["same_tool_label_and_smax_slots_merge_compatible_dates_without_cal_last_rollback"] = True

            result = page.evaluate("""async()=>{
              const changed=makeUpdate('TEST0001',1,'smax',{updateDue:true,due:'2028-02-15'});
              const latest=await ToolWorkbook.apply(baselineModel,[changed]);
              const incoming=makeUpdate('TEST0001',2,'smax',{updateDue:false,last:'1999-01-01',due:'2000-01-01'});
              return reportError(()=>ToolSharedMerge.rebase({baselineModel,baselineBytes:originalBytes,currentBytes:latest,updates:[incoming]}));
            }""")
            assert result and result["code"] == "shared-rebase-conflict", result
            cases["same_smax_slot_conflict_even_when_photo_only"] = True

            for label, script in {
                "same_photo_slot_changed_by_other_user": "const current=firstBytes;const incoming=makeUpdate('TEST0001',2);",
                "same_slot_equal_incoming_also_held_conservatively": "const current=firstBytes;const incoming=makeUpdate('TEST0001',0);",
                "same_date_field_changed_to_conflicting_value": "const current=firstBytes;const incoming=makeUpdate('TEST0001',2,'smax',{updateDue:true,due:'2026-10-01'});",
                "older_calibration_date_stays_rejected": "const current=originalBytes;const incoming=makeUpdate('TEST0001',2,'calibration',{last:'2024-01-01',due:'2025-01-01'});",
            }.items():
                result = page.evaluate("async()=>{" + script + "return reportError(()=>ToolSharedMerge.rebase({baselineModel,baselineBytes:originalBytes,currentBytes:current,updates:[incoming]}));}")
                assert result and result["code"] == "shared-rebase-conflict", (label, result)
                cases[label] = True

            # Mutations here represent edits made by another user in Excel,
            # not altered OCR values or a replacement implementation.
            mutations = {
                "changed_owner": "const cell=doc.getElementsByTagNameNS('*','c');Array.from(cell).find(c=>c.getAttribute('r')==='D3').getElementsByTagNameNS('*','t')[0].textContent='Region_Fixture_새담당';",
                "changed_serial": "Array.from(doc.getElementsByTagNameNS('*','c')).find(c=>c.getAttribute('r')==='C3').getElementsByTagNameNS('*','t')[0].textContent='CHANGED999';",
                "duplicate_serial_in_other_row": "Array.from(doc.getElementsByTagNameNS('*','c')).find(c=>c.getAttribute('r')==='C4').getElementsByTagNameNS('*','t')[0].textContent='TEST0002';",
                "changed_calibration_cycle": "Array.from(doc.getElementsByTagNameNS('*','c')).find(c=>c.getAttribute('r')==='E3').getElementsByTagNameNS('*','v')[0].textContent='6';",
                "retired_target": "const c=Array.from(doc.getElementsByTagNameNS('*','c')).find(c=>c.getAttribute('r')==='A3');while(c.firstChild)c.removeChild(c.firstChild);c.setAttribute('t','inlineStr');const is=doc.createElementNS(doc.documentElement.namespaceURI,'is'),t=doc.createElementNS(doc.documentElement.namespaceURI,'t');t.textContent='retired';is.appendChild(t);c.appendChild(is);",
            }
            for label, mutation in mutations.items():
                result = page.evaluate("""async mutation=>{
                  const zip=await JSZip.loadAsync(originalBytes),doc=new DOMParser().parseFromString(await zip.file('xl/worksheets/sheet1.xml').async('string'),'application/xml');
                  new Function('doc',mutation)(doc);zip.file('xl/worksheets/sheet1.xml',new XMLSerializer().serializeToString(doc));
                  const current=await zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
                  return reportError(()=>ToolSharedMerge.rebase({baselineModel,baselineBytes:originalBytes,currentBytes:current,updates:[second]}));
                }""", mutation)
                assert result and result["code"] == "shared-rebase-conflict", (label, result)
                cases[label] = True

            result = page.evaluate("""async()=>{
              const zip=await JSZip.loadAsync(originalBytes),part='xl/drawings/drawing1.xml';
              const doc=new DOMParser().parseFromString(await zip.file(part).async('string'),'application/xml');
              const anchor=Array.from(doc.documentElement.children).find(a=>a.getElementsByTagNameNS('*','row')[0].textContent==='2'&&a.getElementsByTagNameNS('*','col')[0].textContent==='6');
              let transform=anchor.getElementsByTagNameNS('*','xfrm')[0];
              if(!transform){transform=doc.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main','a:xfrm');anchor.getElementsByTagNameNS('*','spPr')[0].appendChild(transform);}
              transform.setAttribute('rot','5400000');zip.file(part,new XMLSerializer().serializeToString(doc));
              const current=await zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
              return reportError(()=>ToolSharedMerge.rebase({baselineModel,baselineBytes:originalBytes,currentBytes:current,updates:[second]}));
            }""")
            assert result and result["code"] == "shared-rebase-conflict", result
            cases["same_photo_rotation_edit_detected_even_with_unchanged_image_bytes"] = True

            # Moved rows retain identity; row IDs from the first reader must
            # not cause the update to be written to an unrelated row.
            result = page.evaluate("""async()=>{
              const zip=await JSZip.loadAsync(originalBytes),part='xl/worksheets/sheet1.xml';
              const doc=new DOMParser().parseFromString(await zip.file(part).async('string'),'application/xml');
              const row=Array.from(doc.getElementsByTagNameNS('*','row')).find(row=>row.getAttribute('r')==='3');
              row.setAttribute('r','5');for(const c of Array.from(row.children))c.setAttribute('r',c.getAttribute('r').replace(/3$/,'5'));row.parentNode.appendChild(row);zip.file(part,new XMLSerializer().serializeToString(doc));
              const draw='xl/drawings/drawing1.xml',d=new DOMParser().parseFromString(await zip.file(draw).async('string'),'application/xml');
              for(const anchor of Array.from(d.documentElement.children)){const from=Array.from(anchor.children).find(c=>c.localName==='from');const n=from&&Array.from(from.children).find(c=>c.localName==='row');if(n?.textContent==='2')n.textContent='4';}
              zip.file(draw,new XMLSerializer().serializeToString(d));const current=await zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
              const result=await ToolSharedMerge.rebase({baselineModel,baselineBytes:originalBytes,currentBytes:current,updates:[second]});
              return {row:result.model.tools.find(t=>t.serial==='TEST0002').row,toolId:result.updates[0].toolId,last:result.model.tools.find(t=>t.serial==='TEST0002').last};
            }""")
            assert result == {"row": 5, "toolId": "xl/worksheets/sheet1.xml:5", "last": "2026-02-15"}, result
            cases["moved_row_rebound_by_unique_serial_and_owner"] = True

            result = page.evaluate("""async()=>({
              missing:await reportError(()=>ToolSharedMerge.rebase({baselineModel,currentBytes:firstBytes,updates:[second]})),
              mismatched:await reportError(()=>ToolSharedMerge.rebase({baselineModel,baselineBytes:firstBytes,currentBytes:firstBytes,updates:[first]}))
            })""")
            assert result["missing"]["code"] == result["mismatched"]["code"] == "shared-rebase-invalid", result
            cases["missing_or_mismatched_baseline_fails_closed"] = True

            # Exercise the same parser against the uploaded Excel-produced
            # workbook (shared strings, multiple sheets and existing photos).
            # Its values/owner names remain runtime data outside this test.
            company = next(Path('/workspace/attachments').rglob('*.xlsx'), None)
            if company:
                company_hash = sha256(company.read_bytes()).hexdigest()
                page.locator('#book').set_input_files(company)
                selection = page.evaluate("""async()=>{
                  const source=new Uint8Array(await document.getElementById('book').files[0].arrayBuffer());
                  const model=await ToolWorkbook.load(source);
                  const eligible=model.tools.filter(tool=>tool.photoCount===1&&/^\\d{4}-\\d{2}-\\d{2}$/.test(tool.last)&&/^\\d{4}-\\d{2}-\\d{2}$/.test(tool.due)&&tool.due>tool.last&&model.tools.filter(other=>ToolWorkbook.normalizeSerial(other.serial)===ToolWorkbook.normalizeSerial(tool.serial)).length===1);
                  if(eligible.length<2)throw new Error('Need two unambiguous fixture rows');
                  const selected=[eligible[0],eligible.find(tool=>tool.sheet!==eligible[0].sheet)||eligible[1]];
                  const changes=selected.map((tool,index)=>({toolId:tool.id,name:tool.name,serial:tool.serial,last:tool.last,due:tool.due,photoKind:'calibration',photoBytes:photoBytes[index],photoExtension:index===0?'jpg':'png',photoWidth:80,photoHeight:50}));
                  const first=await ToolWorkbook.apply(model,[changes[0]]);
                  window.companyRebased=await ToolSharedMerge.rebase({baselineModel:model,baselineBytes:source,currentBytes:first,updates:[changes[1]]});
                  return selected.map(tool=>({sheet:tool.sheet,row:tool.row}));
                }""")
                output = directory / 'company-rebased.xlsx'
                download_expression(page, 'companyRebased.bytes', output)
                before, after = load_workbook(company), load_workbook(output)
                cells = 0
                for old_sheet in before:
                    assert old_sheet.max_row == after[old_sheet.title].max_row
                    assert old_sheet.max_column == after[old_sheet.title].max_column
                    for row in old_sheet:
                        for old in row:
                            new = after[old_sheet.title][old.coordinate]
                            assert old.value == new.value and old._style == new._style
                            cells += 1
                images = sum(len(sheet._images) for sheet in before)
                assert images == sum(len(sheet._images) for sheet in after)
                with ZipFile(company) as old_zip, ZipFile(output) as new_zip:
                    changed = {(entry['sheet'], entry['row']): index for index, entry in enumerate(selection)}
                    for title in before.sheetnames:
                        if not before[title]._images:
                            continue
                        _, _, old_pictures = drawing_info(old_zip, title)
                        _, _, new_pictures = drawing_info(new_zip, title)
                        assert set(old_pictures) == set(new_pictures)
                        headings = {cell.value: cell.column for cell in before[title][1]}
                        for location, old in old_pictures.items():
                            key = (title, location[0])
                            if key in changed and location[1] == headings['Label Photo']:
                                photo = directory / ('new-a.jpg' if changed[key] == 0 else 'new-b.png')
                                assert new_pictures[location][0]['bytes'] == photo.read_bytes()
                            else:
                                assert old == new_pictures[location]
                before.close()
                after.close()
                assert sha256(company.read_bytes()).hexdigest() == company_hash
                cases['uploaded_excel_two_sheet_photo_rebase_preserves_all_cells_styles_and_other_pictures'] = {'cells_and_styles': cells, 'pictures': images, 'original_unchanged': True}
            assert not errors and not unexpected, (errors, unexpected)
            browser.close()
        assert sha256(original.read_bytes()).hexdigest() == original_hash
    print(json.dumps({"cases": cases, "case_count": len(cases), "original_unchanged": True,
                      "page_errors": errors, "unexpected_requests": unexpected,
                      "scope": "Actual browser writer/ZIP/OOXML rebase; no real company or Graph service used."}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
