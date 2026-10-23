"""Real app/OPFS mixed SMAX and calibration-photo uploads.

Only ToolOCR.recognize responses are controlled; each response is parsed by the
production ToolOCR parser. Classification, assessment, duplicate/conflict
guards, image conversion, workbook writing, OPFS backups and UI run unchanged.
Generated fixtures do not reproduce the user's original SMAX screenshot.
"""
from hashlib import sha256
import json
from pathlib import Path
import tempfile
from urllib.parse import urlsplit
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from PIL import Image, ImageDraw
from playwright.sync_api import sync_playwright

from browser_test import arguments, serve_program
from connected_file_test import PICKER_SCRIPT, export_shared, ready, seed_file
from smax_workbook_test import verify_output, xml_signature
from torque_dates_test import SERIAL, LAST, DUE


def fixture(temporary, scenario, kind, index):
    path = temporary / f"{scenario}--{kind}-{index}.png"
    image = Image.new("RGB", (420, 280), (180, 220, 245) if kind == "smax" else (245, 220, 180))
    draw = ImageDraw.Draw(image)
    draw.rectangle((30, 40, 390, 240), fill="white", outline="black", width=3)
    draw.text((55, 95), f"{kind.upper()} {scenario} {index}", fill="black", font_size=20)
    image.save(path)
    return path


def install_ocr(page, due):
    page.evaluate("""({serial,last,due,labelDue})=>{
      const parse=ToolOCR.parse;
      window.smaxMockCalls=[];
      window.ToolOCR={...ToolOCR,recognize:async(file,tools)=>{
        const tool=tools.find(item=>item.serial===serial),label=/--label-\\d+\\.png$/i.test(file.name);
        const text=label?tool.name+'\\nSerial: '+serial+'\\n교정일자 '+last.replaceAll('-','.')+
          '\\n차기교정예정일자 '+labelDue.replaceAll('-','.'):
          'SMAX\\nTool Details\\nProduct: TEST TORQUE TOOL\\nSerial / Asset Number: '+serial+
          '\\nTool Owner: GEONWOOK KANG\\nCalibration Due Date: '+due+
          '\\nCreated By: UNRELATED USER\\nCreated On: 01/01/2025\\nModified On: 11/03/2027';
        const result=parse(text,tools,{confidence:95,passes:[{text,confidence:95,source:'original'}]});
        smaxMockCalls.push({file:file.name,type:result.documentType,serial:result.serial});return result;
      }};
    }""", {"serial": SERIAL, "last": LAST, "due": due, "labelDue": DUE})


def record(page):
    return page.evaluate("""async()=>{
      const file=await(await navigator.storage.getDirectory()).getFileHandle('company-shared.xlsx');
      const backup=await ToolConnectedFile.takeBackup();
      return {hash:await testHash(await(await file.getFile()).arrayBuffer()),
        backupHash:backup?await testHash(backup.bytes):null,saves:ToolApp.state.source.saveCount,
        prepared:!!ToolApp.state.preparedResult,error:ToolApp.state.source.error,calls:smaxMockCalls,
        photos:ToolApp.state.photos.map(p=>({file:p.file.name,kind:p.review?.photoKind||p.assessment?.photoKind,
          approval:p.approval,saved:!!p.saved,review:!!p.review,assessment:p.assessment,draft:p.draft,
          bytes:p.saved?Array.from(p.image.photoBytes):null}))};
    }""")


def expected_photos(data, temporary, prefix):
    images = {}
    for index, photo in enumerate(data["photos"]):
        if not photo["saved"]:
            continue
        output = temporary / f"{prefix}-{index}-prepared.jpg"
        output.write_bytes(bytes(photo["bytes"]))
        images[photo["kind"]] = output
    return images


def verify_app_output(original, updated, smax, calibration=None, **dates):
    # Unlike direct writer fixtures, app photos are JPEG-converted. Its first
    # .jpg needs a MIME declaration; validate only that required addition before
    # reusing the strict writer verifier for every cell, anchor and other part.
    baseline, mime_added = original, False
    with ZipFile(original) as old, ZipFile(updated) as new:
        part = "[Content_Types].xml"
        if old.read(part) != new.read(part):
            before, after = ET.fromstring(old.read(part)), ET.fromstring(new.read(part))
            original_rules = {xml_signature(node) for node in before}
            extra = [node for node in after if xml_signature(node) not in original_rules]
            assert extra and all(node.tag.endswith("}Default")
                and node.attrib == {"Extension": node.attrib.get("Extension"), "ContentType": "image/jpeg"}
                and node.attrib["Extension"] in {"jpg", "jpeg"} for node in extra), "Unexpected MIME/part changes"
            for node in extra:
                after.remove(node)
            assert xml_signature(before) == xml_signature(after), "Existing MIME rules changed"
            baseline = updated.with_name(updated.stem + "-jpeg-mime-baseline.xlsx")
            with ZipFile(baseline, "w") as normalized:
                for info in old.infolist():
                    normalized.writestr(info, new.read(part) if info.filename == part else old.read(info.filename))
            mime_added = True
    result = verify_output(baseline, updated, smax, calibration, **dates)
    if mime_added:
        result["changed_zip_parts"].append("[Content_Types].xml")
        result["only_required_jpeg_mime_added"] = True
    return result


def run_case(browser, url, args, temporary, scenario):
    context = browser.new_context(accept_downloads=True, locale="ko-KR")
    context.add_init_script("(" + PICKER_SCRIPT + ")()")
    errors, requests = [], []
    context.on("request", lambda request: requests.append(request.url)
               if urlsplit(request.url).scheme in {"http", "https"}
               and urlsplit(request.url).netloc != urlsplit(url).netloc else None)
    page = context.new_page()
    page.on("pageerror", lambda error: errors.append(str(error)))
    page.on("dialog", lambda dialog: dialog.accept())
    page.goto(url)
    ready(page)
    assert page.evaluate("!!window.ToolSMAX"), "Build SMAX modules before running this app suite"
    baseline = seed_file(page, args.workbook)
    page.locator("#connect-onedrive").click()
    page.wait_for_function("ToolApp.state.source.connected&&!!ToolApp.state.model")
    ready(page)
    target = page.evaluate("serial=>{const t=ToolApp.state.model.tools.find(t=>t.serial===serial);return {name:t.name,last:t.last,due:t.due};}", SERIAL)
    due = "01/22/2027" if scenario == "pair-conflict" else "unreadable" if scenario == "photo-only" else "12/22/2026"
    install_ocr(page, due)
    kinds = ["label", "smax"] if scenario in {"pair", "pair-conflict"} else ["smax", "smax"] if scenario == "duplicate" else ["label"] if scenario == "subsequent" else ["smax"]
    if scenario == "manual-photo-only":
        page.locator("#automatic-mode").uncheck()
    inputs = [fixture(temporary, scenario, kind, index) for index, kind in enumerate(kinds)]
    page.locator("#photos").set_input_files(inputs)
    page.wait_for_function("count=>ToolApp.state.photos.length===count&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", arg=len(inputs), timeout=30000)
    if scenario == "manual-photo-only":
        page.locator("#recognize-one").click()
        ready(page)
        assert page.locator("#photo-kind").input_value() == "smax"
        assert page.locator("#smax-update-due").is_checked()
        page.locator("#smax-update-due").uncheck()
        assert page.locator("#last").evaluate("el=>el.readOnly&&!el.required")
        assert page.locator("#due").evaluate("el=>el.readOnly&&!el.required")
        page.locator("#confirm").check()
        page.locator("#save-review").click()
        ready(page)
    first = record(page)
    if scenario in {"pair-conflict", "duplicate"}:
        assert first["hash"] == baseline and first["saves"] == 0 and not first["prepared"], first
        assert not any(p["saved"] or p["review"] for p in first["photos"]), first
        assert all(p["assessment"]["reasons"] for p in first["photos"]), first
        result = {"blocked_and_source_byte_identical": True}
    else:
        assert first["saves"] == 1 and not first["error"], first
        assert first["backupHash"] == baseline, first
        assert all(p["saved"] and p["review"] for p in first["photos"]), first
        assert all(call["type"] == ("calibration" if kind == "label" else "smax")
                   for call, kind in zip(first["calls"], kinds)), first["calls"]
        if scenario != "subsequent":
            smax_index = next(i for i, p in enumerate(first["photos"]) if p["kind"] == "smax")
            page.locator("#photo-list .photo-card").nth(smax_index).click()
            assert page.locator("#photo-kind").input_value() == "smax"
            assert page.locator("#last").evaluate("el=>el.readOnly&&!el.required")
            assert page.locator("#name").input_value() == target["name"]
            assert page.locator("#serial").input_value() == SERIAL
            assert "GEONWOOK KANG" in page.locator("#photo-kind-note").inner_text()
            assert target["name"] in page.locator("#photo-kind-note").inner_text()
        images = expected_photos(first, temporary, scenario)
        output = temporary / f"{scenario}-first.xlsx"
        export_shared(page, output)
        dates = {} if scenario in {"photo-only", "manual-photo-only"} else {"due": DUE}
        if scenario in {"pair", "subsequent"}:
            dates["last"] = LAST
        result = verify_app_output(args.workbook, output, images.get("smax"), images.get("calibration"), **dates)
        if scenario == "subsequent":
            first_bytes = output.read_bytes()
            install_ocr(page, "12/22/2026")
            following = fixture(temporary, scenario, "smax", 1)
            page.locator("#photos").set_input_files(following)
            page.wait_for_function("ToolApp.state.photos.length===2&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=30000)
            final = record(page)
            assert final["saves"] == 2 and final["backupHash"] == sha256(first_bytes).hexdigest(), final
            assert final["photos"][1]["draft"]["last"] == LAST
            images = expected_photos(final, temporary, "subsequent-final")
            updated = temporary / "subsequent-final.xlsx"
            export_shared(page, updated)
            result = verify_app_output(args.workbook, updated, images["smax"], images["calibration"], last=LAST, due=DUE)
            result["later_smax_preserves_new_label_cal_last"] = True
    assert not errors, errors
    assert not requests, requests
    assert not page.evaluate("ToolApp.state.source.error")
    context.close()
    return result


def main():
    args = arguments(__doc__)
    original_hash = sha256(args.workbook.read_bytes()).hexdigest()
    with tempfile.TemporaryDirectory(prefix="tool-smax-app-") as directory, serve_program(args.program) as url, sync_playwright() as playwright:
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
        result = {scenario: run_case(browser, url, args, Path(directory), scenario)
                  for scenario in ["smax-only", "pair", "pair-conflict", "duplicate", "photo-only", "subsequent", "manual-photo-only"]}
        browser.close()
    assert sha256(args.workbook.read_bytes()).hexdigest() == original_hash
    print(json.dumps({"program": str(args.program), "program_sha256": sha256(args.program.read_bytes()).hexdigest(),
                      "scenarios": result, "original_company_workbook_unchanged": True,
                      "recognition_accuracy": "not measured: OCR responses controlled; parser/app/storage/writer are real"}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
