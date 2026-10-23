"""SMAX field roles, document classification and OCR orchestration policies.

The real bundled parser, classifier, assessor and recognizer execute. Synthetic
OCR text/worker responses verify field roles and safety; this test does not
measure OCR accuracy on the user's photographs. The company XLSX is read only.
Writer and real connected-file preservation cases live in smax_workbook_test.
"""
from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path
import tempfile
from urllib.parse import urlsplit

from PIL import Image, ImageDraw
from playwright.sync_api import sync_playwright

from browser_test import arguments, serve_program


SERIAL, DUE = "0716805829", "2026-12-22"
ENGLISH_OWNER = "Ken Kang"


POLICIES = r"""() => {
    const model=ToolApp.state.model,target=model.tools.find(tool=>tool.serial==='0716805829');
    if(!target)throw new Error('The company fixture needs the torque tool.');
    const other=model.tools.find(tool=>tool.serial!==target.serial);
    if(!other)throw new Error('A second registered serial is needed for conflict checks.');
    const base=(serial=target.serial,due='12/22/2026',owner='Ken Kang')=>
        'SMAX\nTool Detail\nSerial/Asset Number '+serial+'\nTool Owner '+owner+
        '\nCalibration Due Date '+due+'\nTool Code TQ6\nSGPN ID 884422\nCreated By Demo Operator'+
        '\nModified 06/22/2026 10:15 AM';
    const bbox={x0:20,y0:20,x1:200,y1:50};
    const crop=(role,text,confidence=95,verified=true)=>({text,confidence,source:'smax-field-'+role+'-test',
        smaxRole:role,smaxLabelVerified:verified,
        smaxLabelText:role==='serial'?'Serial/Asset Number':role==='owner'?'Tool Owner':'Calibration Due Date',
        smaxLabelBBox:bbox,smaxValueBBox:{x0:220,y0:20,x1:450,y1:50}});
    const specs={
        exact_registered_serial_and_us_due:{text:base(),eligible:true,due:'2026-12-22',updateDue:true},
        iso_calibration_due:{text:base(target.serial,'2026-12-22'),eligible:true,due:'2026-12-22',updateDue:true},
        existing_due_can_keep_the_same_date_while_replacing_photo:{text:base(target.serial,target.due),eligible:true,due:target.due,updateDue:true},
        us_date_order_is_month_then_day:{text:base(target.serial,'06/12/2027'),eligible:true,due:'2027-06-12',updateDue:true},
        english_owner_is_displayed_and_korean_owner_is_resolved_by_serial:{text:base(target.serial,'12/22/2026','Unrelated English Owner'),eligible:true,due:'2026-12-22',updateDue:true,owner:'Unrelated English Owner'},
        toolcode_sgpn_and_creator_other_serials_are_not_primary_identity:{text:base()+'\nTool Code '+other.serial+'\nSGPN ID '+other.serial+'\nCreated By '+other.serial,eligible:true,due:'2026-12-22',updateDue:true},
        modified_and_created_dates_cannot_replace_calibration_due:{text:base()+'\nCreated 01/01/2030\nModified 12/31/2031',eligible:true,due:'2026-12-22',updateDue:true},
        adjacent_audit_timestamps_end_the_due_field:{text:'SMAX\nSerial/Asset Number '+target.serial+'\nTool Owner Ken Kang\nCalibration Due Date 12/22/2026\nCreated 01/01/2030\nModified 12/31/2031',eligible:true,due:'2026-12-22',updateDue:true},
        missing_due_is_photo_only:{text:base(target.serial,'[unreadable]'),eligible:true,updateDue:false},
        absent_due_field_is_photo_only:{text:'SMAX\nTool Detail\nSerial/Asset Number '+target.serial+'\nTool Owner Ken Kang',eligible:true,updateDue:false},
        low_confidence_due_crop_is_photo_only:{passes:[
            {text:base(target.serial,'[unreadable]'),confidence:95,source:'original'},crop('due','12/22/2026',20)],eligible:true,updateDue:false},
        verified_due_crop_updates_next_only:{passes:[
            {text:base(target.serial,'[unreadable]'),confidence:95,source:'original'},crop('due','12/22/2026')],eligible:true,updateDue:true,due:'2026-12-22'},
        unverified_value_crop_cannot_write_dates:{passes:[
            {text:base(target.serial,'[unreadable]'),confidence:95,source:'original'},crop('due','12/22/2026',95,false)],eligible:true,updateDue:false},
        modified_field_crop_cannot_be_due_evidence:{passes:[
            {text:base(target.serial,'[unreadable]'),confidence:95,source:'original'},
            {...crop('due','12/22/2026'),smaxLabelText:'Modified'}],eligible:true,updateDue:false},
        explicit_audit_field_in_crop_body_cannot_inherit_due_role:{passes:[
            {text:base(target.serial,'[unreadable]'),confidence:95,source:'original'},
            crop('due','Modified 12/31/2031')],eligible:true,updateDue:false},
        high_confidence_due_disagreement:{passes:[
            {text:base(),confidence:95,source:'original'},crop('due','12/23/2026')],eligible:false},
        interrupted_smax_field_read_is_not_automatically_approved:{passes:[
            {text:base(),confidence:95,source:'original'},
            {text:'',confidence:null,source:'smax-field-incomplete'}],eligible:false},
        exploratory_smax_text_cannot_prove_missing_primary:{passes:[
            {text:'SMAX\nTool Detail\nTool Owner Ken Kang',confidence:95,source:'original'},
            {text:base(),confidence:95,source:'date-probe-wide'},
            {text:base(),confidence:95,source:'label-probe'}],eligible:false,documentType:'unknown-smax',noProbeEvidence:true},
        stale_due_before_existing_workbook_due:{text:base(target.serial,'01/01/2026'),eligible:false},
        due_before_existing_calibration_last:{text:base(target.serial,'01/01/2025'),eligible:false},
        unknown_exact_serial_cannot_use_english_owner:{text:base('9999999999'),eligible:false},
        exact_korean_owner_cannot_replace_unknown_serial:{text:base('9999999999','12/22/2026',target.name),eligible:false},
        partial_serial_is_not_a_smax_primary_key:{text:base('071680'),eligible:false},
        dropped_leading_zero_is_not_exact_serial:{text:base('716805829'),eligible:false},
        letter_o_is_not_zero_in_serial:{text:base('O716805829'),eligible:false},
        excel_asset_owner_string_is_not_the_smax_serial_number:{text:base(target.asset),eligible:false},
        two_different_high_confidence_serial_fields:{text:base()+'\nSerial/Asset Number '+other.serial,eligible:false},
        low_confidence_serial_and_header:{text:base(),confidence:20,eligible:false},
        no_confidence_is_not_automatic_evidence:{text:base(),confidence:null,eligible:false},
        duplicate_registered_serial_is_ambiguous:{text:base(),mutation:'duplicate',eligible:false},
        disposed_registered_tool_is_not_updated:{text:base(),mutation:'disposed',eligible:false},
        toolcode_only_screenshot_is_not_a_label_photo:{text:'SMAX\nTool Detail\nTool Code '+target.serial+'\nTool Owner Ken Kang\nCalibration Due Date 12/22/2026',eligible:false,documentType:'unknown-smax'},
        plain_smax_text_without_structured_primary_field_is_not_label:{text:'SMAX\nTool Owner Ken Kang\n'+target.serial+'\n12/22/2026',eligible:false,documentType:'unknown-smax'},
        generic_serial_and_due_are_not_sufficient_smax_structure:{text:'Serial Number '+target.serial+'\nCalibration Due Date 12/22/2026',eligible:false,documentType:'calibration'},
    };
    const output={};
    for(const [name,spec]of Object.entries(specs)){
        let selected=model;
        if(spec.mutation==='duplicate')selected={...model,tools:[...model.tools,{...target,id:target.id+'-duplicate',row:999}]};
        if(spec.mutation==='disposed')selected={...model,tools:model.tools.map(tool=>tool.id===target.id?{...tool,due:'Disposed'}:tool)};
        const passes=spec.passes||[{text:spec.text,confidence:Object.hasOwn(spec,'confidence')?spec.confidence:95,source:'original'}];
        const parsed=ToolOCR.parse(passes.map(pass=>pass.text).join('\n'),selected.tools,{passes});
        const direct=ToolSMAX.assess(selected,parsed),assessment=ToolAutomation.assess(selected,parsed);
        const replay=ToolOCR.parse(parsed.text,selected.tools,{confidence:parsed.confidence,passes:parsed.passes});
        output[name]={expected:spec,assessment,direct,documentType:parsed.documentType,
            evidence:parsed.smaxEvidence,replay:ToolAutomation.assess(selected,replay),targetId:target.id,
            targetName:target.name,targetLast:target.last,targetDue:target.due};
    }
    const calibration='Serial: '+target.serial+'\n교정일자 2026.06.22\n차기교정예정일자 2026.12.22';
    const label=ToolOCR.parse(calibration,model.tools,{passes:[{text:calibration,confidence:95,source:'original'}]});
    output.calibration_label_stays_calibration={assessment:ToolAutomation.assess(model,label),documentType:label.documentType};
    const physical='Serial Number '+target.serial+'\nDate of Calibration 2026.06.22\nCalibration Due Date 2026.12.22';
    const english=ToolOCR.parse(physical,model.tools,{passes:[{text:physical,confidence:95,source:'original'}]});
    output.calibration_english_label_stays_calibration={assessment:ToolAutomation.assess(model,english),documentType:english.documentType};
    return output;
}"""


def policies(page):
    outcomes = page.evaluate(POLICIES)
    for name, outcome in outcomes.items():
        assessment = outcome["assessment"]
        if name.startswith("calibration_"):
            assert outcome["documentType"] == "calibration", outcome
            assert assessment["eligible"] and assessment["last"] == "2026-06-22" and assessment["due"] == DUE, outcome
            continue
        expected = outcome["expected"]
        assert assessment["eligible"] is expected["eligible"], (name, outcome)
        assert outcome["direct"]["eligible"] is expected["eligible"], (name, outcome)
        assert outcome["documentType"] == expected.get("documentType", "smax"), (name, outcome)
        assert outcome["replay"]["eligible"] is assessment["eligible"], (name, outcome)
        if expected.get("noProbeEvidence"):
            assert not outcome["evidence"]["serialReadings"] and not outcome["evidence"]["dueReadings"], (name, outcome)
        if expected["eligible"]:
            assert assessment["toolId"] == outcome["targetId"], (name, outcome)
            assert assessment["name"] == outcome["targetName"], (name, outcome)
            assert assessment["serial"] == SERIAL and assessment["photoKind"] == "smax", (name, outcome)
            assert assessment["last"] == outcome["targetLast"], (name, outcome)
            assert assessment["updateDue"] is expected["updateDue"], (name, outcome)
            assert assessment["due"] == expected.get("due", outcome["targetDue"]), (name, outcome)
            if "owner" in expected:
                assert assessment["ownerEnglish"] == expected["owner"], (name, outcome)
            for field in ["toolId", "name", "serial", "last", "due", "photoKind", "updateDue"]:
                assert outcome["replay"][field] == assessment[field], (name, field, outcome)
        else:
            assert assessment["reasons"], (name, outcome)
    return {"cases": len(outcomes), "eligible": sum(result["assessment"]["eligible"] for result in outcomes.values()),
            "english_owner_resolved_by_exact_registered_serial": True,
            "only_calibration_due_role_can_update_next": True,
            "unknown_and_low_due_are_photo_only": True,
            "partial_unknown_duplicate_disposed_and_conflicting_serials_blocked": True,
            "classification_and_pass_replay_preserved": True}


def make_picture(path):
    image = Image.new("RGB", (1200, 900), "white")
    draw = ImageDraw.Draw(image)
    for point, text in [((40, 35), "SMAX Tool Detail"), ((60, 200), "Serial/Asset Number"),
                        ((600, 200), SERIAL), ((60, 350), "Tool Owner"), ((600, 350), ENGLISH_OWNER),
                        ((60, 500), "Calibration Due Date"), ((600, 500), "12/22/2026")]:
        draw.text(point, text, fill="black", font_size=30)
    image.save(path)
    return path


WORKER = r"""({scenario,serial})=>{
    const map=new WeakMap(),native=CanvasRenderingContext2D.prototype.drawImage;
    window.smaxCalls=[];window.smaxParameters=[];
    CanvasRenderingContext2D.prototype.drawImage=function(source,...args){
        const returned=native.call(this,source,...args);
        const from=map.get(source)||(source.width===1200&&source.height===900?{x0:0,y0:0,x1:1200,y1:900}:null);
        if(from){
            let sx=0,sy=0,sw=source.width,sh=source.height,dx=0,dy=0,dw=sw,dh=sh;
            if(args.length===2)[dx,dy]=args;
            else if(args.length===4)[dx,dy,dw,dh]=args;
            else if(args.length===8)[sx,sy,sw,sh,dx,dy,dw,dh]=args;
            const ox=x=>from.x0+(sx+(x-dx)*sw/dw)/source.width*(from.x1-from.x0);
            const oy=y=>from.y0+(sy+(y-dy)*sh/dh)/source.height*(from.y1-from.y0);
            map.set(this.canvas,{x0:ox(0),y0:oy(0),x1:ox(this.canvas.width),y1:oy(this.canvas.height)});
        }
        return returned;
    };
    const fields=[
        ['Serial/Asset Number',60,200,400,240], [serial,600,200,920,240],
        ['Tool Owner',60,350,300,390], ['Ken Kang',600,350,900,390],
        ['Calibration Due Date',60,500,410,540], ['12/22/2026',600,500,920,540]];
    const response=(text,words=[],confidence=95)=>({data:{text,words,confidence}});
    Tesseract.createWorker=async()=>({
        setParameters:async params=>smaxParameters.push({...params}),
        recognize:async canvas=>{
            const roi=map.get(canvas);
            smaxCalls.push({roi,width:canvas.width,height:canvas.height,parameters:smaxParameters.at(-1)});
            if(!roi)throw new Error('Missing provenance for the actual OCR canvas');
            const full=roi.x0<=0&&roi.y0<=0&&roi.x1>=1200&&roi.y1>=900;
            if(full){
                if(scenario==='unknown')return response('SMAX\nTool Detail\nTool Code '+serial+'\nTool Owner Ken Kang\nCalibration Due Date 12/22/2026');
                return response('SMAX\nTool Detail\nSerial/Asset Number '+serial+'\nTool Owner Ken Kang\nCalibration Due Date 12/22/2026',
                    fields.map(([text,x0,y0,x1,y1])=>({text,confidence:95,bbox:{
                        x0:x0/1200*canvas.width,y0:y0/900*canvas.height,
                        x1:x1/1200*canvas.width,y1:y1/900*canvas.height}})));
            }
            const covers=(x,y)=>roi.x0<=x&&roi.x1>=x&&roi.y0<=y&&roi.y1>=y;
            if(covers(740,220))return response(serial);
            if(covers(740,370))return response('Ken Kang');
            if(covers(740,520))return response('12/22/2026');
            return response('',[],0);
        },terminate:async()=>{}
    });
}"""


def recognition_flow(page, url, workbook, photo, scenario):
    page.goto(url)
    page.wait_for_function("window.ToolApp&&!ToolApp.state.busy")
    page.locator("#workbook").set_input_files(workbook)
    page.wait_for_function("ToolApp.state.model&&!ToolApp.state.busy")
    page.evaluate(WORKER, {"scenario": scenario, "serial": SERIAL})
    files = [photo]
    if scenario == "duplicate":
        duplicate = photo.with_name("second-smax.png")
        duplicate.write_bytes(photo.read_bytes())
        files.append(duplicate)
    page.locator("#photos").set_input_files(files)
    page.wait_for_function("ToolApp.state.photos.length>0&&ToolApp.state.photos.every(p=>p.result)&&!ToolApp.state.busy&&!ToolApp.state.preparing", timeout=30000)
    result = page.evaluate("""() => ({photos:ToolApp.state.photos.map(p=>({
        documentType:p.result.documentType,passes:p.result.passes,assessment:p.assessment,
        approval:p.approval,review:p.review,draft:p.draft})),prepared:!!ToolApp.state.preparedResult,
        calls:smaxCalls,target:ToolApp.state.model.tools.find(t=>t.serial==='0716805829')})""")
    assert result["calls"], result
    for photo_result in result["photos"]:
        assert photo_result["documentType"] == ("unknown-smax" if scenario == "unknown" else "smax"), result
        assert not any(p["source"].startswith(("date-", "identity-")) for p in photo_result["passes"]), result
    if scenario == "unknown":
        assert not result["prepared"], result
        assert all(not p["assessment"]["eligible"] and not p["approval"] and not p["review"] for p in result["photos"]), result
        assert result["photos"][0]["draft"]["photoKind"] == "smax", result
        return {"unknown_smax_never_enters_calibration_retries_or_label_photo_output": True}
    if scenario == "duplicate":
        assert not result["prepared"], result
        assert all(not p["assessment"]["eligible"] and not p["approval"] for p in result["photos"]), result
        return {"two_smax_images_for_one_slot_are_both_held": True}
    photo_result = result["photos"][0]
    assert result["prepared"] and photo_result["approval"] == "automatic", result
    assert photo_result["review"]["photoKind"] == "smax" and photo_result["review"]["updateDue"], result
    assert photo_result["review"]["last"] == result["target"]["last"], result
    assert photo_result["review"]["due"] == DUE, result
    field_passes = [p for p in photo_result["passes"] if p["source"].startswith("smax-field-")]
    assert field_passes, result
    assert all(p["smaxLabelVerified"] and p["smaxLabelBBox"] and p["smaxLabelText"] for p in field_passes), result
    assert any(p["smaxRole"] == "due" for p in field_passes), result
    assert any(call["roi"]["y0"] <= 520 <= call["roi"]["y1"] and
               call["roi"]["x0"] <= 740 <= call["roi"]["x1"] for call in result["calls"][1:]), result
    return {"actual_recognizer_uses_verified_smax_field_crops": True,
            "calibration_retries_are_not_used": True, "default_automatic_output_targets_smax_photo_only": True,
            "cal_last_kept_and_cal_next_from_due_role": True, "worker_calls": len(result["calls"])}


def main():
    args = arguments(__doc__)
    original_hash = sha256(args.workbook.read_bytes()).hexdigest()
    errors, remote = [], []
    with tempfile.TemporaryDirectory(prefix="tool-smax-policy-") as directory, serve_program(args.program) as url, sync_playwright() as playwright:
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
        context = browser.new_context(locale="ko-KR")
        context.on("request", lambda request: remote.append(request.url)
            if urlsplit(request.url).scheme in {"http", "https"} and urlsplit(request.url).netloc != urlsplit(url).netloc else None)
        page = context.new_page()
        page.on("pageerror", lambda error: errors.append(str(error)))
        page.on("dialog", lambda dialog: dialog.accept())
        page.goto(url)
        page.wait_for_function("window.ToolApp&&!ToolApp.state.busy")
        page.locator("#workbook").set_input_files(args.workbook)
        page.wait_for_function("ToolApp.state.model&&!ToolApp.state.busy")
        result = {"parser_and_assessor": policies(page)}
        picture = make_picture(Path(directory) / "synthetic-smax.png")
        for scenario in ["smax", "unknown", "duplicate"]:
            result["recognition_" + scenario] = recognition_flow(page, url, args.workbook, picture, scenario)
        assert not errors, errors
        assert not remote, remote
        context.close()
        browser.close()
    assert sha256(args.workbook.read_bytes()).hexdigest() == original_hash
    result.update({"original_company_workbook_unchanged": True, "page_errors": errors, "external_network_requests": remote,
        "real_photo_accuracy": "not measured: controlled OCR text and synthetic worker responses"})
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
