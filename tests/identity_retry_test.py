"""Recognition retries must recover an upper owner/serial tag after dates work.

ToolOCR.recognize, canvas preparation, parser, assessment and app save guards
run unchanged. Tesseract's worker is a response test double, so these tests
verify retry decisions and write safety rather than handwriting accuracy.
The company workbook is read only; connected writes use real browser OPFS.
"""
from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path
import tempfile
from urllib.parse import urlsplit

from PIL import Image, ImageDraw
from playwright.sync_api import sync_playwright

from browser_test import REPO, arguments, serve_program
from connected_file_test import PICKER_SCRIPT, export_shared, ready, seed_file
from torque_dates_test import SERIAL, LAST, DUE, LABELS, verify_torque_workbook


def fixture_photo(path, two_tags=False):
    # A visible upper tag makes the input realistic for region proposals. Text
    # recognition remains controlled by worker responses, as stated above.
    image = Image.new("RGB", (1200, 1600), (230, 215, 180))
    draw = ImageDraw.Draw(image)
    draw.rectangle((660, 100, 1110, 340), fill=(180, 215, 235), outline=(30, 45, 55), width=3)
    draw.text((700, 260), SERIAL, fill="black", font_size=35)
    if two_tags:
        draw.rectangle((70, 100, 520, 340), fill=(180, 215, 235), outline=(30, 45, 55), width=3)
        draw.text((105, 260), "SECOND TAG", fill="black", font_size=32)
    draw.rectangle((300, 1120, 970, 1450), fill="white", outline="black", width=3)
    draw.text((340, 1190), LAST, fill="black", font_size=38)
    draw.text((340, 1340), DUE, fill="black", font_size=38)
    image.save(path)
    return path


def identity_policies(page, url, workbook):
    address = urlsplit(url)
    page.route("**/identity-policy.html", lambda route: route.fulfill(status=200,
        content_type="text/html", body='<!doctype html><input id="fixture" type="file">'))
    page.goto(f"{address.scheme}://{address.netloc}/identity-policy.html")
    for source in ["vendor/jszip.min.js", "workbook.js", "ocr.js", "automation.js"]:
        page.add_script_tag(path=REPO / "web" / source)
    page.locator("#fixture").set_input_files(workbook)
    outcomes = page.evaluate("""async ({serial,dates}) => {
        const model=await ToolWorkbook.load(await document.querySelector('#fixture').files[0].arrayBuffer());
        const target=model.tools.find(tool=>tool.serial===serial);
        const other=model.tools.find(tool=>tool.serial!==serial&&tool.name===target.name);
        if(!other)throw new Error('Fixture requires two tools for the owner');
        const cases={
            upper_tag_exact_leading_zero:[serial],
            upper_tag_dropped_leading_zero:['716805829'],
            upper_tag_letter_o_instead_of_zero:['O716805829'],
            matching_tag_and_unknown_explicit_serial:[serial+'\\nSerial: 999999999'],
            two_known_tool_tags:[serial,other.serial],
            two_known_tags_in_reverse_order:[other.serial,serial],
            dates_without_identity:[],
            single_probe_numeric_noise_does_not_contradict_known_tag:[
                {text:serial,source:'identity-probe-panel-0'},
                {text:'999999999',source:'identity-probe-line-1'}],
            repeated_probe_unknown_number_remains_a_conflict:[
                {text:serial,source:'identity-probe-panel-0'},
                {text:'999999999',source:'identity-probe-line-1'},
                {text:'999999999',source:'identity-probe-line-2'}],
            single_probe_explicit_unknown_serial_remains_a_conflict:[
                {text:serial,source:'identity-probe-panel-0'},
                {text:'Serial: 999999999',source:'identity-probe-line-1'}],
            name_cycle_without_exact_serial_does_not_ignore_probe_number:[
                {text:target.name,source:'identity-probe-panel-0'},
                {text:'999999999',source:'identity-probe-line-1'}],
        };
        const outcomes={};
        for(const [name,tags] of Object.entries(cases)){
            const passes=[{text:dates,confidence:95,source:'original'},
                ...tags.map((tag,index)=>typeof tag==='string'?{text:tag,confidence:95,source:'identity-crop-upper-tag-'+index}:
                    {...tag,confidence:95})];
            const parsed=ToolOCR.parse(dates,model.tools,{confidence:95,passes});
            outcomes[name]={assessment:ToolAutomation.assess(model,parsed),
                candidates:parsed.candidateIds,serials:parsed.identityEvidence.serialMatches,
                targetId:target.id};
        }
        // Reported text fragments have no supplied confidence or positions.
        // Preserve that absence; only the added crop responses below are
        // explicitly synthetic, independent, high-confidence test evidence.
        const reported='071680582qg\\n한거으\\n2026.06.22\\n2026.12.22\\n0716805808';
        const raw=ToolOCR.parse(reported,model.tools);
        outcomes.reported_fragments_are_not_forced_to_the_nearest_serial={
            assessment:ToolAutomation.assess(model,raw),candidates:raw.candidateIds,
            serials:raw.identityEvidence.serialMatches,confidence:raw.confidence,targetId:target.id};
        const recovered=ToolOCR.parse(reported,model.tools,{passes:[
            {text:reported,source:'original',confidence:null},
            {text:dates,source:'date-crop',confidence:95},
            {text:serial,source:'identity-crop-upper-tag',confidence:95}]});
        outcomes.independent_controlled_crops_can_recover_reported_fragments={
            assessment:ToolAutomation.assess(model,recovered),candidates:recovered.candidateIds,
            serials:recovered.identityEvidence.serialMatches,targetId:target.id};
        const geometryCases={
            same_row_serial_fragments:{tokens:['07168','05829'],eligible:true},
            serial_fragments_on_separate_rows:{tokens:['07168','05829'],rows:[0,1],eligible:false},
            serial_fragment_minimum_confidence:{tokens:['07168','05829'],confidences:[95,25],eligible:false},
            serial_fragment_suffix_digit:{tokens:['07168','05829','7'],eligible:false},
            full_serial_token_suffix_digit:{tokens:[serial,'7'],eligible:false},
            serial_fragment_suffix_letters:{tokens:['07168','05829','qg'],eligible:false},
            owner_and_full_serial_suffix_digit:{tokens:[target.name,serial,'7'],rows:[0,1,1],eligible:false},
            owner_and_full_serial_suffix_letters:{tokens:[target.name,serial,'qg'],rows:[0,1,1],eligible:false},
            owner_and_unknown_serial_suffix_digit:{tokens:[target.name,'999999999','7'],rows:[0,1,1],eligible:false},
            serial_label_can_precede_exact_value:{tokens:['S/N',serial],eligible:true},
            same_row_name_fragments:{tokens:Array.from(target.name),eligible:true},
            name_fragments_on_separate_rows:{tokens:Array.from(target.name),rows:Array.from(target.name,(_,i)=>i),eligible:false},
            full_name_token_honorific_suffix:{tokens:[target.name,'님'],eligible:false},
            name_fragments_honorific_suffix:{tokens:[...Array.from(target.name),'님'],eligible:false},
            owner_label_is_not_a_name_prefix:{tokens:['성명',target.name],eligible:true},
            explicit_serial_all_fragments_not_known_prefix:{text:'Serial: 07168 05829 99',tokens:[],eligible:false},
        };
        for(const [name,spec] of Object.entries(geometryCases)){
            let x=20;
            const words=spec.tokens.map((text,index)=>{
                const width=Math.max(18,Array.from(text).length*16),y=40+(spec.rows?.[index]||0)*80;
                const word={text,confidence:spec.confidences?.[index]??95,bbox:{x0:x,y0:y,x1:x+width,y1:y+30}};
                x+=width+5;return word;
            });
            const text=spec.text??spec.tokens.join('\\n');
            const parsed=ToolOCR.parse(dates,model.tools,{confidence:95,passes:[
                {text:dates,confidence:95,source:'original'},
                {text,confidence:95,words,source:'identity-probe-line-0'}]});
            outcomes['geometry_'+name]={assessment:ToolAutomation.assess(model,parsed),
                candidates:parsed.candidateIds,serials:parsed.identityEvidence.serialMatches,
                targetId:target.id,expectedEligible:spec.eligible};
        }
        return outcomes;
    }""", {"serial": SERIAL, "dates": LABELS})
    for name, outcome in outcomes.items():
        assessment = outcome["assessment"]
        expected = outcome.get("expectedEligible", name in {"upper_tag_exact_leading_zero",
            "single_probe_numeric_noise_does_not_contradict_known_tag",
            "independent_controlled_crops_can_recover_reported_fragments"})
        if expected:
            assert assessment["eligible"], (name, outcome)
            assert assessment["toolId"] == outcome["targetId"]
            assert assessment["serial"] == SERIAL
            assert assessment["last"] == LAST and assessment["due"] == DUE
        else:
            assert not assessment["eligible"], (name, outcome)
            assert assessment["reasons"], (name, outcome)
    assert outcomes["two_known_tool_tags"]["assessment"]["toolId"] == ""
    assert outcomes["two_known_tags_in_reverse_order"]["assessment"]["toolId"] == ""
    assert outcomes["reported_fragments_are_not_forced_to_the_nearest_serial"]["confidence"] is None
    assert outcomes["reported_fragments_are_not_forced_to_the_nearest_serial"]["serials"] == []
    return {"policy_cases": len(outcomes), "exact_upper_tag_and_single_exploratory_number": "accepted",
            "dropped_zero_o_unknown_conflict_two_tag_candidates_and_no_identity": "pending",
            "repeated_or_explicit_unknown_probe_serials": "pending",
            "reported_fragments_without_confidence": "pending; no nearest-serial substitution",
            "independent_controlled_crop_recovery": "accepted",
            "same_row_fragments_and_minimum_confidence": "verified",
            "different_rows_suffix_digits_letters_and_honorifics": "pending",
            "name_cycle_without_exact_serial_cannot_ignore_probe_number": "pending",
            "candidate_order_cannot_choose_a_different_tool": True}


def install_worker(page, scenario):
    page.evaluate("""({scenario,serial,dates}) => {
        const target=ToolApp.state.model.tools.find(tool=>tool.serial===serial);
        const other=ToolApp.state.model.tools.find(tool=>tool.serial!==serial&&tool.name===target.name);
        window.identityWorkerCalls=[];window.identityWorkerParameters=[];window.identityCoarsePanels=0;
        let currentTagSerial=serial;
        Tesseract.createWorker=async()=>({
            setParameters:async parameters=>{identityWorkerParameters.push({...parameters});},
            recognize:async canvas=>{
                const index=identityWorkerCalls.length;
                const parameters=identityWorkerParameters.at(-1)||{};
                identityWorkerCalls.push({width:canvas.width,height:canvas.height,parameters});
                if(scenario==='incomplete_tag_scan'&&index===5)throw new Error('TEST_SECOND_TAG_RETRY_FAILURE');
                let text=index<2?dates:'';
                if(index>=4){
                    if(scenario==='upper_serial')text=serial;
                    else if(scenario==='upper_name_then_serial')text=index===4?target.name:serial;
                    else if(scenario==='upper_line_serial')text=String(parameters.tessedit_pageseg_mode)==='6'?target.name:serial;
                    else if(scenario==='incomplete_tag_scan')text=serial;
                    else if(scenario==='two_tag_candidates'){
                        if(String(parameters.tessedit_pageseg_mode)==='6'){
                            identityCoarsePanels++;
                            currentTagSerial=identityCoarsePanels===1?serial:other.serial;
                        }
                        text=currentTagSerial;
                    }
                }
                if(index>=2&&scenario==='unknown_serial_conflict')text=serial+'\\nSerial: 999999999';
                if(index>=2&&scenario==='known_serial_conflict')text=serial+'\\n'+other.serial;
                return {data:{text,confidence:95,words:[]}};
            },
            terminate:()=>{}
        });
    }""", {"scenario": scenario, "serial": SERIAL, "dates": LABELS})


def recognize_retry(page, url, args, photo, scenario):
    page.goto(url)
    page.wait_for_function("window.ToolApp&&!ToolApp.state.busy")
    page.locator("#workbook").set_input_files(args.workbook)
    page.wait_for_function("ToolApp.state.model&&!ToolApp.state.busy")
    install_worker(page, scenario)
    page.evaluate("() => {const input=document.createElement('input');input.type='file';input.id='retry-input';document.body.append(input);}")
    page.locator("#retry-input").set_input_files(photo)
    result = page.evaluate("""async()=>{
        const model=ToolApp.state.model;
        const parsed=await ToolOCR.recognize(document.querySelector('#retry-input').files[0],model.tools);
        const replay=ToolOCR.parse(parsed.text,model.tools,{confidence:parsed.confidence,passes:parsed.passes});
        return {assessment:ToolAutomation.assess(model,parsed),failed:parsed.failed,error:parsed.error,
            serials:parsed.identityEvidence.serialMatches,calls:identityWorkerCalls,coarsePanels:identityCoarsePanels,
            passes:parsed.passes.map(pass=>({source:pass.source,text:pass.text})),
            incomplete:parsed.identityRecoveryIncomplete,replayIncomplete:replay.identityRecoveryIncomplete,
            last:parsed.last,due:parsed.due,
            targetId:model.tools.find(tool=>tool.serial==='0716805829').id};
    }""")
    assert not result["failed"], result
    assert len(result["calls"]) > 4, result
    if scenario == "incomplete_tag_scan":
        assert result["incomplete"] and result["replayIncomplete"], result
        assert result["last"] == LAST and result["due"] == DUE, result
        assert any(proof["serial"] == SERIAL and proof["confidence"] >= 70 for proof in result["serials"]), result
        assert not result["assessment"]["eligible"], result
        return {"partial_scan_does_not_approve_first_known_tag": True, "dates_and_exact_serial_are_preserved": True,
                "incomplete_marker_survives_reparse": True}
    if scenario == "two_tag_candidates":
        assert result["coarsePanels"] >= 2, result
        assert not result["assessment"]["eligible"], result
        assert len({proof["toolId"] for proof in result["serials"]}) >= 2, result
        return {"later_tag_is_still_read_after_first_known_serial": True,
                "multiple_known_tags_remain_ambiguous": True, "coarse_panel_reads": result["coarsePanels"]}
    assert result["assessment"]["eligible"], result
    assert result["assessment"]["toolId"] == result["targetId"], result
    assert result["assessment"]["serial"] == SERIAL, result
    assert result["assessment"]["last"] == LAST and result["assessment"]["due"] == DUE, result
    assert any(proof["serial"] == SERIAL and proof["confidence"] >= 70 for proof in result["serials"]), result
    if scenario == "upper_line_serial":
        assert any(p["source"].startswith("identity-probe-line-") for p in result["passes"]), result
    return {"worker_calls": len(result["calls"]), "pass_sources": [p["source"] for p in result["passes"]],
            "generic_identity_failure_does_not_stop_upper_tag_retries": True,
            "recovered_serial": result["assessment"]["serial"]}


def connected_gate(context, page, url, args, photo, temporary, scenario):
    context.add_init_script("(" + PICKER_SCRIPT + ")()")
    page.goto(url)
    page.wait_for_function("window.ToolApp&&!ToolApp.state.busy")
    baseline = seed_file(page, args.workbook)
    page.locator("#connect-onedrive").click()
    page.wait_for_function("ToolApp.state.source.connected&&ToolApp.state.model")
    ready(page)
    install_worker(page, scenario)
    page.locator("#photos").set_input_files(photo)
    page.wait_for_function("ToolApp.state.photos.length===1&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=30000)
    result = page.evaluate("""async()=>{
        const p=ToolApp.state.photos[0],handle=await(await navigator.storage.getDirectory()).getFileHandle('company-shared.xlsx');
        return {review:!!p.review,saved:!!p.saved,approval:p.approval,assessment:p.assessment,
            generated:!!ToolApp.state.preparedResult,saves:ToolApp.state.source.saveCount,
            sourceHash:await testHash(await(await handle.getFile()).arrayBuffer())};
    }""")
    if scenario == "upper_serial":
        assert result["review"] and result["saved"] and result["generated"], result
        assert result["approval"] == "automatic" and result["saves"] == 1, result
        output = temporary / "upper-tag-connected.xlsx"
        export_shared(page, output)
        return {"automatic_connected_write": verify_torque_workbook(args.workbook, output)}
    assert not result["review"] and not result["saved"] and not result["generated"], result
    assert result["saves"] == 0 and result["sourceHash"] == baseline, result
    assert not page.locator("#automatic-result-download").is_visible()
    assert page.locator("#export").is_disabled()
    return {"conflict_blocks_review_generation_and_connected_write": True, "connected_file_byte_identical": True}


def main():
    args = arguments(__doc__)
    before = sha256(args.workbook.read_bytes()).hexdigest()
    errors, remote = [], []
    with tempfile.TemporaryDirectory(prefix="tool-identity-retry-") as directory, serve_program(args.program) as url, sync_playwright() as playwright:
        temporary = Path(directory)
        photo = fixture_photo(temporary / "upper-tag-fixture.png")
        two_tags = fixture_photo(temporary / "two-upper-tags-fixture.png", two_tags=True)
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
        def context_and_page():
            context = browser.new_context(accept_downloads=True, locale="ko-KR")
            context.on("request", lambda request: remote.append(request.url)
                if urlsplit(request.url).scheme in {"http", "https"} and urlsplit(request.url).netloc != urlsplit(url).netloc else None)
            page = context.new_page()
            page.on("dialog", lambda dialog: dialog.accept())
            page.on("pageerror", lambda error: errors.append(str(error)))
            return context, page
        context, page = context_and_page()
        result = {"identity_policies": identity_policies(page, url, args.workbook)}
        context.close()
        for scenario in ["upper_serial", "upper_name_then_serial", "upper_line_serial", "two_tag_candidates", "incomplete_tag_scan"]:
            context, page = context_and_page()
            result[scenario] = recognize_retry(page, url, args, two_tags if scenario == "two_tag_candidates" else photo, scenario)
            context.close()
        for scenario in ["upper_serial", "unknown_serial_conflict", "known_serial_conflict", "two_tag_candidates", "incomplete_tag_scan"]:
            context, page = context_and_page()
            result["connected_" + scenario] = connected_gate(context, page, url, args,
                two_tags if scenario == "two_tag_candidates" else photo, temporary, scenario)
            context.close()
        assert not errors, errors
        assert not remote, remote
        browser.close()
    assert sha256(args.workbook.read_bytes()).hexdigest() == before
    result.update({"original_company_workbook_unchanged": True, "page_errors": errors, "external_network_requests": remote,
                   "recognition_accuracy": "not measured: worker responses are controlled"})
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
