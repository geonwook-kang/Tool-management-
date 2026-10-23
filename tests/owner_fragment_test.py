"""Owner and four observed digits: controlled OCR policy and real OPFS writes.

The parser, assessment, approval, XLSX generation and OPFS storage are real.
Recognition responses are deterministic test doubles, not a measurement of
the user's photo OCR accuracy. The reported text lacks full dates and OCR
confidence; tests preserve those absences and never replace 0713 with 0716.
The attached company workbook is read only; generated files stay temporary.
"""
from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path
import tempfile
from urllib.parse import urlsplit
from zipfile import ZipFile

from PIL import Image, ImageDraw
from playwright.sync_api import sync_playwright

from browser_test import REPO, arguments, serve_program
from connected_file_test import PICKER_SCRIPT, export_shared, ready, seed_file
from torque_dates_test import SERIAL, LAST, DUE, LABELS, verify_torque_workbook


def policy_checks(page, url, workbook, source_modules=False):
    if source_modules:
        address = urlsplit(url)
        page.route("**/owner-fragment-policy.html", lambda route: route.fulfill(
            status=200, content_type="text/html", body="<!doctype html>"))
        page.goto(f"{address.scheme}://{address.netloc}/owner-fragment-policy.html")
        for source in ["vendor/jszip.min.js", "workbook.js", "ocr.js", "automation.js"]:
            page.add_script_tag(path=REPO / "web" / source)
    else:
        page.goto(url)
        page.wait_for_function("window.ToolOCR&&window.ToolAutomation&&window.ToolWorkbook")
    page.evaluate("""()=>{const input=document.createElement('input');input.type='file';
        input.id='owner-fragment-fixture';input.hidden=true;document.body.append(input);} """)
    page.locator("#owner-fragment-fixture").set_input_files(workbook)
    outcomes = page.evaluate("""async({serial,dates})=>{
        const model=await ToolWorkbook.load(await document.querySelector('#owner-fragment-fixture').files[0].arrayBuffer());
        const target=model.tools.find(tool=>tool.serial===serial);
        const other=model.tools.find(tool=>tool.name!==target.name&&tool.serial&&tool.due!=='폐기');
        const fragment=serial.slice(0,4),cases={};
        const sameCycleTools=model.tools.map(tool=>tool.name===target.name?
            {...tool,months:6,interval:'6',cycle:'MONTHS'}:tool);
        const pass=(text,source,confidence=95)=>({text,source,confidence});
        const repeated=(value=fragment,label=true)=>[
            pass((label?'Serial: ':'')+value,'identity-crop-upper-left'),
            pass((label?'Serial: ':'')+value,'identity-crop-upper-right')];
        function run(name,{reads=repeated(),owner=target.name,dateText=dates,tools=model.tools,
                          mutate=null,eligible=false,method='owner-partial-serial',actual=fragment}={}){
            const passes=[pass(dateText,'original'),...(owner?[pass(owner,'identity-crop-owner')]:[]),...reads];
            const parsed=ToolOCR.parse(dateText,tools,{confidence:95,passes});
            if(mutate)mutate(parsed);
            const before=JSON.stringify(parsed.identityEvidence);
            cases[name]={assessment:ToolAutomation.assess({...model,tools},parsed),
                expected:{eligible,method,actual},evidence:parsed.identityEvidence,
                parserSerial:parsed.serial,evidenceUnchanged:before===JSON.stringify(parsed.identityEvidence)};
        }
        run('exact_owner_and_repeated_explicit_four_digits',{eligible:true});
        run('standalone_four_digits_are_actual_observations',{reads:repeated(fragment,false),eligible:true});
        run('spaces_inside_one_serial_line_preserve_actual_digits',{
            reads:repeated('071 6',false),eligible:true});
        run('spaced_explicit_serial_preserves_actual_digits',{
            reads:repeated('07 1 6'),eligible:true});
        run('spaced_wrong_0713_is_not_changed_to_0716',{
            reads:repeated('071 3'),actual:'0713'});
        run('newlines_are_not_joined_to_invent_four_digits',{
            reads:repeated('071\\n6')});
        run('split_word_fragment_uses_its_lowest_observed_confidence',{
            tools:sameCycleTools,reads:repeated('071 6').map(proof=>({...proof,words:[
                {text:'071',confidence:95,bbox:{x0:20,y0:20,x1:80,y1:50}},
                {text:'6',confidence:25,bbox:{x0:85,y0:20,x1:105,y1:50}}]}))});
        run('contiguous_four_digit_suffix',{reads:repeated(serial.slice(-4)),actual:serial.slice(-4),eligible:true});
        run('contiguous_four_digit_middle',{reads:repeated(serial.slice(3,7)),actual:serial.slice(3,7),eligible:true});
        run('four_digit_fragment_at_each_confidence_threshold',{
            reads:repeated().map(proof=>({...proof,confidence:70})),eligible:true});
        run('explicit_wrong_0713_is_not_changed_to_0716',{reads:repeated('0713'),actual:'0713'});
        run('wrong_standalone_digits_preserve_independent_name_cycle',{
            reads:repeated('0713',false),actual:'0713',eligible:true,method:'name-cycle'});
        run('one_pass_cannot_resolve_explicit_fragment',{reads:repeated().slice(0,1)});
        run('same_source_is_not_independent',{reads:repeated().map(proof=>({...proof,source:'identity-crop-upper-left'}))});
        run('one_read_below_confidence_threshold',{
            reads:repeated().map((proof,index)=>({...proof,confidence:index?25:95}))});
        run('weak_fragments_cannot_distinguish_same_owner_same_cycle_tools',{
            tools:sameCycleTools,reads:repeated().map(proof=>({...proof,confidence:69}))});
        run('absent_fragment_confidence_cannot_distinguish_same_cycle_tools',{
            tools:sameCycleTools,mutate:parsed=>{
                for(const proof of [...parsed.identityEvidence.observedSerialFragments||[],
                                    ...parsed.identityEvidence.observedSerialEvidence||[]]){
                    proof.confidence=null;for(const reading of proof.sourceReadings||[])reading.confidence=null;
                }
            }});
        run('legacy_aggregate_confidence_is_not_two_high_reads',{mutate:parsed=>{
            for(const proof of parsed.identityEvidence.observedSerialFragments||[])delete proof.sourceReadings;
        }});
        run('listed_sources_must_contain_actual_high_readings',{mutate:parsed=>{
            for(const proof of parsed.identityEvidence.observedSerialFragments||[])proof.passSources=['other-source'];
        }});
        run('source_confidences_cannot_be_omitted',{mutate:parsed=>{
            for(const proof of parsed.identityEvidence.observedSerialFragments||[])
                for(const reading of proof.sourceReadings||[])reading.confidence=null;
        }});
        run('fragment_array_cannot_be_inferred_from_candidates',{mutate:parsed=>{
            parsed.identityEvidence.observedSerialFragments=[];parsed.candidateIds=[target.id];
        }});
        run('owner_must_be_an_exact_canonical_match',{owner:'Owner: UNKNOWN TEST OWNER'});
        run('another_known_owner_cannot_select_this_fragment',{owner:other.name});
        run('owner_confidence_must_be_high',{mutate:parsed=>{
            for(const proof of parsed.identityEvidence.nameMatches)proof.confidence=25;
        }});
        run('learned_owner_alone_cannot_enable_four_digit_path',{mutate:parsed=>{
            for(const proof of parsed.identityEvidence.nameMatches)proof.learned=true;
        }});
        run('missing_last_day_is_not_filled_in',{dateText:'교정일자 2026.06\\n차기교정예정일자 2026.12.22'});
        run('missing_due_day_is_not_filled_in',{dateText:'교정일자 2026.06.22\\n차기교정예정일자 2026.12'});
        run('both_month_only_dates_are_incomplete',{dateText:'교정일자 2026.06\\n차기교정예정일자 2026.12'});
        run('one_full_date_does_not_enable_calculated_cycle',{dateText:'교정일자 2026.06.22'});
        run('complete_pair_must_match_current_excel_cycle',{
            dateText:'교정일자 2026.06.22\\n차기교정예정일자 2027.06.22'});
        run('low_date_confidence_cannot_enable_four_digits',{mutate:parsed=>{parsed.dateEvidence.confidence=25;}});
        run('conflicting_full_dates_remain_pending',{reads:[...repeated(),
            pass('교정일자 2026.06.22\\n차기교정예정일자 2026.12.21','date-crop-6')]});
        run('identity_recovery_must_be_complete',{mutate:parsed=>{parsed.identityRecoveryIncomplete=true;}});
        run('date_crop_numbers_do_not_become_identity',{reads:[
            pass('Serial: '+fragment,'date-crop-6'),pass('Serial: '+fragment,'date-crop-11')],owner:''});
        run('label_probe_numbers_do_not_become_identity',{reads:[
            pass('Serial: '+fragment,'label-probe-6'),pass('Serial: '+fragment,'label-probe-11')],owner:''});
        run('year_numbers_are_not_serial_fragments',{reads:repeated('2026',false),owner:'',actual:'2026'});
        run('certificate_context_numbers_are_not_serial_fragments',{
            reads:repeated('성적서번호\\n'+fragment,false),owner:''});
        run('unknown_full_serial_still_conflicts',{reads:[...repeated(),pass('Serial: 999999999','identity-crop-other')]});
        run('another_known_full_serial_still_conflicts',{
            reads:[...repeated(),pass('Serial: '+other.serial,'identity-crop-other')]});
        run('longer_unknown_number_is_not_cut_into_four_digits',{
            reads:repeated('99'+fragment+'99'),actual:'99'+fragment+'99'});
        const collision={...target,id:'synthetic-owner-fragment-collision',serial:'99'+fragment+'88',
            months:12,interval:'12',cycle:'MONTHS',due:'2027-12-22'};
        run('same_owner_collision_is_checked_before_cycle',{tools:[...model.tools,collision]});
        run('same_owner_disposed_collision_still_counts',{tools:[...model.tools,{...collision,due:'폐기'}]});
        run('two_different_credible_owner_fragments_cannot_be_filtered_by_cycle',{
            reads:[...repeated(fragment,false),...repeated('0713',false)],
            tools:[...model.tools,{...collision,serial:'99071388'}]});
        run('compatible_candidate_and_repeated_unexplained_fragment_remain_pending',{
            reads:[...repeated(fragment,false),...repeated('9898',false)]});
        run('another_owners_matching_fragment_does_not_change_owner_uniqueness',{
            tools:[...model.tools,{...collision,name:other.name}],eligible:true});
        const longSerial='99'+serial;
        run('four_digits_below_forty_percent_are_insufficient',{
            tools:model.tools.map(tool=>tool.id===target.id?{...tool,serial:longSerial}:tool)});
        run('existing_exact_full_serial_path_is_preserved',{
            reads:[pass('Serial: '+serial,'identity-crop-full')],eligible:true,method:'serial',actual:serial});
        run('existing_global_six_digit_partial_path_is_preserved',{
            reads:repeated(serial.slice(-6)),owner:'',eligible:true,method:'partial-serial',actual:serial.slice(-6)});
        run('existing_owner_five_digit_partial_path_is_preserved',{
            reads:repeated(serial.slice(-5)),eligible:true,method:'partial-serial',actual:serial.slice(-5)});
        const raw='0713\\n'+target.name.slice(0,-1)+'우) / '+target.name+'|\\n2026.06\\n2026.12';
        const parsed=ToolOCR.parse(raw,model.tools);
        cases.reported_text_has_no_full_dates_or_supplied_confidence={
            assessment:ToolAutomation.assess(model,parsed),expected:{eligible:false,method:'',actual:'0713'},
            evidence:parsed.identityEvidence,parserSerial:parsed.serial,evidenceUnchanged:true,
            confidence:parsed.confidence,dates:parsed.dateEvidence};
        return cases;
    }""", {"serial": SERIAL, "dates": LABELS})
    for name, outcome in outcomes.items():
        assessment, expected = outcome["assessment"], outcome["expected"]
        assert assessment["eligible"] is expected["eligible"], (name, outcome)
        assert outcome["evidenceUnchanged"], (name, outcome)
        if expected["eligible"]:
            assert assessment["serial"] == SERIAL, (name, outcome)
            assert assessment["matchMethod"] == expected["method"], (name, outcome)
            assert assessment["last"] == LAST and assessment["due"] == DUE, (name, outcome)
            if expected["method"] == "owner-partial-serial":
                assert expected["actual"] in [proof["fragment"] for proof in outcome["evidence"]["observedSerialFragments"]], (name, outcome)
                assert SERIAL not in outcome["evidence"]["observedSerials"], (name, outcome)
                assert outcome["evidence"]["serialMatches"] == [] and outcome["parserSerial"] == "", (name, outcome)
        else:
            assert assessment["reasons"], (name, outcome)
        if name in {"date_crop_numbers_do_not_become_identity", "label_probe_numbers_do_not_become_identity",
                    "year_numbers_are_not_serial_fragments", "certificate_context_numbers_are_not_serial_fragments"}:
            assert outcome["evidence"].get("observedSerialFragments", []) == [], (name, outcome)
        if name in {"longer_unknown_number_is_not_cut_into_four_digits", "newlines_are_not_joined_to_invent_four_digits"}:
            assert "0716" not in [proof["fragment"] for proof in outcome["evidence"]["observedSerialFragments"]], (name, outcome)
        if name == "split_word_fragment_uses_its_lowest_observed_confidence":
            assert all(proof["confidence"] == 25 for proof in outcome["evidence"]["observedSerialFragments"]), (name, outcome)
        if expected["actual"] == "0713":
            fragments = outcome["evidence"].get("observedSerialFragments", [])
            assert "0713" in [proof["fragment"] for proof in fragments], (name, outcome)
            assert "0716" not in [proof["fragment"] for proof in fragments], (name, outcome)
            assert SERIAL not in outcome["evidence"]["observedSerials"], (name, outcome)
    raw = outcomes["reported_text_has_no_full_dates_or_supplied_confidence"]
    assert raw["confidence"] is None and raw["dates"]["dates"] == [], raw
    assert all(proof["confidence"] is None for proof in raw["evidence"]["observedSerialFragments"]), raw
    return {"cases": len(outcomes),
            "accepted": [name for name, value in outcomes.items() if value["expected"]["eligible"]],
            "pending": [name for name, value in outcomes.items() if not value["expected"]["eligible"]],
            "reported_0713_is_never_rewritten_to_0716": True,
            "reported_partial_dates_and_absent_confidence_remain_pending": True}


def fixture_photo(path):
    image = Image.new("RGB", (600, 800), (230, 220, 195))
    draw = ImageDraw.Draw(image)
    draw.rectangle((80, 70, 520, 210), fill=(175, 220, 240), outline="black", width=4)
    draw.text((130, 110), "0716", fill="black", font_size=46)
    draw.text((100, 550), LAST, fill="black", font_size=34)
    draw.text((100, 650), DUE, fill="black", font_size=34)
    image.save(path)
    return path


def connected_check(browser, url, args, temporary, scenario):
    context = browser.new_context(accept_downloads=True, locale="ko-KR")
    context.add_init_script("(" + PICKER_SCRIPT + ")()")
    errors, remote = [], []
    context.on("request", lambda request: remote.append(request.url)
               if urlsplit(request.url).scheme in {"http", "https"}
               and urlsplit(request.url).netloc != urlsplit(url).netloc else None)
    page = context.new_page()
    page.on("pageerror", lambda error: errors.append(str(error)))
    page.on("dialog", lambda dialog: dialog.accept())
    page.goto(url)
    page.wait_for_function("window.ToolApp&&!ToolApp.state.busy")
    original_hash = seed_file(page, args.workbook)
    page.locator("#connect-onedrive").click()
    page.wait_for_function("ToolApp.state.source.connected&&!!ToolApp.state.model")
    ready(page)
    page.evaluate("""({scenario,serial,dates})=>{
        const parse=ToolOCR.parse;window.ownerFragmentRecognizeCalls=0;
        window.ToolOCR={...ToolOCR,recognize:async(file,tools)=>{
            ownerFragmentRecognizeCalls++;
            const target=tools.find(tool=>tool.serial===serial);
            if(scenario==='ambiguous'){
                ToolApp.state.model={...ToolApp.state.model,tools:[...tools,{...target,
                    id:'synthetic-owner-fragment-collision',serial:'99071688',months:12,
                    interval:'12',cycle:'MONTHS',due:'폐기'}]};tools=ToolApp.state.model.tools;
            }
            if(scenario==='raw-incomplete')return parse('0713\\n'+target.name+'|\\n2026.06\\n2026.12',tools);
            const fragment=scenario==='wrong-fragment'?'0713':'0716';
            return parse(dates,tools,{confidence:95,passes:[
                {text:dates,source:'original',confidence:95},
                {text:target.name,source:'identity-crop-owner',confidence:95},
                {text:'Serial: '+fragment,source:'identity-crop-left',confidence:95},
                {text:'Serial: '+fragment,source:'identity-crop-right',confidence:95}]});
        }};
    }""", {"scenario": scenario, "serial": SERIAL, "dates": LABELS})
    photo = fixture_photo(temporary / f"owner-fragment-{scenario}.png")
    page.locator("#photos").set_input_files(photo)
    page.wait_for_function("ToolApp.state.photos.length===1&&!!ToolApp.state.photos[0].result&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=45000)
    record = page.evaluate("""async()=>{const photo=ToolApp.state.photos[0];return {
        calls:ownerFragmentRecognizeCalls,saved:!!photo.saved,approval:photo.approval,
        assessment:photo.assessment,evidence:photo.result.identityEvidence,
        confidence:photo.result.confidence,photoHash:await testHash(photo.image.photoBytes),
        sourceHash:await testHash(await(await testHandle.getFile()).arrayBuffer()),pickerCalls:testPickerCalls};} """)
    output = temporary / f"owner-fragment-{scenario}.xlsx"
    export_shared(page, output)
    assert record["calls"] == 1 and record["pickerCalls"] == 1, record
    assert not errors and not remote, (errors, remote)
    result = {"real_opfs": True, "recognition_calls": 1}
    if scenario == "success":
        assert record["saved"] and record["approval"] == "automatic", record
        assert record["assessment"]["matchMethod"] == "owner-partial-serial", record
        assert SERIAL not in record["evidence"]["observedSerials"], record
        assert record["sourceHash"] != original_hash, record
        result.update(verify_torque_workbook(args.workbook, output))
        with ZipFile(args.workbook) as old, ZipFile(output) as new:
            added = [name for name in set(new.namelist()) - set(old.namelist())
                     if name.startswith("xl/media/") and not name.endswith("/")]
            assert len(added) == 1, added
            assert sha256(new.read(added[0])).hexdigest() == record["photoHash"]
        with page.expect_download() as event:
            page.locator("#download-backup").click()
        backup = temporary / "owner-fragment-pre-write-backup.xlsx"
        event.value.save_as(backup)
        assert backup.read_bytes() == args.workbook.read_bytes()
        result.update({"dates_and_actual_photo_saved": True, "exact_pre_write_backup": True})
    else:
        assert not record["saved"] and record["approval"] is None, record
        assert not record["assessment"]["eligible"] and record["assessment"]["reasons"], record
        assert record["sourceHash"] == original_hash, record
        assert output.read_bytes() == args.workbook.read_bytes()
        result["pending_workbook_byte_identical"] = True
        if scenario == "raw-incomplete":
            assert record["confidence"] is None, record
            assert all(proof["confidence"] is None for proof in record["evidence"]["observedSerialFragments"]), record
    context.close()
    return result


def main():
    args = arguments(__doc__)
    original_hash = sha256(args.workbook.read_bytes()).hexdigest()
    with tempfile.TemporaryDirectory(prefix="tool-owner-fragment-") as directory, serve_program(args.program) as url, sync_playwright() as playwright:
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
        context = browser.new_context()
        result = {"owner_fragment_policy": policy_checks(context.new_page(), url, args.workbook)}
        context.close()
        result["connected_automatic"] = {scenario: connected_check(browser, url, args, Path(directory), scenario)
            for scenario in ["success", "ambiguous", "wrong-fragment", "raw-incomplete"]}
        browser.close()
    assert sha256(args.workbook.read_bytes()).hexdigest() == original_hash
    result.update({"original_company_workbook_unchanged": True,
                   "actual_image_ocr_accuracy": "not measured: deterministic parser/evidence test doubles"})
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
