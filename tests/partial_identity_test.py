"""Partial serial policy and connected XLSX safety, using controlled OCR evidence.

ToolOCR.parse, ToolAutomation.assess, app approval and XLSX/OPFS writes are real.
Recognition responses are deterministic test doubles; image OCR accuracy is not
measured. The attached company workbook remains read only, and generated files
stay in a temporary directory. Owner names come from the workbook at runtime.
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

from browser_test import arguments, serve_program
from connected_file_test import PICKER_SCRIPT, export_shared, ready, seed_file
from torque_dates_test import SERIAL, LAST, DUE, LABELS, verify_torque_workbook


def policy_checks(page, url, workbook):
    page.goto(url)
    page.wait_for_function("window.ToolOCR&&window.ToolAutomation&&window.ToolWorkbook")
    page.evaluate("""()=>{
        const input=document.createElement('input');input.type='file';input.id='fixture';input.hidden=true;
        document.body.append(input);
    }""")
    page.locator("#fixture").set_input_files(workbook)
    outcomes = page.evaluate("""async ({serial,dates}) => {
        const model=await ToolWorkbook.load(await document.querySelector('#fixture').files[0].arrayBuffer());
        const target=model.tools.find(tool=>tool.serial===serial);
        const other=model.tools.find(tool=>tool.name!==target.name&&tool.serial&&tool.due!=='폐기');
        if(!target||!other)throw new Error('Company fixture needs target and another owner');
        const fragment=serial.slice(-6),short=serial.slice(-5),cases={};
        const copy=value=>JSON.parse(JSON.stringify(value));
        const pass=(text,source,confidence=95)=>({text,source,confidence});
        const partial=(value=fragment,confidence=95)=>[
            pass('Serial: '+value,'identity-crop-upper-left',confidence),
            pass('Serial: '+value,'identity-crop-upper-right',confidence)];
        function run(name,{passes=partial(),owner=false,tools=model.tools,dateText=dates,mutate=null,eligible=false,
                          method='partial-serial',observed=fragment,expectedSerial=serial}={}){
            const all=[pass(dateText,'original'),...passes];
            if(owner)all.push(pass(target.name,'identity-crop-owner'));
            const parsed=ToolOCR.parse(dateText,tools,{confidence:95,passes:all});
            if(mutate)mutate(parsed);
            const before=JSON.stringify(parsed.identityEvidence);
            const assessment=ToolAutomation.assess({...model,tools},parsed);
            cases[name]={assessment,expected:{eligible,method,observed,serial:expectedSerial},
                evidence:parsed.identityEvidence,parserSerial:parsed.serial,parserCandidates:parsed.candidateIds,
                evidenceUnchanged:before===JSON.stringify(parsed.identityEvidence)};
        }
        run('six_digit_suffix_repeated_at_sixty_percent',{eligible:true});
        run('six_digit_prefix_repeated',{passes:partial(serial.slice(0,6)),observed:serial.slice(0,6),eligible:true});
        run('six_digit_middle_repeated',{passes:partial(serial.slice(2,8)),observed:serial.slice(2,8),eligible:true});
        run('hyphenated_observed_serial_fragment',{passes:partial(fragment.slice(0,3)+'-'+fragment.slice(3)),eligible:true});
        run('spaced_observed_serial_fragment',{passes:partial(fragment.slice(0,3)+' '+fragment.slice(3)),eligible:true});
        run('tiny_spaced_tokens_without_geometry_are_not_guessed',{
            passes:partial(fragment.slice(0,2)+' '+fragment.slice(2,4)+' '+fragment.slice(4))});
        run('dropped_leading_zero_explicitly_partial',{passes:partial(serial.slice(1)),observed:serial.slice(1),eligible:true});
        run('owner_allows_five_digits_at_fifty_percent',{passes:partial(short),owner:true,observed:short,eligible:true});
        run('five_digits_without_owner',{passes:partial(short),observed:short});
        // 4.5 adds the separately corroborated owner/four-digit path.
        run('four_digits_with_owner',{passes:partial(serial.slice(-4)),owner:true,observed:serial.slice(-4),
            eligible:true,method:'owner-partial-serial'});
        run('only_one_observation',{passes:partial().slice(0,1)});
        run('one_dropped_zero_observation',{passes:partial(serial.slice(1)).slice(0,1),observed:serial.slice(1)});
        run('same_source_is_not_two_independent_reads',{passes:[
            pass('Serial: '+fragment,'identity-crop-upper-left'),pass('Serial: '+fragment,'identity-crop-upper-left')]});
        run('one_of_two_readings_is_low_confidence',{passes:[
            pass('Serial: '+fragment,'identity-crop-upper-left'),pass('Serial: '+fragment,'identity-crop-upper-right',25)]});
        run('both_readings_below_threshold',{passes:partial(fragment,69)});
        run('both_readings_at_threshold',{passes:partial(fragment,70),eligible:true});
        run('no_source_confidence',{mutate:parsed=>{
            for(const proof of parsed.identityEvidence.observedSerialEvidence){proof.confidence=null;
                for(const reading of proof.sourceReadings||[])reading.confidence=null;}
        }});
        run('legacy_aggregate_confidence_is_insufficient',{mutate:parsed=>{
            for(const proof of parsed.identityEvidence.observedSerialEvidence)delete proof.sourceReadings;
        }});
        run('observed_serial_strings_without_actual_evidence_are_insufficient',{mutate:parsed=>{
            delete parsed.identityEvidence.observedSerialEvidence;
        }});
        run('source_names_without_source_confidence_are_insufficient',{mutate:parsed=>{
            for(const proof of parsed.identityEvidence.observedSerialEvidence)
                for(const reading of proof.sourceReadings||[])delete reading.confidence;
        }});
        run('confidence_in_an_unlisted_source_is_insufficient',{mutate:parsed=>{
            for(const proof of parsed.identityEvidence.observedSerialEvidence)proof.passSources=['different-source'];
        }});
        run('date_only_crops_cannot_supply_identity',{passes:[
            pass('Serial: '+fragment,'date-crop-6'),pass('Serial: '+fragment,'date-crop-11')]});
        run('label_only_crops_cannot_supply_identity',{passes:[
            pass('Serial: '+fragment,'label-crop-6'),pass('Serial: '+fragment,'label-crop-11')]});
        run('anonymous_text_sources_are_not_independent',{passes:[
            pass('Serial: '+fragment,''),pass('Serial: '+fragment,'text')]});
        run('incomplete_identity_recovery_blocks_partial',{mutate:parsed=>{parsed.identityRecoveryIncomplete=true;}});
        run('dates_without_explicit_due_do_not_resolve_partial',{dateText:'교정일자 2026.06.22'});
        run('valid_dates_with_wrong_workbook_cycle_do_not_resolve_partial',{
            dateText:'교정일자 2026.06.22\\n차기교정예정일자 2027.06.22'});
        run('low_confidence_dates_do_not_resolve_partial',{mutate:parsed=>{parsed.dateEvidence.confidence=25;}});
        run('conflicting_dates_do_not_resolve_partial',{passes:[...partial(),
            pass('교정일자 2026.06.22\\n차기교정예정일자 2026.12.21','date-crop-6')]});
        run('partial_does_not_replace_missing_workbook_cycle',{
            tools:model.tools.map(tool=>tool.id===target.id?{...tool,months:null,interval:'',cycle:''}:tool)});
        run('partial_does_not_override_a_different_owner',{passes:[...partial(),pass(other.name,'identity-crop-owner')]});
        run('partial_does_not_override_another_known_serial',{passes:[...partial(),pass('Serial: '+other.serial,'identity-crop-other')]});
        run('unknown_full_number_is_not_truncated_to_its_tail',{
            passes:partial('991'+serial),observed:'991'+serial});
        run('unknown_full_number_conflicts_with_partial',{passes:[...partial(),pass('Serial: 999999999','identity-crop-other')]});
        run('independently_supported_compatible_fragments_resolve_same_tool',{passes:[
            ...partial(),...partial(serial.slice(0,6))],eligible:true});
        run('contradictory_fragment_is_not_repaired_to_known_digits',{passes:[
            ...partial(),...partial('071689')]});
        run('leading_letter_o_is_not_zero',{passes:partial('O'+serial.slice(1)),observed:'O'+serial.slice(1)});
        const collision={...target,id:'synthetic-partial-collision',serial:'99'+fragment,name:other.name,months:12,
            interval:'12',cycle:'MONTHS',due:'2027-12-22'};
        run('global_collision_cannot_be_removed_by_owner_or_cycle',{
            owner:true,tools:[...model.tools,collision]});
        run('disposed_tool_still_counts_for_global_uniqueness',{
            owner:true,tools:[...model.tools,{...collision,due:'폐기'}]});
        run('fragment_that_is_another_tools_full_serial_is_not_partial',{
            owner:true,tools:[...model.tools,{...collision,serial:fragment}]});
        const longSerial='A1B2C3D4E5F6',longTools=model.tools.map(tool=>tool.id===target.id?{...tool,serial:longSerial}:tool);
        run('six_digits_below_sixty_percent_are_insufficient',{
            passes:partial(longSerial.slice(0,6)),tools:longTools,observed:longSerial.slice(0,6),expectedSerial:longSerial});
        run('longer_alphanumeric_contiguous_partial_is_supported',{
            passes:partial(longSerial.slice(0,8)),tools:longTools,observed:longSerial.slice(0,8),expectedSerial:longSerial,eligible:true});
        run('owner_supports_half_of_longer_alphanumeric_serial',{
            passes:partial(longSerial.slice(0,6)),owner:true,tools:longTools,observed:longSerial.slice(0,6),expectedSerial:longSerial,eligible:true});
        run('owner_five_digits_below_fifty_percent_are_insufficient',{
            passes:partial(longSerial.slice(0,5)),owner:true,tools:longTools,observed:longSerial.slice(0,5),expectedSerial:longSerial});
        run('independent_full_read_resolves_same_candidate_partial',{
            passes:[...partial().slice(0,1),pass('Serial: '+serial,'identity-crop-full')],eligible:true,method:'serial'});
        run('full_and_partial_from_same_source_are_not_independent',{
            passes:[pass('Serial: '+fragment+'\\nSerial: '+serial,'identity-crop-full')]});
        run('full_read_does_not_resolve_short_partial',{
            passes:[pass('Serial: '+short,'identity-crop-partial'),pass('Serial: '+serial,'identity-crop-full')],observed:short});
        run('full_read_does_not_resolve_conflicting_fragment',{
            passes:[...partial().slice(0,1),pass('Serial: '+serial+'\\nSerial: 999999999','identity-crop-full')]});
        return cases;
    }""", {"serial": SERIAL, "dates": LABELS})
    for name, outcome in outcomes.items():
        assessment, expected = outcome["assessment"], outcome["expected"]
        assert assessment["eligible"] is expected["eligible"], (name, outcome)
        assert outcome["evidenceUnchanged"], (name, outcome)
        if expected["eligible"]:
            assert assessment["serial"] == expected["serial"], (name, outcome)
            assert assessment["matchMethod"] == expected["method"], (name, outcome)
            assert assessment["last"] == LAST and assessment["due"] == DUE, (name, outcome)
            observed = outcome["evidence"]["observedSerials"]
            assert expected["observed"] in observed, (name, outcome)
            if expected["method"] == "partial-serial":
                # Canonical serial belongs in the approved row, never in OCR evidence.
                assert expected["serial"] not in observed, (name, outcome)
                assert outcome["evidence"]["serialMatches"] == [], (name, outcome)
                assert outcome["parserSerial"] == "", (name, outcome)
        else:
            assert assessment["reasons"], (name, outcome)
    return {"cases": len(outcomes),
            "accepted": [name for name, value in outcomes.items() if value["expected"]["eligible"]],
            "pending": [name for name, value in outcomes.items() if not value["expected"]["eligible"]],
            "canonical_digits_are_not_added_to_ocr_evidence": True,
            "source_confidence_and_global_uniqueness_guards": True}


def fixture_photo(path):
    image = Image.new("RGB", (600, 800), (220, 230, 240))
    draw = ImageDraw.Draw(image)
    draw.rectangle((100, 60, 520, 180), fill="white", outline="black", width=4)
    draw.text((120, 100), "805829", fill="black", font_size=42)
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
        const parse=ToolOCR.parse,target=ToolApp.state.model.tools.find(tool=>tool.serial===serial);
        window.partialRecognizeCalls=0;
        window.ToolOCR={...ToolOCR,recognize:async(file,tools)=>{
            partialRecognizeCalls++;
            // Connected recognition rereads OPFS first. Add this explicit
            // synthetic collision to that fresh reference model, not to a
            // stale pre-read model that the app is required to replace.
            if(scenario==='ambiguous'){
                ToolApp.state.model={...ToolApp.state.model,
                    tools:[...tools,{...target,id:'synthetic-connected-collision',
                        serial:'99'+serial.slice(-6),due:'폐기'}]};
                tools=ToolApp.state.model.tools;
            }
            const value=scenario==='leading-zero'?serial.slice(1):scenario==='owner-five'?serial.slice(-5):serial.slice(-6);
            const passes=[{text:dates,confidence:95,source:'original'},
                {text:'Serial: '+value,confidence:95,source:'identity-crop-left'},
                {text:'Serial: '+value,confidence:scenario==='low-confidence'?25:95,source:'identity-crop-right'}];
            if(scenario==='owner-five')passes.push({text:target.name,confidence:95,source:'identity-crop-owner'});
            if(scenario==='conflicting')passes.push({text:'Serial: 999999999',confidence:95,source:'identity-crop-other'});
            return parse(dates,tools,{confidence:95,passes});
        }};
    }""", {"scenario": scenario, "serial": SERIAL, "dates": LABELS})
    photo = fixture_photo(temporary / f"partial-{scenario}.png")
    page.locator("#photos").set_input_files(photo)
    page.wait_for_function("ToolApp.state.photos.length===1&&!!ToolApp.state.photos[0].result&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=45000)
    record = page.evaluate("""async()=>{
        const photo=ToolApp.state.photos[0];
        return {calls:partialRecognizeCalls,saved:!!photo.saved,approval:photo.approval,
            assessment:photo.assessment,evidence:photo.result.identityEvidence,
            sourceHash:await testHash(await(await testHandle.getFile()).arrayBuffer()),
            photoHash:await testHash(photo.image.photoBytes),pickerCalls:testPickerCalls};
    }""")
    output = temporary / f"partial-{scenario}-shared.xlsx"
    export_shared(page, output)
    assert record["calls"] == 1 and record["pickerCalls"] == 1, record
    assert not errors and not remote, (errors, remote)
    result = {"real_opfs": True, "recognition_calls": 1}
    if scenario in {"suffix", "leading-zero", "owner-five"}:
        assert record["saved"] and record["approval"] == "automatic", record
        assert record["assessment"]["matchMethod"] == "partial-serial", record
        assert record["assessment"]["serial"] == SERIAL, record
        assert SERIAL not in record["evidence"]["observedSerials"], record
        assert record["sourceHash"] != original_hash, record
        result.update(verify_torque_workbook(args.workbook, output))
        with ZipFile(args.workbook) as old, ZipFile(output) as new:
            added_media = [name for name in set(new.namelist()) - set(old.namelist())
                           if name.startswith("xl/media/") and not name.endswith("/")]
            assert len(added_media) == 1, added_media
            assert sha256(new.read(added_media[0])).hexdigest() == record["photoHash"]
        with page.expect_download() as event:
            page.locator("#download-backup").click()
        backup = temporary / f"partial-{scenario}-backup.xlsx"
        event.value.save_as(backup)
        assert backup.read_bytes() == args.workbook.read_bytes()
        result.update({"dates_and_photo_saved": True, "exact_pre_write_backup": True})
    else:
        assert not record["saved"] and record["approval"] is None, record
        assert not record["assessment"]["eligible"] and record["assessment"]["reasons"], record
        assert record["sourceHash"] == original_hash, record
        assert output.read_bytes() == args.workbook.read_bytes()
        result["pending_workbook_byte_identical"] = True
    context.close()
    return result


def main():
    args = arguments(__doc__)
    original_hash = sha256(args.workbook.read_bytes()).hexdigest()
    with tempfile.TemporaryDirectory(prefix="tool-partial-identity-") as directory, serve_program(args.program) as url, sync_playwright() as playwright:
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
        context = browser.new_context()
        page = context.new_page()
        result = {"partial_serial_policy": policy_checks(page, url, args.workbook)}
        context.close()
        result["connected_automatic"] = {scenario: connected_check(browser, url, args, Path(directory), scenario)
            for scenario in ["suffix", "leading-zero", "owner-five", "ambiguous", "conflicting", "low-confidence"]}
        browser.close()
    assert sha256(args.workbook.read_bytes()).hexdigest() == original_hash
    result.update({"original_company_workbook_unchanged": True,
                   "actual_image_ocr_accuracy": "not measured: deterministic parser/evidence test doubles"})
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
