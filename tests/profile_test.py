"""Browser tests for local confirmed-match memory, not external OCR training.

Uses real localStorage and real OCR parsing with synthetic identity fixtures.
No user workbook, photo, or calibration history is written to profile storage.
"""
import argparse
from hashlib import sha256
import json
from pathlib import Path
import tempfile
from urllib.parse import urlsplit

from playwright.sync_api import sync_playwright

from browser_test import REPO, fixture_photo, serve_program, verify_workbook
from connected_file_test import PICKER_SCRIPT, export_shared, seed_file


HELPERS = r"""()=>{
  window.modelFixture=()=>({tools:[
    {id:'a',sheet:'광주',row:6,name:'강건욱',serial:'31550075',asset:'South_광주_CT_강건욱',months:12,interval:12,cycle:'MONTHS',last:'2025-12-29',due:'2026-12-29',photoCount:1},
    {id:'b',sheet:'광주',row:7,name:'김태수',serial:'44350257SV',asset:'South_광주_MR_김태수',months:12,interval:12,cycle:'MONTHS',last:'2025-12-29',due:'2026-12-29',photoCount:1},
    {id:'c',sheet:'광주',row:8,name:'강건욱',serial:'0716805829',asset:'South_광주_CT_강건욱',months:6,interval:6,cycle:'MONTHS',last:'2025-12-29',due:'2026-06-29',photoCount:1},
    {id:'d',sheet:'대전',row:2,name:'이휘준',serial:'001234',asset:'South_대전_CT_이휘준',months:12,interval:12,cycle:'MONTHS',last:'2025-12-29',due:'2026-12-29',photoCount:1}
  ]});
  window.parseFixture=(model,text,confidence=95)=>{
    text+='\n교정일자 2026.12.29\n차기교정예정일자 2027.12.29';
    return ToolOCR.parse(text,model.tools,{confidence,passes:[{text,confidence},{text,confidence}]});
  };
  window.photoFixture=(model,id,result,approval='manual')=>{
    const tool=model.tools.find(t=>t.id===id);
    return {approval,review:{toolId:id,name:tool.name,serial:tool.serial,last:'2026-12-29',due:'2027-12-29'},
      result,assessment:{matchMethod:'serial'},file:{name:'NEVER_STORE_PHOTO_FILENAME_MARKER.jpg'},url:'blob:NEVER_STORE_RAW_IMAGE',image:{photoBytes:new Uint8Array([1,2,3,4]),photoExtension:'jpg'}};
  };
  window.serialTargets=result=>(result.identityEvidence.serialMatches||[]).map(p=>p.toolId);
  window.nameTargets=result=>(result.identityEvidence.nameMatches||[]).map(p=>p.toolId);
  window.trainFixture=async()=>{
    const model=modelFixture();
    await ToolProfiles.remember(model,photoFixture(model,'a',parseFixture(model,'이름: 강건웁\nSerial: 31550075')));
    await ToolProfiles.remember(model,photoFixture(model,'a',parseFixture(model,'이름: 강건욱\nSerial: 3155007S')));
    return model;
  };
  window.profilePayload=()=>JSON.parse(localStorage.getItem('tool-management-profiles-v1')||'{"profiles":[]}');
}"""


def harness(page, url, from_artifact=False):
    if from_artifact:
        page.goto(url)
        page.wait_for_function("window.ToolProfiles&&window.ToolOCR&&window.ToolAutomation")
        page.evaluate(HELPERS)
        return
    origin = f"{urlsplit(url).scheme}://{urlsplit(url).netloc}"
    page.route("**/profile-fixture.html", lambda route: route.fulfill(status=200, content_type="text/html", body="<!doctype html><title>Local match memory tests</title>"))
    page.goto(origin + "/profile-fixture.html")
    for name in ["ocr.js", "automation.js", "profiles.js"]:
        page.add_script_tag(path=REPO / "web" / name)
    page.evaluate(HELPERS)


def persistence_checks(page, url, from_artifact=False):
    harness(page, url, from_artifact)
    first = page.evaluate(r"""async()=>{
      await ToolProfiles.clear();const model=await trainFixture(),raw=parseFixture(model,'이름: 강건웁\nSerial: 3155007S');
      const original=JSON.stringify(raw),enhanced=ToolProfiles.enhance(model,raw),state=ToolProfiles.inspect();
      const storage=localStorage.getItem('tool-management-profiles-v1');
      return {state,serials:serialTargets(enhanced),names:nameTargets(enhanced),assessment:ToolAutomation.assess(model,enhanced),
        originalUnchanged:JSON.stringify(raw)===original,textUnchanged:enhanced.text===raw.text,dateUnchanged:JSON.stringify(enhanced.dateEvidence)===JSON.stringify(raw.dateEvidence),
        forbiddenStored:['2026-12-29','2027-12-29','2026.12.29','2027.12.29','NEVER_STORE_PHOTO_FILENAME_MARKER','NEVER_STORE_RAW_IMAGE','photoBytes','dateEvidence','review','text'].filter(s=>storage.includes(s)),payload:profilePayload()};
    }""")
    assert first["state"]["count"] == 1 and first["state"]["aliasCount"] >= 2, first
    assert first["state"]["persistent"] is True
    assert "a" in first["serials"] and "a" in first["names"]
    assert first["assessment"]["eligible"] and first["assessment"]["toolId"] == "a", first
    assert first["originalUnchanged"] and first["textUnchanged"] and first["dateUnchanged"]
    assert not first["forbiddenStored"], first["forbiddenStored"]
    page.reload()
    if from_artifact:
        page.wait_for_function("window.ToolProfiles&&window.ToolOCR&&window.ToolAutomation")
    else:
        for name in ["ocr.js", "automation.js", "profiles.js"]:
            page.add_script_tag(path=REPO / "web" / name)
    page.evaluate(HELPERS)
    restored = page.evaluate(r"""async()=>{
      const model=modelFixture(),raw=parseFixture(model,'이름: 강건웁\nSerial: 3155007S'),enhanced=ToolProfiles.enhance(model,raw),state=ToolProfiles.inspect();
      await ToolProfiles.clear();const cleared=ToolProfiles.inspect(),after=ToolProfiles.enhance(model,raw);
      return {state,serials:serialTargets(enhanced),cleared,after:serialTargets(after)};
    }""")
    assert restored["state"]["count"] == 1 and restored["state"]["aliasCount"] >= 2
    assert "a" in restored["serials"]
    assert restored["cleared"]["count"] == restored["cleared"]["aliasCount"] == 0
    assert "a" not in restored["after"]
    return {"actual_storage_reload": True, "count_clear": True, "photo_text_dates_not_stored": True, "raw_ocr_and_date_evidence_unchanged": True}


def policy_checks(page):
    outcomes = page.evaluate(r"""async()=>{
      const outcomes={};
      const fresh=async()=>{await ToolProfiles.clear();return modelFixture();};
      let m=await fresh();await ToolProfiles.remember(m,photoFixture(m,'a',parseFixture(m,'강건욱\nSerial: 31550075'),'automatic'));
      outcomes.trustedAutomatic=ToolProfiles.inspect().count===1;
      for(const [key,alter] of [
        ['unapproved',p=>p.approval=null],['missingReview',p=>p.review=null],['wrongReviewName',p=>p.review.name='다른사람'],
        ['wrongReviewSerial',p=>p.review.serial='99999999'],['automaticFailedRecognition',p=>{p.approval='automatic';p.result.failed=true;}],
        ['automaticLowConfidence',p=>{p.approval='automatic';p.result=parseFixture(m,'강건욱\nSerial: 31550075',20);}]
      ]){
        m=await fresh();const p=photoFixture(m,'a',parseFixture(m,'강건욱\nSerial: 31550075'));alter(p);await ToolProfiles.remember(m,p);outcomes[key]=ToolProfiles.inspect().count===0;
      }
      m=await fresh();const manuallyVerified=photoFixture(m,'a',parseFixture(m,'이름: 강건웁\nSerial: 31550075'));manuallyVerified.result.failed=true;
      await ToolProfiles.remember(m,manuallyVerified);outcomes.failedManualRemembersOnlyConfirmedCanonical=ToolProfiles.inspect().count===1&&ToolProfiles.inspect().aliasCount===0;
      // A learned serial is not a licence to repair digit counts or leading zeroes.
      for(const [id,alias] of [['a','315500075'],['a','3155075'],['d','01234'],['d','0001234'],['d','1234']]){
        m=await fresh();const t=m.tools.find(t=>t.id===id),r=parseFixture(m,'이름: '+t.name+'\nSerial: '+alias);
        await ToolProfiles.remember(m,photoFixture(m,id,r));const e=ToolProfiles.enhance(m,r);
        outcomes['length-'+alias]=!serialTargets(e).includes(id);
      }
      m=await fresh();await trainFixture();
      for(const [key,change] of [
        ['ownerChanged',a=>a.name='최진혁'],['assetChanged',a=>a.asset='DIFFERENT ASSET'],['monthsChanged',a=>{a.months=6;a.interval=6;}],
        ['disposed',a=>a.due='폐기']
      ]){
        const current=modelFixture();change(current.tools[0]);const r=parseFixture(current,'이름: 강건욱\nSerial: 3155007S');
        outcomes[key]=!serialTargets(ToolProfiles.enhance(current,r)).includes('a');
      }
      const deleted=modelFixture();deleted.tools=deleted.tools.filter(t=>t.id!=='a');
      outcomes.deleted=!serialTargets(ToolProfiles.enhance(deleted,parseFixture(deleted,'Serial: 3155007S'))).includes('a');
      const duplicated=modelFixture();duplicated.tools.push({...duplicated.tools[0],id:'duplicate',name:'최진혁'});
      outcomes.duplicateCanonical=!serialTargets(ToolProfiles.enhance(duplicated,parseFixture(duplicated,'Serial: 3155007S'))).includes('a');
      const moved=modelFixture();moved.tools[0].id='new-row-id';moved.tools[0].row=20;moved.tools[0].last='2026-12-29';moved.tools[0].due='2027-12-29';
      outcomes.rowAndDatesMayChange=serialTargets(ToolProfiles.enhance(moved,parseFixture(moved,'Serial: 3155007S'))).includes('new-row-id');
      outcomes.lowConfidenceNotPromoted=!serialTargets(ToolProfiles.enhance(m,parseFixture(m,'Serial: 3155007S',20))).includes('a');
      const failed=parseFixture(m,'Serial: 3155007S');failed.failed=true;
      outcomes.engineFailureNotPromoted=!serialTargets(ToolProfiles.enhance(m,failed)).includes('a');

      // New exact workbook identities always outrank remembered OCR variants.
      const knownSerial=modelFixture();knownSerial.tools[1].serial='3155007S';
      const exact=ToolProfiles.enhance(knownSerial,parseFixture(knownSerial,'Serial: 3155007S'));
      outcomes.newKnownSerialWins=serialTargets(exact).includes('b')&&!serialTargets(exact).includes('a');
      const knownName=modelFixture();knownName.tools[1].name='강건웁';
      const named=ToolProfiles.enhance(knownName,parseFixture(knownName,'이름: 강건웁'));
      outcomes.newKnownNameWins=nameTargets(named).includes('b')&&!nameTargets(named).includes('a');
      const partialSerial=parseFixture(m,'Serial: 3155007S\nSerial: 44350257SV');partialSerial.identityEvidence.serialMatches=[];
      outcomes.observedKnownSerialWinsWithoutMatches=!serialTargets(ToolProfiles.enhance(m,partialSerial)).includes('a');
      const partialName=parseFixture(m,'이름: 김태수\nSerial: 3155007S');partialName.identityEvidence.nameMatches=[];
      outcomes.observedKnownNameWinsWithoutMatches=!serialTargets(ToolProfiles.enhance(m,partialName)).includes('a');

      m=await fresh();m.tools[1].serial='3155007S';
      await ToolProfiles.remember(m,photoFixture(m,'a',parseFixture(m,'이름: 강건욱\nSerial: 3155007S')));
      outcomes.cannotLearnAnotherKnownSerial=!profilePayload().profiles.some(p=>p.serial==='31550075'&&(p.serialAliases||[]).includes('3155007S'));
      m=await fresh();const contradictory=parseFixture(m,'이름: 김태수\nSerial: 31550075');contradictory.identityEvidence.nameMatches=[];
      await ToolProfiles.remember(m,photoFixture(m,'a',contradictory,'automatic'));
      outcomes.observedOtherOwnerBlocksAutomaticRemember=ToolProfiles.inspect().count===0;
      m=await fresh();m.tools.push({...m.tools[0],id:'date-tool',name:'날짜담당',serial:'20261228',asset:'DATE TOOL'});
      const compactDate=parseFixture(m,'이름: 날짜담당\nSerial: 20261229');await ToolProfiles.remember(m,photoFixture(m,'date-tool',compactDate));
      outcomes.dateCannotBecomeSerialAlias=!serialTargets(ToolProfiles.enhance(m,compactDate)).includes('date-tool');
      m=await fresh();const certificate=parseFixture(m,'이름: 강건욱\nCertificate No: 3155007S\nSerial: 3155007S');
      await ToolProfiles.remember(m,photoFixture(m,'a',certificate));
      outcomes.certificateCannotBecomeSerialAlias=!profilePayload().profiles.some(p=>(p.serialAliases||[]).includes('3155007S'));

      // A repeated ambiguous variant is blocked, not assigned to the first record.
      await fresh();await trainFixture();m=modelFixture();m.tools[1].serial='3155007B';
      await ToolProfiles.remember(m,photoFixture(m,'b',parseFixture(m,'이름: 김태수\nSerial: 3155007S')));
      const collided=ToolProfiles.enhance(m,parseFixture(m,'Serial: 3155007S'));
      outcomes.serialCollisionBlocked=ToolProfiles.inspect().conflictCount>=1&&!serialTargets(collided).some(id=>['a','b'].includes(id));

      await fresh();m=modelFixture();m.tools[1].name='강건웅';
      await ToolProfiles.remember(m,photoFixture(m,'a',parseFixture(m,'이름: 강건웁\nSerial: 31550075')));
      await ToolProfiles.remember(m,photoFixture(m,'b',parseFixture(m,'이름: 강건웁\nSerial: 44350257SV')));
      const collidedName=ToolProfiles.enhance(m,parseFixture(m,'이름: 강건웁'));
      outcomes.nameCollisionBlocked=!nameTargets(collidedName).some(id=>['a','b'].includes(id));
      await fresh();m=await trainFixture();const remove=Storage.prototype.removeItem;
      Storage.prototype.removeItem=function(key){if(key==='tool-management-profiles-v1')throw new DOMException('TEST_CLEAR_DENIED','SecurityError');return remove.call(this,key);};
      let clearState;try{clearState=await ToolProfiles.clear();}finally{Storage.prototype.removeItem=remove;}
      outcomes.clearFailureWarnsAndClearsSession=clearState.count===0&&!clearState.persistent&&!!clearState.error&&!serialTargets(ToolProfiles.enhance(m,parseFixture(m,'Serial: 3155007S'))).includes('a')&&profilePayload().profiles.length===1;
      await ToolProfiles.clear();return outcomes;
    }""")
    failed = [name for name, passed in outcomes.items() if not passed]
    assert not failed, {"failed_policy_checks": failed, "all": outcomes}
    return {"policy_cases_passed": len(outcomes)}


def storage_unavailable(context, page, url, from_artifact=False):
    context.add_init_script("Object.defineProperty(window,'localStorage',{configurable:true,get(){throw new DOMException('TEST_STORAGE_DENIED','SecurityError');}})")
    harness(page, url, from_artifact)
    result = page.evaluate(r"""async()=>{
      const m=modelFixture(),r=parseFixture(m,'강건욱\nSerial: 31550075');
      await ToolProfiles.remember(m,photoFixture(m,'a',r));const e=ToolProfiles.enhance(m,r),state=ToolProfiles.inspect();await ToolProfiles.clear();
      return {state,eligible:ToolAutomation.assess(m,e).eligible,serials:serialTargets(e)};
    }""")
    assert result["eligible"] and "a" in result["serials"] and not result["state"]["persistent"], result
    return {"storage_denied_keeps_normal_ocr": True}


def export_privacy(page, url):
    page.goto(url)
    page.wait_for_function("window.ToolProfiles&&window.ToolDistribution")
    page.evaluate(HELPERS)
    result = page.evaluate(r"""async()=>{
      await ToolProfiles.clear();const m=modelFixture();m.tools[0].name='홍길동';m.tools[0].asset='PRIVATE PROFILE ASSET';m.tools[0].serial='L0C4L9A1';
      const r=parseFixture(m,'이름: 홍길돔\nSerial: L0C4L9A1');await ToolProfiles.remember(m,photoFixture(m,'a',r));
      return {state:ToolProfiles.inspect(),leaks:['홍길돔','PRIVATE PROFILE ASSET','L0C4L9A1'].filter(s=>ToolDistribution.html().includes(s))};
    }""")
    assert result["state"]["count"] == 1 and result["state"]["aliasCount"] >= 1, result
    assert not result["leaks"], result
    return {"exported_html_excludes_learned_values": True}


def app_integration(context, page, url, args, temporary):
    """Only the app calls remember/enhance; OCR responses are test doubles."""
    context.add_init_script("(" + PICKER_SCRIPT + ")()")
    page.goto(url)
    page.wait_for_function("window.ToolApp&&window.ToolProfiles")
    seed_file(page, args.workbook)
    page.locator("#connect-onedrive").click()
    page.wait_for_function("ToolApp.state.source.connected&&!!ToolApp.state.model&&!ToolApp.state.busy")
    assert page.locator("#remember-matches").is_checked()
    assert page.evaluate("ToolProfiles.inspect().count") == 0

    def stub(text):
        page.evaluate("""text=>{
          const parse=ToolOCR.parse;
          window.ToolOCR={...ToolOCR,recognize:async(file,tools)=>parse(text,tools,{confidence:95,passes:[{text,confidence:95},{text,confidence:95}]})};
        }""", text)

    def wait_saved():
        page.wait_for_function("ToolApp.state.photos.length===1&&!!ToolApp.state.photos[0].saved&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=30000)

    photo = fixture_photo(args, temporary)
    stub("강건육\nSerial: 31550075\n교정일자 2025.12.29\n차기교정예정일자 2026.12.29")
    page.locator("#photos").set_input_files(photo)
    wait_saved()
    first = temporary / "profile-first-source.xlsx"
    export_shared(page, first)
    verify_workbook(args.workbook, first, last="2025-12-29", due="2026-12-29")
    assert page.evaluate("ToolApp.state.photos[0].approval") == "automatic"
    learned = page.evaluate("ToolProfiles.inspect()")
    assert learned["count"] == 1 and learned["aliasCount"] >= 1, learned
    assert "기억한 툴 1개" in page.locator("#memory-status").inner_text()
    first_save_count = page.evaluate("ToolApp.state.source.saveCount")

    # A new photo has only the learned handwriting variant and a 12-month span.
    # The two real 강건욱 rows have different cycles (12 and 6 months).
    page.locator("#remove-photo").click()
    second_photo = temporary / "profile-name-only.png"
    second_photo.write_bytes(photo.read_bytes())
    stub("강건육\n교정일자 2026.12.29\n차기교정예정일자 2027.12.29")
    page.locator("#photos").set_input_files(second_photo)
    wait_saved()
    second_record = page.evaluate("""()=>{
      const p=ToolApp.state.photos[0];return {approval:p.approval,serial:p.saved.serial,row:p.saved.row,
        rememberedName:p.result.identityEvidence.resolvedNames.some(v=>v.learned&&v.alias==='강건육'),
        observedSerials:p.result.identityEvidence.observedSerials,saveCount:ToolApp.state.source.saveCount};
    }""")
    assert second_record["approval"] == "automatic" and second_record["serial"] == "31550075" and second_record["row"] == 6, second_record
    assert second_record["rememberedName"] and not second_record["observedSerials"], second_record
    assert second_record["saveCount"] == first_save_count + 1
    second = temporary / "profile-name-only-source.xlsx"
    export_shared(page, second)
    verify_workbook(first, second, last="2026-12-29", due="2027-12-29")

    page.locator("#remove-photo").click()
    page.locator("#clear-matches").click()
    assert page.evaluate("ToolProfiles.inspect().count") == 0
    assert "기억한 툴 0개" in page.locator("#memory-status").inner_text()
    third_photo = temporary / "profile-after-clear.png"
    third_photo.write_bytes(photo.read_bytes())
    stub("강건육\n교정일자 2027.12.29\n차기교정예정일자 2028.12.29")
    page.locator("#photos").set_input_files(third_photo)
    page.wait_for_function("ToolApp.state.photos.length===1&&!!ToolApp.state.photos[0].result&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving", timeout=30000)
    cleared = page.evaluate("""()=>{const p=ToolApp.state.photos[0];return {review:!!p.review,saved:!!p.saved,eligible:p.assessment.eligible,
      prepared:!!ToolApp.state.preparedResult,saveCount:ToolApp.state.source.saveCount,count:ToolProfiles.inspect().count};}""")
    assert not cleared["review"] and not cleared["saved"] and not cleared["eligible"] and not cleared["prepared"], cleared
    assert cleared["saveCount"] == second_record["saveCount"] and cleared["count"] == 0
    unchanged = temporary / "profile-cleared-source.xlsx"
    export_shared(page, unchanged)
    assert unchanged.read_bytes() == second.read_bytes()
    return {"ocr_responses_are_test_doubles": True, "app_remembers_after_automatic_source_save": True,
            "remembered_name_and_cycle_saves_without_serial": True, "ui_clear_blocks_same_alias_and_preserves_source": True,
            "photos": 59, "cells_checked_per_save": 352}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--program", type=Path, default=REPO / "web/dist/Tool-Manager.html")
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    parser.add_argument("--module-only", action="store_true")
    parser.add_argument("--workbook", type=Path)
    args = parser.parse_args()
    args.photo = None
    if not args.module_only and args.workbook is None:
        args.workbook = next(Path("/workspace/attachments").rglob("*.xlsx"), None)
    if not args.module_only and args.workbook is None:
        parser.error("Use --workbook with the company workbook for app integration, or --module-only")
    original_hash = sha256(args.workbook.read_bytes()).hexdigest() if args.workbook else None
    with serve_program(args.program) as url, sync_playwright() as p:
        browser = p.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
        context = browser.new_context(locale="ko-KR")
        page = context.new_page()
        errors = []
        page.on("pageerror", lambda error: errors.append(str(error)))
        results = {"persistence": persistence_checks(page, url, from_artifact=not args.module_only), "policies": policy_checks(page)}
        context.close()
        context = browser.new_context(locale="ko-KR")
        page = context.new_page()
        page.on("pageerror", lambda error: errors.append(str(error)))
        results["storage_unavailable"] = storage_unavailable(context, page, url, from_artifact=not args.module_only)
        context.close()
        if not args.module_only:
            context = browser.new_context(locale="ko-KR")
            page = context.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))
            results["privacy"] = export_privacy(page, url)
            context.close()
            context = browser.new_context(locale="ko-KR", accept_downloads=True)
            page = context.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))
            with tempfile.TemporaryDirectory(prefix="tool-profile-integration-") as directory:
                results["app_integration"] = app_integration(context, page, url, args, Path(directory))
            context.close()
        browser.close()
    assert not errors, errors
    if original_hash:
        assert original_hash == sha256(args.workbook.read_bytes()).hexdigest()
    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
