"""Exercise fast-field stopping and failure paths with controlled OCR responses.

An attached high-resolution photograph supplies unchanged pixels to production
paper/table detection and cropping. Only worker responses are controlled. These
checks are orchestration/safety checks, not camera OCR accuracy measurements.
The photo and company workbook are runtime inputs and remain outside Git.
"""
from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path
from urllib.parse import urlsplit

from playwright.sync_api import sync_playwright

from browser_test import arguments, serve_program


HARNESS = r"""scenario => {
 const nativeNow=Date.now.bind(Date);let offset=0;
 Date.now=()=>nativeNow()+offset;
 window.fastSafetyCalls=[];window.fastSafetyParameters={};window.fastSafetyTerminations=0;
 window.fastSafetyHung=false;window.fastSafetyClockJump=false;
 const tools=ToolApp.state.model.tools;
 const usable=tools.filter(t=>/^[가-힣]{2,4}$/.test(t.name)&&/^20\d{2}-\d{2}-\d{2}$/.test(t.last)
   &&/^20\d{2}-\d{2}-\d{2}$/.test(t.due)&&t.due>t.last&&t.photoCount===1
   &&tools.filter(o=>o.serial===t.serial).length===1);
 const primary=usable[0],other=usable.find(t=>t.serial!==primary.serial&&t.name!==primary.name);
 if(!primary||!other)throw Error('Need two distinct registered active tools');
 const dates='교정일자 '+primary.last.replaceAll('-','.')+'\n차기교정예정일자 '+primary.due.replaceAll('-','.');
 const identity=t=>t.name+'\n'+t.serial;
 const clear=identity(primary)+'\n'+dates;
 Tesseract.createWorker=async()=>({
  setParameters:async p=>Object.assign(fastSafetyParameters,p),
  recognize:async canvas=>{
   const index=fastSafetyCalls.length;
   fastSafetyCalls.push({width:canvas.width,height:canvas.height,parameters:{...fastSafetyParameters}});
   if(index===0&&scenario==='cancel'){
    window.fastSafetyHung=true;return new Promise(()=>{});
   }
   if(index===0){
    if(scenario==='budget'){offset=26000;window.fastSafetyClockJump=true;}
    return {data:{text:identity(primary),confidence:95,words:[]}};
   }
   if(index===1&&scenario==='conflict')return {data:{text:identity(other),confidence:95,words:[]}};
   if(index===1&&scenario==='clear')return {data:{text:'',confidence:0,words:[]}};
   // Once the deadline has skipped tag 2, the genuine normal-route overview
   // still reads a clear identity and labelled date pair. It must not erase
   // the uninspected-identity marker left by the fast stage.
   if(scenario==='budget')return {data:{text:clear,confidence:95,words:[]}};
   return {data:{text:dates,confidence:95,words:[]}};
  },
  terminate:()=>{fastSafetyTerminations++;}
 });
 window.fastSafetyPromise=ToolOCR.recognize(document.querySelector('#fast-safety-photo').files[0],tools)
  .then(result=>({result,assessment:ToolAutomation.assess({tools},result),primary,other,
    calls:fastSafetyCalls,terminations:fastSafetyTerminations,clockJump:fastSafetyClockJump}))
  .finally(()=>{Date.now=nativeNow;});
 return true;
}"""


def main():
    args = arguments(__doc__)
    if args.photo is None or not args.photo.is_file():
        raise SystemExit("Supply --photo with an attached high-resolution calibration photograph")
    before_book = sha256(args.workbook.read_bytes()).hexdigest()
    before_photo = sha256(args.photo.read_bytes()).hexdigest()
    errors, remote, outcomes = [], [], {}
    with serve_program(args.program) as url, sync_playwright() as playwright:
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
        for scenario in ["clear", "conflict", "budget", "cancel"]:
            context = browser.new_context(locale="ko-KR")
            page = context.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))
            context.on("request", lambda request: remote.append(request.url)
                       if urlsplit(request.url).scheme in {"http", "https"}
                       and urlsplit(request.url).netloc != urlsplit(url).netloc else None)
            page.goto(url)
            page.wait_for_function("window.ToolApp&&window.ToolOCR&&window.ToolFastOCR")
            page.locator("#workbook").set_input_files(args.workbook)
            page.wait_for_function("ToolApp.state.model&&!ToolApp.state.busy")
            page.evaluate("""()=>{const e=document.createElement('input');e.id='fast-safety-photo';
              e.type='file';document.body.append(e)}""")
            page.locator("#fast-safety-photo").set_input_files(args.photo)
            page.evaluate(HARNESS, scenario)
            if scenario == "cancel":
                page.wait_for_function("window.fastSafetyHung", timeout=30000)
                page.evaluate("ToolOCR.cancel()")
            evidence = page.evaluate("window.fastSafetyPromise")
            result, assessment = evidence["result"], evidence["assessment"]
            sources = [p["source"] for p in result["passes"]]
            if scenario == "clear":
                assert result.get("fastPath") is True and assessment["eligible"], evidence
                assert "original" not in sources, sources
                assert len(evidence["calls"]) == 4, evidence
                assert assessment["toolId"] == evidence["primary"]["id"], evidence
            elif scenario == "conflict":
                assert not assessment["eligible"], evidence
                observed = {p["serial"] for p in result["identityEvidence"]["serialMatches"]}
                assert {evidence["primary"]["serial"], evidence["other"]["serial"]} <= observed, evidence
                assert not result.get("failed"), evidence
            elif scenario == "budget":
                assert evidence["clockJump"] and not result.get("failed"), evidence
                assert "identity-probe-incomplete" in sources and "original" in sources, evidence
                assert result["identityRecoveryIncomplete"] and not assessment["eligible"], evidence
                assert assessment["toolId"] == evidence["primary"]["id"], evidence
                assert (assessment["last"], assessment["due"]) == (evidence["primary"]["last"], evidence["primary"]["due"]), evidence
                assert len(assessment["reasons"]) == 1, evidence
            else:
                assert result["failed"] and result["cancelled"] and result["code"] == "cancelled", evidence
                assert not assessment["eligible"] and evidence["terminations"] >= 1, evidence
                assert len(evidence["calls"]) == 1, evidence
            outcomes[scenario] = {"eligible": assessment["eligible"], "failed": result.get("failed"),
                "cancelled": result.get("cancelled", False), "incomplete": result["identityRecoveryIncomplete"],
                "pass_sources": sources, "ocr_calls": len(evidence["calls"]), "terminations": evidence["terminations"]}
            context.close()
        browser.close()
    assert not errors and not remote, (errors, remote)
    assert before_book == sha256(args.workbook.read_bytes()).hexdigest()
    assert before_photo == sha256(args.photo.read_bytes()).hexdigest()
    print(json.dumps({"cases": outcomes, "page_errors": errors, "external_requests": remote,
        "original_photo_and_workbook_unchanged": True,
        "accuracy_scope": "Real photo pixels and production fast pipeline; controlled OCR responses, not photo accuracy"}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
