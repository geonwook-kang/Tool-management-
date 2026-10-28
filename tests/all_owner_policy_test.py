"""Check every registered workbook identity through the real matching policy.

This uses controlled, high-confidence OCR text from the loaded workbook, not
camera photos. It measures row selection, duplicate handling and disposed-tool
rejection; actual photo recognition and timings are separate tests. Company
workbook content is read at runtime and is never embedded in this source.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
import json
from pathlib import Path
from urllib.parse import urlsplit

from playwright.sync_api import sync_playwright

from browser_test import REPO, serve_program


def arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--program", type=Path, default=REPO / "web/dist/Tool-Manager.html")
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    parser.add_argument("--diagnostics", type=Path, default=Path("/tmp/tool-all-owner-policy.json"))
    args = parser.parse_args()
    if args.workbook is None:
        args.workbook = next(Path("/workspace/attachments").rglob("*.xlsx"), None)
    for label, path in [("program", args.program), ("workbook", args.workbook)]:
        if path is None or not path.is_file():
            parser.error(f"Supply --{label} with an existing file")
    return args


CHECK = r"""() => {
 const model=ToolApp.state.model,tools=model.tools;
 const iso=value=>/^20\d{2}-\d{2}-\d{2}$/.test(String(value||''));
 const disposed=tool=>/폐기|disposed|retired/i.test(String(tool.due||''));
 const months=tool=>Number.isInteger(tool.months)?tool.months:null;
 const text=(tool,{serial=true,last=tool.last,due=tool.due}={})=>
   '이름: '+tool.name+'\n'+(serial?'Serial: '+tool.serial+'\n':'')+
   '교정일자 '+last.replaceAll('-','.')+'\n차기교정예정일자 '+due.replaceAll('-','.');
 const assess=(candidate,tool,options={})=>{
   const source=text(tool,options),parsed=ToolOCR.parse(source,candidate.tools,
     {confidence:95,passes:[{source:'original',text:source,confidence:95}]});
   return ToolAutomation.assess(candidate,parsed);
 };
 const rows=[],duplicates=[],ambiguous=[],retired=[],skipped=[];
 const active=tools.filter(tool=>!disposed(tool)&&tool.name&&tool.serial&&iso(tool.last)&&iso(tool.due)&&tool.due>tool.last);
 for(const tool of tools){
   if(disposed(tool))continue;
   if(!active.includes(tool)){skipped.push({id:tool.id,reason:'Missing registered identity or complete valid calibration dates'});continue;}
   const found=assess(model,tool);
   rows.push({id:tool.id,eligible:found.eligible,expectedEligible:!(Number(tool.photoCount)>1),ambiguousExistingPhotoSlot:Number(tool.photoCount)>1,correctRow:found.toolId===tool.id,
     correctFields:found.name===tool.name&&found.serial===tool.serial&&found.last===tool.last&&found.due===tool.due,
     reasons:found.reasons});
   const duplicate={...tool,id:tool.id+':duplicate-for-test'};
   const duplicateModel={...model,tools:tools.concat([duplicate])};
   const collision=assess(duplicateModel,tool);
   duplicates.push({id:tool.id,eligible:collision.eligible,reasons:collision.reasons});
 }
 const groups=new Map();
 for(const tool of active){
   const key=tool.name+'|'+months(tool);
   if(!groups.has(key))groups.set(key,[]);groups.get(key).push(tool);
 }
 for(const group of groups.values()){
   if(group.length<2||!months(group[0]))continue;
   const tool=group[0],found=assess(model,tool,{serial:false});
   ambiguous.push({ids:group.map(tool=>tool.id),eligible:found.eligible,reasons:found.reasons});
 }
 const dateReference=active[0];
 for(const tool of tools.filter(disposed)){
   const found=assess(model,tool,{last:dateReference.last,due:dateReference.due});
   retired.push({id:tool.id,eligible:found.eligible,reasons:found.reasons});
 }
 return {totalTools:tools.length,activeTools:active.length,activeOwners:new Set(active.map(tool=>tool.name)).size,
   rows,duplicates,ambiguous,retired,skipped};
}"""


def main():
    args = arguments()
    original_hash = sha256(args.workbook.read_bytes()).hexdigest()
    errors, requests = [], []
    with serve_program(args.program) as url, sync_playwright() as playwright:
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
        context = browser.new_context(locale="ko-KR")
        page = context.new_page()
        page.on("pageerror", lambda error: errors.append(str(error)))
        context.on("request", lambda request: requests.append(request.url)
                   if urlsplit(request.url).scheme in {"http", "https"}
                   and urlsplit(request.url).netloc != urlsplit(url).netloc else None)
        page.goto(url)
        page.wait_for_function("window.ToolApp&&window.ToolOCR&&window.ToolAutomation")
        page.locator("#workbook").set_input_files(args.workbook)
        page.wait_for_function("ToolApp.state.model&&!ToolApp.state.busy")
        result = page.evaluate(CHECK)
        browser.close()
    assert not errors and not requests, (errors, requests)
    assert sha256(args.workbook.read_bytes()).hexdigest() == original_hash
    report = {"scope": "Controlled OCR text through unmodified parser and matcher; no camera accuracy or speed claim",
              "program_sha256": sha256(args.program.read_bytes()).hexdigest(),
              "company_workbook_unchanged": True, "page_errors": errors, "external_requests": requests, **result}
    args.diagnostics.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    assert result["rows"], "The runtime workbook contains no active tools with valid registered dates"
    assert all(row["eligible"] == row["expectedEligible"] and row["correctRow"] and row["correctFields"]
               and (row["eligible"] or row["reasons"]) for row in result["rows"]), result
    assert all(not row["eligible"] and any("시리얼" in reason for reason in row["reasons"]) for row in result["duplicates"]), result
    assert all(not row["eligible"] and row["reasons"] for row in result["ambiguous"]), result
    assert all(not row["eligible"] and any("폐기" in reason for reason in row["reasons"]) for row in result["retired"]), result
    print(json.dumps({key: report[key] for key in ["scope", "totalTools", "activeTools", "activeOwners", "company_workbook_unchanged", "page_errors", "external_requests"]}
                     | {"duplicate_serial_cases":len(result["duplicates"]), "ambiguous_owner_cycle_cases":len(result["ambiguous"]),
                        "disposed_rejection_cases":len(result["retired"]), "skipped":result["skipped"], "diagnostics":str(args.diagnostics)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
