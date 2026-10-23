"""Actual canvas coverage and OCR retry orchestration for difficult label ROIs.

Synthetic pictures contain located date/tag markers and white distractors.
ToolOCR.recognize, cropping, parsing and assessment run unchanged. Only worker
responses are controlled. The tests measure ROI coverage/retry safety, not OCR
accuracy on the user's unavailable inline JPEG. Company data remains read only.
"""
from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path
import tempfile
from time import monotonic
from urllib.parse import urlsplit

from PIL import Image, ImageDraw
from playwright.sync_api import sync_playwright

from browser_test import arguments, serve_program
from torque_dates_test import SERIAL, LAST, DUE, LABELS


WIDTH, HEIGHT = 1200, 1600
DATE_POINTS = [[825, 1125], [1105, 1375]]
TAG_POINTS = [[810, 110], [1110, 250]]


def make_picture(path, white_holes=False, table_borders=False):
    image = Image.new("RGB", (WIDTH, HEIGHT), (140, 110, 70))
    draw = ImageDraw.Draw(image)
    draw.rectangle((800, 100, 1120, 260), fill=(110, 150, 185))
    draw.text((835, 165), SERIAL, fill="black", font_size=35)
    draw.text((300, 900), "Calibration Label", fill="black", font_size=26)
    if table_borders:
        draw.rectangle((760, 1110, 1140, 1380), fill="white")
        draw.rectangle((790, 1110, 792, 1380), fill="black")
        draw.rectangle((1134, 1110, 1136, 1380), fill="black")
    draw.text((850, 1130), LAST.replace("-", "."), fill="black", font_size=34)
    draw.text((850, 1320), DUE.replace("-", "."), fill="black", font_size=34)
    for x, y in DATE_POINTS + TAG_POINTS:
        draw.rectangle((x - 4, y - 4, x + 4, y + 4), fill="black")
    if white_holes:
        # These are image components, not owner tags. Their presence used to
        # suppress the independent upper-image bands altogether.
        for x, y, width in [(100, 450, 80), (300, 520, 80), (500, 590, 51), (700, 660, 51)]:
            draw.rectangle((x, y, x + width, y + 20), fill="white")
    image.save(path)
    return path


HARNESS = r"""({scenario,serial,labels,width,height,datePoints,tagPoints}) => {
    window.roiCalls=[];window.roiParameters=[];window.roiTerminations=0;window.roiHung=false;
    const mappings=new WeakMap(),nativeDraw=CanvasRenderingContext2D.prototype.drawImage;
    // This instrumentation observes the real drawImage source rectangles. It
    // does not change pixels, crop bounds, recognition inputs or production code.
    CanvasRenderingContext2D.prototype.drawImage=function(source,...args){
        const output=nativeDraw.call(this,source,...args);
        const from=mappings.get(source)||(source.width===width&&source.height===height?
            {x0:0,y0:0,x1:width,y1:height}:null);
        if(from){
            let sx=0,sy=0,sw=source.width,sh=source.height,dx=0,dy=0,dw=sw,dh=sh;
            if(args.length===2)[dx,dy]=args;
            else if(args.length===4)[dx,dy,dw,dh]=args;
            else if(args.length===8)[sx,sy,sw,sh,dx,dy,dw,dh]=args;
            const ox=x=>from.x0+(sx+(x-dx)*sw/dw)/source.width*(from.x1-from.x0);
            const oy=y=>from.y0+(sy+(y-dy)*sh/dh)/source.height*(from.y1-from.y0);
            mappings.set(this.canvas,{x0:ox(0),y0:oy(0),x1:ox(this.canvas.width),y1:oy(this.canvas.height)});
        }
        return output;
    };
    const within=(point,roi)=>roi&&point[0]>=roi.x0&&point[0]<=roi.x1&&point[1]>=roi.y0&&point[1]<=roi.y1;
    const inspect=canvas=>{
        const roi=mappings.get(canvas),context=canvas.getContext('2d'),pixels=context.getImageData(0,0,canvas.width,canvas.height).data;
        let hash=2166136261;
        for(let i=0;i<pixels.length;i++)hash=Math.imul(hash^pixels[i],16777619)>>>0;
        const blackMarkers=points=>points.filter(point=>{
            if(!within(point,roi))return false;
            const x=Math.min(canvas.width-1,Math.max(0,Math.floor((point[0]-roi.x0)/(roi.x1-roi.x0)*canvas.width)));
            const y=Math.min(canvas.height-1,Math.max(0,Math.floor((point[1]-roi.y0)/(roi.y1-roi.y0)*canvas.height)));
            const at=(y*canvas.width+x)*4;
            return (pixels[at]+pixels[at+1]+pixels[at+2])/3<45;
        }).length;
        return {width:canvas.width,height:canvas.height,roi,hash,
            dateMarkers:blackMarkers(datePoints),tagMarkers:blackMarkers(tagPoints),
            parameters:{...(roiParameters.at(-1)||{})}};
    };
    const box=(rect,entry)=>({
        x0:(rect[0]-entry.roi.x0)/(entry.roi.x1-entry.roi.x0)*entry.width,
        y0:(rect[1]-entry.roi.y0)/(entry.roi.y1-entry.roi.y0)*entry.height,
        x1:(rect[2]-entry.roi.x0)/(entry.roi.x1-entry.roi.x0)*entry.width,
        y1:(rect[3]-entry.roi.y0)/(entry.roi.y1-entry.roi.y0)*entry.height});
    const data=(text,confidence=95,words=[])=>({data:{text,confidence,words}});
    const dateResponse=(entry,discovery=false,confidence=95)=>{
        const values=discovery?['2026.08.22','2026.12.22']:['2026.06.22','2026.12.22'];
        return data(values.join('\n'),confidence,[
            {text:values[0],confidence,bbox:box([850,1130,1100,1170],entry)},
            {text:values[1],confidence,bbox:box([850,1320,1100,1360],entry)}]);
    };
    let focusedReads=0,tagFound=false;
    Tesseract.createWorker=async()=>({
        setParameters:async parameters=>{roiParameters.push({...parameters});},
        recognize:async canvas=>{
            const entry=inspect(canvas),index=roiCalls.length;roiCalls.push(entry);
            if(scenario==='date_column_priority'){
                if(index===0)return data('Serial: '+serial);
                if(index===1)return data('교정일자\n차기교정예정일자 2026.12.22',95,[
                    {text:'교정일자',confidence:95,bbox:box([600,1130,760,1170],entry)},
                    {text:'차기교정예정일자',confidence:95,bbox:box([600,1320,760,1360],entry)},
                    {text:'2026.12.22',confidence:95,bbox:box([850,1320,1100,1360],entry)}]);
                // Only an actual crop containing both date locations and
                // excluding the two printed table rules can read this pair.
                if(entry.dateMarkers===2&&entry.roi.x0>792&&entry.roi.x1<1134)
                    return dateResponse(entry);
                return data('',0);
            }
            if(scenario.startsWith('date_')){
                if(index===0)return data('Serial: '+serial+'\nCalibration Label\n2026.',95,[
                    {text:serial,confidence:95,bbox:box([80,100,350,135],entry)},
                    {text:'Calibration Label',confidence:80,bbox:box([300,900,450,930],entry)},
                    {text:'2026.',confidence:20,bbox:box([330,900,420,930],entry)}]);
                if(index<=2)return data('2026.06',20);
                if(entry.dateMarkers!==2)return data('',0);
                const broad=(entry.roi.x1-entry.roi.x0)*(entry.roi.y1-entry.roi.y0)>width*height*0.4;
                if(broad)return dateResponse(entry,true,80);
                focusedReads++;
                return dateResponse(entry,false,scenario==='date_one_weak_confirmation'&&focusedReads===2?25:95);
            }
            if(index<2)return data(labels);
            if(index<4)return data('',0);
            if(tagFound&&(scenario==='tag_timeout'||scenario==='tag_cancel')){
                window.roiHung=true;
                if(scenario==='tag_timeout'){
                    // Accelerate this one real deadline after initialization.
                    // No recognition result or error classification is mocked.
                    const nativeTimeout=window.setTimeout;
                    window.setTimeout=(callback,ms,...rest)=>nativeTimeout(callback,ms>=10000?25:ms,...rest);
                }
                return new Promise(()=>{});
            }
            if(entry.tagMarkers===2){tagFound=true;return data(serial);}
            return data('',0);
        },
        terminate:()=>{roiTerminations++;}
    });
}"""


def prepare(page, url, workbook, photo, scenario):
    page.goto(url)
    page.wait_for_function("window.ToolApp&&!ToolApp.state.busy")
    page.locator("#workbook").set_input_files(workbook)
    page.wait_for_function("ToolApp.state.model&&!ToolApp.state.busy")
    page.evaluate(HARNESS, {"scenario": scenario, "serial": SERIAL, "labels": LABELS,
                           "width": WIDTH, "height": HEIGHT, "datePoints": DATE_POINTS, "tagPoints": TAG_POINTS})
    page.evaluate("() => {const input=document.createElement('input');input.type='file';input.id='roi-fixture';document.body.append(input);}")
    page.locator("#roi-fixture").set_input_files(photo)


READ_RESULT = """async () => {
    const parsed=await ToolOCR.recognize(document.querySelector('#roi-fixture').files[0],ToolApp.state.model.tools);
    const replay=ToolOCR.parse(parsed.text,ToolApp.state.model.tools,{confidence:parsed.confidence,passes:parsed.passes});
    return {failed:parsed.failed,cancelled:parsed.cancelled,error:parsed.error,code:parsed.code,
        incomplete:parsed.identityRecoveryIncomplete,replayIncomplete:replay.identityRecoveryIncomplete,
        last:parsed.last,due:parsed.due,dates:parsed.dateEvidence,
        assessment:ToolAutomation.assess(ToolApp.state.model,parsed),calls:roiCalls,terminations:roiTerminations,
        passes:parsed.passes.map(pass=>({source:pass.source,text:pass.text}))};
}"""


def date_recovery(page, url, workbook, photo, weak=False):
    prepare(page, url, workbook, photo, "date_one_weak_confirmation" if weak else "date_wide")
    result = page.evaluate(READ_RESULT)
    assert not result["failed"], result
    sources = [p["source"] for p in result["passes"]]
    assert "date-probe-wide" in sources, result
    wide_index = sources.index("date-probe-wide")
    assert result["calls"][wide_index]["dateMarkers"] == 2, result
    assert result["calls"][wide_index]["roi"] != result["calls"][1]["roi"], result
    focused = [result["calls"][i] for i, source in enumerate(sources) if source.startswith("date-crop-wide-refine-")]
    assert len(focused) == 2, result
    assert all(call["dateMarkers"] == 2 for call in focused), result
    assert focused[0]["roi"] == focused[1]["roi"], result
    assert focused[0]["hash"] == focused[1]["hash"], "Independent OCR passes must use identical, unchanged focused pixels"
    assert "2026-08-22" not in result["dates"]["dates"], result
    if weak:
        assert not result["assessment"]["eligible"], result
        return {"wide_and_focused_rois_cover_both_date_markers": True,
                "one_weak_focused_confirmation_remains_pending": True}
    assert result["assessment"]["eligible"], result
    assert result["assessment"]["last"] == LAST and result["assessment"]["due"] == DUE, result
    assert result["assessment"]["serial"] == SERIAL, result
    return {"bad_primary_probe_retries_independent_wide_roi": True,
            "both_focused_reads_cover_actual_date_pixels": True, "focused_pixels_unchanged": True,
            "exploratory_wrong_month_is_not_date_evidence": True, "worker_calls": len(result["calls"])}


def column_priority_recovery(page, url, workbook, photo):
    prepare(page, url, workbook, photo, "date_column_priority")
    result = page.evaluate(READ_RESULT)
    assert not result["failed"], result
    sources = [p["source"] for p in result["passes"]]
    assert "date-probe" in sources and "date-probe-wide" not in sources, result
    focused = [result["calls"][i] for i, source in enumerate(sources)
               if source.startswith("date-crop-refine-")]
    assert len(focused) == 2, result
    assert [call["parameters"]["tessedit_pageseg_mode"] for call in focused] == ["6", "11"], result
    assert all(call["dateMarkers"] == 2 for call in focused), result
    assert all(call["roi"]["x0"] > 792 and call["roi"]["x1"] < 1134 for call in focused), result
    assert focused[0]["roi"] == focused[1]["roi"] and focused[0]["hash"] == focused[1]["hash"], result
    assert result["assessment"]["eligible"], result
    assert result["assessment"]["last"] == LAST and result["assessment"]["due"] == DUE, result
    return {"single_probe_date_and_two_role_rows_preserve_both_actual_date_locations": True,
            "long_table_rules_trimmed_before_psm6_and_psm11_confirmation": True,
            "focused_pixel_hashes_identical": True, "wide_search_not_needed": True}


def tag_recovery(page, url, workbook, photo, scenario):
    prepare(page, url, workbook, photo, scenario)
    if scenario == "tag_cancel":
        page.evaluate("script => {window.roiPending=(0,eval)('('+script+')')();}", READ_RESULT)
        page.wait_for_function("window.roiHung", timeout=15000)
        began = monotonic()
        page.evaluate("ToolOCR.cancel()")
        result = page.evaluate("async()=>await roiPending")
        assert monotonic() - began < 5
        assert result["failed"] and result["cancelled"] and result["code"] == "cancelled", result
        assert not result["assessment"]["eligible"], result
        assert result["last"] == LAST and result["due"] == DUE, result
        assert result["terminations"] >= 1, result
        return {"cancel_stops_hung_optional_roi_retry": True, "successful_dates_preserved_without_approval": True}
    result = page.evaluate(READ_RESULT)
    assert not result["failed"], result
    sources = [p["source"] for p in result["passes"]]
    panel_calls = [result["calls"][i] for i, source in enumerate(sources) if source.startswith("identity-probe-panel-")]
    assert 2 <= len(panel_calls) <= 4, result
    assert any(call["tagMarkers"] == 2 for call in panel_calls), result
    assert result["last"] == LAST and result["due"] == DUE, result
    if scenario == "tag_timeout":
        assert result["incomplete"] and result["replayIncomplete"], result
        assert not result["assessment"]["eligible"], result
        assert result["terminations"] >= 1, result
        return {"optional_roi_deadline_preserves_dates": True, "incomplete_recovery_cannot_be_approved": True}
    assert result["assessment"]["eligible"] and result["assessment"]["serial"] == SERIAL, result
    return {"white_components_do_not_disable_independent_upper_bands": True,
            "upper_band_contains_actual_tag_marker_pixels": True, "paper_and_band_candidates": len(panel_calls)}


def main():
    args = arguments(__doc__)
    original_hash = sha256(args.workbook.read_bytes()).hexdigest()
    errors, remote = [], []
    with tempfile.TemporaryDirectory(prefix="tool-roi-recovery-") as directory, serve_program(args.program) as url, sync_playwright() as playwright:
        temporary = Path(directory)
        dates = make_picture(temporary / "date-location-fixture.png")
        table = make_picture(temporary / "date-table-fixture.png", table_borders=True)
        tags = make_picture(temporary / "white-components-fixture.png", white_holes=True)
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
        result = {}
        for scenario in ["date_column_priority", "date_wide", "date_one_weak_confirmation", "tag_wide", "tag_timeout", "tag_cancel"]:
            context = browser.new_context(locale="ko-KR")
            context.on("request", lambda request: remote.append(request.url)
                if urlsplit(request.url).scheme in {"http", "https"} and urlsplit(request.url).netloc != urlsplit(url).netloc else None)
            page = context.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))
            page.on("dialog", lambda dialog: dialog.accept())
            if scenario == "date_column_priority":
                result[scenario] = column_priority_recovery(page, url, args.workbook, table)
            elif scenario.startswith("date_"):
                result[scenario] = date_recovery(page, url, args.workbook, dates, scenario.endswith("confirmation"))
            else:
                result[scenario] = tag_recovery(page, url, args.workbook, tags, scenario)
            context.close()
        assert not errors, errors
        assert not remote, remote
        browser.close()
    assert sha256(args.workbook.read_bytes()).hexdigest() == original_hash
    result.update({"original_company_workbook_unchanged": True, "page_errors": errors, "external_network_requests": remote,
                   "actual_inline_photo_accuracy": "not measured: synthetic images and controlled worker responses"})
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
