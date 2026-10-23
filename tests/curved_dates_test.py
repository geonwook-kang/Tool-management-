"""Curved gray-paper table recovery and fail-closed date evidence.

The source is a generated generic ruled table with visible ink markers in every
row. Production cropping, curve measurement, pixel shifts, binarization, parsing
and assessment run unchanged. Only worker responses are controlled; this is not
an OCR accuracy claim. actual_photo_ocr_test.py covers the supplied camera JPEG.
"""
from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path
import tempfile
from urllib.parse import urlsplit

from PIL import Image, ImageDraw, ImageFont
from playwright.sync_api import sync_playwright

from browser_test import arguments, serve_program
from torque_dates_test import SERIAL

WIDTH, HEIGHT = 1200, 1600
VALUES = ["2027.03.17", "2027.09.17", "2028.03.17"]


def fixture(path, rows=2, rules=True):
    image = Image.new("RGB", (WIDTH, HEIGHT), (205, 200, 195))
    draw = ImageDraw.Draw(image)
    draw.rectangle((300, 730, 1000, 1230), fill=(165, 165, 165))
    font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 46)
    for row in range(rows):
        draw.text((500, 817 + row * 100), VALUES[row], fill=(15, 15, 15), font=font)
        # Independent solid ink markers permit pixel coverage checks without
        # recognizing or repainting characters. Every row has a different width.
        draw.rectangle((760, 875 + row * 100, 779 + row * 10, 882 + row * 100), fill=(0, 0, 0))
    if rules:
        for row in range(rows + 1):
            draw.line((300, 800 + row * 100, 1000, 800 + row * 100), fill=(0, 0, 0), width=3)
    warped = Image.new("RGB", image.size, (205, 205, 205))
    for x in range(WIDTH):
        shift = round(0.00055 * (x - 650) ** 2)
        if shift < HEIGHT:
            warped.paste(image.crop((x, 0, x + 1, HEIGHT - shift)), (x, shift))
    warped.save(path)
    return path


HARNESS = r"""({rows,scenario,serial,values,width,height})=>{
 window.curveCalls=[];window.curveParameters={};
 const regions=new WeakMap(),draw=CanvasRenderingContext2D.prototype.drawImage;
 CanvasRenderingContext2D.prototype.drawImage=function(source,...a){
  const value=draw.call(this,source,...a);
  const from=regions.get(source)||(source.width===width&&source.height===height?{x0:0,y0:0,x1:width,y1:height}:null);
  if(from){let sx=0,sy=0,sw=source.width,sh=source.height,dx=0,dy=0,dw=sw,dh=sh;
   if(a.length===2)[dx,dy]=a;else if(a.length===4)[dx,dy,dw,dh]=a;else if(a.length===8)[sx,sy,sw,sh,dx,dy,dw,dh]=a;
   const ox=x=>from.x0+(sx+(x-dx)*sw/dw)/source.width*(from.x1-from.x0);
   const oy=y=>from.y0+(sy+(y-dy)*sh/dh)/source.height*(from.y1-from.y0);
   regions.set(this.canvas,{x0:ox(0),y0:oy(0),x1:ox(this.canvas.width),y1:oy(this.canvas.height)});
  }return value;
 };
 function inspect(canvas){
  const pixels=canvas.getContext('2d').getImageData(0,0,canvas.width,canvas.height).data;
  let binary=true,black=0;for(let i=0;i<pixels.length;i+=4){if(pixels[i]!==0&&pixels[i]!==255)binary=false;if(pixels[i]<40)black++;}
  const markerWidths=[],w=canvas.width,h=canvas.height,mask=new Uint8Array(w*h),queue=new Int32Array(w*h);
  if(binary){
   for(let at=0;at<mask.length;at++)mask[at]=pixels[at*4]<40?1:0;
   for(let start=0;start<mask.length;start++){if(!mask[start])continue;let size=1,k=0,x0=w,x1=0,y0=h,y1=0;mask[start]=0;queue[0]=start;
    while(k<size){let at=queue[k++],x=at%w,y=Math.floor(at/w);x0=Math.min(x0,x);x1=Math.max(x1,x);y0=Math.min(y0,y);y1=Math.max(y1,y);
     for(let next of [x>0?at-1:-1,x<w-1?at+1:-1,y>0?at-w:-1,y<h-1?at+w:-1])if(next>=0&&mask[next]){mask[next]=0;queue[size++]=next;}}
    const cw=x1-x0+1,ch=y1-y0+1;
    if(cw>=18&&cw<=43&&ch>=7&&ch<=12&&size/(cw*ch)>.85)markerWidths.push(cw);
   }
  }
  return {width:w,height:h,binary,black,markerWidths:markerWidths.sort((a,b)=>a-b),roi:regions.get(canvas),parameters:{...curveParameters}};
 }
 const word=(text,entry,row,confidence=95)=>({text,confidence,bbox:{
  x0:(500-entry.roi.x0)/(entry.roi.x1-entry.roi.x0)*entry.width,
  x1:(680-entry.roi.x0)/(entry.roi.x1-entry.roi.x0)*entry.width,
  y0:(830+row*100-entry.roi.y0)/(entry.roi.y1-entry.roi.y0)*entry.height,
  y1:(894+row*100-entry.roi.y0)/(entry.roi.y1-entry.roi.y0)*entry.height}});
 let curveReads=0;
 Tesseract.createWorker=async()=>({setParameters:async p=>Object.assign(curveParameters,p),terminate:()=>{},
  recognize:async canvas=>{
   const entry=inspect(canvas),index=curveCalls.length;curveCalls.push(entry);
   if(index===0)return {data:{text:serial+'\n2027.0',confidence:95,words:[{text:serial,confidence:95},word('2027.0',entry,0)]}};
   if(entry.binary&&entry.markerWidths.length===rows){
    curveReads++;const confidence=scenario==='weak'&&curveReads===2?35:95;
    return {data:{text:values.slice(0,rows).join('\n'),confidence,words:values.slice(0,rows).map(text=>({text,confidence}))}};
   }
   if(index===1){
    if(scenario==='probe_only')return {data:{text:values.slice(0,2).join('\n'),confidence:95,words:values.slice(0,2).map((text,row)=>word(text,entry,row))}};
    return {data:{text:values.slice(0,rows).map(value=>value.slice(0,6)+' '+value.slice(6)).join('\n'),confidence:95,
     words:values.slice(0,rows).map((value,row)=>word(value.slice(0,6),entry,row))}};
   }
   return {data:{text:'',confidence:0,words:[]}};
  }});
}"""


def run(page, url, workbook, photo, rows, scenario):
    page.goto(url)
    page.wait_for_function("window.ToolApp&&window.ToolOCR")
    page.locator("#workbook").set_input_files(workbook)
    page.wait_for_function("ToolApp.state.model&&!ToolApp.state.busy")
    page.evaluate("()=>{const e=document.createElement('input');e.type='file';e.id='curve-fixture';document.body.append(e)}")
    page.locator("#curve-fixture").set_input_files(photo)
    page.evaluate(HARNESS, {"rows": rows, "scenario": scenario, "serial": SERIAL, "values": VALUES, "width": WIDTH, "height": HEIGHT})
    result = page.evaluate("""async()=>{
      const r=await ToolOCR.recognize(document.querySelector('#curve-fixture').files[0],ToolApp.state.model.tools);
      return {assessment:ToolAutomation.assess(ToolApp.state.model,r),dates:r.dateEvidence,
        failed:r.failed,passes:r.passes.map(p=>({source:p.source,text:p.text})),calls:curveCalls};
    }""")
    sources = [item["source"] for item in result["passes"]]
    curved = [result["calls"][i] for i, source in enumerate(sources) if source.startswith("date-crop-curve-")]
    if scenario in {"normal", "three_dates", "weak"}:
        assert len(curved) == 2, result
        assert all(item["binary"] and len(item["markerWidths"]) == rows for item in curved), result
        assert all(item["markerWidths"] == [20, 30, 40][:rows] for item in curved), result
    else:
        assert not curved, result
    assert not result["failed"], result
    if scenario == "normal":
        assert result["assessment"]["eligible"], result
        assert result["assessment"]["serial"] == SERIAL, result
        assert (result["assessment"]["last"], result["assessment"]["due"]) == ("2027-03-17", "2027-09-17"), result
        assert result["dates"]["method"] == "repeated-dates" and result["dates"]["repeatCount"] == 2, result
    else:
        assert not result["assessment"]["eligible"], result
        if scenario == "three_dates":
            assert result["dates"]["dates"] == [value.replace(".", "-") for value in VALUES], result
        if scenario == "weak":
            assert result["dates"]["repeatCount"] == 1, result
        if scenario == "probe_only":
            assert result["dates"]["dates"] == [], result
    return {"eligible": result["assessment"]["eligible"], "preserved_row_markers": [item["markerWidths"] for item in curved],
            "date_evidence": result["dates"], "pass_sources": sources}


def main():
    args = arguments(__doc__)
    original_hash = sha256(args.workbook.read_bytes()).hexdigest()
    errors, remote = [], []
    with tempfile.TemporaryDirectory(prefix="curved-date-policy-") as directory, serve_program(args.program) as url, sync_playwright() as playwright:
        folder = Path(directory)
        fixtures = {"two": fixture(folder / "two.png"), "three": fixture(folder / "three.png", rows=3), "unruled": fixture(folder / "unruled.png", rules=False)}
        browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
        outcomes = {}
        for scenario, rows, photo in [("normal",2,"two"),("three_dates",3,"three"),("weak",2,"two"),("no_rules",2,"unruled"),("probe_only",2,"unruled")]:
            context = browser.new_context(locale="ko-KR")
            page = context.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))
            context.on("request",lambda request: remote.append(request.url) if urlsplit(request.url).scheme in {"http","https"} and urlsplit(request.url).netloc!=urlsplit(url).netloc else None)
            outcomes[scenario] = run(page,url,args.workbook,fixtures[photo],rows,scenario)
            context.close()
        browser.close()
    assert not errors and not remote, (errors,remote)
    assert sha256(args.workbook.read_bytes()).hexdigest() == original_hash
    print(json.dumps({"cases":outcomes,"company_workbook_unchanged":True,"page_errors":errors,"external_requests":remote,
        "accuracy_scope":"Synthetic actual pixels and controlled OCR; real camera accuracy tested separately."},ensure_ascii=False,indent=2))


if __name__ == "__main__":
    main()
