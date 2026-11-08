"""A lone clear date row must be confirmed by reading its actual pixels again.

The worker replies and region proposals are controlled. Production image
cropping, date parsing, source promotion and automatic approval run unchanged.
Colored row markers and different dark paper/ink levels prove that focused
inputs contain source row pixels, including independently thresholded strips.
This test does not measure camera OCR accuracy or company storage behavior.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import tempfile
from urllib.parse import urlsplit

from playwright.sync_api import sync_playwright

from browser_test import REPO, serve_program


FIXTURE = r"""scenario => {
 window.testReads=[];window.dateCalls=0;window.tokenCalls=0;window.binaryCalls=0;
 window.ToolRuntime={ready:async()=>({workerSource:'/* controlled date proof test */',coreSource:'function TesseractCore(){}',languages:{kor:'fixture',eng:'fixture'}})};
 window.testTools=[{id:'tool-a',name:'홍길동',serial:'6492873150',months:12,interval:12,cycle:'MONTHS',last:'2026-01-14',due:'2027-01-14',photoCount:1}];
 const values=['2027.01.14','2028.01.14'],dark=scenario.startsWith('binary-');
 const summary=input=>{
  const pixels=input.getContext('2d').getImageData(0,0,input.width,input.height).data;let red=0,blue=0,black=0,white=0;
  const mask=new Uint8Array(input.width*input.height),components=[];
  for(let at=0;at<pixels.length;at+=4){if(pixels[at]>190&&pixels[at+1]<80&&pixels[at+2]<80)red++;if(pixels[at]<80&&pixels[at+1]<80&&pixels[at+2]>190)blue++;
   if(pixels[at]<20&&pixels[at+1]<20&&pixels[at+2]<20){black++;mask[at/4]=1;}
   if(pixels[at]===255&&pixels[at+1]===255&&pixels[at+2]===255)white++;
  }
  for(let start=0;start<mask.length;start++)if(mask[start]){
   const queue=[start];mask[start]=0;let x0=input.width,y0=input.height,x1=0,y1=0,count=0;
   for(let index=0;index<queue.length;index++){const at=queue[index],x=at%input.width,y=Math.floor(at/input.width);count++;x0=Math.min(x0,x);x1=Math.max(x1,x);y0=Math.min(y0,y);y1=Math.max(y1,y);
    for(const next of [x?at-1:-1,x+1<input.width?at+1:-1,y?at-input.width:-1,y+1<input.height?at+input.width:-1])if(next>=0&&mask[next]){mask[next]=0;queue.push(next);}
   }components.push({x0,y0,x1,y1,count});
  }
  return {width:input.width,height:input.height,red,blue,black,white,components,bilevel:black+white===input.width*input.height};
 };
 window.Tesseract={createWorker:async()=>{
  let parameters={};return{setParameters:async p=>{parameters={...parameters,...p};},terminate:()=>{},recognize:async input=>{
   const entry={...summary(input),parameters:{...parameters}};testReads.push(entry);
   if(input.width===240&&input.height===100)return{data:{text:testTools[0].serial,confidence:96,words:[]}};
   if(input.width===500&&input.height===300){
    dateCalls++;
    if(dateCalls!==1)return{data:{text:'',confidence:0,words:[]}};
    const weakLocator=scenario.startsWith('weak-first')||['binary-weak-first','binary-one-proof','binary-low-second','binary-zero-second','binary-partial-second','binary-low-word-second'].includes(scenario);
    const confidence=weakLocator?35:96;
    const words=scenario==='missing-geometry'?[]:values.map((text,index)=>({text,confidence,bbox:{x0:100,y0:50+index*120,x1:380,y1:90+index*120}}));
    return{data:{text:values.join('\n'),confidence,words}};
   }
   // Only a real focused conversion retaining all four unequal source ink
   // markers can produce a binary proof. Blank canvases, global thresholds
   // that swallow a dark paper row, and the colored fixture cannot do so.
   if(dark&&input.width<800&&input.width!==500&&entry.bilevel&&entry.components.length===4){
    binaryCalls++;
    if(scenario==='binary-no-second'||scenario==='binary-one-proof'&&binaryCalls>1)return{data:{text:'',confidence:0,words:[]}};
    if(scenario==='binary-partial-second'&&binaryCalls>1)return{data:{text:values[0],confidence:96,words:[]}};
    const text=scenario==='binary-contradictory'?'2027.01.14\n2028.01.15':values.join('\n');
    const confidence=binaryCalls>1&&scenario==='binary-low-second'?35:binaryCalls>1&&scenario==='binary-zero-second'?0:96;
    const words=binaryCalls>1&&scenario==='binary-low-word-second'?values.map(text=>({text,confidence:20})):[];
    return{data:{text,confidence,words}};
   }
   if(entry.red>100&&entry.blue>100&&input.width<800){
    tokenCalls++;
    if(scenario==='weak-first-one-proof'&&tokenCalls>1)return{data:{text:'',confidence:0,words:[]}};
    if(scenario==='weak-first-low-second'&&tokenCalls>1)return{data:{text:values.join('\n'),confidence:35,words:[]}};
    if(scenario==='empty-second')return{data:{text:'',confidence:0,words:[]}};
    if(scenario==='one-date-second')return{data:{text:values[0],confidence:96,words:[]}};
    const text=scenario==='contradictory-second'?'2027.01.14\n2028.01.15':values.join('\n');
    return{data:{text,confidence:scenario==='weak-second'?35:96,words:[]}};
   }
   return{data:{text:'',confidence:0,words:[]}};
  }};
 }};
 window.ToolFastOCR={regions:()=>[
  {kind:'tag',bbox:{x0:100,y0:100,x1:340,y1:200}},
  {kind:'dates',bbox:{x0:100,y0:400,x1:600,y1:700},cells:[{x0:100,y0:400,x1:600,y1:480},{x0:100,y0:500,x1:600,y1:580},{x0:100,y0:600,x1:600,y1:680}]}],
  prepare:(_canvas,region)=>{
   const input=document.createElement('canvas');input.width=region.kind==='tag'?240:500;input.height=region.kind==='tag'?100:300;
   const ctx=input.getContext('2d');ctx.fillStyle='#bbb';ctx.fillRect(0,0,input.width,input.height);
   if(region.kind==='dates'){
    if(dark){
     ctx.fillStyle='rgb(80,80,80)';ctx.fillRect(88,42,304,56);
     ctx.fillStyle='rgb(20,20,20)';ctx.fillRect(120,60,12,20);ctx.fillRect(180,60,20,20);
     ctx.fillStyle='rgb(180,180,180)';ctx.fillRect(88,162,304,56);
     ctx.fillStyle='rgb(120,120,120)';ctx.fillRect(130,180,16,20);ctx.fillRect(260,180,30,20);
    }else{ctx.fillStyle='rgb(225,25,25)';ctx.fillRect(100,50,280,40);ctx.fillStyle='rgb(25,25,225)';ctx.fillRect(100,170,280,40);}
   }
   return input;
  }};
 window.testPhoto=async()=>{
  const c=document.createElement('canvas');c.width=1800;c.height=2400;
  const ctx=c.getContext('2d');ctx.fillStyle='#bbb';ctx.fillRect(0,0,c.width,c.height);
  return new Promise(resolve=>c.toBlob(resolve,'image/png'));
 };
}"""


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--chromium', default='/usr/bin/chromium')
    args = parser.parse_args()
    cases, errors, external = {}, [], []
    scenarios = ['valid', 'weak-second', 'one-date-second', 'contradictory-second',
                 'empty-second', 'missing-geometry', 'weak-first',
                 'weak-first-one-proof', 'weak-first-low-second',
                 'binary-clear', 'binary-weak-first', 'binary-one-proof',
                 'binary-low-second', 'binary-zero-second', 'binary-partial-second',
                 'binary-low-word-second', 'binary-contradictory', 'binary-no-second']
    with tempfile.TemporaryDirectory(prefix='date-token-proof-') as directory:
        fixture = Path(directory) / 'fixture.html'
        fixture.write_text('<!doctype html><title>Date token confirmation fixture</title>')
        with serve_program(fixture) as url, sync_playwright() as pw:
            browser = pw.chromium.launch(executable_path=args.chromium, headless=True, args=['--no-sandbox'])
            for scenario in scenarios:
                page = browser.new_page()
                page.on('pageerror', lambda error: errors.append(str(error)))
                page.on('request', lambda request: external.append(request.url)
                        if urlsplit(request.url).scheme in {'http', 'https'} and urlsplit(request.url).netloc != urlsplit(url).netloc else None)
                page.goto(url)
                page.evaluate(FIXTURE, scenario)
                for source in ['web/automation.js', 'web/ocr.js']:
                    page.add_script_tag(path=REPO / source)
                result = page.evaluate(r"""async()=>{
                  const parsed=await ToolOCR.recognize(await testPhoto(),testTools);
                  return {assessment:ToolAutomation.assess({tools:testTools},parsed),sources:parsed.passes.map(p=>p.source),
                   failed:parsed.failed,tokenCalls,binaryCalls,dateCalls,reads:testReads,dates:parsed.dateEvidence};
                }""")
                if scenario in {'valid', 'weak-first'}:
                    assert result['assessment']['eligible'] and not result['failed'], result
                    assert result['assessment']['last'] == '2027-01-14' and result['assessment']['due'] == '2028-01-14', result
                    expected_reads = 1 if scenario == 'valid' else 2
                    assert result['tokenCalls'] == expected_reads, result
                    assert 'date-crop-fast-token-6' in result['sources'], result
                    if scenario == 'weak-first':
                        assert 'date-probe-fast-gray' in result['sources'], result
                        assert 'date-crop-fast-gray' not in result['sources'], result
                        assert 'date-crop-fast-token-11' in result['sources'], result
                    assert 'original' not in result['sources'], result
                    focused = [row for row in result['reads'] if row['red'] > 100 and row['blue'] > 100 and row['width'] != 500]
                    assert len(focused) == expected_reads and focused[0]['parameters']['tessedit_pageseg_mode'] == '6', result
                    if scenario == 'weak-first':
                        assert focused[1]['parameters']['tessedit_pageseg_mode'] == '11', result
                    assert result['dates']['repeatCount'] == 2, result
                elif scenario.startswith('binary-'):
                    eligible = scenario in {'binary-clear', 'binary-weak-first'}
                    assert result['assessment']['eligible'] == eligible, (scenario, result)
                    binary = [row for row in result['reads'] if row['width'] != 500
                              and row['bilevel'] and len(row['components']) == 4]
                    assert result['binaryCalls'] == len(binary) and binary, (scenario, result)
                    assert [row['parameters']['tessedit_pageseg_mode'] for row in binary] == ['6', '11'][:len(binary)], (scenario, result)
                    # Ink markers keep their distinct widths, all share one
                    # height, and paper between/around them becomes white.
                    for row in binary:
                        components = sorted(row['components'], key=lambda part: (part['y0'], part['x0']))
                        assert [part['x1']-part['x0']+1 for part in components] == [18, 30, 24, 45], (scenario, row)
                        assert {part['y1']-part['y0']+1 for part in components} == {30}, (scenario, row)
                        assert len({part['y0'] for part in components}) == 2 and row['white'] > row['black'] * 10, (scenario, row)
                    if eligible:
                        assert result['dates']['repeatCount'] == 2, result
                        assert 'date-crop-fast-token-binary-6' in result['sources'], result
                        assert 'original' not in result['sources'], result
                        assert result['assessment']['last'] == '2027-01-14' and result['assessment']['due'] == '2028-01-14', result
                        assert result['binaryCalls'] == (1 if scenario == 'binary-clear' else 2), result
                        if scenario == 'binary-weak-first':
                            assert 'date-probe-fast-gray' in result['sources'] and 'date-crop-fast-gray' not in result['sources'], result
                            assert 'date-crop-fast-token-binary-11' in result['sources'], result
                    else:
                        assert result['assessment']['reasons'], result
                else:
                    assert not result['assessment']['eligible'], result
                    if scenario == 'missing-geometry':
                        assert result['tokenCalls'] == 0, result
                    else:
                        assert result['tokenCalls'] >= 1, result
                cases[scenario] = {key: result[key] for key in ['assessment', 'sources', 'tokenCalls', 'binaryCalls', 'dates']}
                page.close()
            browser.close()
    assert not errors and not external, (errors, external)
    print(json.dumps({'case_count': len(cases), 'cases': cases, 'page_errors': errors,
                      'external_requests': external,
                      'scope': 'Controlled OCR/proposals; actual focused row pixels and independent date proof'},
                     ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
