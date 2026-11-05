"""Bound repeated OCR work after identity/date evidence is already sufficient.

This is a recognition-orchestration test, not an OCR-accuracy benchmark. Only
the worker responses and pixel region proposals are controlled fixtures. The
production image loading, parser, evidence gates, automatic resolver, worker
lifetime and recognition pipeline run unchanged. Real camera photographs are
covered separately by multi_owner_ocr_test.py and the PC final photo tests.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import tempfile
from urllib.parse import urlsplit

from playwright.sync_api import sync_playwright

from browser_test import REPO, serve_program


FIXTURE = r"""() => {
  window.testReads=[];window.testWorkers=0;window.testTerminations=0;window.testReplies=[];window.testReplyForInput=null;
  window.testClockStep=0;window.testClockElapsed=0;window.testClockOriginalNow=null;window.testClockReject=false;
  window.ToolRuntime={ready:async()=>({workerSource:'/* controlled orchestration fixture */',coreSource:'function TesseractCore(){}',languages:{kor:'fixture',eng:'fixture'}})};
  window.Tesseract={createWorker:async()=>{
    testWorkers++;let parameters={};
    return {setParameters:async value=>{parameters={...parameters,...value};},terminate:()=>{testTerminations++;},recognize:async input=>{
      if(testClockStep){
        if(!testClockOriginalNow){testClockOriginalNow=Date.now;const start=Date.now();Date.now=()=>start+testClockElapsed;}
        testClockElapsed+=testClockStep;
      }
      const reply=testReplyForInput?testReplyForInput(input):testReplies.length?testReplies.shift():{text:'',confidence:95};
      testReads.push({width:input.width,height:input.height,parameters:{...parameters},text:reply.text});
      if(testClockStep&&testClockReject&&testClockElapsed>=35000)throw new Error('Controlled pending operation rejected at budget exhaustion');
      return {data:{text:reply.text,confidence:reply.confidence??95,words:[]}};
    }};
  }};
  window.testTools=[{id:'tool-a',name:'홍길동',serial:'8765432109',months:12,interval:12,cycle:'MONTHS',last:'2026-03-17',due:'2027-03-17',photoCount:1},
    {id:'tool-b',name:'김하늘',serial:'4198765023',months:12,interval:12,cycle:'MONTHS',last:'2026-03-17',due:'2027-03-17',photoCount:1}];
  window.testDateReply={text:'2027.03.17\n2028.03.17',confidence:96};
  window.testPhoto=async large=>{
    const canvas=document.createElement('canvas');canvas.width=large?1800:600;canvas.height=large?2400:900;
    const context=canvas.getContext('2d');context.fillStyle='#bbb';context.fillRect(0,0,canvas.width,canvas.height);
    return await new Promise(resolve=>canvas.toBlob(resolve,'image/png'));
  };
  window.testQuickProposals=()=>{
    ToolFastOCR={regions:()=>[{kind:'tag',bbox:{x0:100,y0:100,x1:600,y1:300}},
      {kind:'tag',bbox:{x0:100,y0:400,x1:600,y1:600}},
      {kind:'dates',bbox:{x0:100,y0:800,x1:600,y1:1100},cells:[{x0:100,y0:800,x1:600,y1:880},{x0:100,y0:900,x1:600,y1:980},{x0:100,y0:1000,x1:600,y1:1080}]}],
      prepare:(_canvas,region)=>{const c=document.createElement('canvas');c.width=500;c.height=region.kind==='tag'?200:300;return c;}};
  };
  window.testSummary=result=>({eligible:ToolAutomation.assess({tools:testTools},result).eligible,
    assessment:ToolAutomation.assess({tools:testTools},result),sources:result.passes.map(pass=>pass.source),
    reads:testReads.length,workers:testWorkers,fastPath:!!result.fastPath,failed:result.failed});
}"""


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--chromium', default='/usr/bin/chromium')
    args = parser.parse_args()
    cases, errors, external = {}, [], []
    with tempfile.TemporaryDirectory(prefix='pc-speed-orchestration-') as directory:
        fixture = Path(directory) / 'fixture.html'
        fixture.write_text('<!doctype html><title>Recognition orchestration fixture</title>')
        with serve_program(fixture) as url, sync_playwright() as playwright:
            browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=['--no-sandbox'])

            def new_page():
                page = browser.new_page()
                page.on('pageerror', lambda error: errors.append(str(error)))
                page.on('request', lambda request: external.append(request.url)
                        if urlsplit(request.url).scheme in {'http', 'https'} and urlsplit(request.url).netloc != urlsplit(url).netloc else None)
                page.goto(url)
                page.evaluate(FIXTURE)
                for source in ['web/automation.js', 'web/ocr.js']:
                    page.add_script_tag(path=REPO / source)
                return page

            page = new_page()
            full = page.evaluate(r"""async()=>{
              testQuickProposals();testReplies=[{text:'8765432109'}, {text:''}, testDateReply, testDateReply];
              const result=await ToolOCR.recognize(await testPhoto(true),testTools);return testSummary(result);
            }""")
            assert full['eligible'] and full['fastPath'] and not full['failed'], full
            assert full['reads'] == 4 and 'original' not in full['sources'], full
            cases['serial_without_name_stops_after_valid_tag_and_dates'] = full
            warm = page.evaluate(r"""async()=>{
              const before=testReads.length;testReplies=[{text:'8765432109'},{text:''},testDateReply,testDateReply];
              const result=await ToolOCR.recognize(await testPhoto(true),testTools);return {...testSummary(result),newReads:testReads.length-before};
            }""")
            assert warm['eligible'] and warm['workers'] == 1 and warm['newReads'] == 4, warm
            cases['warm_photo_reuses_worker_and_same_bounded_read_count'] = warm
            page.close()

            page = new_page()
            partial = page.evaluate(r"""async()=>{
              testQuickProposals();let tags=0;testReplyForInput=input=>input.height===200?{text:tags++%2?'':'65432109'}:input.height===300?testDateReply:{text:''};
              const result=await ToolOCR.recognize(await testPhoto(true),testTools);return testSummary(result);
            }""")
            assert partial['eligible'] and partial['fastPath'] and not partial['failed'], partial
            assert partial['assessment']['matchMethod'] == 'partial-serial', partial
            assert partial['reads'] <= 6 and 'original' not in partial['sources'], partial
            cases['independent_partial_serial_proof_avoids_full_image_recovery'] = partial
            page.close()

            page = new_page()
            dates = page.evaluate(r"""async()=>{
              testReplies=[{text:'교정일자 2027.03.17\n차기교정예정일자 2028.03.17',confidence:96},{text:'8765432109',confidence:96}];
              const result=await ToolOCR.recognize(await testPhoto(false),testTools);return testSummary(result);
            }""")
            assert dates['eligible'] and not dates['failed'], dates
            assert not any(source.startswith('date-') for source in dates['sources']), dates
            assert dates['reads'] == 2 and dates['sources'] == ['original', 'identity-crop'], dates
            cases['clear_existing_dates_are_not_repeated_to_recover_identity'] = dates
            page.close()

            page = new_page()
            conflict = page.evaluate(r"""async()=>{
              testQuickProposals();let tags=0;testReplyForInput=input=>input.height===200?{text:tags++%2?'4198765023':'8765432109'}:input.height===300?testDateReply:{text:''};
              const result=await ToolOCR.recognize(await testPhoto(true),testTools);return testSummary(result);
            }""")
            assert not conflict['eligible'] and not conflict['fastPath'], conflict
            assert conflict['sources'][:2] == ['identity-probe-fast-binary-0', 'identity-probe-fast-binary-1'], conflict
            cases['second_initial_tag_can_block_a_conflicting_serial'] = conflict
            page.close()

            page = new_page()
            budget = page.evaluate(r"""async()=>{
              // Advance the wall clock only inside actual recognition calls;
              // image loading and engine initialization retain their real clock.
              testClockStep=12000;testClockReject=true;const began=performance.now();let result;
              try{result=await ToolOCR.recognize(await testPhoto(false),testTools);}
              finally{if(testClockOriginalNow)Date.now=testClockOriginalNow;testClockOriginalNow=null;testClockStep=0;testClockReject=false;}
              return {...testSummary(result),realMilliseconds:performance.now()-began,
                simulatedRecognitionMilliseconds:testClockElapsed,terminated:testTerminations};
            }""")
            assert not budget['eligible'] and budget['failed'], budget
            assert budget['simulatedRecognitionMilliseconds'] >= 35000, budget
            assert budget['reads'] <= 4 and budget['terminated'] == 1, budget
            assert budget['realMilliseconds'] < 3000, budget
            # A started operation may reject when the shared deadline cancels
            # the worker. It must be observed, rather than an unhandled error.
            page.wait_for_timeout(50)
            assert not errors, errors
            cases['unreadable_photo_has_one_shared_35_second_recognition_budget'] = budget
            recovered = page.evaluate(r"""async()=>{
              testQuickProposals();testReplies=[{text:'8765432109'},{text:''},testDateReply,testDateReply];
              const result=await ToolOCR.recognize(await testPhoto(true),testTools);return testSummary(result);
            }""")
            assert recovered['eligible'] and recovered['workers'] == 2, recovered
            cases['budget_exhaustion_discards_worker_before_next_photo'] = recovered
            page.close()
            browser.close()
    assert not errors and not external, (errors, external)
    print(json.dumps({'cases':cases,'case_count':len(cases),'page_errors':errors,
                      'unexpected_requests':external,
                      'scope':'Controlled OCR/proposal orchestration; no accuracy or physical device claim'}, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
