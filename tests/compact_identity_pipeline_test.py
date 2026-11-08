"""Compact numeric crops use original gray pixels and preserve identity conflicts.

Only OCR replies and paper proposals are controlled. Production image loading,
crop geometry, parser, approval policy and recognition orchestration run in a
real browser. These checks do not measure camera-photo OCR accuracy.
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
  window.testReads=[];window.testPreparations=[];window.testMode='correct';window.testGlyphHeight=125;
  window.testBinaryTags=0;window.testDateReads=0;window.testWorkers=0;
  window.testTools=[
    {id:'a',name:'홍길동',serial:'8765432109',months:12,interval:12,cycle:'MONTHS',last:'2026-03-17',due:'2027-03-17',photoCount:1},
    {id:'b',name:'김하늘',serial:'4198765023',months:12,interval:12,cycle:'MONTHS',last:'2026-03-17',due:'2027-03-17',photoCount:1}
  ];
  window.testPhoto=async()=>{
    const canvas=document.createElement('canvas');canvas.width=1800;canvas.height=2400;
    const context=canvas.getContext('2d');context.fillStyle='#bbb';context.fillRect(0,0,canvas.width,canvas.height);
    if(testMode==='edge'){
      // Original pixels outside the proposed tag have a distinct shade.
      // White padding or reusing the clipped prepared tag cannot recreate it.
      context.fillStyle='#737373';context.fillRect(0,200,100,300);
      context.fillStyle='#494949';context.fillRect(0,200,80,300);
    }
    return await new Promise(resolve=>canvas.toBlob(resolve,'image/png'));
  };
  window.ToolRuntime={ready:async()=>({workerSource:'/* controlled OCR fixture */',coreSource:'function TesseractCore(){}',languages:{kor:'fixture',eng:'fixture'}})};
  window.ToolFastOCR={
    regions:()=>[
      {kind:'tag',tagIndex:0,bbox:{x0:100,y0:100,x1:860,y1:560}},
      {kind:'tag',tagIndex:1,bbox:{x0:100,y0:650,x1:860,y1:1110}},
      {kind:'dates',bbox:{x0:100,y0:1300,x1:700,y1:1600},cells:[{x0:100,y0:1300,x1:700,y1:1380},{x0:100,y0:1400,x1:700,y1:1480},{x0:100,y0:1500,x1:700,y1:1580}]}
    ],
    prepare:(_canvas,region,options)=>{
      const canvas=document.createElement('canvas');canvas.width=region.kind==='tag'?800:600;canvas.height=region.kind==='tag'?500:300;
      const shade=region.kind==='dates'?220:options?.binary?40:region.tagIndex===1?191:173;
      const context=canvas.getContext('2d');context.fillStyle=`rgb(${shade},${shade},${shade})`;context.fillRect(0,0,canvas.width,canvas.height);
      testPreparations.push({kind:region.kind,tagIndex:region.tagIndex,binary:!!options?.binary,shade});return canvas;
    }
  };
  window.Tesseract={createWorker:async()=>{
    testWorkers++;let parameters={};
    return {setParameters:async value=>{parameters={...parameters,...value};},terminate:()=>{},recognize:async input=>{
      const width=input.width,height=input.height;
      const context=input.getContext('2d',{willReadFrequently:true});
      const shade=context.getImageData(Math.floor(width/2),Math.floor(height/2),1,1).data[0];
      const leftShade=context.getImageData(Math.min(30,width-1),Math.floor(height/2),1,1).data[0];
      let text='',confidence=96,words=[];
      const word=(value,level)=>({text:value,confidence:level,bbox:{x0:testMode==='edge'?20:100,y0:180,x1:testMode==='edge'?620:700,y1:180+testGlyphHeight}});
      const tag=width===800&&height===500,dates=width===600&&height===300;
      const compact=width<500&&height<160&&([173,191].includes(shade)||(testMode==='edge'&&leftShade<110&&shade>140));
      if(tag&&shade===40){
        const index=testBinaryTags++;
        if(testMode==='conflict') {text=index===0?'8765432109':'4198765023';confidence=94;words=[word(text,94)];}
        else if(index===(testMode==='second-tag'?1:0)){
          text=testMode==='punctuation'?'87654)3210q':['alphanumeric','metadata'].includes(testMode)?'728416355V':testMode.startsWith('date-')?'8765432109':'8765432106';
          confidence=['edge','punctuation'].includes(testMode)?52:84;words=[word(text,confidence)];
          if(testMode==='metadata'){
            words.unshift({text:'Certificate',confidence:95,bbox:{x0:20,y0:180,x1:90,y1:180+testGlyphHeight}});
            text='Certificate '+text;
          }
        }
      }else if(dates){
        const index=testDateReads++;
        if(!testMode.startsWith('date-')||index===1||(index===2&&testMode==='date-confirm'))text='2027.03.17\n2028.03.17';
      }
      else if(compact){text=testMode==='wrong'?'8765432106':testTools[0].serial;confidence=93;words=[{text,confidence,bbox:{x0:20,y0:20,x1:width-20,y1:height-20}}];}
      else if(!tag&&width<800&&height<260){text='8765432100';confidence=35;}
      else if(testMode==='small'&&!tag){text='8765432109';confidence=94;}
      testReads.push({width,height,shade,leftShade,psm:parameters.tessedit_pageseg_mode,text,confidence,compact});
      return {data:{text,confidence,words}};
    }};
  }};
  window.testSummary=result=>({assessment:ToolAutomation.assess({tools:testTools},result),
    sources:result.passes.map(pass=>pass.source),fastPath:!!result.fastPath,failed:result.failed,
    serialEvidence:result.identityEvidence,dateEvidence:result.dateEvidence,reads:testReads,preparations:testPreparations,workers:testWorkers});
}"""


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--chromium', default='/usr/bin/chromium')
    args = parser.parse_args()
    cases, errors, external = {}, [], []
    with tempfile.TemporaryDirectory(prefix='compact-identity-pipeline-') as directory:
        fixture = Path(directory) / 'fixture.html'
        fixture.write_text('<!doctype html><title>Compact identity orchestration</title>')
        with serve_program(fixture) as url, sync_playwright() as playwright:
            browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=['--no-sandbox'])

            def run(mode, glyph_height=125):
                page = browser.new_page()
                page.on('pageerror', lambda error: errors.append(str(error)))
                page.on('request', lambda request: external.append(request.url)
                        if urlsplit(request.url).scheme in {'http', 'https'} and urlsplit(request.url).netloc != urlsplit(url).netloc else None)
                page.goto(url)
                page.evaluate(FIXTURE)
                page.evaluate('value=>{testMode=value.mode;testGlyphHeight=value.height;if(["alphanumeric","metadata"].includes(testMode))testTools[0].serial="72841635SV";}', {'mode': mode, 'height': glyph_height})
                for source in ['web/automation.js', 'web/ocr.js']:
                    page.add_script_tag(path=REPO / source)
                result = page.evaluate('async()=>testSummary(await ToolOCR.recognize(await testPhoto(),testTools))')
                page.close()
                return result

            correct = run('correct')
            assert correct['assessment']['eligible'] and correct['assessment']['toolId'] == 'a', correct
            assert correct['assessment']['matchMethod'] == 'serial' and correct['fastPath'] and not correct['failed'], correct
            assert 'original' not in correct['sources'], correct
            assert correct['sources'][-1] == 'identity-probe-fast-compact-0-60', correct
            compact = [read for read in correct['reads'] if read['compact']]
            assert len(compact) == 1 and compact[0]['shade'] == 173 and compact[0]['psm'] == '7', correct
            native = correct['reads'][-2]
            assert native['shade'] == 40 and compact[0]['width'] < native['width'] * .65 and compact[0]['height'] < native['height'] * .7, correct
            assert all(proof['serial'] != '8765432106' for proof in correct['serialEvidence']['observedSerialEvidence']), correct
            cases['compact_gray_pixels_resolve_lone_wrong_locator_without_fuzzy_matching'] = correct

            second = run('second-tag')
            assert second['assessment']['eligible'] and second['fastPath'], second
            compact = [read for read in second['reads'] if read['compact']]
            assert len(compact) == 1 and compact[0]['shade'] == 191, second
            cases['locator_in_second_tag_uses_that_tags_original_gray_pixels'] = second

            conflict = run('conflict')
            assert not conflict['assessment']['eligible'], conflict
            assert {proof['toolId'] for proof in conflict['serialEvidence']['serialMatches']} == {'a', 'b'}, conflict
            assert conflict['sources'][:2] == ['identity-probe-fast-binary-0', 'identity-probe-fast-binary-1'], conflict
            cases['two_registered_serials_remain_held_even_after_correct_compact_read'] = conflict

            wrong = run('wrong')
            assert any(read['compact'] for read in wrong['reads']), wrong
            assert not wrong['assessment']['eligible'] and wrong['assessment']['toolId'] == '', wrong
            assert not wrong['serialEvidence']['serialMatches'], wrong
            cases['compact_unregistered_near_number_does_not_become_a_workbook_guess'] = wrong

            small = run('small', glyph_height=50)
            assert not any(source.startswith('identity-probe-fast-compact-') for source in small['sources']), small
            assert not any(read['compact'] for read in small['reads']), small
            assert small['assessment']['eligible'] and small['assessment']['toolId'] == 'a', small
            cases['small_glyphs_skip_compaction_and_keep_normal_recovery'] = small

            punctuation = run('punctuation')
            assert punctuation['assessment']['eligible'] and punctuation['fastPath'], punctuation
            assert any(source.startswith('identity-probe-fast-compact-') for source in punctuation['sources']), punctuation
            compact = [read for read in punctuation['reads'] if read['compact']]
            assert len(compact) == 1 and compact[0]['text'] == '8765432109', punctuation
            assert all(read['text'] != '8765432109' for read in punctuation['reads'] if not read['compact']), punctuation
            assert punctuation['serialEvidence']['serialMatches'][0]['serial'] == '8765432109', punctuation
            cases['digit_dominant_punctuation_is_only_a_locator_until_real_reread'] = punctuation

            alphanumeric = run('alphanumeric')
            assert alphanumeric['assessment']['eligible'] and alphanumeric['fastPath'], alphanumeric
            assert alphanumeric['assessment']['serial'] == '72841635SV' and alphanumeric['assessment']['name'] == '홍길동', alphanumeric
            assert any(source.startswith('identity-probe-fast-compact-') for source in alphanumeric['sources']), alphanumeric
            compact = [read for read in alphanumeric['reads'] if read['compact']]
            assert len(compact) == 1 and compact[0]['text'] == '72841635SV', alphanumeric
            assert all(read['text'] != '72841635SV' for read in alphanumeric['reads'] if not read['compact']), alphanumeric
            cases['numeric_dominant_alphanumeric_serial_is_a_locator_without_letter_substitution'] = alphanumeric

            metadata = run('metadata')
            assert not metadata['assessment']['eligible'], metadata
            assert not any(source.startswith('identity-probe-fast-compact-') for source in metadata['sources']), metadata
            assert not any(read['compact'] for read in metadata['reads']), metadata
            cases['certificate_heading_excludes_its_numeric_dominant_value_from_identity_retry'] = metadata

            edge = run('edge')
            assert edge['assessment']['eligible'] and edge['fastPath'], edge
            compact = [read for read in edge['reads'] if read['compact']]
            assert len(compact) == 1 and compact[0]['leftShade'] < 90 and compact[0]['shade'] > 140, edge
            assert compact[0]['shade'] not in [40,173,191,255], edge
            assert edge['sources'][-1].startswith('identity-probe-fast-compact-'), edge
            cases['clipped_locator_reads_real_pixels_outside_the_prepared_tag'] = edge

            confirmed_dates = run('date-confirm')
            assert confirmed_dates['assessment']['eligible'] and confirmed_dates['fastPath'], confirmed_dates
            assert 'date-crop-fast-binary-lines' in confirmed_dates['sources'], confirmed_dates
            assert confirmed_dates['dateEvidence']['method'] == 'repeated-dates' and confirmed_dates['dateEvidence']['repeatCount'] == 2, confirmed_dates
            date_reads = [read for read in confirmed_dates['reads'] if read['width'] == 600 and read['height'] == 300]
            assert len(date_reads) == 3 and [read['psm'] for read in date_reads] == ['11','11','6'], confirmed_dates
            assert not date_reads[0]['text'] and date_reads[1]['text'] == date_reads[2]['text'], confirmed_dates
            cases['single_binary_date_pair_is_confirmed_by_independent_psm6'] = confirmed_dates

            unconfirmed_dates = run('date-missing')
            assert not unconfirmed_dates['assessment']['eligible'], unconfirmed_dates
            assert not unconfirmed_dates['dateEvidence']['method'], unconfirmed_dates
            assert 'date-probe-fast-binary-lines' in unconfirmed_dates['sources'] and 'date-crop-fast-binary-lines' not in unconfirmed_dates['sources'], unconfirmed_dates
            cases['one_unconfirmed_date_pair_does_not_gain_automatic_approval'] = unconfirmed_dates
            browser.close()
    assert not errors and not external, (errors, external)
    print(json.dumps({'cases': cases, 'case_count': len(cases), 'page_errors': errors,
                      'unexpected_requests': external,
                      'scope': 'Controlled OCR/proposals; real crop pixels and matching policy; no photo accuracy claim'}, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
