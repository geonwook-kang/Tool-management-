"""Real multi-owner photograph OCR, timing and connected Excel autosave.

Inputs are original camera JPEGs supplied separately from Git. Expected names,
serials and dates are Python assertions only, never recognizer inputs. One real
browser context reads a cold first image and warm subsequent images with the
same worker. The picker uses real OPFS handles; actual OneDrive sync is outside
this test. All worker responses are observed without replacing their content.
"""
from __future__ import annotations

import argparse
from datetime import datetime
from hashlib import sha256
import json
from pathlib import Path
import tempfile
import time
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

from browser_test import REPO, download_expression, serve_program
from connected_file_test import PICKER_SCRIPT, export_shared, ready, seed_file
from smax_workbook_test import drawing_info, rel_path


OBSERVE = r"""() => {
  window.ocrTimings={initializations:[],reinitializations:[],reads:[],assessments:[]};
  const create=Tesseract.createWorker.bind(Tesseract);
  Tesseract.createWorker=async(...args)=>{
    const began=performance.now(),id=ocrTimings.initializations.length;
    const initialization={id,began,milliseconds:null};ocrTimings.initializations.push(initialization);
    let worker;
    try {worker=await create(...args);initialization.milliseconds=performance.now()-began;}
    catch(error){initialization.error=String(error);throw error;}
    const codes=value=>typeof value==='string'?value:Array.isArray(value)?value.map(v=>typeof v==='string'?v:v.code):null;
    let languages=codes(args[0]);
    if(typeof worker.reinitialize==='function'){
      const reinitialize=worker.reinitialize.bind(worker);
      worker.reinitialize=async(...values)=>{const at=performance.now(),result=await reinitialize(...values);languages=codes(values[0]);ocrTimings.reinitializations.push({worker:id,languages,milliseconds:performance.now()-at});return result;};
    }
    let parameters={};const set=worker.setParameters.bind(worker),recognize=worker.recognize.bind(worker);
    worker.setParameters=async values=>{const response=await set(values);parameters={...parameters,...values};return response;};
    worker.recognize=async(...input)=>{
      const at=performance.now(),entry={worker:id,began:at,milliseconds:null,
        input:{width:input[0].width,height:input[0].height},languages,parameters:{...parameters}};
      ocrTimings.reads.push(entry);
      try {
        const response=await recognize(...input),data=response.data||{};
        entry.milliseconds=performance.now()-at;entry.text=data.text;entry.confidence=data.confidence;
        const words=Array.isArray(data.words)?data.words:(data.blocks||[]).flatMap(b=>(b.paragraphs||[]).flatMap(p=>(p.lines||[]).flatMap(l=>l.words||[])));
        entry.words=words.map(w=>({text:w.text,confidence:w.confidence,bbox:w.bbox}));return response;
      }catch(error){entry.milliseconds=performance.now()-at;entry.error=String(error);throw error;}
    };return worker;
  };
  const assess=ToolAutomation.assess.bind(ToolAutomation);
  window.ToolAutomation={...ToolAutomation,assess:(...args)=>{const at=performance.now(),result=assess(...args);ocrTimings.assessments.push({milliseconds:performance.now()-at,eligible:result.eligible});return result;}};
}"""


def arguments():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--program',type=Path,default=REPO/'web/dist/Tool-Manager.html')
    parser.add_argument('--workbook',type=Path)
    parser.add_argument('--manifest',type=Path,required=True,
                        help='JSON array of {photo,name,serial,observed_serial,last,due}; expectations never passed to JS')
    parser.add_argument('--diagnostics',type=Path,default=Path('/tmp/tool-multi-owner-ocr.json'))
    parser.add_argument('--chromium',default='/usr/bin/chromium')
    parser.add_argument('--diagnose-only',action='store_true')
    parser.add_argument('--maximum-seconds',type=float,help='Optional machine-specific per-photo total bound')
    parser.add_argument('--baseline',type=Path,help='Prior report of the same original files/order for timing comparison')
    parser.add_argument('--minimum-speedup',type=float,help='Optional machine-specific total time speedup bound; requires --baseline')
    args=parser.parse_args()
    if args.workbook is None:args.workbook=next(Path('/workspace/attachments').rglob('*.xlsx'),None)
    for label,path in [('program',args.program),('workbook',args.workbook),('manifest',args.manifest)]:
        if path is None or not path.is_file():parser.error('Supply --'+label+' with an existing file')
    if args.minimum_speedup is not None and args.baseline is None:parser.error('--minimum-speedup requires --baseline')
    if args.baseline is not None and not args.baseline.is_file():parser.error('Baseline report missing')
    args.cases=json.loads(args.manifest.read_text())
    for case in args.cases:
        for key in ['photo','name','serial','last','due']:
            if not case.get(key):parser.error('Every manifest case needs '+key)
        if not Path(case['photo']).is_file():parser.error('Photo missing: '+case['photo'])
        datetime.fromisoformat(case['last']);datetime.fromisoformat(case['due'])
    return args


def target_reference(workbook, serial):
    book=load_workbook(workbook);targets=[]
    for sheet in book:
        headings={cell.value:cell.column for cell in sheet[1]}
        if 'Serial' not in headings:continue
        for row in range(2,sheet.max_row+1):
            if str(sheet.cell(row,headings['Serial']).value)==serial:
                targets.append({'sheet':sheet.title,'row':row,'headers':headings,
                    'last_cell':sheet.cell(row,headings['Cal Last']).coordinate,
                    'due_cell':sheet.cell(row,headings['Cal Next']).coordinate})
    book.close();assert len(targets)==1,targets
    return targets[0]


def verify_increment(before_file, after_file, reference, case, photo_hash):
    before,after=load_workbook(before_file),load_workbook(after_file)
    assert before.sheetnames==after.sheetnames
    permitted={(reference['sheet'],reference['last_cell']):datetime.fromisoformat(case['last']),
               (reference['sheet'],reference['due_cell']):datetime.fromisoformat(case['due'])}
    checked=0
    for old_sheet in before:
        new_sheet=after[old_sheet.title]
        assert (old_sheet.max_row,old_sheet.max_column)==(new_sheet.max_row,new_sheet.max_column)
        assert list(old_sheet.merged_cells.ranges)==list(new_sheet.merged_cells.ranges)
        for row in old_sheet:
            for old in row:
                new=new_sheet[old.coordinate]
                assert new.value==permitted.get((old_sheet.title,old.coordinate),old.value),(old_sheet.title,old.coordinate,old.value,new.value)
                assert new._style==old._style,(old_sheet.title,old.coordinate)
                if old.coordinate not in [reference['last_cell'],reference['due_cell']] or old_sheet.title!=reference['sheet']:
                    assert new.data_type==old.data_type,(old_sheet.title,old.coordinate,'storage type changed')
                checked+=1
    old_serial=before[reference['sheet']].cell(reference['row'],reference['headers']['Serial'])
    new_serial=after[reference['sheet']].cell(reference['row'],reference['headers']['Serial'])
    assert str(new_serial.value)==case['serial'] and (new_serial.value,new_serial.data_type)==(old_serial.value,old_serial.data_type)
    count=sum(len(sheet._images) for sheet in after)
    assert count==sum(len(sheet._images) for sheet in before)
    with ZipFile(before_file) as old_zip,ZipFile(after_file) as new_zip:
        changed={name for name in old_zip.namelist() if old_zip.read(name)!=new_zip.read(name)}
        sheet,drawing,old_pictures=drawing_info(old_zip,reference['sheet'])
        new_sheet,new_drawing,new_pictures=drawing_info(new_zip,reference['sheet'])
        assert (sheet,drawing)==(new_sheet,new_drawing)
        assert set(old_pictures)==set(new_pictures)
        target=(reference['row'],reference['headers']['Label Photo'])
        assert target in old_pictures
        for location,old in old_pictures.items():
            new=new_pictures[location]
            if location==target:
                assert len(new)==1 and sha256(new[0]['bytes']).hexdigest()==photo_hash
            else:assert old==new,('Unselected picture/anchor changed',location)
        allowed={sheet,drawing,rel_path(drawing)}
        if '[Content_Types].xml' in changed:
            signature=lambda node:(node.tag,tuple(sorted(node.attrib.items())),(node.text or '').strip())
            old={signature(node) for node in ET.fromstring(old_zip.read('[Content_Types].xml'))}
            new={signature(node) for node in ET.fromstring(new_zip.read('[Content_Types].xml'))}
            assert old<=new
            assert all(dict(a)=={'Extension':'jpg','ContentType':'image/jpeg'} and tag.endswith('}Default') for tag,a,_ in new-old)
            allowed.add('[Content_Types].xml')
        assert changed<=allowed,changed-allowed
        added={name for name in new_zip.namelist() if name not in old_zip.namelist() and not name.endswith('/')}
        assert len(added)==1 and all(name.startswith('xl/media/') for name in added)
    before.close();after.close()
    return {'cells_and_styles_checked':checked,'pictures':count,'serial_value_and_storage_type_preserved':True,
            'other_cells_styles_smax_photos_preserved':True,'correct_label_photo_anchor_and_bytes':True,'changed_zip_parts':sorted(changed)}


def main():
    args=arguments();original_hash=sha256(args.workbook.read_bytes()).hexdigest()
    photo_hashes={case['photo']:sha256(Path(case['photo']).read_bytes()).hexdigest() for case in args.cases}
    report={'program':str(args.program),'program_sha256':sha256(args.program.read_bytes()).hexdigest(),
            'worker_results_replaced':False,'expected_values_passed_to_recognizer':False,
            'actual_onedrive_sync_tested':False,'diagnose_only':args.diagnose_only,'cases':[],'failures':[]}
    with tempfile.TemporaryDirectory(prefix='multi-owner-actual-') as directory,serve_program(args.program) as url,sync_playwright() as pw:
        temporary=Path(directory);prior_file=temporary/'prior.xlsx';prior_file.write_bytes(args.workbook.read_bytes())
        browser=pw.chromium.launch(executable_path=args.chromium,headless=True,args=['--no-sandbox'])
        context=browser.new_context(accept_downloads=True,locale='ko-KR');context.add_init_script('('+PICKER_SCRIPT+')()')
        page=context.new_page();errors=[];requests=[]
        page.on('pageerror',lambda error:errors.append(str(error)))
        page.goto(url);page.wait_for_function('window.ToolApp&&window.ToolOCR&&window.ToolRuntime')
        page.evaluate(OBSERVE)
        context.route('**/*',lambda route:(requests.append(route.request.url),route.abort()) if route.request.url.startswith(('http:','https:')) else route.continue_())
        context.set_offline(True);seed_file(page,args.workbook)
        page.locator('#connect-onedrive').click();page.wait_for_function('ToolApp.state.source.connected&&ToolApp.state.model');ready(page)
        page.locator('#remember-matches').uncheck()
        for index,case in enumerate(args.cases):
            reference=target_reference(prior_file,case['serial'])
            before=page.evaluate('''()=>({reads:ocrTimings.reads.length,inits:ocrTimings.initializations.length,reinit:ocrTimings.reinitializations.length,assess:ocrTimings.assessments.length,saves:ToolApp.state.source.saveCount})''')
            began=time.monotonic();page.locator('#photos').set_input_files(case['photo'])
            page.wait_for_function('ToolApp.state.photos.length===1&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving',timeout=360000)
            evidence=page.evaluate('''async before=>{
                const p=ToolApp.state.photos[0],handle=await(await navigator.storage.getDirectory()).getFileHandle('company-shared.xlsx'),backup=await ToolConnectedFile.takeBackup();
                const parsedAt=performance.now();ToolOCR.parse(p.result.text,ToolApp.state.model.tools,{passes:p.result.passes});const replayMs=performance.now()-parsedAt;
                return {result:p.result,assessment:p.assessment,saved:!!p.saved,approval:p.approval,
                    sourceError:ToolApp.state.source.error,saveDelta:ToolApp.state.source.saveCount-before.saves,
                    preparedPhotoHash:await testHash(p.image.photoBytes),sourceHash:await testHash(await(await handle.getFile()).arrayBuffer()),
                    backupHash:backup&&await testHash(backup.bytes),nameInput:document.querySelector('#name').value,
                    serialInput:document.querySelector('#serial').value,lastInput:document.querySelector('#last').value,dueInput:document.querySelector('#due').value,
                    reads:ocrTimings.reads.slice(before.reads).map(r=>({...r,elapsedMilliseconds:r.milliseconds==null?performance.now()-r.began:r.milliseconds})),initializations:ocrTimings.initializations.slice(before.inits),reinitializations:ocrTimings.reinitializations.slice(before.reinit),
                    assessmentTimings:ocrTimings.assessments.slice(before.assess),parseReplayMilliseconds:replayMs};
            }''',before)
            entry={'index':index,'original_photo_sha256':photo_hashes[case['photo']],'cold_worker':index==0,
                   'expected':{key:case[key] for key in ['name','serial','last','due']},'seconds':round(time.monotonic()-began,3),**evidence}
            entry['ocr_calls']=len(entry['reads']);entry['ocr_milliseconds']=sum(x.get('milliseconds') or x.get('elapsedMilliseconds') or 0 for x in entry['reads'])
            entry['initialization_milliseconds']=sum(x.get('milliseconds') or 0 for x in entry['initializations'])
            assessment=entry['assessment'] or {};result=entry['result'] or {}
            expected=tuple(case[key] for key in ['name','serial','last','due'])
            failures=[]
            if (result.get('last'),result.get('due'))!=(case['last'],case['due']):failures.append('Both complete observed photo dates were not read')
            observed=case.get('observed_serial',case['serial'])
            identity=result.get('identityEvidence') or {}
            strong=[p for p in identity.get('observedSerialEvidence',[]) if p.get('serial')==observed and (p.get('confidence') or 0)>=70]
            if not strong:failures.append('No strong full observed serial proof')
            for label,actual in [('assessment',tuple(assessment.get(key) for key in ['name','serial','last','due'])),('visible UI',tuple(entry.get(key) for key in ['nameInput','serialInput','lastInput','dueInput']))]:
                if actual!=expected:failures.append(label+' has wrong target/values')
            if result.get('failed') or result.get('identityRecoveryIncomplete'):failures.append('OCR failed or identity recovery incomplete')
            if not assessment.get('eligible') or not entry['saved'] or entry['saveDelta']!=1 or entry['approval']!='automatic':failures.append('Actual photo not automatically saved exactly once')
            if args.maximum_seconds is not None and entry['seconds']>args.maximum_seconds:failures.append('Per-photo time exceeded configured machine-specific bound')
            current_hash=sha256(prior_file.read_bytes()).hexdigest()
            if entry['saved']:
                assert entry['backupHash']==current_hash and entry['sourceHash']!=current_hash
                output=temporary/f'saved-{index}.xlsx';export_shared(page,output)
                entry['workbook_verification']=verify_increment(prior_file,output,reference,case,entry['preparedPhotoHash'])
                backup=temporary/f'backup-{index}.xlsx';download_expression(page,'(await ToolConnectedFile.takeBackup()).bytes',backup)
                assert backup.read_bytes()==prior_file.read_bytes()
                entry['backup_byte_exact']=True;prior_file.write_bytes(output.read_bytes())
            else:assert entry['sourceHash']==current_hash and entry['saveDelta']==0
            entry['failures']=failures;report['cases'].append(entry)
            if failures and not args.diagnose_only:report['failures'].extend(f'case {index}: {message}' for message in failures)
            args.diagnostics.write_text(json.dumps(report,ensure_ascii=False,indent=2))
            print(json.dumps({'case':index,'seconds':entry['seconds'],'ocr_calls':entry['ocr_calls'],'initialization_ms':entry['initialization_milliseconds'],'assessment':assessment,'saved':entry['saved'],'failures':failures},ensure_ascii=False),flush=True)
            page.locator('#remove-photo').click();page.wait_for_function('ToolApp.state.photos.length===0&&!ToolApp.state.preparing&&!ToolApp.state.source.saving')
        assert not errors and not requests,(errors,requests)
        report.update({'page_errors':errors,'http_requests':requests,'final_worker_initializations':page.evaluate('ocrTimings.initializations.length')})
        browser.close()
    assert sha256(args.workbook.read_bytes()).hexdigest()==original_hash
    assert all(sha256(Path(path).read_bytes()).hexdigest()==digest for path,digest in photo_hashes.items())
    report.update({'original_company_workbook_unchanged':True,'original_photos_unchanged':True})
    if args.baseline is not None:
        baseline=json.loads(args.baseline.read_text());old=baseline['cases'];new=report['cases']
        assert [c['original_photo_sha256'] for c in old]==[c['original_photo_sha256'] for c in new], 'Speed comparisons need the same original files and order'
        old_total=sum(c['seconds'] for c in old);new_total=sum(c['seconds'] for c in new)
        ratio=old_total/new_total
        report['speed_comparison']={'baseline_program_sha256':baseline['program_sha256'],
            'baseline_total_seconds':round(old_total,3),'current_total_seconds':round(new_total,3),'total_speedup':round(ratio,3),
            'same_original_bytes_and_order':True,'hardware_scope':'Measured on this test machine; actual PCs may differ',
            'per_photo':[{'original_photo_sha256':a['original_photo_sha256'],'baseline_seconds':a['seconds'],'current_seconds':b['seconds'],
                          'speedup':round(a['seconds']/b['seconds'],3),'baseline_ocr_calls':a['ocr_calls'],'current_ocr_calls':b['ocr_calls']} for a,b in zip(old,new)]}
        if args.minimum_speedup is not None and ratio<args.minimum_speedup:report['failures'].append('Total OCR workflow did not meet configured machine-specific speedup bound')
    args.diagnostics.write_text(json.dumps(report,ensure_ascii=False,indent=2))
    assert not report['failures'],report['failures']


if __name__=='__main__':main()
