"""Numeric Excel serial provenance, zero-only matching, and real OPFS preservation.

OCR strings are controlled in this policy/storage suite; camera accuracy is
checked separately using the supplied original photographs. Inputs stay read
only and no company photographs or workbook are committed.
"""
from __future__ import annotations

from datetime import datetime
from hashlib import sha256
import json
from pathlib import Path
import re
import tempfile
from zipfile import ZipFile

from openpyxl import load_workbook
from PIL import Image
from playwright.sync_api import sync_playwright

from browser_test import arguments, download_expression, serve_program
from connected_file_test import PICKER_SCRIPT, export_shared, ready, seed_file
from smax_workbook_test import drawing_info, rel_path


POLICIES = r"""() => {
 const original=ToolApp.state.model,base=original.tools.find(tool=>tool.sheet==='대전'&&tool.row===10);
 if(!base||!base.numericSerial)throw new Error('The fixture needs the numeric-stored torque serial');
 const observed='0'+base.serial,last='2027-06-25',due='2027-12-25',cases={};
 const copy=value=>JSON.parse(JSON.stringify(value));
 function run(name,{tools=original.tools,value=observed,confidence=95,source='identity-probe-field',mutate=null,withDates=true,owner='',extra='',thirdDate='',eligible=false,method='numeric-cell-serial',expectedTool=base.id,smax=false}={}){
  const text=smax?'SMAX\nTool Detail\nSerial/Asset Number '+value+'\nTool Owner Test Owner\nCalibration Due Date '+due:
   [owner,'Serial: '+value,withDates?'교정일자 '+last+'\n차기교정예정일자 '+due:'',extra].filter(Boolean).join('\n');
  const passes=smax?[{text,confidence,source}]:[
   ...(withDates?[{text:'교정일자 '+last+'\n차기교정예정일자 '+due+(thirdDate?'\n'+thirdDate:''),confidence:95,source:'original'}]:[]),
   {text:[owner,'Serial: '+value,extra].filter(Boolean).join('\n'),confidence,source}];
  const result=ToolOCR.parse(passes.map(pass=>pass.text).join('\n'),tools,{passes});
  if(mutate)mutate(result);
  const before=JSON.stringify(result.identityEvidence),assessment=ToolAutomation.assess({...original,tools},result);
  if(assessment.eligible!==eligible)throw new Error(name+': '+JSON.stringify(assessment));
  if(eligible&&(assessment.toolId!==expectedTool||assessment.matchMethod!==method))throw new Error(name+': wrong target/method '+JSON.stringify(assessment));
  if(before!==JSON.stringify(result.identityEvidence))throw new Error(name+': observed evidence mutated');
  cases[name]={eligible:assessment.eligible,method:assessment.matchMethod,toolId:assessment.toolId,
   observedSerial:assessment.observedSerial||'',reasons:assessment.reasons};
 }
 run('numeric_storage_zero_prefix',{eligible:true});
 run('more_than_one_discarded_zero',{value:'00'+base.serial,eligible:true});
 run('exact_unpadded_stays_exact',{value:base.serial,eligible:true,method:'serial'});
 run('threshold_70',{confidence:70,eligible:true});
 run('confidence_69',{confidence:69});
 run('no_dates_does_not_auto_write',{withDates:false});
 run('missing_numeric_provenance',{tools:original.tools.map(tool=>{const c={...tool};delete c.numericSerial;return c;})});
 run('explicit_text_storage_keeps_exact_serial',{tools:original.tools.map(tool=>tool.id===base.id?{...tool,numericSerial:false}:tool)});
 run('metadata_string_is_not_numeric_provenance',{tools:original.tools.map(tool=>tool.id===base.id?{...tool,numericSerial:'true'}:tool)});
 run('changed_middle_digit_is_not_formatting',{value:observed.slice(0,4)+'9'+observed.slice(5)});
 run('extra_trailing_digit_is_not_formatting',{value:observed+'4'});
 run('letter_o_is_not_zero',{value:'O'+base.serial});
 run('date_crop_cannot_supply_identity',{source:'date-crop-6'});
 run('anonymous_text_is_not_numeric_format_proof',{source:'text'});
 run('low_source_cannot_borrow_aggregate_confidence',{mutate:r=>{for(const p of r.identityEvidence.observedSerialEvidence||[])for(const v of p.sourceReadings||[])v.confidence=20;}});
 run('unlisted_source_cannot_prove_format',{mutate:r=>{for(const p of r.identityEvidence.observedSerialEvidence||[])p.passSources=['another-source'];}});
 run('legacy_evidence_without_source_readings',{mutate:r=>{for(const p of r.identityEvidence.observedSerialEvidence||[])delete p.sourceReadings;}});
 run('learned_alias_is_not_storage_provenance',{mutate:r=>{for(const p of r.identityEvidence.observedSerialEvidence||[])p.learned=true;}});
 run('unfinished_recovery_still_blocks',{mutate:r=>r.identityRecoveryIncomplete=true});
 run('third_date_still_blocks',{thirdDate:'차기교정예정일자 2028.06.25'});
 const other=original.tools.find(tool=>tool.name!==base.name&&!/폐기/.test(tool.due));
 run('another_registered_owner_still_conflicts',{owner:other.name});
 run('another_serial_still_conflicts',{extra:'Serial: '+other.serial});
 run('duplicate_numeric_class_blocks',{tools:[...original.tools,{...base,id:'duplicate-numeric',serial:'00'+base.serial,numericSerial:true}]});
 run('retired_text_numeric_class_blocks',{tools:[...original.tools,{...base,id:'retired-text',serial:observed,numericSerial:false,due:'Disposed'}]});
 const exactText={...base,id:'explicit-text',serial:observed,numericSerial:false};
 run('exact_text_wins_without_stealing_numeric_row',{tools:[...original.tools,exactText],eligible:true,method:'serial',expectedTool:exactText.id});
 run('smax_numeric_storage_zero_prefix',{smax:true,eligible:true});
 run('smax_text_storage_not_zero_normalized',{smax:true,tools:original.tools.map(tool=>tool.id===base.id?{...tool,numericSerial:false}:tool)});
 run('smax_retired_equivalent_collision',{smax:true,tools:[...original.tools,{...base,id:'retired-text',serial:observed,numericSerial:false,due:'Disposed'}]});
 run('smax_other_primary_still_conflicts',{smax:true,mutate:r=>r.smaxEvidence.serialReadings.push({labelVerified:true,value:other.serial,source:'original',confidence:95})});
 // Check every registered identity rather than a selected engineer only.
 const allTools=[];
 for(const target of original.tools){
  const l='2028-01-10',d=ToolWorkbook.addMonths(l,target.months||12),text='Serial: '+target.serial+'\n교정일자 '+l+'\n차기교정예정일자 '+d;
  const result=ToolOCR.parse(text,original.tools,{passes:[{text,source:'original',confidence:95}]}),a=ToolAutomation.assess(original,result);
  const shouldPass=!/폐기|disposed|retired/i.test(target.due||'')&&Number(target.photoCount)<=1;
  if(a.eligible!==shouldPass||shouldPass&&a.toolId!==target.id)throw new Error('registered identity '+target.id+':'+JSON.stringify(a));
  if(target.numericSerial&&/^[1-9][0-9]{2,14}$/.test(target.serial)){
   const ptext=text.replace('Serial: '+target.serial,'Serial: 0'+target.serial),p=ToolOCR.parse(ptext,original.tools,{passes:[{text:ptext,source:'original',confidence:95}]}),b=ToolAutomation.assess(original,p);
   if(b.eligible!==shouldPass||shouldPass&&b.toolId!==target.id)throw new Error('numeric registered identity '+target.id+':'+JSON.stringify(b));
  }
  allTools.push({id:target.id,numeric:target.numericSerial,eligible:a.eligible});
 }
 return {cases,allTools,metadata:original.tools.map(tool=>({sheet:tool.sheet,row:tool.row,serial:tool.serial,numericSerial:tool.numericSerial}))};
}"""


def verify_saved(original, updated, prepared_hash):
    before, after = load_workbook(original), load_workbook(updated)
    checked = 0
    for old_sheet in before:
        new_sheet = after[old_sheet.title]
        assert list(old_sheet.merged_cells.ranges) == list(new_sheet.merged_cells.ranges)
        for row in old_sheet:
            for old in row:
                new = new_sheet[old.coordinate]
                expected = old.value
                if old_sheet.title == '대전' and old.coordinate == 'A10': expected = datetime(2027,12,25)
                if old_sheet.title == '대전' and old.coordinate == 'B10': expected = datetime(2027,6,25)
                assert new.value == expected and new._style == old._style, (old_sheet.title,old.coordinate)
                checked += 1
    assert after['대전']['F10'].value == before['대전']['F10'].value
    assert after['대전']['F10'].data_type == before['대전']['F10'].data_type == 'n'
    pictures = sum(len(s._images) for s in after)
    assert pictures == sum(len(s._images) for s in before)
    with ZipFile(original) as old_zip, ZipFile(updated) as new_zip:
        sheet,drawing,old_pictures=drawing_info(old_zip,'대전')
        new_sheet,new_drawing,new_pictures=drawing_info(new_zip,'대전')
        assert (sheet,drawing)==(new_sheet,new_drawing) and set(old_pictures)==set(new_pictures)
        for location,old in old_pictures.items():
            if location==(10,10):assert len(new_pictures[location])==1 and sha256(new_pictures[location][0]['bytes']).hexdigest()==prepared_hash
            else:assert old==new_pictures[location],location
        allowed={sheet,drawing,rel_path(drawing),'[Content_Types].xml'}
        assert {name for name in old_zip.namelist() if old_zip.read(name)!=new_zip.read(name)}<=allowed
    before.close();after.close()
    return {'cells_styles_checked':checked,'pictures':pictures,'numeric_serial_cell_preserved':True,'other_photos_smax_preserved':True}


def main():
    args=arguments(__doc__)
    original_hash=sha256(args.workbook.read_bytes()).hexdigest()
    with tempfile.TemporaryDirectory(prefix='numeric-serial-') as directory,serve_program(args.program) as url,sync_playwright() as pw:
        folder=Path(directory);photo=folder/'controlled-label.png';Image.new('RGB',(320,200),'white').save(photo)
        browser=pw.chromium.launch(executable_path=args.chromium,headless=True,args=['--no-sandbox'])
        context=browser.new_context(accept_downloads=True,locale='ko-KR');context.add_init_script('('+PICKER_SCRIPT+')()');page=context.new_page();errors=[];page.on('pageerror',lambda e:errors.append(str(e)));page.goto(url);page.wait_for_function('window.ToolApp&&window.ToolAutomation')
        page.locator('#workbook').set_input_files(args.workbook);page.wait_for_function('ToolApp.state.model&&!ToolApp.state.busy')
        result=page.evaluate(POLICIES)
        source=load_workbook(args.workbook)
        for item in result['metadata']:
            cell=source[item['sheet']].cell(item['row'],6)
            expected=cell.data_type=='n' and re.fullmatch('[0-9]{3,15}',str(cell.value)) is not None
            assert item['numericSerial']==expected,(item,cell.data_type,cell.value)
        source.close()
        # Raw XML provenance must not follow displayed numeric-looking text,
        # scientific notation or a cached formula result.
        metadata_variants=page.evaluate(r'''async()=>{
          const file=document.querySelector('#workbook').files[0],bytes=await file.arrayBuffer(),out={};
          for(const [kind,body,type] of [['numeric','<v>218800584</v>','n'],['inline','<is><t>218800584</t></is>','inlineStr'],['formula','<f>218800000+584</f><v>218800584</v>','n'],['exponent','<v>2.18800584E8</v>','n']]){
            const z=await JSZip.loadAsync(bytes),name='xl/worksheets/sheet1.xml',doc=new DOMParser().parseFromString(await z.file(name).async('string'),'application/xml'),cell=Array.from(doc.getElementsByTagNameNS('*','c')).find(c=>c.getAttribute('r')==='F10');
            while(cell.firstChild)cell.removeChild(cell.firstChild);cell.setAttribute('t',type);
            const fragment=new DOMParser().parseFromString('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'+body+'</c>','application/xml');
            for(const node of Array.from(fragment.documentElement.childNodes))cell.appendChild(doc.importNode(node,true));z.file(name,new XMLSerializer().serializeToString(doc));
            const m=await ToolWorkbook.load(await z.generateAsync({type:'uint8array'}));out[kind]=m.tools.find(t=>t.sheet==='대전'&&t.row===10).numericSerial;
          }return out;
        }''')
        assert metadata_variants=={'numeric':True,'inline':False,'formula':False,'exponent':False},metadata_variants
        initial=seed_file(page,args.workbook)
        entry='#quick-workbook' if page.evaluate('ToolSources.inspect().pcEntry') else '#connect-onedrive'
        page.locator(entry).click();page.wait_for_function('ToolApp.state.source.connected&&ToolApp.state.model');ready(page)
        page.evaluate("""()=>{ToolOCR.recognize=async(file,tools)=>{const text='Serial: 0218800584\\n교정일자 2027.06.25\\n차기교정예정일자 2027.12.25';return ToolOCR.parse(text,tools,{passes:[{text,source:'original',confidence:95}]});};}""")
        page.locator('#photos').set_input_files(photo);page.wait_for_function('ToolApp.state.photos.length===1&&!ToolApp.state.busy&&!ToolApp.state.preparing&&!ToolApp.state.source.saving')
        saved=page.evaluate("""async()=>{const p=ToolApp.state.photos[0],b=await ToolConnectedFile.takeBackup();return {assessment:p.assessment,saved:!!p.saved,saveCount:ToolApp.state.source.saveCount,backupHash:await testHash(b.bytes),evidence:p.result.identityEvidence,preparedPhotoHash:await testHash(p.image.photoBytes)};}""")
        assert saved['saved'] and saved['saveCount']==1 and saved['backupHash']==initial
        assert saved['assessment']['serial']=='218800584' and saved['assessment']['observedSerial']=='0218800584'
        assert '0218800584' in saved['evidence']['observedSerials']
        output=folder/'updated.xlsx';export_shared(page,output);saved['preservation']=verify_saved(args.workbook,output,saved['preparedPhotoHash'])
        browser.close();assert not errors,errors
    assert sha256(args.workbook.read_bytes()).hexdigest()==original_hash
    print(json.dumps({'policies':result,'provenance_variants':metadata_variants,'connected_save':saved,'company_workbook_unchanged':True,'page_errors':errors,'ocr_accuracy':'Not measured; controlled OCR strings'},ensure_ascii=False,indent=2))


if __name__=='__main__':main()
