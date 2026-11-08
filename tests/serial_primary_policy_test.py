"""Serial lookup chooses the registered engineer, without requiring a photo name.

Production parsing and policy run in Node with synthetic OCR observations.
The checks cover all rows in a mixed-owner/interval fixture and retain real
contradiction, ambiguity and date guards. They do not measure OCR accuracy.
"""
from pathlib import Path
import subprocess


ROOT = Path(__file__).resolve().parents[1]


CASES = r"""
const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
global.window=global;
for(const source of ['web/ocr.js','web/automation.js'])vm.runInThisContext(fs.readFileSync(source,'utf8'),{filename:source});
const owners=['홍길동','김하늘','박지수','최지훈','이해솔'];
const tools=Array.from({length:30},(_,index)=>({id:'tool-'+index,name:owners[index%owners.length],
 serial:'74826'+String(12345+index*37),months:[6,12,24][index%3],
 last:'2026-02-14',due:'2027-02-14',photoCount:1}));
const cases={};
const dateFor=tool=>({last:'2027-02-14',due:new Date(Date.UTC(2027,1+tool.months,14)).toISOString().slice(0,10)});
function parse(tool,readings,records=tools){
 const dates=dateFor(tool),text='교정일자 '+dates.last+'\n차기교정예정일자 '+dates.due;
 const passes=[{text,source:'original',confidence:96},...readings];
 return ToolOCR.parse(passes.map(pass=>pass.text).join('\n'),records,{passes});
}
const serialRead=(value,source='identity-crop-serial-7',confidence=95)=>({text:'Serial: '+value,source,confidence});
function check(name,tool,readings,{records=tools,eligible=true,method='serial',mutate=null}={}){
 const parsed=parse(tool,readings,records);if(mutate)mutate(parsed);
 const evidenceBefore=JSON.stringify(parsed),recordsBefore=JSON.stringify(records);
 const assessment=ToolAutomation.assess({tools:records},parsed);
 assert.equal(assessment.eligible,eligible,name+': '+JSON.stringify(assessment));
 if(eligible){
  assert.equal(assessment.toolId,tool.id,name);
  assert.equal(assessment.name,tool.name,name+': registered engineer must be authoritative');
  assert.equal(assessment.serial,tool.serial,name+': registered serial must be authoritative');
  assert.equal(assessment.matchMethod,method,name);
 }else assert.ok(assessment.reasons.length,name+': held result needs an explanation');
 assert.equal(JSON.stringify(parsed),evidenceBefore,name+': OCR evidence must not be rewritten');
 assert.equal(JSON.stringify(records),recordsBefore,name+': reference records must not be rewritten');
 cases[name]={eligible:assessment.eligible,method:assessment.matchMethod,canonicalName:assessment.name,reasons:assessment.reasons};
}
for(const tool of tools){
 check('complete_serial_without_owner_'+tool.id,tool,[serialRead(tool.serial)]);
 const fragment=tool.serial.slice(-5);
 check('unique_partial_without_owner_'+tool.id,tool,[serialRead(fragment,'identity-crop-serial-7'),serialRead(fragment,'identity-crop-serial-13')],{method:'partial-serial'});
}
const target=tools[0],other=tools[1];
check('garbled_unlabelled_nonregistered_name_does_not_become_engineer',target,[serialRead(target.serial),{text:'홍기동',source:'identity-crop-owner',confidence:95}]);
check('low_confidence_registered_other_name_does_not_override_serial',target,[serialRead(target.serial),{text:other.name,source:'identity-crop-owner',confidence:24}]);
check('low_confidence_labelled_other_name_does_not_override_serial',target,[serialRead(target.serial),{text:'담당자: '+other.name,source:'identity-crop-owner',confidence:24}]);
check('registered_owner_with_parenthetical_annotation_comes_from_excel', {...target,name:target.name+' (공용)',ownerBaseName:target.name},
 [serialRead(target.serial)],{records:tools.map(tool=>tool.id===target.id?{...tool,name:tool.name+' (공용)',ownerBaseName:tool.name}:tool)});
check('clear_registered_other_owner_remains_a_real_conflict',target,[serialRead(target.serial),serialRead(target.serial,'identity-crop-serial-13'),{text:other.name,source:'identity-crop-owner',confidence:95}],{eligible:false});
check('clear_labelled_other_owner_is_not_silently_discarded',target,[serialRead(target.serial),{text:'담당자: '+other.name,source:'identity-crop-owner',confidence:95}],{eligible:false});
check('another_complete_registered_serial_always_stops_approval',target,[serialRead(target.serial),serialRead(other.serial,'identity-crop-second-serial')],{eligible:false});
check('duplicate_registered_full_serial_cannot_be_fixed_by_owner_name',target,[serialRead(target.serial),{text:target.name,source:'identity-crop-owner',confidence:95}],
 {records:[...tools,{...other,id:'duplicate',serial:target.serial}],eligible:false});
check('nearby_unknown_full_number_is_not_replaced_with_a_registered_serial',target,[serialRead(target.serial.slice(0,-1)+'6')],{eligible:false});
const fragment=target.serial.slice(-5);
check('another_owner_and_cycle_cannot_break_a_partial_serial_tie',target,[serialRead(fragment),serialRead(fragment,'identity-crop-serial-13')],
 {records:[...tools,{...other,id:'partial-collision',serial:'93028'+fragment}],eligible:false});
check('missing_dates_are_not_inferred_from_registered_history',target,[serialRead(target.serial)],{eligible:false,mutate:result=>{result.dateEvidence={dates:[],ambiguous:false,method:'',confidence:null};}});
check('retired_tool_is_not_updated_even_with_an_exact_serial', {...target,due:'Disposed'},[serialRead(target.serial)],
 {records:tools.map(tool=>tool.id===target.id?{...tool,due:'Disposed'}:tool),eligible:false});
console.log(JSON.stringify({case_count:Object.keys(cases).length,cases,
 scope:'Controlled OCR policy: canonical Excel owner/serial, no camera accuracy or cloud-write claim'},null,2));
"""


if __name__ == '__main__':
    subprocess.run(['node', '-e', CASES], cwd=ROOT, check=True)
