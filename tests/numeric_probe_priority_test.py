"""Numeric-cell serial provenance gets the same probe handling as exact IDs.

The production parser and automatic matcher run in Node with controlled OCR
observations. This is a policy/provenance test, not a photo accuracy claim.
Fixture records and characters are synthetic; no company inputs are loaded.
"""
from pathlib import Path
import subprocess


ROOT = Path(__file__).resolve().parents[1]


CASES = r"""
const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
global.window=global;
for(const source of ['web/ocr.js','web/automation.js'])vm.runInThisContext(fs.readFileSync(source,'utf8'),{filename:source});
const tools=[
 {id:'numeric-a',name:'홍길동',serial:'583902147',numericSerial:true,months:12,last:'2026-01-14',due:'2027-01-14',photoCount:1},
 {id:'numeric-b',name:'김하늘',serial:'847316259',numericSerial:true,months:12,last:'2026-01-14',due:'2027-01-14',photoCount:1},
 {id:'alpha',name:'박지수',serial:'TK923578',numericSerial:false,months:6,last:'2026-01-14',due:'2026-07-14',photoCount:1}
];
const observed='0'+tools[0].serial, otherObserved='0'+tools[1].serial, noise='8361952407';
const read=(text,source,confidence=95,extra={})=>({text,source,confidence,...extra});
const original=value=>read(value,'original');
const probe=(value,source='identity-probe-fast-token-noise',confidence=84)=>read(value,source,confidence);
const numericProbe=()=>probe(observed,'identity-probe-fast-token-serial',93);
const cases={},failures=[];
function run(name,{reads=[original(observed)],records=tools,eligible=true,method='numeric-cell-serial',mutate=null,
 mustKeep=[],mustDrop=[],expectedTool='numeric-a'}={}){
 const dates=read('교정일자 2027.01.14\n차기교정예정일자 2028.01.14','date-crop-confirmed',96);
 const passes=[dates,...reads],parsed=ToolOCR.parse(passes.map(pass=>pass.text).join('\n'),records,{passes});
 if(mutate)mutate(parsed);
 const before=JSON.stringify(parsed),referenceBefore=JSON.stringify(records);
 const result=ToolAutomation.assess({tools:records},parsed),evidence=parsed.identityEvidence;
 const values=evidence.observedSerialEvidence.map(proof=>proof.serial);
 const checks=[];
 if(result.eligible!==eligible)checks.push('eligible='+result.eligible+' expected '+eligible);
 if(eligible){
  const target=records.find(tool=>tool.id===expectedTool);
  if(result.toolId!==expectedTool)checks.push('wrong registered row');
  if(result.matchMethod!==method)checks.push('method='+result.matchMethod+' expected '+method);
  if(result.name!==target.name||result.serial!==target.serial)checks.push('canonical Excel identity changed');
  if(method==='numeric-cell-serial'&&result.observedSerial!==observed)checks.push('actual zero-prefixed value lost');
 }else if(!result.reasons.length)checks.push('held result has no reason');
 for(const value of mustKeep)if(!values.includes(value))checks.push('valid identity lost '+value);
 for(const value of mustDrop)if(values.includes(value))checks.push('unrelated single probe retained '+value);
 if(JSON.stringify(parsed)!==before)checks.push('raw OCR proof mutated by assessment');
 if(JSON.stringify(records)!==referenceBefore)checks.push('reference workbook records mutated');
 cases[name]={eligible:result.eligible,method:result.matchMethod,reasons:result.reasons,checks};
 if(checks.length)failures.push({name,checks,result});
}
run('proven_numeric_cell_and_original_serial_remain_valid',{mustKeep:[observed]});
run('numeric_original_treats_a_single_unknown_probe_like_an_exact_serial',{
 reads:[original(observed),probe(noise)],mustKeep:[observed],mustDrop:[noise]});
run('single_numeric_probe_itself_remains_valid',{reads:[numericProbe()],mustKeep:[observed]});
run('single_numeric_probe_survives_while_unknown_probe_is_excluded',{
 reads:[numericProbe(),probe(noise)],mustKeep:[observed],mustDrop:[noise]});
run('numeric_priority_does_not_depend_on_probe_order',{
 reads:[probe(noise),numericProbe()],mustKeep:[observed],mustDrop:[noise]});
run('independently_repeated_numeric_reads_still_exclude_lone_noise',{
 reads:[numericProbe(),probe(observed,'identity-probe-fast-token-serial-second',92),probe(noise)],
 mustKeep:[observed],mustDrop:[noise]});
run('explicit_original_numeric_label_has_the_same_priority',{
 reads:[original('Serial: '+observed),probe(noise)],mustKeep:[observed],mustDrop:[noise]});
run('exact_literal_serial_keeps_existing_single_probe_behavior',{
 reads:[original(tools[0].serial),probe(noise)],method:'serial',mustKeep:[tools[0].serial],mustDrop:[noise]});
run('exact_literal_probe_keeps_existing_single_probe_behavior',{
 reads:[probe(tools[0].serial,'identity-probe-fast-exact',93),probe(noise)],method:'serial',mustKeep:[tools[0].serial],mustDrop:[noise]});
run('actual_other_registered_serial_is_never_unknown_noise',{
 reads:[numericProbe(),probe(tools[1].serial,'identity-probe-fast-known-other',94)],eligible:false,
 mustKeep:[observed,tools[1].serial]});
run('explicit_wrong_full_serial_is_not_suppressed',{
 reads:[numericProbe(),probe('Serial: '+noise,'identity-probe-fast-labelled-other',94)],eligible:false,
 mustKeep:[observed,noise]});
run('two_actual_numeric_equivalents_cannot_choose_one_tool',{
 reads:[numericProbe(),probe(otherObserved,'identity-probe-fast-numeric-other',94)],eligible:false,
 mustKeep:[observed,otherObserved]});
run('repeated_unknown_probe_stays_contrary_evidence',{
 reads:[numericProbe(),probe(noise,'identity-probe-noise-a',94),probe(noise,'identity-probe-noise-b',95)],
 eligible:false,mustKeep:[observed,noise]});
run('text_cell_cannot_borrow_numeric_zero_equivalence',{
 reads:[original(observed),probe(noise)],records:tools.map(tool=>tool.id==='numeric-a'?{...tool,numericSerial:false}:tool),eligible:false});
run('text_cell_still_supports_its_own_exact_literal_serial',{
 reads:[original(tools[0].serial)],records:tools.map(tool=>tool.id==='numeric-a'?{...tool,numericSerial:false}:tool),method:'serial'});
run('duplicate_numeric_storage_values_do_not_become_unique',{
 records:[...tools,{...tools[0],id:'duplicate',name:tools[1].name}],eligible:false});
run('missing_numeric_cell_provenance_does_not_infer_zeroes',{
 records:tools.map(tool=>tool.id==='numeric-a'?{...tool,numericSerial:undefined}:tool),eligible:false});
run('low_confidence_serial_cannot_create_numeric_priority',{
 reads:[read(observed,'identity-crop-serial',69),probe(noise)],eligible:false});
run('zero_confidence_serial_cannot_create_numeric_priority',{
 reads:[read(observed,'identity-crop-serial',0),probe(noise)],eligible:false});
run('high_pass_average_cannot_replace_low_actual_word_confidence',{
 reads:[read(observed,'identity-crop-serial',98,{words:[{text:observed,confidence:20}]}),probe(noise)],eligible:false});
run('learned_serial_evidence_cannot_prove_numeric_storage_identity',{
 mutate:parsed=>{for(const proof of parsed.identityEvidence.observedSerialEvidence)if(proof.serial===observed)proof.learned=true;},eligible:false});
run('aggregate_confidence_without_actual_source_confidence_is_not_numeric_proof',{
 mutate:parsed=>{for(const proof of parsed.identityEvidence.observedSerialEvidence)if(proof.serial===observed)proof.sourceReadings=[];},eligible:false});
run('anonymous_reading_is_not_actual_numeric_proof',{
 reads:[read(observed,'text',98),probe(noise)],eligible:false});
run('date_crop_cannot_supply_numeric_identity',{
 reads:[read(observed,'date-crop-serial',98),probe(noise)],eligible:false});
run('incomplete_identity_recovery_cannot_become_numeric_priority',{
 reads:[probe(observed,'identity-probe-incomplete',98)],eligible:false});
console.log(JSON.stringify({case_count:Object.keys(cases).length,cases,failures,
 scope:'Controlled actual parser and numeric provenance policy; no photo accuracy or company-write claim'},null,2));
assert.deepEqual(failures,[],'numeric identity/probe priority regressions');
"""


if __name__ == '__main__':
    subprocess.run(['node', '-e', CASES], cwd=ROOT, check=True)
