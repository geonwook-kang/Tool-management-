"""Name-first literal fragments and serial-first canonical Excel owners.

Production OCR parsing and assessment use controlled observations in Node.
No photograph, company workbook, OCR accuracy or cloud access is simulated.
"""
from pathlib import Path
import subprocess


ROOT = Path(__file__).resolve().parents[1]


CASES = r"""
const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
global.window=global;
for(const source of ['web/ocr.js','web/automation.js'])vm.runInThisContext(fs.readFileSync(source,'utf8'),{filename:source});
const base=[
 {id:'a',name:'홍길동',serial:'4829671930',months:12,last:'2026-02-14',due:'2027-02-14',photoCount:1},
 {id:'b',name:'홍길동',serial:'8531046827',months:6,last:'2026-02-14',due:'2026-08-14',photoCount:1},
 {id:'c',name:'김하늘',serial:'5529671444',months:24,last:'2026-02-14',due:'2028-02-14',photoCount:1}
];
const cases={};
function run(name,{fragment='29671',owner='홍길동',ownerConfidence=95,serialConfidence=95,count=2,
 records=base,due='2028-02-14',otherOwner='',extraReads=[],mutate=null,eligible=false,method='owner-partial-serial',id='a'}={}){
 const passes=[{text:'교정일자 2027.02.14\n차기교정예정일자 '+due,source:'original',confidence:96}];
 if(owner)passes.push({text:owner,source:'identity-crop-owner',confidence:ownerConfidence});
 if(otherOwner)passes.push({text:otherOwner,source:'identity-crop-second-owner',confidence:95});
 for(let index=0;index<count;index++)passes.push({text:'Serial: '+fragment,source:'identity-crop-fragment-'+index,confidence:serialConfidence});
 passes.push(...extraReads);
 const parsed=ToolOCR.parse(passes.map(pass=>pass.text).join('\n'),records,{passes});if(mutate)mutate(parsed);
 const original=JSON.stringify(parsed),registered=JSON.stringify(records);
 const result=ToolAutomation.assess({tools:records},parsed);
 assert.equal(result.eligible,eligible,name+': '+JSON.stringify(result));
 if(eligible){
  assert.equal(result.toolId,id,name);assert.equal(result.matchMethod,method,name);
  const target=records.find(tool=>tool.id===id);assert.equal(result.name,target.name,name);assert.equal(result.serial,target.serial,name);
  if(method==='owner-partial-serial')assert.equal(result.observedSerial,fragment,name);
 }else assert.ok(result.reasons.length,name);
 assert.equal(JSON.stringify(parsed),original,name+': raw proof changed');assert.equal(JSON.stringify(records),registered,name+': Excel reference changed');
 cases[name]={eligible:result.eligible,method:result.matchMethod,reasons:result.reasons};
}
run('clear_owner_filters_a_five_digit_fragment_shared_with_another_person',{eligible:true});
run('clear_owner_also_excludes_another_persons_retired_collision',{records:base.map(tool=>tool.id==='c'?{...tool,due:'Disposed'}:tool),eligible:true});
run('same_fragment_without_owner_stays_globally_ambiguous',{owner:''});
run('owner_below_confidence_threshold_cannot_narrow_the_list',{ownerConfidence:69});
run('two_different_clear_owners_cannot_narrow_the_list',{otherOwner:'김하늘'});
run('same_owner_two_matching_tools_cannot_be_separated_by_cycle',{records:base.map(tool=>tool.id==='b'?{...tool,serial:'8829671555'}:tool)});
run('retired_same_owner_tool_still_counts_for_ambiguity',{records:base.map(tool=>tool.id==='b'?{...tool,serial:'8829671555',due:'Disposed'}:tool)});
run('one_fragment_read_cannot_narrow_a_global_tie',{count:1});
run('a_low_confidence_fragment_is_not_proof',{serialConfidence:69,records:base.map(tool=>tool.id==='b'?{...tool,months:12}:tool)});
run('half_confident_sources_cannot_borrow_the_aggregate',{mutate:parsed=>{
 for(const proof of parsed.identityEvidence.observedSerialEvidence)if(proof.serial==='29671')proof.sourceReadings[1].confidence=30;
}});
run('duplicated_source_names_do_not_prove_two_reads',{mutate:parsed=>{
 for(const proof of parsed.identityEvidence.observedSerialEvidence){proof.passSources=['identity-crop-one'];proof.sourceReadings=[{source:'identity-crop-one',confidence:95},{source:'identity-crop-one',confidence:95}];}
}});
run('learned_owner_alias_cannot_replace_a_clear_owner',{mutate:parsed=>{for(const proof of parsed.identityEvidence.nameMatches)proof.learned=true;}});
run('wrong_whole_serial_is_not_truncated_or_edited',{fragment:'4829671938'});
run('incorrect_visible_fragment_is_not_repaired',{fragment:'29678'});
run('registered_full_serial_of_another_person_blocks_owner_partial',{records:base.map(tool=>tool.id==='c'?{...tool,serial:'29671'}:tool)});
run('registered_full_serial_of_same_owner_also_blocks_partial_collision',{records:base.map(tool=>tool.id==='b'?{...tool,serial:'29671'}:tool)});
run('clear_other_full_serial_blocks_an_owning_name',{fragment:base[2].serial});
run('complete_owner_serial_and_its_literal_fragment_can_agree',{eligible:true,method:'serial',extraReads:[{text:'Serial: '+base[0].serial,source:'identity-crop-whole',confidence:95}]});
run('no_owner_global_unique_full_serial_remains_automatic',{owner:'',fragment:base[0].serial,count:1,eligible:true,method:'serial'});
run('no_owner_global_unique_partial_remains_automatic',{owner:'',fragment:'71930',eligible:true,method:'partial-serial'});
run('name_with_four_actual_digits_keeps_the_existing_owner_path',{fragment:'4829',eligible:true});
run('four_digit_fragment_can_be_shared_with_other_owners',{fragment:'4829',records:base.map(tool=>tool.id==='c'?{...tool,serial:'9948295555'}:tool),eligible:true});
run('three_digits_do_not_become_owner_serial_proof',{fragment:'482'});
run('less_than_forty_percent_is_not_owner_partial_proof',{records:base.map(tool=>tool.id==='a'?{...tool,serial:'884829671930555'}:tool)});
run('wrong_cycle_cannot_override_the_visible_number',{due:'2027-08-14'});
run('missing_complete_date_pair_cannot_confirm_owner_partial',{mutate:parsed=>{parsed.dateEvidence={dates:[],ambiguous:false,method:'',confidence:null};}});
run('date_only_sources_cannot_become_owner_serial_proof',{mutate:parsed=>{
 for(const proof of parsed.identityEvidence.observedSerialEvidence){proof.passSources=['date-crop-7','date-crop-13'];proof.sourceReadings=proof.passSources.map(source=>({source,confidence:95}));}
}});
run('a_four_character_alphanumeric_fragment_uses_the_same_owner_filter',{
 fragment:'AB12',records:base.map(tool=>tool.id==='a'?{...tool,serial:'79AB12XZ45'}:tool.id==='c'?{...tool,serial:'56AB129999'}:tool),eligible:true});
run('longer_alphanumeric_partial_is_not_numeric_only',{
 fragment:'B12XZ',records:base.map(tool=>tool.id==='a'?{...tool,serial:'79AB12XZ45'}:tool.id==='c'?{...tool,serial:'00B12XZ999'}:tool),eligible:true});
run('another_owner_does_not_make_same_owner_duplicates_unique',{
 fragment:'AB12',records:base.map(tool=>tool.id==='a'?{...tool,serial:'79AB12XZ45'}:tool.id==='b'?{...tool,serial:'88AB125678'}:tool.id==='c'?{...tool,serial:'56AB129999'}:tool)});
console.log(JSON.stringify({case_count:Object.keys(cases).length,cases,
 scope:'Controlled actual OCR evidence policy; no nearest-serial editing and no photo accuracy claim'},null,2));
"""


if __name__ == '__main__':
    subprocess.run(['node', '-e', CASES], cwd=ROOT, check=True)
