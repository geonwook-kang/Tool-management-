"""Controlled OCR evidence checks for serial-only and worn serial selection.

The production parser and automatic-assessment module run unchanged in Node.
These are policy tests, not an OCR-accuracy claim. Fixtures contain no company
workbook, photograph, names, identifiers, or network requests.
"""
from pathlib import Path
import subprocess


ROOT = Path(__file__).resolve().parents[1]

CASES = r"""
const assert=require('node:assert/strict'), fs=require('node:fs'), vm=require('node:vm');
global.window=global;
for(const path of ['web/ocr.js','web/automation.js'])vm.runInThisContext(fs.readFileSync(path,'utf8'),{filename:path});
const tools=[
 {id:'a',name:'홍길동',serial:'4829671930',months:12,last:'2025-01-01',due:'2026-01-01',photoCount:1},
 {id:'b',name:'김하늘',serial:'5631408271',months:6,last:'2025-01-01',due:'2025-07-01',photoCount:1},
 {id:'c',name:'홍길동',serial:'9085471632',months:24,last:'2025-01-01',due:'2027-01-01',photoCount:1},
 {id:'d',name:'박지수',serial:'MTX740628',months:12,last:'2025-01-01',due:'2026-01-01',photoCount:1}
];
const dateText='교정일자 2027.02.14\n차기교정예정일자 2028.02.14';
const cases={};
function proof(groups,sources=['identity-crop-top-6','identity-crop-top-11']){
 return {groups,kind:'standalone',confidence:96,gapVerified:true,passSources:sources,
  sourceReadings:sources.map(source=>({source,confidence:96}))};
}
function run(name,{value='71930',owner='',records=tools,confidence=95,count=2,source='identity-crop-top',
 date=dateText,extra='',mutate=null,eligible=false,method='',id='a'}={}){
 const passes=[{text:date,source:'original',confidence:95}];
 if(owner)passes.push({text:owner,source:'identity-crop-owner',confidence:95});
 for(let i=0;i<count;i++)passes.push({text:'Serial: '+value+(extra?'\n'+extra:''),source:source+'-'+i,confidence});
 const result=ToolOCR.parse(passes.map(pass=>pass.text).join('\n'),records,{passes});
 if(mutate)mutate(result);
 const before=JSON.stringify(result.identityEvidence),a=ToolAutomation.assess({tools:records},result);
 assert.equal(a.eligible,eligible,name+': '+JSON.stringify(a));
 if(eligible){assert.equal(a.toolId,id,name);assert.equal(a.matchMethod,method,name);assert.equal(a.serial,records.find(t=>t.id===id).serial,name);}
 else assert.ok(a.reasons.length,name+': missing explanation');
 assert.equal(JSON.stringify(result.identityEvidence),before,name+': input evidence changed');
 cases[name]={eligible:a.eligible,method:a.matchMethod,reasons:a.reasons};
}
run('one_complete_serial_without_any_name',{value:tools[0].serial,count:1,eligible:true,method:'serial'});
run('complete_alphanumeric_serial_without_name',{value:tools[3].serial,count:1,eligible:true,method:'serial',id:'d'});
run('five_visible_digits_without_name',{eligible:true,method:'partial-serial'});
run('five_visible_internal_digits_without_name',{value:'29671',eligible:true,method:'partial-serial'});
run('unknown_name_low_confidence_cannot_block_clear_serial',{value:tools[0].serial,eligible:true,method:'serial',mutate:r=>{
 r.identityEvidence.observedNameEvidence.push({name:'김하늘',kind:'labelled',confidence:24});
}});
run('confident_other_registered_owner_is_a_conflict',{value:tools[0].serial,owner:'김하늘'});
run('confident_other_labelled_owner_is_a_conflict',{value:tools[0].serial,extra:'담당자: 이해솔'});
run('complete_other_serial_is_a_conflict',{value:tools[0].serial,extra:'Serial: '+tools[1].serial});
run('only_one_fragment_read_is_insufficient',{count:1});
run('fragment_confidence_69_is_insufficient',{confidence:69});
run('five_digits_below_half_the_registered_length',{records:tools.map(t=>t.id==='a'?{...t,serial:'8884829671930'}:t)});
run('four_digits_without_name_are_insufficient',{value:'1930'});
run('same_literal_fragment_in_another_persons_tool_is_ambiguous',{records:[...tools,{...tools[1],id:'e',serial:'1271930555'}]});
run('retired_matching_fragment_is_still_ambiguous',{records:[...tools,{...tools[1],id:'e',serial:'1271930555',due:'Disposed'}]});
run('clear_owner_filters_other_persons_fragment_even_with_different_cycle',{owner:'홍길동',records:[...tools,{...tools[1],id:'e',serial:'1271930555',months:6}],eligible:true,method:'owner-partial-serial'});
run('different_calibration_cycle_cannot_resolve_same_owner_ambiguous_fragment',{owner:'홍길동',records:[...tools,{...tools[1],id:'e',name:tools[0].name,serial:'1271930555',months:6}]});
run('high_confidence_missing_date_pair_is_not_guessed',{date:'교정일자 2027.02'});
run('fragment_cycle_mismatch_requires_review',{date:'교정일자 2027.02.14\n차기교정예정일자 2027.08.14'});
run('whole_unknown_serial_is_not_truncated_to_registered_suffix',{value:'887193099'});
run('one_wrong_digit_is_not_replaced_to_force_a_match',{value:'4829671938'});
run('date_crop_cannot_supply_serial_identity',{source:'date-crop'});
run('label_probe_cannot_supply_serial_identity',{source:'label-probe'});
run('anonymous_sources_cannot_supply_repeated_fragment_proof',{mutate:r=>{
 for(const p of r.identityEvidence.observedSerialEvidence){p.passSources=['text'];p.sourceReadings=[{source:'text',confidence:95}];}
}});
run('same_source_twice_is_not_two_independent_reads',{mutate:r=>{
 for(const p of r.identityEvidence.observedSerialEvidence){p.passSources=['identity-crop-one'];p.sourceReadings=[{source:'identity-crop-one',confidence:95},{source:'identity-crop-one',confidence:95}];}
}});
run('low_per_source_confidence_cannot_borrow_aggregate_score',{mutate:r=>{
 for(const p of r.identityEvidence.observedSerialEvidence)for(const reading of p.sourceReadings)reading.confidence=40;
}});
run('legacy_fragment_evidence_without_actual_readings_is_insufficient',{mutate:r=>{
 for(const p of r.identityEvidence.observedSerialEvidence)delete p.sourceReadings;
}});
run('repeated_unvarying_digits_are_not_distinctive',{value:'00000',records:[{...tools[0],serial:'7400000000'}]});
run('calendar_date_digits_are_not_partial_identity',{value:'20270214',records:[{...tools[0],serial:'992027021488'}]});
run('two_visible_groups_with_erased_middle_digits',{value:'4827193',eligible:true,method:'partial-serial-groups',mutate:r=>r.identityEvidence.observedSerialGroups=[proof(['482','7193'])]});
run('groups_reconcile_an_individual_longer_fragment',{value:'48296',eligible:true,method:'partial-serial',mutate:r=>r.identityEvidence.observedSerialGroups=[proof(['48296','930'])]});
run('groups_from_only_one_actual_read_cannot_be_used',{value:'4827193',mutate:r=>r.identityEvidence.observedSerialGroups=[proof(['482','7193'],['identity-crop-one'])]});
run('groups_need_verified_physical_gap',{value:'4827193',mutate:r=>r.identityEvidence.observedSerialGroups=[{...proof(['482','7193']),gapVerified:false}]});
run('groups_from_dates_do_not_match',{value:'4827193',mutate:r=>r.identityEvidence.observedSerialGroups=[proof(['482','7193'],['date-crop-6','date-crop-11'])]});
run('groups_cannot_use_learned_aliases',{value:'4827193',mutate:r=>r.identityEvidence.observedSerialGroups=[{...proof(['482','7193']),learned:true}]});
run('groups_need_each_real_read_high_confidence',{value:'4827193',mutate:r=>{
 const p=proof(['482','7193']);p.sourceReadings[1].confidence=40;r.identityEvidence.observedSerialGroups=[p];
}});
run('groups_need_unique_real_pass_sources',{value:'4827193',mutate:r=>{
 r.identityEvidence.observedSerialGroups=[proof(['482','7193'],['identity-crop-one','identity-crop-one'])];
}});
run('groups_ambiguous_with_another_owner_are_held',{value:'4827193',records:[...tools,{...tools[1],id:'e',serial:'4821171938'}],mutate:r=>r.identityEvidence.observedSerialGroups=[proof(['482','7193'])]});
run('groups_ambiguous_with_retired_tool_are_held',{value:'4827193',records:[...tools,{...tools[1],id:'e',serial:'4821171938',due:'Disposed'}],mutate:r=>r.identityEvidence.observedSerialGroups=[proof(['482','7193'])]});
run('group_order_cannot_be_reversed',{value:'7193482',mutate:r=>r.identityEvidence.observedSerialGroups=[proof(['7193','482'])]});
run('groups_with_less_than_sixty_percent_coverage_are_held',{value:'4827193',records:tools.map(t=>t.id==='a'?{...t,serial:'88482967193099'}:t),mutate:r=>r.identityEvidence.observedSerialGroups=[proof(['482','7193'])]});
run('a_known_shorter_concatenated_serial_cannot_steal_gap_reading',{value:'4827193',records:[...tools,{...tools[1],id:'short',serial:'4827193',months:12}],mutate:r=>r.identityEvidence.observedSerialGroups=[proof(['482','7193'])]});
run('single_gap_reading_also_blocks_wrong_known_concatenation',{value:'4827193',records:[...tools,{...tools[1],id:'short',serial:'4827193',months:12}],mutate:r=>r.identityEvidence.observedSerialGroups=[proof(['482','7193'],['identity-crop-one'])]});
run('two_disagreeing_credible_group_readings_are_held',{value:'4827193',mutate:r=>r.identityEvidence.observedSerialGroups=[proof(['482','7193']),proof(['563','08271'])]});
run('unmatched_credible_group_reading_cannot_be_ignored',{value:tools[0].serial,mutate:r=>r.identityEvidence.observedSerialGroups=[proof(['874','2160'])]});
// The gap channel itself must originate in real same-row OCR word geometry.
// No fabricated groupedEvidence is injected in the following cases.
function geometry(name,{gap=40,secondY=100,count=2,label='Serial:',source='identity-crop-tag',boxes=true,
 low=false,records=tools,eligible=false,expectGroups=false}={}){
 const passes=[{text:dateText,source:'original',confidence:96}];
 for(let i=0;i<count;i++){
  const words=[...(label?[{text:label,confidence:96,bbox:{x0:0,y0:100,x1:50,y1:120}}]:[]),
   {text:'482',confidence:96,bbox:{x0:80,y0:100,x1:110,y1:120}},
   {text:'7193',confidence:low?35:96,bbox:{x0:110+gap,y0:secondY,x1:150+gap,y1:secondY+20}}];
  if(!boxes)for(const word of words)delete word.bbox;
  passes.push({text:[label,'482 7193'].filter(Boolean).join(' '),source:source+'-'+i,confidence:96,words});
 }
 const parsed=ToolOCR.parse(passes.map(p=>p.text).join('\n'),records,{passes}),groups=parsed.identityEvidence.observedSerialGroups||[];
 assert.equal(groups.length>0,expectGroups,name+': '+JSON.stringify(groups));
 const a=ToolAutomation.assess({tools:records},parsed);
 assert.equal(a.eligible,eligible,name+': '+JSON.stringify(a));
 if(eligible){assert.equal(a.toolId,'a',name);assert.equal(a.matchMethod,'partial-serial-groups',name);assert.deepEqual(groups[0].groups,['482','7193']);}
 cases[name]={eligible:a.eligible,method:a.matchMethod,reasons:a.reasons};
}
geometry('actual_same_row_geometry_forms_repeated_groups',{eligible:true,expectGroups:true});
geometry('identity_crop_can_find_unlabelled_serial_groups',{label:'',eligible:true,expectGroups:true});
geometry('ordinary_full_photo_needs_explicit_serial_label',{label:'',source:'original'});
geometry('full_photo_explicit_serial_label_can_prove_groups',{source:'original',eligible:true,expectGroups:true});
geometry('text_only_spaced_digits_do_not_prove_an_erased_gap',{boxes:false});
geometry('digits_on_two_different_rows_are_not_joined',{secondY:155});
geometry('normal_small_ocr_word_spacing_is_not_an_erased_gap',{gap:4});
geometry('huge_gap_is_not_one_serial_label',{gap:120});
geometry('single_geometric_read_cannot_auto_write',{count:1,expectGroups:true});
geometry('weak_geometric_word_cannot_borrow_pass_confidence',{low:true,expectGroups:true});
geometry('certificate_row_is_not_serial_identity',{label:'Certificate No'});
geometry('asset_row_is_not_serial_identity',{label:'Asset Number'});
geometry('date_crop_group_geometry_is_not_identity',{source:'date-crop'});
geometry('physical_gap_cannot_collapse_into_another_known_serial',{expectGroups:true,records:[...tools,{...tools[1],id:'short',serial:'4827193',months:12}]});
// Exercise every synthetic owner without requiring any readable name.
for(const t of tools){const months=t.months,last='2027-02-14';let y=2027,m=2+months;while(m>12){m-=12;y++;}
 run('all_registered_serials_without_owner_'+t.id,{value:t.serial,count:1,id:t.id,eligible:true,method:'serial',date:'교정일자 '+last+'\n차기교정예정일자 '+y+'.'+String(m).padStart(2,'0')+'.14'});
}
console.log(JSON.stringify({case_count:Object.keys(cases).length,cases,scope:'Production parser/assessment with controlled OCR observations; no photo OCR accuracy claim'},null,2));
"""


def main():
    result = subprocess.run(["node", "-"], cwd=ROOT, input=CASES, text=True, capture_output=True)
    print(result.stdout, end="")
    if result.returncode:
        print(result.stderr, end="")
        raise SystemExit(result.returncode)


if __name__ == "__main__":
    main()
