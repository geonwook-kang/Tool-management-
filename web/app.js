(() => {
'use strict';
window.addEventListener('DOMContentLoaded',()=>{
const $=id=>document.getElementById(id);
const state={model:null,workbookName:'',photos:[],selectedId:null,busy:false,preparing:false,preparedResult:null,preparationError:'',revision:0,cancelRequested:false,source:{connected:false,needsPermission:false,needsReconnect:false,name:'',version:'',epoch:0,saving:false,error:'',savedAt:'',saveCount:0,hasBackup:false}};
const current=()=>state.photos.find(p=>p.id===state.selectedId), approved=()=>state.photos.filter(p=>p.review),automatic=()=>$('automatic-mode').checked;
const engineFailed=p=>!!(p&&!p.review&&(p.result?.engineFailure||p.result?.errorCode==='engine-init'||p.result?.code==='engine-init'));
const status=(msg,kind='')=>{$('status').textContent=msg;$('status').className=kind;};
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const methods={serial:'시리얼 번호 일치',name:'이름 일치','name-cycle':'이름·엑셀 교정 주기 일치','partial-serial':'부분 시리얼·엑셀 교정 주기 일치','owner-partial-serial':'이름·부분 번호·엑셀 교정 주기 일치','smax-serial':'SMAX 등록 시리얼 일치','smax-due':'SMAX 차기 교정일 → Cal Next','unchanged-dates':'교정 날짜 유지','remembered-serial':'기억한 시리얼 인식 패턴','remembered-name':'기억한 이름 인식 패턴',labelled:'교정일자 → Cal Last · 차기교정예정일자 → Cal Next','repeated-dates':'두 번 일치한 날짜','calculated-cycle':'엑셀 주기로 계산'};
const photoKind=p=>p?.photoKindOverride||p?.draft?.photoKind||p?.review?.photoKind||p?.assessment?.photoKind||(/smax/i.test(p?.result?.documentType||'')?'smax':'calibration');
const photoSlot=p=>photoKind(p)==='smax'?'SMAX Tool Photo':'Label Photo';
let seq=0,timer,restoreTask=Promise.resolve();
const memoryEnabled=()=>$('remember-matches').checked;
function refreshMemory(){const info=window.ToolProfiles?.inspect();$('memory-status').textContent=info?`기억한 툴 ${info.count}개 · 인식 패턴 ${info.aliasCount}개${!memoryEnabled()?' · 활용 꺼짐':!info.persistent?' · 이번 실행 중에만 기억':''}`:'매칭 기억 기능을 준비하지 못했습니다. 기본 인식은 계속 사용할 수 있습니다.';$('clear-matches').disabled=state.busy||state.preparing||state.source.saving||!(info?.count||info?.conflictCount);$('remember-matches').disabled=state.busy||state.preparing||state.source.saving;}
function rememberApproved(entries){if(!memoryEnabled()||!window.ToolProfiles)return;for(const {p} of entries)if(photoKind(p)!=='smax')try{ToolProfiles.remember(state.model,p);}catch(_){}refreshMemory();}
const pendingRows=()=>approved().filter(p=>!state.source.connected||!p.saved);
function sourceFailure(error){const info=ToolConnectedFile.inspect();state.source.error=error.message;state.source.needsPermission=!!info.needsPermission;state.source.needsReconnect=!!info.needsReconnect;if(info.needsPermission||info.needsReconnect)state.source.connected=false;}

function refreshSource(){
 const c=state.source,locked=state.busy||state.preparing||c.saving;
 for(const id of ['connect-onedrive','reconnect-onedrive','disconnect-onedrive','retry-connected-save'])$(id).disabled=locked;
 $('reconnect-onedrive').hidden=!c.needsPermission;
 $('disconnect-onedrive').hidden=!(c.connected||c.needsPermission||c.needsReconnect);
 $('retry-connected-save').hidden=!c.error||!c.connected;
 $('download-backup').hidden=!c.hasBackup;
 $('connection-status').textContent=c.saving?'연결된 엑셀에 저장 중…':c.error?'자동 저장 확인 필요: '+c.error:c.connected?c.name+' · 연결됨'+(c.savedAt?' · 이 PC 파일 저장 '+c.savedAt:''):c.needsPermission?c.name+' · 파일 권한 재연결이 필요합니다.':'OneDrive 동기화 폴더의 엑셀을 한 번 연결하면 사진만 추가할 수 있습니다.';
 $('connection-status').className='summary'+(c.error?' source-error':'');
 for(const el of $('review-panel').querySelectorAll('input,select,button'))el.disabled=state.busy||c.saving;
}
function resetPhotoConnection(){for(const p of state.photos)Object.assign(p,{saved:null,review:null,approval:null,result:null,draft:{},assessment:null,error:null,edited:false,manualEntry:false});}
async function adoptSource(snapshot,{reset=false}={}){
 const model=await ToolWorkbook.load(snapshot.bytes),c=state.source;
 state.model=model;state.workbookName=snapshot.file.name;c.connected=true;c.needsPermission=false;c.needsReconnect=false;c.name=snapshot.file.name;c.version=snapshot.version;
 if(reset)resetPhotoConnection();
 else for(const p of state.photos){
  if(p.saved){const matches=model.tools.filter(t=>ToolWorkbook.normalizeSerial(t.serial)===ToolWorkbook.normalizeSerial(p.saved.serial)&&t.name===p.saved.name);if(matches.length===1){p.draft.toolId=matches[0].id;if(p.review)p.review.toolId=matches[0].id;}continue;}
  if(p.approval==='manual'){
   const hits=model.tools.filter(t=>ToolWorkbook.normalizeSerial(t.serial)===ToolWorkbook.normalizeSerial(p.draft.serial)&&t.name===p.draft.name);
   if(hits.length===1){p.draft.toolId=hits[0].id;p.review={...p.draft,...p.image};if(ToolWorkbook.validate(model,p.review).length){p.review=null;p.approval=null;}}
   else{p.review=null;p.approval=null;}
  }else if(p.result&&!p.result.failed&&!p.edited){const old=p.result;applyOCR(p,ToolOCR.parse(old.text,model.tools,{confidence:old.confidence,passes:old.passes}));}
 }
 $('workbook-summary').textContent=`${c.name} · 연결 파일 · ${model.tools.length}개 툴`;$('workbook-badge').textContent='연결됨';
}
async function refreshConnectedSource(){const snapshot=await ToolConnectedFile.read();await adoptSource(snapshot);return snapshot;}
async function connectSource(requestPermission=false){
 if(state.busy||state.preparing||state.source.saving)return;
 state.busy=true;refresh();let ok=false;
 try{
  if(requestPermission)await ToolConnectedFile.requestAccess();
  else await ToolConnectedFile.connect({validate:async file=>ToolWorkbook.load(await file.arrayBuffer())});
  const snapshot=await ToolConnectedFile.read();invalidate();state.source.epoch++;state.source.error='';if(!requestPermission){state.source.savedAt='';state.source.saveCount=0;}await adoptSource(snapshot,{reset:!requestPermission});ok=true;
  const info=ToolConnectedFile.inspect();state.source.hasBackup=!!(await ToolConnectedFile.takeBackup());
  status('이 PC의 동기화 엑셀을 연결했습니다. 이후에는 사진만 추가하면 자동 저장합니다. OneDrive의 서버 동기화 상태는 OneDrive에서 확인하세요.'+(info.persistent===false?' 이 브라우저에서는 연결을 기억하지 못해 다음 실행 때 다시 선택해야 합니다.':''),'success');
 }catch(e){status('엑셀 연결 실패: '+e.message,'error');}
 finally{state.busy=false;refresh();render();}
 if(ok&&automatic()&&state.photos.some(p=>!p.saved))await recognize(state.photos.filter(p=>!p.saved).map(p=>p.id));
}
async function restoreSource(){
 state.restoring=true;state.busy=true;refresh();
 try{
  const info=await ToolConnectedFile.restore();state.source.hasBackup=!!(await ToolConnectedFile.takeBackup());
  if(info.connected){state.busy=true;await refreshConnectedSource();status('기억한 엑셀을 불러왔습니다. 사진만 추가하면 됩니다.','success');}
  else if(info.needsPermission){state.source.needsPermission=true;state.source.name=info.name||'';status('기억한 엑셀을 사용하려면 「권한 다시 연결」을 눌러 주세요.');}
  else if(info.needsReconnect){state.source.needsReconnect=true;state.source.name=info.name||'';state.source.error=info.statusMessage||'이전에 연결한 파일을 다시 선택해 주세요.';status(state.source.error,'error');}
 }catch(e){status('이전 엑셀 연결을 복원하지 못했습니다. 다시 연결해 주세요. '+e.message,'error');}
 finally{state.restoring=false;state.busy=false;refresh();render();}
}
async function retrySource(){
 if(state.busy||state.preparing||!state.source.connected)return;
 state.busy=true;invalidate();state.source.error='';refresh();
 try{await refreshConnectedSource();reconcile();}catch(e){sourceFailure(e);status(e.message,'error');}
 finally{state.busy=false;refresh();render();}
 if(!state.source.error)await prepareResult();
}

function invalidate(){state.revision++;state.preparationError='';if(state.preparedResult)URL.revokeObjectURL(state.preparedResult.url);state.preparedResult=null;$('automatic-result-download').hidden=true;$('automatic-result-download').removeAttribute('href');}
function refresh(){
 refreshSource();
 refreshMemory();
 const selected=approved(),autos=selected.filter(p=>p.approval==='automatic').length,engineErrors=state.photos.filter(engineFailed).length,pending=state.photos.length-selected.length-engineErrors;
 $('photo-count').textContent=state.photos.length+'장';
 $('photo-list').innerHTML=state.photos.length?state.photos.map(p=>`<button class="photo-card ${p.id===state.selectedId?'selected':''}" data-id="${p.id}" ${state.busy?'disabled':''}><img src="${p.url}" alt=""><span class="info"><span class="filename">${esc(p.file.name)}</span><span class="photo-state ${p.review?'approved':''}">${p.saved?'✓ 연결 파일 저장됨':p.approval==='automatic'?'✓ 자동 반영':p.review?'✓ 보정 후 반영':engineFailed(p)?'⚠ 인식 엔진 실행 실패':p.result||p.error?'⚠ 확인 필요':'처리 대기'}</span></span></button>`).join(''):'<div class="empty">사진을 추가해 주세요.</div>';
 $('photo-list').querySelectorAll('[data-id]').forEach(b=>b.onclick=()=>selectPhoto(b.dataset.id));
 for(const id of ['workbook','photos','automatic-mode'])$(id).disabled=state.busy||state.source.saving;
 $('photos').disabled=state.busy||state.source.saving||state.source.needsPermission||state.source.needsReconnect;
 $('recognize-all').disabled=state.busy||!state.model||!state.photos.length;
 $('cancel-recognition').hidden=!state.busy||!state.photos.length;
 $('export').disabled=state.busy||state.preparing||(!state.preparedResult&&!(state.preparationError&&selected.length));$('export').textContent=state.preparationError?'엑셀 다시 생성·저장':state.source.connected?'저장된 엑셀 사본 받기 ↓':'자동 생성된 엑셀 저장 ↓';$('download-log').disabled=state.busy||!selected.length;
 $('process-summary').textContent=`자동 반영 ${autos}장 · 보정 반영 ${selected.length-autos}장 · 확인 필요 ${pending}장${engineErrors?' · 엔진 오류 '+engineErrors+'장':''}`;
 $('export-summary').textContent=state.preparing?'날짜와 사진을 엑셀에 반영하는 중…':state.source.error?'연결 파일 자동 저장 확인 필요':state.source.connected&&state.source.savedAt?`${selected.length}장 처리됨 · 연결 파일 자동 저장 완료`:state.preparedResult?`${selected.length}개 툴 반영 완료 · 수정본 준비됨`:`반영할 사진 ${selected.length}장 / 전체 ${state.photos.length}장`;
 $('review-counter').textContent=engineErrors?'엔진 오류 '+engineErrors+'장':pending?'확인 필요 '+pending+'장':selected.length?'자동 처리 완료':'대기';
 const isAuto=current()?.approval==='automatic';$('manual-confirm-row').hidden=isAuto;$('save-review').hidden=isAuto;
 if(state.preparedResult){const a=$('automatic-result-download');a.href=state.preparedResult.url;a.download=state.preparedResult.name;a.textContent=state.source.connected&&!state.source.error?'연결 파일에 저장 완료 · 사본 다운로드 (선택)':'자동 반영된 엑셀 저장 ↓';a.hidden=false;}
}
function saveDraft(){const p=current();if(p&&state.model&&!$('review-panel').hidden&&!$('review-form').hidden)p.draft={...p.draft,...Object.fromEntries(['name','serial','last','due'].map(k=>[k,$(k).value])),toolId:$('tool-select').value,photoKind:$('photo-kind').value,updateDue:$('photo-kind').value!=='smax'||$('smax-update-due').checked};}
function schedule(){clearTimeout(timer);if(approved().length)timer=setTimeout(prepareResult,450);}
function changed(){saveDraft();if(current())current().saved=null;const p=current();if(p){p.review=null;p.approval=null;p.edited=true;$('confirm').checked=false;$('draft-note').textContent='수정한 값을 확인한 뒤 보정한 내용 반영을 누르세요.';$('ocr-warning').classList.remove('auto-success');$('ocr-warning').textContent='입력값이 변경되어 이전 자동 반영을 취소했습니다.';}invalidate();refresh();schedule();}
function renderExisting(){const t=state.model?.tools.find(t=>t.id===$('tool-select').value);$('existing').textContent=t?`엑셀 위치: ${t.sheet} ${t.row}행\n기존 Cal Last: ${t.last||'미등록'} · Cal Next: ${t.due||'미등록'}\n주기: ${t.interval} ${t.cycle}`:'불확실한 사진만 항목과 날짜를 직접 확인하세요.';}
function render(){
 const p=current(),visible=!!(p&&state.model);$('review-panel').hidden=!visible;$('review-placeholder').hidden=visible;if(!visible)return;
 const failed=engineFailed(p);$('ocr-engine-error').hidden=!failed;$('review-form').hidden=failed&&!p.manualEntry;
 $('engine-error-detail').textContent=p.error||p.result?.error||'';
 $('manual-entry-toggle').textContent=p.manualEntry?'직접 입력 닫기':'직접 입력으로 진행 (선택)';
 $('preview').src=p.url;$('preview-caption').textContent=p.file.name;$('ocr-text').textContent=p.result?.text||(failed?'인식 엔진이 시작되지 않아 사진을 읽지 못했습니다.':'인식 전입니다.');
 const recommended=p.result?.candidateIds||[],d=p.draft||{},tools=[...state.model.tools].sort((a,b)=>Number(recommended.includes(b.id))-Number(recommended.includes(a.id)));
 $('tool-select').innerHTML='<option value="">확인이 필요한 툴 선택</option>'+tools.map(t=>`<option value="${esc(t.id)}">${recommended.includes(t.id)?'일치 후보 · ':''}${esc(t.sheet)} · ${esc(t.name)} · ${esc(t.serial)} (행 ${t.row})</option>`).join('');
 $('tool-select').value=d.toolId||'';for(const k of ['name','serial','last','due'])$(k).value=d[k]||'';$('confirm').checked=!!p.review;
 const kind=photoKind(p),smax=kind==='smax';$('photo-kind').value=kind;$('smax-date-option').hidden=!smax;$('smax-update-due').checked=!!d.updateDue;
 $('last').readOnly=smax;$('last').required=!smax;$('due').readOnly=smax&&!d.updateDue;$('due').required=!smax||!!d.updateDue;$('calculate-due').hidden=smax;
 $('last-label').textContent=smax?'기존 Cal Last · 유지':'교정일자 · Cal Last';
 const ownerReadings=p.result?.smaxEvidence?.ownerEnglish;
 const english=p.assessment?.ownerEnglish||(Array.isArray(ownerReadings)?ownerReadings.find(r=>r.value)?.value:typeof ownerReadings==='string'?ownerReadings:'')||'';
 $('photo-kind-note').textContent=`저장 위치: ${photoSlot(p)}`+(smax&&english?` · SMAX 담당자: ${english} · 등록 담당자: ${d.name||'시리얼 확인 후 연결'}`:'');
 $('confirm-description').textContent=smax?'사진의 시리얼·SMAX 저장 위치와 선택한 차기일 반영 여부를 확인했습니다.':'사진과 보정한 이름·시리얼·날짜가 일치함을 확인했습니다.';
 $('draft-note').textContent=p.approval==='automatic'?(smax?`✓ SMAX Tool Photo 자동 반영 · ${d.updateDue?'Cal Next: '+d.due:'교정 날짜 유지'}`:`✓ Cal Last: ${p.draft.last} · Cal Next: ${p.draft.due} · 사진 자동 반영`):p.review?'✓ 보정한 값을 반영했습니다.':'';
 $('ocr-warning').textContent=p.error||(p.approval==='automatic'?`자동 식별: ${methods[p.assessment.matchMethod]||p.assessment.matchMethod} · ${methods[p.assessment.dateMethod]||p.assessment.dateMethod}`:(p.assessment?.reasons||p.result?.warnings||[]).join('\n')||'불확실한 사진만 보정하세요.');
 $('ocr-warning').classList.toggle('auto-success',p.approval==='automatic');renderExisting();
 for(const el of $('review-panel').querySelectorAll('input,select,button'))el.disabled=state.busy;
}
function selectPhoto(id){if(state.busy)return;saveDraft();state.selectedId=id;refresh();render();}
async function readWorkbook(file){
 if(state.restoring)await restoreTask;
 if(!file||state.busy||state.source.saving)return;if(file.size>100*1024*1024){status('엑셀은 100MB 이하로 선택해 주세요.','error');return;}
 state.busy=true;refresh();status('엑셀의 툴 목록을 읽고 있습니다…');let loaded=false;
 try{const model=await ToolWorkbook.load(await file.arrayBuffer());if(state.source.connected||state.source.needsPermission||state.source.needsReconnect){await ToolConnectedFile.disconnect();Object.assign(state.source,{connected:false,needsPermission:false,needsReconnect:false,version:'',name:'',error:'',savedAt:'',epoch:state.source.epoch+1});resetPhotoConnection();}invalidate();state.model=model;state.workbookName=file.name;state.photos.forEach(p=>Object.assign(p,{review:null,approval:null,result:null,draft:{},assessment:null,error:null,edited:false,manualEntry:false}));
 $('workbook-summary').textContent=`${file.name} · ${new Set(model.tools.map(t=>t.sheet)).size}개 지역 · ${model.tools.length}개 툴`;$('workbook-badge').textContent=model.tools.length+'개 툴';loaded=true;status('엑셀을 불러왔습니다. 사진을 올리면 자동 처리를 시작합니다.','success');
 }catch(e){status('엑셀 읽기 실패: '+e.message,'error');}finally{state.busy=false;refresh();render();}
 if(loaded&&automatic()&&state.photos.length)await recognize(state.photos.map(p=>p.id));
}
async function imageData(file){
 let image;
 if(window.createImageBitmap)image=await createImageBitmap(file);
 else image=await new Promise((resolve,reject)=>{const im=new Image(),url=URL.createObjectURL(file);im.onload=()=>{URL.revokeObjectURL(url);resolve(im);};im.onerror=()=>{URL.revokeObjectURL(url);reject(new Error('사진을 열지 못했습니다.'));};im.src=url;});
 try{if(image.width*image.height>40e6)throw new Error('사진은 4천만 픽셀 이하로 선택해 주세요.');const s=Math.min(1,2000/Math.max(image.width,image.height)),c=document.createElement('canvas');c.width=Math.round(image.width*s);c.height=Math.round(image.height*s);c.getContext('2d').drawImage(image,0,0,c.width,c.height);const blob=await new Promise(r=>c.toBlob(r,'image/jpeg',.92));if(!blob)throw new Error('사진 변환에 실패했습니다.');return {photoBytes:new Uint8Array(await blob.arrayBuffer()),photoExtension:'jpg',photoWidth:c.width,photoHeight:c.height};}finally{image.close?.();}
}
async function addPhotos(files){
 if(state.restoring)await restoreTask;
 if(state.busy||state.source.saving)return;saveDraft();state.busy=true;refresh();const ids=[],errors=[];
 try{for(const file of files){if(!/\.(jpe?g|png|webp)$/i.test(file.name)||file.size>25*1024*1024){errors.push(file.name+': JPG/PNG/WebP, 25MB 이하만 가능합니다.');continue;}if(state.photos.some(p=>p.file.name===file.name&&p.file.size===file.size&&p.file.lastModified===file.lastModified))continue;
 try{const image=await imageData(file),id=String(++seq);ids.push(id);state.photos.push({id,file,image,url:URL.createObjectURL(file),draft:{},result:null,review:null,approval:null,assessment:null,error:null,edited:false});}catch(e){errors.push(file.name+': '+e.message);}}
 if(!state.selectedId)state.selectedId=state.photos[0]?.id||null;if(ids.length)invalidate();
 }finally{state.busy=false;$('photos').value='';refresh();render();}
 if(automatic()&&state.model&&ids.length)await recognize(ids);else status(state.model?'사진이 등록되었습니다. 인식 또는 수동 반영을 선택하세요.':'사진이 등록되었습니다. 엑셀을 선택하면 자동으로 처리합니다.','success');if(errors.length)status(errors.join('\n'),'error');
}
function applyOCR(p,result){
 if(!/smax/i.test(result.documentType||'')&&memoryEnabled()&&window.ToolProfiles)try{result=ToolProfiles.enhance(state.model,result);}catch(_){}
 p.result=result;p.error=null;p.edited=false;p.manualEntry=false;p.assessment=ToolAutomation.assess(state.model,result);
 const match=state.model.tools.find(t=>t.id===p.assessment.toolId)||(result.candidateIds?.length===1?state.model.tools.find(t=>t.id===result.candidateIds[0]):null);
 p.photoKindOverride=null;const smax=/smax/i.test(result.documentType||'');
 p.draft={toolId:match?.id||'',name:(smax?p.assessment.name:result.name)||match?.name||'',serial:(smax?p.assessment.serial:result.serial)||match?.serial||'',last:p.assessment.last||result.last||'',due:p.assessment.due||result.due||'',photoKind:p.assessment.photoKind||(smax?'smax':'calibration'),updateDue:smax?!!p.assessment.updateDue:true};p.review=null;p.approval=null;
}
function reconcile(){
 if(!state.model||!automatic())return;
 for(const p of state.photos)if(p.result&&!p.edited&&!p.saved&&p.approval!=='manual'){p.assessment=ToolAutomation.assess(state.model,p.result);p.review=null;p.approval=null;}
 const eligible=state.photos.filter(p=>p.assessment?.eligible&&!p.edited&&!p.saved&&p.approval!=='manual');
 for(const p of eligible){const a=p.assessment;
 const kind=a.photoKind||'calibration',others=eligible.filter(q=>q!==p&&q.assessment.toolId===a.toolId),manual=state.photos.filter(q=>q.approval==='manual'&&!q.saved&&q.review.toolId===a.toolId);
 if(others.some(q=>(q.assessment.photoKind||'calibration')===kind)||manual.some(q=>(q.review.photoKind||'calibration')===kind)){p.assessment={...a,eligible:false,reasons:['같은 툴의 같은 사진 칸에 여러 장이 있습니다. 중복 사진을 제거하거나 직접 확인해 주세요.']};continue;}
 const writesDue=kind!=='smax'||a.updateDue===true;
 if(writesDue&&(others.some(q=>((q.assessment.photoKind||'calibration')!=='smax'||q.assessment.updateDue===true)&&q.assessment.due!==a.due)||manual.some(q=>((q.review.photoKind||'calibration')!=='smax'||q.review.updateDue===true)&&q.review.due!==a.due))){p.assessment={...a,eligible:false,reasons:['같은 툴의 라벨과 SMAX 차기 교정일이 다릅니다. 날짜를 확인해 주세요.']};continue;}
 const update={toolId:a.toolId,name:a.name,serial:a.serial,last:a.last,due:a.due,photoKind:kind,updateDue:writesDue,...p.image},errors=ToolWorkbook.validate(state.model,update);
 if(errors.length){p.assessment={...a,eligible:false,reasons:errors};continue;}p.draft={toolId:a.toolId,name:a.name,serial:a.serial,last:a.last,due:a.due,photoKind:kind,updateDue:writesDue};p.review=update;p.approval='automatic';
 }
}
async function recognize(ids){
 if(state.busy||state.source.saving||!state.model||!ids.length)return;const savesBefore=state.source.saveCount;saveDraft();invalidate();state.busy=true;state.cancelRequested=false;refresh();render();
 try{if(state.source.connected){state.source.error='';try{await refreshConnectedSource();}catch(e){sourceFailure(e);status('연결된 엑셀을 읽지 못했습니다. '+e.message,'error');return;}}for(let i=0;i<ids.length;i++){
 if(state.cancelRequested)break;const p=state.photos.find(p=>p.id===ids[i]);if(!p)continue;p.review=null;p.approval=null;
 const prefix=`${i+1}/${ids.length} · ${p.file.name}`;status(prefix+'\n오프라인 엔진으로 이름·시리얼·교정 날짜를 읽고 있습니다…');
 try{const result=await ToolOCR.recognize(p.file,state.model.tools,event=>{$('progress').style.width=100*(i+(event.progress||0))/ids.length+'%';$('runtime-status').textContent=event.status||'인식 중';status(prefix+'\n'+(event.status||'인식 중'));});
 if(state.cancelRequested)break;applyOCR(p,result);if(result.failed)p.error=result.error||result.warnings?.join('\n')||'인식을 완료하지 못했습니다.';
 }catch(e){p.result={failed:true,text:'',warnings:[]};p.error='인식 실패: '+e.message;p.assessment={eligible:false,reasons:[p.error]};}
 }reconcile();
 }finally{state.busy=false;refresh();render();}
 await prepareResult();const autos=state.photos.filter(p=>p.approval==='automatic').length,engineErrors=state.photos.filter(engineFailed).length,pending=state.photos.filter(p=>!p.review&&!engineFailed(p)).length;
 $('runtime-status').textContent=state.cancelRequested?'인식을 중지했습니다. 다시 인식할 수 있습니다.':engineErrors?'인식 엔진 실행 실패 · 해당 사진은 읽지 못했습니다.':'오프라인 처리 완료 · 외부 서버 전송 없음';
 if(state.source.error)status('연결 파일 저장 결과를 확인해 주세요. '+state.source.error+' 수정본은 별도로 저장할 수 있습니다.','error');else if(state.preparationError)status('엑셀 생성 실패: '+state.preparationError,'error');else if(engineErrors)status(`인식 엔진 실행 실패: ${engineErrors}장의 사진을 읽지 못했습니다. 이름·시리얼 입력 없이 「인식 엔진 다시 시작」을 눌러 주세요.${autos?' 자동 반영된 사진 '+autos+'장은 유지됩니다.':''}`,'error');else if(state.source.connected&&state.source.saveCount>savesBefore)status('이 PC의 연결된 엑셀에 저장했습니다. OneDrive 서버 반영 여부는 OneDrive 동기화 상태를 확인하세요.','success');else status(state.cancelRequested?'인식을 중지했습니다. 처리된 결과는 보존됩니다.':`자동 처리 완료: ${autos}장 자동 반영 · ${pending}장 확인 필요.${state.preparedResult?' 수정본이 준비되었습니다.':''}`,pending?'':'success');
}
async function prepareResult(){
 clearTimeout(timer);if(state.busy||state.preparing){schedule();return null;}const rows=pendingRows();if(!state.model||!rows.length){refresh();return null;}
 const revision=state.revision,model=state.model,sourceEpoch=state.source.epoch,sourceVersion=state.source.version,entries=rows.map(p=>({p,update:{...p.review},meta:{...model.tools.find(t=>t.id===p.review.toolId)}}));state.preparing=true;state.preparationError='';refresh();
 try{const bytes=await ToolWorkbook.apply(model,rows.map(p=>p.review));if(revision!==state.revision||model!==state.model)return null;const blob=new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),stamp=new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Seoul'}).format(new Date()),name=state.workbookName.replace(/\.xlsx$/i,'')+'_수정본_'+stamp+'.xlsx';if(state.preparedResult)URL.revokeObjectURL(state.preparedResult.url);state.preparedResult={bytes,blob,url:URL.createObjectURL(blob),name,count:rows.length};
 if(state.source.connected){
  const updatedModel=await ToolWorkbook.load(bytes);
  if(revision!==state.revision||sourceEpoch!==state.source.epoch)return null;
  state.source.saving=true;refresh();
  try{
   const saved=await ToolConnectedFile.save({bytes,expectedVersion:sourceVersion,verifyCurrent:()=>revision===state.revision&&sourceEpoch===state.source.epoch});
   state.source.version=saved.version;state.source.saveCount++;state.source.error='';state.source.hasBackup=true;state.source.savedAt=new Intl.DateTimeFormat('ko-KR',{timeZone:'Asia/Seoul',hour:'2-digit',minute:'2-digit',second:'2-digit'}).format(new Date());state.model=updatedModel;
   for(const entry of entries){const after=updatedModel.tools.find(t=>t.id===entry.update.toolId)||entry.meta;entry.p.saved={...after,photoKind:entry.update.photoKind||'calibration',version:saved.version};if(entry.update.photoKind==='smax')Object.assign(entry.p.draft,{last:after.last,due:after.due});}
   rememberApproved(entries);
   status('연결한 이 PC의 엑셀에 저장했습니다. OneDrive 동기화는 별도로 진행됩니다.','success');
  }catch(e){sourceFailure(e);status('연결 파일 저장 확인 필요: '+e.message+' 수정본 다운로드는 사용할 수 있습니다.','error');}
  finally{state.source.saving=false;refresh();render();}
 }else rememberApproved(entries);
 return state.preparedResult;
 }catch(e){if(revision===state.revision){state.preparationError=e.message;status('엑셀 생성 실패: '+e.message,'error');}return null;}finally{state.preparing=false;refresh();if(revision!==state.revision)schedule();}
}
async function saveReview(e){
 e.preventDefault();const p=current();if(!p||!state.model||!$('confirm').checked||state.busy)return;saveDraft();const update={...p.draft,...p.image};
 try{const errors=ToolWorkbook.validate(state.model,[...state.photos.filter(q=>q.id!==p.id&&q.review&&!q.saved).map(q=>q.review),update]);if(errors.length)throw new Error(errors.join('\n'));invalidate();p.review=update;p.approval='manual';p.edited=true;refresh();render();await prepareResult();if(!state.preparationError&&!state.source.error)status(state.source.connected?'보정한 내용을 연결한 이 PC의 엑셀에 자동 저장했습니다.':'보정한 내용이 반영되었습니다. 수정본을 저장하세요.','success');
 }catch(e){p.review=null;p.approval=null;invalidate();refresh();schedule();status(e.message,'error');}
}
function download(blob,name){const a=document.createElement('a'),url=URL.createObjectURL(blob);a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);}
async function downloadDiagnostics(){
 const p=current();if(!p||state.busy||state.source.saving)return;
 const button=$('download-diagnostics');button.disabled=true;
 try{
  const original=await p.file.arrayBuffer(),extension=/\.(jpe?g|png|webp|heic|heif|avif)$/i.exec(p.file.name)?.[1].toLowerCase()||'bin';
  const r=p.result||{},report={schema:1,programVersion:'4.5',createdAt:new Date().toISOString(),
   photo:{name:p.file.name,type:p.file.type,size:p.file.size,lastModified:p.file.lastModified,entry:'original.'+extension},
   ocr:{text:r.text||'',documentType:r.documentType||'calibration',smaxEvidence:r.smaxEvidence||null,confidence:r.confidence??null,failed:!!r.failed,cancelled:!!r.cancelled,
    engineFailure:!!r.engineFailure,identityRecoveryIncomplete:!!r.identityRecoveryIncomplete,
    error:r.error||'',code:r.code||'',identityEvidence:r.identityEvidence||null,dateEvidence:r.dateEvidence||null,
    passes:r.passes||[]},assessment:p.assessment||null,processing:{approval:p.approval||null,saved:!!p.saved}};
  const zip=new JSZip();zip.file('original.'+extension,original);zip.file('recognition.json',JSON.stringify(report,null,2));
  const blob=await zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}});
  download(blob,'Tool-Recognition-Diagnostics.zip');status('선택한 원본 사진과 인식 결과를 담은 진단 ZIP 저장을 요청했습니다. 이 ZIP을 첨부하면 같은 사진과 신뢰도로 문제를 확인할 수 있습니다.','success');
 }catch(e){status('인식 진단 ZIP을 저장하지 못했습니다: '+e.message,'error');}
 finally{button.disabled=state.busy||state.source.saving;}
}
$('workbook').onchange=e=>readWorkbook(e.target.files[0]);$('photos').onchange=e=>addPhotos([...e.target.files]);
$('remember-matches').onchange=()=>{try{localStorage.setItem('tool-manager-match-memory-enabled',String(memoryEnabled()));}catch(_){}refreshMemory();status(memoryEnabled()?'확인된 매칭을 기억하고 다음 사진에 활용합니다.':'다음에 추가하거나 다시 인식하는 사진부터 매칭 기억 활용을 끕니다.');};
$('clear-matches').onclick=()=>{if(state.busy||state.preparing||state.source.saving)return;const info=window.ToolProfiles?.clear();refreshMemory();status(info?.error||'기억한 인식 패턴을 지웠습니다. 다음 사진부터 새로 기억합니다.',info?.error?'error':'');};
$('automatic-mode').onchange=()=>{if(automatic()&&state.model&&state.photos.length)recognize(state.photos.filter(p=>!p.review).map(p=>p.id));else refresh();};
$('recognize-all').onclick=()=>recognize(state.photos.filter(p=>!p.review).map(p=>p.id));$('recognize-one').onclick=()=>current()&&recognize([current().id]);
$('download-diagnostics').onclick=downloadDiagnostics;
$('photo-kind').onchange=()=>{const p=current();if(!p)return;p.photoKindOverride=$('photo-kind').value;if(p.photoKindOverride==='smax'){$('smax-update-due').checked=false;const tool=state.model?.tools.find(t=>t.id===$('tool-select').value);$('last').value=tool?.last||'';}changed();render();};
$('smax-update-due').onchange=()=>{changed();render();};
$('retry-engine').onclick=()=>current()&&recognize([current().id]);
$('manual-entry-toggle').onclick=()=>{const p=current();if(!p||state.busy||state.source.saving)return;p.manualEntry=!p.manualEntry;render();};
$('cancel-recognition').onclick=()=>{state.cancelRequested=true;ToolOCR.cancel?.();status('인식을 중지하는 중입니다…');};
$('remove-photo').onclick=()=>{const p=current();if(!p||state.busy)return;invalidate();URL.revokeObjectURL(p.url);state.photos=state.photos.filter(q=>q.id!==p.id);state.selectedId=state.photos[0]?.id||null;reconcile();refresh();render();schedule();status('사진을 제거했습니다.');};
$('tool-select').onchange=()=>{const t=state.model?.tools.find(t=>t.id===$('tool-select').value);$('name').value=t?.name||'';$('serial').value=t?.serial||'';renderExisting();changed();};
for(const k of ['name','serial','last','due'])$(k).oninput=changed;
$('confirm').onchange=()=>{if(!$('confirm').checked&&current()){current().review=null;current().approval=null;current().edited=true;invalidate();refresh();schedule();}};
$('calculate-due').onclick=()=>{try{const t=state.model?.tools.find(t=>t.id===$('tool-select').value);if(!t||!$('last').value)throw new Error('툴 항목과 교정일을 입력해 주세요.');const months=ToolWorkbook.monthInterval(t.interval,t.cycle),due=months?ToolWorkbook.addMonths($('last').value,months):'';if(!due)throw new Error('엑셀 주기로 날짜를 계산할 수 없습니다.');$('due').value=due;changed();}catch(e){status(e.message,'error');}};
$('review-form').onsubmit=saveReview;$('export').onclick=async()=>{if(state.busy||state.preparing||state.source.saving)return;if(!state.preparedResult)await prepareResult();if(!state.preparedResult)return;download(state.preparedResult.blob,state.preparedResult.name);status('수정본 저장을 요청했습니다. 확인 필요 사진은 제외했습니다.','success');};
$('download-log').onclick=()=>{const changes=approved().map(p=>{const u=p.review,t=p.saved||state.model.tools.find(t=>t.id===u.toolId)||{};return {photo:p.file.name,processing:p.approval,photoKind:u.photoKind||'calibration',updateDue:u.photoKind!=='smax'||u.updateDue===true,sheet:t.sheet,row:t.row,name:u.name,serial:u.serial,last:u.last,due:u.due,previousLast:t.last,previousDue:t.due};}),pending=state.photos.filter(p=>!p.review).map(p=>({photo:p.file.name,reasons:p.assessment?.reasons||[p.error||'대기 중']}));download(new Blob([JSON.stringify({workbook:state.workbookName,createdAt:new Date().toISOString(),changes,pending},null,2)],{type:'application/json'}),'툴_변경내역.json');};
window.addEventListener('beforeunload',e=>{if(state.photos.length){e.preventDefault();e.returnValue='';}});
$('connect-onedrive').onclick=()=>connectSource(false);$('reconnect-onedrive').onclick=()=>connectSource(true);$('retry-connected-save').onclick=retrySource;
$('disconnect-onedrive').onclick=async()=>{if(state.busy||state.preparing||state.source.saving)return;await ToolConnectedFile.disconnect();state.source.epoch++;Object.assign(state.source,{connected:false,needsPermission:false,needsReconnect:false,error:'',name:'',version:'',savedAt:''});refresh();render();status('자동 저장 연결을 해제했습니다. 이후 결과는 수정본으로 저장합니다.');};
$('download-backup').onclick=async()=>{try{const backup=await ToolConnectedFile.takeBackup();if(!backup)throw new Error('보관된 백업이 없습니다.');download(new Blob([backup.bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),backup.name.replace(/\.xlsx$/i,'')+'_저장전백업.xlsx');}catch(e){status(e.message,'error');}};
const linkKey='tool-manager-sharepoint-link';
function setCloudLink(raw,persist=false){const value=raw.trim();if(!value){$('open-sharepoint').hidden=true;if(persist)try{localStorage.removeItem(linkKey);}catch(_){};return;}const url=new URL(value);if(url.protocol!=='https:')throw new Error('HTTPS 회사 링크를 입력하세요.');$('sharepoint-url').value=url.href;$('open-sharepoint').href=url.href;$('open-sharepoint').hidden=false;if(persist)try{localStorage.setItem(linkKey,url.href);}catch(_){};}
$('remember-sharepoint').onclick=()=>{try{setCloudLink($('sharepoint-url').value,true);status('회사 문서 링크를 이 브라우저에 저장했습니다. 엑셀은 별도로 OneDrive 동기화 폴더에서 연결하세요.','success');}catch(e){status(e.message,'error');}};
try{setCloudLink(localStorage.getItem(linkKey)||'');}catch(_){}
try{$('remember-matches').checked=localStorage.getItem('tool-manager-match-memory-enabled')!=='false';}catch(_){}
window.ToolApp={state,readWorkbook,addPhotos,selectPhoto,prepareResult,connectSource,retrySource};refresh();
restoreTask=restoreSource();
const missing=['ToolWorkbook','ToolSMAX','ToolOCR','ToolAutomation','ToolRuntime','ToolConnectedFile'].filter(k=>!window[k]);
if(missing.length)status('실행 파일이 완전하지 않습니다: '+missing.join(', ')+'. 전체 HTML을 다시 저장해 Chrome/Edge에서 여세요.','error');
});
})();
