/* Capture immutable program source before app boot; never export uploaded data. */
(() => {
'use strict';
const $=id=>document.getElementById(id);
let pristine='';
const meta=()=>JSON.parse($('distribution-meta').textContent);
const html=()=>{if(!pristine)throw new Error('프로그램이 준비 중입니다. 잠시 후 다시 저장하세요.');return pristine;};
const feedback=t=>{$('download-feedback').textContent=t;};
async function zip(){const data=meta(),z=new JSZip();z.file('Tool-Manager.html',html());z.file('사용법.txt',data.manual);for(const [name,text] of Object.entries(data.licenses))z.file('licenses/'+name,text);return z.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6},mimeType:'application/zip'});}
function reveal(){try{$('save-source-panel').hidden=false;$('program-source-text').value=html();feedback('아래 코드는 빈 프로그램입니다. 전체 복사 후 메모장에 Tool-Manager.html로 저장할 수 있습니다.');}catch(e){feedback(e.message);}}
function download(blob,name){const a=document.createElement('a'),url=URL.createObjectURL(blob);a.href=url;a.download=name;a.style.display='none';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);feedback(name+' 저장을 브라우저에 요청했습니다. 차단되면 「다운로드가 막히면」을 이용하세요.');}
async function save(kind){const button=$(kind==='zip'?'save-program-zip':'save-program-html'),filename=kind==='zip'?'Tool-Manager.zip':'Tool-Manager.html';button.disabled=true;
 try{let picker=null;if(typeof window.showSaveFilePicker==='function'){try{picker=window.showSaveFilePicker({suggestedName:filename,types:[{description:'툴 검교정 프로그램',accept:kind==='zip'?{'application/zip':['.zip']}:{'text/html':['.html']}}]});}catch(_){}}
 const blobPromise=kind==='zip'?zip():Promise.resolve(new Blob([html()],{type:'text/html;charset=utf-8'}));let handle=null;
 if(picker){try{handle=await picker;}catch(e){if(e.name==='AbortError'){await blobPromise;feedback('저장을 취소했습니다.');return;}}}
 const blob=await blobPromise;if(handle){const writable=await handle.createWritable();await writable.write(blob);await writable.close();feedback(filename+' 파일을 저장했습니다. Chrome/Edge에서 여세요.');}else download(blob,filename);
 }catch(e){feedback('저장을 완료하지 못했습니다: '+e.message);reveal();}finally{button.disabled=false;}}
window.ToolDistribution={html,zip};
window.addEventListener('DOMContentLoaded',()=>{
 // This listener is registered before app.js. Capture once, before any uploads.
 pristine='<!doctype html>\n'+document.documentElement.outerHTML;
 $('save-program-html').onclick=()=>save('html');$('save-program-zip').onclick=()=>save('zip');$('show-save-source').onclick=reveal;
 $('copy-program-source').onclick=async()=>{reveal();const area=$('program-source-text');try{if(!navigator.clipboard?.writeText)throw new Error();await navigator.clipboard.writeText(area.value);feedback('전체 코드를 복사했습니다. 메모장에 붙여넣으세요.');}catch(_){area.focus();area.select();feedback('코드를 모두 선택했습니다. Ctrl+C 또는 맥에서 ⌘C를 누르세요.');}};
 $('program-source-text').onclick=e=>e.target.select();
});
})();
