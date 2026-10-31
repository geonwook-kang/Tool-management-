/* Production host helpers and lifecycle; SPFx/DOM are controlled test fixtures. */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');
const root = path.resolve(__dirname, '..');

function sourceModule(file, imports = {}, globals = {}) {
  const source = fs.readFileSync(path.join(root, file), 'utf8');
  const emitted = ts.transpileModule(source, { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 } });
  const module = { exports: {} };
  const context = { module, exports: module.exports, require: name => {
    if (!(name in imports)) throw new Error('Unexpected source import ' + name);
    return imports[name];
  }, ...globals };
  vm.runInNewContext(emitted.outputText, context, { filename: file });
  return module.exports;
}
const host = sourceModule('src/hostFrame.ts');
const context = { siteUrl: 'https://example.sharepoint.com/sites/team', fileUrl: '/sites/team/Shared Documents/</script><script>bad()</script>.xlsx', applicationSha256: 'a'.repeat(64) };
const documentText = host.hostedDocument('<!doctype html><html><head></head><body></body></html>', context);
assert(documentText.includes('data-sharepoint-shell'));
assert(!documentText.includes('</script><script>bad()'));
assert(documentText.includes('Object.freeze(') && documentText.includes('writable:false,configurable:false'));
const injected = documentText.match(/<script data-sharepoint-shell>(.*?)<\/script>/s)[1];
const frameGlobal = {};
vm.runInNewContext(injected, { window: frameGlobal });
assert(Object.isFrozen(frameGlobal.ToolSharePointContext));
assert.equal(frameGlobal.ToolSharePointContext.fileUrl, context.fileUrl);
assert.equal(Object.getOwnPropertyDescriptor(frameGlobal, 'ToolSharePointContext').writable, false);
assert.equal(Object.getOwnPropertyDescriptor(frameGlobal, 'ToolSharePointContext').configurable, false);
for (const [input, expected] of [[200,650],[1000.4,1000],[40000,30000],[NaN,undefined],[Infinity,undefined],['1200',undefined],[undefined,undefined]]) assert.equal(host.safeFrameHeight(input), expected);
for (const state of [{busy:true},{preparing:true},{restoring:true},{source:{saving:true}},{photos:[{saved:null}]}]) assert(host.frameHasPendingWork({contentWindow:{ToolApp:{state}}}));
assert(!host.frameHasPendingWork({contentWindow:{ToolApp:{state:{photos:[{saved:{serial:'fixture'}}]}}}}));
assert(!host.frameHasPendingWork({contentWindow:null}));
for (const state of [{pendingCheckout:true},{uncertain:true},{busy:true}]) assert(host.frameHasPendingWork({contentWindow:{ToolSharePointFile:{inspect:()=>state}}}));
const inaccessible = {};Object.defineProperty(inaccessible,'contentWindow',{get(){throw new Error('cross-origin')}});assert(host.frameHasPendingWork(inaccessible));

class Element {
  constructor(tag) { this.tagName=tag;this.children=[];this.style={};this.attributes={};this.hidden=false;this.contentWindow={};this.textContent=''; }
  appendChild(child) { this.children.push(child); return child; }
  replaceChildren(...children) { this.children=[...children]; }
  setAttribute(name,value) { this.attributes[name]=value; }
}
const listeners = new Set(), timers = new Map();let sequence=0;
const browser = { location:{origin:'https://example.sharepoint.com'}, addEventListener:(_,fn)=>listeners.add(fn), removeEventListener:(_,fn)=>listeners.delete(fn),
  setTimeout:(fn,delay)=>{assert.equal(delay,45000);const id=++sequence;timers.set(id,fn);return id;},clearTimeout:id=>timers.delete(id) };
class Base {
  constructor(){this.properties={fileUrl:'/sites/team/shared.xlsx'};this.context={pageContext:{web:{absoluteUrl:'https://example.sharepoint.com/sites/team'}}};this.domElement=new Element('div');}
  onDispose(){}
}
const production = sourceModule('src/webparts/toolManager/ToolManagerWebPart.ts', {
  '@microsoft/sp-core-library':{Version:{parse:value=>value}},
  '@microsoft/sp-property-pane':{PropertyPaneTextField:(name,options)=>({name,options})},
  '@microsoft/sp-webpart-base':{BaseClientSideWebPart:Base},
  '../../generated/appPayload':{APP_HTML:'<!doctype html><html><head></head><body>Fixture app</body></html>',APP_SHA256:'a'.repeat(64)},
  '../../hostFrame':host
}, {window:browser,document:{createElement:tag=>new Element(tag)},URL,Number});
const app = new production.default();app.render();
assert.equal(app.domElement.children.length,2);const firstFrame=app.domElement.children[1];
assert.equal(firstFrame.attributes.sandbox,'allow-scripts allow-same-origin allow-forms allow-downloads allow-modals');
assert.equal(firstFrame.attributes.allow,'camera; clipboard-write; web-share');
assert.equal(listeners.size,1);assert.equal(timers.size,1);
app.render();assert.equal(app.domElement.children[1],firstFrame);assert.equal(timers.size,1);
firstFrame.contentWindow.ToolApp={state:{photos:[{saved:null}],source:{}}};
app.properties.fileUrl='/sites/team/different.xlsx';app.render();
assert.equal(app.domElement.children[1],firstFrame);assert.equal(app.properties.fileUrl,'/sites/team/shared.xlsx');assert(app.domElement.children[0].textContent.includes('사진 작업'));
for(const state of [{busy:true},{preparing:true},{restoring:true},{source:{saving:true}}]){
 firstFrame.contentWindow.ToolApp.state=state;app.properties.fileUrl='/sites/team/different.xlsx';app.render();assert.equal(app.domElement.children[1],firstFrame);assert.equal(app.properties.fileUrl,'/sites/team/shared.xlsx');
}
const receive=[...listeners][0];
receive({source:{},origin:browser.location.origin,data:{type:'TOOL_SHAREPOINT_READY'}});assert.equal(timers.size,1);
receive({source:firstFrame.contentWindow,origin:'https://other.example',data:{type:'TOOL_SHAREPOINT_READY'}});assert.equal(timers.size,1);
receive({source:firstFrame.contentWindow,origin:browser.location.origin,data:{type:'TOOL_SHAREPOINT_READY'}});assert.equal(timers.size,0);assert(app.domElement.children[0].hidden);
receive({source:firstFrame.contentWindow,origin:browser.location.origin,data:{type:'TOOL_SHAREPOINT_HEIGHT',height:1200}});assert.equal(firstFrame.style.height,'1200px');
receive({source:firstFrame.contentWindow,origin:'https://other.example',data:{type:'TOOL_SHAREPOINT_HEIGHT',height:9000}});assert.equal(firstFrame.style.height,'1200px');
firstFrame.contentWindow.ToolApp.state={photos:[{saved:{}}]};app.properties.fileUrl='/sites/team/different.xlsx';app.render();assert.notEqual(app.domElement.children[1],firstFrame);assert.equal(timers.size,1);
[...timers.values()][0]();assert(app.domElement.children[0].textContent.includes('시작 응답'));app.onDispose();assert.equal(timers.size,0);assert.equal(listeners.size,0);assert.equal(app.domElement.children.length,0);
const outside = new production.default();outside.context.pageContext.web.absoluteUrl='https://other.sharepoint.com/sites/team';outside.render();assert.equal(outside.domElement.children.length,1);assert.equal(timers.size,0);
console.log(JSON.stringify({checks:['context_json_escaped_frozen_and_nonreplaceable','height_limits','all_pending_work_states_preserve_frame','same_properties_keep_existing_frame','busy_target_change_rolls_back_persisted_property','source_origin_guard_for_messages','ready_clears_timeout','timeout_does_not_claim_connected','completed_target_change_replaces_frame','dispose_removes_listeners_timers','cross_origin_host_refused'],physical_sharepoint_page_tested:false},null,2));
