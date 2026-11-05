"""Build a self-contained OCR application with verified, pinned offline assets."""
from pathlib import Path
import argparse,base64,json,re,subprocess,sys
root=Path(__file__).resolve().parent
parser=argparse.ArgumentParser();parser.add_argument('--output',type=Path,default=root/'dist/Tool-Manager.html');parser.add_argument('--entry',choices=['default','pc'],default='default');args=parser.parse_args()
subprocess.run([sys.executable,str(root.parent/'scripts/prepare_assets.py'),'--check'],check=True)
v=root/'vendor'
runtime={'workerSource':(v/'worker.min.js').read_text(),'coreSource':(v/'tesseract-core-lstm.wasm.js').read_text(),'languages':{lang:base64.b64encode((v/(lang+'.traineddata.gz')).read_bytes()).decode() for lang in ('kor','eng')}}
meta={'manual':(root/('PC-사용법.txt' if args.entry=='pc' else '사용법.txt')).read_text(),'licenses':{p.name:p.read_text() for p in sorted(v.glob('*.txt'))}}
def as_json(data):return json.dumps(data,ensure_ascii=False,separators=(',',':')).replace('<','\\u003c')
html=(root/'index.template.html').read_text().replace('/*__RUNTIME_ASSETS__*/',as_json(runtime)).replace('/*__DISTRIBUTION_META__*/',as_json(meta))
for token,path in [('JSZIP','vendor/jszip.min.js'),('TESSERACT','vendor/tesseract.min.js'),('RUNTIME','runtime.js'),('WORKBOOK','workbook.js'),('SHARED_MERGE','shared-merge.js'),('SMAX','smax.js'),('FAST_OCR','fast-ocr.js'),('OCR','ocr.js'),('AUTOMATION','automation.js'),('PROFILES','profiles.js'),('DISTRIBUTE','distribute.js'),('CONNECTED_FILE','connected-file.js'),('TEAM_CONFIG','team-config.js'),('M365_AUTH','m365-auth.js'),('SHARED_WORKSPACE','shared-workspace.js'),('GRAPH_FILE','graph-file.js'),('SHARED_FILE','shared-file.js'),('SHAREPOINT_FILE','sharepoint-file.js'),('MOBILE_FILE','mobile-file.js'),('RESUME','resume.js'),('SOURCES','sources.js'),('SHARED_UI','shared-ui.js'),('APP','app.js')]:
 text=(root/path).read_text();text=re.sub(r'</script',lambda _:r'<\/script',text,flags=re.I);html=html.replace('/*__'+token+'__*/',text)
assert not re.search(r'/\*__[A-Z_]+__\*/',html),'Unresolved build placeholder'
if args.entry=='pc':
 html=html.replace('<head>','<head><meta name="tool-manager-direct-entry" content="true"><meta name="tool-manager-pc-entry" content="true"><script>document.documentElement.classList.add("tool-pc");</script>',1)
 html=re.sub(r'<title>[^<]*</title>','<title>툴 검교정 매니저 · PC 마무리 버전 7.0</title>',html,count=1)
 html=re.sub(r'<footer>.*?</footer>','<footer>TOOL MANAGER · PC 7.0 · 사진은 이 PC에서 인식합니다. 엑셀 저장 대상은 연결 상태에서 확인하세요.</footer>',html,count=1,flags=re.S)
args.output.parent.mkdir(parents=True,exist_ok=True);args.output.write_text(html,encoding='utf-8')
print(f'Built {args.output}: {args.output.stat().st_size:,} bytes; offline OCR included')
