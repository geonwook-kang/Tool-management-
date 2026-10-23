"""Build a self-contained OCR application with verified, pinned offline assets."""
from pathlib import Path
import argparse,base64,json,re,subprocess,sys
root=Path(__file__).resolve().parent
parser=argparse.ArgumentParser();parser.add_argument('--output',type=Path,default=root/'dist/Tool-Manager.html');args=parser.parse_args()
subprocess.run([sys.executable,str(root.parent/'scripts/prepare_assets.py'),'--check'],check=True)
v=root/'vendor'
runtime={'workerSource':(v/'worker.min.js').read_text(),'coreSource':(v/'tesseract-core-lstm.wasm.js').read_text(),'languages':{lang:base64.b64encode((v/(lang+'.traineddata.gz')).read_bytes()).decode() for lang in ('kor','eng')}}
meta={'manual':(root/'사용법.txt').read_text(),'licenses':{p.name:p.read_text() for p in sorted(v.glob('*.txt'))}}
def as_json(data):return json.dumps(data,ensure_ascii=False,separators=(',',':')).replace('<','\\u003c')
html=(root/'index.template.html').read_text().replace('/*__RUNTIME_ASSETS__*/',as_json(runtime)).replace('/*__DISTRIBUTION_META__*/',as_json(meta))
for token,path in [('JSZIP','vendor/jszip.min.js'),('TESSERACT','vendor/tesseract.min.js'),('RUNTIME','runtime.js'),('WORKBOOK','workbook.js'),('OCR','ocr.js'),('AUTOMATION','automation.js'),('PROFILES','profiles.js'),('DISTRIBUTE','distribute.js'),('CONNECTED_FILE','connected-file.js'),('APP','app.js')]:
 text=(root/path).read_text();text=re.sub(r'</script',lambda _:r'<\/script',text,flags=re.I);html=html.replace('/*__'+token+'__*/',text)
assert not re.search(r'/\*__[A-Z_]+__\*/',html),'Unresolved build placeholder'
args.output.parent.mkdir(parents=True,exist_ok=True);args.output.write_text(html,encoding='utf-8')
print(f'Built {args.output}: {args.output.stat().st_size:,} bytes; offline OCR included')
