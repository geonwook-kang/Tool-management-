from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import argparse,hashlib,subprocess,sys
root=Path(__file__).resolve().parent
p=argparse.ArgumentParser();p.add_argument('--output-dir',type=Path,default=Path('/workspace/outputs'));a=p.parse_args();a.output_dir.mkdir(parents=True,exist_ok=True)
html=a.output_dir/'Tool-Manager-Fixed.html';subprocess.run([sys.executable,str(root/'build.py'),'--output',str(html)],check=True)
for name in ('Tool-Manager.html','Tool-Manager-Auto.html','Tool-Manager-Download.html'):(a.output_dir/name).write_bytes(html.read_bytes())
with ZipFile(a.output_dir/'Tool-Manager.zip','w',compression=ZIP_DEFLATED,compresslevel=6) as z:
 z.write(html,'Tool-Manager.html');z.write(root/'사용법.txt','사용법.txt')
 for file in sorted((root/'vendor').glob('*.txt')):z.write(file,'licenses/'+file.name)
(a.output_dir/'사용법.txt').write_bytes((root/'사용법.txt').read_bytes())
onedrive=a.output_dir/'Tool-Manager-OneDrive.zip';onedrive.write_bytes((a.output_dir/'Tool-Manager.zip').read_bytes())
fixed=a.output_dir/'Tool-Manager-Auto-Fix-4.1.zip';fixed.write_bytes((a.output_dir/'Tool-Manager.zip').read_bytes())
files=[html,a.output_dir/'Tool-Manager.zip',onedrive,fixed];(a.output_dir/'SHA256SUMS.txt').write_text(''.join(hashlib.sha256(f.read_bytes()).hexdigest()+'  '+f.name+'\n' for f in files))
print('Packaged',*[f'{f.name}: {f.stat().st_size:,}B' for f in files])
