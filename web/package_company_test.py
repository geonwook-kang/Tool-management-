"""Build a separate copy-only, large-workbook test entry from frozen PC code.

Production PC/mobile releases are not rebuilt or overwritten. Explicit source
substitutions are guarded so later source changes cannot silently bypass a hook.
Company files and private photographs are never packaged into this program.
"""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from hashlib import sha256
import argparse
import json
import re
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent
VERSION = '7.2-T1'


def replace_once(html, old, new):
    if html.count(old) != 1:
        raise ValueError('Expected exactly one company-test build hook: '+old[:90])
    return html.replace(old, new, 1)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir', type=Path, default=Path('/workspace/outputs'))
    args = parser.parse_args()
    output = args.output_dir.resolve(); output.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='tool-company-build-') as directory:
        base = Path(directory)/'pc-base.html'
        subprocess.run([sys.executable, str(ROOT/'build.py'), '--entry', 'pc', '--output', str(base)], check=True)
        html = base.read_text()
    html = replace_once(html, '<head>', '<head><meta name="tool-manager-company-test-entry" content="true"><script>document.documentElement.classList.add("tool-company-test");</script>')
    html = replace_once(html, '<title>툴 검교정 매니저 · PC 마무리 버전 7.2</title>',
                        '<title>툴 검교정 매니저 · 전사 파일 테스트판 '+VERSION+'</title>')
    html = replace_once(html, '<span class="badge">처음 한 번 선택</span>', '<span class="badge">이번 창에서 테스트</span>')
    html = replace_once(html, '저장한 HTML을 Chrome 또는 Edge에서 여세요. 사진 인식은 오프라인으로 실행합니다. OneDrive 동기화는 회사 계정을 사용합니다.',
                        '저장한 HTML을 Chrome 또는 Edge에서 여세요. 이 테스트판은 원본 연결 없이 별도 사본만 변환합니다.')
    html = replace_once(html, "if(file.size>100*1024*1024){status('엑셀은 100MB 이하로 선택해 주세요.','error');return;}",
                        "if(file.size>250*1024*1024){status('전사 테스트 파일은 250 MiB 이하로 선택해 주세요.','error');return;}")
    html = replace_once(html, "state.model=model;state.workbookName=file.name;", "state.model=model;state.workbookName=file.name.replace(/\\.xlsx$/i,'')+'_전사테스트.xlsx';")
    html = replace_once(html, "if (bytes.byteLength > 100 * 1024 * 1024) throw new Error('엑셀 파일은 100 MB 이하로 올려 주세요.');",
                        "if (bytes.byteLength > 250 * 1024 * 1024) throw new Error('전사 테스트 엑셀은 250 MiB 이하로 선택해 주세요.');")
    html = replace_once(html, 'entries.length > 5000 || entries.reduce((n, f) => n + (f._data && f._data.uncompressedSize || 0), 0) > 250 * 1024 * 1024',
                        'entries.length > 20000 || entries.reduce((n, f) => n + (f._data && f._data.uncompressedSize || 0), 0) > 1024 * 1024 * 1024')
    # Count picture locations once per sheet; large regional files otherwise
    # scan the same anchors twice for every registered tool.
    html = replace_once(html, 'sheets.push(record);\n      for (const row of rows) {', '''sheets.push(record);
      const pictureCounts = new Map();
      for (const drawing of drawingRecords) for (const anchor of drawing.doc.documentElement.children) {
        const loc = anchorLocation(anchor);
        if (loc && elements(anchor, 'pic').length) { const key = loc.row + ':' + loc.col; pictureCounts.set(key, (pictureCounts.get(key) || 0) + 1); }
      }
      for (const row of rows) {''')
    html = replace_once(html, '''const picturesAt = (col) => !col ? [] : drawingRecords.flatMap((d) => Array.from(d.doc.documentElement.children).filter((a) => {
          const loc = anchorLocation(a);
          return loc && loc.row === rowNum && loc.col === col && elements(a, 'pic').length;
        }));''', "const pictureCount = (col) => col ? (pictureCounts.get(rowNum + ':' + col) || 0) : 0;")
    html = replace_once(html, "photoCount: picturesAt(headers['Label Photo']).length,", "photoCount: pictureCount(headers['Label Photo']),")
    html = replace_once(html, "smaxPhotoCount: picturesAt(headers['SMAX Tool Photo']).length,", "smaxPhotoCount: pictureCount(headers['SMAX Tool Photo']),")
    html = replace_once(html, 'state.set(model, { bytes, zip, sheets, date1904 });', '''model.companyTest = Object.freeze({
      inputBytes: bytes.byteLength,
      allSheetNames: elements(workbook, 'sheet').map(sheet => sheet.getAttribute('name')),
      mediaCount: entries.filter(entry => !entry.dir && /^xl\\/media\\//.test(entry.name)).length,
      parts: Object.fromEntries(entries.filter(entry => !entry.dir).map(entry => [entry.name, { crc32: (entry._data?.crc32 || 0) >>> 0, bytes: entry._data?.uncompressedSize || 0 }])),
      sheetParts: sheets.map(sheet => ({ name: sheet.name, path: sheet.path, drawings: sheet.drawings.map(drawing => drawing.path) }))
    });
    state.set(model, { bytes, zip, sheets, date1904 });''')
    html = replace_once(html, "const KEY = 'tool-management-profiles-v1'", "const KEY = 'tool-management-company-test-profiles-v1'")
    html = replace_once(html, "programVersion:ToolSources.inspect().pcEntry?'7.2':'6.0'", "programVersion:ToolSources.inspect().pcEntry?'7.2-T1':'6.0'")
    module = (ROOT/'company-test.js').read_text()
    module = re.sub(r'</script', lambda _:r'<\/script', module, flags=re.I)
    app = (ROOT/'app.js').read_text()
    marker = '<script>'+app[:60]
    # Only use the beginning of APP to locate its script after all adapters.
    html = replace_once(html, marker, '<script>'+module+'</script>\n'+marker)
    # Distribution captures this immutable source, including the test entry.
    manual = (ROOT/'전사테스트-사용법.txt').read_text()
    match = re.search(r'<script id="distribution-meta" type="application/json">(.*?)</script>', html, re.S)
    if not match:
        raise ValueError('Missing distribution metadata')
    meta = json.loads(match.group(1)); meta['manual'] = manual
    meta_json = json.dumps(meta, ensure_ascii=False, separators=(',', ':')).replace('<', '\\u003c')
    html = html[:match.start(1)] + meta_json + html[match.end(1):]
    html = re.sub(r'<footer>.*?</footer>', '<footer>TOOL MANAGER · 전사 테스트 '+VERSION+' · 현재 창에서 사본만 처리합니다. 결과 엑셀을 다운로드하세요.</footer>', html, count=1, flags=re.S)
    target = output/'Tool-Manager-Company-Test.html'
    target.write_text(html)
    archive = output/('Tool-Manager-Company-Test-'+VERSION+'.zip')
    with ZipFile(archive, 'w', compression=ZIP_DEFLATED, compresslevel=6) as bundle:
        bundle.write(target, target.name)
        bundle.write(ROOT/'전사테스트-사용법.txt', '전사테스트-사용법.txt')
        for license_file in sorted((ROOT/'vendor').glob('*.txt')):
            bundle.write(license_file, 'licenses/'+license_file.name)
    with ZipFile(archive) as bundle:
        assert bundle.testzip() is None
        assert bundle.read(target.name) == target.read_bytes()
    sums=''.join(sha256(path.read_bytes()).hexdigest()+'  '+path.name+'\n' for path in [target, archive])
    (output/'Company-Test-SHA256SUMS.txt').write_text(sums)
    print('Company test HTML:', target, target.stat().st_size)
    print('Company test ZIP:', archive, archive.stat().st_size)


if __name__ == '__main__':
    main()
