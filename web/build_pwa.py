"""Build the public HTTPS/PWA shell without company workbooks or photographs.

python web/build_pwa.py --output-dir /workspace/outputs/mobile-app
The existing standalone builder remains the source of the offline OCR program.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
import json
import re
from pathlib import Path
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent


def published_settings():
    """Keep a separately published target; never package sharing URLs/secrets."""
    path = ROOT.parent / 'docs/team-settings.json'
    value = json.loads(path.read_text(encoding='utf-8')) if path.exists() else {'enabled': False}
    if value == {'enabled': False}:
        return b'{"enabled":false}\n'
    if not isinstance(value, dict) or set(value) - {'enabled', 'v', 'clientId', 'tenantId', 'name', 'target'} or value.get('enabled') is not True or value.get('v') != 1:
        raise ValueError('Published team settings must contain only public app IDs and an exact workbook target')
    guid = r'[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}'
    tenant = value.get('tenantId', '')
    if not isinstance(value.get('clientId'), str) or not re.fullmatch(guid, value['clientId']) or not isinstance(tenant, str) or not (re.fullmatch(guid, tenant) or tenant == 'organizations' or re.fullmatch(r'[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.onmicrosoft\.com', tenant)):
        raise ValueError('Published client and tenant IDs are invalid')
    target = value.get('target')
    if not isinstance(target, dict) or set(target) != {'driveId', 'itemId', 'name'} or any(not isinstance(target.get(k), str) or not re.fullmatch(r'[A-Za-z0-9!._~-]{1,200}', target[k]) for k in ('driveId', 'itemId')):
        raise ValueError('Published workbook target is invalid')
    for name, maximum in ((value.get('name', ''), 120), (target.get('name', ''), 255)):
        if not isinstance(name, str) or len(name) > maximum or re.search(r'[\x00-\x1f\x7f]', name):
            raise ValueError('Published workbook name is invalid')
    if not target['name'].lower().endswith('.xlsx') or re.search(r'[/\\]', target['name']):
        raise ValueError('Published workbook must be an .xlsx file')
    return (json.dumps(value, ensure_ascii=False, indent=2) + '\n').encode()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir', type=Path, default=ROOT / 'pwa-dist')
    parser.add_argument('--standalone', type=Path, help='Use an already built standalone HTML instead of rebuilding it')
    args = parser.parse_args()
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=True)
    if args.standalone:
        base = args.standalone.read_text(encoding='utf-8')
    else:
        with tempfile.TemporaryDirectory(prefix='tool-pwa-build-') as temporary:
            standalone = Path(temporary) / 'Tool-Manager.html'
            subprocess.run([sys.executable, str(ROOT / 'build.py'), '--output', str(standalone)], check=True)
            base = standalone.read_text(encoding='utf-8')
    head = '''<meta data-pwa-shell name="theme-color" content="#123e37">
<meta data-pwa-shell name="apple-mobile-web-app-capable" content="yes">
<meta data-pwa-shell name="apple-mobile-web-app-status-bar-style" content="default">
<meta data-pwa-shell name="apple-mobile-web-app-title" content="Tool Manager">
<meta data-pwa-shell name="tool-manager-build" content="__PWA_BUILD__">
<link data-pwa-shell rel="manifest" href="./manifest.webmanifest">
<link data-pwa-shell rel="apple-touch-icon" href="./icons/apple-touch-icon.png">
<link data-pwa-shell rel="icon" type="image/png" sizes="192x192" href="./icons/icon-192.png">
<style data-pwa-shell>.pwa-panel{margin:16px 0}.pwa-actions{display:flex;flex-wrap:wrap;gap:8px;margin:10px 0}.pwa-panel p{margin:5px 0}.pwa-panel button{min-height:44px}.pwa-panel .pwa-help{padding:10px;background:#f2f6f3;border-radius:8px}@media(display-mode:standalone){.pwa-panel .pwa-intro{display:none}}</style>
'''
    panel = '''<section class="panel pwa-panel" data-pwa-shell aria-label="앱 설치와 오프라인 상태"><p class="pwa-intro"><strong>휴대폰 홈 화면에서 앱으로 사용</strong></p><div class="pwa-actions"><button id="pwa-install" type="button">홈 화면에 추가</button><button id="pwa-update" type="button" hidden>새 버전 적용</button></div><p id="pwa-state" class="muted" role="status" aria-live="polite">오프라인 앱 저장을 준비하고 있습니다.</p><p id="pwa-install-help" class="muted pwa-help" hidden></p></section>'''
    canonical = '''<script data-pwa-shell>(function(){if(/^https?:$/.test(location.protocol)&&location.pathname.endsWith('/'))history.replaceState(null,'','index.html'+location.search+location.hash);})();</script>'''
    if '<head>' not in base:
        raise ValueError('The standalone program has no head element')
    html = base.replace('<head>', '<head>' + canonical, 1).replace('</head>', head + '</head>', 1)
    if '<main>' not in html:
        raise ValueError('The standalone program has no application main element')
    html = html.replace('<main>', '<main>' + panel, 1)
    html = html.replace('</body>', '<script data-pwa-shell src="./pwa.js"></script>\n</body>', 1)
    pwa = (ROOT / 'pwa.js').read_text(encoding='utf-8')
    worker = (ROOT / 'service-worker.js').read_text(encoding='utf-8')
    static = {'manifest.webmanifest': (ROOT / 'manifest.webmanifest').read_bytes()}
    for path in sorted((ROOT / 'icons').glob('*.png')):
        static['icons/' + path.name] = path.read_bytes()
    seed = html.encode() + pwa.encode() + worker.encode() + b''.join(name.encode() + data for name, data in sorted(static.items()))
    version = sha256(seed).hexdigest()[:24]
    assets = {'index.html': html.replace('__PWA_BUILD__', version).encode('utf-8'),
              'pwa.js': pwa.replace('__PWA_BUILD__', version).encode('utf-8'), **static}
    manifest = [{'path': './' + name, 'sha256': sha256(data).hexdigest(), 'bytes': len(data)} for name, data in sorted(assets.items())]
    allowed = set(assets) | {'service-worker.js', '.nojekyll', 'build-info.json', 'team-settings.json', 'direct.html'}
    unexpected = [str(path.relative_to(output)) for path in output.rglob('*') if path.is_file() and str(path.relative_to(output)) not in allowed]
    if unexpected:
        raise ValueError('PWA output contains unrelated files; choose a separate empty directory: ' + ', '.join(unexpected))
    for name, data in assets.items():
        path = output / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
    # Fetch afresh on launch; do not put mutable team configuration in the
    # immutable service-worker shell or its install checksums.
    (output / 'team-settings.json').write_bytes(published_settings())
    # The direct entry opens the standalone program and bypasses an older
    # installed shell's index cache. It needs no site installation or manifest.
    direct = base.replace('<head>', '<head><meta name="tool-manager-direct-entry" content="true">', 1)
    (output / 'direct.html').write_text(direct, encoding='utf-8')
    script = worker.replace('__PWA_BUILD__', version).replace('/*__PWA_ASSETS__*/[]', json.dumps(manifest, separators=(',', ':')))
    (output / 'service-worker.js').write_text(script, encoding='utf-8')
    (output / '.nojekyll').write_text('', encoding='utf-8')
    info = {'version': version, 'standalone_sha256': sha256(base.encode()).hexdigest(), 'direct_entry': './direct.html',
            'index_sha256': sha256(assets['index.html']).hexdigest(), 'shell': manifest,
            'private_data_included': False}
    (output / 'build-info.json').write_text(json.dumps(info, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(f'Built PWA {version}: {output} ({sum(map(len, assets.values())):,} bytes; public application shell only)')


if __name__ == '__main__':
    main()
