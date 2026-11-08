"""Package the standalone PC final entry without rebuilding mobile releases."""
from __future__ import annotations

import argparse
from hashlib import sha256
from pathlib import Path
import shutil
import subprocess
import sys
from zipfile import ZIP_DEFLATED, ZipFile


ROOT = Path(__file__).resolve().parent
VERSION = "7.1"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=Path("/workspace/outputs"))
    args = parser.parse_args()
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=True)

    html = output / "Tool-Manager-PC-Final.html"
    subprocess.run([sys.executable, str(ROOT / "build.py"), "--entry", "pc", "--output", str(html)], check=True)
    text = html.read_text(encoding="utf-8")
    if 'name="tool-manager-pc-entry"' not in text:
        raise ValueError("The PC release must contain its dedicated entry marker")
    guide = ROOT / "pc-guide.html"
    manual = ROOT / "PC-사용법.txt"
    shutil.copyfile(guide, output / guide.name)
    shutil.copyfile(manual, output / manual.name)
    archive = output / f"Tool-Manager-PC-Final-{VERSION}.zip"
    with ZipFile(archive, "w", compression=ZIP_DEFLATED, compresslevel=6) as bundle:
        bundle.write(html, "Tool-Manager.html")
        bundle.write(manual, manual.name)
        bundle.write(guide, guide.name)
        for license_file in sorted((ROOT / "vendor").glob("*.txt")):
            bundle.write(license_file, "licenses/" + license_file.name)
    with ZipFile(archive) as bundle:
        if bundle.testzip() is not None:
            raise ValueError("The PC release archive did not pass its CRC check")
    artifacts = [html, archive]
    checksums = "".join(sha256(path.read_bytes()).hexdigest() + "  " + path.name + "\n" for path in artifacts)
    (output / "PC-SHA256SUMS.txt").write_text(checksums, encoding="utf-8")
    for path in artifacts:
        print(f"Packaged {path}: {path.stat().st_size:,} bytes")


if __name__ == "__main__":
    main()
