"""Run the real PC OCR/storage harness on one unchanged uploaded photograph.

Photo paths, expected serials and dates are command-line inputs kept outside
the repository. Expectations are used by Python assertions only; the OCR sees
the original image bytes and the workbook. Owner identity is read from the
selected workbook row. The shared harness verifies actual OPFS writes, exact
backups and preservation of all other cells, styles and pictures.
"""
from __future__ import annotations

import argparse
from datetime import date
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unicodedata

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]


def owner_from_workbook(workbook: Path, serial: str) -> str:
    owners = []
    book = load_workbook(workbook, read_only=True, data_only=True)
    try:
        for sheet in book:
            rows = sheet.iter_rows(values_only=True)
            headers = next(rows, ())
            if "Serial" not in headers or "Asset #" not in headers:
                continue
            serial_column, asset_column = headers.index("Serial"), headers.index("Asset #")
            for row in rows:
                if str(row[serial_column]) == serial:
                    owner = unicodedata.normalize("NFKC", str(row[asset_column])).split("_")[-1].strip()
                    owners.append(owner)
    finally:
        book.close()
    if len(owners) != 1 or not owners[0]:
        raise ValueError("Expected serial must identify one named workbook row")
    return owners[0]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--program", type=Path, default=ROOT / "docs/pc.html")
    parser.add_argument("--photo", type=Path, required=True)
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--serial", required=True)
    parser.add_argument("--last", required=True)
    parser.add_argument("--due", required=True)
    parser.add_argument("--diagnostics", type=Path, default=Path("/tmp/tool-blurry-photo-regression.json"))
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    parser.add_argument("--diagnose-only", action="store_true")
    parser.add_argument("--baseline", type=Path)
    parser.add_argument("--maximum-seconds", type=float)
    args = parser.parse_args()
    args.workbook = args.workbook or next(Path("/workspace/attachments").rglob("*.xlsx"), None)
    for label, path in [("program", args.program), ("photo", args.photo), ("workbook", args.workbook)]:
        if path is None or not path.is_file():
            parser.error("Supply --" + label + " with an existing input")
    if date.fromisoformat(args.last) >= date.fromisoformat(args.due):
        parser.error("Expected next calibration date must follow the last date")
    owner = owner_from_workbook(args.workbook, args.serial)
    with tempfile.TemporaryDirectory(prefix="tool-blurry-regression-") as directory:
        manifest = Path(directory) / "private-expectations.json"
        manifest.write_text(json.dumps([{"label": "uploaded-blurry-label", "photo": str(args.photo.resolve()),
                                         "name": owner, "serial": args.serial,
                                         "last": args.last, "due": args.due}], ensure_ascii=False))
        command = [sys.executable, str(ROOT / "tests/pc_final_ocr_test.py"),
                   "--program", str(args.program.resolve()), "--workbook", str(args.workbook.resolve()),
                   "--manifest", str(manifest), "--diagnostics", str(args.diagnostics),
                   "--chromium", args.chromium]
        if args.diagnose_only:
            command.append("--diagnose-only")
        if args.baseline:
            command += ["--baseline", str(args.baseline)]
        if args.maximum_seconds is not None:
            command += ["--maximum-seconds", str(args.maximum_seconds)]
        subprocess.run(command, check=True)


if __name__ == "__main__":
    main()
