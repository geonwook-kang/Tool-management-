"""Exercise ordinary capture variations through actual offline PC OCR/storage.

The supplied photograph and expected serial/dates remain runtime inputs. PIL
only resizes, mildly blurs, changes brightness or rotates existing pixels in
temporary files; no label text or missing characters are drawn. Expectations
are Python assertions, never recognizer inputs. The existing PC harness checks
real OPFS saves, exact backups and all unrelated workbook cells/pictures.
"""
from __future__ import annotations

import argparse
from datetime import date
from hashlib import sha256
import json
from pathlib import Path
import subprocess
import sys
import tempfile

from PIL import Image, ImageEnhance, ImageFilter, ImageOps

from blur_photo_regression_test import owner_from_workbook


ROOT = Path(__file__).resolve().parents[1]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--program", type=Path, default=Path("/workspace/outputs/Tool-Manager-PC-Final.html"))
    parser.add_argument("--photo", type=Path, required=True)
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--serial", required=True)
    parser.add_argument("--last", required=True)
    parser.add_argument("--due", required=True)
    parser.add_argument("--diagnostics", type=Path, default=Path("/tmp/tool-blur-variants-regression.json"))
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    parser.add_argument("--diagnose-only", action="store_true")
    parser.add_argument("--maximum-seconds", type=float)
    parser.add_argument("--variant", action="append", help="Run only named variants; repeat to select several")
    args = parser.parse_args()
    args.workbook = args.workbook or next(Path("/workspace/attachments").rglob("*.xlsx"), None)
    for label, path in [("program", args.program), ("photo", args.photo), ("workbook", args.workbook)]:
        if path is None or not path.is_file():
            parser.error("Supply --" + label + " with an existing file")
    if date.fromisoformat(args.last) >= date.fromisoformat(args.due):
        parser.error("Expected next calibration date must follow the last date")
    owner = owner_from_workbook(args.workbook, args.serial)
    original_hash = sha256(args.photo.read_bytes()).hexdigest()
    with Image.open(args.photo) as original:
        image = ImageOps.exif_transpose(original).convert("RGB")
    variants = [
        ("resize-75-percent", image.resize((round(image.width * .75), round(image.height * .75)), Image.Resampling.LANCZOS)),
        ("resize-60-percent", image.resize((round(image.width * .6), round(image.height * .6)), Image.Resampling.LANCZOS)),
        ("mild-gaussian-blur-035", image.filter(ImageFilter.GaussianBlur(.35))),
        ("mild-gaussian-blur-060", image.filter(ImageFilter.GaussianBlur(.6))),
        ("brightness-80-percent", ImageEnhance.Brightness(image).enhance(.8)),
        ("brightness-115-percent", ImageEnhance.Brightness(image).enhance(1.15)),
        ("rotate-plus-1p5-degrees", image.rotate(1.5, resample=Image.Resampling.BICUBIC, expand=True, fillcolor="white")),
        ("rotate-minus-1p5-degrees", image.rotate(-1.5, resample=Image.Resampling.BICUBIC, expand=True, fillcolor="white")),
    ]
    if args.variant:
        unknown = set(args.variant) - {label for label, _ in variants}
        if unknown:
            parser.error("Unknown variants: " + ", ".join(sorted(unknown)))
        variants = [(label, variant) for label, variant in variants if label in args.variant]
    with tempfile.TemporaryDirectory(prefix="tool-blur-variants-") as directory:
        temporary = Path(directory)
        cases = []
        for label, variant in variants:
            photo = temporary / (label + ".jpg")
            variant.save(photo, format="JPEG", quality=97, subsampling=0)
            cases.append({"label": label, "photo": str(photo), "name": owner,
                          "serial": args.serial, "last": args.last, "due": args.due})
        manifest = temporary / "private-expectations.json"
        manifest.write_text(json.dumps(cases, ensure_ascii=False))
        command = [sys.executable, str(ROOT / "tests/pc_final_ocr_test.py"),
                   "--program", str(args.program.resolve()), "--workbook", str(args.workbook.resolve()),
                   "--manifest", str(manifest), "--diagnostics", str(args.diagnostics),
                   "--chromium", args.chromium]
        if args.diagnose_only:
            command.append("--diagnose-only")
        if args.maximum_seconds is not None:
            command += ["--maximum-seconds", str(args.maximum_seconds)]
        subprocess.run(command, check=True)
    assert sha256(args.photo.read_bytes()).hexdigest() == original_hash


if __name__ == "__main__":
    main()
