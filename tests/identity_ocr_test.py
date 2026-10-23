"""Exercise the real offline identity reader on original and synthetic images.

The actual control is the unmodified tool-system screenshot embedded in the
company workbook. The separated-card fixtures are explicitly synthetic metal
layouts: they preserve the original calibration photo pixels, but rearrange
printed owner/serial tags. They do not reproduce the unavailable new handwritten
JPEG. Real Tesseract responses are observed, never replaced or corrected.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
import json
import math
from pathlib import Path
import tempfile
import time

from PIL import Image, ImageDraw, ImageFont
from playwright.sync_api import sync_playwright

from browser_test import REPO, serve_program
from torque_ocr_test import extract_sources


FONT = Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc")


def arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--program", type=Path, default=REPO / "web/dist/Tool-Manager.html")
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--serial", default="0716805829")
    parser.add_argument("--chromium", default="/usr/bin/chromium")
    parser.add_argument("--diagnose-only", action="store_true")
    parser.add_argument("--diagnostics", type=Path, default=Path("/tmp/tool-identity-ocr.json"))
    args = parser.parse_args()
    if args.workbook is None:
        args.workbook = next(Path("/workspace/attachments").rglob("*.xlsx"), None)
    if args.workbook is None or not args.workbook.is_file():
        parser.error("Supply --workbook with the attached company XLSX.")
    if not args.program.is_file():
        parser.error("Supply --program with the standalone HTML program.")
    if not FONT.is_file():
        parser.error("Install fonts-noto-cjk before running the printed synthetic fixtures.")
    return args


def curved_text(image, text, position, size, amplitude):
    """Draw explicitly generated characters on a shallow curved baseline."""
    font = ImageFont.truetype(str(FONT), size)
    x, y = position
    for index, character in enumerate(text):
        phase = index / max(1, len(text) - 1) * math.pi
        glyph = Image.new("RGBA", (size * 2, size * 2), (255, 255, 255, 0))
        ImageDraw.Draw(glyph).text((size * 0.2, 0), character, font=font, fill=(25, 25, 25, 255))
        glyph = glyph.rotate(math.cos(phase) * 3, resample=Image.Resampling.BICUBIC)
        image.paste(glyph, (round(x - size * 0.2), round(y + math.sin(phase) * amplitude)), glyph)
        x += font.getlength(character)


def synthetic_layout(output, expected, originals, curved=False):
    width, height = 1800, 2600
    image = Image.new("RGB", (width, height))
    pixels = image.load()
    for x in range(width):
        metal = 135 + int(20 * math.sin(x / 70))
        if 600 <= x <= 1250:
            metal += max(0, 70 - int(abs(x - 925) * 0.22))
        for y in range(height):
            shade = max(85, min(235, metal + int(5 * math.cos(y / 210))))
            pixels[x, y] = (shade, shade, shade)
    draw = ImageDraw.Draw(image)
    for x in (610, 622, 1260, 1274):
        draw.line([(x, 0), (x, height)], fill=(90, 90, 90), width=5)
    draw.rounded_rectangle((90, 155, 485, 340), radius=8, fill="white", outline=(100, 100, 100), width=2)
    font = ImageFont.truetype(str(FONT), 90)
    if curved:
        curved_text(image, expected["name"], (130, 185), 90, 7)
    else:
        draw.text((130, 185), expected["name"], font=font, fill=(25, 25, 25))

    # The native 80x24 serial snippet is intentionally tiny in the whole image.
    # Its location is a test-fixture crop in this particular original system
    # screenshot, not a production recognition coordinate or string correction.
    system = next(item for item in originals if item["kind"] == "system-screenshot")
    with Image.open(system["path"]) as screenshot:
        assert screenshot.size == (1519, 484), "Review the actual-source crop if the workbook changes"
        serial = screenshot.convert("RGB").crop((1295, 360, 1375, 384))
    if curved:
        draw.rectangle((1350, 935, 1755, 1060), fill="white", outline=(100, 100, 100), width=2)
        curved_text(image, expected["serial"], (1390, 970), 42, 5)
        serial_description = "generated printed serial; shallow curved baseline, not source pixels"
    else:
        draw.rectangle((1450, 935, 1755, 1040), fill="white", outline=(100, 100, 100), width=2)
        image.paste(serial, (1500, 980))
        assert image.crop((1500, 980, 1580, 1004)).tobytes() == serial.tobytes()
        serial_description = "unmodified native 80x24 pixels cropped from the actual system screenshot"

    source = next(item for item in originals if item["kind"] == "calibration-label")
    with Image.open(source["path"]) as original:
        label = original.convert("RGB")
    image.paste(label, (1340, 2050))
    assert image.crop((1340, 2050, 1340 + label.width, 2050 + label.height)).tobytes() == label.tobytes()
    image.save(output)
    # Check saved pixels too; no source label characters or dates are replaced.
    with Image.open(output) as saved:
        assert saved.crop((1340, 2050, 1340 + label.width, 2050 + label.height)).tobytes() == label.tobytes()
    return {"kind": "synthetic-curved-tags" if curved else "synthetic-separated-native-tags",
            "synthetic": True, "path": output, "dimensions": [width, height],
            "sha256": sha256(output.read_bytes()).hexdigest(),
            "serial_pixels": serial_description,
            "calibration_photo": {"media": source["media"], "sha256": source["sha256"],
                                  "original_pixels_preserved": True},
            "not_a_reproduction_of_the_new_inline_photo": True}


OBSERVE_WORKER = """() => {
  window.identityReadings=[];const create=Tesseract.createWorker.bind(Tesseract);
  Tesseract.createWorker=async(...args)=>{
    const worker=await create(...args),recognize=worker.recognize.bind(worker);
    let parameters={};const set=worker.setParameters.bind(worker);
    worker.setParameters=async values=>{parameters={...parameters,...values};return set(values);};
    worker.recognize=async(...input)=>{
      const began=performance.now(),response=await recognize(...input),data=response.data||{};
      identityReadings.push({text:data.text,confidence:data.confidence,
        milliseconds:Math.round(performance.now()-began),
        input:{width:input[0].width,height:input[0].height},parameters:{...parameters},
        words:(data.words||[]).map(word=>({text:word.text,confidence:word.confidence,bbox:word.bbox,
          symbols:(word.symbols||[]).map(symbol=>({text:symbol.text,confidence:symbol.confidence,bbox:symbol.bbox}))}))});
      return response;
    };return worker;
  };
}"""


def main():
    args = arguments()
    original_hash = sha256(args.workbook.read_bytes()).hexdigest()
    diagnostics = {"program": str(args.program), "program_sha256": sha256(args.program.read_bytes()).hexdigest(),
                   "diagnose_only": args.diagnose_only, "new_inline_photo_tested": False,
                   "worker_results_replaced": False, "cases": []}
    failures, errors, requests = [], [], []
    with tempfile.TemporaryDirectory(prefix="tool-real-identity-") as directory:
        temporary = Path(directory)
        expected, originals = extract_sources(args, temporary)
        diagnostics["worksheet_reference"] = expected
        system = next(item for item in originals if item["kind"] == "system-screenshot")
        cases = [{**system, "synthetic": False},
                 synthetic_layout(temporary / "separated-native-tags.png", expected, originals),
                 synthetic_layout(temporary / "separated-curved-tags.png", expected, originals, curved=True)]
        with serve_program(args.program) as url, sync_playwright() as playwright:
            browser = playwright.chromium.launch(executable_path=args.chromium, headless=True, args=["--no-sandbox"])
            context = browser.new_context(accept_downloads=True, locale="ko-KR")
            page = context.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))
            page.goto(url)
            page.wait_for_function("window.ToolApp&&window.ToolOCR&&window.ToolRuntime")
            page.evaluate(OBSERVE_WORKER)
            context.route("**/*", lambda route: (requests.append(route.request.url), route.abort())
                          if route.request.url.startswith(("http:", "https:")) else route.continue_())
            context.set_offline(True)
            page.locator("#workbook").set_input_files(args.workbook)
            page.wait_for_function("ToolApp.state.model&&!ToolApp.state.busy")
            target_id = page.evaluate("serial=>ToolApp.state.model.tools.find(tool=>tool.serial===serial).id", expected["serial"])
            # These are reported OCR fragments, not an OCR-accuracy fixture.
            # With confidence absent, they must not become a nearest-serial hit.
            negative = page.evaluate("""() => {
              const text='071680582qg\\n한거으\\n2026.06.22\\n2026.12.22\\n0716805808';
              const parsed=ToolOCR.parse(text,ToolApp.state.model.tools);
              return {confidence:parsed.confidence,serial:parsed.serial,
                matches:parsed.identityEvidence.serialMatches,
                assessment:ToolAutomation.assess(ToolApp.state.model,parsed)};
            }""")
            assert negative["confidence"] is None and negative["serial"] == "" and not negative["matches"], negative
            assert not negative["assessment"]["eligible"], negative
            diagnostics["reported_fragments_without_pixel_evidence"] = negative
            for case in cases:
                page.evaluate("identityReadings=[]")
                began = time.monotonic()
                page.locator("#photos").set_input_files(case["path"])
                page.wait_for_function("ToolApp.state.photos.length===1&&!ToolApp.state.busy&&!ToolApp.state.preparing", timeout=360000)
                evidence = page.evaluate("""() => {const photo=ToolApp.state.photos[0];return {
                  result:photo.result,assessment:photo.assessment,approval:photo.approval,
                  prepared:!!ToolApp.state.preparedResult,readings:identityReadings};}""")
                report = {key: value for key, value in case.items() if key != "path"}
                report.update(evidence)
                report["seconds"] = round(time.monotonic() - began, 2)
                diagnostics["cases"].append(report)
                args.diagnostics.write_text(json.dumps(diagnostics, ensure_ascii=False, indent=2), encoding="utf-8")
                result = evidence["result"] or {}
                proofs = [proof for proof in (result.get("identityEvidence") or {}).get("serialMatches", [])
                          if proof.get("toolId") == target_id and proof.get("serial") == expected["serial"]]
                print(json.dumps({"kind": case["kind"], "synthetic": case["synthetic"],
                                  "name": result.get("name"), "serial": result.get("serial"),
                                  "serial_proof": proofs, "prepared": evidence["prepared"],
                                  "seconds": report["seconds"],
                                  "pass_sources": [item["source"] for item in result.get("passes", [])]}, ensure_ascii=False, indent=2), flush=True)
                if not args.diagnose_only:
                    if result.get("failed"):
                        failures.append(case["kind"] + ": real OCR failed")
                    if result.get("serial") != expected["serial"] or not proofs or (proofs[0].get("confidence") or 0) < 70:
                        failures.append(case["kind"] + ": exact leading-zero serial was not read confidently")
                    if case["kind"] == "synthetic-separated-native-tags":
                        if result.get("name") != expected["name"]:
                            failures.append(case["kind"] + ": separate owner card was not read")
                        if not any(item["source"].startswith("identity-probe-panel-") and expected["serial"] in item["text"] for item in result.get("passes", [])):
                            failures.append(case["kind"] + ": serial was not recovered from an independent panel")
                    if not case["synthetic"]:
                        assert not evidence["prepared"], "The actual system screenshot lacks Cal Last"
                    if evidence["prepared"]:
                        assessment = evidence["assessment"] or {}
                        assert assessment.get("eligible") and assessment.get("toolId") == target_id, assessment
                        assert (assessment.get("last"), assessment.get("due")) == (expected["last"], expected["due"]), assessment
                page.locator("#remove-photo").click()
                page.wait_for_function("ToolApp.state.photos.length===0&&!ToolApp.state.preparing")
            assert not requests, requests
            assert not errors, errors
            browser.close()
    assert original_hash == sha256(args.workbook.read_bytes()).hexdigest(), "Original company workbook changed"
    diagnostics.update({"failures": failures, "http_requests": requests, "page_errors": errors, "original_unchanged": True})
    args.diagnostics.write_text(json.dumps(diagnostics, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"actual_original_cases": 1, "explicit_synthetic_cases": 2,
                      "new_inline_photo_tested": False, "failures": failures,
                      "http_requests": requests, "original_unchanged": True,
                      "diagnostics": str(args.diagnostics)}, ensure_ascii=False, indent=2))
    assert not failures, failures


if __name__ == "__main__":
    main()
