#!/usr/bin/env python3
"""Fetch pinned, integrity-verified npm assets for the offline HTML program.

Only web/vendor is written. No package manifest, application source, or lockfile
is changed. npm's normal TLS and integrity verification remain enabled; the
downloaded tarballs are also checked against the SHA-512 values pinned below.
"""
from __future__ import annotations

import argparse
import base64
import gzip
import hashlib
import json
from pathlib import Path
import subprocess
import tarfile
from urllib.parse import urlparse


ROOT = Path(__file__).resolve().parents[1]
REGISTRY = "https://registry.npmjs.org"
PACKAGES = {
    "jszip@3.10.1": {
        "name": "jszip", "version": "3.10.1", "archive": "jszip-3.10.1.tgz",
        "integrity": "sha512-xXDvecyTpGLrqFrvkrUSoxxfJI5AH7U8zxxtVclpsUtMCq4JQ290LY8AW5c7Ggnr/Y/oK+bQMbqK2qmtk3pN4g==",
        "assets": {"package/dist/jszip.min.js": "jszip.min.js", "package/LICENSE.markdown": "JSZIP-LICENSE.txt"},
    },
    "tesseract.js@5.1.1": {
        "name": "tesseract.js", "version": "5.1.1", "archive": "tesseract.js-5.1.1.tgz",
        "integrity": "sha512-lzVl/Ar3P3zhpUT31NjqeCo1f+D5+YfpZ5J62eo2S14QNVOmHBTtbchHm/YAbOOOzCegFnKf4B3Qih9LuldcYQ==",
        "assets": {
            "package/dist/tesseract.min.js": "tesseract.min.js",
            "package/dist/worker.min.js": "worker.min.js",
            "package/LICENSE.md": "TESSERACT-LICENSE.txt",
            "package/dist/tesseract.min.js.LICENSE.txt": "TESSERACT-BROWSER-NOTICES.txt",
            "package/dist/worker.min.js.LICENSE.txt": "TESSERACT-WORKER-NOTICES.txt",
        },
    },
    "tesseract.js-core@5.1.1": {
        "name": "tesseract.js-core", "version": "5.1.1", "archive": "tesseract.js-core-5.1.1.tgz",
        "integrity": "sha512-KX3bYSU5iGcO1XJa+QGPbi+Zjo2qq6eBhNjSGR5E5q0JtzkoipJKOUQD7ph8kFyteCEfEQ0maWLu8MCXtvX5uQ==",
        "assets": {"package/tesseract-core-lstm.wasm.js": "tesseract-core-lstm.wasm.js", "package/LICENSE": "TESSERACT-CORE-LICENSE.txt"},
    },
    "@tesseract.js-data/eng@1.0.0": {
        "name": "@tesseract.js-data/eng", "version": "1.0.0", "archive": "tesseract.js-data-eng-1.0.0.tgz",
        "integrity": "sha512-mbTumm6KQPUHyzTPQaF3ObXYnx0SqqfV2nabqFVQBwD6Kl7PhGSLSzOlfFTWy0P3BjghaSKA2W9GB19Jk+ZcTg==",
        "assets": {"package/4.0.0_best_int/eng.traineddata.gz": "eng.traineddata.gz"},
    },
    "@tesseract.js-data/kor@1.0.0": {
        "name": "@tesseract.js-data/kor", "version": "1.0.0", "archive": "tesseract.js-data-kor-1.0.0.tgz",
        "integrity": "sha512-tepYhsYiRJ8E44R27MHxVteEDDLfYQxuvM6u4JlEHcCYqwpefvI5TGzoMgB67kunOvtyEJYtT5zfJ/6R6FVU9w==",
        "assets": {"package/4.0.0_best_int/kor.traineddata.gz": "kor.traineddata.gz"},
    },
}


def npm_json(arguments: list[str], cache: Path):
    command = ["npm", "--registry=" + REGISTRY, "--cache=" + str(cache), "--strict-ssl=true", "--ignore-scripts", *arguments]
    result = subprocess.run(command, check=True, capture_output=True, text=True)
    return json.loads(result.stdout)


def verify_archive(path: Path, integrity: str) -> bytes:
    data = path.read_bytes()
    actual = "sha512-" + base64.b64encode(hashlib.sha512(data).digest()).decode("ascii")
    if actual != integrity:
        raise ValueError(f"Integrity verification failed for {path.name}; refusing to use this archive.")
    return data


def check_assets(vendor: Path) -> dict:
    manifest = json.loads((vendor / "checksums.json").read_text(encoding="utf-8"))
    expected_files = {filename for package in PACKAGES.values() for filename in package["assets"].values()}
    if set(manifest["files"]) != expected_files:
        raise ValueError("The vendor manifest contains unexpected or missing assets.")
    if set(manifest["packages"]) != set(PACKAGES):
        raise ValueError("The vendor manifest contains unexpected or missing packages.")
    for spec, pin in PACKAGES.items():
        if manifest["packages"][spec]["integrity"] != pin["integrity"]:
            raise ValueError(f"Pinned npm integrity differs for {spec}.")
    for name, expected in manifest["files"].items():
        data = (vendor / name).read_bytes()
        if len(data) != expected["size"] or hashlib.sha256(data).hexdigest() != expected["sha256"]:
            raise ValueError(f"Asset checksum verification failed: {name}.")
    return manifest


def prepare(vendor: Path, cache: Path, archives: Path) -> dict:
    vendor.mkdir(parents=True, exist_ok=True)
    archives.mkdir(parents=True, exist_ok=True)
    manifest = {"format": 1, "registry": REGISTRY, "packages": {}, "files": {}}
    staged = {}
    for spec, pin in PACKAGES.items():
        metadata = npm_json(["view", spec, "version", "dist.integrity", "dist.tarball", "--json"], cache)
        if metadata.get("version") != pin["version"] or metadata.get("dist.integrity") != pin["integrity"]:
            raise ValueError(f"Official registry metadata does not match the pinned version/integrity for {spec}.")
        tarball_url = metadata.get("dist.tarball", "")
        parsed = urlparse(tarball_url)
        if parsed.scheme != "https" or parsed.hostname != "registry.npmjs.org" or parsed.username or parsed.password:
            raise ValueError(f"Unexpected tarball origin for {spec}; only the official HTTPS npm registry is accepted.")
        archive = archives / pin["archive"]
        if archive.exists():
            verify_archive(archive, pin["integrity"])
        else:
            packed = npm_json(["pack", spec, "--json", "--pack-destination=" + str(archives)], cache)
            if len(packed) != 1 or packed[0].get("integrity") != pin["integrity"] or packed[0].get("filename") != pin["archive"]:
                raise ValueError(f"npm pack returned an unexpected artifact for {spec}.")
        archive_data = verify_archive(archive, pin["integrity"])
        with tarfile.open(archive, "r:gz") as package:
            package_info = json.load(package.extractfile("package/package.json"))
            if package_info["name"] != pin["name"] or package_info["version"] != pin["version"]:
                raise ValueError(f"The verified archive has unexpected package metadata: {spec}.")
            for member, name in pin["assets"].items():
                entry = package.getmember(member)
                if not entry.isfile() or entry.size > 10_000_000:
                    raise ValueError(f"Unexpected archive member: {spec}/{member}.")
                data = package.extractfile(entry).read()
                if name.endswith(".traineddata.gz"):
                    unpacked = gzip.decompress(data)
                    if len(unpacked) < 100_000:
                        raise ValueError(f"Unexpected OCR model size: {name}.")
                if name == "tesseract-core-lstm.wasm.js" and b"data:application/octet-stream;base64,AGFzbQ" not in data:
                    raise ValueError("The selected core does not contain embedded WASM; refusing a network-dependent build.")
                staged[name] = data
                manifest["files"][name] = {"size": len(data), "sha256": hashlib.sha256(data).hexdigest(), "package": spec, "path": member}
        manifest["packages"][spec] = {"integrity": pin["integrity"], "archiveSha256": hashlib.sha256(archive_data).hexdigest(), "tarball": tarball_url, "license": package_info.get("license", "")}
        print(f"Verified {spec}", flush=True)
    # Write only after every archive, model, and embedded core has been verified.
    for name, data in staged.items():
        temporary = vendor / (name + ".tmp")
        temporary.write_bytes(data)
        temporary.replace(vendor / name)
    (vendor / "checksums.json").write_text(json.dumps(manifest, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return check_assets(vendor)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="Verify existing assets without network access or npm.")
    parser.add_argument("--vendor", type=Path, default=ROOT / "web/vendor")
    parser.add_argument("--cache", type=Path, default=Path("/tmp/tool-manager-npm-cache"))
    parser.add_argument("--archives", type=Path, default=Path("/tmp/tool-manager-npm-packages"))
    args = parser.parse_args()
    manifest = check_assets(args.vendor) if args.check else prepare(args.vendor, args.cache, args.archives)
    total = sum(item["size"] for item in manifest["files"].values())
    print(f"Verified {len(manifest['files'])} assets ({total:,} bytes) in {args.vendor}.")


if __name__ == "__main__":
    try:
        main()
    except (ValueError, OSError, subprocess.CalledProcessError) as error:
        raise SystemExit(f"Asset preparation stopped: {error}") from error
