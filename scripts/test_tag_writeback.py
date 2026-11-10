#!/usr/bin/env python3
"""Smoke-test catalog-style MWG tag write-back with exiftool -diff.

Copies a photo to /tmp, applies a sample Rating/Description/Keywords payload
(same shape as Append), and prints a filtered diff. Asserts expected tags
changed and camera Make/Model/DateTimeOriginal did not.

Example:
  python3 scripts/test_tag_writeback.py /path/to/photo.jpg
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def find_exiftool() -> str:
    for candidate in ("/usr/local/bin/exiftool", "/opt/homebrew/bin/exiftool", "exiftool"):
        path = shutil.which(candidate) if candidate == "exiftool" else candidate
        if path and Path(path).exists():
            return path
    raise SystemExit("exiftool not found. Install with: brew install exiftool")


NOISY = re.compile(
    r"(FileModifyDate|FileAccessDate|FileInodeChangeDate|Directory|FilePermissions|"
    r"ExifToolVersion|EncodingProcess|YCbCrSubSampling)",
    re.I,
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="MWG write-back smoke test via exiftool -diff.")
    parser.add_argument("photo", type=Path, help="Source photo to copy and tag")
    args = parser.parse_args(argv)

    src = args.photo.expanduser().resolve()
    if not src.is_file():
        raise SystemExit(f"Not a file: {src}")

    exiftool = find_exiftool()
    with tempfile.TemporaryDirectory(prefix="strider_writeback_") as tmp:
        copy = Path(tmp) / src.name
        shutil.copy2(src, copy)
        before = Path(tmp) / "before.jpg"
        shutil.copy2(copy, before)

        cmd = [
            exiftool,
            "-m",
            "-P",
            "-overwrite_original",
            "-api",
            "LargeFileSupport=1",
            "-use",
            "MWG",
            "-charset",
            "iptc=UTF8",
            "-IPTC:CodedCharacterSet=UTF8",
            "-sep",
            "||",
            "-MWG:Rating=3",
            "-MWG:Description=writeback smoke notes",
            "-MWG:Keywords=Species:Aquarius||Behavior:stationary||Sequence:no",
            "--",
            str(copy),
        ]
        proc = subprocess.run(cmd, capture_output=True, text=True)
        if proc.returncode not in (0, 1):
            print(proc.stderr or proc.stdout, file=sys.stderr)
            return 1

        diff = subprocess.run(
            [exiftool, "-diff", str(before), str(copy)],
            capture_output=True,
            text=True,
        )
        lines = [
            ln
            for ln in (diff.stdout or "").splitlines()
            if ln.strip() and not NOISY.search(ln)
        ]
        text = "\n".join(lines)
        print(text)

        lower = text.lower()
        ok = True
        for needle in ("rating", "description", "keywords", "species"):
            if needle not in lower:
                print(f"FAIL: expected {needle} in diff", file=sys.stderr)
                ok = False
        for frozen in ("make", "model", "datetimeoriginal"):
            # -diff shows changes only; frozen tags should not appear as changed values
            # Look for "<" / ">" change lines mentioning these poorly — simpler read-back:
            pass

        make_b = subprocess.run(
            [exiftool, "-s3", "-Make", "-Model", "-DateTimeOriginal", "--", str(before)],
            capture_output=True,
            text=True,
        ).stdout
        make_a = subprocess.run(
            [exiftool, "-s3", "-Make", "-Model", "-DateTimeOriginal", "--", str(copy)],
            capture_output=True,
            text=True,
        ).stdout
        if make_b != make_a:
            print("FAIL: Make/Model/DateTimeOriginal changed", file=sys.stderr)
            print("before:", make_b, file=sys.stderr)
            print("after:", make_a, file=sys.stderr)
            ok = False
        else:
            print("OK: Make/Model/DateTimeOriginal unchanged")

        if ok:
            print("OK: write-back smoke passed")
            return 0
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
