#!/usr/bin/env python3
"""Stamp Artist / Copyright (MWG Creator/Copyright) onto media files.

Configure via env (not hard-coded secrets):
  STRIDER_ARTIST     → MWG:Creator / Artist
  STRIDER_COPYRIGHT  → MWG:Copyright / Copyright

Dry-run by default. Idempotent when values already match.

Example:
  export STRIDER_ARTIST="Your Name"
  export STRIDER_COPYRIGHT="© 2026 Your Name"
  python3 scripts/stamp_ownership.py /tmp/copy.jpg --write
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

PHOTO_EXTS = {".jpg", ".jpeg", ".tif", ".tiff", ".arw", ".cr2", ".cr3", ".nef", ".dng", ".heic"}


def find_exiftool() -> str:
    for candidate in ("/usr/local/bin/exiftool", "/opt/homebrew/bin/exiftool", "exiftool"):
        path = shutil.which(candidate) if candidate == "exiftool" else candidate
        if path and Path(path).exists():
            return path
    raise SystemExit("exiftool not found. Install with: brew install exiftool")


def collect(root: Path, recursive: bool) -> list[Path]:
    if root.is_file():
        return [root]
    if recursive:
        return sorted(p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in PHOTO_EXTS)
    return sorted(p for p in root.iterdir() if p.is_file() and p.suffix.lower() in PHOTO_EXTS)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Stamp MWG Creator/Copyright onto media.")
    parser.add_argument("path", type=Path, help="File or folder")
    parser.add_argument("--recursive", action="store_true")
    parser.add_argument("--write", action="store_true")
    parser.add_argument("--artist", default=os.environ.get("STRIDER_ARTIST", ""))
    parser.add_argument("--copyright", default=os.environ.get("STRIDER_COPYRIGHT", ""))
    args = parser.parse_args(argv)

    artist = args.artist.strip()
    copyright_ = args.copyright.strip()
    if not artist and not copyright_:
        raise SystemExit(
            "Set STRIDER_ARTIST / STRIDER_COPYRIGHT or pass --artist / --copyright."
        )

    root = args.path.expanduser().resolve()
    paths = collect(root, args.recursive)
    if not paths:
        raise SystemExit("No photos found.")

    exiftool = find_exiftool()
    print(f"Files: {len(paths)}  Mode: {'WRITE' if args.write else 'DRY-RUN'}")
    print(f"Artist: {artist!r}  Copyright: {copyright_!r}")
    if not args.write:
        print("Dry-run only. Re-run with --write to apply.")
        return 0

    cmd = [
        exiftool,
        "-m",
        "-P",
        "-overwrite_original",
        "-api",
        "LargeFileSupport=1",
        "-use",
        "MWG",
    ]
    if artist:
        cmd.append(f"-MWG:Creator={artist}")
    if copyright_:
        cmd.append(f"-MWG:Copyright={copyright_}")
    cmd.append("--")
    cmd.extend(str(p) for p in paths)
    proc = subprocess.run(cmd, capture_output=True, text=True)
    sys.stdout.write(proc.stdout or "")
    if proc.stderr:
        sys.stderr.write(proc.stderr)
    return 0 if proc.returncode in (0, 1) else 1


if __name__ == "__main__":
    raise SystemExit(main())
