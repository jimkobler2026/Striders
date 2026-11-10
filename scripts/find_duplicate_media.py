#!/usr/bin/env python3
"""Report duplicate image *content* via exiftool ImageDataHash.

Hashes image pixels only (ignores metadata and filename). Default is a
report-only dry-run — this tool does not delete files.

Example:
  python3 scripts/find_duplicate_media.py ~/Pictures/environmental\\ caustics/temp\\ photos\\ 11052026
  python3 scripts/find_duplicate_media.py /path/to/folder --recursive
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from collections import defaultdict
from pathlib import Path

PHOTO_EXTS = {".jpg", ".jpeg", ".tif", ".tiff", ".arw", ".cr2", ".cr3", ".nef", ".dng", ".heic"}


def find_exiftool() -> str:
    for candidate in ("/usr/local/bin/exiftool", "/opt/homebrew/bin/exiftool", "exiftool"):
        path = shutil.which(candidate) if candidate == "exiftool" else candidate
        if path and Path(path).exists():
            return path
    raise SystemExit("exiftool not found. Install with: brew install exiftool")


def collect_paths(root: Path, recursive: bool) -> list[Path]:
    if recursive:
        paths = [p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in PHOTO_EXTS]
    else:
        paths = [p for p in root.iterdir() if p.is_file() and p.suffix.lower() in PHOTO_EXTS]
    return sorted(paths)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Report ImageDataHash duplicates (metadata-independent; report only)."
    )
    parser.add_argument("directory", type=Path, help="Folder of photos")
    parser.add_argument("--recursive", action="store_true", help="Scan subfolders")
    args = parser.parse_args(argv)

    root = args.directory.expanduser().resolve()
    if not root.is_dir():
        raise SystemExit(f"Not a directory: {root}")

    paths = collect_paths(root, args.recursive)
    if not paths:
        print("No photos found.")
        return 0

    exiftool = find_exiftool()
    cmd = [
        exiftool,
        "-json",
        "-ImageDataHash",
        "-FileName",
        "-m",
        "--",
        *[str(p) for p in paths],
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode not in (0, 1):
        raise SystemExit(f"exiftool failed:\n{proc.stderr or proc.stdout}")

    rows = json.loads(proc.stdout or "[]")
    by_hash: dict[str, list[str]] = defaultdict(list)
    missing = 0
    for row in rows:
        digest = str(row.get("ImageDataHash") or "").strip()
        src = row.get("SourceFile") or row.get("FileName") or ""
        if not digest:
            missing += 1
            continue
        by_hash[digest].append(str(Path(src).resolve()) if src else src)

    groups = {h: files for h, files in by_hash.items() if len(files) > 1}
    print(f"Files scanned: {len(paths)}")
    print(f"Unique hashes: {len(by_hash)}  (no hash: {missing})")
    print(f"Duplicate groups: {len(groups)}")
    if not groups:
        print("No content duplicates found.")
        return 0

    print()
    for i, (digest, files) in enumerate(sorted(groups.items(), key=lambda kv: (-len(kv[1]), kv[0])), start=1):
        print(f"Group {i}  ({len(files)} files)  hash={digest}")
        for f in sorted(files):
            print(f"  {f}")
        print()
    print("Report only — no files deleted.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
