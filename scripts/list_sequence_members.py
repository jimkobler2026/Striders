#!/usr/bin/env python3
"""List photo paths in a folder that share a SequenceID (from Keywords).

Machine output for Excel (sorted by SequencePos index):
  OK
  /abs/path.jpg<TAB>1/12
  /abs/path2.jpg<TAB>2/12

Example:
  python3 scripts/list_sequence_members.py /path/to/folder --id 20261105-102425
  python3 scripts/list_sequence_members.py /path/to/file.jpg --id 20261105-102425 --machine
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

PHOTO_EXTS = {".jpg", ".jpeg", ".tif", ".tiff", ".arw", ".cr2", ".cr3", ".nef", ".dng", ".heic"}
POS_RE = re.compile(r"^(\d+)/(\d+)$")


def find_exiftool() -> str:
    for candidate in ("/usr/local/bin/exiftool", "/opt/homebrew/bin/exiftool", "exiftool"):
        path = shutil.which(candidate) if candidate == "exiftool" else candidate
        if path and Path(path).exists():
            return path
    raise SystemExit("exiftool not found. Install with: brew install exiftool")


def parse_seq_keywords(keywords) -> dict[str, str]:
    out: dict[str, str] = {}
    if keywords is None:
        return out
    items = keywords if isinstance(keywords, list) else [keywords]
    for item in items:
        text = str(item).strip()
        if ":" not in text:
            continue
        key, _, val = text.partition(":")
        key = key.strip()
        if key in ("Sequence", "SequenceID", "SequencePos"):
            out[key] = val.strip()
    return out


def pos_index(pos: str) -> int:
    m = POS_RE.match(pos or "")
    return int(m.group(1)) if m else 10**9


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="List files in a folder with a given SequenceID.")
    parser.add_argument("path", type=Path, help="Folder, or any file inside the folder")
    parser.add_argument("--id", required=True, dest="sequence_id", help="SequenceID to match")
    parser.add_argument(
        "--machine",
        action="store_true",
        help="Excel-friendly: OK then path<TAB>pos lines on stdout",
    )
    args = parser.parse_args(argv)

    anchor = args.path.expanduser().resolve()
    folder = anchor if anchor.is_dir() else anchor.parent
    if not folder.is_dir():
        print(f"ERROR\tNot a directory: {folder}", file=sys.stderr)
        return 2

    want = args.sequence_id.strip()
    if not want:
        print("ERROR\tEmpty SequenceID", file=sys.stderr)
        return 2

    paths = sorted(
        p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in PHOTO_EXTS
    )
    if not paths:
        if args.machine:
            print("OK")
        else:
            print("No photos in folder")
        return 0

    exiftool = find_exiftool()
    proc = subprocess.run(
        [exiftool, "-json", "-Keywords", "-Subject", "--", *[str(p) for p in paths]],
        capture_output=True,
        text=True,
    )
    if proc.returncode not in (0, 1):
        print(f"ERROR\t{proc.stderr or proc.stdout or 'exiftool failed'}", file=sys.stderr)
        return 1
    rows = json.loads(proc.stdout or "[]")
    by_src = {Path(r.get("SourceFile", "")).resolve(): r for r in rows}

    members: list[tuple[int, str, Path]] = []
    for p in paths:
        row = by_src.get(p.resolve()) or {}
        tags = parse_seq_keywords(row.get("Keywords")) or parse_seq_keywords(row.get("Subject"))
        if tags.get("Sequence", "").lower() != "yes":
            continue
        if tags.get("SequenceID", "") != want:
            continue
        pos = tags.get("SequencePos", "")
        members.append((pos_index(pos), pos, p))

    members.sort(key=lambda t: (t[0], t[2].name))

    if args.machine:
        print("OK")
        for _, pos, p in members:
            print(f"{p}\t{pos}")
        return 0

    print(f"Folder: {folder}")
    print(f"SequenceID: {want}")
    print(f"Members: {len(members)}")
    for _, pos, p in members:
        print(f"  {pos or '-':>8}  {p.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
