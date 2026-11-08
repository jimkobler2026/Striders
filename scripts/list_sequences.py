#!/usr/bin/env python3
"""List sequence IDs in a folder as CSV (SequenceID, StartFile, EndFile, Frames).

Reads Sequence* Keywords written by tag_sequences.py.

Example:
  python3 scripts/list_sequences.py "/path/to/photos" > sequences.csv
  python3 scripts/list_sequences.py "/path/to/photos" -o sequences.csv
"""

from __future__ import annotations

import argparse
import csv
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
    parser = argparse.ArgumentParser(description="CSV list of sequence IDs with start/end filenames.")
    parser.add_argument("directory", type=Path, help="Folder of photos")
    parser.add_argument("-o", "--output", type=Path, help="Write CSV to this file (default: stdout)")
    parser.add_argument("--recursive", action="store_true", help="Scan subfolders")
    args = parser.parse_args(argv)

    root = args.directory.expanduser().resolve()
    if not root.is_dir():
        raise SystemExit(f"Not a directory: {root}")

    paths: list[Path] = []
    iterator = root.rglob("*") if args.recursive else root.iterdir()
    for p in iterator:
        if p.is_file() and p.suffix.lower() in PHOTO_EXTS:
            paths.append(p)
    paths.sort()
    if not paths:
        raise SystemExit(f"No photos found in {root}")

    exiftool = find_exiftool()
    proc = subprocess.run(
        [exiftool, "-json", "-Keywords", "-Subject", "--", *[str(p) for p in paths]],
        capture_output=True,
        text=True,
    )
    if proc.returncode not in (0, 1):
        raise SystemExit(proc.stderr or proc.stdout or "exiftool failed")
    rows = json.loads(proc.stdout or "[]")
    by_src = {Path(r.get("SourceFile", "")).resolve(): r for r in rows}

    # SequenceID -> list of (pos_index, filename)
    groups: dict[str, list[tuple[int, str]]] = {}
    for p in paths:
        row = by_src.get(p.resolve()) or {}
        tags = parse_seq_keywords(row.get("Keywords"))
        if not tags:
            tags = parse_seq_keywords(row.get("Subject"))
        if tags.get("Sequence", "").lower() != "yes":
            continue
        seq_id = tags.get("SequenceID") or ""
        if not seq_id:
            continue
        groups.setdefault(seq_id, []).append((pos_index(tags.get("SequencePos", "")), p.name))

    out_rows = []
    for seq_id in sorted(groups):
        frames = sorted(groups[seq_id], key=lambda t: (t[0], t[1]))
        out_rows.append(
            {
                "SequenceID": seq_id,
                "StartFile": frames[0][1],
                "EndFile": frames[-1][1],
                "Frames": len(frames),
            }
        )

    fieldnames = ["SequenceID", "StartFile", "EndFile", "Frames"]

    def write_csv(fh) -> None:
        fh.write(f"Folder,{root}\n")
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(out_rows)

    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        with args.output.open("w", newline="", encoding="utf-8") as fh:
            write_csv(fh)
        print(f"Wrote {len(out_rows)} sequences → {args.output}", file=sys.stderr)
    else:
        write_csv(sys.stdout)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
