#!/usr/bin/env python3
"""Copy user catalog tags from one sequence frame onto SequenceID siblings.

Copies MWG Keywords (except each file's own Sequence*), plus Description /
Rating / Location from a source frame onto siblings that share SequenceID.
Each target keeps its own Sequence / SequenceID / SequencePos.

Default dry-run. Paths may be any folder (document unrestricted — use care).

Example:
  python3 scripts/copy_sequence_tags.py /path/to/1ofN.jpg
  python3 scripts/copy_sequence_tags.py /path/to/1ofN.jpg --write
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
SEQ_KEYS = {"Sequence", "SequenceID", "SequencePos"}
POS_RE = re.compile(r"^(\d+)/(\d+)$")


def find_exiftool() -> str:
    for candidate in ("/usr/local/bin/exiftool", "/opt/homebrew/bin/exiftool", "exiftool"):
        path = shutil.which(candidate) if candidate == "exiftool" else candidate
        if path and Path(path).exists():
            return path
    raise SystemExit("exiftool not found. Install with: brew install exiftool")


def keyword_items(raw) -> list[str]:
    if raw is None:
        return []
    if isinstance(raw, list):
        return [str(x).strip() for x in raw if str(x).strip()]
    text = str(raw).strip()
    return [text] if text else []


def parse_seq(keywords: list[str]) -> dict[str, str]:
    out: dict[str, str] = {}
    for item in keywords:
        if ":" not in item:
            continue
        key, _, val = item.partition(":")
        key = key.strip()
        if key in SEQ_KEYS:
            out[key] = val.strip()
    return out


def user_keywords(keywords: list[str]) -> list[str]:
    return [k for k in keywords if k.split(":", 1)[0].strip() not in SEQ_KEYS]


def read_meta(exiftool: str, path: Path) -> dict:
    proc = subprocess.run(
        [
            exiftool,
            "-json",
            "-Keywords",
            "-Subject",
            "-MWG:Description",
            "-MWG:Rating",
            "-MWG:Location",
            "--",
            str(path),
        ],
        capture_output=True,
        text=True,
    )
    if proc.returncode not in (0, 1) or not proc.stdout:
        raise RuntimeError(proc.stderr or proc.stdout or "exiftool read failed")
    return json.loads(proc.stdout)[0]


def list_members(exiftool: str, folder: Path, sequence_id: str) -> list[tuple[Path, dict[str, str]]]:
    paths = sorted(p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in PHOTO_EXTS)
    if not paths:
        return []
    proc = subprocess.run(
        [exiftool, "-json", "-Keywords", "-Subject", "--", *[str(p) for p in paths]],
        capture_output=True,
        text=True,
    )
    if proc.returncode not in (0, 1):
        raise SystemExit(f"exiftool failed:\n{proc.stderr or proc.stdout}")
    rows = json.loads(proc.stdout or "[]")
    by = {Path(r.get("SourceFile", "")).resolve(): r for r in rows}
    out: list[tuple[Path, dict[str, str]]] = []
    for p in paths:
        row = by.get(p.resolve()) or {}
        kws = keyword_items(row.get("Keywords"))
        if not kws:
            kws = keyword_items(row.get("Subject"))
        seq = parse_seq(kws)
        if seq.get("SequenceID") == sequence_id:
            out.append((p, seq))
    out.sort(key=lambda t: int(POS_RE.match(t[1].get("SequencePos", "") or "").group(1)) if POS_RE.match(t[1].get("SequencePos", "") or "") else 10**9)
    return out


def write_target(
    exiftool: str,
    path: Path,
    *,
    merged_keywords: list[str],
    description: str | None,
    rating: str | None,
    location: str | None,
) -> None:
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
        f"-MWG:Keywords={'||'.join(merged_keywords)}",
    ]
    if description is not None:
        cmd.append(f"-MWG:Description={description}")
    if rating is not None and str(rating).strip() != "":
        cmd.append(f"-MWG:Rating={rating}")
    if location is not None:
        cmd.append(f"-MWG:Location={location}")
    cmd.extend(["--", str(path)])  # single-shot CLI: -- is fine (no -execute)
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode not in (0, 1):
        raise RuntimeError(proc.stderr or proc.stdout or "exiftool write failed")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Copy user MWG tags from one sequence frame onto SequenceID siblings."
    )
    parser.add_argument("source", type=Path, help="Source photo (usually the 1/n frame)")
    parser.add_argument("--write", action="store_true", help="Apply tags (default dry-run)")
    parser.add_argument(
        "--id",
        dest="sequence_id",
        default="",
        help="SequenceID (default: read from source)",
    )
    args = parser.parse_args(argv)

    source = args.source.expanduser().resolve()
    if not source.is_file():
        raise SystemExit(f"Not a file: {source}")

    exiftool = find_exiftool()
    src_meta = read_meta(exiftool, source)
    src_kws = keyword_items(src_meta.get("Keywords"))
    if not src_kws:
        src_kws = keyword_items(src_meta.get("Subject"))
    src_seq = parse_seq(src_kws)
    sequence_id = (args.sequence_id or src_seq.get("SequenceID") or "").strip()
    if not sequence_id:
        raise SystemExit("Source has no SequenceID; pass --id or tag sequences first.")

    user_kws = user_keywords(src_kws)
    description = src_meta.get("Description")
    if description is None:
        description = src_meta.get("MWG:Description")
    rating = src_meta.get("Rating")
    if rating is None:
        rating = src_meta.get("MWG:Rating")
    location = src_meta.get("Location")
    if location is None:
        location = src_meta.get("MWG:Location")

    members = list_members(exiftool, source.parent, sequence_id)
    if not members:
        raise SystemExit(f"No members found for SequenceID {sequence_id}")

    print(f"Source:      {source.name}")
    print(f"SequenceID:  {sequence_id}")
    print(f"User tags:   {len(user_kws)}  Rating={rating!r}  Location={location!r}")
    print(f"Members:     {len(members)}")
    print(f"Mode:        {'WRITE' if args.write else 'DRY-RUN'}")

    written = 0
    skipped = 0
    for path, seq in members:
        if path.resolve() == source.resolve():
            skipped += 1
            print(f"  skip source  {path.name}")
            continue
        merged = list(user_kws)
        if seq.get("Sequence"):
            merged.append(f"Sequence:{seq['Sequence']}")
        if seq.get("SequenceID"):
            merged.append(f"SequenceID:{seq['SequenceID']}")
        if seq.get("SequencePos"):
            merged.append(f"SequencePos:{seq['SequencePos']}")
        print(f"  {'write' if args.write else 'would'}  {path.name}  keep {seq.get('SequencePos', '-')}")
        if args.write:
            try:
                write_target(
                    exiftool,
                    path,
                    merged_keywords=merged,
                    description="" if description is None else str(description),
                    rating=None if rating is None else str(rating),
                    location="" if location is None else str(location),
                )
                written += 1
            except Exception as exc:  # noqa: BLE001
                print(f"ERROR {path.name}: {exc}", file=sys.stderr)
                return 1

    if not args.write:
        print("Dry-run only. Re-run with --write to apply.")
    else:
        print(f"Wrote: {written}  skipped source: {skipped}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
