#!/usr/bin/env python3
"""Tag photo sequences from capture-time proximity (hi / hi+ / Continuous).

Dry-run by default. Writes Keywords/Subject:
  Sequence:yes|no, SequenceID:YYYYMMDD-HHMMSS, SequencePos:i/n

Clusters shorter than --min-frames (default 4) are Sequence:no — filters out
accidental 2–3 frame taps in continuous mode.

Example:
  python3 scripts/tag_sequences.py /path/to/folder
  python3 scripts/tag_sequences.py /path/to/folder --write
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

PHOTO_EXTS = {".jpg", ".jpeg", ".tif", ".tiff", ".arw", ".cr2", ".cr3", ".nef", ".dng", ".heic"}
SEQ_KEYS = ("Sequence", "SequenceID", "SequencePos")
CONTINUOUS_RE = re.compile(r"(continuous|\bhi\+?\b)", re.I)


@dataclass
class Frame:
    path: Path
    captured: datetime
    release_mode: str = ""
    sequence_number: str = ""
    existing: dict[str, str] = field(default_factory=dict)


@dataclass
class Assignment:
    frame: Frame
    sequence: str  # yes / no
    sequence_id: str
    sequence_pos: str
    cluster_size: int


def find_exiftool() -> str:
    for candidate in ("/usr/local/bin/exiftool", "/opt/homebrew/bin/exiftool", "exiftool"):
        path = shutil.which(candidate) if candidate == "exiftool" else candidate
        if path and Path(path).exists():
            return path
    raise SystemExit("exiftool not found. Install with: brew install exiftool")


def parse_exif_dt(raw: str | None) -> datetime | None:
    if not raw:
        return None
    raw = raw.strip()
    if not raw or raw.startswith("0000"):
        return None
    for fmt in ("%Y:%m:%d %H:%M:%S.%f", "%Y:%m:%d %H:%M:%S"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def is_continuous_mode(release_mode: str) -> bool:
    return bool(CONTINUOUS_RE.search(release_mode or ""))


def parse_keyword_map(keywords) -> dict[str, str]:
    out: dict[str, str] = {}
    if keywords is None:
        return out
    if isinstance(keywords, str):
        items = [keywords]
    elif isinstance(keywords, list):
        items = keywords
    else:
        items = [str(keywords)]
    for item in items:
        text = str(item).strip()
        if ":" not in text:
            continue
        key, _, val = text.partition(":")
        key = key.strip()
        if key in SEQ_KEYS:
            out[key] = val.strip()
    return out


def collect_frames(root: Path, recursive: bool, exiftool: str) -> list[Frame]:
    paths: list[Path] = []
    if recursive:
        for p in root.rglob("*"):
            if p.is_file() and p.suffix.lower() in PHOTO_EXTS:
                paths.append(p)
    else:
        for p in root.iterdir():
            if p.is_file() and p.suffix.lower() in PHOTO_EXTS:
                paths.append(p)
    paths.sort()
    if not paths:
        return []

    # Batch metadata read
    # Do not use -n: ReleaseMode becomes a bare integer (e.g. 2) instead of "Continuous".
    cmd = [
        exiftool,
        "-json",
        "-DateTimeOriginal",
        "-CreateDate",
        "-ReleaseMode",
        "-ReleaseMode2",
        "-ReleaseMode3",
        "-SequenceNumber",
        "-SequenceImageNumber",
        "-Keywords",
        "-Subject",
        "--",
        *[str(p) for p in paths],
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode not in (0, 1):  # 1 = minor warnings
        raise SystemExit(f"exiftool failed:\n{proc.stderr or proc.stdout}")
    rows = json.loads(proc.stdout or "[]")

    frames: list[Frame] = []
    by_name = {Path(r.get("SourceFile", "")).resolve(): r for r in rows}
    for p in paths:
        row = by_name.get(p.resolve()) or {}
        captured = parse_exif_dt(row.get("DateTimeOriginal")) or parse_exif_dt(row.get("CreateDate"))
        if captured is None:
            captured = datetime.fromtimestamp(p.stat().st_mtime)
        release = " ".join(
            str(row.get(k) or "")
            for k in ("ReleaseMode", "ReleaseMode2", "ReleaseMode3")
        ).strip()
        seq_num = str(row.get("SequenceNumber") or row.get("SequenceImageNumber") or "")
        existing = parse_keyword_map(row.get("Keywords"))
        if not existing:
            existing = parse_keyword_map(row.get("Subject"))
        frames.append(
            Frame(
                path=p,
                captured=captured,
                release_mode=release,
                sequence_number=seq_num,
                existing=existing,
            )
        )
    return frames


def cluster_frames(frames: list[Frame], gap_seconds: float) -> list[list[Frame]]:
    if not frames:
        return []
    ordered = sorted(frames, key=lambda f: (f.captured, f.path.name))
    clusters: list[list[Frame]] = [[ordered[0]]]
    for frame in ordered[1:]:
        prev = clusters[-1][-1]
        delta = (frame.captured - prev.captured).total_seconds()
        if delta > gap_seconds:
            clusters.append([frame])
        else:
            clusters[-1].append(frame)
    return clusters


def assign_clusters(clusters: list[list[Frame]], min_frames: int = 4) -> list[Assignment]:
    """Mark clusters with >= min_frames as Sequence:yes; shorter ones as no."""
    out: list[Assignment] = []
    for cluster in clusters:
        n = len(cluster)
        if n < min_frames:
            for frame in cluster:
                out.append(
                    Assignment(
                        frame=frame,
                        sequence="no",
                        sequence_id="",
                        sequence_pos="",
                        cluster_size=n,
                    )
                )
            continue
        seq_id = cluster[0].captured.strftime("%Y%m%d-%H%M%S")
        for i, frame in enumerate(cluster, start=1):
            out.append(
                Assignment(
                    frame=frame,
                    sequence="yes",
                    sequence_id=seq_id,
                    sequence_pos=f"{i}/{n}",
                    cluster_size=n,
                )
            )
    return out


def merge_keywords(existing_keywords, assignment: Assignment) -> list[str]:
    """Replace Sequence* keys; keep other keywords."""
    kept: list[str] = []
    if existing_keywords is None:
        items = []
    elif isinstance(existing_keywords, str):
        items = [existing_keywords]
    elif isinstance(existing_keywords, list):
        items = existing_keywords
    else:
        items = [str(existing_keywords)]
    for item in items:
        text = str(item).strip()
        if not text:
            continue
        key = text.split(":", 1)[0].strip()
        if key in SEQ_KEYS:
            continue
        kept.append(text)
    kept.append(f"Sequence:{assignment.sequence}")
    if assignment.sequence_id:
        kept.append(f"SequenceID:{assignment.sequence_id}")
    if assignment.sequence_pos:
        kept.append(f"SequencePos:{assignment.sequence_pos}")
    return kept


def write_assignment(exiftool: str, assignment: Assignment, existing_keywords) -> None:
    kws = merge_keywords(existing_keywords, assignment)
    # MWG:Keywords syncs IPTC:Keywords + XMP-dc:Subject. Full-list replace
    # (not clear+=) — MWG ignores Keywords= when += appears in the same command.
    cmd = [
        exiftool,
        "-m",
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
        f"-MWG:Keywords={'||'.join(kws)}",
        "--",
        str(assignment.frame.path),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode not in (0, 1):
        raise RuntimeError(proc.stderr or proc.stdout or "exiftool write failed")


def already_tagged(frame: Frame, assignment: Assignment) -> bool:
    return (
        frame.existing.get("Sequence") == assignment.sequence
        and frame.existing.get("SequenceID", "") == assignment.sequence_id
        and frame.existing.get("SequencePos", "") == assignment.sequence_pos
    )


def print_summary(
    *,
    total_files: int,
    skipped_mode: int,
    assignments: list[Assignment],
    gap: float,
    min_frames: int,
    write: bool,
    force: bool,
) -> None:
    multi = [a for a in assignments if a.sequence == "yes"]
    short = [a for a in assignments if a.sequence == "no" and a.cluster_size > 1]
    singles = [a for a in assignments if a.cluster_size == 1]
    # Unique sequence IDs among accepted sequences
    seq_ids = sorted({a.sequence_id for a in multi if a.sequence_id})
    sizes = {}
    for a in multi:
        if a.sequence_pos.endswith(f"/{a.cluster_size}") and a.sequence_pos.startswith("1/"):
            sizes[a.sequence_id] = a.cluster_size

    print()
    print("=== Sequence tagging summary ===")
    print(f"Mode:           {'WRITE' if write else 'DRY-RUN'}" + (" (force)" if force and write else ""))
    print(f"Gap threshold:  {gap:g} s")
    print(f"Min frames:     {min_frames}")
    print(f"Files scanned:  {total_files}")
    print(f"Skipped (mode): {skipped_mode}  (not Continuous/Hi/Hi+)")
    print(f"Eligible:       {len(assignments)}")
    print(f"Sequences:      {len(seq_ids)}  (clusters with {min_frames}+ frames)")
    print(f"In sequences:   {len(multi)} frames")
    print(f"Short clusters: {len(short)} frames  (2–{min_frames - 1} → Sequence:no)")
    print(f"Singles:        {len(singles)}  (eligible Continuous but alone → Sequence:no)")
    if sizes:
        vals = sorted(sizes.values())
        print(f"Seq size min/med/max: {vals[0]} / {vals[len(vals)//2]} / {vals[-1]}")
        # Show up to 15 sequences
        print("Sequences:")
        for sid in seq_ids[:15]:
            print(f"  {sid}  ({sizes.get(sid, '?')} frames)")
        if len(seq_ids) > 15:
            print(f"  … {len(seq_ids) - 15} more")
    print("===============================")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Tag hi/hi+/Continuous photo sequences by capture-time gap.")
    parser.add_argument("directory", type=Path, help="Folder of photos")
    parser.add_argument("--gap", type=float, default=1.0, help="Max seconds between frames in a sequence (default 1.0)")
    parser.add_argument(
        "--min-frames",
        type=int,
        default=4,
        help="Minimum frames to count as a sequence (default 4; shorter → Sequence:no)",
    )
    parser.add_argument("--recursive", action="store_true", help="Scan subfolders")
    parser.add_argument("--write", action="store_true", help="Write tags (default is dry-run)")
    parser.add_argument("--force", action="store_true", help="Rewrite even if Sequence* tags already match")
    parser.add_argument(
        "--include-all",
        action="store_true",
        help="Do not gate on Continuous/Hi/Hi+ release mode",
    )
    args = parser.parse_args(argv)
    if args.min_frames < 2:
        raise SystemExit("--min-frames must be >= 2")

    root = args.directory.expanduser().resolve()
    if not root.is_dir():
        raise SystemExit(f"Not a directory: {root}")

    exiftool = find_exiftool()
    frames = collect_frames(root, args.recursive, exiftool)
    total = len(frames)

    if args.include_all:
        eligible = frames
        skipped = 0
    else:
        eligible = [f for f in frames if is_continuous_mode(f.release_mode)]
        skipped = total - len(eligible)

    clusters = cluster_frames(eligible, args.gap)
    assignments = assign_clusters(clusters, min_frames=args.min_frames)

    # For dry-run detail, show a few examples
    shown = 0
    for a in assignments:
        is_seq = a.sequence == "yes"
        if not is_seq and shown >= 5:
            continue
        if is_seq or shown < 5:
            mark = "yes" if is_seq else "no "
            print(
                f"{mark}  {a.sequence_pos or '-':>8}  {a.sequence_id or '-':16}  {a.frame.path.name}"
            )
            shown += 1
            if shown >= 40:
                print("…")
                break

    print_summary(
        total_files=total,
        skipped_mode=skipped,
        assignments=assignments,
        gap=args.gap,
        min_frames=args.min_frames,
        write=args.write,
        force=args.force,
    )

    if not args.write:
        print("Dry-run only. Re-run with --write to apply tags.")
        return 0

    # Need Keywords for merge — re-read per file is expensive; use existing from first pass
    # But existing only has Sequence* parsed. Re-fetch Keywords for writes.
    written = 0
    skipped_same = 0
    errors = 0
    for a in assignments:
        if not args.force and already_tagged(a.frame, a):
            skipped_same += 1
            continue
        # Fetch full keywords for merge
        proc = subprocess.run(
            [exiftool, "-json", "-Keywords", "-Subject", "--", str(a.frame.path)],
            capture_output=True,
            text=True,
        )
        kws = None
        if proc.returncode in (0, 1) and proc.stdout:
            row = json.loads(proc.stdout)[0]
            kws = row.get("Keywords")
            if kws is None:
                kws = row.get("Subject")
        try:
            write_assignment(exiftool, a, kws)
            written += 1
        except Exception as exc:  # noqa: BLE001
            errors += 1
            print(f"ERROR {a.frame.path.name}: {exc}", file=sys.stderr)

    print(f"Wrote: {written}  unchanged: {skipped_same}  errors: {errors}")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
