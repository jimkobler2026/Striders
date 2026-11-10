#!/usr/bin/env python3
"""Tag photo sequences from Continuous / Hi / Hi+ bursts.

Primary boundary (Sony and similar): maker-note SequenceNumber, which resets
to 1 at each new burst. Capture-time --gap is the fallback when SequenceNumber
is absent (phones, non-Sony). Within a burst, order follows the original camera
number embedded in the filename (…_DSC7659), not timestamp alone.

Dry-run by default. Writes Keywords/Subject via one persistent exiftool
(-stay_open) session:
  Sequence:yes|no, SequenceID:YYYYMMDD-HHMMSS, SequencePos:i/n

Bursts shorter than --min-frames (default 4) are Sequence:no — filters out
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

_SCRIPTS = Path(__file__).resolve().parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

from exif_datetime import gap_seconds as dt_gap_seconds  # noqa: E402
from exif_datetime import parse_exif_dt, stamp_local  # noqa: E402

PHOTO_EXTS = {".jpg", ".jpeg", ".tif", ".tiff", ".arw", ".cr2", ".cr3", ".nef", ".dng", ".heic"}
SEQ_KEYS = ("Sequence", "SequenceID", "SequencePos")
CONTINUOUS_RE = re.compile(r"(continuous|\bhi\+?\b)", re.I)
# Tail of rename_media_unique stem: YYYYMMDD-HHMMSS_DSC7659 → 7659
CAMERA_NUM_RE = re.compile(r"(\d+)$")


@dataclass
class Frame:
    path: Path
    captured: datetime
    release_mode: str = ""
    sequence_number: str = ""
    existing: dict[str, str] = field(default_factory=dict)
    # Full Keywords/Subject from collect pass (for MWG merge on write).
    keywords_raw: object | None = None


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


def is_continuous_mode(release_mode: str) -> bool:
    return bool(CONTINUOUS_RE.search(release_mode or ""))


def parse_sequence_number(raw: str | None) -> int | None:
    """Parse maker-note SequenceNumber; ignore non-numeric placeholders."""
    if raw is None:
        return None
    text = str(raw).strip()
    if not text or not text.isdigit():
        return None
    return int(text)


def camera_number(path: Path) -> int:
    """Original camera file number from basename (e.g. …_DSC7659 → 7659)."""
    stem = path.stem
    tail = stem.rsplit("_", 1)[-1] if "_" in stem else stem
    match = CAMERA_NUM_RE.search(tail)
    if match:
        return int(match.group(1))
    match = CAMERA_NUM_RE.search(stem)
    return int(match.group(1)) if match else -1


def frame_sort_key(frame: Frame) -> tuple:
    return (frame.captured, camera_number(frame.path), frame.path.name)


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
        "-OffsetTimeOriginal",
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
        captured = parse_exif_dt(
            row.get("DateTimeOriginal"), row.get("OffsetTimeOriginal")
        ) or parse_exif_dt(row.get("CreateDate"))
        if captured is None:
            captured = datetime.fromtimestamp(p.stat().st_mtime)
        release = " ".join(
            str(row.get(k) or "")
            for k in ("ReleaseMode", "ReleaseMode2", "ReleaseMode3")
        ).strip()
        seq_num = str(row.get("SequenceNumber") or row.get("SequenceImageNumber") or "")
        keywords_raw = row.get("Keywords")
        if keywords_raw is None:
            keywords_raw = row.get("Subject")
        existing = parse_keyword_map(keywords_raw)
        frames.append(
            Frame(
                path=p,
                captured=captured,
                release_mode=release,
                sequence_number=seq_num,
                existing=existing,
                keywords_raw=keywords_raw,
            )
        )
    return frames


def cluster_frames(frames: list[Frame], gap_seconds: float) -> list[list[Frame]]:
    """Group bursts: SequenceNumber reset when available, else capture-time gap."""
    if not frames:
        return []
    ordered = sorted(frames, key=frame_sort_key)
    clusters: list[list[Frame]] = [[ordered[0]]]
    for frame in ordered[1:]:
        prev = clusters[-1][-1]
        curr_sn = parse_sequence_number(frame.sequence_number)
        prev_sn = parse_sequence_number(prev.sequence_number)
        if curr_sn is not None and prev_sn is not None:
            # New burst when counter returns to 1 or otherwise drops.
            new_burst = curr_sn == 1 or curr_sn < prev_sn
        else:
            new_burst = dt_gap_seconds(prev.captured, frame.captured) > gap_seconds
        if new_burst:
            clusters.append([frame])
        else:
            clusters[-1].append(frame)
    return clusters


def assign_clusters(clusters: list[list[Frame]], min_frames: int = 4) -> list[Assignment]:
    """Mark bursts with >= min_frames as Sequence:yes; shorter ones as no."""
    out: list[Assignment] = []
    for cluster in clusters:
        # Filename camera number is authoritative within a burst.
        ordered = sorted(cluster, key=lambda f: (camera_number(f.path), f.path.name))
        n = len(ordered)
        if n < min_frames:
            for frame in ordered:
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
        seq_id = stamp_local(ordered[0].captured)
        for i, frame in enumerate(ordered, start=1):
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


def mwg_keyword_args(keywords: list[str], path: Path) -> list[str]:
    """Arg lines for one MWG full-list Keyword replace (stay_open / -@).

    Do not insert a bare ``--`` before the path: with ``-execute`` / stay_open,
    everything after ``--`` is treated as a filename (including ``-execute``),
    which hangs the session. Paths are absolute (or ./…) so option parsing is safe.
    """
    # MWG:Keywords syncs IPTC:Keywords + XMP-dc:Subject. Full-list replace
    # (not clear+=) — MWG ignores Keywords= when += appears in the same command.
    file_arg = str(path)
    if file_arg.startswith("-"):
        file_arg = f"./{file_arg}"
    return [
        "-m",
        "-P",  # preserve FileModifyDate / FileCreateDate
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
        f"-MWG:Keywords={'||'.join(keywords)}",
        file_arg,
    ]


class ExifToolStayOpen:
    """One long-lived exiftool process for many per-file writes."""

    def __init__(self, exiftool: str):
        self.exiftool = exiftool
        self.proc: subprocess.Popen[str] | None = None

    def __enter__(self) -> ExifToolStayOpen:
        # Merge stderr into stdout so the pipe cannot fill and deadlock while
        # we wait for `{ready}` on stdout alone.
        self.proc = subprocess.Popen(
            [self.exiftool, "-stay_open", "True", "-@", "-"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
        )
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def close(self) -> None:
        proc = self.proc
        self.proc = None
        if proc is None:
            return
        try:
            if proc.stdin and not proc.stdin.closed:
                proc.stdin.write("-stay_open\nFalse\n")
                proc.stdin.flush()
                proc.stdin.close()
            proc.wait(timeout=60)
        except Exception:  # noqa: BLE001
            try:
                proc.kill()
                proc.wait(timeout=5)
            except Exception:  # noqa: BLE001
                pass

    def _await_ready(self) -> str:
        assert self.proc and self.proc.stdout
        collected: list[str] = []
        while True:
            line = self.proc.stdout.readline()
            if line == "":
                raise RuntimeError(
                    "exiftool closed unexpectedly: " + ("".join(collected).strip() or "no output")
                )
            if line.strip() == "{ready}":
                break
            collected.append(line)
        return "".join(collected)

    def write_keywords(self, path: Path, keywords: list[str]) -> None:
        assert self.proc and self.proc.stdin
        args = mwg_keyword_args(keywords, path) + ["-execute"]
        self.proc.stdin.write("\n".join(args) + "\n")
        self.proc.stdin.flush()
        out = self._await_ready()
        if self.proc.poll() is not None:
            raise RuntimeError(out.strip() or "exiftool exited during write")
        # exiftool prints "Error:" on missing files etc. before {ready}.
        if re.search(r"(?i)^Error:", out, re.M):
            raise RuntimeError(out.strip() or "exiftool write failed")


def write_assignments(exiftool: str, pending: list[Assignment]) -> tuple[int, int]:
    """Write Sequence* tags via stay_open. Returns (written, errors)."""
    if not pending:
        return 0, 0
    written = 0
    errors = 0
    with ExifToolStayOpen(exiftool) as session:
        for assignment in pending:
            try:
                kws = merge_keywords(assignment.frame.keywords_raw, assignment)
                session.write_keywords(assignment.frame.path, kws)
                written += 1
            except Exception as exc:  # noqa: BLE001
                errors += 1
                print(f"ERROR {assignment.frame.path.name}: {exc}", file=sys.stderr)
    return written, errors


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
    print(f"Boundaries:     SequenceNumber reset (gap {gap:g}s fallback if absent)")
    print(f"Min frames:     {min_frames}")
    print(f"Files scanned:  {total_files}")
    print(f"Skipped (mode): {skipped_mode}  (not Continuous/Hi/Hi+)")
    print(f"Eligible:       {len(assignments)}")
    print(f"Sequences:      {len(seq_ids)}  (bursts with {min_frames}+ frames)")
    print(f"In sequences:   {len(multi)} frames")
    print(f"Short bursts:   {len(short)} frames  (2–{min_frames - 1} → Sequence:no)")
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
    parser = argparse.ArgumentParser(
        description="Tag Continuous/Hi/Hi+ bursts using SequenceNumber (gap fallback)."
    )
    parser.add_argument("directory", type=Path, help="Ingest folder of photos")
    parser.add_argument(
        "--gap",
        type=float,
        default=1.0,
        help="Max seconds between frames when SequenceNumber is absent (default 1.0)",
    )
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
            sn = a.frame.sequence_number or "-"
            print(
                f"{mark}  {a.sequence_pos or '-':>8}  {a.sequence_id or '-':16}  "
                f"camSeq={sn:>4}  {a.frame.path.name}"
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

    pending: list[Assignment] = []
    skipped_same = 0
    for a in assignments:
        if not args.force and already_tagged(a.frame, a):
            skipped_same += 1
            continue
        pending.append(a)

    written, errors = write_assignments(exiftool, pending)
    print(f"Wrote: {written}  unchanged: {skipped_same}  errors: {errors}")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
