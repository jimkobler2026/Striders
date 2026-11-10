#!/usr/bin/env python3
"""Prefix capture datetime onto camera filenames so catalog basenames stay unique.

Camera names like _DSC0320.JPG recycle after card format. Excel Append matches by
filename, so rename dump folders before Preview cull / catalog entry.

Allowed root: ~/Pictures/environmental caustics
New shape: YYYYMMDD-HHMMSS_ORIGINALSTEM.ext

Dry-run by default:
  python3 scripts/rename_media_unique.py "~/Pictures/environmental caustics/dump"
  python3 scripts/rename_media_unique.py "~/Pictures/environmental caustics/dump" --write
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

ALLOWED_ROOT = Path.home() / "Pictures" / "environmental caustics"

PHOTO_EXTS = {
    ".jpg",
    ".jpeg",
    ".tif",
    ".tiff",
    ".arw",
    ".cr2",
    ".cr3",
    ".nef",
    ".dng",
    ".heic",
}
VIDEO_EXTS = {".mov", ".mp4", ".m4v", ".avi"}
MEDIA_EXTS = PHOTO_EXTS | VIDEO_EXTS

ALREADY_RENAMED_RE = re.compile(r"^\d{8}-\d{6}_")


@dataclass
class PlanItem:
    path: Path
    new_name: str
    stamp_source: str
    action: str  # rename | skip | collide


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


def assert_under_allowed_root(folder: Path) -> Path:
    root = ALLOWED_ROOT.resolve()
    try:
        resolved = folder.resolve()
    except FileNotFoundError:
        raise SystemExit(f"Folder not found: {folder}")
    if not resolved.is_dir():
        raise SystemExit(f"Not a directory: {folder}")
    if resolved != root and root not in resolved.parents:
        raise SystemExit(
            f"Refusing to rename outside allowed root.\n"
            f"  folder: {resolved}\n"
            f"  allowed: {root}"
        )
    return resolved


def collect_paths(root: Path, recursive: bool) -> list[Path]:
    if recursive:
        paths = [p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in MEDIA_EXTS]
    else:
        paths = [p for p in root.iterdir() if p.is_file() and p.suffix.lower() in MEDIA_EXTS]
    paths.sort()
    return paths


def read_capture_times(exiftool: str, paths: list[Path]) -> dict[Path, tuple[datetime | None, str]]:
    """Return path -> (datetime or None, source label)."""
    if not paths:
        return {}
    cmd = [
        exiftool,
        "-json",
        "-DateTimeOriginal",
        "-CreateDate",
        "-MediaCreateDate",
        "--",
        *[str(p) for p in paths],
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode not in (0, 1):
        raise SystemExit(f"exiftool failed:\n{proc.stderr or proc.stdout}")
    rows = json.loads(proc.stdout or "[]")
    by_name = {Path(r.get("SourceFile", "")).resolve(): r for r in rows}

    out: dict[Path, tuple[datetime | None, str]] = {}
    for p in paths:
        row = by_name.get(p.resolve()) or {}
        for key in ("DateTimeOriginal", "CreateDate", "MediaCreateDate"):
            dt = parse_exif_dt(row.get(key))
            if dt is not None:
                out[p] = (dt, key)
                break
        else:
            out[p] = (None, "mtime")
    return out


def unique_target(directory: Path, base_name: str, reserved: set[str]) -> str:
    """Return base_name or base_name with -2, -3, … before the extension."""
    stem = Path(base_name).stem
    ext = Path(base_name).suffix
    candidate = base_name
    n = 2
    while candidate.lower() in reserved or (directory / candidate).exists():
        candidate = f"{stem}-{n}{ext}"
        n += 1
    return candidate


def build_plan(paths: list[Path], stamps: dict[Path, tuple[datetime | None, str]]) -> list[PlanItem]:
    items: list[PlanItem] = []
    reserved: set[str] = {p.name.lower() for p in paths}

    for path in paths:
        if ALREADY_RENAMED_RE.match(path.name):
            items.append(PlanItem(path, path.name, "-", "skip"))
            continue

        dt, source = stamps.get(path, (None, "mtime"))
        if dt is None:
            dt = datetime.fromtimestamp(path.stat().st_mtime)
            source = "mtime"

        stamp = dt.strftime("%Y%m%d-%H%M%S")
        # Nikon often uses _DSC… — avoid YYYYMMDD-HHMMSS__DSC…
        original = path.name[1:] if path.name.startswith("_") else path.name
        proposed = f"{stamp}_{original}"
        final = unique_target(path.parent, proposed, reserved - {path.name.lower()})
        reserved.add(final.lower())
        action = "collide" if final != proposed else "rename"
        items.append(PlanItem(path, final, source, action if final != path.name else "skip"))
    return items


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Rename media under ~/Pictures/environmental caustics with capture-datetime prefixes."
    )
    parser.add_argument("folder", help="Dump folder under ~/Pictures/environmental caustics")
    parser.add_argument("--write", action="store_true", help="Apply renames (default is dry-run)")
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Include nested subfolders (e.g. DCIM-style dumps)",
    )
    args = parser.parse_args()

    folder = assert_under_allowed_root(Path(args.folder).expanduser())
    paths = collect_paths(folder, recursive=args.recursive)
    if not paths:
        print(f"No media files found in {folder}")
        return 0

    exiftool = find_exiftool()
    stamps = read_capture_times(exiftool, paths)
    plan = build_plan(paths, stamps)

    rename_count = sum(1 for i in plan if i.action in ("rename", "collide") and i.new_name != i.path.name)
    skip_count = len(plan) - rename_count
    mtime_count = sum(1 for i in plan if i.stamp_source == "mtime" and i.new_name != i.path.name)

    print(f"Folder: {folder}")
    print(f"Files: {len(plan)}  rename: {rename_count}  skip: {skip_count}")
    if mtime_count:
        print(f"Warning: {mtime_count} file(s) used file mtime (no EXIF/media create date).")
    print()

    width = max((len(i.path.name) for i in plan), default=8)
    for item in plan:
        if item.new_name == item.path.name:
            print(f"  skip  {item.path.name}")
            continue
        tag = "ren" if item.action == "rename" else "col"
        print(f"  {tag}   {item.path.name:<{width}}  →  {item.new_name}  [{item.stamp_source}]")

    if not args.write:
        print()
        print("Dry-run only. Re-run with --write to apply renames.")
        return 0

    applied = 0
    errors = 0
    for item in plan:
        if item.new_name == item.path.name:
            continue
        dest = item.path.with_name(item.new_name)
        if dest.exists():
            print(f"ERROR: target exists, skipped: {dest}", file=sys.stderr)
            errors += 1
            continue
        try:
            item.path.rename(dest)
            applied += 1
        except OSError as exc:
            print(f"ERROR: {item.path.name} → {item.new_name}: {exc}", file=sys.stderr)
            errors += 1

    print()
    print(f"Renamed {applied} file(s)." + (f" Errors: {errors}." if errors else ""))
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
