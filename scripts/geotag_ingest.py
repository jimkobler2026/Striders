#!/usr/bin/env python3
"""Geotag photos in an ingest folder from a GPX track (exiftool -geotag).

Destructive to GPS fields. Dry-run by default. Prefer a /tmp copy first.

Example:
  python3 scripts/geotag_ingest.py /path/to/stills --gpx track.gpx
  python3 scripts/geotag_ingest.py /path/to/stills --gpx track.gpx --write
  python3 scripts/geotag_ingest.py /path/to/stills --gpx track.gpx --write --geosync +65
"""

from __future__ import annotations

import argparse
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


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Geotag ingest-folder photos from a GPX track (dry-run default)."
    )
    parser.add_argument("directory", type=Path, help="Ingest folder of photos")
    parser.add_argument("--gpx", type=Path, required=True, help="GPX track file")
    parser.add_argument("--recursive", action="store_true")
    parser.add_argument("--write", action="store_true", help="Apply GPS (default dry-run)")
    parser.add_argument(
        "--geosync",
        default="",
        help="Time sync offset for exiftool -geosync (e.g. +65 or 00:01:05)",
    )
    args = parser.parse_args(argv)

    root = args.directory.expanduser().resolve()
    gpx = args.gpx.expanduser().resolve()
    if not root.is_dir():
        raise SystemExit(f"Not a directory: {root}")
    if not gpx.is_file():
        raise SystemExit(f"GPX not found: {gpx}")

    if args.recursive:
        paths = [p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in PHOTO_EXTS]
    else:
        paths = [p for p in root.iterdir() if p.is_file() and p.suffix.lower() in PHOTO_EXTS]
    paths.sort()
    if not paths:
        raise SystemExit("No photos found.")

    exiftool = find_exiftool()
    cmd = [exiftool, "-m", "-P"]
    if not args.write:
        cmd.append("-n")  # still need a no-write dry path — use -v2 without overwrite
    if args.write:
        cmd.append("-overwrite_original")
    else:
        # Dry-run: print what would be written without modifying files
        cmd.extend(["-if", "$GPSLatitude", "-p", "already has GPS: $FileName"])
        print(f"DRY-RUN: would geotag {len(paths)} file(s) from {gpx.name}")
        print(f"  geosync={args.geosync or '(none)'}")
        print("Re-run with --write to apply. Tip: test on a /tmp copy first.")
        # Show a sample geotag command for transparency
        sample = [
            exiftool,
            "-m",
            "-P",
            "-overwrite_original",
            f"-geotag={gpx}",
        ]
        if args.geosync:
            sample.append(f"-geosync={args.geosync}")
        sample.extend(["--", str(paths[0])])
        print("Sample command:")
        print(" ", " ".join(sample))
        return 0

    cmd.append(f"-geotag={gpx}")
    if args.geosync:
        cmd.append(f"-geosync={args.geosync}")
    cmd.append("--")
    cmd.extend(str(p) for p in paths)
    print(f"WRITE: geotagging {len(paths)} file(s) from {gpx.name} …")
    proc = subprocess.run(cmd, capture_output=True, text=True)
    sys.stdout.write(proc.stdout or "")
    if proc.stderr:
        sys.stderr.write(proc.stderr)
    if proc.returncode not in (0, 1):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
