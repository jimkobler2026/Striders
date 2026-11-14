#!/usr/bin/env python3
"""Detect sibling camera RAW + pixel size for catalog photos.

RAW siblings share the same basename in the same folder (Sony .arw, Canon .cr2/.cr3,
Nikon .nef/.nrw, Fuji .raf, plus .dng). HaveRaw? is Y if the file itself is RAW
or a sibling RAW exists.

Examples:
  python3 scripts/fill_raw_and_pixels.py ~/Pictures/shot.jpg
  python3 scripts/fill_raw_and_pixels.py --folder ~/Pictures/ingest --recursive
  python3 scripts/fill_raw_and_pixels.py --catalog   # fill open Photos sheet (xlwings)
  python3 scripts/fill_raw_and_pixels.py --catalog --write-keywords
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

RAW_EXTS = {".arw", ".cr2", ".cr3", ".nef", ".nrw", ".raf", ".dng"}


def find_exiftool() -> str:
    for candidate in ("/usr/local/bin/exiftool", "/opt/homebrew/bin/exiftool", "exiftool"):
        path = shutil.which(candidate) if candidate == "exiftool" else candidate
        if path and Path(path).exists():
            return path
    raise SystemExit("exiftool not found. Install with: brew install exiftool")


def have_raw(path: Path) -> str:
    if path.suffix.lower() in RAW_EXTS:
        return "Y"
    stem = path.stem
    parent = path.parent
    for ext in sorted(RAW_EXTS | {e.upper() for e in RAW_EXTS}):
        if (parent / f"{stem}{ext}").is_file():
            return "Y"
    return "N"


def read_pixels(exiftool: str, path: Path) -> tuple[str, str]:
    proc = subprocess.run(
        [exiftool, "-m", "-s3", "-ImageHeight", "-ImageWidth", "--", str(path)],
        capture_output=True,
        text=True,
    )
    lines = [ln.strip() for ln in (proc.stdout or "").splitlines() if ln.strip()]
    vert = lines[0] if len(lines) > 0 else ""
    hor = lines[1] if len(lines) > 1 else ""
    return vert, hor


def iter_images(root: Path, recursive: bool) -> list[Path]:
    patterns = ("*.jpg", "*.JPG", "*.jpeg", "*.JPEG", "*.heic", "*.HEIC", "*.tif", "*.TIF", "*.tiff", "*.TIFF")
    found: list[Path] = []
    for pat in patterns:
        found.extend(root.rglob(pat) if recursive else root.glob(pat))
    # De-dupe while preserving order
    seen: set[Path] = set()
    out: list[Path] = []
    for p in sorted(found):
        rp = p.resolve()
        if rp not in seen:
            seen.add(rp)
            out.append(p)
    return out


def write_keywords(exiftool: str, path: Path, raw_flag: str, vert: str, hor: str) -> None:
    """Merge rawYes/rawNo + VertPix/HorPix into MWG:Keywords (preserve other keywords)."""
    proc = subprocess.run(
        [exiftool, "-m", "-s3", "-MWG:Keywords", "--", str(path)],
        capture_output=True,
        text=True,
    )
    existing = []
    for line in (proc.stdout or "").splitlines():
        for part in line.split(","):
            t = part.strip()
            if t:
                existing.append(t)

    kept = []
    for t in existing:
        low = t.lower()
        if low in ("rawyes", "rawno"):
            continue
        if low.startswith("vertpix:") or low.startswith("horpix:"):
            continue
        kept.append(t)

    if raw_flag.upper() in ("Y", "YES"):
        kept.append("rawYes")
    else:
        kept.append("rawNo")
    if vert:
        kept.append(f"VertPix:{vert}")
    if hor:
        kept.append(f"HorPix:{hor}")

    joined = "||".join(kept)
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
        f"-MWG:Keywords={joined}",
        "--",
        str(path),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr or proc.stdout or "exiftool write failed")


def fill_catalog(write_keywords_flag: bool) -> None:
    try:
        import xlwings as xw
    except ImportError as e:
        raise SystemExit("xlwings required for --catalog. pip install xlwings") from e

    books = []
    for app in xw.apps:
        for b in app.books:
            try:
                name = str(b.fullname)
            except Exception:
                continue
            if "Strider_Media_Catalog" in name and not name.lower().endswith(".xlsx"):
                books.append(b)
    if not books:
        raise SystemExit("Open a Strider .xlsm catalog in Excel first.")
    book = None
    for b in books:
        if "blank" not in Path(str(b.fullname)).name.lower():
            book = b
            break
    if book is None:
        book = books[0]

    ws = book.sheets["Photos"]
    headers = {}
    for c in range(1, 80):
        val = ws.range((1, c)).value
        if val is None:
            continue
        headers[str(val).strip()] = c

    for need in ("Path", "HaveRaw?", "VertPix", "HorPix"):
        if need not in headers:
            raise SystemExit(
                f"Photos is missing column {need!r}. "
                "In Excel run macro EnsureRawAndPixelFields first."
            )

    path_col = headers["Path"]
    raw_col = headers["HaveRaw?"]
    vert_col = headers["VertPix"]
    hor_col = headers["HorPix"]

    exiftool = find_exiftool()
    last = ws.used_range.last_cell.row
    ok = 0
    missing = 0
    for r in range(2, last + 1):
        path_val = ws.range((r, path_col)).value
        if not path_val:
            continue
        path = Path(str(path_val).strip())
        if not path.is_file():
            missing += 1
            continue
        flag = have_raw(path)
        vert, hor = read_pixels(exiftool, path)
        ws.range((r, raw_col)).value = flag
        ws.range((r, vert_col)).value = vert or None
        ws.range((r, hor_col)).value = hor or None
        if write_keywords_flag:
            write_keywords(exiftool, path, flag, vert, hor)
        ok += 1
        print(f"{flag}\t{vert}x{hor}\t{path.name}")

    print(f"Updated {ok} Photos row(s); missing files: {missing}", file=sys.stderr)


def main() -> None:
    parser = argparse.ArgumentParser(description="Fill HaveRaw? / VertPix / HorPix for catalog photos.")
    parser.add_argument("paths", nargs="*", type=Path, help="Image file(s) to inspect")
    parser.add_argument("--folder", type=Path, help="Scan a folder of JPEGs")
    parser.add_argument("--recursive", action="store_true", help="With --folder, recurse")
    parser.add_argument("--catalog", action="store_true", help="Fill open Photos sheet via xlwings")
    parser.add_argument(
        "--write-keywords",
        action="store_true",
        help="Also write rawYes/rawNo + VertPix:/HorPix: into MWG:Keywords",
    )
    args = parser.parse_args()

    if args.catalog:
        fill_catalog(args.write_keywords)
        return

    files: list[Path] = []
    if args.folder:
        if not args.folder.is_dir():
            raise SystemExit(f"Not a folder: {args.folder}")
        files.extend(iter_images(args.folder, args.recursive))
    files.extend(p for p in args.paths if p.is_file())

    if not files:
        parser.error("Pass image paths, --folder, or --catalog")

    exiftool = find_exiftool()
    for path in files:
        flag = have_raw(path)
        vert, hor = read_pixels(exiftool, path)
        print(f"{flag}\t{vert}\t{hor}\t{path}")
        if args.write_keywords:
            write_keywords(exiftool, path, flag, vert, hor)


if __name__ == "__main__":
    main()
