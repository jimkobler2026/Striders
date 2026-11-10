#!/usr/bin/env python3
"""Add a click-pick catalog tag to the open Strider workbook (safe / additive).

Inserts Lists column, Entry + VideoEntry rows, Photos + Videos columns.
Does not clear existing catalog data. Patches scripts/build_catalog.py LISTS
and field arrays so skeleton rebuilds stay aligned.

Requires the working .xlsm open in Excel (xlwings).

Example:
  python3 scripts/add_catalog_field.py --name "Substrate" --options "mud,sand,rock"
  python3 scripts/add_catalog_field.py --name "Substrate" --options "mud,sand" --after Locale --dry-run
"""

from __future__ import annotations

import argparse
import ast
import re
import sys
from pathlib import Path

import xlwings as xw

ROOT = Path(__file__).resolve().parents[1]
BUILD_CATALOG = ROOT / "scripts" / "build_catalog.py"
BLANK_XLSM = ROOT / "catalog" / "Strider_Media_Catalog_blank.xlsm"

FONT = "Times New Roman"
SIZE = 14
LISTS_HEADER_ROW = 2
LISTS_FIRST_OPTION_ROW = 3


def find_books() -> list:
    books = []
    for app in xw.apps:
        for b in app.books:
            try:
                name = str(b.fullname)
            except Exception:
                continue
            if "Strider_Media_Catalog" in name and "/build/" not in name:
                books.append(b)
    return books


def find_working_book():
    books = find_books()
    if not books:
        raise SystemExit(
            "Open catalog/Strider_Media_Catalog.xlsm in Excel first "
            "(xlwings needs the workbook open)."
        )
    # Prefer non-blank working copy
    for b in books:
        full = str(b.fullname)
        if "blank" not in Path(full).name.lower():
            return b
    return books[0]


def find_blank_book():
    for b in find_books():
        if "blank" in Path(str(b.fullname)).name.lower():
            return b
    return None


def style(cell, bold=False):
    cell.font.name = FONT
    cell.font.size = SIZE
    cell.font.bold = bold


def col_letter(col: int) -> str:
    letter = ""
    n = col
    while n:
        n, rem = divmod(n - 1, 26)
        letter = chr(65 + rem) + letter
    return letter


def entry_labels(ws) -> dict[str, int]:
    labels: dict[str, int] = {}
    for r in range(1, 120):
        val = ws.range((r, 1)).value
        if val is None:
            continue
        text = str(val).strip()
        if text:
            labels[text] = r
    return labels


def catalog_headers(ws, header_row: int = 1) -> list[str]:
    headers: list[str] = []
    for c in range(1, 80):
        val = ws.range((header_row, c)).value
        if val is None or str(val).strip() == "":
            if headers:
                break
            continue
        headers.append(str(val).strip())
    return headers


def lists_headers(ws) -> list[str]:
    headers: list[str] = []
    for c in range(1, 40):
        val = ws.range((LISTS_HEADER_ROW, c)).value
        if val is None or str(val).strip() == "":
            if headers:
                break
            continue
        headers.append(str(val).strip())
    return headers


def parse_options(raw: str) -> list[str]:
    parts = [p.strip() for p in raw.split(",")]
    return [p for p in parts if p]


def plan_insert_row(labels: dict[str, int], name: str, after: str | None) -> int:
    if name in labels:
        raise SystemExit(f"Field already exists on form: {name}")
    if after:
        if after not in labels:
            raise SystemExit(f"--after field not found on form: {after}")
        return labels[after] + 1
    if "notes" in labels:
        return labels["notes"]  # insert above notes
    return max(labels.values()) + 1 if labels else 9


def plan_insert_col(headers: list[str], name: str, after: str | None) -> int:
    if name in headers:
        raise SystemExit(f"Column already exists: {name}")
    if after and after in headers:
        return headers.index(after) + 2  # 1-based after
    if "notes" in headers:
        return headers.index("notes") + 1  # before notes
    return len(headers) + 1


def add_lists_column(ws, name: str, options: list[str], dry_run: bool) -> None:
    headers = lists_headers(ws)
    if name in headers:
        raise SystemExit(f"Lists already has column: {name}")
    col = len(headers) + 1
    print(f"  Lists: add column {col_letter(col)} ({name}) with {len(options)} option(s)")
    if dry_run:
        return
    cell = ws.range((LISTS_HEADER_ROW, col))
    cell.value = name
    style(cell, bold=True)
    for i, opt in enumerate(options):
        c = ws.range((LISTS_FIRST_OPTION_ROW + i, col))
        c.value = opt
        style(c)
    try:
        ws.range(f"{col_letter(col)}:{col_letter(col)}").column_width = 12
    except Exception:
        pass


def keyword_key_for_label(label: str) -> str:
    aliases = {
        "Send to UK": "SendToUK",
        "Caustic camera aim": "CausticAim",
        "Caustic appearance": "CausticAppearance",
        "notes": "notes",
        "Note": "notes",
    }
    if label in aliases:
        return aliases[label]
    parts = label.replace("-", " ").replace("_", " ").split()
    out = []
    for w in parts:
        cleaned = "".join(ch for ch in w if ch.isalnum())
        if cleaned:
            out.append(cleaned[0].upper() + cleaned[1:])
    return "".join(out) or label


def add_entry_row(ws, name: str, options: list[str], after: str | None, dry_run: bool) -> None:
    labels = entry_labels(ws)
    insert_at = plan_insert_row(labels, name, after)
    print(f"  {ws.name}: insert row {insert_at} ({name})")
    if dry_run:
        return
    ws.range(f"{insert_at}:{insert_at}").insert("down")
    label_cell = ws.range((insert_at, 1))
    label_cell.value = name
    style(label_cell, bold=True)
    style(ws.range((insert_at, 2)))
    for i, opt in enumerate(options):
        c = ws.range((insert_at, 3 + i))
        c.value = opt
        style(c)
    try:
        ws.range((insert_at, 1), (insert_at, 2)).row_height = 18
    except Exception:
        pass


def add_metadata_preview_row(ws, name: str, dry_run: bool) -> None:
    """Add Keyword preview under 'Metadata written to file on Append'."""
    key = keyword_key_for_label(name)
    label = f"Keyword {key}:"
    section_row = None
    insert_at = None
    for r in range(1, 200):
        a = ws.range((r, 1)).value
        if a is None:
            continue
        text = str(a).strip()
        if text == label:
            print(f"  {ws.name}: metadata preview already has {label}")
            return
        if text == "Metadata written to file on Append":
            section_row = r
        if section_row and r > section_row:
            if text in ("Location → Locale", "Description → notes"):
                insert_at = r
                break
            if text.startswith("Keyword "):
                insert_at = r + 1
    if section_row is None:
        print(f"  {ws.name}: no 'Metadata written…' section — skip preview row")
        return
    if insert_at is None:
        insert_at = section_row + 1
    print(f"  {ws.name}: insert metadata preview row {insert_at} ({label})")
    if dry_run:
        return
    ws.range(f"{insert_at}:{insert_at}").insert("down")
    labels = entry_labels(ws)
    field_row = labels.get(name)
    if not field_row:
        print(f"  {ws.name}: warning — could not find field row for formula")
        return
    cell_a = ws.range((insert_at, 1))
    cell_a.value = label
    style(cell_a, bold=True)
    formula = f'=IF($B${field_row}="","(none)","{key}:"&$B${field_row})'
    cell_b = ws.range((insert_at, 2))
    cell_b.formula = formula
    style(cell_b)


def add_catalog_column(ws, name: str, after: str | None, dry_run: bool) -> None:
    headers = catalog_headers(ws)
    col = plan_insert_col(headers, name, after)
    print(f"  {ws.name}: insert column {col_letter(col)} ({name})")
    if dry_run:
        return
    letter = col_letter(col)
    ws.range(f"{letter}:{letter}").insert("right")
    cell = ws.range((1, col))
    cell.value = name
    style(cell, bold=True)


def apply_to_workbook(wb, name: str, options: list[str], after: str | None, dry_run: bool) -> None:
    print(f"Workbook: {wb.fullname}")
    names = [s.name for s in wb.sheets]
    required = ["Entry", "Photos", "Lists"]
    missing = [n for n in required if n not in names]
    if missing:
        raise SystemExit(f"Workbook missing sheets: {missing}")

    add_lists_column(wb.sheets["Lists"], name, options, dry_run)
    add_entry_row(wb.sheets["Entry"], name, options, after, dry_run)
    add_metadata_preview_row(wb.sheets["Entry"], name, dry_run)
    if "VideoEntry" in names:
        add_entry_row(wb.sheets["VideoEntry"], name, options, after, dry_run)
        add_metadata_preview_row(wb.sheets["VideoEntry"], name, dry_run)
    add_catalog_column(wb.sheets["Photos"], name, after, dry_run)
    if "Videos" in names:
        add_catalog_column(wb.sheets["Videos"], name, after, dry_run)


def _insert_into_string_list(src: str, list_name: str, item: str, before: str = "notes") -> str:
    """Insert quoted item into a Python list assigned to list_name, before `before` if present."""
    pattern = rf"({list_name}\s*=\s*\[)(.*?)(\])"
    m = re.search(pattern, src, flags=re.S)
    if not m:
        raise SystemExit(f"Could not find {list_name} in build_catalog.py")
    body = m.group(2)
    if re.search(rf'["\']{re.escape(item)}["\']', body):
        return src  # already present
    before_re = re.compile(rf'^([ \t]*)(["\']{re.escape(before)}["\'])', re.M)
    bm = before_re.search(body)
    if bm:
        indent = bm.group(1)
        insertion = f'{indent}"{item}",\n'
        new_body = body[: bm.start()] + insertion + body[bm.start() :]
    else:
        insertion = f'    "{item}",\n'
        new_body = body.rstrip() + "\n" + insertion
    return src[: m.start(2)] + new_body + src[m.end(2) :]


def _insert_into_lists_dict(src: str, name: str, options: list[str]) -> str:
    if re.search(rf'["\']{re.escape(name)}["\']\s*:', src):
        return src  # already present
    m = re.search(r"(LISTS\s*=\s*\{)(.*?)(\n\})", src, flags=re.S)
    if not m:
        raise SystemExit("Could not find LISTS dict in build_catalog.py")
    body = m.group(2)
    opts_literal = ", ".join(f'"{o}"' for o in options)
    # Keep Filters as last known key if present; insert before closing
    entry = f'    "{name}": [{opts_literal}],\n'
    # Prefer insert before final entry isn't needed — append before closing brace
    new_body = body.rstrip() + "\n" + entry
    return src[: m.start(2)] + new_body + src[m.end(2) :]


def patch_build_catalog(name: str, options: list[str], dry_run: bool) -> None:
    src = BUILD_CATALOG.read_text(encoding="utf-8")
    new_src = _insert_into_lists_dict(src, name, options)
    new_src = _insert_into_string_list(new_src, "ENTRY_FIELDS", name, before="notes")
    new_src = _insert_into_string_list(new_src, "VIDEO_ENTRY_FIELDS", name, before="notes")
    if new_src == src:
        print("  build_catalog.py: already up to date")
        return
    # Validate still parses
    try:
        ast.parse(new_src)
    except SyntaxError as exc:
        raise SystemExit(f"Refusing to write build_catalog.py — would break syntax: {exc}") from exc
    print(f"  build_catalog.py: add {name!r} to LISTS + ENTRY_FIELDS + VIDEO_ENTRY_FIELDS")
    if not dry_run:
        BUILD_CATALOG.write_text(new_src, encoding="utf-8")


def sync_blank(
    working,
    name: str,
    options: list[str],
    after: str | None,
    dry_run: bool,
) -> None:
    blank = find_blank_book()
    if blank is not None:
        if blank == working:
            print("Working book is the blank template — already updated.")
            return
        print("Blank workbook is open — applying same inserts.")
        apply_to_workbook(blank, name, options, after, dry_run)
        if not dry_run:
            print("  Save the blank .xlsm as well.")
        return

    if dry_run:
        print(f"Blank not open — would open/save {BLANK_XLSM.name} to sync.")
        return

    if not BLANK_XLSM.exists():
        print(f"Warning: blank template not found at {BLANK_XLSM}")
        print("  Open catalog/Strider_Media_Catalog_blank.xlsm later and re-run, or copy layout by hand.")
        return

    print(f"Opening blank template to sync: {BLANK_XLSM}")
    app = xw.App(visible=False)
    try:
        wb = app.books.open(str(BLANK_XLSM))
        apply_to_workbook(wb, name, options, after, dry_run=False)
        wb.save()
        wb.close()
        print("  Blank template saved.")
    except Exception as exc:  # noqa: BLE001
        print(f"Warning: could not sync blank template automatically: {exc}")
        print("  Open catalog/Strider_Media_Catalog_blank.xlsm and re-run this command, or copy layout.")
    finally:
        app.quit()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Add a click-pick tag field to the open Strider catalog (additive, safe)."
    )
    parser.add_argument("--name", required=True, help='Field label, e.g. "Substrate"')
    parser.add_argument(
        "--options",
        required=True,
        help='Comma-separated click-pick options, e.g. "mud,sand,rock"',
    )
    parser.add_argument(
        "--after",
        default=None,
        help="Insert after this existing field (default: before notes)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print planned changes without writing",
    )
    args = parser.parse_args(argv)

    name = args.name.strip()
    if not name:
        raise SystemExit("--name must be non-empty")
    options = parse_options(args.options)
    if not options:
        raise SystemExit("--options must include at least one value")

    print(f"{'DRY-RUN' if args.dry_run else 'APPLY'}: add field {name!r}")
    print(f"  options: {options}")
    if args.after:
        print(f"  after: {args.after}")
    else:
        print("  after: (default — before notes)")

    wb = find_working_book()
    apply_to_workbook(wb, name, options, args.after, args.dry_run)
    patch_build_catalog(name, options, args.dry_run)
    sync_blank(wb, name, options, args.after, args.dry_run)

    print()
    if args.dry_run:
        print("Dry-run only. Re-run without --dry-run to apply.")
    else:
        print("Done. Save the working .xlsm in Excel.")
        print(
            "No macro re-import needed for new fields "
            "(after the one-time dynamic-macro upgrade: re-import CatalogMacros.bas "
            "+ ./scripts/install_mac_helpers.sh + restart Excel)."
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
