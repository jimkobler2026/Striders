#!/usr/bin/env python3
"""Insert Sequence / SequenceID / SequencePos into an open Strider .xlsm via xlwings.

Does not rebuild the workbook. Open catalog/Strider_Media_Catalog.xlsm (or blank) first.
"""

from __future__ import annotations

import xlwings as xw

SEQ_FIELDS = ["Sequence", "SequenceID", "SequencePos"]
FONT = "Times New Roman"
SIZE = 14


def find_book():
    for app in xw.apps:
        for b in app.books:
            try:
                name = str(b.fullname)
            except Exception:
                continue
            if "Strider_Media_Catalog" in name and not name.lower().endswith(".xlsx"):
                return b
    raise SystemExit("Open catalog/Strider_Media_Catalog.xlsm (or blank) in Excel first.")


def style(cell, bold=False):
    cell.font.name = FONT
    cell.font.size = SIZE
    cell.font.bold = bold


def entry_labels(ws):
    labels = {}
    for r in range(1, 80):
        val = ws.range((r, 1)).value
        if val is None:
            continue
        text = str(val).strip()
        if text:
            labels[text] = r
    return labels


def ensure_entry_fields(ws):
    labels = entry_labels(ws)
    if all(f in labels for f in SEQ_FIELDS):
        print(f"{ws.name}: Sequence fields already present")
        return

    # Prefer insert after ISO; else after Time; else append before Rating
    anchor = labels.get("ISO") or labels.get("Time")
    rating_row = labels.get("Rating")
    if anchor is None:
        raise SystemExit(f"{ws.name}: could not find ISO/Time row to insert after")

    insert_at = anchor + 1
    # If fields somehow partially exist, skip
    missing = [f for f in SEQ_FIELDS if f not in labels]
    last = insert_at + len(missing) - 1
    # Mac Excel xlwings requires an explicit shift argument.
    ws.range(f"{insert_at}:{last}").insert("down")
    for i, name in enumerate(missing):
        cell = ws.range((insert_at + i, 1))
        cell.value = name
        style(cell, bold=True)
        val = ws.range((insert_at + i, 2))
        style(val)
        # SequencePos values like 6/54 must stay text, not Excel fractions.
        if name in SEQ_FIELDS:
            val.number_format = "@"
        try:
            ws.range((insert_at + i, 1), (insert_at + i, 2)).row_height = 18
        except Exception:
            pass
    print(f"{ws.name}: inserted {missing} at rows {insert_at}-{last}")
    if rating_row:
        print(f"  (Rating was at {rating_row}; rows shifted down)")


def ensure_photo_headers(ws):
    headers = []
    for c in range(1, 64):
        val = ws.range((1, c)).value
        if val is None or str(val).strip() == "":
            if c > 1 and not headers:
                continue
            if headers:
                break
            continue
        headers.append(str(val).strip())

    missing = [f for f in SEQ_FIELDS if f not in headers]
    if not missing:
        print("Photos: Sequence columns already present")
        return

    # Insert after Time if present, else append
    if "Time" in headers:
        col = headers.index("Time") + 2  # 1-based after Time
    else:
        col = len(headers) + 1

    for i, name in enumerate(missing):
        insert_col = col + i
        # Convert 1-based col index to Excel letter(s) for a full-column range.
        letter = ""
        n = insert_col
        while n:
            n, rem = divmod(n - 1, 26)
            letter = chr(65 + rem) + letter
        ws.range(f"{letter}:{letter}").insert("right")
        cell = ws.range((1, insert_col))
        cell.value = name
        style(cell, bold=True)
        # Entire data column as text so 6/54 is not parsed as a fraction.
        ws.range((2, insert_col), (5000, insert_col)).number_format = "@"
    print(f"Photos: inserted columns {missing}")


def main():
    wb = find_book()
    ensure_entry_fields(wb.sheets["Entry"])
    ensure_photo_headers(wb.sheets["Photos"])
    print("Done. Save the workbook (.xlsm). Re-import CatalogMacros.bas if not already updated.")


if __name__ == "__main__":
    main()
