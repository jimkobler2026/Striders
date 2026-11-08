#!/usr/bin/env python3
"""Apply VideoEntry + Videos sheets to the live .xlsm via xlwings."""

import xlwings as xw

VIDEO_HEADERS = [
    "Name",
    "Path",
    "Date",
    "Time",
    "Type",
    "FileSize",
    "Duration",
    "FrameRate",
    "Resolution",
    "Camera",
    "Speed",
    "f-stop",
    "ISO",
    "Rating",
    "Species",
    "Behavior",
    "Locale",
    "Send to UK",
    "View",
    "Lighting",
    "notes",
]

VIDEO_ENTRY_FIELDS = [
    "Name",
    "Path",
    "Date",
    "Time",
    "Type",
    "FileSize",
    "Duration",
    "FrameRate",
    "Resolution",
    "Camera",
    "Speed",
    "f-stop",
    "ISO",
    "Rating",
    "Species",
    "Behavior",
    "Locale",
    "Send to UK",
    "View",
    "Lighting",
    "notes",
]
VIDEO_ENTRY_AUTO_COUNT = 13

LISTS = {
    "Rating": ["1", "2", "3", "4", "5"],
    "Species": ["Unknown", "Aquarius", "Gerris", "Other"],
    "Behavior": [
        "Skating",
        "Resting",
        "Mating",
        "Predation",
        "PreyCapture",
        "Jumping",
        "Escape",
        "Grooming",
        "Territorial",
        "Grouping",
        "Other",
        "Unknown",
    ],
    "Send to UK": ["Yes", "No", "Unknown"],
    "View": ["surface", "bottom", "both", "from below", "eye level"],
    "Lighting": ["Bright clear sun", "slightly overcast", "overcast", "very overcast", "godox strobe", "2 strobes", "mole spotlight"],
}

ENTRY_COL_WIDTHS = {
    "A": 9.83,
    "B": 28.33,
    "C": 9.5,
    "D": 8.83,
    "E": 9.5,
    "F": 11.33,
    "G": 12.0,
    "H": 8.33,
    "I": 6.83,
    "J": 10.0,
    "K": 10.17,
    "L": 9.33,
    "M": 5.67,
    "N": 9.5,
    "O": 27.67,
    "P": 27.67,
}

FONT = "Times New Roman"
SIZE = 14
DATE_FORMAT = "m/d/yy"
TIME_FORMAT = "[$-F400]h:mm:ss AM/PM"


def find_book():
    path_hint = "Strider_Media_Catalog.xlsm"
    for app in xw.apps:
        for b in app.books:
            try:
                name = str(b.fullname)
            except Exception:
                continue
            if path_hint in name and "/build/" not in name:
                return b
    raise SystemExit("Open catalog/Strider_Media_Catalog.xlsm in Excel first.")


def style_cell(cell, bold=False, italic=False):
    cell.font.name = FONT
    cell.font.size = SIZE
    cell.font.bold = bold
    cell.font.italic = italic
    cell.api.horizontal_alignment.set(left_text_alignment := None) if False else None
    try:
        cell.api.horizontal_alignment.set(xw.constants.HAlign.xlHAlignLeft)
    except Exception:
        try:
            cell.api.properties().horizontal_alignment.set(-4131)  # xlLeft
        except Exception:
            pass


def clear_sheet(ws):
    ws.clear()


def ensure_videos(wb):
    names = [s.name for s in wb.sheets]
    if "Videos" in names:
        ws = wb.sheets["Videos"]
        clear_sheet(ws)
    else:
        # place after Photos if present
        after = wb.sheets["Photos"] if "Photos" in names else wb.sheets[0]
        ws = wb.sheets.add("Videos", after=after)

    for i, h in enumerate(VIDEO_HEADERS, start=1):
        cell = ws.range((1, i))
        cell.value = h
        style_cell(cell, bold=True)

    try:
        ws.api.freeze_panes.set("$A$2")
    except Exception:
        pass
    try:
        ws.api.show_gridlines.set(True)
    except Exception:
        pass
    print("Videos headers:", VIDEO_HEADERS)
    return ws


def ensure_video_entry(wb):
    names = [s.name for s in wb.sheets]
    if "VideoEntry" in names:
        ws = wb.sheets["VideoEntry"]
        clear_sheet(ws)
    else:
        after = wb.sheets["Entry"] if "Entry" in names else wb.sheets[0]
        ws = wb.sheets.add("VideoEntry", after=after)

    ws.range("A1").value = "This sheet collects video data before adding to the Videos catalog"
    style_cell(ws.range("A1"))

    steps = [
        (3, "1) Select video in Finder.  "),
        (4, "2) MACRO: LoadVideoFromFinderSelection. "),
        (5, "3) Click an option in columns C+ (Tags section) to put it in column B.  "),
        (6, "4) MACRO: AppendVideoRow."),
    ]
    for row, text in steps:
        cell = ws.range((row, 1))
        cell.value = text
        style_cell(cell)

    ws.range((8, 1)).value = "Auto (from LoadVideoFromFinderSelection)"
    style_cell(ws.range((8, 1)), italic=True)

    row = 9
    for i, name in enumerate(VIDEO_ENTRY_FIELDS):
        if i == VIDEO_ENTRY_AUTO_COUNT:
            cell = ws.range((row, 1))
            cell.value = "Tags (click option in C+ → fills B)"
            style_cell(cell, italic=True)
            row += 1
        label = ws.range((row, 1))
        label.value = name
        style_cell(label, bold=True)
        value = ws.range((row, 2))
        value.value = None
        style_cell(value)
        if name == "Date":
            value.number_format = DATE_FORMAT
        elif name == "Time":
            value.number_format = TIME_FORMAT
        if name in LISTS:
            for j, opt in enumerate(LISTS[name]):
                cell = ws.range((row, 3 + j))
                cell.value = opt
                style_cell(cell)
        row += 1

    for letter, width in ENTRY_COL_WIDTHS.items():
        ws.range(f"{letter}:{letter}").column_width = width

    try:
        ws.api.show_gridlines.set(True)
    except Exception:
        pass

    print("VideoEntry fields through row", row - 1)
    return ws


def update_lists_view(wb):
    """Ensure Lists View options match surface/bottom/..."""
    if "Lists" not in [s.name for s in wb.sheets]:
        return
    lists = wb.sheets["Lists"]
    # Find View column in row 2
    view_col = None
    for c in range(1, 20):
        if str(lists.range((2, c)).value or "").strip() == "View":
            view_col = c
            break
    if view_col is None:
        return
    opts = LISTS["View"]
    for r in range(3, 20):
        lists.range((r, view_col)).value = None
    for i, opt in enumerate(opts):
        cell = lists.range((3 + i, view_col))
        cell.value = opt
        style_cell(cell)
    print("Lists View updated:", opts)


def update_readme(wb):
    if "ReadMe" not in [s.name for s in wb.sheets]:
        return
    ws = wb.sheets["ReadMe"]
    ws.range("A1").value = "Strider Media Catalog"
    style_cell(ws.range("A1"), bold=True)
    lines = [
        "",
        "Sheets",
        "Entry — add photos (LoadFromFinderSelection → AppendPhotoRow).",
        "VideoEntry — add videos (LoadVideoFromFinderSelection → AppendVideoRow).",
        "Photos / Videos — catalogs.",
        "",
        "Videos columns",
        ", ".join(VIDEO_HEADERS),
    ]
    for i, line in enumerate(lines, start=2):
        cell = ws.range((i, 1))
        cell.value = line
        style_cell(cell)


def reorder(wb):
    desired = ["Entry", "VideoEntry", "Photos", "Videos", "Lists", "ReadMe"]
    # Move sheets to match order as best-effort
    existing = [s.name for s in wb.sheets]
    for i, name in enumerate(desired):
        if name not in existing:
            continue
        try:
            wb.sheets[name].api.move(before=wb.sheets[desired[0]].api if i == 0 else None)
        except Exception:
            pass
    # Simpler: move each to end in order then... skip if fragile
    print("Sheet order now:", [s.name for s in wb.sheets])


def main():
    wb = find_book()
    ensure_videos(wb)
    ensure_video_entry(wb)
    update_lists_view(wb)
    update_readme(wb)
    reorder(wb)
    wb.save()
    print("Saved", wb.fullname)


if __name__ == "__main__":
    main()
