#!/usr/bin/env python3
"""Build the Strider media catalog workbook (.xlsx structure)."""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "catalog" / "build" / "Strider_Media_Catalog.xlsx"

PHOTO_HEADERS = [
    "Name",
    "Path",
    "Date",
    "Time",
    "Rating",
    "Species",
    "Behavior",
    "Locale",
    "Type",
    "FileSize",
    "Send to UK",
    "View",
    "Camera",
    "Mode",
    "Lens",
    "Speed",
    "f-stop",
    "ISO",
    "Lighting",
    "notes",
]

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

# Captured from live Entry sheet tag options (2026-11-07).
LISTS = {
    "Rating": ["1", "2", "3", "4", "5"],
    "Species": ["Unknown", "Aquarius", "Gerris", "Moth", "Ant"],
    "Behavior": [
        "stationary",
        "slow row",
        "medium row",
        "fast row",
        "steering",
        "jump",
        "pair interaction",
        "territorial",
        "predation",
        "larger group",
    ],
    "Locale": ["Purgatory Brook", "Attic aquarium"],
    "Send to UK": ["Send", "Sent", "maybe?"],
    "View": ["surface", "bottom", "both", "from below", "eye level"],
    "Lighting": [
        "Bright clear sun",
        "slightly overcast",
        "overcast",
        "very overcast",
        "godox strobe",
        "2 strobes",
        "mole spotlight",
    ],
}

# Photo Entry: auto first, then tags
ENTRY_FIELDS = [
    "Name",
    "Path",
    "Date",
    "Time",
    "Type",
    "FileSize",
    "Camera",
    "Mode",
    "Lens",
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
ENTRY_AUTO_COUNT = 12

# Video Entry: auto first, then tags
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

# Captured from live workbook column-width tuning (2026-11-07).
ENTRY_COL_WIDTHS = {
    "A": 9.83,
    "B": 23.67,
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
VIDEO_ENTRY_COL_WIDTHS = {
    "A": 12.67,
    "B": 21.67,
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
CATALOG_COL_WIDTH = 12.0
LISTS_COL_WIDTH = 8.5
DATE_FORMAT = "m/d/yy"
# Matches Excel's built-in time style as set on Photos/Videos (locale-aware).
TIME_FORMAT = "[$-F400]h:mm:ss AM/PM"

HEADER_FONT = Font(bold=True, name="Times New Roman", size=14)
LABEL_FONT = Font(bold=True, name="Times New Roman", size=14)
TITLE_FONT = Font(bold=True, name="Times New Roman", size=14)
HINT_FONT = Font(name="Times New Roman", size=14, italic=True)
PLAIN_FONT = Font(name="Times New Roman", size=14)


def style_input(cell):
    cell.font = PLAIN_FONT
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)


def list_range(name: str) -> str:
    col_idx = list(LISTS.keys()).index(name) + 1
    letter = get_column_letter(col_idx)
    return f"Lists!${letter}$3:${letter}$52"


def add_form_sheet(
    wb: Workbook,
    sheet_name: str,
    intro: str,
    steps: list[tuple[int, str]],
    fields: list[str],
    auto_count: int,
    auto_label: str,
    col_widths: dict[str, float] | None = None,
):
    ws = wb.create_sheet(sheet_name, 0)
    ws.sheet_view.showGridLines = True
    ws.sheet_format.defaultRowHeight = 18

    ws["A1"] = intro
    ws["A1"].font = PLAIN_FONT
    ws["A1"].alignment = Alignment(horizontal="left")

    for row, text in steps:
        cell = ws.cell(row, 1, text)
        cell.font = HINT_FONT if row == 7 else PLAIN_FONT
        cell.alignment = Alignment(horizontal="left")

    ws["A8"] = auto_label
    ws["A8"].font = HINT_FONT
    ws["A8"].alignment = Alignment(horizontal="left")

    field_rows: dict[str, int] = {}
    row = 9
    for i, name in enumerate(fields):
        if i == auto_count:
            section = ws.cell(row, 1, "Tags (click option in C+ → fills B)")
            section.font = HINT_FONT
            section.alignment = Alignment(horizontal="left")
            row += 1
        label = ws.cell(row, 1, name)
        label.font = LABEL_FONT
        label.alignment = Alignment(horizontal="left")
        value_cell = ws.cell(row, 2)
        style_input(value_cell)
        field_rows[name] = row
        if name == "Date":
            value_cell.number_format = DATE_FORMAT
        elif name == "Time":
            value_cell.number_format = TIME_FORMAT
        if name in LISTS:
            for j, opt in enumerate(LISTS[name]):
                cell = ws.cell(row, 3 + j, opt)
                cell.font = PLAIN_FONT
                cell.alignment = Alignment(horizontal="left")
        row += 1

    add_tag_write_preview(ws, field_rows, start_row=32)

    widths = col_widths or ENTRY_COL_WIDTHS
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = width

    return ws


def add_tag_write_preview(ws, field_rows: dict[str, int], start_row: int = 32):
    """Show live preview of metadata written to the media file on Append."""

    def bref(name: str) -> str:
        return f"$B${field_rows[name]}"

    def put(row: int, col: int, value, font=PLAIN_FONT):
        cell = ws.cell(row, col, value)
        cell.font = font
        cell.alignment = Alignment(horizontal="left", wrap_text=True)
        return cell

    r = start_row
    put(r, 1, "Metadata written to file on Append", HINT_FONT)
    r += 1

    preview_rows = [
        ("Rating → Rating / XMP:Rating", f'=IF({bref("Rating")}="","(none)",{bref("Rating")})'),
        ("Keyword Species:", f'=IF({bref("Species")}="","(none)","Species:"&{bref("Species")})'),
        ("Keyword Behavior:", f'=IF({bref("Behavior")}="","(none)","Behavior:"&{bref("Behavior")})'),
        ("Keyword Locale:", f'=IF({bref("Locale")}="","(none)","Locale:"&{bref("Locale")})'),
        ("Keyword SendToUK:", f'=IF({bref("Send to UK")}="","(none)","SendToUK:"&{bref("Send to UK")})'),
        ("Keyword View:", f'=IF({bref("View")}="","(none)","View:"&{bref("View")})'),
        ("Keyword Lighting:", f'=IF({bref("Lighting")}="","(none)","Lighting:"&{bref("Lighting")})'),
        ("Location → Locale", f'=IF({bref("Locale")}="","(none)",{bref("Locale")})'),
        ("Description → notes", f'=IF({bref("notes")}="","(none)",{bref("notes")})'),
    ]
    for label, formula in preview_rows:
        put(r, 1, label, LABEL_FONT)
        put(r, 2, formula)
        r += 1


def add_list_sheet(wb: Workbook):
    ws = wb.create_sheet("Lists", 0)
    ws["A1"] = (
        "Option lists for Entry / VideoEntry click-to-pick (columns C+). "
        "Edit/add values below; keep a contiguous block per column."
    )
    ws["A1"].font = HINT_FONT
    ws.merge_cells("A1:I1")

    for col_idx, (name, values) in enumerate(LISTS.items(), start=1):
        header = ws.cell(2, col_idx, name)
        header.font = HEADER_FONT
        header.alignment = Alignment(horizontal="left")
        for row_idx, value in enumerate(values, start=3):
            cell = ws.cell(row_idx, col_idx, value)
            cell.font = PLAIN_FONT
            cell.alignment = Alignment(horizontal="left")

    for col_idx in range(1, len(LISTS) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = LISTS_COL_WIDTH

    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = True
    return ws


def add_catalog_sheet(wb: Workbook, name: str, headers: list[str]):
    ws = wb.create_sheet(name, 0)
    ws.sheet_view.showGridLines = True

    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(1, col_idx, header)
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = CATALOG_COL_WIDTH
        if header == "Date":
            for r in range(1, 201):
                ws.cell(r, col_idx).number_format = DATE_FORMAT
        elif header == "Time":
            for r in range(1, 201):
                ws.cell(r, col_idx).number_format = TIME_FORMAT

    hint = ws.cell(
        1,
        len(headers) + 1,
        "Shortcuts: Control+Shift+O Preview · Control+Shift+J Finder",
    )
    hint.font = HINT_FONT
    hint.alignment = Alignment(horizontal="left", vertical="center")
    ws.column_dimensions[get_column_letter(len(headers) + 1)].width = 52

    ws.freeze_panes = "A2"

    for list_name in LISTS:
        if list_name not in headers:
            continue
        col_idx = headers.index(list_name) + 1
        col_letter = get_column_letter(col_idx)
        dv = DataValidation(
            type="list",
            formula1=list_range(list_name),
            allow_blank=True,
            showDropDown=False,
        )
        ws.add_data_validation(dv)
        dv.add(f"{col_letter}2:{col_letter}5000")

    return ws


def add_readme_sheet(wb: Workbook):
    ws = wb.create_sheet("ReadMe")
    ws.sheet_view.showGridLines = True
    ws["A1"] = "Strider Media Catalog"
    ws["A1"].font = TITLE_FONT

    lines = [
        "",
        "Sheets",
        "Entry — add photos (LoadFromFinderSelection → AppendPhotoRow).",
        "VideoEntry — add videos (LoadVideoFromFinderSelection → AppendVideoRow).",
        "Photos / Videos — catalogs.",
        "",
        "Photos columns",
        ", ".join(PHOTO_HEADERS),
        "",
        "Videos columns",
        ", ".join(VIDEO_HEADERS),
    ]
    for i, line in enumerate(lines, start=2):
        ws[f"A{i}"] = line
        ws[f"A{i}"].font = PLAIN_FONT
        ws[f"A{i}"].alignment = Alignment(horizontal="left")
    return ws


def main():
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    add_list_sheet(wb)
    add_readme_sheet(wb)
    add_catalog_sheet(wb, "Videos", VIDEO_HEADERS)
    add_catalog_sheet(wb, "Photos", PHOTO_HEADERS)
    add_form_sheet(
        wb,
        "VideoEntry",
        "This sheet collects video data before adding to the Videos catalog",
        [
            (3, "1) Select video in Finder, or open it in Preview."),
            (4, "2) LoadVideoFromFinderSelection (Ctrl+Shift+E) or LoadVideoFromPreviewSelection (Ctrl+Shift+U)."),
            (5, "3) Click an option in columns C+ (Tags) to put it in column B."),
            (6, "4) AppendVideoRow  (Control+Shift+R)"),
            (7, "Review on Videos: Control+Shift+O Preview · Control+Shift+J Finder"),
        ],
        VIDEO_ENTRY_FIELDS,
        VIDEO_ENTRY_AUTO_COUNT,
        "Auto (from Load / Preview)",
        VIDEO_ENTRY_COL_WIDTHS,
    )
    add_form_sheet(
        wb,
        "Entry",
        "This sheet collects photo data before adding to the Photos catalog",
        [
            (3, "1) Select photo in Finder, or step to it in Preview."),
            (4, "2) LoadFromFinderSelection (Ctrl+Shift+Q) or LoadFromPreviewSelection (Ctrl+Shift+I)."),
            (5, "3) Click an option in columns C+ (Tags) to put it in column B."),
            (6, "4) AppendPhotoRow  (Control+Shift+W)"),
            (7, "Review on Photos: Control+Shift+O Preview · Control+Shift+J Finder"),
        ],
        ENTRY_FIELDS,
        ENTRY_AUTO_COUNT,
        "Auto (from Load / Preview)",
        ENTRY_COL_WIDTHS,
    )

    order = ["Entry", "VideoEntry", "Photos", "Videos", "Lists", "ReadMe"]
    for idx, name in enumerate(order):
        wb.move_sheet(name, offset=idx - wb.sheetnames.index(name))

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
