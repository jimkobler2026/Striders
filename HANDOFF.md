# Striders media catalog — handoff

**Status:** Ready for real use. Next: one real Preview cull session on the working-copy `.xlsm` (see Suggested next steps).

**Canonical files (same story as README):**

| Priority | File | Role |
|---|---|---|
| 1 | `catalog/Strider_Media_Catalog.xlsm` | **Working copy** — open this for culling |
| 2 | `catalog/Strider_Media_Catalog_blank.xlsm` | **Template of record** — empty + macros; copy for a fresh catalog |
| 3 | `catalog/build/Strider_Media_Catalog_blank.xlsm` | Mirror of the blank |
| — | `catalog/build/Strider_Media_Catalog.xlsx` | Skeleton only (no macros); input to cold rebuild |

Always open a `.xlsm`, never the build `.xlsx`. Day-to-day user guide: `README.md`.

---

## What we built

A Mac Excel media catalog for water-strider **photos** and **videos**, focused on behavior, locale, lighting/caustics, and “Send to UK” tracking.

| Piece | Role |
|---|---|
| `catalog/Strider_Media_Catalog.xlsm` | Working copy (macros + sheets + data) |
| `catalog/Strider_Media_Catalog_blank.xlsm` | Template of record (empty + same layout/macros) |
| `catalog/macros/CatalogMacros.bas` | All VBA (import into VBE) |
| `catalog/macros/*Sheet_SelectionChange.txt` | Click-to-pick handlers for Entry / VideoEntry |
| `catalog/macros/StriderFinderMeta.applescript` (+ Open/Reveal) | Excel → Finder/Preview/exiftool bridge |
| `README.md` | Setup + day-to-day workflow |

### Scripts

| Script | When to use | Touches live `.xlsm`? |
|---|---|---|
| `scripts/install_mac_helpers.sh` | After AppleScript helper edits; safe anytime | No — only `~/Library/Application Scripts/com.microsoft.Excel/` |
| `scripts/build_catalog.py` | Regenerate sheet skeleton / Lists into `catalog/build/*.xlsx` | No |
| `scripts/tag_sequences.py` | Tag hi/Continuous photo sequences by capture-time gap (dry-run default; `--write` to apply) | No — only the folder you pass |
| `scripts/add_sequence_fields.py` | Insert Sequence\* rows/cols into an **open** `.xlsm` via xlwings | **Yes** — open workbook |
| `scripts/setup_workbook.sh` | Cold rebuild of a new `.xlsm` shell from that `.xlsx` | **Yes — overwrites** (refuses unless `--force`) |
| `scripts/apply_video_sheets.py` | One-shot: add/refresh VideoEntry + Videos via xlwings | **Yes** — open workbook must already be the live `.xlsm`; not for daily use |
| `scripts/setup_xlsm.applescript` | Incomplete stub; superseded by `setup_workbook.sh` | Do not run |
| `scripts/upgrade_entry_layout.applescript` | Obsolete Entry layout (old field names); superseded by current blank/live workbooks | Do not run — would wipe Entry contents |

Helpers are installed at:

`~/Library/Application Scripts/com.microsoft.Excel/`  
(`StriderFinderMeta.scpt`, `StriderOpenPreview.scpt`, `StriderRevealFinder.scpt`)

Requires **exiftool** (`brew install exiftool`).

---

## Sheets

- **Entry** / **VideoEntry** — load one file, tag in column B, click options in C+, append to catalog  
- **Photos** / **Videos** — catalog tables (filter, open, reveal)  
- **Lists** — vocabulary source for click-picks (edit here to change options)  
- **ReadMe** — short in-workbook notes  

**Photos columns:** Name, Path, Date, Time, Sequence, SequenceID, SequencePos, Rating, Species, Behavior, Locale, Type, FileSize, Send to UK, View, Camera, Mode, Lens, Speed, f-stop, ISO, Lighting, notes  

**Sequence fields** come from Keywords written by `scripts/tag_sequences.py` (`Sequence:yes|no`, `SequenceID:…`, `SequencePos:i/n`). Load fills them; Append preserves them.

**Videos columns:** Name, Path, Date, Time, Type, FileSize, Duration, FrameRate, Resolution, Camera, Speed, f-stop, ISO, Rating, Species, Behavior, Locale, Send to UK, View, Lighting, notes  

**Date/Time formats (Photos & Videos):** Date `m/d/yy`, Time `[$-F400]h:mm:ss AM/PM` — applied on Load/Append and in the blank template.

Form sheets use **Times New Roman 14**, row height **18**.

---

## Daily workflow (culling)

1. Open files in **Preview**; step through candidates.  
2. **Ctrl+Shift+I** — load front Preview photo into **Entry** (or **U** for video).  
3. Click tag chips in columns C+ (fills B).  
4. **Ctrl+Shift+W** — append photo row (or **R** for video).  
   - Overwrites an existing catalog row with the **same filename**.  
   - Writes tags back into the media file via exiftool.  
5. Reject in Preview with Preview’s own trash (**Edit → Move Selected Image to Trash**), or **Ctrl+Shift+D** from Excel (drives that same Preview menu).

Finder load shortcuts still work: **Q** (photo), **E** (video).

| Shortcut | Action |
|---|---|
| Ctrl+Shift+Q | Load photo from Finder |
| Ctrl+Shift+I | Load photo from front Preview |
| Ctrl+Shift+W | Append photo (+ write tags) |
| Ctrl+Shift+E | Load video from Finder |
| Ctrl+Shift+U | Load video from front Preview |
| Ctrl+Shift+R | Append video (+ write tags) |
| Ctrl+Shift+O | Open selected catalog row(s) in Preview |
| Ctrl+Shift+J | Reveal selected in Finder |
| Ctrl+Shift+D | Trash front Preview document (via Preview menu) |

Shortcuts use **Control+Shift**, not ⌘/⌥ (those collide with Excel on Mac). Re-import `CatalogMacros.bas` after macro edits so `Auto_Open` rebinds them.

---

## Tag write-back (on Append)

Written into the original file:

- Rating → Rating / XMP:Rating  
- notes → Description / Caption  
- Locale → Location  
- Species, Behavior, Locale, Send to UK, View, Lighting → Keywords / Subject as `Species:…`, `Behavior:…`, etc.

---

## Re-import macros after edits

1. Excel → Tools → Macro → Visual Basic…  
2. Remove old **CatalogMacros** module if present.  
3. File → Import → `catalog/macros/CatalogMacros.bas`  
4. Confirm Entry / VideoEntry still have the SelectionChange snippets from the `.txt` files.  
5. Save `.xlsm`.  

If click-picks stop filling B, run **`FixEnableEvents`** once (events can get stuck off after a macro error).

---

## Regenerating layout / blank

```bash
./scripts/install_mac_helpers.sh   # after AppleScript changes (safe; does not touch .xlsm)
python3 scripts/build_catalog.py   # sheet skeleton → catalog/build/*.xlsx only
```

**Do not** run `./scripts/setup_workbook.sh` casually — it rebuilds and **overwrites** the working-copy `.xlsm` (macros/data lost). It refuses if that file exists unless you pass `--force`. Canonical empty template is the hand-tuned **blank `.xlsm`**, not the build `.xlsx`. After meaningful layout/macro changes, update `catalog/Strider_Media_Catalog_blank.xlsm` and mirror it into `catalog/build/`.

---

## Known quirks

- Do **not** import or keep an **AppEvents** class module — it breaks on Mac Excel. The old `catalog/macros/AppEvents.cls` source was removed; if AppEvents still appears in the VBE for a workbook, delete that module and save.  
- `scripts/setup_workbook.sh` overwrites the live `.xlsm`; it requires `--force` if that file already exists. Use `install_mac_helpers.sh` for helper-only updates.  
- Old OnKey letters (L/A/M/F/…) conflicted with Excel built-ins; current set is Q/I/W/E/U/R/O/J/D.  
- Photos “dark line” between rows 4–5 was linked to **vertical page breaks** / view mode — use **View → Normal** if a line reappears.  
- Append finds rows by **filename** for overwrite; rename files carefully.  
- Tag vocab lives on **Lists** and the Entry/VideoEntry option columns — keep them in sync when you add terms.

---

## Sequence tagging

```bash
# Always dry-run first; use a temp copy if testing --write on valuable originals
python3 scripts/tag_sequences.py "/path/to/photos"
python3 scripts/tag_sequences.py "/path/to/photos" --write
```

- Gates on Sony `ReleaseMode` Continuous / Hi / Hi+ (`--include-all` to disable).
- Gap default 1.0 s (`--gap`).
- Prints counts: scanned, skipped, sequences, frames in sequences, singles, size min/med/max.
- After layout update: open `.xlsm` → `python3 scripts/add_sequence_fields.py` → re-import `CatalogMacros.bas` → `./scripts/install_mac_helpers.sh` → restart Excel.

## Suggested next steps

1. **Cull session** — open the working-copy `.xlsm`; Preview → Ctrl+Shift+I → tag → Ctrl+Shift+W → Ctrl+Shift+D for rejects (video: U / R).  
2. **Watch these during the first real session** (fix in macros/helpers if any fail):  
   - Preview load vs Finder load (I/Q and U/E)  
   - Append overwrite of an existing same-filename row  
   - Tag write-back via exiftool (Rating / keywords / notes)  
   - Video duration / frame-rate / resolution on Load  
3. **Vocab** — add terms on **Lists** and mirror them into Entry / VideoEntry option columns when real media needs them.  
4. After layout or macro tweaks you want to keep, refresh the blank `.xlsm` (template of record) and mirror into `catalog/build/`.  
5. Commit after meaningful tag-list or macro/helper changes (docs-only gap fixes can wait for that commit).
