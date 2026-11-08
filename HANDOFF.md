# Striders media catalog — handoff

**Status:** Ready for real use. Next step is to catalog media and tweak tags/macros as needed.

**Working copy:** `catalog/Strider_Media_Catalog.xlsm`  
**Blank template:** `catalog/Strider_Media_Catalog_blank.xlsm` (mirrored in `catalog/build/`)

Always open the `.xlsm`, not the build `.xlsx`.

---

## What we built

A Mac Excel media catalog for water-strider **photos** and **videos**, focused on behavior, locale, lighting/caustics, and “Send to UK” tracking.

| Piece | Role |
|---|---|
| `catalog/Strider_Media_Catalog.xlsm` | Live workbook (macros + sheets) |
| `catalog/Strider_Media_Catalog_blank.xlsm` | Empty starter with same layout/macros |
| `catalog/macros/CatalogMacros.bas` | All VBA (import into VBE) |
| `catalog/macros/*Sheet_SelectionChange.txt` | Click-to-pick handlers for Entry / VideoEntry |
| `catalog/macros/StriderFinderMeta.applescript` (+ Open/Reveal) | Excel → Finder/Preview/exiftool bridge |
| `scripts/install_mac_helpers.sh` | Compiles `.scpt` into Excel’s Application Scripts folder |
| `scripts/build_catalog.py` | Regenerates sheet layout / lists into `catalog/build/*.xlsx` |
| `README.md` | Setup + day-to-day workflow |

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

**Photos columns:** Name, Path, Date, Time, Rating, Species, Behavior, Locale, Type, FileSize, Send to UK, View, Camera, Mode, Lens, Speed, f-stop, ISO, Lighting, notes  

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
./scripts/install_mac_helpers.sh   # after AppleScript changes
python3 scripts/build_catalog.py   # sheet skeleton → catalog/build/*.xlsx
```

Prefer keeping the hand-tuned **blank `.xlsm`** as the template of record once macros are imported. Mirror a fresh blank into `catalog/build/` when you save one.

---

## Known quirks

- Do **not** keep an **AppEvents** class module — it breaks on Mac Excel.  
- Old OnKey letters (L/A/M/F/…) conflicted with Excel built-ins; current set is Q/I/W/E/U/R/O/J/D.  
- Photos “dark line” between rows 4–5 was linked to **vertical page breaks** / view mode — use **View → Normal** if a line reappears.  
- Append finds rows by **filename** for overwrite; rename files carefully.  
- Tag vocab lives on **Lists** and the Entry/VideoEntry option columns — keep them in sync when you add terms.

---

## Suggested next steps

1. Run a real cull session tomorrow (Preview → Load → tag → Append → trash rejects).  
2. Adjust Lists / click-pick options as real media demands.  
3. Note any Load/Append/tag-write failures and fix in `CatalogMacros.bas` + AppleScript helpers.  
4. Optionally drop the obsolete build `.xlsx` once the blank `.xlsm` is the only template you use.  
5. Commit again after meaningful tag or macro changes.
