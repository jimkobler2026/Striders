# Striders media catalog — handoff

**Date:** 2026-11-10  
**For:** Fresh agent taking the exiftool feature queue below  
**Repo:** `/Users/jim/Development/projects/Striders` (branch `main`)  
**Day-to-day user guide:** `README.md` (includes “Sequence tagging and bulk import”)

---

## Current state (read first)

### Canonical catalog files

| File | Role |
|---|---|
| `catalog/Strider_Media_Catalog_blank.xlsm` | **Template of record** — duplicate this for a working catalog |
| `catalog/build/Strider_Media_Catalog_blank.xlsm` | Mirror of the blank |
| `catalog/macros/CatalogMacros.bas` | VBA source of record |
| `catalog/macros/*.applescript` | Excel helper sources |

There is **no** checked-in working `Strider_Media_Catalog.xlsm` anymore (removed). Users duplicate the blank. Do **not** run `setup_workbook.sh` casually.

### Already shipped on `main` (`e2de3f4` and earlier)

- Unique renames: `scripts/rename_media_unique.py`
- Sequence tagging (≥4 frames): `scripts/tag_sequences.py`
- Bulk sequence import: Ctrl+Shift+**S** / `AppendPhotoSequence` + `list_sequence_members.py` + `StriderSequence.sh`
- Dynamic catalog fields: Ctrl+Shift+**K** / `AddCatalogField` + `add_catalog_field.py`
- AI captions: Ctrl+Shift+**Y**
- Labeled tag write-back via AppleScriptTask helpers

### MWG write-path (item 10 — DONE on `main`)

- `catalog/macros/StriderFinderMeta.applescript` — Append uses `-use MWG` (`MWG:Rating` / `Description` / `Location` / `Keywords`)
- `scripts/tag_sequences.py` — Sequence* Keywords via `MWG:Keywords` full-list replace (`-sep "||"`)

**Quirk already handled:** with MWG, `-MWG:Keywords=` + `-MWG:Keywords+=…` in one command does **not** clear; must replace the full list in one `-MWG:Keywords=…` (or two-pass clear then add).

**Exiftool feature queue is complete** (item 14 HierarchicalSubject skipped — only if user wants LR trees).

**NOTE:/CAPTION: Description embed (2026-11-10):** `MWG:Description` holds optional `NOTE:` (Entry notes / Append) and `CAPTION:` (Ctrl+Shift+Y / `writeCaption`). Append merges NOTE and preserves CAPTION; Y merges CAPTION and preserves NOTE; Load fills Entry Note from NOTE (legacy unprefixed Description → Note). Also: `caption_photo.py --write`. Helpers: reinstall `install_mac_helpers.sh`.

User action still needed: restart Excel, re-import `CatalogMacros.bas`, run **`AssignMacroShortcuts`**, smoke **I → W** and **Y** (embed + sheet).

### Media root

Ingest folders live under `~/Pictures/environmental caustics/`. Sequence re-identify with SequenceNumber boundaries completed **2026-11-10** (`tag_sequences.py --recursive --write --force`; stay_open). Largest: `purgatory_striders_11022026` (1663 files / ~104 sequences). Prefer the term **ingest folder** over “dump.”

**Entry auto layout (blank):** row 12 Name | Camera | Serial Number (E12/F12, bold label); row 18 Focus Distance | GPS; blank spacer before tag section (currently row 22).

### Bootstrap after pull / helper edits

```bash
./scripts/install_mac_helpers.sh
# User: quit/restart Excel, re-import CatalogMacros.bas if VBA changed, Save
```

Smoke (user or agent-guided): Load (I/Q) → click-pick → Append (W); sequence bulk (S); caption (Y); add field (K).

---

## Rules for the next agent

1. **One suggestion at a time** in the numbered order below (skip only if marked done or user says skip).
2. After **each** item: run that item’s **Test** section, then the **Regression gate** at the bottom. Do not start the next item until both pass.
3. Prefer small diffs; match existing style; update `README.md` briefly when behavior changes.
4. Never commit secrets (`.env`). Only commit when the user asks.
5. Do not run `setup_workbook.sh` with `--force` unless explicitly requested.
6. Sandbox: renaming/writing under `~/Pictures/…` needs full permissions (`all`).
7. Leave a short note under **Progress log** (bottom) after each item.
8. **Multi-camera:** Expect Canon and other bodies alongside Sony; phones include **iPhone 12 Pro Max** and **iPhone 18 Pro Max** (often geotagged; usually not Continuous sequence-tagged). Prefer generic exiftool tags and fallbacks (e.g. SequenceNumber when present, gap otherwise; filename camera number for order). Do **not** hard-code Sony-only assumptions or chase per-model maker-note details until a real ingest folder shows a problem.

---

## Exiftool feature queue (15)

Context: catalog already uses exiftool for dates, `ReleaseMode`, Keywords/Subject, Rating/Description/Location, and Load exposure/lens fields. These items deepen that. Camera samples are mostly **Sony ILCE-7RM4**; phone ingest folders are **Apple** (iPhone 12/18 Pro Max) with GPS.

---

### 1. Sony `SequenceNumber` for sequence tagging

**Status:** **DONE** on `main`  
**Goal:** Use maker-note `SequenceNumber` (resets to **1** at each burst start on A7RIV) to harden `tag_sequences.py` beyond the 1 s capture-time gap heuristic.

**Done in:** `scripts/tag_sequences.py`, `README.md` sequence section.  
**Behavior:** Burst split on `SequenceNumber == 1` or counter drop when both frames have the tag; `--gap` only if absent. Within a burst, `SequencePos` order uses the filename camera number (`…_DSC7659`), not timestamp alone.  
**Tests:** Dry-run on `temp photos 11052026/DCIM/100MSDCF` (32 sequences); `/tmp` write of 31 frames → 3 bursts (5/15/11) with `camSeq` matching `SequencePos`; 3-frame folder → `Sequence:no`; MWG round-trip OK.

---

### 2. Faster exiftool writes (`-stay_open` / batched `-execute`)

**Status:** **DONE** on `main` (uncommitted until user asks)  
**Goal:** Speed up per-file writes in `tag_sequences.py` by avoiding one process spawn per file.

**Done in:** `scripts/tag_sequences.py` (`ExifToolStayOpen`), README. AppleScript Append left unchanged.  
**Behavior:** Reuses Keywords/Subject from the collect JSON pass; `--write` uses one `-stay_open True -@ -` session with per-file MWG full-list replace. **Do not put `--` before the path** in stay_open/`-execute` arg lists — it makes `-execute` look like a filename and hangs.  
**Tests:** 80 files `/tmp` write ~10.5s; 31-frame burst write OK (5/15/11); empty folder / bad path clean; MWG round-trip OK.

---

### 3. `FocusDistance2` on Load / catalog

**Status:** **DONE** (helpers reinstalled; user Excel smoke pending)  
**Goal:** Surface Sony `FocusDistance2` as read-only auto field **Focus Distance** on Entry Load + Photos column.

**Done in:** `StriderFinderMeta.applescript` (meta index 12; Sequence* → 13–15), `CatalogMacros.bas` (Load + exclusions), blank `.xlsm` Entry row 18 + Photos column, `build_catalog.py`, README.  
**Tests:** Helper `readPath` returns `2.621 m` for Sony sample; phone blank FocusDistance slot; MWG/sequence regression OK. **User:** restart Excel, re-import `CatalogMacros.bas`, smoke **I → W** (Focus Distance fills; Keywords not polluted).

---

### 4. Phone GPS → catalog / Locale assist

**Status:** **DONE** (helpers reinstalled; user Excel smoke pending)  
**Goal:** On Load, read `GPSPosition` into a catalog **GPS** field (no Locale auto-fill).

**Done in:** `StriderFinderMeta.applescript` (meta idx 13), `CatalogMacros.bas`, blank Entry C18 + Photos column, `build_catalog.py`, README. Locale never overwritten from GPS.  
**Tests:** Phone `readPath` includes `42 deg 17'…`; Sony GPS slot blank; Focus Distance still idx 12.

---

### 5. Preserve / restore file times after metadata writes

**Status:** **DONE**  
**Goal:** Stop Append / `tag_sequences` from making Finder “Date Modified” = now.

**Policy:** exiftool **`-P`** (preserve existing FileModifyDate/FileCreateDate).  
**Done in:** `StriderFinderMeta.applescript` `writeTagsToFile`, `tag_sequences.py` `mwg_keyword_args`, README.  
**Tests:** `/tmp` copy touched to 2026-01-01 12:00; after `--write` FileModifyDate unchanged; Keywords updated.

---

### 6. `ImageDataHash` for duplicate detection

**Status:** **DONE** — `scripts/find_duplicate_media.py` (report only).

### 7. `-TagsFromFile` sibling copy (CLI complement to Ctrl+Shift+S)

**Status:** **DONE** — `scripts/copy_sequence_tags.py` (MWG user tags; keeps each SequencePos; dry-run default; unrestricted paths — use care).

### 8. Embedded preview extract for captions / thumbs

**Status:** **DONE** — `caption_photo.py` prefers `PreviewImage` / `JpgFromRaw` / `OtherImage`; `--debug-preview`; phone falls back to master.

### 9. Honor `OffsetTimeOriginal` in rename / sequence IDs

**Status:** **DONE** — `scripts/exif_datetime.py`; used by `tag_sequences.py` + `rename_media_unique.py`. SequenceID / rename stamp = camera wall clock (`stamp_local`); gaps use offset-aware datetimes when OffsetTimeOriginal present.

---

### 10. MWG write path

**Status:** **DONE** on `main`  
**Goal:** Write Rating / Description / Location / Keywords via `-use MWG` so EXIF/IPTC/XMP stay synchronized.

**Done in:** `StriderFinderMeta.applescript`, `tag_sequences.py`, docs.  
**Tests already run:** `/tmp` smoke showed IPTC+XMP+EXIF Description aligned; Keywords full-list replace works; helpers reinstalled.

**If revisiting:** re-run `/tmp` MWG smoke; ensure `IPTC:CodedCharacterSet=UTF8` so virgin files get IPTC Keywords.

---

### 11. Geotag mirrorless ingest folders from GPX / phone track

**Status:** **DONE** — `scripts/geotag_ingest.py` (dry-run default; `--geosync`; destructive GPS — test on `/tmp` first).

### 12. `SerialNumber` on Load / catalog

**Status:** **DONE** — meta idx 14 → Entry **Serial Number** (E18) + Photos column; not Keyword write-back. Helper returns `03391080` on sample Sony.

### 13. Artist / Copyright stamp on ingest

**Status:** **DONE** — `scripts/stamp_ownership.py` via `STRIDER_ARTIST` / `STRIDER_COPYRIGHT` or flags.

### 14. HierarchicalSubject (Lightroom-style trees)

**Status:** **SKIPPED** — only if user explicitly wants LR hierarchy later.

### 15. `-diff` smoke test after write-back changes

**Status:** **DONE** — `scripts/test_tag_writeback.py`.

---

## Regression gate (run after every item)

Minimum bar — adapt paths if blank/working book differs:

1. **Helpers compile:** `./scripts/install_mac_helpers.sh` exits 0 (if AppleScript/shell touched).
2. **Sequence dry-run:**  
   `python3 scripts/tag_sequences.py "~/Pictures/environmental caustics/PURGATORY_BROOK" --recursive`  
   exits 0 and prints a summary (no writes).
3. **MWG keyword round-trip on a copy:**  
   Copy any JPG → `/tmp/strider_reg.jpg` → write a harmless Keyword set with the **current** write path (MWG) → read back IPTC:Keywords and XMP-dc:Subject → both contain the test tokens → delete `/tmp` file.
4. **If VBA/AppleScript Load/Append touched:** ask user to smoke **I → tag → W** on one photo (or drive via documented AppleScriptTask if available). Confirm Entry still fills Date/Camera/Lens and Append doesn’t error.
5. **If captions touched:** `python3 scripts/caption_photo.py --help` exits 0; optional one `--plain` if API keys exist.
6. Note results in **Progress log**.

---

## Progress log

| Item | Date | Result | Notes |
|---|---|---|---|
| 10 MWG write path | 2026-11-10 | Done | Full-list Keyword replace; helpers reinstalled |
| 1 SequenceNumber | 2026-11-10 | Done | Burst on cam SeqNum reset; filename # for order; gap fallback |
| 2 stay_open / batch writes | 2026-11-10 | Done | stay_open MWG writes; reuse collect keywords; no `--` before path |
| 3 FocusDistance2 | 2026-11-10 | Done | Auto field; meta idx 12; Excel smoke pending |
| 4 Phone GPS | 2026-11-10 | Done | GPSPosition → GPS; Locale untouched; Excel smoke pending |
| 5 File times `-P` | 2026-11-10 | Done | -P on Append + tag_sequences |
| 6 ImageDataHash | 2026-11-10 | Done | find_duplicate_media.py report-only |
| 7 TagsFromFile siblings | 2026-11-10 | Done | copy_sequence_tags.py |
| 8 PreviewImage captions | 2026-11-10 | Done | embedded preview + --debug-preview |
| 9 OffsetTimeOriginal | 2026-11-10 | Done | exif_datetime.py shared helper |
| 11 Geotag GPX | 2026-11-10 | Done | geotag_ingest.py dry-run default |
| 12 SerialNumber | 2026-11-10 | Done | Serial Number Load field; Excel smoke pending |
| 13 Artist/Copyright | 2026-11-10 | Done | stamp_ownership.py |
| 14 HierarchicalSubject | 2026-11-10 | Skipped | Wait for explicit LR ask |
| 15 `-diff` writeback test | 2026-11-10 | Done | test_tag_writeback.py |
| NOTE/CAPTION Description embed | 2026-11-10 | Done | writeCaption + Append NOTE merge; Load NOTE; caption_photo --write |

---

## Quick reference

| Shortcut | Macro |
|---|---|
| Q / I | Load photo Finder / Preview |
| W | Append photo + tag write-back |
| S | AppendPhotoSequence (tags from Entry / 1/n) |
| E / U / R | Video load / append |
| O / J / D | Preview / Finder / trash |
| Y | Captions sheet + embed CAPTION: |
| K | AddCatalogField |

Helpers: `~/Library/Application Scripts/com.microsoft.Excel/`  
(`StriderFinderMeta`, `StriderOpenPreview`, `StriderRevealFinder`, `StriderCaption`, `StriderCaption.sh`, `StriderSequence.sh`)

API keys for captions: `../Virtual_Nexus/.env` (preferred), else repo `.env` (gitignored).
