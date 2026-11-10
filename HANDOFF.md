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

Next agent: start at **item 1**. Run `./scripts/install_mac_helpers.sh` and ask the user to restart Excel after any helper edit.

### Media root

Dumps live under `~/Pictures/environmental caustics/`. Sequence retag of all image folders was completed 2026-11-10 (largest: `purgatory_striders_11022026`, 104 sequences / 1551 frames). Prefer **copies under `/tmp`** for write tests — do not rewrite whole dumps unless the user asks.

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

---

## Exiftool feature queue (15)

Context: catalog already uses exiftool for dates, `ReleaseMode`, Keywords/Subject, Rating/Description/Location, and Load exposure/lens fields. These items deepen that. Camera samples are mostly **Sony ILCE-7RM4**; phone dumps are **Apple** with GPS.

---

### 1. Sony `SequenceNumber` for sequence tagging

**Status:** TODO  
**Goal:** Use maker-note `SequenceNumber` (resets to **1** at each burst start on A7RIV) to harden `tag_sequences.py` beyond the 1 s capture-time gap heuristic; optionally align `SequencePos` with camera counters when trustworthy.

**Why:** Gap-only clustering can mis-merge/split bursts; Sony already numbers frames within a continuous burst.

**Likely files:** `scripts/tag_sequences.py`; maybe `README.md` sequence section.

**Hints:**
- Read `-SequenceNumber` (and keep `-ReleaseMode`) in the collect pass.
- Treat a return to `SequenceNumber == 1` (while still Continuous) as a strong burst boundary; still require `--min-frames` (default 4).
- JPEG+RAW pairs can share a capture instant but may share or duplicate sequence numbers — verify on a real dump before assuming 1:1 with `SequencePos`.
- Keep gap threshold as fallback for cameras without `SequenceNumber`.

**Test:**
1. Dry-run on a known folder, e.g. `purgatory_striders_11022026` (or a small subfolder):  
   `python3 scripts/tag_sequences.py "<path>" --recursive`  
   Compare sequence count / IDs to a prior run or spot-check that burst starts match files where `SequenceNumber=1`.
2. On a **copy** of 20–40 frames under `/tmp`: `--write` and confirm Keywords `Sequence*` via `exiftool -Keywords -Subject` and `SequenceNumber`.
3. Confirm `--min-frames 4` still marks 2–3 frame clusters as `Sequence:no`.

**Regression gate** after this item.

---

### 2. Faster exiftool writes (`-stay_open` / batched `-execute`)

**Status:** TODO  
**Goal:** Speed up per-file writes in `tag_sequences.py` (and optionally Append helpers) by avoiding one process spawn per file.

**Why:** Full-tree sequence write took ~46 minutes largely due to process overhead.

**Likely files:** `scripts/tag_sequences.py`; optionally a tiny shared helper module; **not** required to change AppleScript on the first pass if CLI gains are enough.

**Hints:**
- Prefer batching many files in fewer invocations, or `-stay_open True -@ -` with `-execute`.
- Preserve MWG full-list Keyword replace semantics from item 10.
- Keep dry-run path unchanged.

**Test:**
1. Time write on a `/tmp` copy of ~50–100 previously tagged frames; confirm wall time drops vs one-process-per-file baseline.
2. Spot-check 5 files: Keywords/Subject Sequence* correct; no `_original` leftovers (`-overwrite_original`).
3. Run once with intentional bad path / empty folder — must error cleanly, not hang on stay_open.

**Regression gate** after this item.

---

### 3. `FocusDistance2` on Load / catalog

**Status:** TODO  
**Goal:** Surface Sony `FocusDistance2` (subject distance) on Entry Load and as a Photos column (and write-back only if it is treated as a user field — prefer **read-only auto** like Lens/Speed).

**Why:** Highly relevant to caustics / strider working distance; present on sampled A7RIV frames.

**Likely files:** `catalog/macros/StriderFinderMeta.applescript` (photo metadata format string); `CatalogMacros.bas` (parse Load); `add_catalog_field` / blank layout or `build_catalog.py` if adding a column; Lists only if click-pick (probably not).

**Hints:**
- Prefer auto field (A/B or C/D) filled on Load, not a Keyword write-back, unless user asks to persist it.
- Display units clearly (exiftool usually returns e.g. `1.654 m`).

**Test:**
1. `exiftool -FocusDistance2` on a known Sony JPG; Load that file into Entry (I/Q) and confirm the field fills.
2. Append (W): Photos row gets the value if it’s a catalog column; confirm Keywords were **not** polluted with FocusDistance unless intentionally designed.
3. Load a phone JPG without the tag — field blank, no error.

**Regression gate** after this item.

---

### 4. Phone GPS → catalog / Locale assist

**Status:** TODO  
**Goal:** On Load (and optionally batch script), read `GPSPosition` / lat / lon from geotagged files (iPhone dumps already have GPS). Optionally suggest or fill Locale.

**Why:** Phone caustics under `environmental caustics/phone caustics*` are geotagged; mirrorless dumps usually are not.

**Likely files:** `StriderFinderMeta.applescript`; `CatalogMacros.bas`; maybe Photos columns `GPSLatitude` / `GPSLongitude` or a single `GPS` text field; README.

**Hints:**
- Do not overwrite a non-empty Locale without an explicit user action.
- Reverse-geocode is optional (exiftool geolocation API / external); lat/lon storage alone is valuable.
- Videos: QuickTime GPS may differ — check phone MOV samples.

**Test:**
1. Load a known phone file with GPS (e.g. under `phone caustics 11042026`); Entry shows coordinates (and Locale only if auto-fill is implemented and Locale was empty).
2. Load a Sony brook JPG without GPS — blank GPS fields, existing Locale untouched.
3. Append (W): row stores GPS as designed; MWG Keywords still OK.

**Regression gate** after this item.

---

### 5. Preserve / restore file times after metadata writes

**Status:** TODO  
**Goal:** Stop Append / `tag_sequences` from making Finder “Date Modified” = now; use `-P` and/or set `FileModifyDate` from `DateTimeOriginal` after writes.

**Why:** Sequence retag already bumped mtimes on large dumps; breaks capture-time Finder sorting.

**Likely files:** `StriderFinderMeta.applescript` (`writeTagsToFile`); `scripts/tag_sequences.py`; any other writers (`rename_media_unique` should keep content times sensible).

**Hints:**
- `-P` preserves existing FileModifyDate/FileCreateDate through the write.
- Or `-FileModifyDate<DateTimeOriginal` (and Mac creation date APIs if desired).
- Decide one policy and document it.

**Test:**
1. Copy a JPG to `/tmp`; note `exiftool -FileModifyDate -DateTimeOriginal`.
2. Run writeTags-equivalent / `tag_sequences --write` on the copy; confirm FileModifyDate policy holds (unchanged with `-P`, or equals capture time if restoring).
3. Confirm Keywords still updated.

**Regression gate** after this item.

---

### 6. `ImageDataHash` for duplicate detection

**Status:** TODO  
**Goal:** Add a script (or flag) that reports duplicate image **content** via `ImageDataHash` (hash of image data, ignoring metadata/filename).

**Why:** Overlapping `temp photos*` folders can hold duplicate JPGs under different names after renames. Not for JPEG↔RAW matching.

**Likely files:** new `scripts/find_duplicate_media.py` (or similar); README one-liner.

**Hints:**
- `exiftool -ImageDataHash -r -csv` / JSON in batches; group by hash.
- Default dry-run report only; deletion must be explicit and cautious.

**Test:**
1. Copy one JPG twice to `/tmp/a.jpg` `/tmp/b.jpg`; script reports them as duplicates.
2. Change Keywords on a third copy — hash should still match (metadata-independent).
3. Two different photos — different hashes. Exit 0 on report-only.

**Regression gate** (script-only: run sequence dry-run + one MWG keyword read still works).

---

### 7. `-TagsFromFile` sibling copy (CLI complement to Ctrl+Shift+S)

**Status:** TODO  
**Goal:** CLI to copy catalog Keywords (and optionally notes/rating/location) from a tagged **1/n** (or any) file onto SequenceID siblings or onto matching RAW/JPG mates — without Excel.

**Why:** Bulk metadata on disk when Excel isn’t open; useful for RAW+JPG pairs.

**Likely files:** new `scripts/copy_sequence_tags.py` or extend `list_sequence_members.py`; README.

**Hints:**
- Prefer copying MWG:Keywords / Description / Rating / Location, not entire EXIF (don’t clobber camera Make/Model/datetime).
- Default dry-run.

**Test:**
1. `/tmp` sequence of 4 copies; tag only frame 1 with Species/Behavior Keywords; run tool; frames 2–4 get same user Keywords; each keeps own SequencePos if already set (define behavior clearly).
2. Dry-run prints actions and writes nothing.
3. Refuse to run outside allowed roots if you mirror `rename_media_unique` safety — or document unrestricted paths.

**Regression gate** after this item.

---

### 8. Embedded preview extract for captions / thumbs

**Status:** TODO  
**Goal:** Prefer Sony/JPEG embedded `PreviewImage` (or similar) for `caption_photo.py` instead of sending full ~40 MB masters when a preview exists.

**Why:** Faster/cheaper vision API calls; less I/O.

**Likely files:** `scripts/caption_photo.py`; maybe `StriderCaption.sh`.

**Hints:**
- `exiftool -b -PreviewImage` → temp JPG; fall back to full file if missing.
- Clean up temp files; keep `--plain` / `--machine` behavior identical aside from source pixels.

**Test:**
1. On a Sony JPG with PreviewImage: caption run uses preview (log or debug flag); still returns a caption with API keys present — if no keys, at least verify preview extraction size ≪ original.
2. Phone JPG with no PreviewImage: falls back to full file, no crash.
3. Excel Ctrl+Shift+Y path still works if shell bridge unchanged (or update bridge and smoke once).

**Regression gate** after this item.

---

### 9. Honor `OffsetTimeOriginal` in rename / sequence IDs

**Status:** TODO  
**Goal:** Apply timezone offset when parsing capture times for `rename_media_unique.py` and `tag_sequences.py` SequenceID / clustering.

**Why:** Sample files carry `OffsetTimeOriginal` (e.g. `-04:00`); naive local parse can drift across DST/travel.

**Likely files:** `scripts/rename_media_unique.py`, `scripts/tag_sequences.py` (shared datetime helper if practical).

**Hints:**
- Prefer timezone-aware datetimes; SequenceID format can stay `YYYYMMDD-HHMMSS` in local/offset-adjusted time — document which.
- Don’t change filenames on disk unless user runs rename `--write` again.

**Test:**
1. Unit-style: file with known `DateTimeOriginal` + `OffsetTimeOriginal` → parsed instant matches expectation.
2. Dry-run rename on a small folder — prefixes stable/sensible vs previous behavior (call out intentional changes).
3. Sequence dry-run still clusters the same bursts on a Sony continuous folder (no mass split/merge surprise).

**Regression gate** after this item.

---

### 10. MWG write path

**Status:** **DONE** on `main`  
**Goal:** Write Rating / Description / Location / Keywords via `-use MWG` so EXIF/IPTC/XMP stay synchronized.

**Done in:** `StriderFinderMeta.applescript`, `tag_sequences.py`, docs.  
**Tests already run:** `/tmp` smoke showed IPTC+XMP+EXIF Description aligned; Keywords full-list replace works; helpers reinstalled.

**If revisiting:** re-run `/tmp` MWG smoke; ensure `IPTC:CodedCharacterSet=UTF8` so virgin files get IPTC Keywords.

---

### 11. Geotag mirrorless dumps from GPX / phone track

**Status:** TODO (situational)  
**Goal:** Optional script: `exiftool -geotag track.gpx` (with time drift options) onto Sony dump folders when a phone/GPS track exists for the shoot.

**Why:** A7RIV brook/river stills lack GPS; phone in pocket can provide a track.

**Likely files:** new `scripts/geotag_dump.py`; README safety notes.

**Hints:**
- Dry-run / backup; geotag is destructive to GPS fields.
- Document timezone/`-geosync` needs.
- Only run when user supplies a track file.

**Test:**
1. With a small GPX fixture and `/tmp` images with known times, dry-run then write; `exiftool -GPSPosition` populated.
2. Images outside track time range left without GPS (or warned).
3. No change to Keywords/Sequence*.

**Regression gate** after this item.

---

### 12. `SerialNumber` on Load / catalog

**Status:** TODO (situational)  
**Goal:** Show camera body serial on Load / Photos for multi-body distinction.

**Likely files:** `StriderFinderMeta.applescript` metadata format; Entry + Photos column via AddCatalogField pattern (auto field).

**Test:** Load Sony JPG → serial matches `exiftool -SerialNumber`. Phone may differ or blank — no error. Append doesn’t write serial into Keywords.

**Regression gate** after this item.

---

### 13. Artist / Copyright stamp on ingest

**Status:** TODO (situational)  
**Goal:** Optional ingest stamp (`MWG:Creator` / `MWG:Copyright` or Artist/Copyright) via rename or a small `stamp_ownership.py`, configured by env or repo config (not hard-coded secrets).

**Test:** `/tmp` copy stamped; `exiftool -Artist -Copyright -MWG:Creator -MWG:Copyright` shows values; second run idempotent; unrelated Keywords preserved.

**Regression gate** after this item.

---

### 14. HierarchicalSubject (Lightroom-style trees)

**Status:** TODO (situational — only if user wants LR hierarchy)  
**Goal:** Optionally mirror flat `Key:value` Keywords into `XMP-lr:HierarchicalSubject` or MWG hierarchical keywords (e.g. `Species|Aquarius remigis`).

**Hints:** Do **not** replace flat Keywords the catalog depends on; additive only. Confirm Preview/Excel Load still parse Sequence* from IPTC/XMP Subject.

**Test:** After write, flat Keywords unchanged; hierarchical tag present; Load + Append still round-trip Species/Sequence*.

**Regression gate** after this item.

---

### 15. `-diff` smoke test after write-back changes

**Status:** TODO (situational / tooling)  
**Goal:** Add a small script or documented command that copies a file, runs the catalog write payload, and uses `exiftool -diff` (minus noisy system tags) to show what changed — for CI-less regression when touching write-back.

**Likely files:** `scripts/test_tag_writeback.py` or a README “Dev testing” snippet.

**Test:** Run the tool itself against `/tmp` copy; assert expected tags appear in diff (Rating/Description/Keywords) and camera Make/Model/DateTimeOriginal do not flip unexpectedly.

**Regression gate** after this item.

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
| 1 SequenceNumber | | | |
| 2 stay_open / batch writes | | | |
| 3 FocusDistance2 | | | |
| 4 Phone GPS | | | |
| 5 File times `-P` | | | |
| 6 ImageDataHash | | | |
| 7 TagsFromFile siblings | | | |
| 8 PreviewImage captions | | | |
| 9 OffsetTimeOriginal | | | |
| 11 Geotag GPX | | | |
| 12 SerialNumber | | | |
| 13 Artist/Copyright | | | |
| 14 HierarchicalSubject | | | |
| 15 `-diff` writeback test | | | |

---

## Quick reference

| Shortcut | Macro |
|---|---|
| Q / I | Load photo Finder / Preview |
| W | Append photo + tag write-back |
| S | AppendPhotoSequence (tags from Entry / 1/n) |
| E / U / R | Video load / append |
| O / J / D | Preview / Finder / trash |
| Y | Captions |
| K | AddCatalogField |

Helpers: `~/Library/Application Scripts/com.microsoft.Excel/`  
(`StriderFinderMeta`, `StriderOpenPreview`, `StriderRevealFinder`, `StriderCaption`, `StriderCaption.sh`, `StriderSequence.sh`)

API keys for captions: `../Virtual_Nexus/.env` (preferred), else repo `.env` (gitignored).
