# Striders media catalog — handoff

**Date:** 2026-11-16  
**For:** Fresh agent — T tag-write batched (speed); B/T + status-bar + friend docs shipped; sequence queue still open below  
**Repo:** `/Users/jim/Development/projects/Striders` (branch `main`)  
**Day-to-day user guide:** `README.md` (includes “Sequence tagging and bulk import”)  
**Friend / non-technical setup:** `SETUP_FOR_FRIENDS.txt` (also linked from README “Sharing with another person”)

---

## Current state (read first)

### Canonical catalog files

| File | Role |
|---|---|
| `catalog/build/Strider_Media_Catalog_blank.xlsm` | **Template of record** — duplicate this for a working catalog |
| `catalog/main catalog/Strider_Media_Catalog.xlsm` | Working catalog (tracked) |
| `catalog/backup of blank catalog/` | Optional blank snapshots (gitignored) |
| `catalog/macros/CatalogMacros.bas` | VBA source of record |
| `catalog/macros/*.applescript` | Excel helper sources |
| `catalog/macros/PhotosSheet_SelectionChange.txt` | Photos stub (Caption click + Name double-click) |
| `catalog/macros/VideosSheet_SelectionChange.txt` | Videos stub (same) |

Working catalog: `catalog/main catalog/Strider_Media_Catalog.xlsm`. Duplicate the blank for new books; do **not** run `setup_workbook.sh` casually.

### Shipped recently (2026-11-13)

- **Caption column + Mac edit dialog:** Ctrl+Shift+Y → Photos **Caption** + embed `CAPTION:`; click Caption → `editCaption` dialog; Save re-embeds. Prompt is Scientific American–style plain prose (`caption_prompt.txt` + VBA `CaptionPromptText`).
- **Name open = double-click:** No Name hyperlinks; double-click Name → Preview; Ctrl+Shift+O still multi-opens. `StripCatalogNameHyperlinks` / `ReapplyUiOnlyProtection` clears legacy links. Photos/Videos data area unlocked so Name/Caption are selectable under `EnableSelection=xlUnlockedCells`.
- **HaveRaw? / VertPix / HorPix:** Columns on Photos (HaveRaw? after Path; VertPix/HorPix after FileSize). Load detects sibling RAW (`.arw`/`.cr2`/`.cr3`/`.nef`/`.nrw`/`.raf`/`.dng`) + ImageHeight/Width. Append writes keywords `rawYes`|`rawNo` and `VertPix:`/`HorPix:`. Macro **`EnsureRawAndPixelFields`** (from `AssignMacroShortcuts`). Backfill: `scripts/fill_raw_and_pixels.py --catalog` (`--write-keywords` optional).
- **Batch import + tag copy (2026-11-14):** Ctrl+Shift+**B** `BatchImportSelectedPhotos` — Finder (or Preview) multi-select → Photos metadata rows only. Ctrl+Shift+**T** `CopyTagsToSelectedPhotos` — select Photos Name cells; copy user tags from seed row + Keyword write-back. No SequenceID auto-expand (**S** still expands). Helper: `listBatchImportPaths` in StriderFinderMeta.
- **Status-bar progress (2026-11-15):** Slow macros **B** / **T** / **S** / **Y** use `ProgressBegin` / `ProgressStep` / `ProgressDone` (Excel status bar + wait cursor). Look at the bottom of the Excel window while batches run.
- **Click-pick after empty Entry stub:** If Option clicks stop filling Entry after a batch, paste `EntrySheet_SelectionChange.txt` / `VideoEntrySheet_SelectionChange.txt` into the sheet modules, or run **`FixEnableEvents`**. Empty stubs break click-pick even when EnableEvents is True.
- **Sharing docs (2026-11-15):** README “Sharing with another person”; plain-language **`SETUP_FOR_FRIENDS.txt`**. Share blank + macros/scripts, not personal working catalog / `.env`.
- **T write speed (2026-11-16):** `CopyTagsToSelectedPhotos` batches Keyword writes via `writeCatalogTagsBatch` (chunks of 25) instead of one `AppleScriptTask` per file; exiftool path resolved once per chunk.

### Already on `main` (earlier)

- Unique renames, sequence tagging (`tag_sequences.py` SequenceNumber boundaries), bulk Append **S**, dynamic fields **K**, MWG write-back, NOTE:/CAPTION: Description layout, exiftool feature queue (item 14 HierarchicalSubject skipped).

### Media root

Ingest folders under `~/Pictures/environmental caustics/`. Prefer **ingest folder** over “dump.” Largest tagged set: `purgatory_striders_11022026` (~1663 files / ~104 sequences).

### Bootstrap after pull / helper edits

```bash
./scripts/install_mac_helpers.sh
# User: quit/restart Excel, re-import CatalogMacros.bas if VBA changed,
# paste Photos/Videos stubs if missing, run AssignMacroShortcuts, Save
```

Smoke: Load (I/Q) → click-pick → Append (W); sequence bulk (S); batch **B** then **T**; caption (Y); Name double-click; Caption click; HaveRaw?/pixels on Load.

User action after pull: re-import `CatalogMacros.bas`, `./scripts/install_mac_helpers.sh` + restart Excel, run **`AssignMacroShortcuts`**.

---

## Rules for the next agent

1. **One suggestion at a time** in the sequence queue below (skip only if marked done or user says skip).
2. After **each** item: run that item’s **Test**, then a short regression (Load / Append / S still OK). Do not start the next item until both pass.
3. Prefer small diffs; match existing style; update `README.md` briefly when behavior changes.
4. Never commit secrets (`.env`). Only commit when the user asks.
5. Do not run `setup_workbook.sh` with `--force` unless explicitly requested.
6. Sandbox: renaming/writing under `~/Pictures/…` needs full permissions (`all`).
7. Leave a short note under **Progress log** after each item.
8. **Multi-camera:** Canon and others alongside Sony; phones **iPhone 12/18 Pro Max** (often geotagged; usually not Continuous-tagged). Prefer generic exiftool tags; do not hard-code Sony-only assumptions until a real ingest shows a problem.

---

## Next queue — sequence handling

User priority after 2026-11-13 caption/raw/pixel work: **improve how sequences are handled** in tagging, catalog, and bulk Append. Propose and implement **one item at a time**; confirm with the user before large UX changes.

### Context already in place

| Piece | Role |
|---|---|
| `scripts/tag_sequences.py` | Disk Keywords: Sequence / SequenceID / SequencePos (SequenceNumber bursts; gap fallback; `--min-frames` default 4) |
| Ctrl+Shift+**S** / `AppendPhotoSequence` | Bulk-append all frames of Entry’s SequenceID; shared tags from Entry (use 1/n); per-frame EXIF + SequencePos |
| `list_sequence_members.py` + `StriderSequence.sh` | List members for S |
| `scripts/copy_sequence_tags.py` | Copy catalog Keywords across sequence siblings |

### Suggested sequence work (order — refine with user)

#### S1. Inventory / clarify current sequence UX gaps

**Goal:** With the user, list what’s painful today (e.g. re-tag after renames, Photos filters by SequenceID, reviewing a burst in Preview, RAW+JPEG pairs inside a sequence, partial Append, SequencePos display).  
**Test:** Short written punch-list agreed before coding.

#### S2. Harden bulk Append (S) edge cases

**Goal:** Clearer errors when SequenceID missing / folder has no members; preserve HaveRaw?/VertPix/HorPix per frame (already excluded from seed); optional progress StatusBar; don’t clear Entry mid-failure.  
**Test:** Append a known SequenceID from `purgatory_striders_*` or a small temp burst; spot-check SequencePos + HaveRaw? on a few rows.

#### S3. Sequence-aware review helpers

**Goal:** From a Photos selection or Entry SequenceID, open all (or N) frames in Preview / reveal folder — without fighting Name double-click.  
**Test:** One sequence opens; Cancel/limit behavior documented in README.

#### S4. Catalog filters / columns for sequences

**Goal:** Document or add helpers to filter Photos by SequenceID / Sequence=yes; ensure Sequence* columns stay text (no fraction coercion).  
**Test:** Filter a SequenceID; SequencePos still shows `6/54`-style text.

#### S5. Retag / repair after ingest changes

**Goal:** Safe re-run path for `tag_sequences.py --force` on one ingest folder; optional sync of Sequence* from disk → existing Photos rows (by Path/Name).  
**Test:** Dry-run then write on a small folder; Photos rows update or a report lists mismatches.

#### S6. JPEG+RAW within sequences

**Goal:** Decide whether sequences are JPEG-only in the catalog (HaveRaw? notes sibling) or whether RAW frames should be excluded from SequenceID membership / Append. Align `tag_sequences.py` extensions with that decision.  
**Test:** Folder with JPG+ARW pairs → expected Sequence membership + HaveRaw?=Y on JPGs.

---

## Regression gate (any sequence item that touches Load/Append/S)

1. `python3 -m py_compile scripts/tag_sequences.py scripts/list_sequence_members.py scripts/fill_raw_and_pixels.py`
2. Dry-run `tag_sequences.py` on a small ingest folder exits 0.
3. If VBA/AppleScript touched: user smoke **I → W** and **S** on one short sequence; Caption click + Name double-click still work.
4. Note results in **Progress log**.

---

## Progress log

| Item | Date | Result | Notes |
|---|---|---|---|
| Exiftool queue 1–15 | 2026-11-10 | Done / 14 skipped | See prior log; HierarchicalSubject deferred |
| NOTE/CAPTION Description | 2026-11-10 | Done | writeCaption + Append NOTE merge |
| Caption column + edit dialog | 2026-11-13 | Done | SciAm prompt; Mac dialog; Captions sheet retired |
| Name double-click open | 2026-11-13 | Done | Hyperlinks stripped; data area unlock |
| HaveRaw? / VertPix / HorPix | 2026-11-13 | Done | Load + Append keywords; fill_raw_and_pixels.py |
| Batch import B + tag copy T | 2026-11-14 | Done | listBatchImportPaths; metadata-only B; Photos selection T |
| Status-bar progress B/T/S/Y | 2026-11-15 | Done | ProgressBegin/Step/Done; wait cursor |
| Entry/VideoEntry stub paste + FixEnableEvents note | 2026-11-15 | Done | Empty SelectionChange broke click-pick after B |
| Friend sharing README + SETUP_FOR_FRIENDS.txt | 2026-11-15 | Done | Non-technical install + daily-use guide |
| T batched Keyword writes | 2026-11-16 | Done | writeCatalogTagsBatch + VBA chunks of 25; re-install helpers + re-import BAS |
| Sequence handling queue | 2026-11-13 | Next | S1–S6 above — start with user punch-list |

---

## Quick reference

| Shortcut | Macro |
|---|---|
| Q / I | Load photo Finder / Preview |
| W | Append photo + tag write-back |
| S | AppendPhotoSequence (tags from Entry / 1/n) |
| E / U / R | Video load / append |
| O / J / D | Preview / Finder / trash |
| Y | Photos Caption column + embed CAPTION: |
| B | BatchImportSelectedPhotos (Finder/Preview → metadata rows) |
| T | CopyTagsToSelectedPhotos (seed row tags → selection) |
| K | AddCatalogField |

Helpers: `~/Library/Application Scripts/com.microsoft.Excel/`  
(`StriderFinderMeta`, `StriderOpenPreview`, `StriderRevealFinder`, `StriderCaption`, `StriderCaption.sh`, `StriderSequence.sh`)

API keys for captions: `../Virtual_Nexus/.env` (preferred), else repo `.env` (gitignored).
