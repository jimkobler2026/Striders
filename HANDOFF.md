# Striders media catalog — handoff

**Date:** 2026-11-13  
**For:** Fresh agent — next focus is **sequence handling** (see queue below)  
**Repo:** `/Users/jim/Development/projects/Striders` (branch `main`)  
**Day-to-day user guide:** `README.md` (includes “Sequence tagging and bulk import”)

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

Smoke: Load (I/Q) → click-pick → Append (W); sequence bulk (S); caption (Y); Name double-click; Caption click; HaveRaw?/pixels on Load.

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
| K | AddCatalogField |

Helpers: `~/Library/Application Scripts/com.microsoft.Excel/`  
(`StriderFinderMeta`, `StriderOpenPreview`, `StriderRevealFinder`, `StriderCaption`, `StriderCaption.sh`, `StriderSequence.sh`)

API keys for captions: `../Virtual_Nexus/.env` (preferred), else repo `.env` (gitignored).
