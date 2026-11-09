# Striders media catalog — handoff

**Date:** 2026-11-09 (evening) → continue morning of 2026-11-10  
**Status:** Captions + labeled tag write-back are in git. **Excel still needs a morning bootstrap** (re-import `.bas`, install helpers, restart). Do that before testing Ctrl+Shift+Y.

**Canonical files**

| Priority | File | Role |
|---|---|---|
| 1 | `catalog/Strider_Media_Catalog.xlsm` | **Working copy** — open this for culling |
| 2 | `catalog/Strider_Media_Catalog_blank.xlsm` | **Template of record** — empty + macros |
| 3 | `catalog/build/Strider_Media_Catalog_blank.xlsm` | Mirror of the blank |
| — | `catalog/build/Strider_Media_Catalog.xlsx` | Skeleton only (no macros) |

Always open a `.xlsm`. Day-to-day guide: `README.md`.

---

## Morning checklist (do first)

1. **Quit Excel** if open.
2. Install helpers (adds `StriderCaption.scpt` + `StriderCaption.sh` — not installed yet as of last session):

   ```bash
   ./scripts/install_mac_helpers.sh
   ```

3. Open `catalog/Strider_Media_Catalog.xlsm`.
4. **Re-import macros:** VBE → remove old **CatalogMacros** → File → Import → `catalog/macros/CatalogMacros.bas` → Save.
5. Confirm Entry / VideoEntry still have SelectionChange snippets from the `.txt` files.
6. Restart Excel once after helper install; reopen the `.xlsm`.
7. Smoke-test:
   - Select a Photos row → **Ctrl+Shift+Y** → should create/update **Captions** sheet.
   - Captions sheet top: title (row 1), **prompt** (row 2, merged), hint (row 3), headers (row 4), data from row 5.
   - Default provider is **Gemini Flash** (`google`). Claude: set `DEFAULT_CAPTION_PROVIDER = "anthropic"` in CatalogMacros and re-import.

API keys: sibling project `../Virtual_Nexus/.env` (wired by `StriderCaption.sh`). Fallback: repo `.env` (gitignored).

---

## What landed this session

### AI captions (new)

| Piece | Role |
|---|---|
| `scripts/caption_photo.py` | CLI vision captioner; `--machine` for Excel; loads `caption_prompt.txt`; default provider **google** (Flash) |
| `catalog/macros/caption_prompt.txt` | Canonical prompt text (keep in sync with VBA `CaptionPromptText`) |
| `catalog/macros/StriderCaption.applescript` | Excel → shell bridge |
| `scripts/install_mac_helpers.sh` | Also installs `StriderCaption.scpt` + `StriderCaption.sh` |
| `CatalogMacros.bas` → `CaptionSelectedPhotos` | Ctrl+Shift+Y; writes Captions sheet |
| Captions sheet layout | Row 1 title · row 2 prompt · row 3 hint · **row 4 headers** · row 5+ data (`CAPTIONS_HEADER_ROW = 4`) |
| `SheetHeaderRow()` | Catalog helpers use header row 4 on Captions, row 1 elsewhere |

CLI examples:

```bash
python3 scripts/caption_photo.py /path/to/photo.JPG --plain
python3 scripts/caption_photo.py /path/to/photo.JPG --provider anthropic --plain
python3 scripts/caption_photo.py /path/to/photo.JPG --all
```

Prompt refinements already in place: ~60 words; trust photographer notes; count every visible strider; note patterns; no camera/settings dump.

### Tag write-back fix (critical)

`StriderFinderMeta.applescript` now accepts a **labeled** payload `key=value|…` (order-independent). Empty notes **clears** Description/Caption. Prevents Sequence* / caustic fields from shifting into wrong keyword slots.

### Caustic / Filters fields

Photos & Videos (+ Entry forms, Lists, build_catalog) include: **Caustic camera aim**, **Caustic appearance**, **Filters**. Live + blank `.xlsm` updated in repo.

### Sequences (already on main from earlier today)

`scripts/tag_sequences.py`, `list_sequences.py`, `add_sequence_fields.py` — Sequence / SequenceID / SequencePos via Keywords. SequencePos stored as **text** (e.g. `6/54`), not Excel fraction.

---

## Sheets

- **Entry** / **VideoEntry** — load, tag, append  
- **Photos** / **Videos** — catalogs  
- **Captions** — AI captions (created on first CaptionSelectedPhotos if missing)  
- **Lists** — click-pick vocab  
- **ReadMe** — in-workbook notes  

**Photos columns:** Name, Path, Date, Time, Sequence, SequenceID, SequencePos, Rating, Species, Behavior, Locale, Type, FileSize, Send to UK, View, Camera, Mode, Lens, Speed, f-stop, ISO, Lighting, Caustic camera aim, Caustic appearance, Filters, notes  

**Captions columns (row 4):** Name, Path, Provider, Model, Words, Generated, Caption  

---

## Shortcuts (Control+Shift)

| Shortcut | Action |
|---|---|
| Q / I | Load photo Finder / Preview |
| W | Append photo (+ write tags) |
| E / U | Load video Finder / Preview |
| R | Append video |
| O / J | Open Preview / Reveal Finder |
| D | Trash front Preview doc |
| **Y** | **CaptionSelectedPhotos** → Captions sheet |

Re-import `CatalogMacros.bas` after macro edits so `Auto_Open` rebinds keys.

---

## Scripts quick map

| Script | Touches live `.xlsm`? |
|---|---|
| `install_mac_helpers.sh` | No — only Application Scripts |
| `caption_photo.py` | No — files / stdout only |
| `tag_sequences.py` | No — folder you pass |
| `build_catalog.py` | No — `catalog/build/*.xlsx` |
| `add_sequence_fields.py` | **Yes** — open workbook |
| `setup_workbook.sh` | **Yes — overwrites** unless careful; needs `--force` |

Helpers path: `~/Library/Application Scripts/com.microsoft.Excel/`  
(`StriderFinderMeta`, `StriderOpenPreview`, `StriderRevealFinder`, **`StriderCaption`**, **`StriderCaption.sh`**)

---

## Prompt sync rule

Edit prompt in **both** places (or regenerate one from the other):

1. `catalog/macros/caption_prompt.txt` — used by `caption_photo.py`
2. `CaptionPromptText()` in `CatalogMacros.bas` — written into Captions!A2 by `EnsureCaptionsSheet`

After VBA prompt edits: re-import `.bas`. Existing Captions sheets already have the old prompt in A2 — re-run `EnsureCaptionsSheet` logic by calling CaptionSelectedPhotos once (it overwrites A2), or paste manually.

---

## Known quirks / watch list

- **Captions sheet may not exist** on the open working copy until first Ctrl+Shift+Y after re-import.
- Helpers **must** be reinstalled for captioning; last session confirmed only the three older `.scpt` files were present.
- Do **not** keep an **AppEvents** class module (breaks Mac Excel).
- Do **not** run `setup_workbook.sh` casually — overwrites working `.xlsm`.
- Append matches by **filename**; empty notes clears file caption.
- Claude sometimes undercounts tiny striders; Gemini Flash is the cheap default and was acceptable in testing.
- Test images used: `_DSC9895`, `_DSC0320` (notes + count), `_DSC7924` (notes incorporation).

---

## Suggested next steps (tomorrow)

1. Morning checklist above (import + helpers + smoke Ctrl+Shift+Y).  
2. Confirm Captions prompt is visible at top and data lands under row-4 headers.  
3. Optional: switch one caption to `anthropic` and compare quality vs Flash.  
4. Real Preview cull session (I → tag → W → D).  
5. If Captions layout looks good, ensure blank `.xlsm` includes a Captions sheet (or rely on EnsureCaptionsSheet) and mirror into `catalog/build/` if desired.  
6. Commit only if further edits — today’s work should already be on `origin/main`.

---

## Regenerating layout (rare)

```bash
./scripts/install_mac_helpers.sh
python3 scripts/build_catalog.py   # skeleton only → catalog/build/
```

Canonical empty template remains the hand-tuned **blank `.xlsm`**, not the build `.xlsx`.
