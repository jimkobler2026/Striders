# Striders

Media catalog and field notes for water strider biology, with a focus on behavior and caustics in photos and videos.

The **Strider Media Catalog** is a Mac Excel (`.xlsm`) workbook wired to Finder, Preview, exiftool, and a few Python tools. It is “smart” in that cataloguing is not just typing into a spreadsheet: loading a file pulls camera metadata, click-pick tags write back into the image, sequences can be pre-tagged on disk, filenames can be made unique across card reformats, and selected photos can be AI-captioned into a Captions sheet.

Agent session notes: `[HANDOFF.md](HANDOFF.md)`.

---



## How the smart catalog works

### Big picture

```text
Camera card → temp ingest folder under ~/Pictures/environmental caustics
    → rename_media_unique.py   (unique basenames)
    → tag_sequences.py         (optional: Sequence* keywords)
    → Preview cull + Excel Entry/VideoEntry
         Load (Finder or Preview) → click-pick tags → Append
    → Photos / Videos sheets (+ Keywords written into the file)
    → optional CaptionSelectedPhotos → Captions sheet
```

Excel never shells out to arbitrary scripts by path alone. On Mac, VBA calls **AppleScriptTask** helpers installed under:

`~/Library/Application Scripts/com.microsoft.Excel/`

Those helpers run exiftool (and, for captions, a small shell bridge into `scripts/caption_photo.py`).

### What makes it “smart”

1. **Load fills the form from the file** — Date/Time, camera, exposure, and any Sequence* keywords already on the file.
2. **Click-pick vocabularies** — Options live on **Lists** and on Entry/VideoEntry columns C+; clicking an option copies into the value cell (column B, or D/F for side-by-side auto fields).
3. **Append does two things** — Adds/updates a row on **Photos** or **Videos**, and writes tagged metadata back into the media file (Rating, Description NOTE section, Location, Keywords).
4. **Fields are discovered at runtime** — Macros read Entry/VideoEntry labels (columns A/C/E) and Lists headers. New click-pick tags do not require editing VBA or AppleScript key lists (after the one-time dynamic-field upgrade).
5. **Append matches by filename** — Re-appending the same basename updates that catalog row. That is why ingest folders are renamed before cataloguing when camera names recycle.
6. **Captions are optional and on demand** — Select Photos rows → Ctrl+Shift+Y; writes the Captions sheet and embeds `CAPTION:` into `MWG:Description` (preserves any `NOTE:`).



### Recommended ingest loop

1. Copy card contents into an ingest folder under `~/Pictures/environmental caustics/…`.
2. Rename for uniqueness (`rename_media_unique.py --recursive --write`).
3. Optionally tag continuous bursts (`tag_sequences.py --recursive --write`; default minimum **4** frames).
4. Open a **copy** of the blank `.xlsm` (your working catalog).
5. Cull in Preview: **I** load → tag → **W** append → **D** trash reject (photos). Videos: **U** / **E** → tag → **R**.
6. **Sequences:** load/tag the **1/n** frame on Entry, then **Ctrl+Shift+S** (`AppendPhotoSequence`) to import every frame with those tags (each keeps its own SequencePos).
7. Later: filter on Photos/Videos; **O** / **J** to reopen or reveal; **Y** to caption keepers.

---



## Sequence tagging and bulk import

Two separate steps: **tag bursts on disk**, then **bulk-append into Excel** using the first frame’s user tags.

### 1. Tag sequences on disk (`tag_sequences.py`)

Scans an ingest folder for Continuous / Hi / Hi+ frames and writes Keywords. Burst boundaries prefer the camera’s `SequenceNumber` (resets to **1** at each new burst on Sony A7RIV); capture-time `--gap` (default 1 s) is used only when that tag is absent. Order inside a burst follows the original camera number in the filename (`…_DSC7659`).

| Keyword | Example | Meaning |
|---------|---------|---------|
| `Sequence` | `yes` or `no` | In a qualifying burst, or Continuous but too short / alone |
| `SequenceID` | `20261102-091557` | Shared id for the burst (usually start timestamp) |
| `SequencePos` | `3/22` | Position within the burst (stored as **text**, not an Excel fraction) |

Rules that matter in practice:

- Only Continuous/Hi/Hi+ are eligible; single-shot and phone ingest folders usually stay untagged.
- Default **`--min-frames 4`**: bursts of 2–3 get `Sequence:no` (no ID/pos).
- JPEG+RAW pairs at the same instant count as two frames toward the minimum.
- Dry-run by default; add `--write` (and `--force` to rewrite existing Sequence* tags).
- `--write` keeps one exiftool process open (`-stay_open`) so large ingest folders are not one-spawn-per-file.
- Use `--recursive` for nested DCIM-style ingest folders.

```bash
# One ingest folder
python3 scripts/tag_sequences.py "~/Pictures/environmental caustics/my-shoot" --recursive --write

# All ingest folders under environmental caustics
for d in ~/Pictures/environmental\ caustics/*/; do
  python3 scripts/tag_sequences.py "$d" --recursive --write --force
done
```

Inspect afterward with `list_sequences.py` (CSV of SequenceID / bounds / counts).

### 2. Bulk import into the catalog (`Ctrl+Shift+S`)

Goal: tag the representative frame once, then append every sibling with the same user metadata.

1. Run `tag_sequences.py --write` on the ingest folder first (so SequenceID / SequencePos exist).
2. In Preview (or Finder), open the **1/n** frame of a sequence.
3. **Ctrl+Shift+I** (or **Q**) to load it into Entry — Sequence* fields fill from the file.
4. Click-pick Species, Behavior, Locale, caustic fields, notes, etc. on that one frame.
5. **Ctrl+Shift+S** (`AppendPhotoSequence`):
   - Finds every file in the same folder with that SequenceID (`list_sequence_members.py` via `StriderSequence.sh`).
   - Appends each to Photos (updates if the basename already exists).
   - Copies **your Entry user tags** onto every frame.
   - Keeps each frame’s own camera EXIF and its own SequencePos (`1/n`, `2/n`, …).
6. Continue culling; use **W** for singles, **S** whenever Entry is sitting on a tagged sequence member.

Prerequisites: helpers installed (`./scripts/install_mac_helpers.sh`), Excel restarted, current `CatalogMacros.bas` imported so **S** is bound.

---



## Workbooks and sheets



### Which Excel file to use


| File / folder                                                                              | Role                                                                                                        |
| ------------------------------------------------------------------------------------------ | ----------------------------------------------------------------------------------------------------------- |
| `[catalog/build/Strider_Media_Catalog_blank.xlsm](catalog/build/Strider_Media_Catalog_blank.xlsm)` | **Template of record** — empty starter with layout + macros. **Duplicate this** to start a working catalog. |
| `catalog/main catalog/Strider_Media_Catalog.xlsm`                                          | Working catalog — day-to-day use (tracked).                                                                 |
| `catalog/backup of blank catalog/`                                                         | Optional local snapshots of the blank (gitignored).                                                         |


Always open a `.xlsm`. Do not use a build `.xlsx` skeleton for daily work (no macros). Prefer duplicating the blank over running `setup_workbook.sh` (that script can overwrite a working catalog).

### Sheets


| Sheet          | Purpose                                                                                                         |
| -------------- | --------------------------------------------------------------------------------------------------------------- |
| **Entry**      | Photo staging form: load metadata, click-pick tags, append to Photos.                                           |
| **VideoEntry** | Same for videos → Videos.                                                                                       |
| **Photos**     | Photo catalog (one row per file).                                                                               |
| **Videos**     | Video catalog.                                                                                                  |
| **Lists**      | Vocabularies for click-pick tags (header row 2; options from row 3).                                            |
| **Captions**   | AI captions (created on first Ctrl+Shift+Y). Row 1 title, row 2 prompt, row 3 hint, row 4 headers, row 5+ data. |
| **ReadMe**     | Short in-workbook notes.                                                                                        |


**Entry / VideoEntry layout notes**

- Instructions and a Ctrl+Shift cheat sheet sit above the fields.
- Auto/camera fields may use **A/B**, **C/D**, and **E/F** on the same row (e.g. Name | Camera | Serial Number).
- Current Entry auto block (blank template): Name/Path/Date/Time/Type/FileSize with Camera/Mode/Lens/Speed/f-stop/ISO beside them; **Serial Number** at **E12/F12**; **Focus Distance** | **GPS** on the row below FileSize/ISO; then Sequence / SequenceID / SequencePos; a blank spacer row; then user tags.
- Sequence fields may share one row: Sequence (A/B), SequenceID (C/D), SequencePos (E/F).
- User tags use label in **A**, value in **B**, options in **C+**.
- `Note` on Entry maps to the Photos/Videos column `notes`, and on Append merges into `MWG:Description` as `NOTE:` (see Description layout above).
- **Metadata written to file on Append** is a live preview of what Append will write. `Description → NOTE:` tracks the Entry Note; `Description → CAPTION:` is filled on Load (Append preserves any CAPTION already on the file; Ctrl+Shift+Y sets it).

Typical **Photos** columns include: Name, Path, Date, Time, Sequence, SequenceID, SequencePos, Rating, Species, Behavior, Locale, Type, FileSize, Send to UK, View, Camera, Mode, Lens, Speed, f-stop, ISO, Focus Distance, GPS, Serial Number, Lighting, Caustic camera aim, Caustic appearance, Filters, Panopticon portfolio, notes (plus any fields you add). Focus Distance, GPS, and Serial Number are read-only Load fields (not Keyword write-back); Load never overwrites Locale from GPS.

---



## Keyboard shortcuts

Control key, not ⌘. Bound by `Auto_Open` → `AssignMacroShortcuts` after `CatalogMacros.bas` is imported.


| Shortcut         | Macro                           | What it does                                    |
| ---------------- | ------------------------------- | ----------------------------------------------- |
| Ctrl+Shift+**Q** | `LoadFromFinderSelection`       | Load Finder-selected photo into Entry           |
| Ctrl+Shift+**I** | `LoadFromPreviewSelection`      | Load front Preview photo into Entry             |
| Ctrl+Shift+**W** | `AppendPhotoRow`                | Append Entry → Photos + write tags to file      |
| Ctrl+Shift+**E** | `LoadVideoFromFinderSelection`  | Load Finder-selected video into VideoEntry      |
| Ctrl+Shift+**U** | `LoadVideoFromPreviewSelection` | Load front Preview video into VideoEntry        |
| Ctrl+Shift+**R** | `AppendVideoRow`                | Append VideoEntry → Videos + write tags to file |
| Ctrl+Shift+**O** | `OpenSelectedInPreview`         | Open selected Photos/Videos paths in Preview    |
| Ctrl+Shift+**J** | `RevealSelectedInFinder`        | Reveal selected paths in Finder                 |
| Ctrl+Shift+**D** | `TrashFrontPreviewDocument`     | Preview → Move Selected Image to Trash          |
| Ctrl+Shift+**Y** | `CaptionSelectedPhotos`         | Caption → Captions sheet + embed `CAPTION:`     |
| Ctrl+Shift+**K** | `AddCatalogField`               | Prompt to add a new click-pick tag everywhere   |
| Ctrl+Shift+**S** | `AppendPhotoSequence`           | Append all frames of Entry’s SequenceID (tags from 1/n) |


If OnKey does nothing on your Excel build: **Tools → Customize Keyboard…** → Macros → assign the same keys.

After **re-importing** `CatalogMacros.bas` while the workbook is already open, OnKey bindings are not refreshed automatically (`Auto_Open` only runs on open). Run macro **`AssignMacroShortcuts`** once, or close and reopen the `.xlsm`.

---



## Macros (`catalog/macros/CatalogMacros.bas`)

Import this module into the VBA project (remove any old **CatalogMacros** / **AppEvents** first; AppEvents breaks Mac Excel).


| Macro                                                            | Role                                                                             |
| ---------------------------------------------------------------- | -------------------------------------------------------------------------------- |
| `Auto_Open` / `Auto_Close`                                       | Enable events; assign / clear OnKey shortcuts                                    |
| `AssignMacroShortcuts` / `ClearMacroShortcuts`                   | Bind or clear Ctrl+Shift keys                                                    |
| `FixEnableEvents`                                                | Re-enable sheet events if click-pick dies                                        |
| `HandleEntryOptionClick` / `PickSelectedOption`                  | Click-pick: option cell → value cell (Lists-driven)                              |
| `LoadFromFinderSelection` / `LoadFromPreviewSelection`           | Photo load via AppleScript + exiftool                                            |
| `LoadVideoFromFinderSelection` / `LoadVideoFromPreviewSelection` | Video load                                                                       |
| `AppendPhotoRow` / `AppendCatalogRow`                            | Write Photos row + tag write-back                                                |
| `AppendVideoRow`                                                 | Write Videos row + tag write-back                                                |
| `TrashFrontPreviewDocument`                                      | Reject current Preview document                                                  |
| `OpenSelectedInPreview` / `RevealSelectedInFinder`               | Review helpers                                                                   |
| `CaptionSelectedPhotos`                                          | Batch AI captions for selection                                                  |
| `AddCatalogField`                                                | Add Lists column, Entry/VideoEntry rows, Photos/Videos columns, metadata preview |
| `AppendPhotoSequence`                                            | Bulk-append every photo sharing Entry’s SequenceID; shared tags from Entry (use 1/n); per-frame EXIF + SequencePos |


**Runtime field discovery**

- Form field lists come from Entry/VideoEntry labels in columns A/C/E (instructions and “Metadata written…” rows are skipped).
- Click-pick fields are those whose label matches a **Lists** header.
- Tag write-back sends labeled `key=value|…` payloads; Keyword keys use legacy aliases where needed (`Send to UK` → `SendToUK`, caustic fields, `Note`/`notes` → `notes`).

**Sheet event stubs** (paste into each sheet’s code module):

- `[catalog/macros/EntrySheet_SelectionChange.txt](catalog/macros/EntrySheet_SelectionChange.txt)`
- `[catalog/macros/VideoEntrySheet_SelectionChange.txt](catalog/macros/VideoEntrySheet_SelectionChange.txt)`

---



## AppleScript helpers (source + installed copies)



### Source (repo)


| File                                                                                               | Role                                                                                                         |
| -------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------ |
| `[catalog/macros/StriderFinderMeta.applescript](catalog/macros/StriderFinderMeta.applescript)`     | Finder/Preview paths; photo/video exiftool metadata; trash front Preview; **generic** catalog tag write-back |
| `[catalog/macros/StriderOpenPreview.applescript](catalog/macros/StriderOpenPreview.applescript)`   | Open paths in Preview                                                                                        |
| `[catalog/macros/StriderRevealFinder.applescript](catalog/macros/StriderRevealFinder.applescript)` | Reveal paths in Finder                                                                                       |
| `[catalog/macros/StriderCaption.applescript](catalog/macros/StriderCaption.applescript)`           | Bridge Excel → `StriderCaption.sh` → `caption_photo.py`                                                      |




### Installed location (required for macros)

`~/Library/Application Scripts/com.microsoft.Excel/`


| Installed file             | From                                                     |
| -------------------------- | -------------------------------------------------------- |
| `StriderFinderMeta.scpt`   | `StriderFinderMeta.applescript`                          |
| `StriderOpenPreview.scpt`  | `StriderOpenPreview.applescript`                         |
| `StriderRevealFinder.scpt` | `StriderRevealFinder.applescript`                        |
| `StriderCaption.scpt`      | `StriderCaption.applescript`                             |
| `StriderCaption.sh`        | Generated by `install_mac_helpers.sh` (paths + provider) |
| `StriderSequence.sh`       | Generated by `install_mac_helpers.sh` (list sequence members) |


Install / refresh:

```bash
./scripts/install_mac_helpers.sh
```

Then restart Excel.

**Tag write-back behavior (on Append):** Uses exiftool **MWG** composites so overlapping namespaces stay in sync — Rating → `MWG:Rating` (XMP); Locale → `MWG:Location` (+ Keywords); other tags → `MWG:Keywords` (`Key:value`, IPTC+XMP). Blank Sequence* on the form preserves Sequence* already on the file. Writes use **`-P`** so Finder Date Modified / file create times are not bumped to “now” (`tag_sequences.py` likewise).

**`MWG:Description` layout (NOTE / CAPTION):** One Description field may hold optional sections:

```text
NOTE: field notes from Entry
CAPTION: Nat Geo–style AI caption
```

- **Ctrl+Shift+W (Append):** merges the Entry Note into the `NOTE:` section; preserves any existing `CAPTION:`. Empty Note removes `NOTE:` only (does not clear `CAPTION:`). Non-empty Note is also written as a `notes:…` Keyword.
- **Ctrl+Shift+Y (Caption):** merges the AI caption into the `CAPTION:` section; preserves any existing `NOTE:`; still writes the Captions sheet.
- **Load:** Entry Note is filled from the `NOTE:` body (or from legacy unprefixed Description text). CAPTION text stays on the file only.
- Sections may appear in either order; either may be absent. Delimiters are the literal prefixes `NOTE:` and `CAPTION:` (case-sensitive).

---



## Python and shell scripts (`scripts/`)


| Script                                                     | Touches workbook?             | Purpose                                                                                                                                         |
| ---------------------------------------------------------- | ----------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| `[install_mac_helpers.sh](scripts/install_mac_helpers.sh)` | No                            | Compile/install AppleScript helpers + caption shell bridge                                                                                      |
| `[rename_media_unique.py](scripts/rename_media_unique.py)` | No                            | Prefix `YYYYMMDD-HHMMSS_` onto media under `~/Pictures/environmental caustics` (dry-run default; `--write`; `--recursive` for DCIM-style ingest folders) |
| `[tag_sequences.py](scripts/tag_sequences.py)`             | No                            | Tag Continuous/Hi/Hi+ bursts via `SequenceNumber` (gap fallback); `--min-frames` default **4**; writes Sequence* Keywords via one exiftool `-stay_open` session |
| `[list_sequences.py](scripts/list_sequences.py)`           | No                            | CSV of SequenceID / start / end / frame count from Keywords                                                                                     |
| `[list_sequence_members.py](scripts/list_sequence_members.py)` | No                        | Paths + SequencePos for one SequenceID (used by AppendPhotoSequence)                                                                            |
| `[find_duplicate_media.py](scripts/find_duplicate_media.py)` | No                          | Report ImageDataHash content duplicates (report only; metadata-independent)                                                                   |
| `[copy_sequence_tags.py](scripts/copy_sequence_tags.py)`   | No                            | Copy user MWG tags from one sequence frame onto SequenceID siblings (dry-run default)                                                           |
| `[caption_photo.py](scripts/caption_photo.py)`             | No*                           | Vision caption CLI (`--machine` for Excel; `--write` embeds `CAPTION:` into Description). Prefers embedded `PreviewImage` when present          |
| `[exif_datetime.py](scripts/exif_datetime.py)`             | No                            | Shared DateTimeOriginal + OffsetTimeOriginal parsing for rename / sequence tools                                                                |
| `[geotag_ingest.py](scripts/geotag_ingest.py)`             | No                            | Optional GPX geotag for mirrorless ingest folders (dry-run default; destructive to GPS)                                                         |
| `[stamp_ownership.py](scripts/stamp_ownership.py)`         | No                            | Optional MWG Creator/Copyright stamp via `STRIDER_ARTIST` / `STRIDER_COPYRIGHT`                                                                 |
| `[test_tag_writeback.py](scripts/test_tag_writeback.py)`   | No                            | Dev smoke: copy file, MWG write payload, `exiftool -diff` assertions                                                                            |
| `[add_catalog_field.py](scripts/add_catalog_field.py)`     | **Yes** (open book)           | CLI alternative to Ctrl+Shift+K; also patches `build_catalog.py` / blank when possible                                                          |
| `[add_sequence_fields.py](scripts/add_sequence_fields.py)` | **Yes** (open book)           | One-time insert of Sequence columns into an older workbook                                                                                      |
| `[build_catalog.py](scripts/build_catalog.py)`             | Writes `catalog/build/*.xlsx` | Rebuild spreadsheet **skeleton** only (no macros)                                                                                               |
| `[setup_workbook.sh](scripts/setup_workbook.sh)`           | **Dangerous**                 | Cold rebuild of a `.xlsm` shell; refuses overwrite unless `--force`                                                                             |
| `[apply_video_sheets.py](scripts/apply_video_sheets.py)`   | **Yes** (open book)           | Legacy helper to add/fix VideoEntry + Videos (usually already present)                                                                          |


**Rare / maintenance AppleScripts in** `scripts/`


| File                               | Purpose                       |
| ---------------------------------- | ----------------------------- |
| `setup_xlsm.applescript`           | Used by cold workbook setup   |
| `upgrade_entry_layout.applescript` | One-off layout upgrade helper |




### Example commands

```bash
# Unique names (nested DCIM ingest folders)
python3 scripts/rename_media_unique.py "~/Pictures/environmental caustics/my-shoot" --recursive --write

# Sequences (≥4 frames)
python3 scripts/tag_sequences.py "~/Pictures/environmental caustics/my-shoot" --recursive --write

# Captions from Terminal
python3 scripts/caption_photo.py /path/to/photo.JPG --plain
python3 scripts/caption_photo.py /path/to/photo.JPG --provider anthropic --plain
python3 scripts/caption_photo.py /path/to/photo.JPG --plain --write   # also embed CAPTION: into the file
```

API keys for captions: prefer sibling `../Virtual_Nexus/.env` (wired by `StriderCaption.sh`); fallback repo `.env` (gitignored). Providers: `google` (Gemini Flash, default), `anthropic`, `openai`.

---



## Ancillary files and locations



### In the repo


| Path                                                 | What it is                                                                              |
| ---------------------------------------------------- | --------------------------------------------------------------------------------------- |
| `catalog/build/Strider_Media_Catalog_blank.xlsm`     | Canonical empty catalog template (template of record)                                   |
| `catalog/main catalog/Strider_Media_Catalog.xlsm`    | Working catalog (tracked)                                                               |
| `catalog/backup of blank catalog/`                   | Optional blank snapshots (gitignored)                                                   |
| `catalog/macros/CatalogMacros.bas`                   | VBA source of record                                                                    |
| `catalog/macros/*.applescript`                       | Helper sources                                                                          |
| `catalog/macros/EntrySheet_SelectionChange.txt`      | Entry sheet event stub                                                                  |
| `catalog/macros/VideoEntrySheet_SelectionChange.txt` | VideoEntry sheet event stub                                                             |
| `catalog/macros/caption_prompt.txt`                  | Canonical caption prompt (keep in sync with VBA `CaptionPromptText` if both are edited) |
| `scripts/*`                                          | Tools listed above                                                                      |
| `HANDOFF.md`                                         | Agent/session handoff notes                                                             |
| `.env` (optional, gitignored)                        | Local API keys fallback                                                                 |




### Outside the repo (machine-local)


| Path                                                                  | What it is                                      |
| --------------------------------------------------------------------- | ----------------------------------------------- |
| `~/Library/Application Scripts/com.microsoft.Excel/Strider*.scpt`     | Compiled helpers Excel may call                 |
| `~/Library/Application Scripts/com.microsoft.Excel/StriderCaption.sh` | Caption bridge script                           |
| `~/Pictures/environmental caustics/`                                  | Allowed root for `rename_media_unique.py` ingest folders |
| `../Virtual_Nexus/.env`                                               | Preferred API key file for captions             |




### Dependencies

- Mac Microsoft Excel with macros enabled  
- [exiftool](https://exiftool.org/) (`brew install exiftool`)  
- Python 3 (scripts); `xlwings` for scripts that drive an open workbook; vision API access for captions

---



## First-time setup

1. Duplicate `catalog/build/Strider_Media_Catalog_blank.xlsm` → e.g. `catalog/main catalog/Strider_Media_Catalog.xlsm`.
2. Install helpers and restart Excel:
  ```bash
   chmod +x scripts/*.sh
   ./scripts/install_mac_helpers.sh
  ```
3. Open the working `.xlsm`; enable macros.
4. If macros are missing or outdated: VBE → remove old **CatalogMacros** / **AppEvents** → Import `catalog/macros/CatalogMacros.bas` → paste Entry/VideoEntry SelectionChange stubs → Save.
5. Smoke-test: Load (I/Q), click-pick, Append (W), AddCatalogField (K), Caption (Y), sequence bulk Append (S).



### Adding a catalog tag later

Prefer **Ctrl+Shift+K** inside Excel (updates Lists, Entry, VideoEntry, Photos, Videos, and the metadata preview). Existing catalog rows stay blank in the new column.

### Safety

- Do **not** casually run `setup_workbook.sh` — it can overwrite a working `.xlsm` (requires `--force`).
- Deleting Photos/Videos **rows** only removes catalog entries, not files on disk. Captions rows are not auto-deleted with them.
- Keep `catalog/build/Strider_Media_Catalog_blank.xlsm` updated when Entry layout or macros change.

