# Striders

Media catalog and notes for water strider biology, with a focus on behavior and caustics in photos and videos.

## Catalog workbook

**Which file to open**

| File | Role |
|---|---|
| [`catalog/Strider_Media_Catalog.xlsm`](catalog/Strider_Media_Catalog.xlsm) | **Working copy** — day-to-day catalog (macros + your rows) |
| [`catalog/Strider_Media_Catalog_blank.xlsm`](catalog/Strider_Media_Catalog_blank.xlsm) | **Template of record** — empty starter with same layout/macros; copy this for a fresh catalog |
| `catalog/build/Strider_Media_Catalog.xlsx` | Build skeleton only (no macros) — do not use for daily work |
| `catalog/build/Strider_Media_Catalog_blank.xlsm` | Mirror of the blank template |

Always open a `.xlsm`. Session notes for agents: [`HANDOFF.md`](HANDOFF.md).

Sheets:

- **Entry** — add photos (`LoadFromFinderSelection` → tag → `AppendPhotoRow`)
- **VideoEntry** — add videos (`LoadVideoFromFinderSelection` → tag → `AppendVideoRow`)
- **Photos** / **Videos** — catalogs
- **Captions** — AI photo captions (prompt at top; Ctrl+Shift+Y from Photos). Default: Gemini Flash
- **Lists** — click-to-pick vocabularies (edit here to change options)
- **ReadMe** — short in-workbook notes

**Photos columns:**  
Name, Path, Date, Time, Sequence, SequenceID, SequencePos, Rating, Species, Behavior, Locale, Type, FileSize, Send to UK, View, Camera, Mode, Lens, Speed, f-stop, ISO, Lighting, Caustic camera aim, Caustic appearance, Filters, notes

**Videos columns:**  
Name, Path, Date, Time, Type, FileSize, Duration, FrameRate, Resolution, Camera, Speed, f-stop, ISO, Rating, Species, Behavior, Locale, Send to UK, View, Lighting, Caustic camera aim, Caustic appearance, Filters, notes

**View options:** surface, bottom, both, from below, eye level

### First-time setup

Prefer an existing `.xlsm` (working copy or blank template). Only rebuild from the `.xlsx` skeleton as a last resort.

```bash
chmod +x scripts/*.sh
./scripts/install_mac_helpers.sh
# Last resort — overwrites the working-copy .xlsm (refused unless --force):
# ./scripts/setup_workbook.sh --force
```

1. Open the **working copy** or a **copy of the blank template**; enable macros.
2. **Tools → Macro → Visual Basic…**
3. Delete old **CatalogMacros** / **AppEvents** if present (AppEvents breaks on Mac; there is no AppEvents source in `catalog/macros/` anymore).
4. **File → Import File…** → `catalog/macros/CatalogMacros.bas`
5. One-click picks:
   - **Microsoft Excel Objects → Entry** → paste `catalog/macros/EntrySheet_SelectionChange.txt`
   - **Microsoft Excel Objects → VideoEntry** → paste `catalog/macros/VideoEntrySheet_SelectionChange.txt`
6. Save as `.xlsm`. Restart Excel after installing helpers.

After layout/macro tweaks you care about, keep the blank `.xlsm` in sync (and optionally mirror it into `catalog/build/`).

Macros:

- `LoadFromFinderSelection` / `LoadFromPreviewSelection` / `AppendPhotoRow` (alias: `AppendCatalogRow`)
- `LoadVideoFromFinderSelection` / `LoadVideoFromPreviewSelection` / `AppendVideoRow`
- `TrashFrontPreviewDocument` — Preview’s own **Edit → Move Selected Image to Trash** (Ctrl+Shift+D)
- `PickSelectedOption` — click an option in C+, then run this to fill column B
- `OpenSelectedInPreview` / `RevealSelectedInFinder` (Photos or Videos)
- `CaptionSelectedPhotos` — Nat Geo–style captions for selected Photos rows → **Captions** sheet

Keyboard shortcuts (Control key, not ⌘). Old L/A/M/F combos conflict with Excel built-ins (filter, formula args, Copilot, fonts).

Installed by `AssignMacroShortcuts` / `Auto_Open` after re-importing `CatalogMacros.bas`:

| Shortcut | Macro |
|---|---|
| Control+Shift+Q | LoadFromFinderSelection |
| Control+Shift+I | LoadFromPreviewSelection (front Preview doc) |
| Control+Shift+W | AppendPhotoRow |
| Control+Shift+E | LoadVideoFromFinderSelection |
| Control+Shift+U | LoadVideoFromPreviewSelection |
| Control+Shift+R | AppendVideoRow |
| Control+Shift+O | OpenSelectedInPreview |
| Control+Shift+J | RevealSelectedInFinder |
| Control+Shift+D | TrashFrontPreviewDocument (via Preview menu) |
| Control+Shift+Y | CaptionSelectedPhotos (Gemini Flash by default) |

If OnKey still does nothing on your Mac Excel build, use **Tools → Customize Keyboard…** → category **Macros** → assign the same keys (or ⌥ letter keys, which that dialog supports).

Requires [exiftool](https://exiftool.org/) (`brew install exiftool`).

### Tag photo sequences (hi / hi+ / Continuous)

Before culling, tag a folder of photos by capture-time proximity. Skips non-continuous drive modes by default. Dry-run first:

```bash
python3 scripts/tag_sequences.py "/path/to/DCIM/folder"
python3 scripts/tag_sequences.py "/path/to/DCIM/folder" --write
```

Writes Keywords/Subject: `Sequence:yes|no`, `SequenceID:YYYYMMDD-HHMMSS`, `SequencePos:6/54`.  
Load into Entry picks these up; Append preserves them.

To add the Sequence columns to an existing open workbook:

```bash
python3 scripts/add_sequence_fields.py   # workbook must already be open in Excel
```

Then re-import `catalog/macros/CatalogMacros.bas` and reinstall helpers (`./scripts/install_mac_helpers.sh`).

### Add a photo

1. Select the file in **Finder**, *or* step to it in **Preview** (multi-file cull).
2. Run **LoadFromFinderSelection** (Ctrl+Shift+Q) or **LoadFromPreviewSelection** (Ctrl+Shift+I).
3. Click an option in columns **C+** (tags) — it copies into column B.
4. Run **AppendPhotoRow** (also writes tags into the original file via exiftool).

### Add a video

1. Select the video in **Finder**, or open it in **Preview**.
2. Run **LoadVideoFromFinderSelection** (Ctrl+Shift+E) or **LoadVideoFromPreviewSelection** (Ctrl+Shift+U).
3. Click tag options in columns **C+** on **VideoEntry**.
4. Run **AppendVideoRow** (also writes tags into the original file).

**Tag write-back on Append:** Rating → `Rating`/`XMP:Rating`; notes → Description/Caption; Locale → Location; Species/Behavior/Locale/SendToUK/View/Lighting/CausticAim/CausticAppearance/Filters/Sequence* → Keywords / Subject (`Species:…`, `CausticAim:…`, etc.). Modifies the original file in place.

### Review

1. On **Photos** or **Videos**, filter as needed.
2. Select **Name** / **Path** cells or rows.
3. Run **OpenSelectedInPreview** or **RevealSelectedInFinder**.
