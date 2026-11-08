# Striders

Media catalog and notes for water strider biology, with a focus on behavior and caustics in photos and videos.

## Catalog workbook

[`catalog/Strider_Media_Catalog.xlsm`](catalog/Strider_Media_Catalog.xlsm)

Sheets:

- **Entry** — add photos (`LoadFromFinderSelection` → tag → `AppendPhotoRow`)
- **VideoEntry** — add videos (`LoadVideoFromFinderSelection` → tag → `AppendVideoRow`)
- **Photos** / **Videos** — catalogs
- **Lists** — click-to-pick vocabularies (edit here to change options)

**Photos columns:**  
Name, Path, Date, Time, Rating, Species, Behavior, Locale, Type, FileSize, Send to UK, View, Camera, Mode, Lens, Speed, f-stop, ISO, Lighting, notes

**Videos columns:**  
Name, Path, Date, Time, Type, FileSize, Duration, FrameRate, Resolution, Camera, Speed, f-stop, ISO, Rating, Species, Behavior, Locale, Send to UK, View, Lighting, notes

**View options:** surface, bottom, both, from below, eye level

**Always open the `.xlsm` file.**

### First-time setup

```bash
chmod +x scripts/*.sh
./scripts/setup_workbook.sh
./scripts/install_mac_helpers.sh
```

Then import macros:

1. Open `catalog/Strider_Media_Catalog.xlsm` and enable macros.
2. **Tools → Macro → Visual Basic…**
3. Delete old **CatalogMacros** / **AppEvents** if present (AppEvents breaks on Mac).
4. **File → Import File…** → `catalog/macros/CatalogMacros.bas`
5. One-click picks:
   - **Microsoft Excel Objects → Entry** → paste `catalog/macros/EntrySheet_SelectionChange.txt`
   - **Microsoft Excel Objects → VideoEntry** → paste `catalog/macros/VideoEntrySheet_SelectionChange.txt`
6. Save as `.xlsm`. Restart Excel after installing helpers.

Macros:

- `LoadFromFinderSelection` / `LoadFromPreviewSelection` / `AppendPhotoRow` (alias: `AppendCatalogRow`)
- `LoadVideoFromFinderSelection` / `LoadVideoFromPreviewSelection` / `AppendVideoRow`
- `TrashFrontPreviewDocument` — Preview’s own **Edit → Move Selected Image to Trash** (Ctrl+Shift+D)
- `PickSelectedOption` — click an option in C+, then run this to fill column B
- `OpenSelectedInPreview` / `RevealSelectedInFinder` (Photos or Videos)

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

If OnKey still does nothing on your Mac Excel build, use **Tools → Customize Keyboard…** → category **Macros** → assign the same keys (or ⌥ letter keys, which that dialog supports).

Requires [exiftool](https://exiftool.org/) (`brew install exiftool`).

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

**Tag write-back on Append:** Rating → `Rating`/`XMP:Rating`; notes → Description/Caption; Locale → Location; Species/Behavior/Locale/SendToUK/View/Lighting → Keywords / Subject (`Species:…`, `Behavior:…`, etc.). Modifies the original file in place.

### Review

1. On **Photos** or **Videos**, filter as needed.
2. Select **Name** / **Path** cells or rows.
3. Run **OpenSelectedInPreview** or **RevealSelectedInFinder**.
