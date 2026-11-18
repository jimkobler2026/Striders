# Workflow: image sequence → dual-speed GIF

Efficient path for strider (and similar) bursts. Goal: good motion GIF at roughly **~0.45–1 MB per frame** in the final file. **No picture-in-picture / insert overlays.**

Requires: **Fiji**, **gifsicle** (`brew install gifsicle`), **Pillow**, repo script `scripts/make_dual_interval_gif.py`.

---

## A. Find sequences

For a selected folder, list bursts as a text file:

- first frame filename  
- last frame filename  
- number of frames  

Repo helpers (SequenceNumber / Keywords): `scripts/tag_sequences.py`, `scripts/list_sequences.py`, `scripts/list_sequence_members.py`. A dedicated “folder → text report” wrapper can be added if needed.

## B. Cull in Preview

Open the range; keep only frames worth animating. Trim start/end as needed.

## C. Process

1. **Crop 3:2** to the **minimum** window that still encompasses the action (full-res Sony is typically **9504×6336**; example crop used in testing: **6000×4000**).
2. **Scale to 2400×1600** in place (same filenames). Keeps Fiji align fast; working JPEGs land near ~1 MB each.
3. **Fiji:** import as stack → **align** → **recrop** 3:2 if needed → save stack (TIFF or image sequence).  
   Do **not** bother resizing to 1200×800 in Fiji — the GIF script does that.
4. **Cursor / gifsicle** via `make_dual_interval_gif.py` (or equivalent):
   - normalize frames to fixed **1200×800** (3:2)
   - build **0.1 s** and **1.0 s** GIFs
   - concatenate playback, e.g. **3× fast + 1× slow**, **loop forever**
   - optimize: **`--colors 256 --lossy=40 -O3`**

Example:

```bash
python3 scripts/make_dual_interval_gif.py "/path/to/aligned_crop_stack.tif"
# defaults: --width 1200 → 1200×800, --fast-delay 0.1, --slow-delay 1.0,
#           --fast-repeats 3, --lossy 40
```

Outputs land next to the input (interval GIFs + combined + lossy deliverable).

---

## Size notes

| Stage | Typical size |
|---|---|
| Working JPEG after step 2 (2400×1600) | ~0.9–1 MB per frame |
| Final GIF at 1200×800 + lossy40 | ~**0.45 MB per frame** (e.g. 40 frames → ~18 MB) |
| 2400px GIF with little/no lossy | ~2 MB+ per frame — avoid for sharing |

Fiji does **not** expose gifsicle’s `--lossy` / `-O3`; keep size reduction in the script/gifsicle step.

---

## Why this order

- Aligning at full resolution is very slow; **2400×1600 before Fiji** cuts pixel count ~15× vs 9504-wide.
- Final **1200×800 + lossy40** is the quality/size sweet spot we validated.
- Dual timing (fast ×3, then slow ×1) shows motion then lets you study frames.
