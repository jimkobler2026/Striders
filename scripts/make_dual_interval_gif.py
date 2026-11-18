#!/usr/bin/env python3
"""Build a looping dual-speed GIF from a Fiji cropped image stack.

Pipeline:
  1. Load frames (Fiji TIFF stack, GIF, or a folder of images)
  2. Center-crop to exact 3:2, then scale to a fixed size (default 1200×800)
  3. Write .1s and 1.0s GIFs
  4. Concatenate: fast ×3, then slow ×1; loop forever
  5. Apply gifsicle lossy optimize (default lossy=40)
  6. Save all outputs next to the input

Requires: gifsicle (brew install gifsicle), Pillow.

Examples:
  python3 scripts/make_dual_interval_gif.py "/path/to/aligned_crop.tif"
  python3 scripts/make_dual_interval_gif.py "/path/to/stack.gif" --width 1200 --lossy 40
  python3 scripts/make_dual_interval_gif.py "/path/to/frames_folder"
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

from PIL import Image

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".gif", ".webp", ".bmp"}
ASPECT_W = 3
ASPECT_H = 2


def find_gifsicle() -> str:
    path = shutil.which("gifsicle")
    if not path:
        raise SystemExit("gifsicle not found. Install with: brew install gifsicle")
    return path


def natural_key(path: Path):
    return [int(p) if p.isdigit() else p.lower() for p in re.split(r"(\d+)", path.name)]


def load_frames(source: Path) -> list[Image.Image]:
    if source.is_dir():
        files = sorted(
            (p for p in source.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_EXTS),
            key=natural_key,
        )
        if not files:
            raise SystemExit(f"No images found in folder: {source}")
        frames = [Image.open(p).convert("RGB") for p in files]
        for im in frames:
            im.load()
        return frames

    if not source.is_file():
        raise SystemExit(f"Input not found: {source}")

    suffix = source.suffix.lower()
    if suffix in {".tif", ".tiff", ".gif"}:
        im = Image.open(source)
        frames: list[Image.Image] = []
        index = 0
        while True:
            try:
                im.seek(index)
            except EOFError:
                break
            frames.append(im.convert("RGB"))
            index += 1
        if not frames:
            raise SystemExit(f"No frames readable in: {source}")
        return frames

    if suffix in IMAGE_EXTS:
        return [Image.open(source).convert("RGB")]

    raise SystemExit(f"Unsupported input type: {source}")


def output_size(width: int) -> tuple[int, int]:
    """Fixed 3:2 frame size from width (height = width * 2 / 3)."""
    if width < ASPECT_W:
        raise SystemExit(f"--width must be >= {ASPECT_W}")
    # Exact 3:2 integer pixels: width divisible by 3 keeps height even/clean.
    width = width - (width % ASPECT_W)
    height = width * ASPECT_H // ASPECT_W
    return width, height


def fit_frame_3x2(im: Image.Image, out_w: int, out_h: int) -> Image.Image:
    """Center-crop to 3:2, then scale to exact out_w×out_h."""
    src_w, src_h = im.size
    target_aspect = out_w / out_h
    src_aspect = src_w / src_h
    if src_aspect > target_aspect:
        # Too wide: crop left/right
        crop_h = src_h
        crop_w = max(1, round(src_h * target_aspect))
    else:
        # Too tall (or equal): crop top/bottom
        crop_w = src_w
        crop_h = max(1, round(src_w / target_aspect))
    left = max(0, (src_w - crop_w) // 2)
    top = max(0, (src_h - crop_h) // 2)
    cropped = im.crop((left, top, left + crop_w, top + crop_h))
    if cropped.size == (out_w, out_h):
        return cropped
    return cropped.resize((out_w, out_h), Image.Resampling.LANCZOS)


def delay_cs(seconds: float) -> int:
    """gifsicle --delay unit is centiseconds (1/100 s)."""
    cs = int(round(seconds * 100))
    if cs < 1:
        raise SystemExit(f"Delay too small: {seconds}s")
    return cs


def run(cmd: list[str]) -> None:
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "").strip()
        raise SystemExit(f"Command failed ({proc.returncode}): {' '.join(cmd)}\n{err}")


def stem_label(source: Path) -> str:
    if source.is_dir():
        return source.name
    return source.stem


def build_outputs(
    source: Path,
    out_dir: Path,
    out_w: int,
    out_h: int,
    fast_s: float,
    slow_s: float,
    lossy: int,
    fast_repeats: int,
) -> dict[str, Path]:
    fast_label = f"{fast_s:g}s"
    slow_label = f"{slow_s:g}s"
    base = stem_label(source)
    combo = f"{fast_repeats}x {fast_label} then {slow_label}"
    size_tag = f"{out_w}x{out_h}"
    return {
        "fast": out_dir / f"{base} {fast_label} interval {size_tag}.gif",
        "slow": out_dir / f"{base} {slow_label} interval {size_tag}.gif",
        "combined": out_dir / f"{base} combined {combo} {size_tag}.gif",
        "final": out_dir / f"{base} combined {combo} {size_tag} lossy{lossy}.gif",
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Normalize a Fiji stack to fixed 3:2 and build looping fast×N + slow GIF."
    )
    parser.add_argument(
        "input",
        type=Path,
        help="Fiji TIFF/GIF stack, or a folder of cropped frames",
    )
    parser.add_argument(
        "--width",
        type=int,
        default=1200,
        help="Output width in pixels; height is 2/3 of width for 3:2 (default 1200 → 1200×800)",
    )
    parser.add_argument("--fast-delay", type=float, default=0.1, help="Fast GIF delay seconds (default 0.1)")
    parser.add_argument("--slow-delay", type=float, default=1.0, help="Slow GIF delay seconds (default 1.0)")
    parser.add_argument(
        "--fast-repeats",
        type=int,
        default=3,
        help="Play the fast GIF this many times before the slow GIF (default 3)",
    )
    parser.add_argument("--lossy", type=int, default=40, help="gifsicle --lossy for final GIF (default 40)")
    parser.add_argument("--colors", type=int, default=256, help="GIF palette size (default 256)")
    parser.add_argument(
        "--no-parts",
        action="store_true",
        help="Only write the final optimized combined GIF (skip .1s / 1s / raw combined)",
    )
    args = parser.parse_args()

    if args.fast_repeats < 1:
        raise SystemExit("--fast-repeats must be >= 1")

    out_w, out_h = output_size(args.width)
    source = args.input.expanduser().resolve()
    gifsicle = find_gifsicle()
    out_dir = source if source.is_dir() else source.parent
    outputs = build_outputs(
        source,
        out_dir,
        out_w,
        out_h,
        args.fast_delay,
        args.slow_delay,
        args.lossy,
        args.fast_repeats,
    )

    print(f"Input:   {source}")
    print(f"Out dir: {out_dir}")
    print(
        f"Frame:   {out_w}x{out_h} (3:2)  fast={args.fast_delay:g}s ×{args.fast_repeats}  "
        f"slow={args.slow_delay:g}s ×1  lossy={args.lossy}"
    )

    frames = load_frames(source)
    print(f"Frames:  {len(frames)}  ({frames[0].width}x{frames[0].height})")
    fitted = [fit_frame_3x2(im, out_w, out_h) for im in frames]
    print(f"Fitted:  {fitted[0].width}x{fitted[0].height}")

    fast_cs = delay_cs(args.fast_delay)
    slow_cs = delay_cs(args.slow_delay)

    with tempfile.TemporaryDirectory(prefix="dual_gif_") as tmp:
        tmp_path = Path(tmp)
        frame_paths: list[Path] = []
        for i, im in enumerate(fitted):
            path = tmp_path / f"frame_{i:04d}.png"
            im.save(path, format="PNG")
            frame_paths.append(path)

        fast_tmp = tmp_path / "fast.gif"
        slow_tmp = tmp_path / "slow.gif"
        combined_tmp = tmp_path / "combined.gif"
        final_tmp = tmp_path / "final.gif"

        common = [gifsicle, f"--colors={args.colors}", "--loopcount=0"]
        run(common + [f"--delay={fast_cs}", "-O3", "-o", str(fast_tmp), *[str(p) for p in frame_paths]])
        run(common + [f"--delay={slow_cs}", "-O3", "-o", str(slow_tmp), *[str(p) for p in frame_paths]])
        concat_inputs = [str(fast_tmp)] * args.fast_repeats + [str(slow_tmp)]
        run([gifsicle, "--loopcount=0", *concat_inputs, "-o", str(combined_tmp)])
        run(
            [
                gifsicle,
                "--loopcount=0",
                f"--colors={args.colors}",
                f"--lossy={args.lossy}",
                "-O3",
                str(combined_tmp),
                "-o",
                str(final_tmp),
            ]
        )

        if not args.no_parts:
            shutil.copy2(fast_tmp, outputs["fast"])
            shutil.copy2(slow_tmp, outputs["slow"])
            shutil.copy2(combined_tmp, outputs["combined"])
        shutil.copy2(final_tmp, outputs["final"])

    print("Wrote:")
    if not args.no_parts:
        for key in ("fast", "slow", "combined"):
            path = outputs[key]
            print(f"  {path.name}  ({path.stat().st_size / (1024 * 1024):.1f} MB)")
    final = outputs["final"]
    print(f"  {final.name}  ({final.stat().st_size / (1024 * 1024):.1f} MB)  ← deliverable")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


