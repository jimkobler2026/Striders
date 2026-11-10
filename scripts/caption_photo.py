#!/usr/bin/env python3
"""Generate a National Geographic–style caption (~60 words) for a photo.

Uses vision-capable LLMs plus helpful file metadata (EXIF / Keywords).
Does not dump every technical setting — metadata is context for the prose.

Examples:
  python3 scripts/caption_photo.py path/to/photo.jpg
  python3 scripts/caption_photo.py path/to/photo.jpg --provider anthropic,google
  python3 scripts/caption_photo.py path/to/photo.jpg --all

Environment (any subset):
  ANTHROPIC_API_KEY   → Claude vision
  GOOGLE_API_KEY or GEMINI_API_KEY → Gemini vision
  OPENAI_API_KEY      → GPT-4o vision
  OPENROUTER_API_KEY  → OpenRouter vision models
  GROQ_API_KEY        → Groq (if a vision model is available)

Optional: place keys in a project `.env` (loaded automatically).
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PROMPT_PATH = ROOT / "catalog" / "macros" / "caption_prompt.txt"


def load_prompt() -> str:
    if PROMPT_PATH.is_file():
        text = PROMPT_PATH.read_text(encoding="utf-8").strip()
        if text:
            return text
    return (
        "You are writing a single National Geographic magazine photo caption. "
        "About 60 words. Output ONLY the caption text."
    )


PROMPT = load_prompt()


@dataclass
class ProviderResult:
    provider: str
    model: str
    caption: str
    error: str = ""


def load_dotenv(env_path: Path | None = None) -> None:
    """Load KEY=VALUE pairs. Does not override existing environment variables."""
    candidates: list[Path] = []
    if env_path is not None:
        candidates.append(env_path.expanduser())
    candidates.append(ROOT / ".env")
    # Sibling Virtual_Nexus project often holds shared API keys
    candidates.append(ROOT.parent / "Virtual_Nexus" / ".env")

    path = next((p for p in candidates if p.is_file()), None)
    if path is None:
        return
    try:
        from dotenv import load_dotenv as _load

        _load(path, override=False)
    except Exception:
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, _, v = line.partition("=")
            k, v = k.strip(), v.strip().strip("'").strip('"')
            os.environ.setdefault(k, v)


def find_exiftool() -> str | None:
    for c in ("/usr/local/bin/exiftool", "/opt/homebrew/bin/exiftool", "exiftool"):
        p = shutil.which(c) if c == "exiftool" else c
        if p and Path(p).exists():
            return p
    return None


def read_metadata(path: Path) -> dict:
    exiftool = find_exiftool()
    meta: dict = {"SourceFile": str(path)}
    if not exiftool:
        return meta
    tags = [
        "DateTimeOriginal",
        "CreateDate",
        "Model",
        "LensID",
        "LensModel",
        "ExposureTime",
        "FNumber",
        "ISO",
        "Keywords",
        "Subject",
        "Caption-Abstract",
        "ImageDescription",
        "XPComment",
        "Megapixels",
        "ImageSize",
    ]
    proc = subprocess.run(
        [exiftool, "-json", "-s", *[f"-{t}" for t in tags], "--", str(path)],
        capture_output=True,
        text=True,
    )
    if proc.returncode in (0, 1) and proc.stdout:
        rows = json.loads(proc.stdout)
        if rows:
            meta.update(rows[0])
    return meta


def parse_catalog_keywords(meta: dict) -> dict[str, str]:
    """Pull Species:/Behavior:/… style keywords into a clean map."""
    out: dict[str, str] = {}
    raw = meta.get("Keywords") or meta.get("Subject") or []
    items = raw if isinstance(raw, list) else [raw]
    for item in items:
        text = str(item).strip()
        if ":" not in text:
            continue
        k, _, v = text.partition(":")
        k, v = k.strip(), v.strip()
        if k and v:
            out[k] = v
    return sanitize_tag_map(out)


def sanitize_tag_map(tags: dict[str, str]) -> dict[str, str]:
    """Drop clearly swapped/corrupt Sequence-vs-tag values from older write-backs."""
    out = dict(tags)
    # Filters should not look like a SequenceID stamp
    if re.fullmatch(r"\d{8}-\d{6}", out.get("Filters", "") or ""):
        out.pop("Filters", None)
    # CausticAppearance should not be a bare yes/no when Sequence carries that meaning
    if (out.get("CausticAppearance") or "").lower() in {"yes", "no"} and out.get("Sequence"):
        out.pop("CausticAppearance", None)
    if (out.get("CausticAim") or "").lower() in {"yes", "no"} and out.get("Sequence"):
        out.pop("CausticAim", None)
    return out


def read_catalog_row(photo: Path, workbook: Path | None = None) -> dict[str, str]:
    """If this file is in the Photos catalog, prefer those curated tags."""
    candidates = []
    if workbook is not None:
        candidates.append(workbook.expanduser())
    candidates.extend(
        [
            ROOT / "catalog" / "Strider_Media_Catalog.xlsm",
            ROOT / "catalog" / "Strider_Media_Catalog_blank.xlsm",
        ]
    )
    try:
        import openpyxl
    except ImportError:
        return {}

    want = photo.resolve()
    want_name = photo.name
    for wb_path in candidates:
        if not wb_path.is_file():
            continue
        try:
            wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=True)
            if "Photos" not in wb.sheetnames:
                continue
            ws = wb["Photos"]
            rows = ws.iter_rows(values_only=True)
            headers = [str(h).strip() if h is not None else "" for h in next(rows)]
            for row in rows:
                data = {headers[i]: row[i] for i in range(min(len(headers), len(row))) if headers[i]}
                name = str(data.get("Name") or "")
                path = str(data.get("Path") or "")
                if name == want_name or (path and Path(path).resolve() == want):
                    mapped = {
                        "Species": data.get("Species"),
                        "Behavior": data.get("Behavior"),
                        "Locale": data.get("Locale"),
                        "View": data.get("View"),
                        "Lighting": data.get("Lighting"),
                        "CausticAim": data.get("Caustic camera aim"),
                        "CausticAppearance": data.get("Caustic appearance"),
                        "Filters": data.get("Filters"),
                        "Sequence": data.get("Sequence"),
                        "SequenceID": data.get("SequenceID"),
                        "SequencePos": data.get("SequencePos"),
                        "notes": data.get("notes"),
                    }
                    return sanitize_tag_map(
                        {k: str(v).strip() for k, v in mapped.items() if v not in (None, "")}
                    )
        except Exception:
            continue
    return {}


def metadata_brief(meta: dict, catalog: dict[str, str] | None = None) -> str:
    kw = parse_catalog_keywords(meta)
    if catalog:
        # Catalog wins on curated fields
        kw = {**kw, **catalog}
        kw = sanitize_tag_map(kw)
    lines = []
    mapping = [
        ("Species", kw.get("Species")),
        ("Behavior", kw.get("Behavior")),
        ("Locale", kw.get("Locale")),
        ("View", kw.get("View")),
        ("Lighting", kw.get("Lighting")),
        ("CausticAim", kw.get("CausticAim")),
        ("CausticAppearance", kw.get("CausticAppearance")),
        ("Filters", kw.get("Filters")),
        ("Sequence", kw.get("Sequence")),
        ("SequenceID", kw.get("SequenceID")),
        ("SequencePos", kw.get("SequencePos")),
    ]
    for label, val in mapping:
        if val:
            lines.append(f"{label}: {val}")
    tech = []
    for k in ("DateTimeOriginal", "Model", "LensID", "LensModel", "ExposureTime", "FNumber", "ISO"):
        if meta.get(k) not in (None, ""):
            tech.append(f"{k}={meta[k]}")
    if tech:
        lines.append("Camera context (do not list unless essential): " + "; ".join(tech))
    notes = kw.get("notes") or meta.get("Caption-Abstract") or meta.get("ImageDescription") or ""
    if notes and str(notes).strip() not in ("", kw.get("SequencePos", "")):
        if not re.fullmatch(r"\d+/\d+", str(notes).strip()):
            lines.append(f"Photographer notes: {notes}")
    return "\n".join(lines) if lines else "(no useful metadata)"


def extract_embedded_preview(path: Path) -> Path | None:
    """Pull Sony/JPEG PreviewImage (or similar) via exiftool; None if absent."""
    exif = shutil.which("exiftool") or next(
        (c for c in ("/opt/homebrew/bin/exiftool", "/usr/local/bin/exiftool") if Path(c).exists()),
        None,
    )
    if not exif:
        return None
    for tag in ("PreviewImage", "JpgFromRaw", "OtherImage"):
        proc = subprocess.run(
            [exif, "-b", f"-{tag}", "--", str(path)],
            capture_output=True,
        )
        if proc.returncode not in (0, 1) or not proc.stdout or len(proc.stdout) < 1000:
            continue
        # JPEG SOI
        if not proc.stdout.startswith(b"\xff\xd8"):
            continue
        fd, name = tempfile.mkstemp(prefix="caption_embed_", suffix=".jpg")
        os.write(fd, proc.stdout)
        os.close(fd)
        return Path(name)
    return None


def make_preview_jpeg(path: Path, max_edge: int = 2048, *, prefer_embedded: bool = True) -> Path:
    """Build a small JPEG for vision APIs (embedded preview when present)."""
    from PIL import Image

    embedded: Path | None = None
    source = path
    if prefer_embedded:
        embedded = extract_embedded_preview(path)
        if embedded is not None:
            source = embedded
    try:
        im = Image.open(source)
        im = im.convert("RGB")
        im.thumbnail((max_edge, max_edge))
        fd, name = tempfile.mkstemp(prefix="caption_", suffix=".jpg")
        os.close(fd)
        out = Path(name)
        im.save(out, "JPEG", quality=85, optimize=True)
        return out
    finally:
        if embedded is not None:
            try:
                embedded.unlink(missing_ok=True)
            except Exception:
                pass


def b64_jpeg(path: Path) -> str:
    return base64.standard_b64encode(path.read_bytes()).decode("ascii")


def parse_description_sections(raw: str) -> tuple[str, str]:
    """Split MWG:Description into (note, caption). Legacy unprefixed → note."""
    text = (raw or "").strip()
    if not text:
        return "", ""
    note_at = text.find("NOTE:")
    cap_at = text.find("CAPTION:")
    if note_at < 0 and cap_at < 0:
        return text, ""
    note, caption = "", ""
    if note_at >= 0 and (cap_at < 0 or note_at < cap_at):
        if cap_at > note_at:
            note = text[note_at + 5 : cap_at].strip()
            caption = text[cap_at + 8 :].strip()
        else:
            note = text[note_at + 5 :].strip()
    elif cap_at >= 0:
        if note_at > cap_at:
            caption = text[cap_at + 8 : note_at].strip()
            note = text[note_at + 5 :].strip()
        else:
            caption = text[cap_at + 8 :].strip()
    return note, caption


def format_description(note: str, caption: str) -> str:
    parts: list[str] = []
    n = (note or "").strip()
    c = (caption or "").strip()
    if n:
        parts.append(f"NOTE: {n}")
    if c:
        parts.append(f"CAPTION: {c}")
    return "\n".join(parts)


def find_exiftool() -> str:
    for candidate in ("/usr/local/bin/exiftool", "/opt/homebrew/bin/exiftool", "exiftool"):
        path = shutil.which(candidate) if candidate == "exiftool" else candidate
        if path and Path(path).exists():
            return path
    raise RuntimeError("exiftool not found")


def embed_caption(path: Path, caption: str) -> None:
    """Merge CAPTION: into MWG:Description; preserve NOTE:."""
    exif = find_exiftool()
    # -b keeps real newlines; -s3 replaces them with "." and breaks NOTE/CAPTION splits.
    proc = subprocess.run(
        [exif, "-m", "-b", "-MWG:Description", "--", str(path)],
        capture_output=True,
        text=True,
    )
    current = (proc.stdout or "").strip() if proc.returncode in (0, 1) else ""
    note, _old_cap = parse_description_sections(current)
    new_desc = format_description(note, caption)
    cmd = [
        exif,
        "-m",
        "-P",
        "-overwrite_original",
        "-api",
        "LargeFileSupport=1",
        "-use",
        "MWG",
        "-charset",
        "iptc=UTF8",
        "-IPTC:CodedCharacterSet=UTF8",
        f"-MWG:Description={new_desc}",
        "--",
        str(path),
    ]
    written = subprocess.run(cmd, capture_output=True, text=True)
    if written.returncode not in (0, 1):
        raise RuntimeError(written.stderr or written.stdout or "caption embed failed")


def word_count(text: str) -> int:
    return len(re.findall(r"\b[\w’'-]+\b", text))


def available_providers() -> dict[str, bool]:
    return {
        "anthropic": bool(os.environ.get("ANTHROPIC_API_KEY")),
        "google": bool(os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY")),
        "openai": bool(os.environ.get("OPENAI_API_KEY")),
        "openrouter": bool(os.environ.get("OPENROUTER_API_KEY")),
        "groq": bool(os.environ.get("GROQ_API_KEY")),
    }


def run_anthropic(preview: Path, meta_text: str, model: str) -> ProviderResult:
    import anthropic

    client = anthropic.Anthropic()
    data = b64_jpeg(preview)
    msg = client.messages.create(
        model=model,
        max_tokens=500,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": data}},
                    {
                        "type": "text",
                        "text": f"{PROMPT}\n\nHelpful metadata:\n{meta_text}",
                    },
                ],
            }
        ],
    )
    text = "".join(getattr(b, "text", "") for b in msg.content).strip()
    return ProviderResult("anthropic", model, text)


def run_google(preview: Path, meta_text: str, model: str) -> ProviderResult:
    # Prefer new google.genai; fall back to generativeai
    api_key = os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY")
    prompt = f"{PROMPT}\n\nHelpful metadata:\n{meta_text}"
    try:
        from google import genai
        from google.genai import types

        client = genai.Client(api_key=api_key)
        uploaded = preview.read_bytes()
        resp = client.models.generate_content(
            model=model,
            contents=[
                types.Content(
                    role="user",
                    parts=[
                        types.Part.from_bytes(data=uploaded, mime_type="image/jpeg"),
                        types.Part.from_text(text=prompt),
                    ],
                )
            ],
        )
        text = (resp.text or "").strip()
        return ProviderResult("google", model, text)
    except Exception:
        import google.generativeai as genai

        genai.configure(api_key=api_key)
        gmodel = genai.GenerativeModel(model)
        resp = gmodel.generate_content(
            [
                {"mime_type": "image/jpeg", "data": preview.read_bytes()},
                prompt,
            ]
        )
        text = (getattr(resp, "text", None) or "").strip()
        return ProviderResult("google", model, text)


def run_openai(preview: Path, meta_text: str, model: str) -> ProviderResult:
    try:
        from openai import OpenAI
    except ImportError as e:
        return ProviderResult("openai", model, "", f"openai package missing: {e}")

    client = OpenAI()
    data = b64_jpeg(preview)
    resp = client.chat.completions.create(
        model=model,
        max_tokens=300,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": f"{PROMPT}\n\nHelpful metadata:\n{meta_text}"},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/jpeg;base64,{data}"},
                    },
                ],
            }
        ],
    )
    text = (resp.choices[0].message.content or "").strip()
    return ProviderResult("openai", model, text)


def run_openrouter(preview: Path, meta_text: str, model: str) -> ProviderResult:
    import httpx

    data = b64_jpeg(preview)
    headers = {
        "Authorization": f"Bearer {os.environ['OPENROUTER_API_KEY']}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": f"{PROMPT}\n\nHelpful metadata:\n{meta_text}"},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/jpeg;base64,{data}"},
                    },
                ],
            }
        ],
    }
    with httpx.Client(timeout=120.0) as client:
        r = client.post("https://openrouter.ai/api/v1/chat/completions", headers=headers, json=payload)
        r.raise_for_status()
        body = r.json()
    text = body["choices"][0]["message"]["content"].strip()
    return ProviderResult("openrouter", model, text)


def run_groq(preview: Path, meta_text: str, model: str) -> ProviderResult:
    import httpx

    data = b64_jpeg(preview)
    headers = {
        "Authorization": f"Bearer {os.environ['GROQ_API_KEY']}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": f"{PROMPT}\n\nHelpful metadata:\n{meta_text}"},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/jpeg;base64,{data}"},
                    },
                ],
            }
        ],
    }
    with httpx.Client(timeout=120.0) as client:
        r = client.post("https://api.groq.com/openai/v1/chat/completions", headers=headers, json=payload)
        r.raise_for_status()
        body = r.json()
    text = body["choices"][0].message["content"].strip() if isinstance(body["choices"][0]["message"], dict) else body["choices"][0]["message"]["content"].strip()
    # groq returns dict message
    if isinstance(body["choices"][0]["message"], dict):
        text = body["choices"][0]["message"]["content"].strip()
    return ProviderResult("groq", model, text)


DEFAULT_MODELS = {
    "anthropic": "claude-sonnet-4-5-20250929",
    "google": "gemini-2.5-flash",
    "openai": "gpt-4o",
    "openrouter": "openai/gpt-4o",
    "groq": "meta-llama/llama-4-scout-17b-16e-instruct",
}


def run_provider(name: str, preview: Path, meta_text: str, model: str | None) -> ProviderResult:
    model = model or DEFAULT_MODELS[name]
    try:
        if name == "anthropic":
            return run_anthropic(preview, meta_text, model)
        if name == "google":
            return run_google(preview, meta_text, model)
        if name == "openai":
            return run_openai(preview, meta_text, model)
        if name == "openrouter":
            return run_openrouter(preview, meta_text, model)
        if name == "groq":
            return run_groq(preview, meta_text, model)
        return ProviderResult(name, model, "", f"Unknown provider: {name}")
    except Exception as exc:  # noqa: BLE001
        return ProviderResult(name, model, "", str(exc))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Nat Geo–style photo caption via vision LLMs + metadata.")
    parser.add_argument("photo", type=Path, help="Path to image file")
    parser.add_argument(
        "--env-file",
        type=Path,
        default=None,
        help="Optional .env path (default: Striders/.env, then ../Virtual_Nexus/.env)",
    )
    parser.add_argument(
        "--provider",
        default="",
        help="Comma-separated: anthropic,google,openai,openrouter,groq "
        "(default: google if keyed, else first available; use --all for every keyed provider)",
    )
    parser.add_argument("--all", action="store_true", help="Run every provider that has an API key set")
    parser.add_argument("--model", default="", help="Override model id for a single provider run")
    parser.add_argument("--max-edge", type=int, default=2048, help="Max preview edge length for vision upload")
    parser.add_argument(
        "--catalog",
        type=Path,
        default=None,
        help="Optional Strider .xlsm to pull curated Photos-row tags",
    )
    parser.add_argument(
        "--plain",
        action="store_true",
        help="Print only the caption text (single provider; for Excel bridge)",
    )
    parser.add_argument(
        "--machine",
        action="store_true",
        help="Print OK\\tprovider\\tmodel\\twords\\tcaption (or ERROR\\tmessage) for Excel",
    )
    parser.add_argument("--json", action="store_true", help="Emit JSON results")
    parser.add_argument(
        "--debug-preview",
        action="store_true",
        help="Print whether an embedded PreviewImage was used and preview byte size",
    )
    parser.add_argument(
        "--write",
        action="store_true",
        help="Embed caption into MWG:Description as 'CAPTION: …' (preserves any 'NOTE: …')",
    )
    args = parser.parse_args(argv)
    load_dotenv(args.env_file)
    # Gemini clients often expect GOOGLE_API_KEY
    if os.environ.get("GEMINI_API_KEY") and not os.environ.get("GOOGLE_API_KEY"):
        os.environ["GOOGLE_API_KEY"] = os.environ["GEMINI_API_KEY"]

    photo = args.photo.expanduser().resolve()
    if not photo.is_file():
        raise SystemExit(f"Not a file: {photo}")

    avail = available_providers()
    if args.provider:
        wanted = [p.strip().lower() for p in args.provider.split(",") if p.strip()]
    elif args.all:
        wanted = [p for p, ok in avail.items() if ok]
    else:
        # Spreadsheet-friendly default: cheapest solid vision model first
        preferred = ["google", "openai", "anthropic", "openrouter", "groq"]
        wanted = [p for p in preferred if avail.get(p)][:1]

    if args.plain or args.machine:
        wanted = wanted[:1]

    if not wanted:
        print("No providers available. Set one or more of:", file=sys.stderr)
        for k in ("ANTHROPIC_API_KEY", "GOOGLE_API_KEY", "GEMINI_API_KEY", "OPENAI_API_KEY", "OPENROUTER_API_KEY", "GROQ_API_KEY"):
            print(f"  {k}", file=sys.stderr)
        print("Or add them to .env in the project root / Virtual_Nexus/.env.", file=sys.stderr)
        return 2

    missing = [p for p in wanted if not avail.get(p)]
    if missing:
        print(f"Missing API keys for: {', '.join(missing)}", file=sys.stderr)
        wanted = [p for p in wanted if avail.get(p)]
        if not wanted:
            return 2

    meta = read_metadata(photo)
    catalog = read_catalog_row(photo, args.catalog)
    meta_text = metadata_brief(meta, catalog)
    if args.debug_preview:
        embedded = extract_embedded_preview(photo)
        if embedded is not None:
            print(
                f"debug-preview: embedded {embedded.stat().st_size} bytes "
                f"(master {photo.stat().st_size} bytes)",
                file=sys.stderr,
            )
            try:
                embedded.unlink(missing_ok=True)
            except Exception:
                pass
        else:
            print(
                f"debug-preview: no embedded preview; will downscale master "
                f"({photo.stat().st_size} bytes)",
                file=sys.stderr,
            )
    preview = make_preview_jpeg(photo, max_edge=args.max_edge)
    results: list[ProviderResult] = []
    try:
        for name in wanted:
            model = args.model if args.model and len(wanted) == 1 else None
            results.append(run_provider(name, preview, meta_text, model))
    finally:
        try:
            preview.unlink(missing_ok=True)
        except Exception:
            pass

    def maybe_embed(text: str) -> None:
        if not args.write:
            return
        embed_caption(photo, " ".join(text.split()))
        print(f"Embedded CAPTION: into {photo.name}", file=sys.stderr)

    if args.machine:
        r = results[0]
        if r.error or not r.caption:
            print(f"ERROR\t{r.error or 'empty caption'}")
            return 1
        # Flatten newlines so Excel/AppleScript can split on tab
        cap = " ".join(r.caption.split())
        try:
            maybe_embed(cap)
        except Exception as exc:  # noqa: BLE001
            print(f"ERROR\tembed failed: {exc}")
            return 1
        print(f"OK\t{r.provider}\t{r.model}\t{word_count(cap)}\t{cap}")
        return 0

    if args.plain:
        r = results[0]
        if r.error or not r.caption:
            print(r.error or "empty caption", file=sys.stderr)
            return 1
        try:
            maybe_embed(r.caption)
        except Exception as exc:  # noqa: BLE001
            print(f"embed failed: {exc}", file=sys.stderr)
            return 1
        print(r.caption.strip())
        return 0

    if args.json:
        print(
            json.dumps(
                [
                    {
                        "provider": r.provider,
                        "model": r.model,
                        "words": word_count(r.caption) if r.caption else 0,
                        "caption": r.caption,
                        "error": r.error,
                    }
                    for r in results
                ],
                indent=2,
            )
        )
        if args.write:
            r0 = next((r for r in results if r.caption and not r.error), None)
            if r0:
                try:
                    maybe_embed(r0.caption)
                except Exception as exc:  # noqa: BLE001
                    print(f"embed failed: {exc}", file=sys.stderr)
                    return 1
        return 0 if all(r.caption and not r.error for r in results) else 1

    print(f"Photo: {photo}")
    print(f"Metadata used:\n{meta_text}\n")
    for r in results:
        print("=" * 72)
        print(f"{r.provider} · {r.model}")
        if r.error:
            print(f"ERROR: {r.error}")
        else:
            print(f"({word_count(r.caption)} words)")
            print(r.caption)
        print()
    if args.write:
        r0 = next((r for r in results if r.caption and not r.error), None)
        if r0:
            try:
                maybe_embed(r0.caption)
            except Exception as exc:  # noqa: BLE001
                print(f"embed failed: {exc}", file=sys.stderr)
                return 1
    return 0 if any(r.caption and not r.error for r in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
