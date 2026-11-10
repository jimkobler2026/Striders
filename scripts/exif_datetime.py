"""Parse exiftool DateTimeOriginal with optional OffsetTimeOriginal."""

from __future__ import annotations

import re
from datetime import datetime, timedelta, timezone

OFFSET_RE = re.compile(r"^([+-])(\d{2}):(\d{2})$")


def parse_offset(raw: str | None) -> timezone | None:
    if not raw:
        return None
    text = str(raw).strip()
    match = OFFSET_RE.match(text)
    if not match:
        return None
    sign = 1 if match.group(1) == "+" else -1
    hours = int(match.group(2))
    minutes = int(match.group(3))
    return timezone(sign * timedelta(hours=hours, minutes=minutes))


def parse_exif_dt(raw: str | None, offset_raw: str | None = None) -> datetime | None:
    """Parse camera datetime; attach OffsetTimeOriginal when present.

    DateTimeOriginal is the camera's local wall clock. With an offset, the
    result is timezone-aware so gap math stays correct across DST/travel.
    SequenceID / rename prefixes should use ``stamp_local`` (wall clock).
    """
    if not raw:
        return None
    text = str(raw).strip()
    if not text or text.startswith("0000"):
        return None
    naive: datetime | None = None
    for fmt in ("%Y:%m:%d %H:%M:%S.%f", "%Y:%m:%d %H:%M:%S"):
        try:
            naive = datetime.strptime(text, fmt)
            break
        except ValueError:
            continue
    if naive is None:
        return None
    tz = parse_offset(offset_raw)
    if tz is not None:
        return naive.replace(tzinfo=tz)
    return naive


def stamp_local(dt: datetime) -> str:
    """YYYYMMDD-HHMMSS in the datetime's wall-clock (offset) components."""
    return dt.strftime("%Y%m%d-%H%M%S")


def gap_seconds(a: datetime, b: datetime) -> float:
    """Absolute seconds between two datetimes (aware or naive)."""
    if a.tzinfo is None and b.tzinfo is not None:
        a = a.replace(tzinfo=b.tzinfo)
    elif b.tzinfo is None and a.tzinfo is not None:
        b = b.replace(tzinfo=a.tzinfo)
    return abs((b - a).total_seconds())
