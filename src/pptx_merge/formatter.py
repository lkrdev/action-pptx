from __future__ import annotations

import re
from datetime import datetime

# Matches Looker's DashboardZIPProducer filename sanitizer:
# gsub(/(^\.+$)|[\/?*\\:|\"><\s%]/, '_')
_SLUG_RE = re.compile(r'(^\.+$)|[/?*\\:|"<>\s%]')

# Matches {{tile_slug.col}}, {{tile_slug[0].col | $#,##0}}, or prototype {{.col | 0.0%}}
TAG_RE = re.compile(
    r"\{\{\s*(?P<tile>[a-zA-Z0-9_]*)(?:\[(?P<row>\d+)\])?\.(?P<col>[^}|]+?)"
    r"(?:\s*\|\s*(?P<fmt>[^}]+?))?\s*\}\}"
)

# (scale, decimals, is_currency, suffix)
_NUMERIC_MASKS: dict[str, tuple[float, int, bool, str]] = {
    "#,##0": (1.0, 0, False, ""),
    "0.0": (1.0, 1, False, ""),
    "0.00": (1.0, 2, False, ""),
    "$#,##0": (1.0, 0, True, ""),
    "$#,##0.00": (1.0, 2, True, ""),
    "0%": (100.0, 0, False, "%"),
    "0.0%": (100.0, 1, False, "%"),
    "0.00%": (100.0, 2, False, "%"),
}

SUPPORTED_MASKS: frozenset[str] = frozenset({*_NUMERIC_MASKS, "yyyy-mm-dd"})


def sanitize_tile_slug(title: str) -> str:
    """Normalize a dashboard tile title to Looker's csv_zip filename stem."""
    return _SLUG_RE.sub("_", title.strip()).lower()


def format_value(raw: str | int | float | None, mask: str | None = None) -> str:
    """Format a raw csv_zip cell value using a deterministic inline mask."""
    if raw is None or (text := str(raw)) == "":
        return ""
    if not mask:
        return text

    m = mask.strip()
    if m == "yyyy-mm-dd":
        s = text.strip()
        return (
            s[:10]
            if re.match(r"^\d{4}-\d{2}-\d{2}", s)
            else datetime.fromisoformat(s).strftime("%Y-%m-%d")
        )

    spec = _NUMERIC_MASKS.get(m)
    if spec is None:
        raise ValueError(f"Unsupported format mask: {m!r}")

    scale, decimals, is_currency, suffix = spec
    val = float(text.strip().replace(",", "")) * scale
    r = round(val, decimals) if decimals else round(val)
    if abs(r) < 1e-12:
        r = 0
    if is_currency:
        sign = "-$" if r < 0 else "$"
        return f"{sign}{abs(r):,.{decimals}f}"
    return f"{r:,.{decimals}f}{suffix}"
