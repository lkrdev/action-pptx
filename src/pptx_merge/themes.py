from __future__ import annotations

import re
from dataclasses import dataclass
from typing import cast

from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import BaseOxmlElement
from pptx.presentation import Presentation
from pptx.shapes.graphfrm import GraphicFrame

from pptx_merge.inspector import _iter_shapes


@dataclass(slots=True, frozen=True)
class PptxTheme:
    name: str
    brand: str
    mode: str
    series_palette: tuple[str, str, str, str, str]
    bg_color: str = "090D16"
    card_bg: str = "111827"
    card_border: str = "1E293B"
    row_alt_a: str = "0F172A"
    row_alt_b: str = "162032"
    badge_bg: str = "0B2622"
    text_primary: str = "F8FAFC"
    text_secondary: str = "E2E8F0"
    text_muted: str = "94A3B8"
    text_dim: str = "64748B"


_LIGHT = {
    "bg_color": "F0F4F9",
    "card_bg": "FFFFFF",
    "card_border": "CBD5E1",
    "row_alt_a": "F8FAFC",
    "row_alt_b": "F1F5F9",
    "badge_bg": "DCFCE7",
    "text_primary": "0F172A",
    "text_secondary": "334155",
    "text_muted": "475569",
    "text_dim": "64748B",
}

_BRAND_SPECS: tuple[
    tuple[str, str, tuple[str, str, str, str, str], tuple[str, str, str, str, str]], ...
] = (
    (
        "Embed_Demo",
        "Default",
        ("38BDF8", "818CF8", "34D399", "FBBF24", "FB7185"),
        ("0B57D0", "4F46E5", "059669", "D97706", "E11D48"),
    ),
    (
        "Levis",
        "Levi's",
        ("E11D48", "3B82F6", "2DD4BF", "F59E0B", "A855F7"),
        ("C41230", "0052CC", "0D9488", "D97706", "7C3AED"),
    ),
    (
        "Calvin_Klein",
        "Calvin Klein",
        ("94A3B8", "38BDF8", "CBD5E1", "64748B", "A1A1AA"),
        ("334155", "0284C7", "64748B", "475569", "0F766E"),
    ),
    (
        "Columbia",
        "Columbia",
        ("00A896", "38BDF8", "F4A261", "34D399", "E76F51"),
        ("0066B2", "00A896", "D97706", "059669", "DC2626"),
    ),
    (
        "Allegra_K",
        "Allegra K",
        ("F43F5E", "C084FC", "EC4899", "FBBF24", "2DD4BF"),
        ("BE123C", "9333EA", "DB2777", "D97706", "0D9488"),
    ),
)

PPTX_THEMES: dict[str, PptxTheme] = {}
for _key, _label, _dark_pal, _light_pal in _BRAND_SPECS:
    PPTX_THEMES[f"{_key}_Dark"] = PptxTheme(f"{_key}_Dark", _label, "dark", _dark_pal)
    PPTX_THEMES[f"{_key}_Light"] = PptxTheme(f"{_key}_Light", _label, "light", _light_pal, **_LIGHT)

BASE_TEMPLATE_THEME = PPTX_THEMES["Embed_Demo_Dark"]


def get_embed_theme_name(is_dark: bool, brand: str | None = None) -> str:
    """Resolve Looker Embed Demo theme name (<SanitizedBrand>_Light / _Dark)."""
    suffix = "Dark" if is_dark else "Light"
    cleaned = (brand or "").strip()
    if not cleaned or cleaned.lower() in ("default", "all", "%", "embed_demo"):
        return f"Embed_Demo_{suffix}"
    sanitized = re.sub(r"\s+", "_", re.sub(r"[^a-zA-Z0-9\s]", "", cleaned).strip())
    candidate = f"{sanitized}_{suffix}"
    return candidate if candidate in PPTX_THEMES else f"Embed_Demo_{suffix}"


def _hex6(raw: object, default: str) -> str:
    cleaned = str(raw or "").strip().lstrip("#")
    if len(cleaned) in (3, 4) and all(c in "0123456789abcdefABCDEF" for c in cleaned):
        return "".join(c * 2 for c in cleaned[:3]).upper()
    if len(cleaned) >= 6 and all(c in "0123456789abcdefABCDEF" for c in cleaned[:6]):
        return cleaned[:6].upper()
    return default


def _is_dark_hex(hex6: str) -> bool:
    r, g, b = int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)
    return (0.299 * r + 0.587 * g + 0.114 * b) < 128.0


def theme_from_looker_palette(name: str, palette: dict[str, object]) -> PptxTheme:
    """Build a PptxTheme dynamically from a Looker Theme + ColorCollection payload."""
    bg = _hex6(palette.get("bg_color"), "090D16")
    dark = _is_dark_hex(bg)
    defaults = BASE_TEMPLATE_THEME if dark else PPTX_THEMES["Embed_Demo_Light"]
    raw_series = palette.get("series_colors")
    parsed_series = [
        _hex6(c, "") for c in (raw_series if isinstance(raw_series, list) else []) if _hex6(c, "")
    ]
    accent = _hex6(palette.get("accent"), defaults.series_palette[0])
    fallback_series = [accent, *defaults.series_palette]
    merged_series = (parsed_series + fallback_series)[:5]
    return PptxTheme(
        name=name or "Custom_Looker_Theme",
        brand=name or "Custom",
        mode="dark" if dark else "light",
        series_palette=cast(tuple[str, str, str, str, str], tuple(merged_series)),
        bg_color=bg,
        card_bg=_hex6(palette.get("card_bg"), defaults.card_bg),
        card_border=defaults.card_border,
        row_alt_a=_hex6(palette.get("card_bg"), defaults.row_alt_a),
        row_alt_b=defaults.row_alt_b,
        badge_bg=defaults.badge_bg,
        text_primary=_hex6(palette.get("title_color"), defaults.text_primary),
        text_secondary=_hex6(palette.get("text_color"), defaults.text_secondary),
        text_muted=_hex6(palette.get("text_color"), defaults.text_muted),
        text_dim=defaults.text_dim,
    )


def resolve_pptx_theme(
    theme: str | PptxTheme | None = None,
    *,
    brand: str | None = None,
    is_dark: bool = True,
) -> PptxTheme | None:
    """Resolve a theme preset name, brand name, or PptxTheme instance."""
    if isinstance(theme, PptxTheme):
        return theme
    raw = (theme or "").strip()
    if raw.lower() in ("", "auto", "none", "default"):
        if brand and brand.strip() and brand.strip() not in ("%", "all"):
            return PPTX_THEMES[get_embed_theme_name(is_dark, brand)]
        return None
    for key, val in PPTX_THEMES.items():
        if key.lower() == raw.lower():
            return val
    if raw.lower() in ("dark", "light"):
        return PPTX_THEMES[get_embed_theme_name(raw.lower() == "dark", brand)]
    return PPTX_THEMES[get_embed_theme_name(is_dark, raw)]


def apply_pptx_theme(prs: Presentation, theme: PptxTheme) -> None:
    """Recolor a Presentation's OpenXML fills, borders, text, tables, and charts to `theme`."""
    base = BASE_TEMPLATE_THEME
    color_map: dict[str, str] = {
        base.bg_color: theme.bg_color,
        base.card_bg: theme.card_bg,
        base.card_border: theme.card_border,
        base.row_alt_a: theme.row_alt_a,
        base.row_alt_b: theme.row_alt_b,
        base.badge_bg: theme.badge_bg,
        base.text_primary: theme.text_primary,
        base.text_secondary: theme.text_secondary,
        base.text_muted: theme.text_muted,
        base.text_dim: theme.text_dim,
        **dict(zip(base.series_palette, theme.series_palette, strict=False)),
    }
    palette_rgbs = [RGBColor.from_string(c) for c in theme.series_palette]

    def _recolor(el_root: BaseOxmlElement) -> None:
        for node in el_root.xpath(".//a:srgbClr"):
            el = cast(BaseOxmlElement, node)
            if (val := (el.get("val") or "").upper()) in color_map:
                el.set("val", color_map[val])

    for slide in prs.slides:
        _recolor(slide._element)
        for shape in _iter_shapes(slide.shapes):
            if not (isinstance(shape, GraphicFrame) and shape.has_chart):
                continue
            chart = shape.chart
            _recolor(chart._chartSpace)

            if chart._chartSpace.xpath(".//c:doughnutChart//c:dPt | .//c:pieChart//c:dPt"):
                for idx, pt in enumerate(chart.series[0].points):
                    pt.format.fill.solid()
                    pt.format.fill.fore_color.rgb = palette_rgbs[idx % len(palette_rgbs)]
            else:
                for idx, ser in enumerate(chart.series):
                    clr = palette_rgbs[idx % len(palette_rgbs)]
                    ser.format.fill.solid()
                    ser.format.fill.fore_color.rgb = clr
                    ser.format.line.color.rgb = clr
