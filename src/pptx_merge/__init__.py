from pptx_merge.engine import MergeError, MergeResult, merge_pptx
from pptx_merge.formatter import SUPPORTED_MASKS, format_value, sanitize_tile_slug
from pptx_merge.inspector import (
    DeckValidationReport,
    SlideBinding,
    TileSchema,
    ValidationIssue,
    inspect_and_validate,
)
from pptx_merge.themes import (
    PPTX_THEMES,
    PptxTheme,
    apply_pptx_theme,
    get_embed_theme_name,
    resolve_pptx_theme,
)

__all__ = [
    "DeckValidationReport",
    "MergeError",
    "MergeResult",
    "PPTX_THEMES",
    "PptxTheme",
    "SUPPORTED_MASKS",
    "SlideBinding",
    "TileSchema",
    "ValidationIssue",
    "apply_pptx_theme",
    "format_value",
    "get_embed_theme_name",
    "inspect_and_validate",
    "merge_pptx",
    "resolve_pptx_theme",
    "sanitize_tile_slug",
]
