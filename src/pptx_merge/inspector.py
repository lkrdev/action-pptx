from __future__ import annotations

import io
import zipfile
from collections.abc import Iterable, Iterator
from dataclasses import asdict, dataclass, field

from pptx import Presentation
from pptx.shapes.base import BaseShape
from pptx.shapes.group import GroupShape
from pptx.table import Table

from pptx_merge.formatter import SUPPORTED_MASKS, TAG_RE, sanitize_tile_slug

PLAN_FIELDS: frozenset[str] = frozenset(
    {"title", "url", "run_at", "scheduled_plan_id", "webhook_id"}
)


@dataclass(slots=True)
class TileSchema:
    title: str
    columns: list[str] = field(default_factory=list)
    slug: str = ""

    def __post_init__(self) -> None:
        if not self.slug:
            self.slug = sanitize_tile_slug(self.title)


@dataclass(slots=True)
class SlideBinding:
    slide_index: int
    kind: str  # "chart" | "table" | "text"
    shape_name: str
    tile_slug: str
    column: str | None = None
    row_index: int = 0
    format_mask: str | None = None
    raw_tag: str | None = None


@dataclass(slots=True)
class ValidationIssue:
    code: str
    message: str
    slide_index: int | None = None
    shape_name: str | None = None


@dataclass(slots=True)
class DeckValidationReport:
    valid: bool
    slide_count: int
    bindings: list[SlideBinding] = field(default_factory=list)
    issues: list[ValidationIssue] = field(default_factory=list)

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


def _iter_shapes(shapes: Iterable[BaseShape]) -> Iterator[BaseShape]:
    for shape in shapes:
        yield shape
        if isinstance(shape, GroupShape):
            yield from _iter_shapes(shape.shapes)


def _scan_text_tags(
    text: str,
    *,
    slide_index: int,
    shape_name: str,
    default_tile: str,
    bindings: list[SlideBinding],
    issues: list[ValidationIssue],
) -> None:
    for match in TAG_RE.finditer(text):
        raw_tile = match.group("tile")
        tile_slug = sanitize_tile_slug(raw_tile) if raw_tile else default_tile
        col = match.group("col").strip()
        row_str = match.group("row")
        row_idx = int(row_str) if row_str is not None else 0
        fmt = match.group("fmt")
        mask = fmt.strip() if fmt else None

        if not tile_slug:
            issues.append(
                ValidationIssue(
                    code="missing_tile_in_tag",
                    message=(
                        f"Tag '{match.group(0)}' omits tile slug outside a "
                        "'table:<tile>' prototype row."
                    ),
                    slide_index=slide_index,
                    shape_name=shape_name,
                )
            )
            continue

        if mask and mask not in SUPPORTED_MASKS:
            issues.append(
                ValidationIssue(
                    code="invalid_format_mask",
                    message=f"Unsupported format mask '{mask}' in tag '{match.group(0)}'.",
                    slide_index=slide_index,
                    shape_name=shape_name,
                )
            )

        bindings.append(
            SlideBinding(
                slide_index=slide_index,
                kind="text",
                shape_name=shape_name,
                tile_slug=tile_slug,
                column=col,
                row_index=row_idx,
                format_mask=mask,
                raw_tag=match.group(0),
            )
        )


def _scan_table_paragraphs(
    table: Table,
    *,
    slide_index: int,
    shape_name: str,
    proto_tile: str,
    bindings: list[SlideBinding],
    issues: list[ValidationIssue],
) -> None:
    rows = list(table.rows)
    for r_idx, row in enumerate(rows):
        default_tile = proto_tile if r_idx == len(rows) - 1 else ""
        for cell in row.cells:
            for para in cell.text_frame.paragraphs:
                _scan_text_tags(
                    para.text,
                    slide_index=slide_index,
                    shape_name=shape_name,
                    default_tile=default_tile,
                    bindings=bindings,
                    issues=issues,
                )


def inspect_and_validate(
    pptx_bytes: bytes,
    dashboard_tiles: list[TileSchema] | None = None,
) -> DeckValidationReport:
    """Inspect a .pptx template for bindings and validate against dashboard tiles."""
    issues: list[ValidationIssue] = []
    bindings: list[SlideBinding] = []

    try:
        with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as zf:
            for name in sorted(zf.namelist()):
                if name.startswith("ppt/charts/chartEx") and name.endswith(".xml"):
                    issues.append(
                        ValidationIssue(
                            code="unsupported_chart_ex",
                            message=(
                                f"Part '{name}' uses Office 2016+ Extended Chart XML (chartEx). "
                                "Convert waterfall/funnel/treemap charts to standard Office charts."
                            ),
                        )
                    )
        prs = Presentation(io.BytesIO(pptx_bytes))
    except Exception as exc:
        return DeckValidationReport(
            valid=False,
            slide_count=0,
            issues=[ValidationIssue(code="invalid_pptx", message=f"Invalid .pptx archive: {exc}")],
        )

    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape in _iter_shapes(slide.shapes):
            name = (shape.name or "").strip()
            lower_name = name.lower()

            if lower_name.startswith(("chart:", "table:")):
                kind, raw_slug = lower_name.split(":", 1)[0], name.split(":", 1)[1].strip()
                tile_slug = sanitize_tile_slug(raw_slug)
                has_kind = bool(getattr(shape, f"has_{kind}", False))
                if not tile_slug:
                    issues.append(
                        ValidationIssue(
                            code=f"empty_{kind}_slug",
                            message=f"{kind.capitalize()} shape '{name}' has an empty tile slug.",
                            slide_index=slide_idx,
                            shape_name=name,
                        )
                    )
                elif not has_kind:
                    issues.append(
                        ValidationIssue(
                            code=f"invalid_{kind}_shape",
                            message=(
                                f"Shape '{name}' is named '{kind}:...' "
                                f"but is not a PowerPoint {kind}."
                            ),
                            slide_index=slide_idx,
                            shape_name=name,
                        )
                    )
                else:
                    bindings.append(
                        SlideBinding(
                            slide_index=slide_idx,
                            kind=kind,
                            shape_name=name,
                            tile_slug=tile_slug,
                        )
                    )
                    if kind == "table":
                        _scan_table_paragraphs(
                            getattr(shape, "table"),  # noqa: B009
                            slide_index=slide_idx,
                            shape_name=name,
                            proto_tile=tile_slug,
                            bindings=bindings,
                            issues=issues,
                        )
                continue

            if getattr(shape, "has_table", False):
                _scan_table_paragraphs(
                    getattr(shape, "table"),  # noqa: B009
                    slide_index=slide_idx,
                    shape_name=name,
                    proto_tile="",
                    bindings=bindings,
                    issues=issues,
                )

            if getattr(shape, "has_text_frame", False):
                tf = getattr(shape, "text_frame")  # noqa: B009
                for para in tf.paragraphs:
                    _scan_text_tags(
                        para.text,
                        slide_index=slide_idx,
                        shape_name=name,
                        default_tile="",
                        bindings=bindings,
                        issues=issues,
                    )

    if dashboard_tiles is not None:
        tile_map: dict[str, TileSchema] = {}
        duplicate_slugs: set[str] = set()
        for t in dashboard_tiles:
            slug = sanitize_tile_slug(t.slug or t.title)
            if slug in tile_map:
                duplicate_slugs.add(slug)
            tile_map[slug] = t

        referenced = {b.tile_slug for b in bindings if b.tile_slug and b.tile_slug != "_plan"}
        for dup in sorted(duplicate_slugs & referenced):
            issues.append(
                ValidationIssue(
                    code="duplicate_dashboard_tile",
                    message=(
                        f"Multiple dashboard tiles sanitize to '{dup}'. "
                        f"Looker renames colliding CSV files to '{dup}_1.csv'; "
                        "give each tile a unique title."
                    ),
                )
            )

        for b in bindings:
            if b.tile_slug == "_plan":
                if b.column and b.column not in PLAN_FIELDS:
                    issues.append(
                        ValidationIssue(
                            code="invalid_plan_field",
                            message=(
                                f"Unknown _plan field '{b.column}' "
                                f"(expected one of {sorted(PLAN_FIELDS)})."
                            ),
                            slide_index=b.slide_index,
                            shape_name=b.shape_name,
                        )
                    )
                continue

            tile = tile_map.get(b.tile_slug)
            if tile is None:
                issues.append(
                    ValidationIssue(
                        code="missing_tile",
                        message=(
                            f"Referenced tile '{b.tile_slug}' was not found "
                            "in the dashboard schema."
                        ),
                        slide_index=b.slide_index,
                        shape_name=b.shape_name,
                    )
                )
                continue

            if b.kind == "chart" and len(tile.columns) < 2:
                issues.append(
                    ValidationIssue(
                        code="insufficient_chart_columns",
                        message=(
                            f"Chart '{b.shape_name}' binds to tile '{b.tile_slug}', which has "
                            f"{len(tile.columns)} column(s); at least 2 are required."
                        ),
                        slide_index=b.slide_index,
                        shape_name=b.shape_name,
                    )
                )

            if b.column is not None and b.column not in tile.columns:
                issues.append(
                    ValidationIssue(
                        code="missing_column",
                        message=(
                            f"Column '{b.column}' referenced by '{b.raw_tag or b.shape_name}' "
                            f"was not found in tile '{b.tile_slug}' (available: {tile.columns})."
                        ),
                        slide_index=b.slide_index,
                        shape_name=b.shape_name,
                    )
                )

    return DeckValidationReport(
        valid=len(issues) == 0,
        slide_count=len(prs.slides),
        bindings=bindings,
        issues=issues,
    )
