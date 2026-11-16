from __future__ import annotations

import copy
import csv
import io
import math
import posixpath
import re
import zipfile
from dataclasses import dataclass, field
from typing import cast

import openpyxl
from pptx import Presentation
from pptx.chart.chart import Chart
from pptx.chart.data import CategoryChartData
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import BaseOxmlElement
from pptx.parts.chart import ChartPart
from pptx.shapes.graphfrm import GraphicFrame
from pptx.text.text import _Paragraph

from pptx_merge.formatter import TAG_RE, format_value, sanitize_tile_slug
from pptx_merge.inspector import TileSchema, _iter_shapes, inspect_and_validate
from pptx_merge.themes import PptxTheme, apply_pptx_theme, resolve_pptx_theme

MAX_UNCOMPRESSED_ZIP_BYTES = 100 * 1024 * 1024  # 100 MiB guard
LOOKER_ROW_LIMIT = 5000
ACCESS_DENIED_SENTINEL = "You do not have access to this data"


class MergeError(Exception):
    """Raised when a .pptx merge fails validation, row limits, or XML/XLSX parity."""


@dataclass(slots=True)
class TileCSVData:
    slug: str
    headers: list[str]
    rows: list[dict[str, str]]
    positional_rows: list[list[str]]


@dataclass(slots=True)
class MergeResult:
    pptx_bytes: bytes
    slides_processed: int
    charts_updated: int
    tables_updated: int
    text_tags_replaced: int
    tile_row_counts: dict[str, int] = field(default_factory=dict)


def _unpack_csv_zip(csv_zip_bytes: bytes) -> tuple[dict[str, TileCSVData], set[str], set[str]]:
    """Return (tiles_by_slug, raw_stems, access_denied_slugs) from Looker's csv_zip payload."""
    try:
        zf = zipfile.ZipFile(io.BytesIO(csv_zip_bytes))
    except zipfile.BadZipFile as exc:
        raise MergeError("Attachment is not a valid ZIP archive.") from exc

    with zf:
        total_uncompressed = 0
        tiles: dict[str, TileCSVData] = {}
        raw_stems: set[str] = set()
        denied_slugs: set[str] = set()

        for info in zf.infolist():
            if info.is_dir():
                continue
            total_uncompressed += info.file_size
            if total_uncompressed > MAX_UNCOMPRESSED_ZIP_BYTES:
                raise MergeError(
                    f"csv_zip uncompressed size exceeds {MAX_UNCOMPRESSED_ZIP_BYTES} bytes limit."
                )

            base = posixpath.basename(info.filename)
            if not base.lower().endswith(".csv"):
                continue

            stem = base[:-4]
            raw_stems.add(stem.lower())
            slug = sanitize_tile_slug(stem)

            text = zf.read(info.filename).decode("utf-8-sig", errors="replace")
            if ACCESS_DENIED_SENTINEL in text:
                denied_slugs.add(slug)
                continue

            reader = list(csv.reader(io.StringIO(text)))
            if not reader:
                tiles[slug] = TileCSVData(slug=slug, headers=[], rows=[], positional_rows=[])
                continue

            headers = [h.strip() for h in reader[0]]
            pos_rows = reader[1:]
            dict_rows = [
                {headers[i]: (row[i] if i < len(row) else "") for i in range(len(headers))}
                for row in pos_rows
            ]
            tiles[slug] = TileCSVData(
                slug=slug,
                headers=headers,
                rows=dict_rows,
                positional_rows=pos_rows,
            )

        return tiles, raw_stems, denied_slugs


def _replace_paragraph_tags(
    para: _Paragraph,
    *,
    tiles: dict[str, TileCSVData],
    plan_meta: dict[str, str],
    default_tile: str = "",
    default_row_index: int | None = None,
) -> int:
    full_text = para.text
    if "{{" not in full_text:
        return 0

    matches = list(TAG_RE.finditer(full_text))
    if not matches:
        return 0

    def _sub(match: re.Match[str]) -> str:
        raw_tile = match.group("tile")
        tile_slug = sanitize_tile_slug(raw_tile) if raw_tile else default_tile
        col = match.group("col").strip()
        row_str = match.group("row")
        row_idx = (
            int(row_str)
            if row_str is not None
            else (default_row_index if default_row_index is not None else 0)
        )
        fmt = match.group("fmt")
        mask = fmt.strip() if fmt else None

        if tile_slug == "_plan":
            return format_value(plan_meta.get(col, ""), mask)

        tile = tiles.get(tile_slug)
        if tile is None:
            raise MergeError(f"Tile '{tile_slug}' is missing from csv_zip.")
        if col not in tile.headers:
            raise MergeError(f"Column '{col}' is missing from tile '{tile_slug}'.")
        if row_idx < 0 or row_idx >= len(tile.rows):
            raise MergeError(
                f"Tag '{match.group(0)}' requests row index {row_idx}, "
                f"but tile '{tile_slug}' has {len(tile.rows)} row(s)."
            )
        return format_value(tile.rows[row_idx].get(col, ""), mask)

    new_text = TAG_RE.sub(_sub, full_text)
    runs = list(para.runs)
    if runs:
        runs[0].text = new_text
        for extra_run in runs[1:]:
            parent = extra_run._r.getparent()
            if parent is not None:
                parent.remove(extra_run._r)
    else:
        para.text = new_text
    return len(matches)


def _set_tc_text_preserving_rpr(tc: BaseOxmlElement, value: str) -> None:
    tx_body = tc.find(qn("a:txBody"))
    if tx_body is None or not (p_list := tx_body.findall(qn("a:p"))):
        return
    first_p = p_list[0]
    for extra_p in p_list[1:]:
        tx_body.remove(extra_p)
    r_list = first_p.findall(qn("a:r"))
    if r_list:
        t_el = r_list[0].find(qn("a:t"))
        if t_el is not None:
            t_el.text = value
        for extra_r in r_list[1:]:
            first_p.remove(extra_r)
    else:
        r_el = first_p.makeelement(qn("a:r"), {})
        t_el = first_p.makeelement(qn("a:t"), {})
        t_el.text = value
        r_el.append(t_el)
        first_p.append(r_el)


def _populate_dynamic_table(
    shape: GraphicFrame,
    tile: TileCSVData,
    tiles: dict[str, TileCSVData],
    plan_meta: dict[str, str],
) -> int:
    if not tile.rows:
        raise MergeError(f"Tile '{tile.slug}' returned 0 rows for dynamic table '{shape.name}'.")

    table = shape.table
    replaced_tags = 0
    all_rows = list(table.rows)

    for header_row in all_rows[:-1]:
        for cell in header_row.cells:
            for para in cell.text_frame.paragraphs:
                replaced_tags += _replace_paragraph_tags(para, tiles=tiles, plan_meta=plan_meta)

    tbl_el = table._tbl
    proto_tr = tbl_el.tr_lst[-1]
    proto_has_tags = "{{" in "".join(proto_tr.itertext())

    for pos_row in tile.positional_rows:
        new_tr = copy.deepcopy(proto_tr)
        proto_tr.addprevious(new_tr)
        if not proto_has_tags:
            for col_idx, tc in enumerate(new_tr.findall(qn("a:tc"))):
                _set_tc_text_preserving_rpr(tc, pos_row[col_idx] if col_idx < len(pos_row) else "")

    tbl_el.remove(proto_tr)

    if proto_has_tags:
        header_count = len(all_rows) - 1
        for row_idx, data_row in enumerate(list(table.rows)[header_count:]):
            for cell in data_row.cells:
                for para in cell.text_frame.paragraphs:
                    replaced_tags += _replace_paragraph_tags(
                        para,
                        tiles=tiles,
                        plan_meta=plan_meta,
                        default_tile=tile.slug,
                        default_row_index=row_idx,
                    )

    return replaced_tags


def _verify_chart_xml_xlsx_parity(
    chart: Chart,
    categories: list[str],
    series_headers: list[str],
    series_values: list[list[float | None]],
) -> None:
    xlsx_blob = cast(ChartPart, chart.part).chart_workbook.xlsx_part.blob
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_blob), data_only=True)
    try:
        ws = wb.active
        if ws is None:
            raise MergeError("Embedded chart workbook has no active worksheet.")

        for s_idx, expected_header in enumerate(series_headers):
            actual_header = ws.cell(row=1, column=2 + s_idx).value
            if str(actual_header or "") != expected_header:
                raise MergeError(
                    f"Embedded .xlsx header mismatch at column {2 + s_idx}: "
                    f"expected {expected_header!r}, got {actual_header!r}."
                )

        for r_idx, expected_cat in enumerate(categories):
            actual_cat = ws.cell(row=2 + r_idx, column=1).value
            if str(actual_cat if actual_cat is not None else "") != expected_cat:
                raise MergeError(
                    f"Embedded .xlsx category mismatch at row {2 + r_idx}: "
                    f"expected {expected_cat!r}, got {actual_cat!r}."
                )

        for s_idx, vals in enumerate(series_values):
            for r_idx, expected_val in enumerate(vals):
                actual_val = ws.cell(row=2 + r_idx, column=2 + s_idx).value
                if expected_val is None:
                    if actual_val is not None and str(actual_val).strip() != "":
                        raise MergeError(
                            f"Embedded .xlsx expected blank at ({2 + r_idx}, {2 + s_idx}), "
                            f"got {actual_val!r}."
                        )
                elif (
                    actual_val is None
                    or not isinstance(actual_val, int | float | str)
                    or not math.isclose(float(actual_val), expected_val, rel_tol=1e-9, abs_tol=1e-9)
                ):
                    raise MergeError(
                        f"Embedded .xlsx value mismatch at ({2 + r_idx}, {2 + s_idx}): "
                        f"expected {expected_val}, got {actual_val!r}."
                    )
    finally:
        wb.close()

    ser_elements = chart._chartSpace.xpath(".//c:ser")
    if len(ser_elements) != len(series_values):
        raise MergeError(
            f"Chart XML series count ({len(ser_elements)}) does not match "
            f"CSV series count ({len(series_values)})."
        )

    for s_idx, ser_el in enumerate(ser_elements):
        pts = ser_el.xpath("./c:val/c:numRef/c:numCache/c:pt | ./c:yVal/c:numRef/c:numCache/c:pt")
        pt_map: dict[int, float] = {}
        for pt in pts:
            idx_attr = pt.get("idx")
            v_nodes = pt.xpath("./c:v/text()")
            if idx_attr is not None and v_nodes:
                pt_map[int(str(idx_attr))] = float(str(v_nodes[0]))
        for r_idx, expected_val in enumerate(series_values[s_idx]):
            if expected_val is None:
                if r_idx in pt_map:
                    raise MergeError(
                        f"Chart XML cache has unexpected point at series {s_idx}, idx {r_idx}."
                    )
            elif r_idx not in pt_map or not math.isclose(
                pt_map[r_idx], expected_val, rel_tol=1e-9, abs_tol=1e-9
            ):
                raise MergeError(
                    f"Chart XML <c:numCache> mismatch at series {s_idx}, idx {r_idx}: "
                    f"expected {expected_val}, got {pt_map.get(r_idx)!r}."
                )


def _update_chart(shape: GraphicFrame, tile: TileCSVData) -> None:
    if len(tile.headers) < 2:
        raise MergeError(
            f"Chart '{shape.name}' requires at least 2 columns in tile '{tile.slug}', "
            f"found {len(tile.headers)}."
        )
    if not tile.positional_rows:
        raise MergeError(f"Tile '{tile.slug}' returned 0 rows for chart '{shape.name}'.")

    categories = [row[0] if row else "" for row in tile.positional_rows]
    series_headers = tile.headers[1:]
    series_values: list[list[float | None]] = []

    chart_data = CategoryChartData()
    chart_data.categories = categories

    for col_idx in range(1, len(tile.headers)):
        vals: list[float | None] = []
        for row_idx, row in enumerate(tile.positional_rows):
            raw_cell = row[col_idx].strip() if col_idx < len(row) else ""
            if raw_cell == "":
                vals.append(None)
            else:
                try:
                    vals.append(float(raw_cell.replace(",", "")))
                except ValueError as exc:
                    raise MergeError(
                        f"Non-numeric chart value {raw_cell!r} in tile '{tile.slug}', "
                        f"column '{tile.headers[col_idx]}', row {row_idx}."
                    ) from exc
        series_values.append(vals)
        chart_data.add_series(tile.headers[col_idx], vals)

    chart = shape.chart
    chart.replace_data(chart_data)
    _verify_chart_xml_xlsx_parity(chart, categories, series_headers, series_values)


def merge_pptx(
    pptx_bytes: bytes,
    csv_zip_bytes: bytes,
    plan_meta: dict[str, str] | None = None,
    *,
    theme: str | PptxTheme | None = None,
) -> MergeResult:
    """Populate a .pptx template from Looker's unformatted csv_zip payload."""
    meta = plan_meta or {}
    tiles, raw_stems, denied_slugs = _unpack_csv_zip(csv_zip_bytes)

    schemas = [TileSchema(title=s, slug=s, columns=t.headers) for s, t in tiles.items()]
    report = inspect_and_validate(pptx_bytes, dashboard_tiles=schemas)

    referenced_slugs = {
        b.tile_slug for b in report.bindings if b.tile_slug and b.tile_slug != "_plan"
    }
    for slug in sorted(referenced_slugs):
        if slug in denied_slugs:
            raise MergeError(
                f"Tile '{slug}' contains Looker's access-denied error sentinel; aborting merge."
            )
        if f"{slug}_1" in raw_stems or f"{slug}_1" in tiles:
            raise MergeError(
                f"Duplicate dashboard tile collision detected for '{slug}' "
                f"('{slug}_1.csv' found in csv_zip)."
            )
        if slug not in tiles:
            raise MergeError(f"Required dashboard tile '{slug}' is missing from csv_zip.")
        if len(tiles[slug].rows) >= LOOKER_ROW_LIMIT:
            raise MergeError(
                f"Tile '{slug}' reached Looker's {LOOKER_ROW_LIMIT}-row dashboard export limit "
                f"({len(tiles[slug].rows)} rows); refusing to merge potentially truncated data."
            )

    if not report.valid:
        first = report.issues[0]
        raise MergeError(f"[{first.code}] {first.message}")

    prs = Presentation(io.BytesIO(pptx_bytes))
    charts_updated = 0
    tables_updated = 0
    text_tags_replaced = 0

    for slide in prs.slides:
        for shape in _iter_shapes(slide.shapes):
            name = (shape.name or "").strip()
            lower_name = name.lower()

            if lower_name.startswith("chart:") and isinstance(shape, GraphicFrame):
                slug = sanitize_tile_slug(name.split(":", 1)[1])
                _update_chart(shape, tiles[slug])
                charts_updated += 1
                continue

            if lower_name.startswith("table:") and isinstance(shape, GraphicFrame):
                slug = sanitize_tile_slug(name.split(":", 1)[1])
                text_tags_replaced += _populate_dynamic_table(shape, tiles[slug], tiles, meta)
                tables_updated += 1
                continue

            if isinstance(shape, GraphicFrame) and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for para in cell.text_frame.paragraphs:
                            text_tags_replaced += _replace_paragraph_tags(
                                para, tiles=tiles, plan_meta=meta
                            )

            if shape.has_text_frame:
                tf = getattr(shape, "text_frame")  # noqa: B009
                for para in tf.paragraphs:
                    text_tags_replaced += _replace_paragraph_tags(para, tiles=tiles, plan_meta=meta)

    resolved_theme = resolve_pptx_theme(theme)
    if resolved_theme is not None:
        apply_pptx_theme(prs, resolved_theme)

    out_buf = io.BytesIO()
    prs.save(out_buf)
    return MergeResult(
        pptx_bytes=out_buf.getvalue(),
        slides_processed=len(prs.slides),
        charts_updated=charts_updated,
        tables_updated=tables_updated,
        text_tags_replaced=text_tags_replaced,
        tile_row_counts={slug: len(tiles[slug].rows) for slug in sorted(referenced_slugs)},
    )
