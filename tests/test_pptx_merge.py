from __future__ import annotations

import io
import json
from pathlib import Path

import openpyxl
import pytest
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt

from pptx_merge import (
    PPTX_THEMES,
    MergeError,
    TileSchema,
    format_value,
    get_embed_theme_name,
    inspect_and_validate,
    merge_pptx,
    resolve_pptx_theme,
    sanitize_tile_slug,
)
from pptx_merge.cli import main as cli_main
from tests.conftest import (
    build_csv_zip_bytes,
    build_embed_demo_csv_zip_bytes,
    build_embed_demo_template_bytes,
    build_sample_template_bytes,
    build_showcase_csv_zip_bytes,
    build_showcase_template_bytes,
)


def test_sanitize_tile_slug_matches_looker_rules() -> None:
    assert sanitize_tile_slug("KPI Summary") == "kpi_summary"
    assert sanitize_tile_slug("Revenue & Margin (%) / Q1:Q4") == "revenue_&_margin_(_)___q1_q4"
    assert sanitize_tile_slug("...") == "_"
    assert sanitize_tile_slug("  Top Brands*  ") == "top_brands_"


@pytest.mark.parametrize(
    ("raw", "mask", "expected"),
    [
        ("34855339.456", "$#,##0", "$34,855,339"),
        ("34855339.456", "$#,##0.00", "$34,855,339.46"),
        ("-1234.5", "$#,##0.00", "-$1,234.50"),
        ("-0.0001", "0.00", "0.00"),
        ("3378", "#,##0", "3,378"),
        ("62.44", "0.0", "62.4"),
        ("0.5568", "0%", "56%"),
        ("0.5568", "0.0%", "55.7%"),
        ("0.5568", "0.00%", "55.68%"),
        ("2026-11-12T08:30:00Z", "yyyy-mm-dd", "2026-11-12"),
        ("Plain String", None, "Plain String"),
        (None, "$#,##0", ""),
        ("", "0.0%", ""),
    ],
)
def test_format_value(raw: str | None, mask: str | None, expected: str) -> None:
    assert format_value(raw, mask) == expected


def test_format_value_rejects_unknown_mask() -> None:
    with pytest.raises(ValueError, match="Unsupported format mask"):
        format_value("100", "eval(1)")


def test_inspect_and_validate_discovers_all_bindings(sample_template_bytes: bytes) -> None:
    schemas = [
        TileSchema(title="KPI Summary", columns=["Total Spend", "Trend Rate"]),
        TileSchema(title="Mixed Metrics", columns=["Metric", "Value"]),
        TileSchema(title="Top Brands", columns=["Brand Name", "Units Sold", "Net Revenue"]),
        TileSchema(title="Monthly Trend", columns=["Quarter", "Direct", "Marketplace"]),
    ]
    report = inspect_and_validate(sample_template_bytes, dashboard_tiles=schemas)
    assert report.valid is True
    assert report.slide_count == 1
    assert report.issues == []

    kinds = [b.kind for b in report.bindings]
    assert "chart" in kinds
    assert "table" in kinds
    assert "text" in kinds


def test_inspect_and_validate_catches_chart_ex_and_schema_mismatches() -> None:
    pptx_with_cx = build_sample_template_bytes(include_chart_ex=True)
    incomplete_schemas = [
        TileSchema(title="KPI Summary", columns=["Total Spend"]),  # missing Trend Rate
        TileSchema(title="Monthly Trend", columns=["Quarter"]),  # only 1 column for chart
    ]
    report = inspect_and_validate(pptx_with_cx, dashboard_tiles=incomplete_schemas)
    assert report.valid is False
    codes = {issue.code for issue in report.issues}
    assert "unsupported_chart_ex" in codes
    assert "missing_tile" in codes
    assert "missing_column" in codes
    assert "insufficient_chart_columns" in codes


def test_merge_pptx_end_to_end_and_xlsx_parity(
    sample_template_bytes: bytes,
    sample_csv_zip_bytes: bytes,
) -> None:
    res = merge_pptx(
        sample_template_bytes,
        sample_csv_zip_bytes,
        plan_meta={"title": "Q4 Executive Review", "url": "https://looker.example/dashboards/12"},
    )
    assert res.slides_processed == 1
    assert res.charts_updated == 1
    assert res.tables_updated == 1
    assert res.text_tags_replaced > 0
    assert res.tile_row_counts == {
        "kpi_summary": 1,
        "mixed_metrics": 3,
        "monthly_trend": 4,
        "top_brands": 3,
    }

    prs = Presentation(io.BytesIO(res.pptx_bytes))
    slide = prs.slides[0]
    shapes_by_name = {s.name: s for s in slide.shapes}

    # 1. Verify split-run text box replacement and preserved <a:rPr> styling on first run
    callout = shapes_by_name["Header Callout"]
    para = callout.text_frame.paragraphs[0]
    assert para.text == "Plan: Q4 Executive Review | Spend: $34,855,339.46 | Rate: 8.4%"
    assert len(para.runs) == 1
    assert para.runs[0].font.name == "Arial"
    assert para.runs[0].font.size == Pt(16)
    assert para.runs[0].font.bold is True
    assert para.runs[0].font.color.rgb == RGBColor(0x1A, 0x73, 0xE8)

    # 2. Verify fixed-row mixed-unit table cells
    ft = shapes_by_name["Fixed Metrics Table"].table
    assert ft.cell(1, 1).text == "3,378"
    assert ft.cell(2, 1).text == "62.4"
    assert ft.cell(3, 1).text == "55.7%"

    # 3. Verify dynamic table cloned 3 data rows + 1 header row = 4 rows
    dt = shapes_by_name["table:top_brands"].table
    assert len(dt.rows) == 4
    assert [dt.cell(1, c).text for c in range(3)] == ["Apex Goods", "120", "$450,200"]
    assert [dt.cell(2, c).text for c in range(3)] == ["Nova Apparel", "340", "$289,100"]
    assert [dt.cell(3, c).text for c in range(3)] == ["Vertex Active", "890", "$195,000"]

    # 4. Verify chart XML cache and embedded openpyxl .xlsx worksheet
    chart = shapes_by_name["chart:monthly_trend"].chart
    xlsx_bytes = chart.part.chart_workbook.xlsx_part.blob
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb.active
    assert ws is not None
    assert ws.cell(1, 2).value == "Direct"
    assert ws.cell(1, 3).value == "Marketplace"
    assert ws.cell(2, 1).value == "2026-Q1"
    assert ws.cell(2, 2).value == pytest.approx(1250000.5)
    assert ws.cell(4, 3).value is None  # Blank cell preserved as None in Q3 Marketplace
    assert ws.cell(5, 3).value == pytest.approx(390000.5)
    wb.close()


def test_merge_pptx_fails_on_missing_tile(
    sample_template_bytes: bytes,
) -> None:
    incomplete_zip = build_csv_zip_bytes({"kpi_summary.csv": "Total Spend,Trend Rate\n100,0.05\n"})
    with pytest.raises(MergeError, match="missing from csv_zip"):
        merge_pptx(sample_template_bytes, incomplete_zip)


def test_merge_pptx_fails_on_access_denied_sentinel(
    sample_template_bytes: bytes,
) -> None:
    denied_zip = build_csv_zip_bytes(
        {
            "kpi_summary.csv": "You do not have access to this data.",
            "mixed_metrics.csv": "Metric,Value\nA,1\nB,2\nC,3\n",
            "top_brands.csv": "Brand Name,Units Sold,Net Revenue\nX,1,1\n",
            "monthly_trend.csv": "Quarter,Direct,Marketplace\nQ1,1,2\n",
        }
    )
    with pytest.raises(MergeError, match="access-denied"):
        merge_pptx(sample_template_bytes, denied_zip)


def test_merge_pptx_fails_on_duplicate_tile_collision(
    sample_template_bytes: bytes,
) -> None:
    collision_zip = build_csv_zip_bytes(
        {
            "kpi_summary.csv": "Total Spend,Trend Rate\n100,0.05\n",
            "kpi_summary_1.csv": "Total Spend,Trend Rate\n200,0.06\n",
            "mixed_metrics.csv": "Metric,Value\nA,1\nB,2\nC,3\n",
            "top_brands.csv": "Brand Name,Units Sold,Net Revenue\nX,1,1\n",
            "monthly_trend.csv": "Quarter,Direct,Marketplace\nQ1,1,2\n",
        }
    )
    with pytest.raises(MergeError, match="Duplicate dashboard tile collision"):
        merge_pptx(sample_template_bytes, collision_zip)


def test_merge_pptx_fails_on_5000_row_ceiling(
    sample_template_bytes: bytes,
) -> None:
    rows_5000 = "Brand Name,Units Sold,Net Revenue\n" + "".join(
        f"Brand_{i},10,100\n" for i in range(5000)
    )
    capped_zip = build_csv_zip_bytes(
        {
            "kpi_summary.csv": "Total Spend,Trend Rate\n100,0.05\n",
            "mixed_metrics.csv": "Metric,Value\nA,1\nB,2\nC,3\n",
            "top_brands.csv": rows_5000,
            "monthly_trend.csv": "Quarter,Direct,Marketplace\nQ1,1,2\n",
        }
    )
    with pytest.raises(MergeError, match="5000-row dashboard export limit"):
        merge_pptx(sample_template_bytes, capped_zip)


def test_cli_inspect_validate_and_merge(
    tmp_path: Path,
    sample_template_bytes: bytes,
    sample_csv_zip_bytes: bytes,
    capsys: pytest.CaptureFixture[str],
) -> None:
    deck_path = tmp_path / "template.pptx"
    zip_path = tmp_path / "export.zip"
    out_path = tmp_path / "merged.pptx"
    deck_path.write_bytes(sample_template_bytes)
    zip_path.write_bytes(sample_csv_zip_bytes)

    assert cli_main(["inspect", str(deck_path)]) == 0
    inspect_out = json.loads(capsys.readouterr().out)
    assert inspect_out["valid"] is True

    assert cli_main(["validate", str(deck_path), "--csv-zip", str(zip_path)]) == 0
    val_out = json.loads(capsys.readouterr().out)
    assert val_out["valid"] is True

    assert (
        cli_main(
            [
                "merge",
                str(deck_path),
                "--csv-zip",
                str(zip_path),
                "-o",
                str(out_path),
                "--title",
                "CLI Test Deck",
            ]
        )
        == 0
    )
    merge_out = json.loads(capsys.readouterr().out)
    assert merge_out["success"] is True
    assert out_path.exists()

    schema_path = tmp_path / "schema.json"
    schema_path.write_text(
        json.dumps(
            {
                "dashboard_tiles": [
                    {"title": "KPI Summary", "columns": ["Total Spend", "Trend Rate"]},
                    {"title": "Mixed Metrics", "columns": ["Metric", "Value"]},
                    {
                        "title": "Top Brands",
                        "columns": ["Brand Name", "Units Sold", "Net Revenue"],
                    },
                    {"title": "Monthly Trend", "columns": ["Quarter", "Direct", "Marketplace"]},
                ]
            }
        ),
        encoding="utf-8",
    )
    assert cli_main(["validate", str(deck_path), "--schema", str(schema_path)]) == 0
    capsys.readouterr()

    # Missing file returns exit code 1
    assert cli_main(["inspect", str(tmp_path / "missing.pptx")]) == 1
    capsys.readouterr()


def test_positional_table_and_group_shape(sample_csv_zip_bytes: bytes) -> None:
    from pptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # Group shape with text tag
    grp = slide.shapes.add_group_shape()
    box = grp.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    box.text_frame.text = "Grouped: {{kpi_summary.Total Spend|$#,##0}}"
    # Positional table (no {{}} tags in prototype row)
    tbl_shape = slide.shapes.add_table(2, 3, Inches(1), Inches(2), Inches(6), Inches(1.5))
    tbl_shape.name = "table:top_brands"
    tbl = tbl_shape.table
    tbl.cell(0, 0).text = "Brand"
    tbl.cell(0, 1).text = "Units"
    tbl.cell(0, 2).text = "Revenue"
    # Prototype row with empty/styled run
    tbl.cell(1, 0).text = "placeholder"

    buf = io.BytesIO()
    prs.save(buf)
    res = merge_pptx(buf.getvalue(), sample_csv_zip_bytes)
    assert res.tables_updated == 1
    assert res.text_tags_replaced == 1
    out_prs = Presentation(io.BytesIO(res.pptx_bytes))
    for s in out_prs.slides[0].shapes:
        if s.has_table:
            assert s.table.cell(1, 0).text == "Apex Goods"
            assert s.table.cell(2, 0).text == "Nova Apparel"


def test_showcase_deck_all_patterns_and_charts() -> None:
    tmpl = build_showcase_template_bytes()
    csv_zip = build_showcase_csv_zip_bytes()
    res = merge_pptx(
        tmpl,
        csv_zip,
        plan_meta={
            "title": "Q3 Executive Business Review",
            "url": "https://dev-argolis.looker.com/dashboards/81",
            "run_at": "2026-11-13T17:15:00Z",
        },
    )
    assert res.slides_processed == 3
    assert res.charts_updated == 5
    assert res.tables_updated == 2
    assert res.text_tags_replaced >= 35
    assert len(res.tile_row_counts) == 9


def test_embed_demo_deck_and_dynamic_theme_recoloring() -> None:
    assert get_embed_theme_name(True, "Levi's") == "Levis_Dark"
    assert get_embed_theme_name(False, "Levi's") == "Levis_Light"
    assert get_embed_theme_name(False, "Calvin Klein") == "Calvin_Klein_Light"
    assert get_embed_theme_name(True, "Allegra K") == "Allegra_K_Dark"
    assert get_embed_theme_name(False, "") == "Embed_Demo_Light"

    assert (
        resolve_pptx_theme("auto", brand="Columbia", is_dark=False)
        == (PPTX_THEMES["Columbia_Light"])
    )
    assert resolve_pptx_theme("Levis_Light") == PPTX_THEMES["Levis_Light"]

    tmpl = build_embed_demo_template_bytes()
    csv_zip = build_embed_demo_csv_zip_bytes()
    res_light = merge_pptx(
        tmpl,
        csv_zip,
        plan_meta={
            "title": "Levi's Brand Overview",
            "url": "https://looker.example.com/dashboards/embed_demo::brand_overview",
            "run_at": "2026-11-15T05:00:00Z",
        },
        theme="Levis_Light",
    )
    assert res_light.slides_processed == 3
    assert res_light.charts_updated == 6
    assert res_light.tables_updated == 1
    assert res_light.text_tags_replaced >= 55
    assert len(res_light.tile_row_counts) == 19

    prs_light = Presentation(io.BytesIO(res_light.pptx_bytes))
    slide1_xml = prs_light.slides[0]._element.xml
    assert "F0F4F9" in slide1_xml  # Light mode background
    assert "FFFFFF" in slide1_xml  # Light mode bento card surface
    assert "C41230" in slide1_xml  # Levi's Crimson primary accent
