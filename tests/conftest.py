from __future__ import annotations

import io
import zipfile

import pytest
from pptx import Presentation
from pptx.chart.chart import Chart
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.slide import Slide
from pptx.table import Table, _Cell
from pptx.util import Inches, Pt

# Modern Executive "Obsidian & Electric Cyan" Palette
BG_DARK = RGBColor(0x09, 0x0D, 0x16)
CARD_BG = RGBColor(0x11, 0x18, 0x27)
CARD_BORDER = RGBColor(0x1E, 0x29, 0x3B)
ROW_ALT_A = RGBColor(0x0F, 0x17, 0x2A)
ROW_ALT_B = RGBColor(0x16, 0x20, 0x32)
TEXT_WHITE = RGBColor(0xF8, 0xFA, 0xFC)
TEXT_LIGHT = RGBColor(0xE2, 0xE8, 0xF0)
TEXT_MUTED = RGBColor(0x94, 0xA3, 0xB8)
TEXT_DIM = RGBColor(0x64, 0x74, 0x8B)
ACCENT_CYAN = RGBColor(0x38, 0xBD, 0xF8)
ACCENT_INDIGO = RGBColor(0x81, 0x8C, 0xF8)
ACCENT_EMERALD = RGBColor(0x34, 0xD3, 0x99)
ACCENT_AMBER = RGBColor(0xFB, 0xBF, 0x24)
ACCENT_ROSE = RGBColor(0xFB, 0x71, 0x85)
FONT_SANS = "Calibri"
NO_STYLE_TABLE_GUID = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"


def build_sample_template_bytes(*, include_chart_ex: bool = False) -> bytes:
    """Create an in-memory .pptx with split-run text tags, fixed table, dynamic table, and chart."""
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 1. Text box with split XML <a:r> runs for {{kpi_summary...}} and {{_plan.title}}
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(8.0), Inches(1.0))
    tb.name = "Header Callout"
    p = tb.text_frame.paragraphs[0]
    r1 = p.add_run()
    r1.text = "Plan: {{_plan.title}} | Spend: {{kpi_"
    r1.font.name = "Arial"
    r1.font.size = Pt(16)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(0x1A, 0x73, 0xE8)
    r2 = p.add_run()
    r2.text = "summary.Total Spend | $#,##0.00}} | Rate: {{kpi_summary[0].Trend Rate | 0.0%}}"

    # 2. Fixed-row mixed-unit table (counts, decimals, percentages in the same column)
    fixed_shape = slide.shapes.add_table(4, 2, Inches(0.5), Inches(1.6), Inches(4.0), Inches(1.5))
    fixed_shape.name = "Fixed Metrics Table"
    ft = fixed_shape.table
    ft.cell(0, 0).text = "Metric"
    ft.cell(0, 1).text = "Current"
    ft.cell(1, 0).text = "Active Customers"
    ft.cell(1, 1).text = "{{mixed_metrics[0].Value | #,##0}}"
    ft.cell(2, 0).text = "Orders per Customer"
    ft.cell(2, 1).text = "{{mixed_metrics[1].Value | 0.0}}"
    ft.cell(3, 0).text = "Gross Margin Rate"
    ft.cell(3, 1).text = "{{mixed_metrics[2].Value | 0.0%}}"

    # 3. Dynamic N-row table bound via Selection Pane name "table:top_brands"
    dyn_shape = slide.shapes.add_table(2, 3, Inches(5.0), Inches(1.6), Inches(4.5), Inches(1.5))
    dyn_shape.name = "table:top_brands"
    dt = dyn_shape.table
    dt.cell(0, 0).text = "Brand Name"
    dt.cell(0, 1).text = "Units Sold"
    dt.cell(0, 2).text = "Net Revenue"
    dt.cell(1, 0).text = "{{.Brand Name}}"
    dt.cell(1, 1).text = "{{.Units Sold | #,##0}}"
    dt.cell(1, 2).text = "{{.Net Revenue | $#,##0}}"

    # 4. Native Office Column Chart bound via Selection Pane name "chart:monthly_trend"
    chart_data = CategoryChartData()
    chart_data.categories = ["Q1", "Q2"]
    chart_data.add_series("Direct", (100.0, 200.0))
    chart_data.add_series("Marketplace", (50.0, 75.0))
    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.5),
        Inches(3.5),
        Inches(8.5),
        Inches(3.2),
        chart_data,
    )
    chart_shape.name = "chart:monthly_trend"

    buf = io.BytesIO()
    prs.save(buf)
    raw_pptx = buf.getvalue()

    if not include_chart_ex:
        return raw_pptx

    out_buf = io.BytesIO()
    with (
        zipfile.ZipFile(io.BytesIO(raw_pptx), "r") as zin,
        zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout,
    ):
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        zout.writestr(
            "ppt/charts/chartEx1.xml",
            b'<?xml version="1.0" encoding="UTF-8"?><cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>',
        )
    return out_buf.getvalue()


def build_csv_zip_bytes(files: dict[str, str]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, content in files.items():
            zf.writestr(f"dashboard-export/{filename}", content.encode("utf-8"))
    return buf.getvalue()


@pytest.fixture
def sample_template_bytes() -> bytes:
    return build_sample_template_bytes()


@pytest.fixture
def sample_csv_zip_bytes() -> bytes:
    return build_csv_zip_bytes(
        {
            "kpi_summary.csv": "Total Spend,Trend Rate\n34855339.456,0.0842\n",
            "mixed_metrics.csv": (
                "Metric,Value\nCustomers,3378\nOrders per Customer,62.44\nGross Margin,0.5568\n"
            ),
            "top_brands.csv": (
                "Brand Name,Units Sold,Net Revenue\n"
                "Apex Goods,120,450200.4\n"
                "Nova Apparel,340,289100.2\n"
                "Vertex Active,890,195000.0\n"
            ),
            "monthly_trend.csv": (
                "Quarter,Direct,Marketplace\n"
                "2026-Q1,1250000.5,320000.25\n"
                "2026-Q2,1310000.0,340500.75\n"
                "2026-Q3,1290000.0,\n"
                "2026-Q4,1420000.25,390000.5\n"
            ),
        }
    )


def _style_cell(
    cell: _Cell,
    text: str,
    *,
    bg: RGBColor,
    fg: RGBColor = TEXT_LIGHT,
    size: float = 10.0,
    bold: bool = False,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    border_hex: str = "1E293B",
) -> None:
    cell.fill.solid()
    cell.fill.fore_color.rgb = bg
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Inches(0.12)
    cell.margin_right = Inches(0.12)
    cell.margin_top = Inches(0.06)
    cell.margin_bottom = Inches(0.06)
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    p.text = ""
    run = p.add_run()
    run.text = text
    run.font.name = FONT_SANS
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = fg

    tc_pr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT"):
        old = tc_pr.find(qn(tag))
        if old is not None:
            tc_pr.remove(old)
        ln = tc_pr.makeelement(qn(tag), {"w": "0"})
        ln.append(tc_pr.makeelement(qn("a:noFill"), {}))
        tc_pr.append(ln)
    old_b = tc_pr.find(qn("a:lnB"))
    if old_b is not None:
        tc_pr.remove(old_b)
    ln_b = tc_pr.makeelement(qn("a:lnB"), {"w": "9525"})
    solid = tc_pr.makeelement(qn("a:solidFill"), {})
    solid.append(tc_pr.makeelement(qn("a:srgbClr"), {"val": border_hex}))
    ln_b.append(solid)
    tc_pr.append(ln_b)


def _strip_default_table_style(table: Table) -> None:
    tbl_pr = table._tbl.find(qn("a:tblPr"))
    if tbl_pr is not None:
        tbl_pr.attrib.clear()
        style_id = tbl_pr.find(qn("a:tableStyleId"))
        if style_id is not None:
            style_id.text = NO_STYLE_TABLE_GUID


def _add_bento_card(
    slide: Slide,
    left: float,
    top: float,
    width: float,
    height: float,
    *,
    title: str = "",
    subtitle: str = "",
    accent: RGBColor | None = None,
) -> None:
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    card.adjustments[0] = 0.02
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = CARD_BORDER
    card.line.width = Pt(1)

    if accent is not None:
        bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(left + 0.18),
            Inches(top),
            Inches(width - 0.36),
            Inches(0.045),
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = accent
        bar.line.fill.background()

    if title:
        tb = slide.shapes.add_textbox(
            Inches(left + 0.22),
            Inches(top + 0.14),
            Inches(width - 0.44),
            Inches(0.55),
        )
        tf = tb.text_frame
        tf.word_wrap = True
        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = title
        r0.font.name = FONT_SANS
        r0.font.size = Pt(13)
        r0.font.bold = True
        r0.font.color.rgb = TEXT_WHITE
        if subtitle:
            p1 = tf.add_paragraph()
            r1 = p1.add_run()
            r1.text = subtitle
            r1.font.name = FONT_SANS
            r1.font.size = Pt(9)
            r1.font.color.rgb = TEXT_MUTED


def _add_kpi_strip(
    slide: Slide,
    specs: list[tuple[float, RGBColor, str, str, str]],
    *,
    grouped: bool = False,
) -> None:
    for x_pos, accent_clr, _, _, _ in specs:
        _add_bento_card(slide, x_pos, 1.16, 2.92, 1.32, accent=accent_clr)
    container = slide.shapes.add_group_shape() if grouped else slide
    if grouped:
        container.name = "Executive KPI Group"
    for idx, (x_pos, accent_clr, label, big_val, sub_val) in enumerate(specs):
        card_tb = container.shapes.add_textbox(
            Inches(x_pos + 0.18), Inches(1.26), Inches(2.56), Inches(1.12)
        )
        card_tb.name = f"KPI Card {idx + 1}"
        tf = card_tb.text_frame
        tf.word_wrap = True
        for p_idx, (txt, sz, clr) in enumerate(
            [(label, 8.5, TEXT_MUTED), (big_val, 20.0, TEXT_WHITE), (sub_val, 8.5, accent_clr)]
        ):
            p = tf.paragraphs[0] if p_idx == 0 else tf.add_paragraph()
            r = p.add_run()
            r.text = txt
            r.font.name = FONT_SANS
            r.font.size = Pt(sz)
            r.font.bold = True
            r.font.color.rgb = clr


def _add_slide_chrome(
    slide: Slide,
    *,
    slide_num: int,
    eyebrow: str,
) -> None:
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5)
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG_DARK
    bg.line.fill.background()

    top_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.055)
    )
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = ACCENT_CYAN
    top_bar.line.fill.background()

    pill = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(0.18), Inches(3.35), Inches(0.26)
    )
    pill.adjustments[0] = 0.5
    pill.fill.solid()
    pill.fill.fore_color.rgb = CARD_BORDER
    pill.line.color.rgb = ACCENT_CYAN
    pill.line.width = Pt(0.75)
    p_pill = pill.text_frame.paragraphs[0]
    p_pill.alignment = PP_ALIGN.CENTER
    r_pill = p_pill.add_run()
    r_pill.text = eyebrow
    r_pill.font.name = FONT_SANS
    r_pill.font.size = Pt(8)
    r_pill.font.bold = True
    r_pill.font.color.rgb = ACCENT_CYAN

    status = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.45), Inches(0.18), Inches(2.38), Inches(0.26)
    )
    status.adjustments[0] = 0.5
    status.fill.solid()
    status.fill.fore_color.rgb = RGBColor(0x0B, 0x26, 0x22)
    status.line.color.rgb = ACCENT_EMERALD
    status.line.width = Pt(0.75)
    p_st = status.text_frame.paragraphs[0]
    p_st.alignment = PP_ALIGN.CENTER
    r_st = p_st.add_run()
    r_st.text = "LIVE  •  {{_plan.run_at | yyyy-mm-dd}}"
    r_st.font.name = FONT_SANS
    r_st.font.size = Pt(8)
    r_st.font.bold = True
    r_st.font.color.rgb = ACCENT_EMERALD

    footer = slide.shapes.add_textbox(Inches(0.5), Inches(7.12), Inches(10.0), Inches(0.28))
    p_f = footer.text_frame.paragraphs[0]
    r_f1 = p_f.add_run()
    r_f1.text = "Source: {{_plan.url}}   •   Looker PowerPoint Action Hub (action-pptx)"
    r_f1.font.name = FONT_SANS
    r_f1.font.size = Pt(8.5)
    r_f1.font.color.rgb = TEXT_DIM

    page_tb = slide.shapes.add_textbox(Inches(10.83), Inches(7.12), Inches(2.0), Inches(0.28))
    p_pg = page_tb.text_frame.paragraphs[0]
    p_pg.alignment = PP_ALIGN.RIGHT
    r_f2 = p_pg.add_run()
    r_f2.text = f"SLIDE 0{slide_num} / 03"
    r_f2.font.name = FONT_SANS
    r_f2.font.size = Pt(8.5)
    r_f2.font.bold = True
    r_f2.font.color.rgb = TEXT_MUTED


def _style_modern_chart(
    chart: Chart,
    series_colors: list[RGBColor],
    *,
    is_doughnut: bool = False,
) -> None:
    chart.has_title = True
    chart.chart_title.text_frame.text = " "
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.name = FONT_SANS
    chart.legend.font.size = Pt(8.5)
    chart.legend.font.color.rgb = TEXT_LIGHT

    # Remove outer chart frame border and make chart background transparent
    c_space = chart._chartSpace
    sp_pr = c_space.find(qn("c:spPr"))
    if sp_pr is not None:
        c_space.remove(sp_pr)
    sp_pr = c_space.makeelement(qn("c:spPr"), {})
    sp_pr.append(c_space.makeelement(qn("a:noFill"), {}))
    ln = c_space.makeelement(qn("a:ln"), {})
    ln.append(c_space.makeelement(qn("a:noFill"), {}))
    sp_pr.append(ln)
    c_space.append(sp_pr)

    if is_doughnut:
        for idx, color in enumerate(series_colors):
            pt = chart.series[0].points[idx]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = color
        return

    for idx, color in enumerate(series_colors):
        if idx < len(chart.series):
            ser = chart.series[idx]
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = color
            ser.format.line.color.rgb = color
            ser.format.line.width = Pt(2.25)

    for ax_attr in ("category_axis", "value_axis"):
        ax = getattr(chart, ax_attr, None)
        if ax is not None:
            ax.tick_labels.font.name = FONT_SANS
            ax.tick_labels.font.size = Pt(8)
            ax.tick_labels.font.color.rgb = TEXT_MUTED
            ax.format.line.color.rgb = CARD_BORDER
            if ax.has_major_gridlines:
                ax.major_gridlines.format.line.color.rgb = CARD_BORDER


def build_showcase_template_bytes() -> bytes:
    """Build a 3-slide modern dark-mode widescreen showcase .pptx covering all spec patterns."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 1: Executive Summary, Bento KPI Strip, Benchmark Matrix & Donut
    # =========================================================================
    s1 = prs.slides.add_slide(blank)
    _add_slide_chrome(s1, slide_num=1, eyebrow="EXECUTIVE OVERVIEW  •  TWO-PERIOD BENCHMARK")

    tb1 = s1.shapes.add_textbox(Inches(0.5), Inches(0.48), Inches(12.33), Inches(0.62))
    tb1.name = "Header Callout"
    p1 = tb1.text_frame.paragraphs[0]
    r1 = p1.add_run()
    r1.text = "{{_plan.title}}   |   Spend: {{kpi_"
    r1.font.name = FONT_SANS
    r1.font.size = Pt(18)
    r1.font.bold = True
    r1.font.color.rgb = TEXT_WHITE
    r2 = p1.add_run()
    r2.text = (
        "summary.Total Spend | $#,##0.00}}   •   YoY Trend: +{{kpi_summary[0].Trend Rate | 0.0%}} "
        "  •   Avg Order Value: {{kpi_summary.Avg Order Value | $#,##0.00}}"
    )
    r2.font.name = FONT_SANS
    r2.font.size = Pt(13)
    r2.font.bold = True
    r2.font.color.rgb = ACCENT_CYAN

    # 4 Hero KPI Bento Cards across the top (y=1.16", h=1.32")
    kpi_specs = [
        (
            0.50,
            ACCENT_CYAN,
            "TOTAL REVENUE SPEND",
            "{{kpi_summary.Total Spend | $#,##0}}",
            "Exact: {{kpi_summary.Total Spend | $#,##0.00}}",
        ),
        (
            3.63,
            ACCENT_EMERALD,
            "AVERAGE ORDER VALUE",
            "{{kpi_summary.Avg Order Value | $#,##0.00}}",
            "YoY Trend: +{{kpi_summary[0].Trend Rate | 0.0%}} vs. Prior",
        ),
        (
            6.77,
            ACCENT_INDIGO,
            "ACTIVE CUSTOMERS",
            "{{mixed_metrics[0].Current | #,##0}}",
            "Prior: {{mixed_metrics[0].Prior | #,##0}} (+{{mixed_metrics[0].Delta | 0.0%}})",
        ),
        (
            9.91,
            ACCENT_AMBER,
            "GROSS MARGIN RATE",
            "{{mixed_metrics[2].Current | 0.0%}}",
            "Orders/Cust: {{mixed_metrics[1].Current | 0.0}} ({{mixed_metrics[1].Delta | 0.0%}})",
        ),
    ]
    _add_kpi_strip(s1, kpi_specs, grouped=True)

    # Left Bento Card: Two-Period Performance Benchmark Matrix
    _add_bento_card(
        s1,
        0.50,
        2.66,
        7.55,
        4.30,
        title="Two-Period Operational & Financial Benchmarks",
        subtitle="Mixed-unit KPI matrix formatted inline per cell (counts, rates, %, currency)",
    )
    ft_shape = s1.shapes.add_table(5, 4, Inches(0.72), Inches(3.38), Inches(7.11), Inches(3.36))
    ft_shape.name = "Fixed Metrics Table"
    ft = ft_shape.table
    _strip_default_table_style(ft)
    col_widths_s1 = [Inches(2.55), Inches(1.52), Inches(1.52), Inches(1.52)]
    for c_idx, w in enumerate(col_widths_s1):
        ft.columns[c_idx].width = w
    for c_idx, h_text in enumerate(
        ["BENCHMARK METRIC", "PRIOR PERIOD", "CURRENT PERIOD", "YOY DELTA"]
    ):
        _style_cell(
            ft.cell(0, c_idx),
            h_text,
            bg=CARD_BORDER,
            fg=ACCENT_CYAN,
            size=9.0,
            bold=True,
            align=PP_ALIGN.LEFT if c_idx == 0 else PP_ALIGN.RIGHT,
            border_hex="38BDF8",
        )
    rows_spec = [
        (
            "Total Active Customers",
            "{{mixed_metrics[0].Prior | #,##0}}",
            "{{mixed_metrics[0].Current | #,##0}}",
            "+{{mixed_metrics[0].Delta | 0.0%}}",
        ),
        (
            "Orders per 1,000 Customers",
            "{{mixed_metrics[1].Prior | 0.0}}",
            "{{mixed_metrics[1].Current | 0.0}}",
            "{{mixed_metrics[1].Delta | 0.0%}}",
        ),
        (
            "Gross Margin Rate (GMR)",
            "{{mixed_metrics[2].Prior | 0.0%}}",
            "{{mixed_metrics[2].Current | 0.0%}}",
            "+{{mixed_metrics[2].Delta | 0.00%}}",
        ),
        (
            "Average Order Value (AOV)",
            "{{mixed_metrics[3].Prior | $#,##0.00}}",
            "{{mixed_metrics[3].Current | $#,##0.00}}",
            "+{{mixed_metrics[3].Delta | 0.0%}}",
        ),
    ]
    for r_idx, r_vals in enumerate(rows_spec, start=1):
        row_bg = ROW_ALT_A if r_idx % 2 == 1 else ROW_ALT_B
        for c_idx, val in enumerate(r_vals):
            _style_cell(
                ft.cell(r_idx, c_idx),
                val,
                bg=row_bg,
                fg=(
                    TEXT_WHITE
                    if c_idx in (0, 2)
                    else (ACCENT_EMERALD if c_idx == 3 else TEXT_MUTED)
                ),
                size=10.5,
                bold=c_idx in (0, 2, 3),
                align=PP_ALIGN.LEFT if c_idx == 0 else PP_ALIGN.RIGHT,
            )

    # Right Bento Card: Spend by Product Category Doughnut Chart
    _add_bento_card(
        s1,
        8.25,
        2.66,
        4.58,
        4.30,
        title="Revenue Share by Category",
        subtitle="Native Office Doughnut Chart  •  100% XML/XLSX Parity",
    )
    donut_data = CategoryChartData()
    donut_data.categories = ["Apparel", "Footwear", "Accessories", "Outerwear"]
    donut_data.add_series("Total Spend", (11250000.0, 9840000.0, 6415339.46, 7350000.0))
    donut_shape = s1.shapes.add_chart(
        XL_CHART_TYPE.DOUGHNUT,
        Inches(8.42),
        Inches(3.32),
        Inches(4.24),
        Inches(3.48),
        donut_data,
    )
    donut_shape.name = "chart:spend_by_category"
    _style_modern_chart(
        donut_shape.chart,
        [ACCENT_CYAN, ACCENT_INDIGO, ACCENT_EMERALD, ACCENT_AMBER],
        is_doughnut=True,
    )

    # =========================================================================
    # SLIDE 2: Multi-Series Column, Line with Markers, and Stacked Column Bridge
    # =========================================================================
    s2 = prs.slides.add_slide(blank)
    _add_slide_chrome(
        s2, slide_num=2, eyebrow="TREND ANALYTICS  •  QUARTERLY TRAJECTORY & VALUE BRIDGE"
    )

    s2_title = s2.shapes.add_textbox(Inches(0.5), Inches(0.48), Inches(12.33), Inches(0.62))
    s2_title.name = "Slide 2 Title"
    p2 = s2_title.text_frame.paragraphs[0]
    r2_title = p2.add_run()
    r2_title.text = "{{_plan.title}}   —   Multi-Series Quarterly Trajectory & AOV Value Bridge"
    r2_title.font.name = FONT_SANS
    r2_title.font.size = Pt(18)
    r2_title.font.bold = True
    r2_title.font.color.rgb = TEXT_WHITE

    # 1. Clustered Column Chart (with sparse None handling)
    _add_bento_card(
        s2,
        0.50,
        1.18,
        3.95,
        5.75,
        title="Quarterly Direct vs. Marketplace",
        subtitle="Clustered Column  •  Sparse Null Handling",
        accent=ACCENT_CYAN,
    )
    col_data = CategoryChartData()
    col_data.categories = ["2026-Q1", "2026-Q2", "2026-Q3", "2026-Q4"]
    col_data.add_series("Direct", (6450000.5, 6810000.0, 6590000.0, 7655339.0))
    col_data.add_series("Marketplace", (1620000.25, 1740500.75, 1680000.0, 1989499.0))
    col_shape = s2.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.65),
        Inches(1.88),
        Inches(3.65),
        Inches(4.85),
        col_data,
    )
    col_shape.name = "chart:monthly_trend"
    _style_modern_chart(col_shape.chart, [ACCENT_CYAN, ACCENT_INDIGO])

    # 2. Line Chart with Markers (3-series AOV trajectory)
    _add_bento_card(
        s2,
        4.69,
        1.18,
        3.95,
        5.75,
        title="AOV Trajectory vs. Benchmark",
        subtitle="3-Series Line with Markers  •  QoQ Trend",
        accent=ACCENT_EMERALD,
    )
    line_data = CategoryChartData()
    line_data.categories = ["2026-Q1", "2026-Q2", "2026-Q3", "2026-Q4"]
    line_data.add_series("Prior Year AOV", (392.1, 401.5, 408.2, 417.4))
    line_data.add_series("Current Year AOV", (418.4, 432.15, 439.8, 465.3))
    line_data.add_series("Target Benchmark", (405.0, 410.0, 414.5, 420.0))
    line_shape = s2.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS,
        Inches(4.84),
        Inches(1.88),
        Inches(3.65),
        Inches(4.85),
        line_data,
    )
    line_shape.name = "chart:aov_trend"
    _style_modern_chart(line_shape.chart, [TEXT_MUTED, ACCENT_CYAN, ACCENT_EMERALD])

    # 3. Stacked Column Chart (Standard <c:chartSpace> Waterfall / Value Bridge)
    _add_bento_card(
        s2,
        8.88,
        1.18,
        3.95,
        5.75,
        title="AOV Value Driver Bridge",
        subtitle="Standard <c:chartSpace> Stacked Bridge",
        accent=ACCENT_AMBER,
    )
    bridge_data = CategoryChartData()
    bridge_data.categories = [
        "Prior AOV",
        "Volume Mix",
        "Unit Price",
        "Premium Mix",
        "Current AOV",
    ]
    bridge_data.add_series("Base", (0.0, 404.8, 419.0, 430.6, 0.0))
    bridge_data.add_series("Impact", (404.8, 14.2, 11.6, 8.32, 438.92))
    bridge_shape = s2.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_STACKED,
        Inches(9.03),
        Inches(1.88),
        Inches(3.65),
        Inches(4.85),
        bridge_data,
    )
    bridge_shape.name = "chart:cost_bridge"
    _style_modern_chart(bridge_shape.chart, [CARD_BORDER, ACCENT_AMBER])

    # =========================================================================
    # SLIDE 3: Horizontal Bar Chart (regional_customers) + Tagged & Positional Tables
    # =========================================================================
    s3 = prs.slides.add_slide(blank)
    _add_slide_chrome(s3, slide_num=3, eyebrow="REGIONAL & BRAND DRIVERS  •  DYNAMIC N-ROW TABLES")

    s3_title = s3.shapes.add_textbox(Inches(0.5), Inches(0.48), Inches(12.33), Inches(0.62))
    s3_title.name = "Slide 3 Title"
    p3 = s3_title.text_frame.paragraphs[0]
    r3_title = p3.add_run()
    r3_title.text = "{{_plan.title}}   —   Regional Growth, Top Brands & Enterprise Accounts"
    r3_title.font.name = FONT_SANS
    r3_title.font.size = Pt(18)
    r3_title.font.bold = True
    r3_title.font.color.rgb = TEXT_WHITE

    # 1. Clustered Horizontal Bar Chart (regional_customers)
    _add_bento_card(
        s3,
        0.50,
        1.18,
        4.85,
        5.75,
        title="Active Customers by Region",
        subtitle="Prior vs. Current Period Customers  •  Horizontal Bar",
        accent=ACCENT_INDIGO,
    )
    bar_data = CategoryChartData()
    bar_data.categories = [
        "North America",
        "EMEA",
        "APAC",
        "LATAM",
    ]
    bar_data.add_series("Prior Customers", (710.0, 820.0, 940.0, 720.0))
    bar_data.add_series("Current Customers", (745.0, 890.0, 988.0, 755.0))
    bar_shape = s3.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.68),
        Inches(1.88),
        Inches(4.48),
        Inches(4.85),
        bar_data,
    )
    bar_shape.name = "chart:regional_customers"
    _style_modern_chart(bar_shape.chart, [ACCENT_INDIGO, ACCENT_CYAN])

    # 2. Tagged Dynamic Table (table:top_brands) with {{.col | fmt}} prototype row
    _add_bento_card(
        s3,
        5.55,
        1.18,
        7.28,
        2.78,
        title="Top Brand Revenue Drivers",
        subtitle="Dynamic N-Row Table (table:top_brands) with prototype row format masks",
        accent=ACCENT_CYAN,
    )
    dyn_tagged = s3.shapes.add_table(2, 5, Inches(5.77), Inches(1.82), Inches(6.84), Inches(0.64))
    dyn_tagged.name = "table:top_brands"
    dt = dyn_tagged.table
    _strip_default_table_style(dt)
    dt.rows[0].height = Inches(0.32)
    dt.rows[1].height = Inches(0.32)
    for c_idx, w in enumerate(
        [Inches(1.75), Inches(1.65), Inches(0.98), Inches(1.28), Inches(1.18)]
    ):
        dt.columns[c_idx].width = w
    for c_idx, h_text in enumerate(
        ["BRAND ({{_plan.title}})", "CATEGORY", "UNITS SOLD", "NET REVENUE", "YOY TREND"]
    ):
        _style_cell(
            dt.cell(0, c_idx),
            h_text,
            bg=CARD_BORDER,
            fg=ACCENT_CYAN,
            size=8.5,
            bold=True,
            align=PP_ALIGN.LEFT if c_idx < 2 else PP_ALIGN.RIGHT,
            border_hex="38BDF8",
        )
    for c_idx, tag_text in enumerate(
        [
            "{{.Brand Name}}",
            "{{.Category}}",
            "{{.Units Sold | #,##0}}",
            "{{.Net Revenue | $#,##0}}",
            "+{{.YoY Trend | 0.0%}}",
        ]
    ):
        _style_cell(
            dt.cell(1, c_idx),
            tag_text,
            bg=ROW_ALT_B,
            fg=(TEXT_WHITE if c_idx in (0, 3) else (ACCENT_AMBER if c_idx == 4 else TEXT_LIGHT)),
            size=9.5,
            bold=c_idx in (0, 3, 4),
            align=PP_ALIGN.LEFT if c_idx < 2 else PP_ALIGN.RIGHT,
        )

    # 3. Positional Dynamic Table (table:top_accounts) without {{}} tags
    _add_bento_card(
        s3,
        5.55,
        4.15,
        7.28,
        2.78,
        title="Enterprise Account Cohort Summary",
        subtitle="Positional Dynamic N-Row Table (table:top_accounts)",
        accent=ACCENT_ROSE,
    )
    dyn_pos = s3.shapes.add_table(2, 4, Inches(5.77), Inches(4.79), Inches(6.84), Inches(0.68))
    dyn_pos.name = "table:top_accounts"
    dp = dyn_pos.table
    _strip_default_table_style(dp)
    dp.rows[0].height = Inches(0.34)
    dp.rows[1].height = Inches(0.34)
    for c_idx, w in enumerate([Inches(2.15), Inches(2.25), Inches(1.04), Inches(1.40)]):
        dp.columns[c_idx].width = w
    for c_idx, h_text in enumerate(["ACCOUNT TIER", "PRIMARY SEGMENT", "ORDERS", "TOTAL REVENUE"]):
        _style_cell(
            dp.cell(0, c_idx),
            h_text,
            bg=CARD_BORDER,
            fg=ACCENT_ROSE,
            size=8.5,
            bold=True,
            align=PP_ALIGN.LEFT if c_idx < 2 else PP_ALIGN.RIGHT,
            border_hex="FB7185",
        )
    for c_idx, proto_text in enumerate(["Tier", "Segment", "0", "$0"]):
        _style_cell(
            dp.cell(1, c_idx),
            proto_text,
            bg=ROW_ALT_B,
            fg=ACCENT_ROSE if c_idx == 3 else (TEXT_WHITE if c_idx == 0 else TEXT_LIGHT),
            size=9.5,
            bold=c_idx in (0, 3),
            align=PP_ALIGN.LEFT if c_idx < 2 else PP_ALIGN.RIGHT,
        )

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def build_showcase_csv_zip_bytes() -> bytes:
    """Build the Looker csv_zip export matching all 9 tiles in the 3-slide showcase deck."""
    return build_csv_zip_bytes(
        {
            "kpi_summary.csv": (
                "Total Spend,Trend Rate,Avg Order Value\n34855339.456,0.0842,438.92\n"
            ),
            "mixed_metrics.csv": (
                "Metric,Prior,Current,Delta\n"
                "Customers,3190,3378,0.0589\n"
                "Orders per Customer,65.1,62.44,-0.0409\n"
                "Gross Margin,0.5312,0.5568,0.0256\n"
                "Avg Order Value,404.80,438.92,0.0842\n"
            ),
            "spend_by_category.csv": (
                "Product Category,Total Spend\n"
                "Apparel,11250000.0\n"
                "Footwear,9840000.0\n"
                "Accessories,6415339.46\n"
                "Outerwear,7350000.0\n"
            ),
            "monthly_trend.csv": (
                "Quarter,Direct,Marketplace\n"
                "2026-Q1,6450000.5,1620000.25\n"
                "2026-Q2,6810000.0,1740500.75\n"
                "2026-Q3,6590000.0,\n"
                "2026-Q4,7655339.0,1989499.0\n"
            ),
            "aov_trend.csv": (
                "Quarter,Prior Year AOV,Current Year AOV,Target Benchmark\n"
                "2026-Q1,392.10,418.40,405.00\n"
                "2026-Q2,401.50,432.15,410.00\n"
                "2026-Q3,408.20,439.80,414.50\n"
                "2026-Q4,417.40,465.30,420.00\n"
            ),
            "cost_bridge.csv": (
                "Bridge Step,Base,Impact\n"
                "Prior AOV,0.0,404.80\n"
                "Volume Mix,404.80,14.20\n"
                "Unit Price,419.00,11.60\n"
                "Premium Mix,430.60,8.32\n"
                "Current AOV,0.0,438.92\n"
            ),
            "regional_customers.csv": (
                "Region,Prior Customers,Current Customers\n"
                "North America,710,745\n"
                "EMEA,820,890\n"
                "APAC,940,988\n"
                "LATAM,720,755\n"
            ),
            "top_brands.csv": (
                "Brand Name,Category,Units Sold,Net Revenue,YoY Trend\n"
                "Apex Goods,Outerwear,120,450200.4,0.124\n"
                "Nova Apparel,Activewear,340,289100.2,0.382\n"
                "Vertex Active,Footwear,890,195000.0,0.065\n"
                "Solstice Co,Accessories,64,182400.0,0.091\n"
            ),
            "top_accounts.csv": (
                "Account Tier,Primary Segment,Orders,Total Revenue\n"
                'Account #1 (>$500k),Enterprise Retail,142,"$684,210"\n'
                'Account #2 ($250k-$500k),Direct Wholesale,88,"$412,900"\n'
                'Account #3 ($250k-$500k),Global Partner,215,"$298,450"\n'
            ),
        }
    )


def build_embed_demo_template_bytes() -> bytes:
    """Build a 3-slide deck bound to ~/lkrdev/looker-embed-demo (embed_demo::brand_overview)."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 1: Sales Pulse (Tab 1 of embed_demo::brand_overview)
    # =========================================================================
    s1 = prs.slides.add_slide(blank)
    _add_slide_chrome(s1, slide_num=1, eyebrow="SALES PULSE  •  LOOKER EMBED DEMO")

    tb1 = s1.shapes.add_textbox(Inches(0.5), Inches(0.48), Inches(12.33), Inches(0.62))
    tb1.name = "Sales Pulse Header"
    p1 = tb1.text_frame.paragraphs[0]
    r1 = p1.add_run()
    r1.text = "{{_plan.title}}   |   Revenue: {{revenue.Current Period | $#,##0.00}}"
    r1.font.name = FONT_SANS
    r1.font.size = Pt(18)
    r1.font.bold = True
    r1.font.color.rgb = TEXT_WHITE
    r2 = p1.add_run()
    r2.text = (
        "   •   P/P Change: +{{revenue[0].P/P Change | 0.0%}}"
        "   •   Current AOV: {{average_order_value.Current AOV | $#,##0.00}}"
    )
    r2.font.name = FONT_SANS
    r2.font.size = Pt(13)
    r2.font.bold = True
    r2.font.color.rgb = ACCENT_CYAN

    s1_kpis = [
        (
            0.50,
            ACCENT_CYAN,
            "REVENUE",
            "{{revenue.Current Period | $#,##0}}",
            "P/P Change: +{{revenue[0].P/P Change | 0.0%}} vs. Prior",
        ),
        (
            3.63,
            ACCENT_EMERALD,
            "ORDERS",
            "{{orders.Current Period | #,##0}}",
            "P/P Change: +{{orders[0].P/P Change | 0.0%}} vs. Prior",
        ),
        (
            6.77,
            ACCENT_INDIGO,
            "AVERAGE ORDER VALUE",
            "{{average_order_value.Current AOV | $#,##0.00}}",
            "P/P Change: +{{average_order_value[0].P/P Change | 0.0%}} vs. Prior",
        ),
        (
            9.91,
            ACCENT_AMBER,
            "UNITS SOLD",
            "{{units_sold.Current Period | #,##0}}",
            "P/P Change: +{{units_sold[0].P/P Change | 0.0%}} vs. Prior",
        ),
    ]
    _add_kpi_strip(s1, s1_kpis)

    _add_bento_card(
        s1,
        0.50,
        2.66,
        7.55,
        4.30,
        title="Top Categories by Revenue & Volume",
        subtitle="embed_demo::order_items  •  Total Sale Price & Purchased Items by Category",
        accent=ACCENT_CYAN,
    )
    cat_data = CategoryChartData()
    cat_data.categories = [
        "Jeans",
        "Outerwear & Coats",
        "Sweaters",
        "Tops & Tees",
        "Accessories",
    ]
    cat_data.add_series("Total Sale Price", (482500.0, 391200.0, 284900.0, 219400.0, 164800.0))
    cat_data.add_series("Purchased Items", (5420.0, 3180.0, 3890.0, 4610.0, 2940.0))
    cat_shape = s1.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.68),
        Inches(3.34),
        Inches(7.18),
        Inches(3.48),
        cat_data,
    )
    cat_shape.name = "chart:top_categories_by_revenue_&_volume"
    _style_modern_chart(cat_shape.chart, [ACCENT_CYAN, ACCENT_INDIGO])

    _add_bento_card(
        s1,
        8.25,
        2.66,
        4.58,
        4.30,
        title="Revenue Share by Country",
        subtitle="embed_demo::order_items  •  users.country",
        accent=ACCENT_INDIGO,
    )
    country_data = CategoryChartData()
    country_data.categories = ["United States", "United Kingdom", "China", "Brasil"]
    country_data.add_series("Total Sale Price", (742900.0, 318400.0, 268100.0, 213400.0))
    country_shape = s1.shapes.add_chart(
        XL_CHART_TYPE.DOUGHNUT,
        Inches(8.42),
        Inches(3.32),
        Inches(4.24),
        Inches(3.48),
        country_data,
    )
    country_shape.name = "chart:revenue_share_by_country"
    _style_modern_chart(
        country_shape.chart,
        [ACCENT_CYAN, ACCENT_INDIGO, ACCENT_EMERALD, ACCENT_AMBER],
        is_doughnut=True,
    )

    # =========================================================================
    # SLIDE 2: Customers (Tab 2 of embed_demo::brand_overview)
    # =========================================================================
    s2 = prs.slides.add_slide(blank)
    _add_slide_chrome(
        s2, slide_num=2, eyebrow="CUSTOMER ANALYTICS  •  ACQUISITION & GEOGRAPHIC FOOTPRINT"
    )

    s2_title = s2.shapes.add_textbox(Inches(0.5), Inches(0.48), Inches(12.33), Inches(0.62))
    s2_title.name = "Customers Header"
    p2 = s2_title.text_frame.paragraphs[0]
    r2_title = p2.add_run()
    r2_title.text = (
        "{{_plan.title}}   —   Active Customers: "
        "{{total_active_customers.Current Period | #,##0}}   •   "
        "Avg Spend: {{customer_average_spend.Current Spend | $#,##0.00}}"
    )
    r2_title.font.name = FONT_SANS
    r2_title.font.size = Pt(18)
    r2_title.font.bold = True
    r2_title.font.color.rgb = TEXT_WHITE

    s2_kpis = [
        (
            0.50,
            ACCENT_CYAN,
            "TOTAL ACTIVE CUSTOMERS",
            "{{total_active_customers.Current Period | #,##0}}",
            "P/P Change: +{{total_active_customers[0].P/P Change | 0.0%}}",
        ),
        (
            3.63,
            ACCENT_EMERALD,
            "CUSTOMER AVERAGE SPEND",
            "{{customer_average_spend.Current Spend | $#,##0.00}}",
            "P/P Change: +{{customer_average_spend[0].P/P Change | 0.0%}}",
        ),
        (
            6.77,
            ACCENT_INDIGO,
            "ORDERS PER CUSTOMER",
            "{{orders_per_customer.Current Period | 0.00}}",
            "P/P Change: +{{orders_per_customer[0].P/P Change | 0.0%}}",
        ),
        (
            9.91,
            ACCENT_AMBER,
            "NEW CUSTOMER ACQUISITIONS",
            "{{new_customer_acquisitions.Current Period | #,##0}}",
            "P/P Change: +{{new_customer_acquisitions[0].P/P Change | 0.0%}}",
        ),
    ]
    _add_kpi_strip(s2, s2_kpis)

    _add_bento_card(
        s2,
        0.50,
        2.66,
        4.15,
        4.30,
        title="Monthly Revenue & AOV Trajectory",
        subtitle="order_items.created_month  •  Sale Price & AOV",
        accent=ACCENT_CYAN,
    )
    traj_data = CategoryChartData()
    traj_data.categories = ["2026-07", "2026-08", "2026-09", "2026-10"]
    traj_data.add_series("Total Sale Price", (342000.0, 368500.0, 394200.0, 438100.0))
    traj_data.add_series("Average Sale Price", (86.4, 88.9, 91.2, 94.8))
    traj_shape = s2.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS,
        Inches(0.65),
        Inches(3.34),
        Inches(3.85),
        Inches(3.48),
        traj_data,
    )
    traj_shape.name = "chart:monthly_revenue_&_aov_trajectory"
    _style_modern_chart(traj_shape.chart, [ACCENT_CYAN, ACCENT_EMERALD])

    _add_bento_card(
        s2,
        4.82,
        2.66,
        3.85,
        4.30,
        title="Customer Traffic Source Composition",
        subtitle="users.traffic_source  •  Active Users",
        accent=ACCENT_EMERALD,
    )
    traffic_data = CategoryChartData()
    traffic_data.categories = ["Search", "Organic", "Email", "Social", "Display"]
    traffic_data.add_series("Users Count", (4820.0, 2910.0, 1840.0, 1420.0, 890.0))
    traffic_shape = s2.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(4.97),
        Inches(3.34),
        Inches(3.55),
        Inches(3.48),
        traffic_data,
    )
    traffic_shape.name = "chart:customer_traffic_source_composition"
    _style_modern_chart(traffic_shape.chart, [ACCENT_EMERALD])

    _add_bento_card(
        s2,
        8.84,
        2.66,
        3.99,
        4.30,
        title="Top Geographic Footprint",
        subtitle="users.state  •  Top Customer States",
        accent=ACCENT_AMBER,
    )
    geo_data = CategoryChartData()
    geo_data.categories = ["California", "Texas", "New York", "England", "Florida"]
    geo_data.add_series("Users Count", (2410.0, 1680.0, 1490.0, 1320.0, 1140.0))
    geo_shape = s2.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(8.99),
        Inches(3.34),
        Inches(3.69),
        Inches(3.48),
        geo_data,
    )
    geo_shape.name = "chart:top_geographic_footprint_by_location"
    _style_modern_chart(geo_shape.chart, [ACCENT_AMBER])

    # =========================================================================
    # SLIDE 3: Orders & Logistics (Tab 3 of embed_demo::brand_overview)
    # =========================================================================
    s3 = prs.slides.add_slide(blank)
    _add_slide_chrome(
        s3, slide_num=3, eyebrow="ORDERS & LOGISTICS  •  FULFILLMENT PIPELINE & CHURN HOTSPOTS"
    )

    s3_title = s3.shapes.add_textbox(Inches(0.5), Inches(0.48), Inches(12.33), Inches(0.62))
    s3_title.name = "Orders & Logistics Header"
    p3 = s3_title.text_frame.paragraphs[0]
    r3_title = p3.add_run()
    r3_title.text = (
        "{{_plan.title}}   —   Shipped: {{shipped_orders.Current Period | #,##0}}   •   "
        "Processing: {{processing_orders.Current Period | #,##0}}   •   "
        "Returned: {{returned_orders.Current Period | #,##0}}"
    )
    r3_title.font.name = FONT_SANS
    r3_title.font.size = Pt(18)
    r3_title.font.bold = True
    r3_title.font.color.rgb = TEXT_WHITE

    s3_kpis = [
        (
            0.50,
            ACCENT_EMERALD,
            "SHIPPED ORDERS",
            "{{shipped_orders.Current Period | #,##0}}",
            "P/P Change: +{{shipped_orders[0].P/P Change | 0.0%}}",
        ),
        (
            3.63,
            ACCENT_CYAN,
            "PROCESSING ORDERS",
            "{{processing_orders.Current Period | #,##0}}",
            "P/P Change: {{processing_orders[0].P/P Change | 0.0%}}",
        ),
        (
            6.77,
            ACCENT_AMBER,
            "RETURNED ORDERS",
            "{{returned_orders.Current Period | #,##0}}",
            "P/P Change: {{returned_orders[0].P/P Change | 0.0%}}",
        ),
        (
            9.91,
            ACCENT_ROSE,
            "CANCELLED ORDERS",
            "{{cancelled_orders.Current Period | #,##0}}",
            "P/P Change: {{cancelled_orders[0].P/P Change | 0.0%}}",
        ),
    ]
    _add_kpi_strip(s3, s3_kpis)

    _add_bento_card(
        s3,
        0.50,
        2.66,
        5.35,
        4.30,
        title="Fulfillment Pipeline Status Breakdown",
        subtitle="order_items.status  •  Order Items & Sale Price",
        accent=ACCENT_INDIGO,
    )
    pipe_data = CategoryChartData()
    pipe_data.categories = ["Complete", "Shipped", "Processing", "Returned", "Cancelled"]
    pipe_data.add_series("Order Items Count", (8420.0, 5890.0, 2140.0, 1180.0, 790.0))
    pipe_shape = s3.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.68),
        Inches(3.34),
        Inches(4.98),
        Inches(3.48),
        pipe_data,
    )
    pipe_shape.name = "chart:fulfillment_pipeline_status_breakdown"
    _style_modern_chart(pipe_shape.chart, [ACCENT_INDIGO])

    _add_bento_card(
        s3,
        6.05,
        2.66,
        6.78,
        4.30,
        title="Return & Cancellation Hotspots by Category",
        subtitle="Dynamic N-Row Table (table:return_&_cancellation_hotspots_by_category)",
        accent=ACCENT_ROSE,
    )
    churn_tbl_shape = s3.shapes.add_table(
        2, 4, Inches(6.27), Inches(3.38), Inches(6.34), Inches(0.68)
    )
    churn_tbl_shape.name = "table:return_&_cancellation_hotspots_by_category"
    ct = churn_tbl_shape.table
    _strip_default_table_style(ct)
    ct.rows[0].height = Inches(0.34)
    ct.rows[1].height = Inches(0.34)
    for c_idx, w in enumerate([Inches(2.34), Inches(1.30), Inches(1.30), Inches(1.40)]):
        ct.columns[c_idx].width = w
    for c_idx, h_text in enumerate(["CATEGORY", "RETURNED", "CANCELLED", "CHURN RATIO"]):
        _style_cell(
            ct.cell(0, c_idx),
            h_text,
            bg=CARD_BORDER,
            fg=ACCENT_CYAN,
            size=8.5,
            bold=True,
            align=PP_ALIGN.LEFT if c_idx == 0 else PP_ALIGN.RIGHT,
            border_hex="38BDF8",
        )
    for c_idx, tag_text in enumerate(
        [
            "{{.Category}}",
            "{{.Returned Count | #,##0}}",
            "{{.Cancelled Count | #,##0}}",
            "{{.Transaction Churn Ratio | 0.0%}}",
        ]
    ):
        _style_cell(
            ct.cell(1, c_idx),
            tag_text,
            bg=ROW_ALT_B,
            fg=ACCENT_ROSE if c_idx == 3 else (TEXT_WHITE if c_idx == 0 else TEXT_LIGHT),
            size=9.5,
            bold=c_idx in (0, 3),
            align=PP_ALIGN.LEFT if c_idx == 0 else PP_ALIGN.RIGHT,
        )

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def build_embed_demo_csv_zip_bytes() -> bytes:
    """Build the Looker csv_zip export matching embed_demo::brand_overview tiles."""
    return build_csv_zip_bytes(
        {
            "revenue.csv": "Current Period,P/P Change\n1542800.50,0.142\n",
            "orders.csv": "Current Period,P/P Change\n16920,0.098\n",
            "average_order_value.csv": "Current AOV,P/P Change\n91.18,0.041\n",
            "units_sold.csv": "Current Period,P/P Change\n20040,0.116\n",
            "top_categories_by_revenue_&_volume.csv": (
                "Category,Total Sale Price,Purchased Items\n"
                "Jeans,482500.0,5420\n"
                "Outerwear & Coats,391200.0,3180\n"
                "Sweaters,284900.0,3890\n"
                "Tops & Tees,219400.0,4610\n"
                "Accessories,164800.0,2940\n"
            ),
            "revenue_share_by_country.csv": (
                "Country,Total Sale Price\n"
                "United States,742900.0\n"
                "United Kingdom,318400.0\n"
                "China,268100.0\n"
                "Brasil,213400.0\n"
            ),
            "total_active_customers.csv": "Current Period,P/P Change\n11880,0.128\n",
            "customer_average_spend.csv": "Current Spend,P/P Change\n129.87,0.064\n",
            "orders_per_customer.csv": "Current Period,P/P Change\n1.42,0.032\n",
            "new_customer_acquisitions.csv": "Current Period,P/P Change\n3490,0.185\n",
            "monthly_revenue_&_aov_trajectory.csv": (
                "Created Month,Total Sale Price,Average Sale Price\n"
                "2026-07,342000.0,86.4\n"
                "2026-08,368500.0,88.9\n"
                "2026-09,394200.0,91.2\n"
                "2026-10,438100.0,94.8\n"
            ),
            "customer_traffic_source_composition.csv": (
                "Traffic Source,Users Count\n"
                "Search,4820\n"
                "Organic,2910\n"
                "Email,1840\n"
                "Social,1420\n"
                "Display,890\n"
            ),
            "top_geographic_footprint_by_location.csv": (
                "State,Users Count\n"
                "California,2410\n"
                "Texas,1680\n"
                "New York,1490\n"
                "England,1320\n"
                "Florida,1140\n"
            ),
            "shipped_orders.csv": "Current Period,P/P Change\n14310,0.114\n",
            "processing_orders.csv": "Current Period,P/P Change\n2140,-0.024\n",
            "returned_orders.csv": "Current Period,P/P Change\n1180,-0.038\n",
            "cancelled_orders.csv": "Current Period,P/P Change\n790,-0.051\n",
            "fulfillment_pipeline_status_breakdown.csv": (
                "Status,Order Items Count\n"
                "Complete,8420\n"
                "Shipped,5890\n"
                "Processing,2140\n"
                "Returned,1180\n"
                "Cancelled,790\n"
            ),
            "return_&_cancellation_hotspots_by_category.csv": (
                "Category,Returned Count,Cancelled Count,Transaction Churn Ratio\n"
                "Jeans,380,210,0.109\n"
                "Outerwear & Coats,290,175,0.146\n"
                "Sweaters,220,160,0.098\n"
                "Tops & Tees,175,140,0.068\n"
                "Accessories,115,105,0.075\n"
            ),
        }
    )
