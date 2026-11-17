---
name: looker-action-pptx-templates
description: >
  Master skill for creating or retrofitting native editable Microsoft PowerPoint (.pptx)
  templates for Looker action-pptx using existing LookML. Supports Mode 1 (start from
  scratch with master slides and a general idea) and Mode 2 (retrofit a built-out static
  .pptx deck with template bindings in place). Emits the templated .pptx and matching
  dashboard tile contract; defaults to NOT building LookML dashboards unless explicitly requested.
---

# Looker Action PowerPoint Template Builder (`looker-action-pptx-templates`)

Build or retrofit native, editable `.pptx` templates with `action-pptx` Selection Pane bindings (`chart:<tile_slug>`, `table:<tile_slug>`) and text tags (`{{tile_slug[row].Column | fmt}}`) mapped to **existing LookML** so every placeholder aligns 1:1 with the Looker dashboard `csv_zip` export you will schedule.

## 0. Scope & Default Guardrails (Read First)

1. **LookML Already Exists — Never Author LookML Models/Views**:
   - Inspect local `.lkml` files (`views/`, `explores/`, `models/`) or live Looker Explore metadata (`lookml_model_explore` via `lkr` / Looker MCP) to ground every column label in real LookML fields.
   - **Never** invent LookML fields or edit `.view.lkml` / `.model.lkml`. If a slide metric has no backing LookML field, flag the gap and point the user to companion LookML skills (`npx skills add lkrdev/lookml_skills`, `npx skills add looker-open-source/looker-skills`, `npx skills add lkrdev/cli`).
2. **Default: Do NOT Build LookML Dashboards (`build_dashboard = false`)**:
   - Unless the user **explicitly** asks you to create, edit, or sync a `.dashboard.lookml` or User-Defined Dashboard (UDD), **stop after producing**:
     1. The templated `.pptx` file
     2. `dashboard_tiles.json` (validated via `pptx-merge validate`)
     3. A **Target LookML Dashboard Tile Spec** Markdown table showing the exact tiles the user needs to make.
   - Only execute **Section 4 (Opt-In LookML Dashboard Authoring)** when explicitly requested.

---

## 1. Mode Selection & End-to-End Workflows

### Mode 1 Workflow — Start from Scratch (Master Slides + Idea + Existing LookML)

1. **Inspect Existing LookML & Slide Masters**:
   - Read target LookML Explore(s) and record `(view.field_name, short_label, type)` where `short_label` is LookML `label:` (or Title Case `field_name` when `label:` is omitted).
   - Inspect the master `.pptx` layouts, dimensions, and placeholders via `uv`:
     ```bash
     uv run --with python-pptx python -c '
     from pptx import Presentation
     prs = Presentation("master.pptx")
     print(f"{prs.slide_width.inches:.3f}x{prs.slide_height.inches:.3f} in")
     for i, layout in enumerate(prs.slide_layouts):
         phs = [(p.placeholder_format.idx, p.name) for p in layout.placeholders]
         print(f"[{i}] {layout.name}: {phs}")
     '
     ```
2. **Design Consolidated Tile Contract (`dashboard_tiles.json`)**:
   - Map the user's narrative across slides (e.g., Slide 1 Executive KPI Strip + Trend Chart, Slide 2 Category/Regional Breakdowns, Slide 3 Top-N Leaderboard Table).
   - **Consolidate scalar KPIs**: Group single-row headline metrics from the same Explore into **one** `KPI Summary` (`kpi_summary`) tile (`limit: 1`) rather than creating separate 1-value tiles.
3. **Build Slides from Master Layouts**:
   - Open `Presentation("master.pptx")` (or `Presentation()` with `13.333"` x `7.5"` 16:9 dimensions if no master deck was supplied), remove any sample starter slides if needed, add slides from `prs.slide_layouts[idx]`, and populate the **Templating Sections (Section 2)**.
4. **Validate & Hand Off**:
   - Write `dashboard_tiles.json`, run the **Validation Gate (Section 3)**, and output the **Target LookML Dashboard Tile Spec** table (do not build `.dashboard.lookml` unless explicitly asked).

### Mode 2 Workflow — Retrofit Existing Built-Out Slides (Static `.pptx` + Existing LookML)

1. **Inventory Every Shape & Check for `chartEx`**:
   - Walk all slides (including nested `<p:grpSp>` `GroupShape`s) and check for unsupported Office 2016+ `chartEx` parts before editing:
     ```bash
     uv run --with python-pptx python -c '
     import zipfile
     from pptx import Presentation
     from pptx.shapes.group import GroupShape

     path = "existing_deck.pptx"
     with zipfile.ZipFile(path) as zf:
         chart_ex = [n for n in zf.namelist() if n.startswith("ppt/charts/chartEx")]
         if chart_ex:
             print("UNSUPPORTED chartEx:", chart_ex)

     def walk(shapes, prefix=""):
         for idx, s in enumerate(shapes):
             loc = f"{prefix}{idx}"
             if isinstance(s, GroupShape):
                 yield from walk(s.shapes, f"{loc}.")
             else:
                 yield loc, s

     prs = Presentation(path)
     for s_idx, slide in enumerate(prs.slides, 1):
         for loc, sh in walk(slide.shapes):
             if sh.has_chart:
                 print(f"Slide {s_idx} [{loc}] {sh.name!r} chart type={sh.chart.chart_type} series={[s.name for s in sh.chart.series]}")
             elif sh.has_table:
                 t = sh.table
                 print(f"Slide {s_idx} [{loc}] {sh.name!r} table {len(t.rows)}x{len(t.columns)} hdr={[c.text.strip() for c in t.rows[0].cells]}")
             elif sh.has_text_frame and sh.text_frame.text.strip():
                 print(f"Slide {s_idx} [{loc}] {sh.name!r} text {sh.text_frame.text.strip()[:80]!r}")
     '
     ```
2. **Map Static Values, Tables & Charts to Existing LookML**:
   - Match each hardcoded callout, table column, and chart series to existing LookML dimensions/measures. Flag any unmapped metric to the user instead of inventing LookML fields.
3. **Mutate Shapes Surgically In Place (Save to `<stem>_templated.pptx`)**:
   - **Never** recreate slides from scratch in Mode 2; mutate the existing `Presentation` object so slide masters, logos, coordinates, gradients, borders, and `<a:rPr>` run formatting remain intact:
     - **Text callouts & titles**: Replace literal numbers/dates inside the target paragraph via `set_para_text_keep_rpr` (Section 2.2).
     - **Charts**: Rename `shape.name = "chart:<tile_slug>"`. Ensure chart series count `K` equals `len(tile.columns) - 1` (call `chart.replace_data(...)` if series count or names differ). If a chart is `chartEx`, replace only that shape at identical `(left, top, width, height)` with a standard Office chart (Section 2.5).
     - **Tables**:
       - *Repeating Top-N / entity rows*: Rename `shape.name = "table:<tile_slug>"`, keep header row(s) `0..H-1` and **first data row `H`**, delete extra static rows `H+1..end` (`pptx-merge` treats `rows[:-1]` as headers and clones only `rows[-1]`), and write `{{.Column | fmt}}` into row `H` via `retrofit_dynamic_table` (Section 2.4).
       - *Fixed mixed-unit metric rows*: Keep original shape name (not `table:...`), keep all rows, and write row-indexed `{{tile_slug[i].Column | fmt}}` into each value cell (Section 2.3).
4. **Validate & Hand Off**:
   - Save as `<stem>_templated.pptx` (do not overwrite input unless requested), write `dashboard_tiles.json`, run the **Validation Gate (Section 3)**, and present the **Target LookML Dashboard Tile Spec** table.

---

## 2. Templating Sections & `python-pptx` Mechanics

### 2.1 Tile Slug & `csv_zip` Column Header Contract

- **Tile Slug Normalization**: Looker converts each dashboard tile `title` into `<tile_slug>.csv` via `re.sub(r'(^\.+$)|[/?*\\:|"<>\s%]', '_', title.strip()).lower()`:
  - `"KPI Summary"` -> `kpi_summary` (`kpi_summary.csv`)
  - `"Spend by Category"` -> `spend_by_category` (`spend_by_category.csv`)
  - Every tile on the dashboard **must** produce a unique `tile_slug` (duplicates collide as `<slug>_1.csv` and fail validation).
- **Column Header Precedence in `<tile_slug>.csv`**:
  1. `vis_config.series_labels[field_name]`
  2. `dynamic_fields[].label` (table calculations / custom fields)
  3. LookML field `label:` (short label, **no view prefix**) when `show_view_names: false` — **always design tags for `show_view_names: false`** (e.g. `Total Spend`, not `Inventory Items Total Spend`).
- **Hidden Sort Columns**: Any LookML dimension in `vis_config.hidden_fields` orders SQL rows without appearing in `<tile_slug>.csv` — use this whenever row-indexed tags (`[0]`, `[1]`) or bridge steps require deterministic row order.

### 2.2 Schedule Metadata & Inline Text Tags (`{{tile_slug[row].Column | fmt}}`)

- **Tag Syntax**:
  - Row `0` shorthand: `{{kpi_summary.Total Spend | $#,##0}}`
  - Explicit zero-based row index: `{{mixed_metrics[2].Delta | 0.0%}}`
  - Built-in schedule metadata (`_plan`): `{{_plan.title}}`, `{{_plan.run_at | yyyy-mm-dd}}`, `{{_plan.url}}`, `{{_plan.scheduled_plan_id}}`, `{{_plan.webhook_id}}`
- **Supported Format Masks (`SUPPORTED_MASKS` — strictly limited to these 9)**:
  - Counts / Decimals: `#,##0`, `0.0`, `0.00`
  - Currency (USD): `$#,##0`, `$#,##0.00` (`-$1,234.50` on negatives)
  - Percentages (multiplies raw Looker `unformatted` ratio by `100`, e.g. `0.0842` -> `8.4%`): `0%`, `0.0%`, `0.00%`
  - ISO Date: `yyyy-mm-dd`
- **Preserving Typography When Injecting Tags (`<a:rPr>`)**:
  Assigning `paragraph.text = "..."` in `python-pptx` destroys run-level font family, size, bold, and RGB color. Always mutate `runs[0]` and prune trailing runs:
  ```python
  def set_para_text_keep_rpr(para, new_text: str) -> None:
      runs = list(para.runs)
      if runs:
          runs[0].text = new_text
          for extra in runs[1:]:
              para._p.remove(extra._r)
      else:
          para.text = new_text
  ```

### 2.3 KPI Cards & Fixed Mixed-Unit Tables

- **Scalar KPI Strip**: Bind 3–6 KPI callout boxes on a slide to a single 1-row `kpi_summary` tile (`{{kpi_summary.Total Spend | $#,##0}}`, `{{kpi_summary.Trend Rate | 0.0%}}`, `{{kpi_summary.Avg Order Value | $#,##0.00}}`).
- **Fixed Mixed-Unit Table** (when Row 0 is an integer count, Row 1 is a decimal, and Row 2 is a `%` in the **same** column):
  - Leave `shape.name` as a normal descriptive name (do **not** prefix with `table:`).
  - Place explicit row-indexed tags per cell: `{{mixed_metrics[0].Current | #,##0}}`, `{{mixed_metrics[1].Current | 0.0}}`, `{{mixed_metrics[2].Current | 0.0%}}`, backed by a deterministically sorted Looker tile (`sorts: [view.metric_sort asc]`, `hidden_fields: [view.metric_sort]`).

### 2.4 Dynamic N-Row Tables (`table:<tile_slug>`)

- **Selection Pane Binding**: Set `table_shape.name = "table:<tile_slug>"` (e.g., `table:top_brands`).
- **Prototype Row Rule**: `pptx-merge` keeps `table.rows[:-1]` as static header row(s), clones `table.rows[-1]` (the **very last row**) once per CSV row preserving cell fills/borders/fonts, and removes the prototype row:
  - Use `{{.Column Label | fmt}}` tags inside `table.rows[-1]` cells to map columns by name and format mask (recommended), or leave cells blank/untagged for raw positional `0..N-1` fill.
  - Keep expected query `limit` sized to the slide's vertical space (typically `5`–`12` rows on a `7.5"` slide; hard engine ceiling is `< 5000` rows).
- **Mode 2 Static Table Retrofit Pattern (Prune Extra Static Data Rows)**:
  ```python
  def retrofit_dynamic_table(
      table_shape, tile_slug: str, proto_tags: list[str], header_rows: int = 1
  ) -> None:
      table_shape.name = f"table:{tile_slug}"
      tbl_el = table_shape.table._tbl
      # Keep header_rows + 1 styled data row as the prototype row; remove remaining static rows
      for extra_tr in list(tbl_el.tr_lst)[header_rows + 1 :]:
          tbl_el.remove(extra_tr)
      proto_row = list(table_shape.table.rows)[header_rows]
      for cell, tag in zip(proto_row.cells, proto_tags, strict=True):
          set_para_text_keep_rpr(cell.text_frame.paragraphs[0], tag)
  ```

### 2.5 Native Editable Office Charts (`chart:<tile_slug>`)

- **Selection Pane Binding**: Set `chart_shape.name = "chart:<tile_slug>"` (e.g., `chart:monthly_trend`).
- **CSV Shape Contract**:
  - Column `0` = category labels (strings, e.g., `Quarter`, `Product Category`, `Region`).
  - Columns `1..K` = numeric series (`float`; empty CSV cells become `None` chart gaps; non-numeric strings fail merge). At least 2 total columns (`1` category + `>= 1` series) and **zero Looker pivots** unless pivoted headers are static and deterministic.
- **Series-Count Parity**: Seed `CategoryChartData` with the **exact `K` series** matching `columns[1:]` so `<c:ser>` XML caches and embedded `ppt/embeddings/*.xlsx` columns stay 1:1:
  ```python
  from pptx.chart.data import CategoryChartData
  from pptx.enum.chart import XL_CHART_TYPE

  chart_data = CategoryChartData()
  chart_data.categories = ["Q1", "Q2"]
  for series_col in ["Direct", "Marketplace"]:  # Must match CSV columns[1:]
      chart_data.add_series(series_col, (0.0, 0.0))

  # Mode 1: add new chart
  chart_shape = slide.shapes.add_chart(
      XL_CHART_TYPE.COLUMN_CLUSTERED, left, top, width, height, chart_data
  )
  chart_shape.name = "chart:monthly_trend"

  # Mode 2: update existing chart in place when series count/headers differ
  existing_chart_shape.name = "chart:monthly_trend"
  existing_chart_shape.chart.replace_data(chart_data)
  ```
- **Ban Office 2016+ Extended Charts (`chartEx`)**:
  - PowerPoint's native Waterfall, Funnel, Treemap, Sunburst, and Box & Whisker charts serialize as `ppt/charts/chartEx*.xml`, which `pptx-merge` rejects (`unsupported_chart_ex`).
  - Use standard `<c:chartSpace>` types (`COLUMN_CLUSTERED`, `COLUMN_STACKED`, `BAR_CLUSTERED`, `BAR_STACKED`, `LINE`, `AREA`, `PIE`, `DOUGHNUT`).
  - **Waterfall / Cost Bridge Pattern**: Build bridges as `XL_CHART_TYPE.COLUMN_STACKED` with 2 numeric series (`Base`, `Impact`) and hide the `Base` series fill/line (`chart.series[0].format.fill.background()`, `chart.series[0].format.line.fill.background()`).

---

## 3. Validation Gate & Default Deliverable (`dashboard_tiles.json`)

Every Mode 1 or Mode 2 run **must** write `dashboard_tiles.json` and pass `pptx-merge` validation before finishing:

1. **Emit `dashboard_tiles.json`**:
   ```json
   {
     "dashboard_tiles": [
       {"title": "KPI Summary", "slug": "kpi_summary", "columns": ["Total Spend", "Trend Rate", "Avg Order Value"]},
       {"title": "Monthly Trend", "slug": "monthly_trend", "columns": ["Quarter", "Direct", "Marketplace"]},
       {"title": "Top Brands", "slug": "top_brands", "columns": ["Brand Name", "Units Sold", "Net Revenue"]}
     ]
   }
   ```
2. **Run `inspect` + `validate` (`uv run` inside `action-pptx`, or `uvx` from any workspace)**:
   ```bash
   uv run pptx-merge inspect template.pptx
   uv run pptx-merge validate template.pptx --schema dashboard_tiles.json
   # External workspace: uvx --from git+https://github.com/lkrdev/action-pptx pptx-merge validate template.pptx --schema dashboard_tiles.json
   ```
   Fix any reported `issues` (`missing_tile`, `missing_column`, `invalid_format_mask`, `insufficient_chart_columns`, `duplicate_dashboard_tile`, `unsupported_chart_ex`, `missing_tile_in_tag`) until `"valid": true`.
3. **Present the Target LookML Dashboard Tile Spec Table (Default Stopping Point)**:
   Summarize the required dashboard tiles so the user (or a downstream dashboard skill) can build the matching LookML dashboard:

   | Tile Title (`tile_slug`) | `model::explore` | Visible Fields -> CSV Columns (`show_view_names: false`) | `hidden_fields` & `sorts` | `limit` | Suggested Vis |
   | :--- | :--- | :--- | :--- | :--- | :--- |
   | `KPI Summary` (`kpi_summary`) | `ecom::orders` | `orders.total_spend` (`Total Spend`), `orders.trend_rate` (`Trend Rate`) | `sorts: [orders.total_spend desc]` | `1` | `looker_grid` |
   | `Monthly Trend` (`monthly_trend`) | `ecom::orders` | `orders.quarter` (`Quarter`), `orders.direct_spend` (`Direct`), `orders.marketplace_spend` (`Marketplace`) | `sorts: [orders.quarter asc]` | `12` | `looker_column` |
   | `Top Brands` (`top_brands`) | `ecom::orders` | `products.brand` (`Brand Name`), `orders.units` (`Units Sold`), `orders.net_revenue` (`Net Revenue`) | `sorts: [orders.net_revenue desc]` | `10` | `looker_grid` |

---

## 4. Explicit Opt-In Only: Authoring the Matching LookML Dashboard

> **STOP UNLESS EXPLICITLY REQUESTED:** Skip this entire section unless the user explicitly asks to generate, update, or sync the `.dashboard.lookml` / Looker dashboard.

When the user explicitly opts into building the LookML dashboard alongside the `.pptx` template:

1. **Pre-Flight Explore & 1-Row SQL Verification**:
   - Confirm every `model`, `explore`, and `view.field` exists via `lookml_model_explore` (or local `.lkml` inspection + `lkr parse lookml-to-api --validate`).
   - Run a 1-row `run_inline_query` per Explore (propagating concrete partition `default_value` filters) when connected to a live Looker instance; exclude null chart categories with `filters: {"view.category_dim": "-NULL"}` when needed.
2. **24-Column `newspaper` Dashboard Structure**:
   - Author `dashboards/<dashboard_name>.dashboard.lookml` with `layout: newspaper`, `preferred_viewer: dashboards-next`, `col + width <= 24`, and ensure `include: "/dashboards/*.dashboard.lookml"` (or `**/*.dashboard.lookml`) is present in `<model>.model.lkml`.
   - Use clean 24-col rows: Row `0` (`height: 4`) for KPI summary tiles; Rows `4+` (`height: 6`–`8`, `width: 8` or `12`) for native chart/table backing tiles; use native `tabs:` if `> 25` tiles.
3. **Mandatory `action-pptx` Tile Properties**:
   - `title:` **must** sanitize to the exact `tile_slug` in `dashboard_tiles.json` with zero duplicates across the dashboard.
   - Set `show_view_names: false` on every tile; use `series_labels:` whenever a template column header intentionally differs from the LookML field's `label:`.
   - Put sort-only helper dimensions in `hidden_fields: [...]` and `sorts: [...]` so row order is deterministic without adding extra CSV columns.
   - Keep `chart:<tile_slug>` tiles unpivoted (`fields: [category_dim, measure_1, ...]`) so CSV column `0` is categories and columns `1..K` match the `K` chart series.
   - Use `type: field_filter` whenever `model` / `explore` / `field` is specified on dashboard `filters:`, set a concrete `default_value` on date partition filters, and wire `listen:` on every element.
   - Validate the LookML project (`validate_project` / `lkr tools lookml push`) and verify `pptx-merge validate <template.pptx> --csv-zip <exported.zip>` passes cleanly.
