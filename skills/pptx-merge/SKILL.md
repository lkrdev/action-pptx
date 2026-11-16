---
name: pptx-merge
description: >
  Author, inspect, validate, and populate native editable Microsoft PowerPoint (.pptx)
  templates from Looker csv_zip exports using the pptx-merge CLI and library.
---

# PowerPoint Template Merge (`pptx-merge`)

`pptx-merge` populates native Microsoft PowerPoint (`.pptx`) templates from Looker's `unformatted` `csv_zip` exports without external mapping files. All bindings live inside standard PowerPoint properties.

## 1. In-PPTX Binding Grammar

### Tile Slug Normalization
Looker's `DashboardZIPProducer` converts each dashboard tile title into a lowercase CSV filename (`<tile_slug>.csv`) by replacing spaces and special characters (`/?*\:|"<> %`) with underscores (`_`).
- Tile title `"KPI Summary"` -> `kpi_summary`
- Tile title `"Monthly Trend"` -> `monthly_trend`

### Text Tags (`{{tile_slug[row].column | fmt}}`)
Type tags directly inside any slide text box, title, footer, or fixed-layout table cell:
- `{{kpi_summary.Total Spend | $#,##0.00}}` (defaults to row `0`)
- `{{mixed_metrics[2].Value | 0.0%}}` (explicit zero-based row index)
- Built-in schedule metadata tags: `{{_plan.title}}`, `{{_plan.url}}`, `{{_plan.run_at}}`

Even if PowerPoint splits a typed `{{...}}` tag across multiple XML `<a:r>` text runs, `pptx-merge` resolves the tag at the paragraph (`<a:p>`) level while preserving the first run's font family, size, bold/italic, and RGB color.

### Supported Format Masks (`| <fmt>`)
| Mask | Input (`unformatted`) | Output |
| :--- | :--- | :--- |
| `#,##0` | `3378.2` | `3,378` |
| `0.0` | `62.44` | `62.4` |
| `0.00` | `62.446` | `62.45` |
| `$#,##0` | `34855339.4` | `$34,855,339` |
| `$#,##0.00` | `-1234.5` | `-$1,234.50` |
| `0%` | `0.5568` | `56%` |
| `0.0%` | `0.5568` | `55.7%` |
| `0.00%` | `0.5568` | `55.68%` |
| `yyyy-mm-dd` | `2026-11-12T08:30:00Z` | `2026-11-12` |

### Dynamic N-Row Tables (`table:<tile_slug>`)
1. Open **Home -> Arrange -> Selection Pane** in PowerPoint and rename the table shape to `table:<tile_slug>` (e.g., `table:top_brands`).
2. Keep header row(s) at the top. The **last row** of the table acts as the prototype row:
   - Place `{{.Column Name | fmt}}` tags inside the prototype row cells to map specific columns and format masks, OR leave prototype cells without tags to fill columns positionally (`0..N-1`).
3. At merge time, `pptx-merge` clones the prototype XML row once per CSV row (preserving borders, fills, and font styling) and removes the prototype row.

### Native Editable Office Charts (`chart:<tile_slug>`)
1. Rename the chart in PowerPoint's **Selection Pane** to `chart:<tile_slug>` (e.g., `chart:monthly_trend`).
2. Use standard Office charts (`<c:chartSpace>`, including stacked-column/bar charts for waterfalls) rather than Office 2016+ Extended Charts (`chartEx`).
3. Column `0` of `<tile_slug>.csv` supplies category labels; columns `1..K` supply numeric series (`float` values, blanks mapped to `None`).
4. `pptx-merge` updates both the `<c:chart>` XML cache and the embedded `ppt/embeddings/*.xlsx` worksheet and verifies 100% parity with `openpyxl` so **Edit Data** in desktop PowerPoint opens cleanly without repair prompts.

## 2. CLI Usage

```bash
# 1. Inspect a .pptx template for Selection Pane and {{...}} bindings
uv run pptx-merge inspect template.pptx

# 2. Validate a .pptx template against a Looker csv_zip or dashboard schema JSON
uv run pptx-merge validate template.pptx --csv-zip export.zip
uv run pptx-merge validate template.pptx --schema dashboard_tiles.json

# 3. Merge a Looker csv_zip export into a .pptx template
uv run pptx-merge merge template.pptx --csv-zip export.zip -o output.pptx --title "Q4 Report"
```
