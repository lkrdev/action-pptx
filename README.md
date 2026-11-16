# PowerPoint Template Merge Looker Action Hub (`action-pptx`)

A standalone Looker Action Hub (`action-pptx`), companion Python CLI (`pptx-merge`), and Looker Template Manager Extension (`extension/`) that populates native, editable Microsoft PowerPoint (`.pptx`) templates from scheduled Looker Dashboards.

---

## Why Use This Action

Standard PDF or PNG dashboard exports from Looker flatten slides into static images, preventing analysts and consultants from editing narrative callouts or clicking **Edit Data** on native Office charts. `action-pptx` enables you to:

- **Bind directly inside PowerPoint:** Connect native Office charts (`chart:<tile_slug>`) and dynamic tables (`table:<tile_slug>`) via PowerPoint's built-in **Selection Pane**, and populate narrative callouts or fixed-layout tables via `{{tile_slug[row].column | fmt}}` text tags.
- **Keep native Office charts editable:** Consumes Looker's `unformatted` `csv_zip` webhook payload, updates both `<c:chart>` XML caches and embedded `ppt/embeddings/*.xlsx` workbooks, and verifies 100% XML-to-XLSX cell parity via `openpyxl` so **Edit Data** in desktop PowerPoint works without repair prompts.
- **Self-service validation & upload:** Includes a Looker Template Manager Extension (`extension/`) that validates `.pptx` bindings against live LookML Dashboard schemas (`core40SDK`) before uploading to GCS.
- **Preserve immutable package history:** Writes every scheduled or *Send Now* run to a timestamped path under `output_gcs_uri` so prior runs and manually edited drafts are never overwritten.

---

## Template Bindings & Placeholders

Design templates in PowerPoint using Selection Pane names and text tags:

| Binding / Expression | Where to Place | Description | Example |
| :--- | :--- | :--- | :--- |
| `chart:<tile_slug>` | Selection Pane (Chart) | Binds a native Office chart to `<tile_slug>.csv` (col 0 = categories, cols 1..K = series) | `chart:monthly_trend` |
| `table:<tile_slug>` | Selection Pane (Table) | Clones the last row of the table as a prototype row for each CSV row | `table:top_brands` |
| `{{tile_slug.column}}` | Any text box or cell | Injects row `0` of `column` from `<tile_slug>.csv` | `{{kpi_summary.Total Spend}}` |
| `{{tile_slug[row].column \| fmt}}` | Any text box or cell | Zero-indexed row lookup with deterministic format mask (`#,##0`, `0.0`, `0.00`, `$#,##0`, `$#,##0.00`, `0%`, `0.0%`, `0.00%`, `yyyy-mm-dd`) | `{{mixed_metrics[0].Delta \| 0.0%}}` |
| `{{.column \| fmt}}` | Prototype row of `table:<tile_slug>` | Formats current row's column inside a dynamic N-row table | `{{.Net Revenue \| $#,##0}}` |
| `{{_plan.title}}` | Any text box or cell | Scheduled plan title | `Q3 Executive Business Review` |
| `{{_plan.run_at \| yyyy-mm-dd}}` | Any text box or cell | UTC timestamp when the schedule ran | `2026-11-13` |
| `{{_plan.url}}` | Any text box or cell | Source Looker dashboard URL | `https://instance.looker.com/dashboards/1` |

---

## Deployment

[![Open in Cloud Shell](https://gstatic.com/cloudssh/images/open-btn.svg)](https://ssh.cloud.google.com/cloudshell/editor?shellonly=true&cloudshell_git_repo=https://github.com/lkrdev/action-pptx)

Deploy directly to Google Cloud Run using the deployment wizard:

```bash
./deploy.sh
```

To view all CLI flags, options, and automated Looker registration parameters, run:

```bash
./deploy.sh --help
```

The script enables required Google Cloud APIs (`run`, `secretmanager`, `storage`, `cloudbuild`, `artifactregistry`), creates or reuses your GCS bucket and uploads `examples/simple_template.pptx`, configures `action-hub-secret` in Secret Manager, deploys `action-pptx` to Cloud Run (Python 3.14), and can optionally register and enable the Action Hub in your Looker instance.

---

## Example LookML & Sample Template

- **Sample 3-Slide Widescreen Template:** [`examples/simple_template.pptx`](examples/simple_template.pptx)
- **Example LookML Project (`action_pptx`):** [`examples/lookml/`](examples/lookml/) queries `bigquery-public-data.thelook_ecommerce.inventory_items` over the `looker-private-demo` connection and provides the 9-tile `executive_business_review` dashboard matching `examples/simple_template.pptx`:

```bash
# Push and deploy the example LookML project to Looker
lkr tools lookml push examples/lookml --project-id=action_pptx --deploy
```

---

## Local Development & CLI

```bash
# Install dependencies (Python 3.14 + uv) and pre-commit hooks
uv sync
uv run pre-commit install

# Run linter, type checker, and test suite
make test

# Start local hot-reloading server + Cloudflare Quick Tunnel
make server   # terminal 1
make tunnel   # terminal 2

# Standalone CLI (inspect, validate, merge)
uv run pptx-merge inspect examples/simple_template.pptx
uv run pptx-merge validate examples/simple_template.pptx --csv-zip examples/sample_looker_export.zip
uv run pptx-merge merge examples/simple_template.pptx --csv-zip examples/sample_looker_export.zip -o examples/merged_output.pptx
```

---

## Looker Template Manager Extension Setup

Add the following [`manifest.lkml`](manifest.lkml) configuration to any LookML project, deploy to production, refresh Looker, and open **Applications > PowerPoint Template Manager (action-pptx)**:

```lookml
application: action_pptx_template_manager {
  label: "PowerPoint Template Manager (action-pptx)"
  url: "https://cdn.lkr.dev/apps/action-pptx/latest/bundle.js"
  mount_points: {
    standalone: yes
  }
  entitlements: {
    core_api_methods: [
      "me",
      "search_dashboards",
      "dashboard",
      "lookml_model_explore",
      "merge_query",
      "run_query",
      "run_inline_query",
      "all_integration_hubs",
      "all_integrations",
      "all_user_attributes",
      "create_user_attribute",
      "set_user_attribute_group_values"
    ]
    scoped_user_attributes: ["action_hub_secret"]
    external_api_urls: [
      "https://*.a.run.app",
      "https://*.trycloudflare.com",
      "https://oauth2.googleapis.com"
    ]
    oauth2_urls: [
      "https://accounts.google.com/o/oauth2/v2/auth",
      "https://oauth2.googleapis.com/token"
    ]
    local_storage: yes
    use_form_submit: yes
  }
}
```

1. As a Looker Admin, open the Extension and navigate to `/config` once:
   - Click **Auto-Discover from Action Hub** to pre-fill the Cloud Run URL and `templates_gcs_uri` from the registered `action-pptx` integration.
   - Enter the Action Hub secret to store it in the hidden `action_hub_secret` scoped user attribute (used by `extensionSDK.serverProxy`), or enable **Google OAuth Mode** if running on a strict RFC1918 Private IP Local FQDN Looker instance.
2. To build the extension bundle locally from source (`extension/dist/bundle.js`):

```bash
cd extension
npm install
npm run typecheck
npm run build
```
