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

## Agent Skills (`npx skills add`)

Install the PowerPoint templating skills from this repository into Gemini CLI, Claude Code, Cursor, Antigravity, or Copilot with [`npx skills`](https://www.npmjs.com/package/skills):

```bash
npx skills add lkrdev/action-pptx
npx skills add lkrdev/action-pptx --skill looker-action-pptx-templates
```

- **[`looker-action-pptx-templates`](skills/looker-action-pptx-templates/SKILL.md)** *(Master Skill)* — Creates or retrofits `.pptx` templates against existing LookML in two modes:
  1. **Start from Scratch:** Takes slide masters/layouts, a general narrative idea, and existing LookML to build native templated slides (`chart:<tile_slug>`, `table:<tile_slug>`, `{{tile_slug[row].Column | fmt}}`) plus a validated `dashboard_tiles.json` spec.
  2. **Retrofit Existing Slides:** Takes a completely built-out static `.pptx` deck without templating and existing LookML, then injects `action-pptx` tags and Selection Pane bindings surgically in place while preserving layouts and typography.
- **[`pptx-merge`](skills/pptx-merge/SKILL.md)** — Reference for `pptx-merge` binding syntax, format masks, and `inspect` / `validate` / `merge` CLI usage.

> **Scope & Companion LookML Skills:** The skills in `lkrdev/action-pptx` **do not build your LookML models, views, or explores**—they assume your LookML already exists and build `.pptx` templates whose tags match the LookML dashboard you need to make (defaulting to **not** generating LookML dashboards unless you explicitly ask for one). To author or refactor LookML models, explores, and dashboards, install the companion Looker skill sets:
> - **[`lkrdev/lookml_skills`](https://github.com/lkrdev/lookml_skills)** (`npx skills add lkrdev/lookml_skills`) — Authoring LookML models, views, explores, fields, refinements, and tests.
> - **[`looker-open-source/looker-skills`](https://github.com/looker-open-source/looker-skills)** (`npx skills add looker-open-source/looker-skills`) — LookML architecture, audits, and LookML dashboard creation (`creating-looker-dashboard`).
> - **[`lkrdev/cli`](https://github.com/lkrdev/cli)** (`npx skills add lkrdev/cli`) — Looker CLI, `lkr-code-mode`, `agentic-lookml`, `looker-field-profiler`, and `lkr-project-helper`.

---

## Deployment

[![Open in Cloud Shell](https://gstatic.com/cloudssh/images/open-btn.svg)](https://ssh.cloud.google.com/cloudshell/editor?shellonly=true&cloudshell_git_repo=https://github.com/lkrdev/action-pptx)

Deploy directly to Google Cloud Run using the deployment wizard:

```bash
./deploy.sh
```

See [deploy.md](deploy.md) for the full deployment walkthrough and redeploy flags, or the auto-generated [Deploy CLI Reference](deploy_cli_docs.md) (`./deploy.sh --help`) for all options.

---

## Configuration & Authentication Modes

`action-pptx` supports both **Service Account** and **Google OAuth** authentication from the same container based on your environment variables:

| Environment Variable | Default | Description |
| :--- | :--- | :--- |
| `ACTION_HUB_SECRET` | *(Required)* | Shared secret verifying Looker Action Hub webhooks (accepts raw secret or `<nonce>/<hmac-sha512>`) and Extension `serverProxy` calls. |
| `PUBLIC_BASE_URL` | `http://localhost:8080` | Public HTTPS origin of your Cloud Run service (e.g. `https://action-pptx-xyz.a.run.app`). |
| `ACTION_HUB_LABEL` | `Looker PowerPoint Action Hub` | Group heading shown in **Looker → Admin → Actions**. |
| `ACTION_LABEL` | `PowerPoint Template Merge (action-pptx)` | Action destination name shown in Looker's Schedule/Send modal and Admin Actions list. |
| `REQUIRED_USER_ATTRIBUTE` | `""` *(disabled)* | Looker User Attribute name (e.g. `can_use_action_pptx`) bound to a `required: true`, `sensitive: true` Action Hub parameter to restrict who can see and execute the action. |
| `RESTRICT_BY_USER_ATTRIBUTE` | `false` | Shorthand boolean (`true`/`false`); when `true` and `REQUIRED_USER_ATTRIBUTE` is unset, defaults to `can_use_action_pptx`. |
| `ALLOWED_GCS_BUCKETS` | `""` *(any bucket)* | Comma-separated allowlist of permitted GCS bucket names (e.g. `my-templates-bucket,my-output-bucket`). |
| `GOOGLE_OAUTH_CLIENT_ID` | `""` *(disabled)* | Google Web OAuth 2.0 Client ID. Leaving this unset runs in **Service Account (no OAuth)** mode; setting it enables **Google OAuth** mode across the Action and Extension. |
| `GOOGLE_OAUTH_CLIENT_SECRET` | `""` | Google Web OAuth 2.0 Client Secret used to exchange authorization codes for offline `refresh_token`s on scheduled deliveries. |
| `OAUTH_GCS_PASSTHROUGH` | `false` | `true` reads/writes GCS as the signed-in Google user (`devstorage.read_write`); `false` uses the Cloud Run Service Account. |
| `ALLOWED_OAUTH_DOMAINS` | `""` *(any verified email)* | Optional comma-separated Google Workspace domain allowlist (e.g. `yourcompany.com`) when OAuth is enabled. |
| `ENCRYPT_PAYLOAD` | `false` | Set `true` to encrypt OAuth token state returned to Looker using `CIPHER_MASTER` (defaults to `ACTION_HUB_SECRET`). |

### Mode 1: Service Account Only (Default — No User OAuth)
Cloud Run accesses GCS via its attached Service Account (`roles/storage.objectAdmin`), and Looker authenticates via `ACTION_HUB_SECRET`. End users never see a Google sign-in prompt in Schedules or the Extension (`useOAuth: false`).

```ini
ACTION_HUB_SECRET=your-64-char-hex-secret
PUBLIC_BASE_URL=https://action-pptx-xyz-uc.a.run.app
ACTION_HUB_LABEL=Looker PowerPoint Action Hub
ACTION_LABEL=PowerPoint Slides
ALLOWED_GCS_BUCKETS=my-project-action-pptx
```

### Mode 2: Google OAuth + Per-User GCS Passthrough
Users sign in once with Google (storing an offline refresh token for unattended schedules), and GCS reads/writes run under each user's own Google IAM permissions (`devstorage.read_write`). Can share a single Web OAuth 2.0 Client ID between the Action and Extension.

```ini
ACTION_HUB_SECRET=your-64-char-hex-secret
PUBLIC_BASE_URL=https://action-pptx-xyz-uc.a.run.app
ACTION_HUB_LABEL=Looker PowerPoint Action Hub
ACTION_LABEL=PowerPoint Slides
ALLOWED_GCS_BUCKETS=my-project-action-pptx
GOOGLE_OAUTH_CLIENT_ID=1234567890-abc.apps.googleusercontent.com
GOOGLE_OAUTH_CLIENT_SECRET=GOCSPX-your-oauth-client-secret
OAUTH_GCS_PASSTHROUGH=true
ALLOWED_OAUTH_DOMAINS=yourcompany.com
ENCRYPT_PAYLOAD=true
```

### Mode 3: Hybrid (Google OAuth Sign-In + Service Account GCS)
Users authenticate with Google OAuth (`openid email` only — handy when `serverProxy` is unavailable on private-IP Looker instances), while Cloud Run still reads/writes GCS through its central Service Account so individual users do not need GCS bucket IAM grants.

```ini
ACTION_HUB_SECRET=your-64-char-hex-secret
PUBLIC_BASE_URL=https://action-pptx-xyz-uc.a.run.app
ACTION_LABEL=PowerPoint Slides
ALLOWED_GCS_BUCKETS=my-project-action-pptx
GOOGLE_OAUTH_CLIENT_ID=1234567890-abc.apps.googleusercontent.com
GOOGLE_OAUTH_CLIENT_SECRET=GOCSPX-your-oauth-client-secret
OAUTH_GCS_PASSTHROUGH=false
ALLOWED_OAUTH_DOMAINS=yourcompany.com
ENCRYPT_PAYLOAD=true
```

### Restricting Action Access to Specific Users or Groups (User Attribute Gate)
Looker does not provide a native per-user/per-group toggle in **Admin → Platform → Actions**. Setting `REQUIRED_USER_ATTRIBUTE=can_use_action_pptx` (or `RESTRICT_BY_USER_ATTRIBUTE=true`) adds a `required: true`, `sensitive: true` parameter bound to that Looker User Attribute (`user_attribute_name`) and enforces it on `/form` and `/execute`:

```ini
REQUIRED_USER_ATTRIBUTE=can_use_action_pptx
```

When enabled, configure the matching attribute in **Looker → Admin → Users → User Attributes**:
1. **Name:** `can_use_action_pptx` (or the value set in `REQUIRED_USER_ATTRIBUTE`), **Data Type:** `String (advanced)`.
2. **Default Value:** Leave **unset (`null`)** — never set an empty string `""`, which Looker treats as populated.
3. **User Access:** Set to **None** or **View** so unauthorized users cannot populate it themselves.
4. **Hide Values:** Keep set to **No** (Looker encrypts hidden user attributes into ciphertext on Action Hub webhooks; `sensitive: true` on the Action Hub parameter masks the field in the Admin UI instead).
5. Assign a non-null value (e.g. `yes`) on the **Group Values** or **User Values** tab only for users/groups who should see and execute the action. Non-admin users without a value will not see the action in the Send/Schedule modal, and direct/scheduled executions without a value are rejected.

---

## Sample Templates & Example LookML (`assets/`)

- **Executive Business Review Template (3 slides):** [`assets/simple_template.pptx`](assets/simple_template.pptx)
- **Looker Embed Demo (`embed_demo::brand_overview`) Template:** [`assets/embed_demo_template.pptx`](assets/embed_demo_template.pptx)
- **Example LookML Project (`action-pptx`):** [`assets/lookml/`](assets/lookml/) queries `bigquery-public-data.thelook_ecommerce.inventory_items` over the `looker-private-demo` connection and provides the 9-tile `executive_business_review` dashboard matching `assets/simple_template.pptx`:

```bash
# Push and deploy the example LookML project to Looker
lkr tools lookml push assets/lookml --project-id=action-pptx --deploy
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
uv run pptx-merge inspect assets/simple_template.pptx
uv run pptx-merge validate assets/simple_template.pptx --csv-zip assets/sample_looker_export.zip
uv run pptx-merge merge assets/simple_template.pptx --csv-zip assets/sample_looker_export.zip -o assets/merged_output.pptx
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
      "all_dashboards",
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
      "delete_user_attribute",
      "all_themes",
      "all_color_collections",
      "search_scheduled_plans",
      "create_scheduled_plan",
      "update_scheduled_plan",
      "delete_scheduled_plan",
      "scheduled_plan_run_once_by_id",
      "artifact",
      "search_artifacts",
      "update_artifacts",
      "delete_artifact"
    ]
    global_user_attributes: ["brand"]
    external_api_urls: [
      # Replace with your deployed Cloud Run URL:
      "https://*.a.run.app"
    ]
    # only required for OAuth2 use in the Action and Extension
    oauth2_urls: [
      "https://accounts.google.com/o/oauth2/v2/auth",
      "https://oauth2.googleapis.com/token"
    ]
    local_storage: yes
    new_window: yes
    use_downloads: yes
    use_form_submit: yes
  }
}
```

1. As a Looker Admin, open the Extension and navigate to `/config` once:
   - Click **Auto-Discover from Action Hub** to pre-fill the Cloud Run URL and `templates_gcs_uri` from the registered `action-pptx` integration.
   - Enter the Action Hub secret to store it in the hidden `action_hub_secret` scoped user attribute (used by `extensionSDK.serverProxy`), or enable **Google OAuth Mode** if running on a strict RFC1918 Private IP Local FQDN Looker instance.
2. Switch between the three built-in views from the header navigation bar:
   - **Builder (`/`)** — Validate `.pptx` bindings against live Looker dashboards, preview themed decks, and upload templates to GCS.
   - **Viewer (`/viewer`)** — Browse, search (`?q=<subpath>`), preview, and download rendered `.pptx` outputs and saved templates.
   - **Scheduler (`/scheduler`)** — Manage **Master Fanout Schedules** (1 dashboard + 1 template + 1 cadence → *N* Looker `ScheduledPlan` filter slices writing to `{output_folder}/{slice_slug}_<timestamp>.pptx`, backed by Looker Artifacts and reconciled via `search_scheduled_plans`). Looker Admins with `see_schedules` can inspect or manage another user's Master Fanout schedules one user at a time by appending `?user_id=<looker_user_id>` (e.g. `/scheduler?user_id=42`).
3. To build the extension bundle locally from source (`extension/dist/bundle.js`):

```bash
cd extension
npm install
npm run typecheck
npm run build
```
