# Deployment

[![Open in Cloud Shell](https://gstatic.com/cloudssh/images/open-btn.svg)](https://ssh.cloud.google.com/cloudshell/editor?shellonly=true&cloudshell_git_repo=https://github.com/lkrdev/action-pptx)

Deploy directly to Google Cloud Run using the interactive wizard:

```bash
./deploy.sh
```

See the auto-generated [Deploy CLI Reference](deploy_cli_docs.md) (or run `./deploy.sh --help`) for every flag, environment variable, and default.

## What `./deploy.sh` Provisions

The wizard enables required Google Cloud APIs (`run`, `secretmanager`, `storage`, `cloudbuild`, `artifactregistry`), creates or reuses your GCS bucket and uploads sample templates from `assets/`, creates or reuses `action-hub-secret` in Secret Manager, grants scoped runtime IAM roles, deploys `action-pptx` to Cloud Run (Python 3.14), and optionally configures Looker end-to-end:

- **Action Hub & Settings:** Registers or updates the Looker Integration Hub (`url` + HMAC-signed `authorization_token`) and configures `templates_gcs_uri` and `output_gcs_uri` on `action-pptx`.
- **Extension Deployment (`--extension-mode new` or `existing`):**
  - **New LookML Project (`new`):** Lists your Looker database connections and prompts you to choose one, creates a bare-Git LookML project, configures the LookML model + connection, deploys `bundle.js` + `manifest.lkml` to production, and pre-populates the path-scoped hidden `action_hub_secret` user attribute.
  - **Existing LookML Project (`existing`):** Prints the ready-to-paste `application: action_pptx_template_manager { ... }` block with your deployed service URL pre-filled in `external_api_urls`.

## Redeploying & Reconfiguring Looker

`./deploy.sh` checks live GCP and Looker resources so it can be re-run safely:

```bash
# Re-configure Looker Action Hub & Extension only (skips GCP deploy; reads existing secret & URL)
./deploy.sh --looker-only -p <PROJECT_ID> -r <REGION> --public-base-url https://your-proxy.example.com/apps/action-pptx

# Redeploy Cloud Run from local source without re-running Looker setup
./deploy.sh --source --no-register-looker -p <PROJECT_ID> -r <REGION> -b <BUCKET>
```

> **Note on Looker Core / Custom Domains:** Use `--public-base-url` when fronting Cloud Run with a load balancer or reverse proxy so Looker's Action Hub registration, webhook callbacks, and Extension `external_api_urls` target the public URL instead of the default `*.a.run.app` hostname.
