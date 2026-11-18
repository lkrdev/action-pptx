# `./deploy.sh`

**Usage**:

```console
$ ./deploy.sh [OPTIONS]
```

**Options**:

* `-p, --project-id <str>`: Google Cloud Project ID
* `-r, --region <str>`: Cloud Run region (e.g. us-central1)
* `-b, --bucket <str>`: GCS bucket for templates and outputs
* `--service-account-email <str>`: Specific service account email to run the service
* `--source / --prebuilt-image`: Build from local source instead of prebuilt Artifact Registry image  [default: prebuilt-image]
* `--register-looker / --no-register-looker`: Automatically register the Action Hub & Extension in Looker
* `--looker-url <str>`: Looker Base URL
* `--looker-client-id <str>`: Looker Client ID
* `--looker-client-secret <str>`: Looker Client Secret
* `--looker-oauth-account <str>`: Configured lkr-dev-cli OAuth account name (alternative to client ID/secret)
* `--extension-mode <str>`: Extension setup mode: 'new' (deploy to new LookML project), 'existing' (print manifest block), or 'skip'
* `--looker-project <str>`: LookML project ID when deploying the Extension to a new project
* `--looker-model <str>`: LookML model name to configure when deploying to a new project
* `--looker-connection <str>`: Looker database connection name to attach to the new project's model
* `--include-sample-lookml / --no-include-sample-lookml`: Include sample LookML view and Executive Business Review dashboard from assets/lookml
* `--oauth-client-id <str>`: Shared Google Web OAuth 2.0 Client ID (used by both Custom Action and Extension)  [env var: GOOGLE_OAUTH_CLIENT_ID]
* `--oauth-client-secret <str>`: Shared Google Web OAuth 2.0 Client Secret (for Custom Action offline tokens)  [env var: GOOGLE_OAUTH_CLIENT_SECRET]
* `--action-label <str>`: Display label shown for the action in Looker Admin & Schedule modal  [env var: ACTION_LABEL; default: PowerPoint Template Merge (action-pptx)]
* `--required-user-attribute <str>`: Optional Looker User Attribute name to restrict who can see and execute this action  [env var: REQUIRED_USER_ATTRIBUTE]
* `--public-base-url <str>`: Public URL if fronted by a load balancer or proxy (defaults to Cloud Run URL)
* `--looker-only`: Skip GCP provisioning/deploy and only configure Looker (reads existing secret & URL)
* `--deploy-mode <str>`: Exposure mode (both support Actions + Extension): 'load-balanced' (recommended; blocks direct .a.run.app access) or 'public' (direct Cloud Run URL)
* `--lb-type <str>`: Load balancer for load-balanced mode: 'external' (Global HTTPS ALB + Cloud Armor), 'internal' (VPC Internal HTTPS ALB), or 'none' (BYO LB)
* `--lb-domain <str>`: Custom HTTPS domain for Load Balancer SSL cert and Action Hub URL
* `--allowed-cidrs <str>`: Comma-separated CIDR ranges for Cloud Armor allowlist (e.g. Looker outbound IPs)
* `--vpc-network <str>`: VPC network name for internal Load Balancer  [default: default]
* `--vpc-subnet <str>`: VPC subnet name for internal Load Balancer
* `--ssl-cert <str>`: Existing regional SSL certificate resource name for internal Load Balancer
* `--install-completion`: Install completion for the current shell.
* `--show-completion`: Show completion for the current shell, to copy it or customize the installation.
* `--help`: Show this message and exit.
