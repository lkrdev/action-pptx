terraform {
  required_version = ">= 1.3.0"
  required_providers {
    google = {
      source  = "hashicorp/google"
      version = ">= 5.0.0, < 7.0.0"
    }
    random = {
      source  = "hashicorp/random"
      version = ">= 3.0.0"
    }
  }
}

provider "google" {
  project = var.project_id
  region  = var.region
}

# 1. Enable Google Cloud APIs
resource "google_project_service" "services" {
  for_each = toset(concat([
    "run.googleapis.com",
    "secretmanager.googleapis.com",
    "storage.googleapis.com",
    "cloudbuild.googleapis.com",
    "artifactregistry.googleapis.com",
    "iam.googleapis.com"
    ],
    var.enable_external_lb ? ["compute.googleapis.com"] : []
  ))
  service            = each.key
  disable_on_destroy = false
}

# 2. Create a dedicated Service Account for Cloud Run
resource "google_service_account" "sa" {
  account_id   = "action-pptx-sa"
  display_name = "PowerPoint Template Merge (action-pptx) Service Account"
  description  = "Dedicated service account for Looker PowerPoint Action Hub"
  depends_on   = [google_project_service.services]
}

# 3. Create GCS Bucket for Templates and Outputs
resource "google_storage_bucket" "pptx_bucket" {
  name                        = var.gcs_bucket_name != "" ? var.gcs_bucket_name : "${var.project_id}-action-pptx"
  location                    = var.region
  uniform_bucket_level_access = true
  depends_on                  = [google_project_service.services]
}

# 4. Generate and store Action Hub Secret Key
resource "random_id" "action_hub_secret" {
  byte_length = 32
}

resource "google_secret_manager_secret" "action_hub_secret" {
  secret_id = "action-hub-secret"
  replication {
    auto {}
  }
  depends_on = [google_project_service.services]
}

resource "google_secret_manager_secret_version" "action_hub_secret" {
  secret      = google_secret_manager_secret.action_hub_secret.id
  secret_data = random_id.action_hub_secret.hex
}

# 5. Grant Secret Manager Secret Accessor role to the Service Account
resource "google_secret_manager_secret_iam_member" "action_hub_secret_accessor" {
  secret_id = google_secret_manager_secret.action_hub_secret.id
  role      = "roles/secretmanager.secretAccessor"
  member    = "serviceAccount:${google_service_account.sa.email}"
}

# 6. Grant Storage Object Admin on the GCS Bucket to the Service Account
resource "google_storage_bucket_iam_member" "bucket_object_admin" {
  bucket = google_storage_bucket.pptx_bucket.name
  role   = "roles/storage.objectAdmin"
  member = "serviceAccount:${google_service_account.sa.email}"
}

# 7. Create Artifact Registry Repository for the Docker image
resource "google_artifact_registry_repository" "repo" {
  location      = var.region
  repository_id = "action-pptx"
  description   = "Docker repository for Looker PowerPoint Action Hub (action-pptx)"
  format        = "DOCKER"
  depends_on    = [google_project_service.services]
}

# 8. Deploy Cloud Run Service
resource "google_cloud_run_v2_service" "service" {
  name                 = "action-pptx"
  location             = var.region
  ingress              = var.ingress
  invoker_iam_disabled = var.invoker_iam_disabled

  template {
    service_account                  = google_service_account.sa.email
    max_instance_request_concurrency = 2
    timeout                          = "540s"

    containers {
      image = var.image

      resources {
        limits = {
          cpu    = "2.0"
          memory = "2Gi"
        }
      }

      env {
        name  = "ALLOWED_GCS_BUCKETS"
        value = google_storage_bucket.pptx_bucket.name
      }

      env {
        name  = "ACTION_HUB_LABEL"
        value = "Looker PowerPoint Action Hub"
      }

      env {
        name = "PUBLIC_BASE_URL"
        value = (
          var.action_hub_base_url != ""
          ? var.action_hub_base_url
          : (var.enable_external_lb && var.lb_domain != "" ? "https://${var.lb_domain}" : "http://placeholder")
        )
      }

      env {
        name = "ACTION_HUB_SECRET"
        value_source {
          secret_key_ref {
            secret  = google_secret_manager_secret.action_hub_secret.secret_id
            version = "latest"
          }
        }
      }

      dynamic "env" {
        for_each = var.enable_iam_scheduler && var.looker_base_url != "" ? [1] : []
        content {
          name  = "LOOKERSDK_BASE_URL"
          value = trimsuffix(var.looker_base_url, "/")
        }
    }
  }

  depends_on = [
    google_project_service.services,
    google_secret_manager_secret_version.action_hub_secret,
    google_secret_manager_secret_iam_member.action_hub_secret_accessor,
    google_storage_bucket_iam_member.bucket_object_admin,
  ]
}

# 9. Grant Logs Writer role to the Service Account
resource "google_project_iam_member" "sa_logging_log_writer" {
  project = var.project_id
  role    = "roles/logging.logWriter"
  member  = "serviceAccount:${google_service_account.sa.email}"
}

# 10. Optional External Global HTTPS Application Load Balancer + Cloud Armor
resource "google_compute_global_address" "lb_ip" {
  count      = var.enable_external_lb ? 1 : 0
  name       = "action-pptx-lb-ip"
  depends_on = [google_project_service.services]
}

resource "google_compute_region_network_endpoint_group" "serverless_neg" {
  count                 = var.enable_external_lb ? 1 : 0
  name                  = "action-pptx-neg"
  network_endpoint_type = "SERVERLESS"
  region                = var.region
  cloud_run {
    service = google_cloud_run_v2_service.service.name
  }
}

resource "google_compute_security_policy" "cloud_armor" {
  count       = var.enable_external_lb ? 1 : 0
  name        = "action-pptx-armor"
  description = "Cloud Armor WAF & CIDR allowlist for action-pptx"

  rule {
    action   = length(var.allowed_source_cidrs) > 0 ? "deny(403)" : "allow"
    priority = 2147483647
    match {
      versioned_expr = "SRC_IPS_V1"
      config {
        src_ip_ranges = ["*"]
      }
    }
    description = "Default rule"
  }

  dynamic "rule" {
    for_each = length(var.allowed_source_cidrs) > 0 ? [1] : []
    content {
      action   = "allow"
      priority = 1000
      match {
        versioned_expr = "SRC_IPS_V1"
        config {
          src_ip_ranges = var.allowed_source_cidrs
        }
      }
      description = "Allowed Looker / corporate CIDRs"
    }
  }

  depends_on = [google_project_service.services]
}

resource "google_compute_backend_service" "backend" {
  count                 = var.enable_external_lb ? 1 : 0
  name                  = "action-pptx-backend"
  load_balancing_scheme = "EXTERNAL_MANAGED"
  security_policy       = google_compute_security_policy.cloud_armor[0].id

  backend {
    group = google_compute_region_network_endpoint_group.serverless_neg[0].id
  }
}

resource "google_compute_url_map" "url_map" {
  count           = var.enable_external_lb ? 1 : 0
  name            = "action-pptx-url-map"
  default_service = google_compute_backend_service.backend[0].id
}

resource "google_compute_managed_ssl_certificate" "ssl_cert" {
  count = var.enable_external_lb ? 1 : 0
  name  = "action-pptx-ssl-cert"
  managed {
    domains = [var.lb_domain]
  }
  depends_on = [google_project_service.services]
}

resource "google_compute_target_https_proxy" "https_proxy" {
  count            = var.enable_external_lb ? 1 : 0
  name             = "action-pptx-https-proxy"
  url_map          = google_compute_url_map.url_map[0].id
  ssl_certificates = [google_compute_managed_ssl_certificate.ssl_cert[0].id]
}

resource "google_compute_global_forwarding_rule" "https_fr" {
  count                 = var.enable_external_lb ? 1 : 0
  name                  = "action-pptx-https-fr"
  load_balancing_scheme = "EXTERNAL_MANAGED"
  target                = google_compute_target_https_proxy.https_proxy[0].id
  ip_address            = google_compute_global_address.lb_ip[0].id
  port_range            = "443"
}
