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
  for_each = toset([
    "run.googleapis.com",
    "secretmanager.googleapis.com",
    "storage.googleapis.com",
    "cloudbuild.googleapis.com",
    "artifactregistry.googleapis.com",
    "iam.googleapis.com"
  ])
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
  ingress              = "INGRESS_TRAFFIC_ALL"
  invoker_iam_disabled = true

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
        name  = "PUBLIC_BASE_URL"
        value = var.action_hub_base_url != "" ? var.action_hub_base_url : "http://placeholder"
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
