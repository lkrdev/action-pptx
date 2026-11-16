variable "project_id" {
  type        = string
  description = "The Google Cloud Project ID to deploy to."
}

variable "region" {
  type        = string
  description = "The Google Cloud region to deploy resources to."
  default     = "us-central1"
}

variable "gcs_bucket_name" {
  type        = string
  description = "The GCS bucket name for storing .pptx templates and generated outputs. Defaults to <project_id>-action-pptx."
  default     = ""
}

variable "image" {
  type        = string
  description = "The Docker image to deploy."
  default     = "us-central1-docker.pkg.dev/lkr-dev-production/looker-action/action-pptx:latest"
}

variable "action_hub_base_url" {
  type        = string
  description = "The URL of the deployed Cloud Run service. Leave empty on the first run, then populate with the outputted service URL and re-apply."
  default     = ""
}
