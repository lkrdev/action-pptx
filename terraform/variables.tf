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
  description = "The URL of the deployed Cloud Run service (or Load Balancer HTTPS URL). Leave empty on first run, then re-apply."
  default     = ""
}

variable "ingress" {
  type        = string
  description = "Cloud Run ingress setting (e.g. INGRESS_TRAFFIC_INTERNAL_AND_CLOUD_LOAD_BALANCING or INGRESS_TRAFFIC_ALL)."
  default     = "INGRESS_TRAFFIC_INTERNAL_AND_CLOUD_LOAD_BALANCING"

  validation {
    condition = contains([
      "INGRESS_TRAFFIC_ALL",
      "INGRESS_TRAFFIC_INTERNAL_ONLY",
      "INGRESS_TRAFFIC_INTERNAL_AND_CLOUD_LOAD_BALANCING",
    ], var.ingress)
    error_message = "ingress must be a valid Cloud Run v2 ingress setting."
  }
}

variable "invoker_iam_disabled" {
  type        = bool
  description = "Disable Cloud Run IAM invoker checks (--no-invoker-iam-check). Must be false when Org Policy constraints/run.managed.requireInvokerIam is enforced."
  default     = true
}

variable "enable_external_lb" {
  type        = bool
  description = "Provision an External Global HTTPS Application Load Balancer, Serverless NEG, Google-managed SSL cert, and Cloud Armor policy."
  default     = false
}

variable "lb_domain" {
  type        = string
  description = "Custom HTTPS domain name for the External Load Balancer Google-managed SSL certificate."
  default     = ""
}

variable "allowed_source_cidrs" {
  type        = list(string)
  description = "Optional CIDR ranges allowed through Cloud Armor to the External Load Balancer. Empty list permits all traffic."
  default     = []
}
