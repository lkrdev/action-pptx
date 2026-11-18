output "service_url" {
  value       = google_cloud_run_v2_service.service.uri
  description = "The URL of the deployed PowerPoint Action Hub service."
}

output "gcs_bucket_name" {
  value       = google_storage_bucket.pptx_bucket.name
  description = "The GCS bucket provisioned for templates and generated presentations."
}

output "templates_gcs_uri" {
  value       = "gs://${google_storage_bucket.pptx_bucket.name}/templates"
  description = "Templates GCS URI to configure in Looker Admin -> Actions."
}

output "output_gcs_uri" {
  value       = "gs://${google_storage_bucket.pptx_bucket.name}/output"
  description = "Output GCS Package Bucket URI to configure in Looker Admin -> Actions."
}

output "artifact_registry_repository_url" {
  value       = "${var.region}-docker.pkg.dev/${var.project_id}/${google_artifact_registry_repository.repo.repository_id}"
  description = "The Artifact Registry repository URL for pushing your container image."
}

output "load_balancer_ip" {
  value       = var.enable_external_lb ? google_compute_global_address.lb_ip[0].address : null
  description = "Global static IP of the External HTTPS Load Balancer (point lb_domain DNS A record here)."
}

output "load_balancer_url" {
  value       = var.enable_external_lb && var.lb_domain != "" ? "https://${var.lb_domain}" : null
  description = "HTTPS URL of the External Load Balancer when enabled."
}
