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
