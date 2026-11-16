-include .env
export

PORT ?= 9119
GCP_PROJECT ?= fifth-inkwell-452619-v2
GCP_REGION ?= us-east1
GCS_BUCKET ?= $(GCP_PROJECT)-action-pptx
TEMPLATES_GCS_URI ?= gs://$(GCS_BUCKET)/templates
OUTPUT_GCS_URI ?= gs://$(GCS_BUCKET)/output

.PHONY: server tunnel setup-gcs demo-pptx test

server:
	uv run --frozen --env-file=.env uvicorn action_pptx.app:app --reload --reload-dir src --host 0.0.0.0 --port $(PORT)

tunnel:
	cloudflared tunnel --protocol http2 --retries 10 --grace-period 30s --url http://localhost:$(PORT)

setup-gcs: demo-pptx
	@gcloud storage buckets describe "gs://$(GCS_BUCKET)" --project="$(GCP_PROJECT)" >/dev/null 2>&1 || \
		gcloud storage buckets create "gs://$(GCS_BUCKET)" --project="$(GCP_PROJECT)" --location="$(GCP_REGION)"
	@gcloud storage cp examples/simple_template.pptx "$(TEMPLATES_GCS_URI)/simple_template.pptx" --project="$(GCP_PROJECT)"
	@printf "ALLOWED_GCS_BUCKETS=%s\nTEMPLATES_GCS_URI=%s\nOUTPUT_GCS_URI=%s\n" \
		"$(GCS_BUCKET)" "$(TEMPLATES_GCS_URI)" "$(OUTPUT_GCS_URI)" > .env

demo-pptx:
	@mkdir -p examples
	@uv run --frozen python -c 'from pathlib import Path; from tests.conftest import build_showcase_template_bytes, build_showcase_csv_zip_bytes, build_embed_demo_template_bytes, build_embed_demo_csv_zip_bytes; Path("examples/simple_template.pptx").write_bytes(build_showcase_template_bytes()); Path("examples/sample_looker_export.zip").write_bytes(build_showcase_csv_zip_bytes()); Path("examples/embed_demo_template.pptx").write_bytes(build_embed_demo_template_bytes()); Path("examples/embed_demo_export.zip").write_bytes(build_embed_demo_csv_zip_bytes())'
	@uv run --frozen pptx-merge merge examples/simple_template.pptx --csv-zip examples/sample_looker_export.zip -o examples/merged_output.pptx --title "Q3 Executive Business Review" --url "https://looker.example.com/dashboards/81" --run-at "2026-11-13T17:15:00Z"
	@uv run --frozen pptx-merge merge examples/embed_demo_template.pptx --csv-zip examples/embed_demo_export.zip -o examples/embed_demo_merged_dark.pptx --title "Levi's Brand Overview" --url "https://looker.example.com/dashboards/embed_demo::brand_overview" --run-at "2026-11-15T05:00:00Z" --theme Levis_Dark
	@uv run --frozen pptx-merge merge examples/embed_demo_template.pptx --csv-zip examples/embed_demo_export.zip -o examples/embed_demo_merged_light.pptx --title "Levi's Brand Overview" --url "https://looker.example.com/dashboards/embed_demo::brand_overview" --run-at "2026-11-15T05:00:00Z" --theme Levis_Light

test:
	uv run --frozen ruff check .
	uv run --frozen ruff format --check .
	uv run --frozen ty check
	uv run --frozen pytest --cov=src --cov-report=term-missing
