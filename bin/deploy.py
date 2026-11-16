#!/usr/bin/env uv run
# /// script
# requires-python = ">=3.14"
# dependencies = [
#     "typer>=0.9.0",
#     "rich>=13.0.0",
#     "questionary>=2.0.0",
# ]
# ///

from __future__ import annotations

import hashlib
import hmac
import os
import secrets
import subprocess
import sys
import time
from pathlib import Path

import questionary
import typer
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

console = Console()


def run_cmd(
    cmd: list[str],
    *,
    check: bool = True,
    capture_output: bool = False,
    env: dict[str, str] | None = None,
    input_text: str | None = None,
) -> subprocess.CompletedProcess[str]:
    """Run a shell command with consistent logging and error handling."""
    try:
        return subprocess.run(
            cmd,
            check=check,
            text=True,
            capture_output=capture_output,
            env=env,
            input=input_text,
            stdin=None if input_text is not None else subprocess.DEVNULL,
        )
    except subprocess.CalledProcessError as exc:
        console.print(f"\n[bold red]Error running command:[/bold red] {' '.join(cmd)}")
        if exc.stderr:
            console.print(f"[red]{exc.stderr.strip()}[/red]")
        sys.exit(1)


def run_cmd_with_retry(
    cmd: list[str],
    error_substrings: list[str],
    max_retries: int = 12,
    delay: int = 5,
) -> subprocess.CompletedProcess[str]:
    """Run a command, retrying if stderr contains any of the transient error_substrings."""
    for attempt in range(1, max_retries + 1):
        res = run_cmd(cmd, check=False, capture_output=True)
        if res.returncode == 0:
            return res
        stderr_msg = res.stderr or ""
        if any(sub in stderr_msg for sub in error_substrings) and attempt < max_retries:
            console.print(
                f"  [yellow]Command failed, retrying in {delay}s... "
                f"(attempt {attempt}/{max_retries})[/yellow]"
            )
            time.sleep(delay)
        else:
            console.print(
                f"\n[bold red]Error running command after {attempt} attempts:[/bold red] "
                f"{' '.join(cmd)}\n[red]{stderr_msg.strip()}[/red]"
            )
            sys.exit(1)
    sys.exit(1)


def main(
    project_id: str | None = typer.Option(
        None, "--project-id", "-p", help="Google Cloud Project ID"
    ),
    region: str | None = typer.Option(
        None, "--region", "-r", help="Cloud Run region (e.g. us-central1)"
    ),
    bucket_name: str | None = typer.Option(
        None, "--bucket", "-b", help="GCS bucket for templates and outputs"
    ),
    service_account_email: str | None = typer.Option(
        None, "--service-account-email", help="Specific service account email to run the service"
    ),
    from_source: bool = typer.Option(
        False,
        "--source/--prebuilt-image",
        help="Build from local source instead of prebuilt Artifact Registry image",
    ),
    register_looker: bool | None = typer.Option(
        None,
        "--register-looker/--no-register-looker",
        help="Automatically register the action in Looker",
    ),
    looker_url: str | None = typer.Option(None, "--looker-url", help="Looker Base URL"),
    looker_client_id: str | None = typer.Option(
        None, "--looker-client-id", help="Looker Client ID"
    ),
    looker_client_secret: str | None = typer.Option(
        None, "--looker-client-secret", help="Looker Client Secret"
    ),
    oauth_client_id: str | None = typer.Option(
        None,
        "--oauth-client-id",
        envvar="GOOGLE_OAUTH_CLIENT_ID",
        help="Shared Google Web OAuth 2.0 Client ID (used by both Custom Action and Extension)",
    ),
    oauth_client_secret: str | None = typer.Option(
        None,
        "--oauth-client-secret",
        envvar="GOOGLE_OAUTH_CLIENT_SECRET",
        help="Shared Google Web OAuth 2.0 Client Secret (for Custom Action offline tokens)",
    ),
) -> None:
    console.print(
        Panel.fit(
            "[bold green]PowerPoint Template Merge (action-pptx) Deployment Wizard[/bold green]",
            border_style="green",
            padding=(0, 4),
        )
    )

    try:
        current_project = run_cmd(
            ["gcloud", "config", "get-value", "project"], capture_output=True
        ).stdout.strip()
    except Exception as exc:
        console.print(
            "[bold red]Error: gcloud is not authenticated or not installed.[/bold red]\n"
            "Please run [bold yellow]gcloud auth login[/bold yellow] first."
        )
        raise typer.Exit(1) from exc

    project_id = (
        project_id
        or questionary.text("Enter Google Cloud Project ID:", default=current_project or "").ask()
    )
    if not project_id:
        console.print("[bold red]Error: Project ID is required.[/bold red]")
        raise typer.Exit(1)

    os.environ["CLOUDSDK_CORE_PROJECT"] = project_id

    if not region:
        region = questionary.select(
            "Select Cloud Run region:",
            choices=[
                "us-central1",
                "us-east1",
                "us-west1",
                "europe-west1",
                "asia-east1",
                "Other (type manually)",
            ],
            default="us-central1",
        ).ask()
        if region == "Other (type manually)":
            region = questionary.text("Enter Cloud Run region:").ask()
    region = region or "us-central1"

    default_bucket = f"{project_id}-action-pptx"
    bucket_name = (
        bucket_name
        or questionary.text(
            "Enter GCS Bucket name for .pptx templates and outputs:", default=default_bucket
        ).ask()
        or default_bucket
    )
    bucket_name = bucket_name.removeprefix("gs://").strip("/")
    templates_gcs_uri = f"gs://{bucket_name}/templates"
    output_gcs_uri = f"gs://{bucket_name}/output"

    if register_looker is None:
        register_looker = bool(
            questionary.confirm(
                "Do you want to automatically register this action in Looker?", default=False
            ).ask()
        )

    if register_looker:
        console.print("\n[bold yellow]Note:[/bold yellow] Looker API credentials are not stored.")
        looker_url = (
            looker_url
            or questionary.text(
                "Enter Looker Base URL (e.g. https://yourcompany.cloud.looker.com):"
            ).ask()
        )
        looker_client_id = looker_client_id or questionary.text("Enter Looker Client ID:").ask()
        looker_client_secret = (
            looker_client_secret or questionary.password("Enter Looker Client Secret:").ask()
        )

    apis = [
        "run.googleapis.com",
        "secretmanager.googleapis.com",
        "storage.googleapis.com",
        "cloudbuild.googleapis.com",
        "artifactregistry.googleapis.com",
    ]
    with console.status("[bold green]Enabling Google Cloud APIs (this may take a minute)..."):
        run_cmd(["gcloud", "services", "enable", *apis], capture_output=True)
    console.print("[bold green]OK[/bold green] Google Cloud APIs enabled.")

    if (
        run_cmd(
            ["gcloud", "storage", "buckets", "describe", f"gs://{bucket_name}"],
            check=False,
            capture_output=True,
        ).returncode
        != 0
    ):
        console.print(
            f"Creating GCS bucket [bold cyan]gs://{bucket_name}[/bold cyan] in {region}..."
        )
        run_cmd(
            [
                "gcloud",
                "storage",
                "buckets",
                "create",
                f"gs://{bucket_name}",
                f"--location={region}",
            ]
        )
    else:
        console.print(f"GCS bucket [bold green]gs://{bucket_name}[/bold green] already exists.")

    repo_root = Path(__file__).resolve().parent.parent
    sample_template = repo_root / "examples" / "simple_template.pptx"
    if sample_template.is_file():
        console.print(
            "Uploading sample template to "
            f"[bold cyan]{templates_gcs_uri}/simple_template.pptx[/bold cyan]..."
        )
        run_cmd(
            [
                "gcloud",
                "storage",
                "cp",
                str(sample_template),
                f"{templates_gcs_uri}/simple_template.pptx",
            ],
            capture_output=True,
        )

    if not service_account_email:
        project_number = run_cmd(
            ["gcloud", "projects", "describe", project_id, "--format=value(projectNumber)"],
            capture_output=True,
        ).stdout.strip()
        service_account_email = f"{project_number}-compute@developer.gserviceaccount.com"

    console.print(f"Granting runtime roles to [bold cyan]{service_account_email}[/bold cyan]...")
    for role in (
        "roles/secretmanager.secretAccessor",
        "roles/storage.objectAdmin",
        "roles/logging.logWriter",
    ):
        run_cmd_with_retry(
            [
                "gcloud",
                "projects",
                "add-iam-policy-binding",
                project_id,
                f"--member=serviceAccount:{service_account_email}",
                f"--role={role}",
            ],
            error_substrings=["does not exist", "INVALID_ARGUMENT"],
        )

    if (
        run_cmd(
            ["gcloud", "secrets", "describe", "action-hub-secret"],
            check=False,
            capture_output=True,
        ).returncode
        == 0
    ):
        console.print("Retrieving existing [bold cyan]action-hub-secret[/bold cyan] secret...")
        action_hub_secret_val = run_cmd(
            ["gcloud", "secrets", "versions", "access", "latest", "--secret=action-hub-secret"],
            capture_output=True,
        ).stdout.strip()
    else:
        action_hub_secret_val = secrets.token_hex(32)
        console.print("Creating secret [bold cyan]action-hub-secret[/bold cyan]...")
        run_cmd(
            ["gcloud", "secrets", "create", "action-hub-secret", "--replication-policy=automatic"]
        )
        run_cmd(
            ["gcloud", "secrets", "versions", "add", "action-hub-secret", "--data-file=-"],
            capture_output=True,
            input_text=action_hub_secret_val,
        )

    console.print("\n[bold cyan]Deploying action-pptx to Google Cloud Run...[/bold cyan]")
    source_args = (
        ["--source", str(repo_root)]
        if from_source
        else [
            "--image",
            "us-central1-docker.pkg.dev/lkr-dev-production/looker-action/action-pptx:latest",
        ]
    )
    env_vars = [
        f"ALLOWED_GCS_BUCKETS={bucket_name}",
        "ACTION_HUB_LABEL=Looker PowerPoint Action Hub",
        "PUBLIC_BASE_URL=http://placeholder",
    ]
    if oauth_client_id:
        env_vars.extend(
            [
                f"GOOGLE_OAUTH_CLIENT_ID={oauth_client_id}",
                f"GOOGLE_OAUTH_CLIENT_SECRET={oauth_client_secret or ''}",
                "OAUTH_GCS_PASSTHROUGH=true",
            ]
        )

    run_cmd(
        [
            "gcloud",
            "run",
            "deploy",
            "action-pptx",
            "--platform=managed",
            f"--region={region}",
            "--no-invoker-iam-check",
            "--cpu=2",
            "--memory=2Gi",
            "--timeout=540s",
            "--concurrency=2",
            f"--service-account={service_account_email}",
            f"--set-env-vars={','.join(env_vars)}",
            "--set-secrets=ACTION_HUB_SECRET=action-hub-secret:latest",
            *source_args,
        ]
    )

    service_url = run_cmd(
        [
            "gcloud",
            "run",
            "services",
            "describe",
            "action-pptx",
            "--platform=managed",
            f"--region={region}",
            "--format=value(status.url)",
        ],
        capture_output=True,
    ).stdout.strip()
    console.print(f"Deployed Service URL: [bold green]{service_url}[/bold green]")

    run_cmd(
        [
            "gcloud",
            "run",
            "services",
            "update",
            "action-pptx",
            "--platform=managed",
            f"--region={region}",
            f"--update-env-vars=PUBLIC_BASE_URL={service_url}",
        ],
        capture_output=True,
    )

    nonce = secrets.token_bytes(32).hex()
    digest = hmac.new(
        action_hub_secret_val.encode("utf-8"), nonce.encode("utf-8"), hashlib.sha512
    ).hexdigest()
    api_key_token = f"{nonce}/{digest}"

    if register_looker:
        console.print("\n[bold cyan]Registering action with Looker Instance...[/bold cyan]")
        env = {
            **os.environ,
            "SERVICE_URL": service_url,
            "API_KEY_TOKEN": api_key_token,
            "TEMPLATES_GCS_URI": templates_gcs_uri,
            "OUTPUT_GCS_URI": output_gcs_uri,
            "LOOKERSDK_CLIENT_ID": looker_client_id or "",
            "LOOKERSDK_CLIENT_SECRET": looker_client_secret or "",
            "LOOKERSDK_BASE_URL": looker_url or "",
        }
        run_cmd(
            ["uv", "run", str(Path(__file__).resolve().parent / "register_integration.py")], env=env
        )

    table = Table(
        title="Looker Action Hub Registration Details",
        show_header=True,
        header_style="bold magenta",
    )
    table.add_column("Parameter", style="cyan", width=32)
    table.add_column("Value", style="green")
    table.add_row("Action Hub URL", service_url)
    table.add_row("Authorization Token (API Key)", api_key_token)
    table.add_row("Templates GCS URI", templates_gcs_uri)
    table.add_row("Output GCS Package Bucket URI", output_gcs_uri)

    console.print(
        "\n",
        Panel.fit(
            table,
            title="[bold green]Deployment & Configuration Successful![/bold green]",
            border_style="green",
            padding=(1, 2),
        ),
    )


if __name__ == "__main__":
    typer.run(main)
