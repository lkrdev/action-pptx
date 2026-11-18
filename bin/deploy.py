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
import json
import os
import secrets
import shutil
import subprocess
import sys
import time
from pathlib import Path

import questionary
import typer
from rich.console import Console
from rich.panel import Panel
from rich.syntax import Syntax
from rich.table import Table

BIN_DIR = Path(__file__).resolve().parent
if str(BIN_DIR) not in sys.path:
    sys.path.insert(0, str(BIN_DIR))

from deploy_policy import (  # noqa: E402
    VALID_DEPLOY_MODES,
    VALID_LB_TYPES,
    ModeAssessment,
    OrgPolicyReport,
    assess_deploy_modes,
    parse_org_policy_constraint,
)
from register_integration import (  # noqa: E402
    build_manifest_application_block,
    list_looker_connections,
    node_modules_in_sync,
    normalize_service_base,
    register_action_hub_and_extension,
)

console = Console()


def run_cmd(
    cmd: list[str],
    *,
    check: bool = True,
    capture_output: bool = False,
    env: dict[str, str] | None = None,
    input_text: str | None = None,
    cwd: Path | None = None,
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
            cwd=str(cwd) if cwd else None,
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


def _maybe_build_extension_bundle(repo_root: Path) -> str | None:
    ext_dir = repo_root / "extension"
    dist_bundle = ext_dir / "dist" / "bundle.js"
    if (ext_dir / "package.json").is_file() and shutil.which("npm"):
        console.print("[bold cyan]Building Looker Extension bundle via webpack...[/bold cyan]")
        if not node_modules_in_sync(ext_dir):
            run_cmd(["npm", "ci", "--no-audit", "--no-fund"], cwd=ext_dir)
        run_cmd(["npm", "run", "build"], cwd=ext_dir)
    return str(dist_bundle) if dist_bundle.is_file() else None


def inspect_org_policies(project_id: str) -> OrgPolicyReport:
    """Query effective GCP Org Policies on project_id without mutating state."""
    notes: list[str] = []
    parsed: dict[str, dict[str, bool | None]] = {}
    constraints = (
        "constraints/run.managed.requireInvokerIam",
        "constraints/run.allowedIngress",
        "constraints/iam.allowedPolicyMemberDomains",
    )
    for constraint in constraints:
        res = run_cmd(
            [
                "gcloud",
                "org-policies",
                "describe",
                constraint,
                f"--project={project_id}",
                "--effective",
                "--format=json",
            ],
            check=False,
            capture_output=True,
        )
        if res.returncode != 0:
            err = (res.stderr or "").strip()
            if "NOT_FOUND" in err or "was not found" in err.lower():
                parsed[constraint] = parse_org_policy_constraint(constraint, {})
            else:
                parsed[constraint] = parse_org_policy_constraint(constraint, None)
                detail_msg = err.splitlines()[-1] if err else "error"
                notes.append(f"{constraint.split('/')[-1]}: unreadable ({detail_msg})")
            continue
        try:
            data = json.loads(res.stdout or "{}")
            payload = data if isinstance(data, dict) else {}
        except json.JSONDecodeError:
            payload = None
            notes.append(f"{constraint.split('/')[-1]}: invalid JSON output")
        parsed[constraint] = parse_org_policy_constraint(constraint, payload)

    return OrgPolicyReport(
        require_invoker_iam=parsed["constraints/run.managed.requireInvokerIam"].get("enforced"),
        ingress_all_allowed=parsed["constraints/run.allowedIngress"].get("allow_all"),
        ingress_lb_allowed=parsed["constraints/run.allowedIngress"].get("lb_allowed"),
        allowed_member_domains_restricted=parsed["constraints/iam.allowedPolicyMemberDomains"].get(
            "restricted"
        ),
        notes=tuple(notes),
    )


def _render_deployment_options_overview(
    report: OrgPolicyReport,
    assessments: list[ModeAssessment],
) -> None:
    def _fmt_flag(val: bool | None, *, true_text: str, false_text: str) -> str:
        if val is None:
            return "[yellow]Unknown (cannot read Org Policy)[/yellow]"
        return true_text if val else false_text

    policy_table = Table(
        title="GCP Org Policy Preflight",
        show_header=True,
        header_style="bold cyan",
    )
    policy_table.add_column("Org Policy Check", style="white", width=32)
    policy_table.add_column("Status", width=18)
    policy_table.add_column("What It Means", style="white")

    policy_table.add_row(
        "run.managed.requireInvokerIam",
        _fmt_flag(
            report.require_invoker_iam,
            true_text="[red]Enforced[/red]",
            false_text="[green]Not enforced[/green]",
        ),
        "Required off for Looker Action Hub webhooks unless an authenticating proxy sits in front.",
    )
    policy_table.add_row(
        "run.allowedIngress (all)",
        _fmt_flag(
            report.ingress_all_allowed,
            true_text="[green]Allowed[/green]",
            false_text="[red]Restricted[/red]",
        ),
        "When restricted, Cloud Run must sit behind a Load Balancer.",
    )
    policy_table.add_row(
        "iam.allowedPolicyMemberDomains",
        _fmt_flag(
            report.allowed_member_domains_restricted,
            true_text="[yellow]Restricted[/yellow]",
            false_text="[green]Unrestricted[/green]",
        ),
        "Informational: restricts public allUsers IAM bindings, not Load Balancer ingress.",
    )
    console.print("\n", policy_table)
    for note in report.notes:
        console.print(f"[dim]Org Policy note: {note}[/dim]")

    status_by_mode = {a.mode: a.status for a in assessments}
    services_guide = (
        "[bold white]Always provisioned (both options support Looker Actions + "
        "Extension):[/bold white]\n"
        "  • [cyan]Cloud Run (action-pptx)[/cyan] — runs the Action Hub webhooks & Extension API\n"
        "  • [cyan]Cloud Storage Bucket[/cyan] — stores .pptx templates & merged outputs\n"
        "  • [cyan]Secret Manager (action-hub-secret)[/cyan] — shared token authenticating Looker\n"
        "  • [cyan]Service Account + Artifact Registry / Cloud Build[/cyan] — "
        "runtime IAM & image\n\n"
        f"[bold green]Option 1: load-balanced[/bold green] [{status_by_mode['load-balanced']}] "
        "— Recommended (Most Secure)\n"
        "  Locks Cloud Run to [bold]internal-and-cloud-load-balancing[/bold] so direct *.a.run.app "
        "requests are blocked.\n"
        "  You can add:\n"
        "  • [cyan]external[/cyan] (Compute Engine): Global Static IP, Google-managed SSL Cert, "
        "Serverless NEG, & [bold]Cloud Armor[/bold]\n"
        "    to filter traffic by Looker/corp IP CIDRs [dim](requires a DNS domain)[/dim].\n"
        "  • [cyan]internal[/cyan] (Compute Engine): Regional Internal HTTPS ALB + Serverless NEG "
        "inside your VPC\n"
        "    for private-IP Looker setups [dim](requires VPC subnet, proxy-only subnet, "
        "SSL cert)[/dim].\n"
        "  • [cyan]none[/cyan]: No extra LB resources created now; bring your own LB via "
        "--public-base-url.\n\n"
        f"[bold cyan]Option 2: public[/bold cyan] [{status_by_mode['public']}] — Simplest\n"
        "  Exposes Cloud Run directly ([bold]--ingress=all[/bold]). Adds [bold]no extra networking "
        "services[/bold]; requests are\n"
        "  still verified at the app layer via ACTION_HUB_SECRET or Google OAuth."
    )
    console.print(
        "\n",
        Panel(
            services_guide,
            title="[bold cyan]Services Added & Why — Choose Exposure Mode[/bold cyan]",
            border_style="cyan",
            padding=(1, 2),
        ),
    )


def _gcloud_exists(args: list[str]) -> bool:
    return run_cmd(["gcloud", *args], check=False, capture_output=True).returncode == 0


def _ensure_gcloud(describe_args: list[str], create_args: list[str]) -> bool:
    """Run create_args via gcloud if describe_args fails; returns True when created."""
    if _gcloud_exists(describe_args):
        return False
    run_cmd(["gcloud", *create_args])
    return True


def _ensure_serverless_neg(region: str, neg_name: str = "action-pptx-neg") -> None:
    _ensure_gcloud(
        ["compute", "network-endpoint-groups", "describe", neg_name, f"--region={region}"],
        [
            "compute",
            "network-endpoint-groups",
            "create",
            neg_name,
            f"--region={region}",
            "--network-endpoint-type=serverless",
            "--cloud-run-service=action-pptx",
        ],
    )


def _provision_external_lb(
    *,
    region: str,
    lb_domain: str,
    allowed_cidrs: str,
) -> tuple[str, str]:
    """Provision External HTTPS ALB + Serverless NEG + Cloud Armor. Returns (base_url, ip)."""
    ip_name = "action-pptx-lb-ip"
    _ensure_gcloud(
        ["compute", "addresses", "describe", ip_name, "--global"],
        ["compute", "addresses", "create", ip_name, "--ip-version=IPV4", "--global"],
    )
    lb_ip = run_cmd(
        [
            "gcloud",
            "compute",
            "addresses",
            "describe",
            ip_name,
            "--global",
            "--format=value(address)",
        ],
        capture_output=True,
    ).stdout.strip()

    _ensure_serverless_neg(region)

    armor_name = "action-pptx-armor"
    _ensure_gcloud(
        ["compute", "security-policies", "describe", armor_name],
        [
            "compute",
            "security-policies",
            "create",
            armor_name,
            "--description=Cloud Armor WAF & CIDR policy for action-pptx",
        ],
    )
    cidrs = [c.strip() for c in allowed_cidrs.split(",") if c.strip()]
    policy_desc = run_cmd(
        ["gcloud", "compute", "security-policies", "describe", armor_name, "--format=json"],
        capture_output=True,
        check=False,
    )
    rule_exists = False
    if policy_desc.returncode == 0:
        try:
            policy_data = json.loads(policy_desc.stdout or "{}")
            rule_exists = any(r.get("priority") == 1000 for r in policy_data.get("rules", []))
        except json.JSONDecodeError:
            pass

    if cidrs:
        run_cmd(
            [
                "gcloud",
                "compute",
                "security-policies",
                "rules",
                "update",
                "2147483647",
                f"--security-policy={armor_name}",
                "--action=deny-403",
            ]
        )
        run_cmd(
            [
                "gcloud",
                "compute",
                "security-policies",
                "rules",
                "update" if rule_exists else "create",
                "1000",
                f"--security-policy={armor_name}",
                f"--src-ip-ranges={','.join(cidrs)}",
                "--action=allow",
                "--description=Allowed Looker / corporate CIDRs",
            ]
        )
    else:
        run_cmd(
            [
                "gcloud",
                "compute",
                "security-policies",
                "rules",
                "update",
                "2147483647",
                f"--security-policy={armor_name}",
                "--action=allow",
            ]
        )
        if rule_exists:
            run_cmd(
                [
                    "gcloud",
                    "compute",
                    "security-policies",
                    "rules",
                    "delete",
                    "1000",
                    f"--security-policy={armor_name}",
                    "--quiet",
                ]
            )

    backend_name = "action-pptx-backend"
    if _ensure_gcloud(
        ["compute", "backend-services", "describe", backend_name, "--global"],
        [
            "compute",
            "backend-services",
            "create",
            backend_name,
            "--load-balancing-scheme=EXTERNAL_MANAGED",
            "--global",
        ],
    ):
        run_cmd(
            [
                "gcloud",
                "compute",
                "backend-services",
                "add-backend",
                backend_name,
                "--global",
                "--network-endpoint-group=action-pptx-neg",
                f"--network-endpoint-group-region={region}",
            ]
        )
    run_cmd(
        [
            "gcloud",
            "compute",
            "backend-services",
            "update",
            backend_name,
            "--global",
            f"--security-policy={armor_name}",
        ]
    )

    url_map = "action-pptx-url-map"
    _ensure_gcloud(
        ["compute", "url-maps", "describe", url_map, "--global"],
        ["compute", "url-maps", "create", url_map, f"--default-service={backend_name}", "--global"],
    )

    cert_name = "action-pptx-ssl-cert"
    _ensure_gcloud(
        ["compute", "ssl-certificates", "describe", cert_name, "--global"],
        ["compute", "ssl-certificates", "create", cert_name, f"--domains={lb_domain}", "--global"],
    )

    proxy_name = "action-pptx-https-proxy"
    _ensure_gcloud(
        ["compute", "target-https-proxies", "describe", proxy_name, "--global"],
        [
            "compute",
            "target-https-proxies",
            "create",
            proxy_name,
            f"--url-map={url_map}",
            f"--ssl-certificates={cert_name}",
            "--global",
        ],
    )

    fr_name = "action-pptx-https-fr"
    _ensure_gcloud(
        ["compute", "forwarding-rules", "describe", fr_name, "--global"],
        [
            "compute",
            "forwarding-rules",
            "create",
            fr_name,
            "--load-balancing-scheme=EXTERNAL_MANAGED",
            f"--address={ip_name}",
            f"--target-https-proxy={proxy_name}",
            "--global",
            "--ports=443",
        ],
    )

    console.print(
        f"\n[bold yellow]DNS Action Required:[/bold yellow] Point "
        f"[bold cyan]{lb_domain}[/bold cyan] (A record) to "
        f"[bold green]{lb_ip}[/bold green] so the Google-managed SSL cert provisions."
    )
    return f"https://{lb_domain}", lb_ip


def _provision_internal_lb(
    *,
    region: str,
    vpc_network: str,
    vpc_subnet: str,
    lb_domain: str,
    ssl_cert_name: str,
) -> tuple[str, str]:
    """Provision Regional Internal HTTPS ALB + Serverless NEG. Returns (base_url, internal_ip)."""
    _ensure_serverless_neg(region)
    proxies = run_cmd(
        [
            "gcloud",
            "compute",
            "networks",
            "subnets",
            "list",
            f"--regions={region}",
            f"--filter=purpose=REGIONAL_MANAGED_PROXY AND network~/{vpc_network}$",
            "--format=value(name)",
        ],
        capture_output=True,
    ).stdout.strip()
    if not proxies:
        console.print(
            f"[bold red]Error:[/bold red] VPC network '{vpc_network}' has no "
            f"REGIONAL_MANAGED_PROXY subnet in {region} (required for Internal HTTPS ALB)."
        )
        raise typer.Exit(1)

    backend_name = "action-pptx-internal-backend"
    if _ensure_gcloud(
        ["compute", "backend-services", "describe", backend_name, f"--region={region}"],
        [
            "compute",
            "backend-services",
            "create",
            backend_name,
            "--load-balancing-scheme=INTERNAL_MANAGED",
            f"--region={region}",
            "--protocol=HTTPS",
        ],
    ):
        run_cmd(
            [
                "gcloud",
                "compute",
                "backend-services",
                "add-backend",
                backend_name,
                f"--region={region}",
                "--network-endpoint-group=action-pptx-neg",
                f"--network-endpoint-group-region={region}",
            ]
        )

    url_map = "action-pptx-internal-url-map"
    _ensure_gcloud(
        ["compute", "url-maps", "describe", url_map, f"--region={region}"],
        [
            "compute",
            "url-maps",
            "create",
            url_map,
            f"--default-service={backend_name}",
            f"--region={region}",
        ],
    )

    proxy_name = "action-pptx-internal-https-proxy"
    _ensure_gcloud(
        ["compute", "target-https-proxies", "describe", proxy_name, f"--region={region}"],
        [
            "compute",
            "target-https-proxies",
            "create",
            proxy_name,
            f"--url-map={url_map}",
            f"--ssl-certificates={ssl_cert_name}",
            f"--region={region}",
        ],
    )

    fr_name = "action-pptx-internal-fr"
    _ensure_gcloud(
        ["compute", "forwarding-rules", "describe", fr_name, f"--region={region}"],
        [
            "compute",
            "forwarding-rules",
            "create",
            fr_name,
            "--load-balancing-scheme=INTERNAL_MANAGED",
            f"--network={vpc_network}",
            f"--subnet={vpc_subnet}",
            f"--target-https-proxy={proxy_name}",
            f"--target-https-proxy-region={region}",
            f"--region={region}",
            "--ports=443",
        ],
    )
    internal_ip = run_cmd(
        [
            "gcloud",
            "compute",
            "forwarding-rules",
            "describe",
            fr_name,
            f"--region={region}",
            "--format=value(IPAddress)",
        ],
        capture_output=True,
    ).stdout.strip()
    return f"https://{lb_domain or internal_ip}", internal_ip


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
        help="Automatically register the Action Hub & Extension in Looker",
    ),
    looker_url: str | None = typer.Option(None, "--looker-url", help="Looker Base URL"),
    looker_client_id: str | None = typer.Option(
        None, "--looker-client-id", help="Looker Client ID"
    ),
    looker_client_secret: str | None = typer.Option(
        None, "--looker-client-secret", help="Looker Client Secret"
    ),
    looker_oauth_account: str | None = typer.Option(
        None,
        "--looker-oauth-account",
        help="Configured lkr-dev-cli OAuth account name (alternative to client ID/secret)",
    ),
    extension_mode: str | None = typer.Option(
        None,
        "--extension-mode",
        help=(
            "Extension setup mode: 'new' (deploy to new LookML project), "
            "'existing' (print manifest block), or 'skip'"
        ),
    ),
    looker_project: str | None = typer.Option(
        None,
        "--looker-project",
        help="LookML project ID when deploying the Extension to a new project",
    ),
    looker_model: str | None = typer.Option(
        None,
        "--looker-model",
        help="LookML model name to configure when deploying to a new project",
    ),
    looker_connection: str | None = typer.Option(
        None,
        "--looker-connection",
        help="Looker database connection name to attach to the new project's model",
    ),
    include_sample_lookml: bool | None = typer.Option(
        None,
        "--include-sample-lookml/--no-include-sample-lookml",
        help=(
            "Include sample LookML view and Executive Business Review dashboard from assets/lookml"
        ),
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
    action_label: str = typer.Option(
        "PowerPoint Template Merge (action-pptx)",
        "--action-label",
        envvar="ACTION_LABEL",
        help="Display label shown for the action in Looker Admin & Schedule modal",
    ),
    required_user_attribute: str = typer.Option(
        "",
        "--required-user-attribute",
        envvar="REQUIRED_USER_ATTRIBUTE",
        help="Optional Looker User Attribute name to restrict who can see and execute this action",
    ),
    public_base_url: str | None = typer.Option(
        None,
        "--public-base-url",
        help="Public URL if fronted by a load balancer or proxy (defaults to Cloud Run URL)",
    ),
    looker_only: bool = typer.Option(
        False,
        "--looker-only",
        help="Skip GCP provisioning/deploy and only configure Looker (reads existing secret & URL)",
    ),
    deploy_mode: str | None = typer.Option(
        None,
        "--deploy-mode",
        help=(
            "Exposure mode (both support Actions + Extension): 'load-balanced' (recommended; "
            "blocks direct .a.run.app access) or 'public' (direct Cloud Run URL)"
        ),
    ),
    lb_type: str | None = typer.Option(
        None,
        "--lb-type",
        help=(
            "Load balancer for load-balanced mode: 'external' (Global HTTPS ALB + Cloud Armor), "
            "'internal' (VPC Internal HTTPS ALB), or 'none' (BYO LB)"
        ),
    ),
    lb_domain: str | None = typer.Option(
        None,
        "--lb-domain",
        help="Custom HTTPS domain for Load Balancer SSL cert and Action Hub URL",
    ),
    allowed_cidrs: str = typer.Option(
        "",
        "--allowed-cidrs",
        help="Comma-separated CIDR ranges for Cloud Armor allowlist (e.g. Looker outbound IPs)",
    ),
    vpc_network: str = typer.Option(
        "default",
        "--vpc-network",
        help="VPC network name for internal Load Balancer",
    ),
    vpc_subnet: str = typer.Option(
        "",
        "--vpc-subnet",
        help="VPC subnet name for internal Load Balancer",
    ),
    ssl_cert: str = typer.Option(
        "",
        "--ssl-cert",
        help="Existing regional SSL certificate resource name for internal Load Balancer",
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

    resolved_deploy_mode = (deploy_mode or "").strip().lower()
    resolved_lb_type = (lb_type or "").strip().lower()
    org_report = OrgPolicyReport(None, None, None, None, ())

    if not looker_only:
        if resolved_deploy_mode and resolved_deploy_mode not in VALID_DEPLOY_MODES:
            console.print(
                f"[bold red]Error:[/bold red] --deploy-mode must be one of {VALID_DEPLOY_MODES}."
            )
            raise typer.Exit(1)

        with console.status("[bold cyan]Checking GCP Org Policies for Cloud Run..."):
            org_report = inspect_org_policies(project_id)
        assessments = assess_deploy_modes(org_report)
        _render_deployment_options_overview(org_report, assessments)

        if not resolved_deploy_mode:
            mode_choice = questionary.select(
                "Select Cloud Run exposure mode (both support Looker Actions + Extension):",
                choices=[
                    questionary.Choice(
                        f"{a.mode} [{a.status}] — {a.summary}",
                        value=a.mode,
                    )
                    for a in assessments
                ],
                default="load-balanced",
            ).ask()
            resolved_deploy_mode = str(mode_choice or "load-balanced")

        if resolved_deploy_mode == "load-balanced":
            if resolved_lb_type and resolved_lb_type not in VALID_LB_TYPES:
                console.print(
                    f"[bold red]Error:[/bold red] --lb-type must be one of {VALID_LB_TYPES}."
                )
                raise typer.Exit(1)
            if not resolved_lb_type:
                resolved_lb_type = str(
                    questionary.select(
                        "Select which Load Balancer to provision in front of Cloud Run:",
                        choices=[
                            questionary.Choice(
                                "external — Adds Global IP, Managed SSL, Serverless NEG & "
                                "Cloud Armor WAF/IP allowlist (needs DNS domain)",
                                value="external",
                            ),
                            questionary.Choice(
                                "internal — Adds VPC Internal HTTPS ALB & Serverless NEG "
                                "(no public IP; needs VPC proxy-only subnet + SSL cert)",
                                value="internal",
                            ),
                            questionary.Choice(
                                "none — Lock Cloud Run to Load Balancers only "
                                "(attach your own LB/proxy via --public-base-url)",
                                value="none",
                            ),
                        ],
                        default="external",
                    ).ask()
                    or "external"
                )
            if resolved_lb_type in ("external", "internal") and not lb_domain:
                lb_domain = (
                    questionary.text(
                        "Enter HTTPS domain name for the Load Balancer SSL cert "
                        "(e.g. pptx.example.com):"
                    ).ask()
                    or ""
                ).strip()
                if resolved_lb_type == "external" and not lb_domain:
                    console.print(
                        "[bold red]Error:[/bold red] --lb-domain is required for external "
                        "Google-managed SSL."
                    )
                    raise typer.Exit(1)
            if resolved_lb_type == "external" and not allowed_cidrs:
                allowed_cidrs = (
                    questionary.text(
                        "Cloud Armor allowed IP CIDRs — e.g. Looker outbound/VPN IPs "
                        "(comma-separated, blank allows all):",
                        default="",
                    ).ask()
                    or ""
                ).strip()
            if resolved_lb_type == "internal":
                if not vpc_subnet:
                    vpc_subnet = (
                        questionary.text(
                            "Enter VPC subnet name for the Internal Load Balancer:"
                        ).ask()
                        or ""
                    ).strip()
                if not ssl_cert:
                    ssl_cert = (
                        questionary.text(
                            "Enter existing regional SSL certificate name for Internal ALB:"
                        ).ask()
                        or ""
                    ).strip()
                if not vpc_subnet or not ssl_cert:
                    console.print(
                        "[bold red]Error:[/bold red] --vpc-subnet and --ssl-cert are required "
                        "for internal ALB."
                    )
                    raise typer.Exit(1)

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
        register_looker = looker_only or bool(
            questionary.confirm(
                "Do you want to automatically configure the Action Hub & Extension in Looker?",
                default=True,
            ).ask()
        )

    resolved_ext_mode = (extension_mode or "").strip().lower()
    if register_looker:
        if not looker_oauth_account:
            console.print(
                "\n[bold yellow]Note:[/bold yellow] Looker API credentials are not stored."
            )
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
            if not looker_url or not looker_client_id or not looker_client_secret:
                console.print(
                    "[bold red]Error: Looker URL, Client ID, and Client Secret are "
                    "required.[/bold red]"
                )
                raise typer.Exit(1)

        if resolved_ext_mode not in ("new", "existing", "skip"):
            ext_choice = questionary.select(
                "How would you like to set up the PowerPoint Template Manager Extension?",
                choices=[
                    questionary.Choice(
                        "Deploy Extension bundle to a NEW LookML project "
                        "(configures model & connection)",
                        value="new",
                    ),
                    questionary.Choice(
                        "Add Extension to an EXISTING LookML project "
                        "(prints manifest.lkml application block)",
                        value="existing",
                    ),
                    questionary.Choice(
                        "Skip Extension setup (register Action Hub only)",
                        value="skip",
                    ),
                ],
                default="new",
            ).ask()
            resolved_ext_mode = str(ext_choice or "existing")

        if resolved_ext_mode == "new":
            looker_project = (
                looker_project
                or questionary.text("Enter new LookML Project ID:", default="action-pptx").ask()
                or "action-pptx"
            )
            looker_model = (
                looker_model
                or questionary.text(
                    "Enter LookML Model name to configure:", default="action_pptx_extension"
                ).ask()
                or "action_pptx_extension"
            )
            if not looker_connection:
                try:
                    with console.status("[bold cyan]Fetching Looker database connections..."):
                        connections = list_looker_connections(
                            looker_url,
                            looker_client_id,
                            looker_client_secret,
                            oauth_account=looker_oauth_account,
                        )
                except Exception as exc:
                    console.print(
                        f"[yellow]Warning: Could not fetch Looker connections ({exc}).[/yellow]"
                    )
                    connections = []
                if connections:
                    conn_table = Table(
                        title="Available Looker Database Connections",
                        show_header=True,
                        header_style="bold cyan",
                    )
                    conn_table.add_column("Connection Name", style="green")
                    conn_table.add_column("Dialect", style="magenta")
                    conn_table.add_column("Database / Project", style="white")
                    for cname, cdialect, cdb in connections:
                        conn_table.add_row(cname, cdialect or "-", cdb or "-")
                    console.print("\n", conn_table)

                    looker_connection = questionary.select(
                        f"Select database connection for model '{looker_model}':",
                        choices=[
                            questionary.Choice(
                                f"{cname} ({cdialect})" if cdialect else cname,
                                value=cname,
                            )
                            for cname, cdialect, _ in connections
                        ],
                    ).ask()
                else:
                    looker_connection = (
                        questionary.text("Enter Looker database connection name:").ask() or ""
                    )
                if not looker_connection:
                    console.print("[bold red]Error: A connection must be selected.[/bold red]")
                    raise typer.Exit(1)
            if include_sample_lookml is None:
                include_sample_lookml = bool(
                    questionary.confirm(
                        "Include sample LookML view and Executive Business Review dashboard "
                        "from assets/lookml?",
                        default=True,
                    ).ask()
                )
    else:
        resolved_ext_mode = resolved_ext_mode or "existing"

    repo_root = BIN_DIR.parent
    lb_ip_address = ""
    if looker_only:
        console.print(
            "\n[bold cyan]Skipping GCP deployment; reading existing secret and URL...[/bold cyan]"
        )
        action_hub_secret_val = run_cmd(
            ["gcloud", "secrets", "versions", "access", "latest", "--secret=action-hub-secret"],
            capture_output=True,
        ).stdout.strip()
        effective_url = (
            normalize_service_base(public_base_url)
            if public_base_url
            else run_cmd(
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
        )
    else:
        apis = [
            "run.googleapis.com",
            "secretmanager.googleapis.com",
            "storage.googleapis.com",
            "cloudbuild.googleapis.com",
            "artifactregistry.googleapis.com",
        ]
        if resolved_deploy_mode == "load-balanced" and resolved_lb_type in ("external", "internal"):
            apis.append("compute.googleapis.com")
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

        for sample_name in ("simple_template.pptx", "embed_demo_template.pptx"):
            sample_file = repo_root / "assets" / sample_name
            if sample_file.is_file():
                console.print(
                    "Uploading sample template to "
                    f"[bold cyan]{templates_gcs_uri}/{sample_name}[/bold cyan]..."
                )
                run_cmd(
                    [
                        "gcloud",
                        "storage",
                        "cp",
                        str(sample_file),
                        f"{templates_gcs_uri}/{sample_name}",
                    ],
                    capture_output=True,
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
                [
                    "gcloud",
                    "secrets",
                    "versions",
                    "access",
                    "latest",
                    "--secret=action-hub-secret",
                ],
                capture_output=True,
            ).stdout.strip()
        else:
            action_hub_secret_val = secrets.token_hex(32)
            console.print("Creating secret [bold cyan]action-hub-secret[/bold cyan]...")
            run_cmd(
                [
                    "gcloud",
                    "secrets",
                    "create",
                    "action-hub-secret",
                    "--replication-policy=automatic",
                ]
            )
            run_cmd(
                ["gcloud", "secrets", "versions", "add", "action-hub-secret", "--data-file=-"],
                capture_output=True,
                input_text=action_hub_secret_val,
            )

        if not service_account_email:
            project_number = run_cmd(
                ["gcloud", "projects", "describe", project_id, "--format=value(projectNumber)"],
                capture_output=True,
            ).stdout.strip()
            service_account_email = f"{project_number}-compute@developer.gserviceaccount.com"

        console.print(
            f"Granting runtime roles to [bold cyan]{service_account_email}[/bold cyan]..."
        )
        for base_cmd, role in (
            (
                ["gcloud", "secrets", "add-iam-policy-binding", "action-hub-secret"],
                "roles/secretmanager.secretAccessor",
            ),
            (
                ["gcloud", "storage", "buckets", "add-iam-policy-binding", f"gs://{bucket_name}"],
                "roles/storage.objectAdmin",
            ),
            (
                ["gcloud", "projects", "add-iam-policy-binding", project_id],
                "roles/logging.logWriter",
            ),
        ):
            run_cmd_with_retry(
                [
                    *base_cmd,
                    f"--member=serviceAccount:{service_account_email}",
                    f"--role={role}",
                    "--condition=None",
                ],
                error_substrings=["does not exist", "INVALID_ARGUMENT"],
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
            f"ACTION_LABEL={action_label}",
            "PUBLIC_BASE_URL=http://placeholder",
        ]
        if required_user_attribute.strip():
            env_vars.append(f"REQUIRED_USER_ATTRIBUTE={required_user_attribute.strip()}")
        if oauth_client_id:
            env_vars.extend(
                [
                    f"GOOGLE_OAUTH_CLIENT_ID={oauth_client_id}",
                    f"GOOGLE_OAUTH_CLIENT_SECRET={oauth_client_secret or ''}",
                    "OAUTH_GCS_PASSTHROUGH=true",
                ]
            )

        ingress_flag = (
            "--ingress=internal-and-cloud-load-balancing"
            if resolved_deploy_mode == "load-balanced"
            else "--ingress=all"
        )
        invoker_flags = [] if org_report.require_invoker_iam is True else ["--no-invoker-iam-check"]
        run_cmd(
            [
                "gcloud",
                "run",
                "deploy",
                "action-pptx",
                "--platform=managed",
                f"--region={region}",
                ingress_flag,
                *invoker_flags,
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
        effective_url = normalize_service_base(public_base_url) if public_base_url else service_url
        if resolved_deploy_mode == "load-balanced":
            if resolved_lb_type == "external":
                effective_url, lb_ip_address = _provision_external_lb(
                    region=region,
                    lb_domain=lb_domain or "",
                    allowed_cidrs=allowed_cidrs,
                )
            elif resolved_lb_type == "internal":
                effective_url, lb_ip_address = _provision_internal_lb(
                    region=region,
                    vpc_network=vpc_network,
                    vpc_subnet=vpc_subnet,
                    lb_domain=lb_domain or "",
                    ssl_cert_name=ssl_cert,
                )
        console.print(f"Cloud Run URL: [bold cyan]{service_url}[/bold cyan]")
        console.print(f"Effective Service URL: [bold green]{effective_url}[/bold green]")

    if not looker_only or public_base_url:
        run_cmd(
            [
                "gcloud",
                "run",
                "services",
                "update",
                "action-pptx",
                "--platform=managed",
                f"--region={region}",
                f"--update-env-vars=PUBLIC_BASE_URL={effective_url}",
            ],
            capture_output=True,
        )

    nonce = secrets.token_bytes(32).hex()
    digest = hmac.new(
        action_hub_secret_val.encode("utf-8"), nonce.encode("utf-8"), hashlib.sha512
    ).hexdigest()
    api_key_token = f"{nonce}/{digest}"

    deployed_extension_id: str | None = None
    looker_warnings: list[str] = []
    if register_looker:
        console.print("\n[bold cyan]Configuring Looker Action Hub & Extension...[/bold cyan]")
        try:
            bundle_path = (
                _maybe_build_extension_bundle(repo_root) if resolved_ext_mode == "new" else None
            )
            deployed_extension_id, looker_warnings = register_action_hub_and_extension(
                service_url=effective_url,
                api_key_token=api_key_token,
                templates_gcs_uri=templates_gcs_uri,
                output_gcs_uri=output_gcs_uri,
                required_user_attribute=required_user_attribute,
                extension_mode=resolved_ext_mode,
                project_id=looker_project or "action-pptx",
                model_name=looker_model or "action_pptx_extension",
                connection_name=looker_connection or "",
                bundle_path=bundle_path,
                include_sample_lookml=bool(include_sample_lookml),
                oauth_client_id=oauth_client_id or None,
                oauth_gcs_passthrough=bool(oauth_client_id),
                base_url=looker_url,
                client_id=looker_client_id,
                client_secret=looker_client_secret,
                oauth_account=looker_oauth_account,
            )
        except Exception as exc:
            looker_warnings.append(f"Looker configuration failed: {exc}")

    table = Table(
        title="Looker Action Hub & Extension Details",
        show_header=True,
        header_style="bold magenta",
    )
    table.add_column("Parameter", style="cyan", width=32)
    table.add_column("Value", style="green", overflow="fold")
    if not looker_only and resolved_deploy_mode:
        mode_detail = (
            f"{resolved_deploy_mode} ({resolved_lb_type})"
            if resolved_deploy_mode == "load-balanced" and resolved_lb_type
            else resolved_deploy_mode
        )
        table.add_row("Deploy Mode", mode_detail)
    if lb_ip_address:
        table.add_row("Load Balancer IP", lb_ip_address)
    table.add_row("Action Hub URL", effective_url)
    table.add_row("Authorization Token (API Key)", api_key_token)
    table.add_row("Templates GCS URI", templates_gcs_uri)
    table.add_row("Output GCS Package Bucket URI", output_gcs_uri)
    ext_link: str | None = None
    sample_dash_link: str | None = None
    base_prefix = looker_url.rstrip("/") if looker_url else ""
    if deployed_extension_id:
        ext_link = f"{base_prefix}/extensions/{deployed_extension_id}"
        table.add_row("Deployed Looker Extension", ext_link)
        if include_sample_lookml:
            resolved_model = looker_model or "action_pptx_extension"
            sample_dash_link = (
                f"{base_prefix}/dashboards/{resolved_model}::executive_business_review"
            )
            table.add_row("Sample LookML Dashboard", sample_dash_link)

    panel_title = (
        "[bold yellow]Deployment Completed with Looker Warnings[/bold yellow]"
        if looker_warnings
        else "[bold green]Deployment & Configuration Successful![/bold green]"
    )
    console.print(
        "\n",
        Panel.fit(
            table,
            title=panel_title,
            border_style="yellow" if looker_warnings else "green",
            padding=(1, 2),
        ),
    )
    console.print(f"\n[bold cyan]Copyable API Key:[/bold cyan]\n{api_key_token}", soft_wrap=True)
    if ext_link:
        console.print(f"\n[bold cyan]Extension URL:[/bold cyan] {ext_link}", soft_wrap=True)
    if sample_dash_link:
        console.print(
            f"[bold cyan]Sample Dashboard URL:[/bold cyan] {sample_dash_link}", soft_wrap=True
        )

    if resolved_ext_mode == "existing":
        manifest_snippet = build_manifest_application_block(effective_url, use_local_bundle=False)
        console.print(
            "\n",
            Panel(
                Syntax(manifest_snippet, "ruby", theme="monokai", word_wrap=True),
                title="[bold cyan]Paste into your existing project's manifest.lkml[/bold cyan]",
                border_style="cyan",
                padding=(1, 2),
            ),
        )
        console.print(
            f"\n[bold cyan]Raw copyable manifest.lkml block:[/bold cyan]\n{manifest_snippet}\n"
        )

    if looker_warnings:
        console.print(
            "\n[bold yellow]Looker warnings (re-run anytime with --looker-only):[/bold yellow]"
        )
        for warning in looker_warnings:
            console.print(f"  [yellow]- {warning}[/yellow]")


if __name__ == "__main__":
    if docs_out := os.environ.pop("GENERATE_DEPLOY_CLI_DOCS", ""):
        import typer.cli

        app = typer.Typer()
        app.command()(main)
        cmd = typer.main.get_command(app)
        docs = typer.cli.get_docs_for_click(obj=cmd, ctx=typer.Context(cmd), name="./deploy.sh")
        Path(docs_out).write_text(f"{docs.strip()}\n", encoding="utf-8")
    else:
        typer.run(main)
