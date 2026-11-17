#!/usr/bin/env uv run
# /// script
# requires-python = ">=3.14"
# dependencies = []
# ///

from __future__ import annotations

import hashlib
import hmac
import json
import os
import re
import secrets
import subprocess
import tempfile
import time
import urllib.parse
import urllib.request
from pathlib import Path

BIN_DIR = Path(__file__).resolve().parent
REPO_ROOT = BIN_DIR.parent
CODEMODE_SCRIPT = BIN_DIR / "looker_codemode.py"
CDN_BUNDLE_URL = "https://cdn.lkr.dev/apps/action-pptx/latest/bundle.js"
EXTENSION_APP_ID = "action_pptx_template_manager"


def normalize_service_base(service_url: str) -> str:
    cleaned = service_url.strip().rstrip("/")
    parsed = urllib.parse.urlparse(cleaned)
    if not parsed.scheme or not parsed.netloc:
        raise ValueError(f"Invalid service URL: {service_url!r}")
    path = parsed.path.rstrip("/")
    return f"{parsed.scheme}://{parsed.netloc}{path}"


def build_secret_domain_allowlist(service_url: str) -> str:
    base = normalize_service_base(service_url)
    return f"{base},{base}/,{base}/*"


def build_manifest_application_block(
    service_url: str,
    *,
    use_local_bundle: bool = False,
) -> str:
    base = normalize_service_base(service_url)
    template = (REPO_ROOT / "manifest.lkml").read_text(encoding="utf-8").strip()
    if use_local_bundle:
        template = re.sub(r'^\s*url:\s*"[^"]+"', '  file: "bundle.js"', template, flags=re.M)
    urls_block = (
        f'    external_api_urls: [\n      "{base}",\n      "{base}/",\n      "{base}/*"\n    ]'
    )
    return re.sub(r"    external_api_urls:\s*\[[^\]]*\]", urls_block, template)


def _load_bundle_js(bundle_path: str | None = None) -> str | None:
    if bundle_path:
        candidate = Path(bundle_path)
        if candidate.is_file():
            return candidate.read_text(encoding="utf-8")
    try:
        with urllib.request.urlopen(CDN_BUNDLE_URL, timeout=20) as resp:
            raw: bytes = resp.read()
            text = raw.decode("utf-8")
            if text.strip():
                return text
    except Exception:
        return None
    return None


def _lkr_env(
    base_url: str | None = None,
    client_id: str | None = None,
    client_secret: str | None = None,
) -> dict[str, str]:
    env = os.environ.copy()
    if base_url:
        env["LOOKERSDK_BASE_URL"] = base_url.strip().rstrip("/")
    if client_id:
        env["LOOKERSDK_CLIENT_ID"] = client_id.strip()
    if client_secret:
        env["LOOKERSDK_CLIENT_SECRET"] = client_secret.strip()
    return env


def _lkr_cli_prefix(
    *,
    extras: str = "code-mode",
    oauth_account: str | None = None,
    dev: bool = True,
) -> list[str]:
    cmd = ["uvx", "--from", f"lkr-dev-cli[{extras}]", "lkr-dev-cli"]
    if oauth_account:
        cmd.append(f"--oauth-account={oauth_account.strip()}")
    if dev:
        cmd.append("--dev")
    return cmd


def _run_lkr_codemode(
    *,
    code: str | None = None,
    file_path: Path | None = None,
    variables: dict[str, str] | None = None,
    base_url: str | None = None,
    client_id: str | None = None,
    client_secret: str | None = None,
    oauth_account: str | None = None,
    dev: bool = True,
) -> object:
    cmd = [
        *_lkr_cli_prefix(extras="tools,code-mode", oauth_account=oauth_account, dev=dev),
        "code-mode",
        "sandbox",
    ]
    if code is not None:
        cmd.extend(["--code", code])
    elif file_path is not None:
        cmd.extend(["--file", str(file_path)])
    else:
        raise ValueError("Either code or file_path is required.")

    for key, val in (variables or {}).items():
        cmd.extend(["--var", f"{key}={val}"])

    proc = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        env=_lkr_env(base_url, client_id, client_secret),
    )
    stdout = proc.stdout.strip()
    if proc.returncode != 0 or stdout.startswith("Error:"):
        detail = stdout or proc.stderr.strip() or f"exit code {proc.returncode}"
        raise RuntimeError(f"lkr-dev-cli code-mode failed: {detail}")
    return json.loads(stdout) if stdout else {}


def list_looker_connections(
    base_url: str | None = None,
    client_id: str | None = None,
    client_secret: str | None = None,
    *,
    oauth_account: str | None = None,
) -> list[tuple[str, str, str]]:
    """Return (name, dialect_label, database) for all configured Looker connections."""
    raw = _run_lkr_codemode(
        code="return all_connections(fields='name,dialect,database')",
        base_url=base_url,
        client_id=client_id,
        client_secret=client_secret,
        oauth_account=oauth_account,
        dev=False,
    )
    if not isinstance(raw, list):
        return []
    rows: list[tuple[str, str, str]] = []
    for conn in raw:
        if not isinstance(conn, dict):
            continue
        name = str(conn.get("name") or "").strip()
        if not name:
            continue
        dialect_obj = conn.get("dialect")
        dialect = ""
        if isinstance(dialect_obj, dict) and dialect_obj.get("label"):
            dialect = str(dialect_obj["label"])
        elif conn.get("dialect_name"):
            dialect = str(conn["dialect_name"])
        database = str(conn.get("database") or "").strip()
        rows.append((name, dialect, database))
    rows.sort(key=lambda r: r[0].lower())
    return rows


def stage_lookml_project_dir(
    target_dir: Path,
    *,
    service_url: str,
    project_id: str,
    model_name: str,
    connection_name: str,
    bundle_path: str | None = None,
    include_sample_lookml: bool = False,
) -> bool:
    """Populate target_dir with manifest.lkml, optional bundle.js, model, and sample LookML."""
    target_dir.mkdir(parents=True, exist_ok=True)

    bundle_js = _load_bundle_js(bundle_path)
    use_local_bundle = bundle_js is not None
    if bundle_js is not None:
        (target_dir / "bundle.js").write_text(bundle_js, encoding="utf-8")

    manifest_content = (
        f'project_name: "{project_id}"\n\n'
        + build_manifest_application_block(service_url, use_local_bundle=use_local_bundle)
        + "\n"
    )
    (target_dir / "manifest.lkml").write_text(manifest_content, encoding="utf-8")

    models_dir = target_dir / "models"
    models_dir.mkdir(exist_ok=True)
    sample_root = REPO_ROOT / "assets" / "lookml"
    sample_model = sample_root / "models" / "action_pptx.model.lkml"

    if include_sample_lookml and sample_model.is_file():
        model_text = re.sub(
            r'^\s*connection:\s*"[^"]*"',
            lambda _m: f'connection: "{connection_name}"',
            sample_model.read_text(encoding="utf-8"),
            count=1,
            flags=re.M,
        )
        (models_dir / f"{model_name}.model.lkml").write_text(model_text, encoding="utf-8")

        sample_view = sample_root / "views" / "inventory_items.view.lkml"
        if sample_view.is_file():
            views_dir = target_dir / "views"
            views_dir.mkdir(exist_ok=True)
            (views_dir / sample_view.name).write_text(
                sample_view.read_text(encoding="utf-8"), encoding="utf-8"
            )

        sample_dash = sample_root / "dashboards" / "executive_business_review.dashboard.lookml"
        if sample_dash.is_file():
            dash_dir = target_dir / "dashboards"
            dash_dir.mkdir(exist_ok=True)
            dash_text = re.sub(
                r"^(\s*model:\s*)\S+",
                lambda m: f"{m.group(1)}{model_name}",
                sample_dash.read_text(encoding="utf-8"),
                flags=re.M,
            )
            (dash_dir / sample_dash.name).write_text(dash_text, encoding="utf-8")
    else:
        (models_dir / f"{model_name}.model.lkml").write_text(
            f'connection: "{connection_name}"\n',
            encoding="utf-8",
        )

    return use_local_bundle


def deploy_extension_to_new_project(
    *,
    service_url: str,
    templates_gcs_uri: str = "",
    output_gcs_uri: str = "",
    action_integration_id: str = "",
    project_id: str,
    model_name: str,
    connection_name: str,
    bundle_path: str | None = None,
    include_sample_lookml: bool = False,
    oauth_client_id: str | None = None,
    oauth_gcs_passthrough: bool = False,
    base_url: str | None = None,
    client_id: str | None = None,
    client_secret: str | None = None,
    oauth_account: str | None = None,
) -> str:
    """Stage LookML project + bundle.js and push/deploy via lkr-dev-cli tools lookml push."""
    with tempfile.TemporaryDirectory(prefix="action-pptx-lookml-") as tmp:
        stage_dir = Path(tmp) / project_id
        use_local_bundle = stage_lookml_project_dir(
            stage_dir,
            service_url=service_url,
            project_id=project_id,
            model_name=model_name,
            connection_name=connection_name,
            bundle_path=bundle_path,
            include_sample_lookml=include_sample_lookml,
        )
        bundle_msg = "local bundle.js" if use_local_bundle else "CDN bundle URL"
        sample_msg = " + sample LookML/dashboard" if include_sample_lookml else ""
        print(
            f"Pushing LookML project '{project_id}' ({bundle_msg}{sample_msg}) "
            "and deploying to production via lkr-dev-cli..."
        )
        cmd = [
            *_lkr_cli_prefix(extras="tools,code-mode", oauth_account=oauth_account, dev=True),
            "tools",
            "lookml",
            "push",
            str(stage_dir),
            f"--project-id={project_id}",
            "--create-project",
            "--deploy",
            "--message=Deploy action-pptx Looker Extension",
        ]
        subprocess.run(
            cmd,
            check=True,
            env=_lkr_env(base_url, client_id, client_secret),
        )

    extension_id = f"{project_id}::{EXTENSION_APP_ID}"
    print(f"Configuring Extension context for '{extension_id}' via update_data_app...")
    cfg_res = _run_lkr_codemode(
        file_path=CODEMODE_SCRIPT,
        variables={
            "action": "configure_extension",
            "service_url": normalize_service_base(service_url),
            "templates_gcs_uri": templates_gcs_uri,
            "output_gcs_uri": output_gcs_uri,
            "action_integration_id": (action_integration_id or "").strip(),
            "project_id": project_id,
            "oauth_client_id": (oauth_client_id or "").strip(),
            "oauth_gcs_passthrough": "true" if oauth_gcs_passthrough else "false",
        },
        base_url=base_url,
        client_id=client_id,
        client_secret=client_secret,
        oauth_account=oauth_account,
        dev=True,
    )
    if isinstance(cfg_res, dict):
        if cfg_res.get("configured"):
            print(f"Extension context updated for {extension_id}!")
        elif cfg_res.get("reason"):
            print(f"Note: Could not auto-configure {extension_id}: {cfg_res['reason']}")
    return extension_id


def register_action_hub_and_extension(
    *,
    service_url: str,
    api_key_token: str,
    templates_gcs_uri: str,
    output_gcs_uri: str,
    required_user_attribute: str = "",
    extension_mode: str = "skip",
    project_id: str = "action-pptx",
    model_name: str = "action_pptx_extension",
    connection_name: str = "",
    bundle_path: str | None = None,
    include_sample_lookml: bool = False,
    oauth_client_id: str | None = None,
    oauth_gcs_passthrough: bool = False,
    base_url: str | None = None,
    client_id: str | None = None,
    client_secret: str | None = None,
    oauth_account: str | None = None,
) -> str | None:
    token = api_key_token
    if "/" not in token:
        nonce = secrets.token_bytes(32).hex()
        digest = hmac.new(token.encode("utf-8"), nonce.encode("utf-8"), hashlib.sha512).hexdigest()
        token = f"{nonce}/{digest}"

    url = normalize_service_base(service_url)
    print(f"Registering Integration Hub {url} via lkr-dev-cli code-mode...")
    res: object = {}
    for attempt in range(3):
        try:
            res = _run_lkr_codemode(
                file_path=CODEMODE_SCRIPT,
                variables={
                    "action": "setup_hub",
                    "service_url": url,
                    "api_key_token": token,
                    "templates_gcs_uri": templates_gcs_uri,
                    "output_gcs_uri": output_gcs_uri,
                    "required_user_attribute": required_user_attribute,
                    "extension_mode": extension_mode,
                    "project_id": project_id,
                },
                base_url=base_url,
                client_id=client_id,
                client_secret=client_secret,
                oauth_account=oauth_account,
                dev=True,
            )
            break
        except RuntimeError as exc:
            if "Could not find action-pptx integration" not in str(exc) or attempt == 2:
                raise
            time.sleep(3)
    action_integration_id = ""
    if isinstance(res, dict):
        action_integration_id = str(res.get("integration_id") or "").strip()
        print(
            f"Integration Hub {res.get('hub_action', 'configured')} "
            f"(ID: {res.get('hub_id')}), enabled {res.get('integration_id')}!"
        )
        if required_user_attribute.strip() and not res.get("bound_user_attribute"):
            print(
                f"Note: User attribute '{required_user_attribute.strip()}' does not exist yet in "
                "Looker; create it under Admin -> User Attributes and bind it on the action."
            )

    if extension_mode == "new":
        if not connection_name:
            raise RuntimeError(
                "A Looker database connection name is required to deploy a new Extension project."
            )
        return deploy_extension_to_new_project(
            service_url=url,
            templates_gcs_uri=templates_gcs_uri,
            output_gcs_uri=output_gcs_uri,
            action_integration_id=action_integration_id,
            project_id=project_id,
            model_name=model_name,
            connection_name=connection_name,
            bundle_path=bundle_path,
            include_sample_lookml=include_sample_lookml,
            oauth_client_id=oauth_client_id,
            oauth_gcs_passthrough=oauth_gcs_passthrough,
            base_url=base_url,
            client_id=client_id,
            client_secret=client_secret,
            oauth_account=oauth_account,
        )
    return None


def main() -> None:
    url = os.environ.get("SERVICE_URL", "")
    token = os.environ.get("API_KEY_TOKEN", "")
    if not url or not token:
        raise RuntimeError("SERVICE_URL and API_KEY_TOKEN environment variables are required.")

    oauth_client_id = os.environ.get("OAUTH_CLIENT_ID") or None
    oauth_passthrough_raw = os.environ.get("OAUTH_GCS_PASSTHROUGH", "").lower()
    oauth_gcs_passthrough = (
        oauth_passthrough_raw in ("1", "true", "yes")
        if oauth_passthrough_raw
        else bool(oauth_client_id)
    )

    register_action_hub_and_extension(
        service_url=url,
        api_key_token=token,
        templates_gcs_uri=os.environ.get("TEMPLATES_GCS_URI", ""),
        output_gcs_uri=os.environ.get("OUTPUT_GCS_URI", ""),
        required_user_attribute=os.environ.get("REQUIRED_USER_ATTRIBUTE", ""),
        extension_mode=os.environ.get("EXTENSION_MODE", "skip"),
        project_id=os.environ.get("LOOKER_PROJECT_ID", "action-pptx"),
        model_name=os.environ.get("LOOKER_MODEL_NAME", "action_pptx_extension"),
        connection_name=os.environ.get("LOOKER_CONNECTION_NAME", ""),
        bundle_path=os.environ.get("EXTENSION_BUNDLE_PATH") or None,
        include_sample_lookml=os.environ.get("INCLUDE_SAMPLE_LOOKML", "").lower()
        in ("1", "true", "yes"),
        oauth_client_id=oauth_client_id,
        oauth_gcs_passthrough=oauth_gcs_passthrough,
        oauth_account=os.environ.get("LOOKER_OAUTH_ACCOUNT") or None,
    )


if __name__ == "__main__":
    main()
