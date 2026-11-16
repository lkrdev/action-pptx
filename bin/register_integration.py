#!/usr/bin/env uv run
# /// script
# requires-python = ">=3.14"
# dependencies = [
#     "looker-sdk>=24.0.0",
# ]
# ///

from __future__ import annotations

import hashlib
import hmac
import os
import re
import secrets
import time
import urllib.parse
import urllib.request
from pathlib import Path
from typing import TYPE_CHECKING, cast

if TYPE_CHECKING:
    from looker_sdk.sdk.api40.methods import Looker40SDK

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
    template = (
        (Path(__file__).resolve().parent.parent / "manifest.lkml")
        .read_text(encoding="utf-8")
        .strip()
    )
    if use_local_bundle:
        template = re.sub(r'^\s*url:\s*"[^"]+"', '  file: "bundle.js"', template, flags=re.M)
    urls_block = (
        f'    external_api_urls: [\n      "{base}",\n      "{base}/",\n      "{base}/*"\n    ]'
    )
    return re.sub(r"    external_api_urls:\s*\[[^\]]*\]", urls_block, template)


def _init_sdk(
    base_url: str | None = None,
    client_id: str | None = None,
    client_secret: str | None = None,
) -> Looker40SDK:
    import looker_sdk

    if base_url:
        os.environ["LOOKERSDK_BASE_URL"] = base_url.strip().rstrip("/")
    if client_id:
        os.environ["LOOKERSDK_CLIENT_ID"] = client_id.strip()
    if client_secret:
        os.environ["LOOKERSDK_CLIENT_SECRET"] = client_secret.strip()
    return looker_sdk.init40()


def list_looker_connections(
    base_url: str,
    client_id: str,
    client_secret: str,
) -> list[tuple[str, str, str]]:
    """Return (name, dialect_label, database) for all configured Looker connections."""
    sdk = _init_sdk(base_url, client_id, client_secret)
    rows: list[tuple[str, str, str]] = []
    for conn in sdk.all_connections():
        name = (conn.name or "").strip()
        if not name:
            continue
        dialect = ""
        if conn.dialect and conn.dialect.label:
            dialect = conn.dialect.label
        elif conn.dialect_name:
            dialect = conn.dialect_name
        database = (conn.database or "").strip()
        rows.append((name, dialect, database))
    rows.sort(key=lambda r: r[0].lower())
    return rows


def _upsert_project_file(
    sdk: Looker40SDK,
    project_id: str,
    existing_paths: set[str],
    file_path: str,
    content: str,
) -> None:
    enc_proj = sdk.encode_path_param(project_id)
    api_path = f"/projects/{enc_proj}/files"
    payload: dict[str, str] = {"path": file_path, "content": content}
    if file_path in existing_paths:
        sdk.put(path=api_path, structure=dict, body=payload)
    else:
        sdk.post(path=api_path, structure=dict, body=payload)
        existing_paths.add(file_path)


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


def _secret_user_attribute_name(project_id: str) -> str:
    prefix = re.sub(r"[^A-Za-z0-9]+", "_", project_id).strip("_").lower()
    return f"{prefix}_{EXTENSION_APP_ID}_action_hub_secret"


def deploy_extension_to_new_project(
    sdk: Looker40SDK,
    *,
    service_url: str,
    api_key_token: str,
    project_id: str,
    model_name: str,
    connection_name: str,
    bundle_path: str | None = None,
) -> str:
    """Create/configure a LookML project, upload bundle + manifest + model, and deploy."""
    from looker_sdk import models40

    print(f"Preparing LookML project '{project_id}' in dev workspace...")
    sdk.update_session(body=models40.WriteApiSession(workspace_id="dev"))

    try:
        proj = sdk.project(project_id=project_id, fields="id,name,uses_git")
    except Exception:
        proj = sdk.create_project(body=models40.WriteProject(name=project_id))
        print(f"Created LookML project '{project_id}'.")

    if not proj.uses_git:
        proj = sdk.update_project(
            project_id=project_id,
            body=models40.WriteProject(git_service_name="bare"),
        )
        print(f"Initialized bare Git repository on project '{project_id}'.")
    else:
        try:
            sdk.reset_project_to_production(project_id=project_id)
        except Exception:
            pass

    enc_proj = sdk.encode_path_param(project_id)
    raw_files = cast(
        list[dict[str, object]],
        sdk.get(path=f"/projects/{enc_proj}/files", structure=list) or [],
    )
    existing_paths = {
        str(f.get("path") or f.get("id") or "")
        for f in raw_files
        if isinstance(f, dict) and (f.get("path") or f.get("id"))
    }

    bundle_js = _load_bundle_js(bundle_path)
    use_local_bundle = bundle_js is not None
    if bundle_js is not None:
        print(f"Uploading extension bundle.js ({len(bundle_js):,} chars) to '{project_id}'...")
        _upsert_project_file(sdk, project_id, existing_paths, "bundle.js", bundle_js)
    else:
        print("Local/CDN bundle download unavailable; falling back to CDN url in manifest.lkml.")

    manifest_content = (
        f'project_name: "{project_id}"\n\n'
        + build_manifest_application_block(service_url, use_local_bundle=use_local_bundle)
        + "\n"
    )
    _upsert_project_file(sdk, project_id, existing_paths, "manifest.lkml", manifest_content)

    model_file_path = f"{model_name}.model.lkml"
    if model_file_path not in existing_paths:
        model_lkml = f'connection: "{connection_name}"\n'
        _upsert_project_file(sdk, project_id, existing_paths, model_file_path, model_lkml)

    print(f"Configuring LookML model '{model_name}' with connection '{connection_name}'...")
    model_body = models40.WriteLookmlModel(
        name=model_name,
        project_name=project_id,
        allowed_db_connection_names=[connection_name],
    )
    try:
        sdk.lookml_model(lookml_model_name=model_name, fields="name")
        sdk.update_lookml_model(lookml_model_name=model_name, body=model_body)
    except Exception:
        sdk.create_lookml_model(body=model_body)

    enc_proj = sdk.encode_path_param(project_id)
    sdk.post(
        path=f"/projects/{enc_proj}/commit",
        structure=str,
        body={"message": "Deploy action-pptx Looker Extension"},
    )
    sdk.deploy_to_production(project_id=project_id)
    print(f"Deployed LookML project '{project_id}' to production!")

    secret_attr_name = _secret_user_attribute_name(project_id)
    allowlist = build_secret_domain_allowlist(service_url)
    existing_attr = next(
        (a for a in sdk.all_user_attributes() if a.name == secret_attr_name and a.id),
        None,
    )
    if existing_attr and existing_attr.id:
        sdk.delete_user_attribute(user_attribute_id=existing_attr.id)
    sdk.create_user_attribute(
        body=models40.WriteUserAttribute(
            name=secret_attr_name,
            label=f"{project_id}::{EXTENSION_APP_ID} Action Hub Secret",
            type="string",
            value_is_hidden=True,
            hidden_value_domain_whitelist=allowlist,
            user_can_edit=False,
            user_can_view=True,
            default_value=api_key_token,
        )
    )
    print(f"Configured hidden secret user attribute '{secret_attr_name}' (allowlist: {allowlist}).")
    return f"{project_id}::{EXTENSION_APP_ID}"


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
    base_url: str | None = None,
    client_id: str | None = None,
    client_secret: str | None = None,
) -> str | None:
    from looker_sdk import models40

    sdk = _init_sdk(base_url, client_id, client_secret)
    token = api_key_token
    if "/" not in token:
        nonce = secrets.token_bytes(32).hex()
        digest = hmac.new(token.encode("utf-8"), nonce.encode("utf-8"), hashlib.sha512).hexdigest()
        token = f"{nonce}/{digest}"

    url = normalize_service_base(service_url)
    print(f"Registering Integration Hub {url} in Looker...")
    body = models40.WriteIntegrationHub(url=url, authorization_token=token)
    try:
        hub = sdk.create_integration_hub(body=body)
        print(f"Successfully registered Integration Hub (ID: {hub.id})")
    except Exception as exc:
        existing = next(
            (h for h in sdk.all_integration_hubs() if h.url and h.url.rstrip("/") == url and h.id),
            None,
        )
        if not existing or not existing.id:
            raise RuntimeError(f"Failed to find or create the Integration Hub: {exc}") from exc
        hub = sdk.update_integration_hub(integration_hub_id=existing.id, body=body)
        print(f"Updated existing Integration Hub (ID: {hub.id})")

    if hub.id and hub.legal_agreement_required and not hub.legal_agreement_signed:
        hub = sdk.accept_integration_hub_legal_agreement(integration_hub_id=hub.id)

    print("Enabling action-pptx integration...")
    integration = None
    for attempt in range(6):
        try:
            integration = sdk.integration(integration_id=f"{hub.id}::action-pptx")
            break
        except Exception:
            if attempt == 5:
                break
            print(f"Integration not found on hub {hub.id} yet; retrying in 3s ({attempt + 1}/6)...")
            time.sleep(3)

    if not integration or not integration.id:
        raise RuntimeError("Could not find action-pptx integration on the registered hub.")

    params = [
        models40.IntegrationParam(name=k, value=v)
        for k, v in (
            ("templates_gcs_uri", templates_gcs_uri),
            ("output_gcs_uri", output_gcs_uri),
        )
        if v
    ]
    req_attr = required_user_attribute.strip()
    if req_attr:
        has_attr = any(a.name == req_attr for a in sdk.all_user_attributes())
        if has_attr:
            params.append(
                models40.IntegrationParam(
                    name="action_access_param",
                    user_attribute_name=req_attr,
                )
            )
        else:
            print(
                f"Note: User attribute '{req_attr}' does not exist yet in Looker; "
                "create it under Admin -> User Attributes and bind it on the action."
            )

    sdk.update_integration(
        integration_id=integration.id,
        body=models40.WriteIntegration(enabled=True, params=params or None),
    )
    print("Successfully enabled and configured action-pptx integration!")

    if extension_mode == "new":
        if not connection_name:
            raise RuntimeError(
                "A Looker database connection name is required to deploy a new Extension project."
            )
        return deploy_extension_to_new_project(
            sdk,
            service_url=url,
            api_key_token=token,
            project_id=project_id,
            model_name=model_name,
            connection_name=connection_name,
            bundle_path=bundle_path,
        )
    return None


def main() -> None:
    url = os.environ.get("SERVICE_URL", "")
    token = os.environ.get("API_KEY_TOKEN", "")
    if not url or not token:
        raise RuntimeError("SERVICE_URL and API_KEY_TOKEN environment variables are required.")

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
    )


if __name__ == "__main__":
    main()
