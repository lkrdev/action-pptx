# ruff: noqa: F821
from __future__ import annotations

import re

EXTENSION_APP_ID = "action_pptx_template_manager"


def _opt_str(val: object, default: str = "") -> str:
    return str(val) if val is not None else default


try:
    _RAW_SERVICE_URL = service_url
except NameError:
    _RAW_SERVICE_URL = ""

try:
    _RAW_API_KEY_TOKEN = api_key_token
except NameError:
    _RAW_API_KEY_TOKEN = ""

try:
    _RAW_TEMPLATES_GCS_URI = templates_gcs_uri
except NameError:
    _RAW_TEMPLATES_GCS_URI = ""

try:
    _RAW_OUTPUT_GCS_URI = output_gcs_uri
except NameError:
    _RAW_OUTPUT_GCS_URI = ""

try:
    _RAW_REQUIRED_USER_ATTR = required_user_attribute
except NameError:
    _RAW_REQUIRED_USER_ATTR = ""

try:
    _RAW_EXTENSION_MODE = extension_mode
except NameError:
    _RAW_EXTENSION_MODE = "skip"

try:
    _RAW_PROJECT_ID = project_id
except NameError:
    _RAW_PROJECT_ID = "action-pptx"


try:
    _RAW_ACTION = action
except NameError:
    _RAW_ACTION = "setup_hub"

try:
    _RAW_OAUTH_CLIENT_ID = oauth_client_id
except NameError:
    _RAW_OAUTH_CLIENT_ID = ""

try:
    _RAW_OAUTH_PASSTHROUGH = oauth_gcs_passthrough
except NameError:
    _RAW_OAUTH_PASSTHROUGH = ""

try:
    _RAW_ACTION_INTEGRATION_ID = action_integration_id
except NameError:
    _RAW_ACTION_INTEGRATION_ID = ""


def _normalize_base(url: str) -> str:
    cleaned = url.strip().rstrip("/")
    if not (cleaned.startswith("https://") or cleaned.startswith("http://")):
        raise ValueError(f"Invalid service URL: {url!r}")
    return cleaned


def _secret_user_attribute_name(proj: str) -> str:
    prefix = re.sub(r"[^A-Za-z0-9]+", "_", proj).strip("_").lower()
    return f"{prefix}_{EXTENSION_APP_ID}_action_hub_secret"


def _ensure_new_project(proj: str) -> None:
    try:
        p = project(project_id=proj)
    except Exception:
        p = create_project(body={"name": proj})
    if p.get("uses_git") is not True and p.get("git_service_name") != "bare":
        update_project(project_id=proj, body={"git_service_name": "bare"})
    try:
        all_project_files(project_id=proj)
    except Exception:
        create_developer_copy(project_id=proj)


def _configure_extension(
    base: str,
    templates_gcs_uri_val: str,
    output_gcs_uri_val: str,
    project_id_val: str,
    oauth_client_id_val: str,
    oauth_passthrough: bool,
    action_integration_id_val: str = "",
) -> dict[str, object]:
    app_id = f"{project_id_val}::{EXTENSION_APP_ID}"
    context_data: dict[str, object] = {
        "cloudRunUrl": base,
        "templatesGcsUri": templates_gcs_uri_val,
        "useOAuth": bool(oauth_client_id_val),
        "oauthGcsPassthrough": oauth_passthrough if oauth_client_id_val else True,
    }
    if output_gcs_uri_val:
        context_data["outputGcsUri"] = output_gcs_uri_val
    if action_integration_id_val:
        context_data["actionIntegrationId"] = action_integration_id_val
    if oauth_client_id_val:
        context_data["oauthClientId"] = oauth_client_id_val

    try:
        updated = update_data_app(data_app_id=app_id, body={"context": context_data})
    except NameError:
        return {
            "extension_id": app_id,
            "configured": False,
            "reason": "update_data_app requires lkr-dev-cli >= 0.0.68",
        }
    return {
        "extension_id": app_id,
        "configured": True,
        "context": updated.get("context") if isinstance(updated, dict) else context_data,
    }


def main() -> dict[str, object]:
    action_val = _opt_str(_RAW_ACTION, "setup_hub").strip().lower() or "setup_hub"
    service_url_val = _opt_str(_RAW_SERVICE_URL)
    if not service_url_val:
        raise ValueError("service_url variable is required.")

    base = _normalize_base(service_url_val)
    templates_gcs_uri_val = _opt_str(_RAW_TEMPLATES_GCS_URI).strip()
    output_gcs_uri_val = _opt_str(_RAW_OUTPUT_GCS_URI).strip()
    project_id_val = _opt_str(_RAW_PROJECT_ID, "action-pptx").strip() or "action-pptx"

    if action_val == "configure_extension":
        oauth_client_id_val = _opt_str(_RAW_OAUTH_CLIENT_ID).strip()
        oauth_passthrough = _opt_str(_RAW_OAUTH_PASSTHROUGH).strip().lower() in (
            "1",
            "true",
            "yes",
        )
        action_integration_id_val = _opt_str(_RAW_ACTION_INTEGRATION_ID).strip()
        return _configure_extension(
            base,
            templates_gcs_uri_val,
            output_gcs_uri_val,
            project_id_val,
            oauth_client_id_val,
            oauth_passthrough,
            action_integration_id_val,
        )

    api_key_token_val = _opt_str(_RAW_API_KEY_TOKEN)
    if not api_key_token_val:
        raise ValueError("api_key_token variable is required.")

    required_user_attribute_val = _opt_str(_RAW_REQUIRED_USER_ATTR).strip()
    extension_mode_val = _opt_str(_RAW_EXTENSION_MODE, "skip").strip().lower() or "skip"

    hub_body = {"url": base, "authorization_token": api_key_token_val}
    try:
        hub = create_integration_hub(body=hub_body)
        hub_action = "created"
    except Exception as exc:
        hubs = all_integration_hubs() or []
        existing = None
        for h in hubs:
            if (h.get("url") or "").rstrip("/") == base and h.get("id"):
                existing = h
                break
        if not existing:
            raise RuntimeError(f"Failed to find or create the Integration Hub: {exc}") from exc
        hub = update_integration_hub(
            integration_hub_id=str(existing["id"]),
            body=hub_body,
        )
        hub_action = "updated"

    hub_id = str(hub["id"])
    if hub.get("legal_agreement_required") and not hub.get("legal_agreement_signed"):
        hub = accept_integration_hub_legal_agreement(integration_hub_id=hub_id)

    integration_id = f"{hub_id}::action-pptx"
    try:
        integ = integration(integration_id=integration_id)
    except Exception:
        integ = None
    if not integ or not integ.get("id"):
        raise RuntimeError(f"Could not find action-pptx integration on hub {hub_id}.")

    params: list[dict[str, str]] = [
        {"name": k, "value": v}
        for k, v in (
            ("templates_gcs_uri", templates_gcs_uri_val),
            ("output_gcs_uri", output_gcs_uri_val),
        )
        if v
    ]
    bound_user_attribute: str | None = None
    if required_user_attribute_val:
        attrs = all_user_attributes() or []
        if any(a.get("name") == required_user_attribute_val for a in attrs):
            params.append(
                {
                    "name": "action_access_param",
                    "user_attribute_name": required_user_attribute_val,
                }
            )
            bound_user_attribute = required_user_attribute_val

    update_integration(
        integration_id=str(integ["id"]),
        body={"enabled": True, "params": params or None},
    )

    secret_attr_name: str | None = None
    extension_id: str | None = None
    if extension_mode_val == "new":
        secret_attr_name = _secret_user_attribute_name(project_id_val)
        allowlist = f"{base},{base}/,{base}/*"
        existing_attr = None
        for a in all_user_attributes() or []:
            if a.get("name") == secret_attr_name and a.get("id"):
                existing_attr = a
                break
        if existing_attr:
            delete_user_attribute(user_attribute_id=str(existing_attr["id"]))

        create_user_attribute(
            body={
                "name": secret_attr_name,
                "label": f"{project_id_val}::{EXTENSION_APP_ID} Action Hub Secret",
                "type": "string",
                "value_is_hidden": True,
                "hidden_value_domain_whitelist": allowlist,
                "user_can_edit": False,
                "user_can_view": True,
                "default_value": api_key_token_val,
            }
        )
        _ensure_new_project(project_id_val)
        extension_id = f"{project_id_val}::{EXTENSION_APP_ID}"

    return {
        "hub_id": hub_id,
        "hub_action": hub_action,
        "integration_id": str(integ["id"]),
        "integration_enabled": True,
        "bound_user_attribute": bound_user_attribute,
        "extension_id": extension_id,
        "secret_user_attribute": secret_attr_name,
    }


main()
