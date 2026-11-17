from __future__ import annotations

import base64
import binascii
import io
import json
import os
import secrets
import urllib.request
import zipfile
from datetime import UTC, datetime

from fastapi import FastAPI, Header, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, Response
from pydantic import BaseModel, Field

from action_pptx import config, gcs
from action_pptx.config import (
    Settings,
    build_google_oauth_url,
    cipher_id,
    decrypt_oauth_state,
    encrypt_oauth_state,
    load_settings,
    log_structured,
    parse_and_validate_gcs_uri,
    resolve_action_oauth_token,
    verify_action_token,
    verify_oauth_bearer,
)
from pptx_merge.engine import MergeError, merge_pptx
from pptx_merge.inspector import TileSchema, inspect_and_validate
from pptx_merge.themes import PPTX_THEMES, resolve_pptx_theme, theme_from_looker_palette

ACTION_NAME = "action-pptx"
ACTION_ICON_SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 96" fill="none">'
    "<defs>"
    '<linearGradient id="a" x1="22" y1="10" x2="88" y2="68" gradientUnits="userSpaceOnUse">'
    '<stop offset="0" stop-color="#FF8A65"/><stop offset="1" stop-color="#D84315"/>'
    "</linearGradient>"
    '<linearGradient id="b" x1="18" y1="18" x2="86" y2="78" gradientUnits="userSpaceOnUse">'
    '<stop offset="0" stop-color="#FFF"/><stop offset="1" stop-color="#F1F5F9"/>'
    "</linearGradient>"
    '<linearGradient id="c" x1="8" y1="36" x2="46" y2="78" gradientUnits="userSpaceOnUse">'
    '<stop offset="0" stop-color="#FF6D42"/><stop offset=".55" stop-color="#E64A19"/>'
    '<stop offset="1" stop-color="#B91C1C"/>'
    "</linearGradient>"
    '<linearGradient id="d" x1="0" y1="34" x2="0" y2="66" gradientUnits="userSpaceOnUse">'
    '<stop offset="0" stop-color="#38BDF8"/><stop offset="1" stop-color="#0284C7"/>'
    "</linearGradient>"
    '<linearGradient id="e" x1="0" y1="28" x2="0" y2="66" gradientUnits="userSpaceOnUse">'
    '<stop offset="0" stop-color="#818CF8"/><stop offset="1" stop-color="#4F46E5"/>'
    "</linearGradient>"
    "</defs>"
    '<rect x="26" y="11" width="62" height="46" rx="8" transform="rotate(5 57 34)" '
    'fill="url(#a)" opacity=".9"/>'
    '<rect x="20" y="19" width="68" height="52" rx="8" fill="url(#b)" '
    'stroke="#CBD5E1" stroke-width="2"/>'
    '<rect x="28" y="27" width="26" height="4.5" rx="2.25" fill="#94A3B8"/>'
    '<circle cx="77" cy="29.25" r="2.5" fill="#FF6D42"/>'
    '<rect x="50" y="49" width="7" height="14" rx="2" fill="#CBD5E1"/>'
    '<rect x="61" y="41" width="7" height="22" rx="2" fill="url(#d)"/>'
    '<rect x="72" y="34" width="7" height="29" rx="2" fill="url(#e)"/>'
    '<path d="M51 44L63 36L76 28" stroke="#F97316" stroke-width="2.6" '
    'stroke-linecap="round" stroke-linejoin="round"/>'
    '<rect x="8" y="38" width="38" height="38" rx="9" fill="url(#c)"/>'
    '<rect x="9" y="39" width="36" height="36" rx="8" stroke="#FFF" '
    'stroke-opacity=".24" stroke-width="1.5"/>'
    '<path d="M20 66V48h7.8c4.4 0 7.2 2.4 7.2 6.1 0 3.8-2.8 6.2-7.2 6.2H24.4V66H20Z'
    'm4.4-11.2h2.9c2 0 3.2-1 3.2-2.7 0-1.6-1.2-2.6-3.2-2.6h-2.9v5.3Z" fill="#FFF"/>'
    "</svg>"
)
ACTION_ICON_DATA_URI = (
    f"data:image/svg+xml;base64,{base64.b64encode(ACTION_ICON_SVG.encode('utf-8')).decode('ascii')}"
)


class ScheduledPlanPayload(BaseModel):
    scheduled_plan_id: int | str | None = None
    title: str | None = None
    type: str | None = None
    url: str | None = None
    query_id: int | str | None = None


class AttachmentPayload(BaseModel):
    data: str | None = None
    mimetype: str | None = None
    extension: str | None = None


class ActionWebhookRequest(BaseModel):
    type: str | None = "dashboard"
    data: dict[str, str | int | float | bool | None] = Field(default_factory=dict)
    form_params: dict[str, str | int | float | bool | None] = Field(default_factory=dict)
    scheduled_plan: ScheduledPlanPayload | None = None
    attachment: AttachmentPayload | None = None
    state_url: str | None = None
    state_redir_url: str | None = None
    state: str | None = None


class TemplateValidateRequest(BaseModel):
    filename: str
    pptx_base64: str
    dashboard_tiles: list[TileSchema] | None = None


class TemplateUploadRequest(TemplateValidateRequest):
    templates_gcs_uri: str


class TemplatePreviewRequest(BaseModel):
    pptx_base64: str
    tile_csvs: dict[str, str]
    plan_title: str = "Dashboard Preview"
    plan_url: str = ""
    theme: str | None = None
    theme_palette: dict[str, object] | None = None


class OutputListRequest(BaseModel):
    output_gcs_uri: str


class OutputDownloadRequest(BaseModel):
    output_gcs_uri: str
    filename: str


class TemplateListRequest(BaseModel):
    templates_gcs_uri: str


class TemplateDownloadRequest(BaseModel):
    templates_gcs_uri: str
    filename: str


def _resolve_base_url(request: Request, settings: Settings) -> str:
    if os.environ.get("PUBLIC_BASE_URL"):
        return settings.public_base_url
    host = request.headers.get("x-forwarded-host") or request.headers.get("host") or ""
    if host and not host.startswith(("localhost", "127.0.0.1", "test")):
        proto = request.headers.get("x-forwarded-proto") or request.url.scheme or "https"
        return f"{proto}://{host}".rstrip("/")
    return settings.public_base_url


def _action_definition(base_url: str, settings: Settings) -> dict[str, object]:
    base = base_url.rstrip("/")
    uses_oauth = bool(settings.google_oauth_client_id)
    return {
        "name": ACTION_NAME,
        "label": settings.action_label if settings else "PowerPoint Template Merge",
        "icon_data_uri": ACTION_ICON_DATA_URI,
        "description": (
            "Populate native, editable PowerPoint (.pptx) templates from LookML Dashboards."
        ),
        "supported_action_types": ["dashboard"],
        "supported_formats": ["csv_zip"],
        "supported_formattings": ["unformatted"],
        "supported_visualization_formattings": ["apply"],
        "supported_download_settings": ["push"],
        "uses_oauth": uses_oauth,
        "delegate_oauth": False,
        "oauth_client_id": settings.google_oauth_client_id,
        "required_fields": [],
        "params": [
            {
                "name": "templates_gcs_uri",
                "label": "Templates GCS URI (gs://bucket/prefix)",
                "required": True,
                "sensitive": False,
                "description": "GCS bucket/prefix storing validated .pptx templates.",
            },
            {
                "name": "output_gcs_uri",
                "label": "Output GCS Package Bucket URI (gs://bucket/prefix)",
                "required": True,
                "sensitive": False,
                "description": (
                    "GCS bucket/prefix where generated timestamped .pptx files are written."
                ),
            },
            *(
                [
                    {
                        "name": "action_access_param",
                        "label": "Action Access",
                        "description": "Controls which users can access this action.",
                        "user_attribute_name": settings.required_user_attribute,
                        "required": True,
                        "sensitive": True,
                    }
                ]
                if settings and settings.required_user_attribute
                else []
            ),
        ],
        "url": f"{base}/actions/pptx/execute",
        "form_url": f"{base}/actions/pptx/form",
        "token_url": f"{base}/actions/pptx/oauth_token" if uses_oauth else "",
    }


def _looker_response(
    webhook_id: str,
    *,
    success: bool,
    message: str,
    validation_errors: dict[str, str] | None = None,
    state_data: str | None = None,
) -> dict[str, object]:
    looker_obj: dict[str, object] = {
        "success": success,
        "refresh_query": False,
        "message": message,
        "validation_errors": validation_errors or {},
        "webhookId": webhook_id,
    }
    resp: dict[str, object] = {"looker": looker_obj}
    if state_data is not None:
        looker_obj["state"] = {"data": state_data}
        resp["state"] = {"data": state_data}
    return resp


def _login_form(
    base_url: str,
    payload: ActionWebhookRequest,
    settings: Settings,
    error_message: str | None = None,
) -> dict[str, object]:
    state_redir_url = str(
        payload.state_redir_url or payload.data.get("state_redir_url") or ""
    ).strip()
    state_url = str(payload.state_url or payload.data.get("state_url") or "").strip()
    state_obj: dict[str, object] = (
        {"tokenurl": state_redir_url} if state_redir_url else {"stateurl": state_url}
    )
    encrypted_state = encrypt_oauth_state(state_obj, settings.cipher_master)

    fields: list[dict[str, object]] = []
    if error_message:
        fields.append(
            {
                "name": "error_message",
                "type": "message",
                "value": f"⚠️ Login failed: {error_message}",
            }
        )
    fields.append(
        {
            "name": "login",
            "type": "oauth_link_google",
            "label": "Log in with Google",
            "description": (
                "Log in once with your Google account to list and generate PowerPoint (.pptx) "
                "presentations using your Google Cloud Storage credentials (includes offline "
                "refresh token for scheduled deliveries)."
            ),
            "oauth_url": f"{base_url.rstrip('/')}/actions/pptx/oauth?state={encrypted_state}",
        }
    )
    resp: dict[str, object] = {"fields": fields}
    if error_message:
        resp["state"] = {"data": "reset"}
    return resp


def _require_action_auth(authorization: str | None, settings: Settings) -> None:
    if not verify_action_token(authorization, settings.action_hub_secret):
        raise HTTPException(status_code=403, detail="Invalid 'Authorization' header.")


def _require_extension_auth(
    authorization: str | None,
    settings: Settings,
) -> tuple[str, str | None]:
    if verify_action_token(authorization, settings.action_hub_secret):
        return "token", None
    if authorization and authorization.strip().lower().startswith("bearer "):
        try:
            _, token = verify_oauth_bearer(authorization, settings.allowed_oauth_domains)
            return "oauth", token
        except PermissionError as exc:
            raise HTTPException(status_code=403, detail=str(exc)) from exc
    raise HTTPException(status_code=403, detail="Invalid 'Authorization' header.")


def _require_viewer_auth(
    authorization: str | None,
    settings: Settings,
) -> str | None:
    if settings.disable_report_viewer:
        raise HTTPException(
            status_code=403, detail="Report viewer is disabled by server configuration."
        )
    auth_mode, oauth_token = _require_extension_auth(authorization, settings)
    return oauth_token if (auth_mode == "oauth" and settings.oauth_gcs_passthrough) else None


def _decode_base64_payload(b64_str: str, label: str) -> bytes:
    cleaned = "".join(b64_str.split())
    if not cleaned:
        raise ValueError(f"Invalid base64 encoding for {label}.")
    try:
        return base64.b64decode(cleaned, validate=True)
    except (binascii.Error, ValueError) as exc:
        raise ValueError(f"Invalid base64 encoding for {label}.") from exc


def create_app() -> FastAPI:
    app = FastAPI(title="Looker PowerPoint Custom Action (action-pptx)", version="0.1.0")
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_methods=["GET", "POST", "OPTIONS"],
        allow_headers=[
            "Authorization",
            "Content-Type",
            "X-Looker-Webhook-Id",
            "X-Looker-Instance",
        ],
    )

    @app.get("/healthz")
    @app.get("/health")
    async def healthz() -> dict[str, object]:
        settings = load_settings()
        return {
            "status": "ok",
            "configured": bool(settings.action_hub_secret),
            "uses_oauth": bool(settings.google_oauth_client_id),
            "oauth_client_id": settings.google_oauth_client_id,
            "oauth_gcs_passthrough": settings.oauth_gcs_passthrough,
            "report_viewer_enabled": not settings.disable_report_viewer,
            "looker": {"success": True, "message": "Health check passed."},
        }

    @app.post("/")
    async def list_integrations(
        request: Request,
        authorization: str | None = Header(default=None),
    ) -> dict[str, object]:
        settings = load_settings()
        _require_action_auth(authorization, settings)
        base = _resolve_base_url(request, settings)
        return {
            "label": settings.action_hub_label,
            "integrations": [_action_definition(base, settings)],
        }

    @app.post("/actions/pptx")
    async def get_integration(
        request: Request,
        authorization: str | None = Header(default=None),
    ) -> dict[str, object]:
        settings = load_settings()
        _require_action_auth(authorization, settings)
        return _action_definition(_resolve_base_url(request, settings), settings)

    @app.get("/actions/pptx/oauth")
    async def action_oauth_start(
        request: Request,
        state: str = "",
    ) -> Response:
        if not state:
            raise HTTPException(status_code=400, detail="Missing state parameter.")
        settings = load_settings()
        if not settings.google_oauth_client_id:
            raise HTTPException(status_code=400, detail="GOOGLE_OAUTH_CLIENT_ID is not configured.")
        base = _resolve_base_url(request, settings)
        redirect_uri = f"{base}/actions/pptx/oauth_redirect"
        nonce = secrets.token_hex(16)
        combined_state = f"{nonce}.{state}"
        auth_url = build_google_oauth_url(
            settings.google_oauth_client_id,
            redirect_uri,
            combined_state,
            passthrough=settings.oauth_gcs_passthrough,
        )
        response = RedirectResponse(url=auth_url, status_code=302)
        response.set_cookie(
            key="action_hub_state",
            value=nonce,
            httponly=True,
            secure=base.startswith("https://"),
            samesite="lax",
            max_age=3600,
        )
        return response

    @app.get("/actions/pptx/oauth_redirect")
    async def action_oauth_redirect(
        request: Request,
        code: str = "",
        state: str = "",
    ) -> Response:
        if not state or not code:
            raise HTTPException(status_code=400, detail="Missing code or state parameter.")
        settings = load_settings()
        base = _resolve_base_url(request, settings)
        redirect_uri = f"{base}/actions/pptx/oauth_redirect"

        cookie_nonce = request.cookies.get("action_hub_state")
        nonce_part, original_state = state.split(".", 1) if "." in state else ("", state)
        if nonce_part and cookie_nonce and nonce_part != cookie_nonce:
            raise HTTPException(status_code=403, detail="CSRF state nonce mismatch.")

        try:
            state_payload = decrypt_oauth_state(original_state, settings.cipher_master)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

        token_url = str(state_payload.get("tokenurl") or "").strip()
        state_url = str(state_payload.get("stateurl") or "").strip()

        if token_url:
            # Looker OAuth V2 flow: redirect back to Looker's state_redir_url with encrypted code
            code_state = encrypt_oauth_state(
                {"code": code, "redirecturi": redirect_uri},
                settings.cipher_master,
            )
            resp: Response = RedirectResponse(
                url=f"{token_url}?state={code_state}", status_code=302
            )
            resp.delete_cookie("action_hub_state")
            return resp

        if state_url:
            # Looker OAuth V1 flow: exchange code directly and POST tokens to Looker's state_url
            tokens = config.exchange_oauth_code(code, redirect_uri, settings)
            token_doc: dict[str, object] = {"tokens": tokens, "redirect": redirect_uri}
            if settings.encrypt_payload:
                token_doc = {
                    "cid": cipher_id(settings.cipher_master),
                    "payload": encrypt_oauth_state(token_doc, settings.cipher_master),
                }
            req = urllib.request.Request(
                state_url,
                data=json.dumps(token_doc).encode("utf-8"),
                headers={"Content-Type": "application/json"},
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=10):
                pass

        html_resp = HTMLResponse(
            "<html><script>window.close()</script><body>You may now close this tab.</body></html>"
        )
        html_resp.delete_cookie("action_hub_state")
        return html_resp

    @app.post("/actions/pptx/oauth_token")
    async def action_oauth_token(
        payload: ActionWebhookRequest,
        authorization: str | None = Header(default=None),
    ) -> dict[str, object]:
        settings = load_settings()
        _require_action_auth(authorization, settings)

        raw_state = str(payload.state or payload.data.get("state") or "").strip()
        if not raw_state:
            raise HTTPException(status_code=400, detail="Request is missing state parameter.")

        try:
            decrypted = decrypt_oauth_state(raw_state, settings.cipher_master)
            code = str(decrypted.get("code") or "").strip()
            redirect_uri = str(decrypted.get("redirecturi") or "").strip()
            if not code or not redirect_uri:
                raise ValueError("Missing code or redirecturi in state.")
            tokens = config.exchange_oauth_code(code, redirect_uri, settings)
        except Exception as exc:
            raise HTTPException(
                status_code=400, detail=f"OAuth token exchange failed: {exc}"
            ) from exc

        if not tokens.get("refresh_token") and not tokens.get("access_token"):
            raise HTTPException(status_code=400, detail="OAuth tokens are invalid.")

        token_payload: dict[str, object] = {"tokens": tokens, "redirect": redirect_uri}
        if settings.encrypt_payload:
            return {
                "cid": cipher_id(settings.cipher_master),
                "payload": encrypt_oauth_state(token_payload, settings.cipher_master),
            }
        return token_payload

    @app.get("/actions/pptx/oauth_check")
    @app.post("/actions/pptx/oauth_check")
    async def action_oauth_check(
        payload: ActionWebhookRequest | None = None,
        authorization: str | None = Header(default=None),
    ) -> bool:
        settings = load_settings()
        _require_action_auth(authorization, settings)
        state_json = str((payload.data.get("state_json") if payload else "") or "").strip()
        if not state_json:
            return False
        try:
            resolve_action_oauth_token(state_json, settings)
            return True
        except Exception:
            return False

    @app.post("/actions/pptx/form")
    async def action_form(
        request: Request,
        payload: ActionWebhookRequest,
        authorization: str | None = Header(default=None),
    ) -> dict[str, object]:
        settings = load_settings()
        _require_action_auth(authorization, settings)
        base = _resolve_base_url(request, settings)

        if (
            settings.required_user_attribute
            and payload.data.get("test") is not True
            and not str(payload.data.get("action_access_param") or "").strip()
        ):
            return {
                "error": (
                    f"Access denied: Looker user attribute '{settings.required_user_attribute}' "
                    "is required to use this action."
                )
            }

        state_json = str(payload.data.get("state_json") or "").strip()
        has_oauth_flow_urls = bool(
            payload.state_redir_url
            or payload.state_url
            or payload.data.get("state_redir_url")
            or payload.data.get("state_url")
        )
        oauth_enabled = bool(settings.google_oauth_client_id or has_oauth_flow_urls or state_json)

        passthrough_token: str | None = None
        if oauth_enabled:
            if not state_json or state_json in ("reset", "null"):
                return _login_form(base, payload, settings)
            try:
                resolved_token = resolve_action_oauth_token(state_json, settings)
                if settings.oauth_gcs_passthrough:
                    passthrough_token = resolved_token
            except Exception as exc:
                return _login_form(base, payload, settings, error_message=str(exc))

        templates_uri = str(payload.data.get("templates_gcs_uri") or "").strip()
        output_uri = str(payload.data.get("output_gcs_uri") or "").strip()
        if not templates_uri or not output_uri:
            return {
                "error": (
                    "Required settings 'templates_gcs_uri' and 'output_gcs_uri' must be configured."
                )
            }

        try:
            parse_and_validate_gcs_uri(output_uri, settings.allowed_gcs_buckets)
            templates = gcs.list_pptx_templates(
                templates_uri,
                settings.allowed_gcs_buckets,
                oauth_token=passthrough_token,
            )
        except Exception as exc:
            if oauth_enabled:
                return _login_form(base, payload, settings, error_message=str(exc))
            return {"error": f"Failed to list templates from '{templates_uri}': {exc}"}

        return {
            "fields": [
                {
                    "name": "template",
                    "label": "PowerPoint Template (.pptx)",
                    "description": "Select a validated .pptx template from the templates bucket.",
                    "type": "select",
                    "required": True,
                    "options": [{"name": t, "label": t} for t in templates],
                },
                {
                    "name": "theme",
                    "label": "Slide Theme / Brand Palette",
                    "description": (
                        "Optional brand/theme palette (or 'auto' to use user attribute / template)."
                    ),
                    "type": "select",
                    "required": False,
                    "default": "auto",
                    "options": [
                        {"name": "auto", "label": "Auto (User Attribute / Template Default)"},
                        *[
                            {
                                "name": key,
                                "label": f"{val.brand} — {val.mode.title()} ({key})",
                            }
                            for key, val in PPTX_THEMES.items()
                        ],
                    ],
                },
                {
                    "name": "filename_prefix",
                    "label": "Output Filename Prefix",
                    "description": (
                        "Optional prefix for the timestamped .pptx written to the output bucket."
                    ),
                    "type": "string",
                    "required": False,
                    "default": "presentation",
                },
            ]
        }

    @app.post("/actions/pptx/execute")
    async def action_execute(
        payload: ActionWebhookRequest,
        authorization: str | None = Header(default=None),
        x_looker_webhook_id: str | None = Header(default=None),
        x_cloud_trace_context: str | None = Header(default=None),
    ) -> dict[str, object]:
        settings = load_settings()
        _require_action_auth(authorization, settings)

        webhook_id = x_looker_webhook_id or "manual"
        plan = payload.scheduled_plan
        plan_id = plan.scheduled_plan_id if plan else None

        if payload.type and payload.type != "dashboard":
            return _looker_response(
                webhook_id,
                success=False,
                message=f"Unsupported action type '{payload.type}'; expected 'dashboard'.",
            )

        if (
            settings.required_user_attribute
            and not str(payload.data.get("action_access_param") or "").strip()
        ):
            return _looker_response(
                webhook_id,
                success=False,
                message=(
                    f"Access denied: Looker user attribute '{settings.required_user_attribute}' "
                    "is required to execute this action."
                ),
            )

        state_json = str(payload.data.get("state_json") or "").strip()
        oauth_required = bool(settings.google_oauth_client_id or state_json)

        passthrough_token: str | None = None
        if oauth_required:
            if not state_json:
                return _looker_response(
                    webhook_id,
                    success=False,
                    message="No state found with OAuth credentials.",
                    state_data="reset",
                )
            try:
                resolved_token = resolve_action_oauth_token(state_json, settings)
                if settings.oauth_gcs_passthrough:
                    passthrough_token = resolved_token
            except Exception as exc:
                return _looker_response(
                    webhook_id,
                    success=False,
                    message=f"OAuth validation failed: {exc}",
                    state_data="reset",
                )

        templates_uri = str(payload.data.get("templates_gcs_uri") or "").strip()
        output_uri = str(payload.data.get("output_gcs_uri") or "").strip()
        template_name = str(payload.form_params.get("template") or "").strip()
        filename_prefix = str(
            payload.form_params.get("filename_prefix")
            or (plan.title if plan and plan.title else "")
            or "presentation"
        ).strip()

        if not templates_uri or not output_uri:
            return _looker_response(
                webhook_id,
                success=False,
                message="Missing 'templates_gcs_uri' or 'output_gcs_uri' in action params.",
            )

        if not template_name:
            return _looker_response(
                webhook_id,
                success=False,
                message="Missing required form parameter 'template'.",
                validation_errors={"template": "Please select a .pptx template."},
            )

        if not payload.attachment or not payload.attachment.data:
            return _looker_response(
                webhook_id,
                success=False,
                message="Missing csv_zip attachment payload.",
            )

        try:
            parse_and_validate_gcs_uri(output_uri, settings.allowed_gcs_buckets)
            csv_zip_bytes = _decode_base64_payload(payload.attachment.data, "attachment.data")
            template_bytes = gcs.download_template_bytes(
                templates_uri,
                template_name,
                settings.allowed_gcs_buckets,
                oauth_token=passthrough_token,
            )

            run_at = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
            plan_meta = {
                "title": (plan.title if plan and plan.title else filename_prefix),
                "url": (plan.url if plan and plan.url else ""),
                "run_at": run_at,
                "scheduled_plan_id": str(plan_id or ""),
                "webhook_id": webhook_id,
            }

            raw_theme = str(
                payload.form_params.get("theme") or payload.data.get("theme") or ""
            ).strip()
            user_brand = str(payload.data.get("brand") or "").strip()
            resolved_theme = resolve_pptx_theme(raw_theme or None, brand=user_brand or None)

            merge_res = merge_pptx(
                template_bytes,
                csv_zip_bytes,
                plan_meta=plan_meta,
                theme=resolved_theme,
            )
            out_gcs_uri = gcs.write_timestamped_output(
                output_uri,
                filename_prefix,
                webhook_id,
                merge_res.pptx_bytes,
                settings.allowed_gcs_buckets,
                oauth_token=passthrough_token,
            )

            log_structured(
                "pptx_merge_succeeded",
                trace_header=x_cloud_trace_context,
                webhook_id=webhook_id,
                scheduled_plan_id=plan_id,
                template=template_name,
                output_gcs_uri=out_gcs_uri,
                slides_processed=merge_res.slides_processed,
                charts_updated=merge_res.charts_updated,
                tables_updated=merge_res.tables_updated,
                text_tags_replaced=merge_res.text_tags_replaced,
                tile_count=len(merge_res.tile_row_counts),
                tile_row_counts=merge_res.tile_row_counts,
                oauth_passthrough=bool(passthrough_token),
            )
            return _looker_response(
                webhook_id,
                success=True,
                message=f"Generated {out_gcs_uri}",
            )
        except Exception as exc:
            is_expected = isinstance(exc, (MergeError, ValueError))
            log_structured(
                "pptx_merge_failed" if is_expected else "pptx_merge_unexpected_error",
                trace_header=x_cloud_trace_context,
                severity="WARNING" if is_expected else "ERROR",
                webhook_id=webhook_id,
                scheduled_plan_id=plan_id,
                template=template_name,
                error=str(exc),
            )
            msg = str(exc) if is_expected else f"Action execution failed: {exc}"
            return _looker_response(webhook_id, success=False, message=msg)

    @app.post("/templates/validate")
    async def validate_template(
        payload: TemplateValidateRequest,
        authorization: str | None = Header(default=None),
    ) -> dict[str, object]:
        settings = load_settings()
        _require_extension_auth(authorization, settings)

        try:
            pptx_bytes = _decode_base64_payload(payload.pptx_base64, "pptx_base64")
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

        report = inspect_and_validate(pptx_bytes, dashboard_tiles=payload.dashboard_tiles)
        return report.to_dict()

    @app.post("/templates/upload", response_model=None)
    async def upload_template(
        payload: TemplateUploadRequest,
        authorization: str | None = Header(default=None),
    ) -> JSONResponse | dict[str, object]:
        settings = load_settings()
        auth_mode, oauth_token = _require_extension_auth(authorization, settings)

        try:
            parse_and_validate_gcs_uri(payload.templates_gcs_uri, settings.allowed_gcs_buckets)
            pptx_bytes = _decode_base64_payload(payload.pptx_base64, "pptx_base64")
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

        report = inspect_and_validate(pptx_bytes, dashboard_tiles=payload.dashboard_tiles)
        if not report.valid:
            return JSONResponse(status_code=422, content=report.to_dict())

        passthrough_token = (
            oauth_token if (auth_mode == "oauth" and settings.oauth_gcs_passthrough) else None
        )
        try:
            uploaded_uri = gcs.upload_template_bytes(
                payload.templates_gcs_uri,
                payload.filename,
                pptx_bytes,
                settings.allowed_gcs_buckets,
                oauth_token=passthrough_token,
            )
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"GCS upload failed: {exc}") from exc

        return {"gcs_uri": uploaded_uri, **report.to_dict()}

    @app.post("/templates/preview", response_model=None)
    async def preview_template(
        payload: TemplatePreviewRequest,
        authorization: str | None = Header(default=None),
    ) -> JSONResponse | dict[str, object]:
        settings = load_settings()
        _require_extension_auth(authorization, settings)

        try:
            pptx_bytes = _decode_base64_payload(payload.pptx_base64, "pptx_base64")
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for slug, csv_content in payload.tile_csvs.items():
                clean_slug = slug.removesuffix(".csv")
                zf.writestr(f"{clean_slug}.csv", csv_content.encode("utf-8"))

        plan_meta = {
            "title": payload.plan_title or "Dashboard Preview",
            "url": payload.plan_url,
            "run_at": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "scheduled_plan_id": "preview",
            "webhook_id": "preview",
        }
        preview_theme = (
            theme_from_looker_palette(payload.theme or "Looker_Theme", payload.theme_palette)
            if payload.theme_palette
            else payload.theme
        )
        try:
            merge_res = merge_pptx(
                pptx_bytes,
                zip_buf.getvalue(),
                plan_meta=plan_meta,
                theme=preview_theme,
            )
        except (MergeError, ValueError) as exc:
            return JSONResponse(status_code=422, content={"error": str(exc)})

        return {
            "pptx_base64": base64.b64encode(merge_res.pptx_bytes).decode("ascii"),
            "slides_processed": merge_res.slides_processed,
            "charts_updated": merge_res.charts_updated,
            "tables_updated": merge_res.tables_updated,
            "text_tags_replaced": merge_res.text_tags_replaced,
            "tile_row_counts": merge_res.tile_row_counts,
        }

    def _list_viewer_bucket(gcs_uri: str, authorization: str | None) -> dict[str, object]:
        settings = load_settings()
        passthrough_token = _require_viewer_auth(authorization, settings)
        try:
            files = gcs.list_pptx_outputs(
                gcs_uri,
                settings.allowed_gcs_buckets,
                oauth_token=passthrough_token,
            )
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"GCS list failed: {exc}") from exc
        return {"files": files}

    def _download_viewer_bucket(
        gcs_uri: str, filename: str, authorization: str | None
    ) -> dict[str, object]:
        settings = load_settings()
        passthrough_token = _require_viewer_auth(authorization, settings)
        try:
            clean_name = config.sanitize_relative_pptx_name(filename)
            pptx_bytes = gcs.download_template_bytes(
                gcs_uri,
                clean_name,
                settings.allowed_gcs_buckets,
                oauth_token=passthrough_token,
            )
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"GCS download failed: {exc}") from exc
        return {
            "filename": clean_name,
            "pptx_base64": base64.b64encode(pptx_bytes).decode("ascii"),
        }

    @app.post("/outputs/list")
    async def list_outputs(
        payload: OutputListRequest,
        authorization: str | None = Header(default=None),
    ) -> dict[str, object]:
        return _list_viewer_bucket(payload.output_gcs_uri, authorization)

    @app.post("/outputs/download")
    async def download_output(
        payload: OutputDownloadRequest,
        authorization: str | None = Header(default=None),
    ) -> dict[str, object]:
        return _download_viewer_bucket(payload.output_gcs_uri, payload.filename, authorization)

    @app.post("/templates/list")
    async def list_templates_viewer(
        payload: TemplateListRequest,
        authorization: str | None = Header(default=None),
    ) -> dict[str, object]:
        return _list_viewer_bucket(payload.templates_gcs_uri, authorization)

    @app.post("/templates/download")
    async def download_template_viewer(
        payload: TemplateDownloadRequest,
        authorization: str | None = Header(default=None),
    ) -> dict[str, object]:
        return _download_viewer_bucket(payload.templates_gcs_uri, payload.filename, authorization)

    return app


app = create_app()
