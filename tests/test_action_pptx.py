from __future__ import annotations

import base64
import hashlib
import hmac
import json
from pathlib import Path

import pytest
from httpx import ASGITransport, AsyncClient

from action_pptx import gcs
from action_pptx.app import create_app


@pytest.fixture
def action_env(monkeypatch: pytest.MonkeyPatch) -> str:
    secret = "test-super-secret-key"
    monkeypatch.setenv("ACTION_HUB_SECRET", secret)
    monkeypatch.setenv("PUBLIC_BASE_URL", "https://action-pptx.a.run.app")
    monkeypatch.setenv("ALLOWED_GCS_BUCKETS", "templates-bucket,packages-bucket")
    monkeypatch.setenv("ALLOWED_OAUTH_DOMAINS", "google.com,example.com")
    monkeypatch.setenv("OAUTH_GCS_PASSTHROUGH", "true")
    return secret


def _make_hmac_token(secret: str, nonce: str = "abc123nonce") -> str:
    digest = hmac.new(secret.encode("utf-8"), nonce.encode("utf-8"), hashlib.sha512).hexdigest()
    return f"{nonce}/{digest}"


@pytest.mark.asyncio
async def test_healthz_and_auth_enforcement(
    action_env: str, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("ACTION_HUB_LABEL", "Custom Slides Hub")
    monkeypatch.setenv("ACTION_LABEL", "PowerPoint Slides")
    app = create_app()
    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://test") as client:
        r_health = await client.get("/health")
        assert r_health.status_code == 200
        assert r_health.json()["looker"]["success"] is True

        # Unauthorized POST /
        r_unauth = await client.post("/")
        assert r_unauth.status_code == 403

        # Authorized with raw secret
        r_raw = await client.post("/", headers={"Authorization": f'Token token="{action_env}"'})
        assert r_raw.status_code == 200
        data = r_raw.json()
        assert data["label"] == "Custom Slides Hub"
        assert data["integrations"][0]["name"] == "action-pptx"
        assert data["integrations"][0]["label"] == "PowerPoint Slides"
        assert data["integrations"][0]["icon_data_uri"].startswith("data:image/svg+xml;base64,")
        assert data["integrations"][0]["supported_formats"] == ["csv_zip"]
        assert data["integrations"][0]["supported_formattings"] == ["unformatted"]

        # Authorized with Action Hub nonce/digest HMAC-SHA512 format
        hmac_token = _make_hmac_token(action_env)
        r_hmac = await client.post(
            "/actions/pptx", headers={"Authorization": f'Token token="{hmac_token}"'}
        )
        assert r_hmac.status_code == 200
        assert r_hmac.json()["name"] == "action-pptx"


@pytest.mark.asyncio
async def test_action_form_lists_gcs_templates_and_checks_buckets(
    action_env: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        gcs,
        "list_pptx_templates",
        lambda uri, allowed, oauth_token=None: [
            "quarterly/two_period.pptx",
            "executive_summary.pptx",
        ],
    )
    app = create_app()
    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://test") as client:
        headers = {"Authorization": f'Token token="{action_env}"'}

        # Valid bucket (including Looker Admin UI boolean "test": True payload)
        r_ok = await client.post(
            "/actions/pptx/form",
            headers=headers,
            json={
                "data": {
                    "templates_gcs_uri": "gs://templates-bucket/team-a",
                    "output_gcs_uri": "gs://packages-bucket/out",
                    "test": True,
                }
            },
        )
        assert r_ok.status_code == 200
        fields = r_ok.json()["fields"]
        assert [f["name"] for f in fields] == ["template", "filename_prefix"]
        assert [o["name"] for o in fields[0]["options"]] == [
            "quarterly/two_period.pptx",
            "executive_summary.pptx",
        ]

        # Disallowed output bucket returns ActionForm error
        r_bad_bucket = await client.post(
            "/actions/pptx/form",
            headers=headers,
            json={
                "data": {
                    "templates_gcs_uri": "gs://templates-bucket/team-a",
                    "output_gcs_uri": "gs://evil-bucket/out",
                }
            },
        )
        assert r_bad_bucket.status_code == 200
        assert "not in ALLOWED_GCS_BUCKETS" in r_bad_bucket.json()["error"]


@pytest.mark.asyncio
async def test_required_user_attribute_restriction(
    action_env: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("RESTRICT_BY_USER_ATTRIBUTE", "true")
    monkeypatch.setattr(
        gcs,
        "list_pptx_templates",
        lambda uri, allowed, oauth_token=None: ["executive_summary.pptx"],
    )
    app = create_app()
    transport = ASGITransport(app=app)
    headers = {"Authorization": f'Token token="{action_env}"'}
    async with AsyncClient(transport=transport, base_url="http://test") as client:
        # 1. Action definition exposes required, sensitive user_attribute_name param
        r_meta = await client.post("/actions/pptx", headers=headers)
        assert r_meta.status_code == 200
        access_params = [p for p in r_meta.json()["params"] if p["name"] == "action_access_param"]
        assert access_params == [
            {
                "name": "action_access_param",
                "label": "Action Access",
                "description": "Controls which users can access this action.",
                "user_attribute_name": "can_use_action_pptx",
                "required": True,
                "sensitive": True,
            }
        ]

        # 2. Custom attribute name via REQUIRED_USER_ATTRIBUTE overrides default
        monkeypatch.setenv("REQUIRED_USER_ATTRIBUTE", "pptx_allowed")
        r_custom = await client.post("/actions/pptx", headers=headers)
        assert r_custom.json()["params"][-1]["user_attribute_name"] == "pptx_allowed"

        # 3. Looker Admin 'test': True check on /form still succeeds before user assignment
        r_admin_test = await client.post(
            "/actions/pptx/form",
            headers=headers,
            json={
                "data": {
                    "templates_gcs_uri": "gs://templates-bucket/team-a",
                    "output_gcs_uri": "gs://packages-bucket/out",
                    "test": True,
                }
            },
        )
        assert "fields" in r_admin_test.json()

        # 4. Unauthorized user without action_access_param is denied on /form and /execute
        r_denied_form = await client.post(
            "/actions/pptx/form",
            headers=headers,
            json={
                "data": {
                    "templates_gcs_uri": "gs://templates-bucket/team-a",
                    "output_gcs_uri": "gs://packages-bucket/out",
                }
            },
        )
        assert "Access denied" in r_denied_form.json()["error"]

        r_denied_exec = await client.post(
            "/actions/pptx/execute",
            headers=headers,
            json={
                "type": "dashboard",
                "data": {
                    "templates_gcs_uri": "gs://templates-bucket/team-a",
                    "output_gcs_uri": "gs://packages-bucket/out",
                },
                "form_params": {"template": "executive_summary.pptx"},
            },
        )
        assert r_denied_exec.json()["looker"]["success"] is False
        assert "Access denied" in r_denied_exec.json()["looker"]["message"]

        # 5. Authorized user with populated action_access_param succeeds on /form
        r_allowed = await client.post(
            "/actions/pptx/form",
            headers=headers,
            json={
                "data": {
                    "templates_gcs_uri": "gs://templates-bucket/team-a",
                    "output_gcs_uri": "gs://packages-bucket/out",
                    "action_access_param": "yes",
                }
            },
        )
        assert r_allowed.json()["fields"][0]["name"] == "template"


@pytest.mark.asyncio
async def test_action_execute_merges_and_writes_timestamped_pptx(
    action_env: str,
    sample_template_bytes: bytes,
    sample_csv_zip_bytes: bytes,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    written_outputs: list[tuple[str, str, bytes]] = []

    monkeypatch.setattr(
        gcs,
        "download_template_bytes",
        lambda uri, filename, allowed, oauth_token=None: sample_template_bytes,
    )

    def _fake_write_output(
        uri: str,
        prefix: str,
        data: bytes,
        allowed: frozenset[str],
        oauth_token: str | None = None,
    ) -> str:
        written_outputs.append((uri, prefix, data))
        return f"{uri}/{prefix}_20261113T000000Z.pptx"

    monkeypatch.setattr(gcs, "write_timestamped_output", _fake_write_output)

    app = create_app()
    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://test") as client:
        headers = {
            "Authorization": f'Token token="{action_env}"',
            "X-Looker-Webhook-Id": "wh-99",
        }
        r = await client.post(
            "/actions/pptx/execute?_theme=%7B%22background_color%22%3A%22%23f6f8fa%22%2C%22primary_button_color%22%3A%22%230b57d0%22%7D",
            headers=headers,
            json={
                "type": "dashboard",
                "data": {
                    "templates_gcs_uri": "gs://templates-bucket/finance",
                    "output_gcs_uri": "gs://packages-bucket/81",
                },
                "form_params": {
                    "template": "executive_master.pptx",
                    "filename_prefix": "QBR_81",
                },
                "scheduled_plan": {
                    "scheduled_plan_id": 42,
                    "title": "QBR Package 81",
                    "url": "https://looker.example/dashboards/9?theme=Levis_Light",
                },
                "attachment": {
                    "data": base64.encodebytes(sample_csv_zip_bytes).decode("ascii"),
                    "mimetype": "application/zip;base64",
                    "extension": "zip",
                },
            },
        )
        assert r.status_code == 200
        looker_resp = r.json()["looker"]
        assert looker_resp["success"] is True
        assert "gs://packages-bucket/81/QBR_81_20261113T000000Z.pptx" in looker_resp["message"]
        assert len(written_outputs) == 1
        assert len(written_outputs[0][2]) > 1000


@pytest.mark.asyncio
async def test_templates_validate_and_upload_with_token_and_oauth(
    action_env: str,
    sample_template_bytes: bytes,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    uploaded_calls: list[tuple[str, str, str | None]] = []

    def _fake_upload(
        uri: str,
        filename: str,
        data: bytes,
        allowed: frozenset[str],
        oauth_token: str | None = None,
    ) -> str:
        uploaded_calls.append((uri, filename, oauth_token))
        return f"{uri}/{filename}"

    monkeypatch.setattr(gcs, "upload_template_bytes", _fake_upload)
    monkeypatch.setattr(
        "action_pptx.config.verify_oauth_token_payload",
        lambda token: {"email": "analyst@example.com", "email_verified": "true"},
    )

    tiles = [
        {"title": "KPI Summary", "columns": ["Total Spend", "Trend Rate"]},
        {"title": "Mixed Metrics", "columns": ["Metric", "Value"]},
        {"title": "Top Brands", "columns": ["Brand Name", "Units Sold", "Net Revenue"]},
        {"title": "Monthly Trend", "columns": ["Quarter", "Direct", "Marketplace"]},
    ]
    pptx_b64 = base64.b64encode(sample_template_bytes).decode("ascii")

    app = create_app()
    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://test") as client:
        # 1. Validate via Action Hub Token
        r_val = await client.post(
            "/templates/validate",
            headers={"Authorization": f'Token token="{action_env}"'},
            json={
                "filename": "executive_review.pptx",
                "pptx_base64": pptx_b64,
                "dashboard_tiles": tiles,
            },
        )
        assert r_val.status_code == 200
        assert r_val.json()["valid"] is True

        # 2. Upload via Google OAuth Bearer token (testing passthrough)
        r_up = await client.post(
            "/templates/upload",
            headers={"Authorization": "Bearer mock-oauth-token"},
            json={
                "filename": "executive_review.pptx",
                "pptx_base64": pptx_b64,
                "dashboard_tiles": tiles,
                "templates_gcs_uri": "gs://templates-bucket/team",
            },
        )
        assert r_up.status_code == 200
        assert r_up.json()["gcs_uri"] == "gs://templates-bucket/team/executive_review.pptx"
        assert uploaded_calls == [
            ("gs://templates-bucket/team", "executive_review.pptx", "mock-oauth-token")
        ]

        # 3. Upload fails with 422 when validation fails
        r_fail = await client.post(
            "/templates/upload",
            headers={"Authorization": f'Token token="{action_env}"'},
            json={
                "filename": "broken.pptx",
                "pptx_base64": pptx_b64,
                "dashboard_tiles": [{"title": "KPI Summary", "columns": ["Total Spend"]}],
                "templates_gcs_uri": "gs://templates-bucket/team",
            },
        )
        assert r_fail.status_code == 422
        assert r_fail.json()["valid"] is False

        # 4. Live preview merges tile_csvs into pptx_base64
        r_prev = await client.post(
            "/templates/preview",
            headers={"Authorization": f'Token token="{action_env}"'},
            json={
                "pptx_base64": pptx_b64,
                "plan_title": "Live Preview Deck",
                "plan_url": "https://looker.example/dashboards/81",
                "tile_csvs": {
                    "kpi_summary": "Total Spend,Trend Rate\n34855339.456,0.0842\n",
                    "mixed_metrics": (
                        "Metric,Value\nCustomers,3378\n"
                        "Orders per Customer,62.44\nGross Margin,0.5568\n"
                    ),
                    "top_brands": (
                        "Brand Name,Units Sold,Net Revenue\n"
                        "Apex Goods,120,450200.4\n"
                        "Nova Apparel,340,289100.2\n"
                    ),
                    "monthly_trend": (
                        "Quarter,Direct,Marketplace\n"
                        "2026-Q1,1250000.5,320000.25\n"
                        "2026-Q2,1310000.0,340500.75\n"
                    ),
                },
            },
        )
        assert r_prev.status_code == 200
        prev_json = r_prev.json()
        assert prev_json["slides_processed"] == 1
        assert prev_json["charts_updated"] == 1
        assert prev_json["tables_updated"] == 1
        assert len(prev_json["pptx_base64"]) > 1000


@pytest.mark.asyncio
async def test_custom_action_oauth_v2_offline_and_passthrough(
    action_env: str,
    sample_template_bytes: bytes,
    sample_csv_zip_bytes: bytes,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("GOOGLE_OAUTH_CLIENT_ID", "shared-client-id.apps.googleusercontent.com")
    monkeypatch.setenv("GOOGLE_OAUTH_CLIENT_SECRET", "shared-client-secret")
    monkeypatch.setenv("ENCRYPT_PAYLOAD", "true")

    monkeypatch.setattr(
        "action_pptx.config.exchange_oauth_code",
        lambda code, redirect_uri, settings: {
            "access_token": "mock-initial-token",
            "refresh_token": "1//offline-refresh-token",
            "expires_in": 3600,
            "token_type": "Bearer",
        },
    )
    monkeypatch.setattr(
        "action_pptx.config.refresh_offline_access_token",
        lambda refresh_token, settings: {
            "access_token": "mock-refreshed-offline-token",
            "expires_in": 3600,
        },
    )
    monkeypatch.setattr(
        "action_pptx.config.verify_oauth_token_payload",
        lambda token: {"email": "analyst@example.com", "email_verified": "true"},
    )

    recorded_tokens: list[tuple[str, str | None]] = []

    def _list_with_oauth(
        uri: str, allowed: frozenset[str], oauth_token: str | None = None
    ) -> list[str]:
        recorded_tokens.append(("list", oauth_token))
        return ["quarterly/two_period.pptx"]

    def _download_with_oauth(
        uri: str, filename: str, allowed: frozenset[str], oauth_token: str | None = None
    ) -> bytes:
        recorded_tokens.append(("download", oauth_token))
        return sample_template_bytes

    def _write_with_oauth(
        uri: str,
        prefix: str,
        data: bytes,
        allowed: frozenset[str],
        oauth_token: str | None = None,
    ) -> str:
        recorded_tokens.append(("write", oauth_token))
        return f"{uri}/{prefix}_20261113T000000Z.pptx"

    monkeypatch.setattr(gcs, "list_pptx_templates", _list_with_oauth)
    monkeypatch.setattr(gcs, "download_template_bytes", _download_with_oauth)
    monkeypatch.setattr(gcs, "write_timestamped_output", _write_with_oauth)

    app = create_app()
    transport = ASGITransport(app=app)
    headers = {"Authorization": f'Token token="{action_env}"'}

    async with AsyncClient(transport=transport, base_url="http://test") as client:
        # 1. Integration metadata exposes uses_oauth, token_url, and shared oauth_client_id
        r_meta = await client.post("/actions/pptx", headers=headers)
        assert r_meta.status_code == 200
        meta = r_meta.json()
        assert meta["uses_oauth"] is True
        assert meta["token_url"] == "https://action-pptx.a.run.app/actions/pptx/oauth_token"
        assert meta["oauth_client_id"] == "shared-client-id.apps.googleusercontent.com"

        # 2. Unauthenticated /form returns oauth_link_google login field
        r_login_form = await client.post(
            "/actions/pptx/form",
            headers=headers,
            json={
                "state_redir_url": "https://looker.example/action_hub_state/123",
                "data": {
                    "templates_gcs_uri": "gs://templates-bucket/team",
                    "output_gcs_uri": "gs://packages-bucket/out",
                },
            },
        )
        assert r_login_form.status_code == 200
        login_field = r_login_form.json()["fields"][0]
        assert login_field["type"] == "oauth_link_google"
        oauth_url = login_field["oauth_url"]
        state_param = oauth_url.split("state=", 1)[1]

        # 3. GET /actions/pptx/oauth redirects to Google with access_type=offline & prompt=consent
        r_oauth = await client.get(f"/actions/pptx/oauth?state={state_param}")
        assert r_oauth.status_code == 302
        google_loc = r_oauth.headers["location"]
        assert "access_type=offline" in google_loc
        assert "prompt=consent" in google_loc
        assert "devstorage.read_write" in google_loc
        combined_state = google_loc.split("state=", 1)[1].split("&", 1)[0]
        csrf_cookie = r_oauth.cookies.get("action_hub_state")
        assert csrf_cookie is not None

        # 4. GET /actions/pptx/oauth_redirect redirects back to Looker tokenurl with encrypted code
        r_redir = await client.get(
            f"/actions/pptx/oauth_redirect?code=auth-code-xyz&state={combined_state}",
            cookies={"action_hub_state": csrf_cookie},
        )
        assert r_redir.status_code == 302
        looker_redir = r_redir.headers["location"]
        assert looker_redir.startswith("https://looker.example/action_hub_state/123?state=")
        code_state = looker_redir.split("state=", 1)[1]

        # 5. POST /actions/pptx/oauth_token exchanges code and returns encrypted token payload
        r_tok = await client.post(
            "/actions/pptx/oauth_token",
            headers=headers,
            json={"state": code_state},
        )
        assert r_tok.status_code == 200
        encrypted_state_json = r_tok.text
        assert "cid" in r_tok.json() and "payload" in r_tok.json()

        # 6. Authenticated /form passes token through to gcs.list_pptx_templates
        r_authed_form = await client.post(
            "/actions/pptx/form",
            headers=headers,
            json={
                "data": {
                    "templates_gcs_uri": "gs://templates-bucket/team",
                    "output_gcs_uri": "gs://packages-bucket/out",
                    "state_json": encrypted_state_json,
                }
            },
        )
        assert r_authed_form.status_code == 200
        assert r_authed_form.json()["fields"][0]["name"] == "template"

        # 7. Offline scheduled /execute refreshes offline refresh_token when access_token is empty
        offline_only_state = '{"tokens": {"refresh_token": "1//offline-refresh-token"}}'
        r_exec = await client.post(
            "/actions/pptx/execute",
            headers={**headers, "X-Looker-Webhook-Id": "wh-offline"},
            json={
                "type": "dashboard",
                "data": {
                    "templates_gcs_uri": "gs://templates-bucket/team",
                    "output_gcs_uri": "gs://packages-bucket/out",
                    "state_json": offline_only_state,
                },
                "form_params": {
                    "template": "executive_master.pptx",
                    "filename_prefix": "OfflineDeck",
                },
                "attachment": {
                    "data": base64.encodebytes(sample_csv_zip_bytes).decode("ascii"),
                    "mimetype": "application/zip;base64",
                    "extension": "zip",
                },
            },
        )
        assert r_exec.status_code == 200
        assert r_exec.json()["looker"]["success"] is True

        assert recorded_tokens == [
            ("list", "mock-initial-token"),
            ("download", "mock-refreshed-offline-token"),
            ("write", "mock-refreshed-offline-token"),
        ]

        # 8. Missing state_json on /execute resets Looker OAuth state
        r_missing = await client.post(
            "/actions/pptx/execute",
            headers=headers,
            json={
                "type": "dashboard",
                "data": {
                    "templates_gcs_uri": "gs://templates-bucket/team",
                    "output_gcs_uri": "gs://packages-bucket/out",
                },
                "form_params": {"template": "executive_master.pptx"},
            },
        )
        assert r_missing.status_code == 200
        assert r_missing.json()["looker"]["success"] is False
        assert r_missing.json()["state"]["data"] == "reset"


def test_gcs_helpers_and_config_edge_cases(monkeypatch: pytest.MonkeyPatch) -> None:
    from action_pptx.config import (
        log_structured,
        parse_and_validate_gcs_uri,
        sanitize_relative_pptx_name,
        verify_oauth_bearer,
    )

    class _FakeBlob:
        def __init__(self, name: str) -> None:
            self.name = name
            self.uploaded: bytes = b""

        def download_as_bytes(self) -> bytes:
            return b"pptx-binary-content"

        def upload_from_string(self, data: bytes, content_type: str = "") -> None:
            self.uploaded = data

    class _FakeBucket:
        def __init__(self) -> None:
            self.blobs: dict[str, _FakeBlob] = {}

        def blob(self, path: str) -> _FakeBlob:
            b = _FakeBlob(path)
            self.blobs[path] = b
            return b

    class _FakeClient:
        def __init__(self) -> None:
            self.bucket_inst = _FakeBucket()

        def list_blobs(self, bucket_name: str, prefix: str = "") -> list[_FakeBlob]:
            return [
                _FakeBlob(f"{prefix}sub/deck.pptx"),
                _FakeBlob(f"{prefix}readme.txt"),
                _FakeBlob(f"{prefix}master.pptx"),
            ]

        def bucket(self, name: str) -> _FakeBucket:
            return self.bucket_inst

    fake_client = _FakeClient()
    monkeypatch.setattr(gcs, "_get_storage_client", lambda oauth_token=None: fake_client)

    allowed = frozenset({"my-bucket"})
    listed = gcs.list_pptx_templates("gs://my-bucket/templates", allowed)
    assert listed == ["master.pptx", "sub/deck.pptx"]

    raw = gcs.download_template_bytes("gs://my-bucket/templates", "master", allowed)
    assert raw == b"pptx-binary-content"

    up_uri = gcs.upload_template_bytes(
        "gs://my-bucket/templates", "sub/new_deck.pptx", b"123", allowed, "tok"
    )
    assert up_uri == "gs://my-bucket/templates/sub/new_deck.pptx"

    out_uri = gcs.write_timestamped_output("gs://my-bucket/out", "Q3 Report!", b"456", allowed)
    assert out_uri.startswith("gs://my-bucket/out/Q3_Report_")
    assert out_uri.endswith("Z.pptx")
    assert "_wh_" not in out_uri

    fanout_uri = gcs.write_timestamped_output(
        "gs://my-bucket/out", "/Q4 EBR/Levi's.pptx", b"789", allowed
    )
    assert fanout_uri.startswith("gs://my-bucket/out/Q4_EBR/Levi_s_")
    assert fanout_uri.endswith("Z.pptx")

    with pytest.raises(ValueError, match="Path traversal rejected in output prefix"):
        gcs.write_timestamped_output("gs://my-bucket/out", "Q4_EBR/../secret", b"000", allowed)

    # Config edge cases
    with pytest.raises(ValueError, match="Invalid GCS URI"):
        parse_and_validate_gcs_uri("https://not-gcs", allowed)
    with pytest.raises(ValueError, match="Invalid GCS path prefix"):
        parse_and_validate_gcs_uri("gs://my-bucket/a/../b", allowed)
    with pytest.raises(ValueError, match="Path traversal rejected"):
        sanitize_relative_pptx_name("../secret.pptx")

    monkeypatch.setattr(
        "action_pptx.config.verify_oauth_token_payload",
        lambda token: {"email": "user@evil.com", "email_verified": "true"},
    )
    with pytest.raises(PermissionError, match="not in ALLOWED_OAUTH_DOMAINS"):
        verify_oauth_bearer("Bearer tok", frozenset({"google.com"}))

    log_structured("test_event", trace_header="105445aa7843bc8bf206b12000100000/1;o=1", ok=True)


def test_looker_palette_preview_and_deploy_manifest_helpers(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from bin.register_integration import (
        build_manifest_application_block,
        build_secret_domain_allowlist,
    )
    from pptx_merge.themes import theme_from_looker_palette

    assert build_secret_domain_allowlist("https://www.lkr.dev/apps/action-pptx/") == (
        "https://www.lkr.dev/apps/action-pptx,"
        "https://www.lkr.dev/apps/action-pptx/,"
        "https://www.lkr.dev/apps/action-pptx/*"
    )

    existing_block = build_manifest_application_block(
        "https://www.lkr.dev/apps/action-pptx", use_local_bundle=False
    )
    assert 'url: "https://cdn.lkr.dev/apps/action-pptx/latest/bundle.js"' in existing_block
    assert '"https://www.lkr.dev/apps/action-pptx/*"' in existing_block
    assert '"all_themes"' in existing_block
    assert "new_window: yes" in existing_block
    assert "use_downloads: yes" in existing_block

    new_proj_block = build_manifest_application_block(
        "https://www.lkr.dev/apps/action-pptx", use_local_bundle=True
    )
    assert 'file: "bundle.js"' in new_proj_block

    custom_theme = theme_from_looker_palette(
        "Shipping_Light",
        {
            "bg_color": "#f6f8fa",
            "card_bg": "#ffffff",
            "title_color": "#3a4245",
            "text_color": "#3a4245",
            "accent": "#1f78b4",
            "series_colors": ["#d19", "#d18a", "#059669"],
        },
    )
    assert custom_theme.mode == "light"
    assert custom_theme.series_palette[0] == "DD1199"
    assert custom_theme.series_palette[1] == "DD1188"

    monkeypatch.setenv("ACTION_HUB_SECRET", "test-secret")
    import base64
    import io
    import zipfile

    from fastapi.testclient import TestClient

    from action_pptx.app import create_app
    from tests.conftest import build_showcase_csv_zip_bytes, build_showcase_template_bytes

    tmpl_b64 = base64.b64encode(build_showcase_template_bytes()).decode("ascii")
    csv_map: dict[str, str] = {}
    with zipfile.ZipFile(io.BytesIO(build_showcase_csv_zip_bytes())) as zf:
        for name in zf.namelist():
            if name.endswith(".csv"):
                csv_map[name.removesuffix(".csv")] = zf.read(name).decode("utf-8")

    client = TestClient(create_app())
    resp = client.post(
        "/templates/preview",
        headers={"Authorization": 'Token token="test-secret"'},
        json={
            "pptx_base64": tmpl_b64,
            "tile_csvs": csv_map,
            "plan_title": "Palette Test",
            "theme": "Shipping_Light",
            "theme_palette": {
                "bg_color": "#f6f8fa",
                "card_bg": "#ffffff",
                "title_color": "#0f172a",
                "text_color": "#334155",
                "accent": "#1f78b4",
                "series_colors": ["#d19", "#1f78b4", "#059669"],
            },
        },
    )
    assert resp.status_code == 200
    assert len(resp.json()["pptx_base64"]) > 1000


def test_stage_lookml_project_dir_and_codemode_ast(tmp_path: Path) -> None:
    import ast

    from bin.register_integration import stage_lookml_project_dir

    fake_bundle = tmp_path / "bundle.js"
    fake_bundle.write_text("console.log('action-pptx');", encoding="utf-8")

    staged_sample = tmp_path / "with_sample"
    used_local = stage_lookml_project_dir(
        staged_sample,
        service_url="https://action-pptx.a.run.app/",
        project_id="my-pptx-proj",
        model_name="custom_pptx_model",
        connection_name="bq_conn",
        bundle_path=str(fake_bundle),
        include_sample_lookml=True,
    )
    assert used_local is True
    assert (staged_sample / "bundle.js").read_text(
        encoding="utf-8"
    ) == "console.log('action-pptx');"
    manifest = (staged_sample / "manifest.lkml").read_text(encoding="utf-8")
    assert 'project_name: "my-pptx-proj"' in manifest
    assert 'file: "bundle.js"' in manifest
    assert 'url: "https://cdn.lkr.dev' not in manifest
    assert '"https://action-pptx.a.run.app/*"' in manifest

    model_txt = (staged_sample / "models" / "custom_pptx_model.model.lkml").read_text(
        encoding="utf-8"
    )
    assert 'connection: "bq_conn"' in model_txt
    assert 'view_label: ""' not in model_txt
    assert "datagroup: ebr_daily_datagroup" in model_txt
    assert (staged_sample / "views" / "inventory_items.view.lkml").is_file()
    dash_txt = (
        staged_sample / "dashboards" / "executive_business_review.dashboard.lookml"
    ).read_text(encoding="utf-8")
    assert "model: custom_pptx_model" in dash_txt
    assert "model: action_pptx" not in dash_txt
    assert "Region: inventory_items.region" in dash_txt

    codemode_path = Path(__file__).resolve().parent.parent / "bin" / "looker_codemode.py"
    codemode_src = codemode_path.read_text(encoding="utf-8")
    ast.parse(codemode_src)
    for key in (
        "cloudRunUrl",
        "templatesGcsUri",
        "outputGcsUri",
        "actionIntegrationId",
        "useOAuth",
        "oauthClientId",
        "oauthGcsPassthrough",
        "update_data_app",
        "update_artifacts",
    ):
        assert key in codemode_src

    import json

    captured_ctx: list[tuple[str, dict[str, dict[str, object]]]] = []
    captured_artifacts: list[tuple[str, list[dict[str, object]]]] = []
    existing_json = json.dumps(
        {
            "template": "simple_template.pptx",
            "dashboards": [{"id": "3965", "title": "Executive Business Review UDD"}],
        }
    )
    ns: dict[str, object] = {
        "action": "configure_extension",
        "service_url": "https://action-pptx.a.run.app/",
        "templates_gcs_uri": "gs://b/templates",
        "output_gcs_uri": "gs://b/output",
        "action_integration_id": "28::action-pptx",
        "project_id": "my-pptx-proj",
        "model_name": "custom_pptx_model",
        "include_sample_lookml": "true",
        "update_data_app": lambda data_app_id, body: (
            captured_ctx.append((data_app_id, body)) or body
        ),
        "artifact": lambda namespace, key: (
            [{"key": key, "version": 4, "value": existing_json}]
            if "simple-5ftemplate" in key
            else []
        ),
        "update_artifacts": lambda namespace, body: (
            captured_artifacts.append((namespace, body)) or body
        ),
        "dashboard": lambda dashboard_id: {
            "id": dashboard_id,
            "title": "Brand Overview",
        },
    }
    exec(compile(codemode_src, str(codemode_path), "exec"), ns)
    app_id, body = captured_ctx[0]
    assert app_id == "my-pptx-proj::action_pptx_template_manager"
    assert body["context"]["actionIntegrationId"] == "28::action-pptx"
    assert len(captured_artifacts) == 2
    simple_ns, simple_body = captured_artifacts[0]
    assert simple_ns == "my-pptx-proj::action_pptx_template_manager"
    assert simple_body[0]["key"] == "tpl_links.simple-5ftemplate-2epptx"
    assert simple_body[0]["version"] == 4
    parsed_simple = json.loads(str(simple_body[0]["value"]))
    assert [d["id"] for d in parsed_simple["dashboards"]] == [
        "custom_pptx_model::executive_business_review",
        "3965",
    ]
    _, embed_body = captured_artifacts[1]
    assert embed_body[0]["key"] == "tpl_links.embed-5fdemo-5ftemplate-2epptx"
    assert "version" not in embed_body[0]
    parsed_embed = json.loads(str(embed_body[0]["value"]))
    assert parsed_embed["dashboards"] == [
        {"id": "embed_demo::brand_overview", "title": "Brand Overview"}
    ]


def test_register_action_hub_passes_integration_id_to_update_data_app(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import subprocess

    from bin import register_integration

    calls: list[dict[str, str]] = []

    def _fake_codemode(
        *, variables: dict[str, str] | None = None, **_: object
    ) -> dict[str, object]:
        v = variables or {}
        calls.append(v)
        if v.get("action") == "setup_hub":
            return {"hub_id": "42", "hub_action": "created", "integration_id": "42::action-pptx"}
        return {"configured": True}

    monkeypatch.setattr(register_integration, "_run_lkr_codemode", _fake_codemode)
    monkeypatch.setattr(register_integration, "_load_bundle_js", lambda _p=None: "console.log(1);")
    monkeypatch.setattr(subprocess, "run", lambda *a, **kw: subprocess.CompletedProcess(a[0], 0))

    ext_id, warnings = register_integration.register_action_hub_and_extension(
        service_url="https://action-pptx.a.run.app",
        api_key_token="nonce/digest",
        templates_gcs_uri="gs://b/templates",
        output_gcs_uri="gs://b/output",
        extension_mode="new",
        project_id="my-proj",
        connection_name="bq_conn",
    )
    assert ext_id == "my-proj::action_pptx_template_manager"
    assert warnings == []
    cfg_call = next(c for c in calls if c.get("action") == "configure_extension")
    assert cfg_call["action_integration_id"] == "42::action-pptx"


@pytest.mark.asyncio
async def test_outputs_viewer_routes_and_disable_env(
    action_env: str,
    sample_template_bytes: bytes,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from datetime import UTC, datetime

    class _OutBlob:
        def __init__(self, name: str, size: int, updated: datetime) -> None:
            self.name = name
            self.size = size
            self.updated = updated

    class _OutClient:
        def list_blobs(self, bucket_name: str, prefix: str = "") -> list[_OutBlob]:
            return [
                _OutBlob(
                    f"{prefix}QBR_old.pptx",
                    1200,
                    datetime(2026, 11, 1, 10, 0, tzinfo=UTC),
                ),
                _OutBlob(f"{prefix}notes.txt", 50, datetime(2026, 11, 15, 10, 0, tzinfo=UTC)),
                _OutBlob(
                    f"{prefix}EBR_new.pptx",
                    2400,
                    datetime(2026, 11, 16, 12, 0, tzinfo=UTC),
                ),
            ]

    recorded_oauth: list[str | None] = []

    def _fake_storage_client(oauth_token: str | None = None) -> _OutClient:
        recorded_oauth.append(oauth_token)
        return _OutClient()

    monkeypatch.setattr(gcs, "_get_storage_client", _fake_storage_client)
    monkeypatch.setattr(
        "action_pptx.config.verify_oauth_token_payload",
        lambda token: {"email": "analyst@example.com", "email_verified": "true"},
    )
    monkeypatch.setattr(
        gcs,
        "download_template_bytes",
        lambda uri, filename, allowed, oauth_token=None: (
            recorded_oauth.append(oauth_token) or sample_template_bytes
        ),
    )
    deleted_blobs: list[tuple[str, str, str | None]] = []
    monkeypatch.setattr(
        gcs,
        "delete_template_blob",
        lambda uri, filename, allowed, oauth_token=None: (
            deleted_blobs.append((uri, filename, oauth_token)) or f"{uri}/{filename}"
        ),
    )

    app = create_app()
    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://test") as client:
        r_health = await client.get("/health")
        assert r_health.status_code == 200
        assert r_health.json()["report_viewer_enabled"] is True

        # List sorted newest first + OAuth passthrough
        r_list = await client.post(
            "/outputs/list",
            headers={"Authorization": "Bearer viewer-oauth-tok"},
            json={"output_gcs_uri": "gs://packages-bucket/out"},
        )
        assert r_list.status_code == 200
        files = r_list.json()["files"]
        assert [f["name"] for f in files] == ["EBR_new.pptx", "QBR_old.pptx"]
        assert files[0]["size"] == 2400

        # Disallowed bucket rejected
        r_bad = await client.post(
            "/outputs/list",
            headers={"Authorization": f'Token token="{action_env}"'},
            json={"output_gcs_uri": "gs://evil-bucket/out"},
        )
        assert r_bad.status_code == 400

        # Download returns base64 pptx
        r_dl = await client.post(
            "/outputs/download",
            headers={"Authorization": "Bearer viewer-oauth-tok"},
            json={"output_gcs_uri": "gs://packages-bucket/out", "filename": "EBR_new.pptx"},
        )
        assert r_dl.status_code == 200
        assert r_dl.json()["filename"] == "EBR_new.pptx"
        assert base64.b64decode(r_dl.json()["pptx_base64"]) == sample_template_bytes
        assert recorded_oauth == ["viewer-oauth-tok", "viewer-oauth-tok"]

        # Path traversal rejected
        r_trav = await client.post(
            "/outputs/download",
            headers={"Authorization": f'Token token="{action_env}"'},
            json={"output_gcs_uri": "gs://packages-bucket/out", "filename": "../secret.pptx"},
        )
        assert r_trav.status_code == 400

        # Template viewer routes share auth and disable env semantics
        r_tpl_list = await client.post(
            "/templates/list",
            headers={"Authorization": "Bearer viewer-oauth-tok"},
            json={"templates_gcs_uri": "gs://templates-bucket/team"},
        )
        assert r_tpl_list.status_code == 200
        assert [f["name"] for f in r_tpl_list.json()["files"]] == ["EBR_new.pptx", "QBR_old.pptx"]

        r_tpl_dl = await client.post(
            "/templates/download",
            headers={"Authorization": "Bearer viewer-oauth-tok"},
            json={"templates_gcs_uri": "gs://templates-bucket/team", "filename": "EBR_new.pptx"},
        )
        assert r_tpl_dl.status_code == 200
        assert base64.b64decode(r_tpl_dl.json()["pptx_base64"]) == sample_template_bytes

        r_tpl_del = await client.post(
            "/templates/delete",
            headers={"Authorization": "Bearer viewer-oauth-tok"},
            json={"templates_gcs_uri": "gs://templates-bucket/team", "filename": "QBR_old.pptx"},
        )
        assert r_tpl_del.status_code == 200
        assert r_tpl_del.json()["deleted"] == "QBR_old.pptx"
        assert deleted_blobs == [("gs://templates-bucket/team", "QBR_old.pptx", "viewer-oauth-tok")]

        r_tpl_del_trav = await client.post(
            "/templates/delete",
            headers={"Authorization": f'Token token="{action_env}"'},
            json={"templates_gcs_uri": "gs://templates-bucket/team", "filename": "../evil.pptx"},
        )
        assert r_tpl_del_trav.status_code == 400

        # Server env var disables viewer routes and /health flag
        monkeypatch.setenv("DISABLE_REPORT_VIEWER", "true")
        r_health_off = await client.get("/health")
        assert r_health_off.status_code == 200
        assert r_health_off.json()["report_viewer_enabled"] is False

        r_list_off = await client.post(
            "/outputs/list",
            headers={"Authorization": f'Token token="{action_env}"'},
            json={"output_gcs_uri": "gs://packages-bucket/out"},
        )
        assert r_list_off.status_code == 403

        r_dl_off = await client.post(
            "/outputs/download",
            headers={"Authorization": f'Token token="{action_env}"'},
            json={"output_gcs_uri": "gs://packages-bucket/out", "filename": "EBR_new.pptx"},
        )
        assert r_dl_off.status_code == 403

        r_tpl_off = await client.post(
            "/templates/list",
            headers={"Authorization": f'Token token="{action_env}"'},
            json={"templates_gcs_uri": "gs://templates-bucket/team"},
        )
        assert r_tpl_off.status_code == 403

        r_tpl_del_off = await client.post(
            "/templates/delete",
            headers={"Authorization": f'Token token="{action_env}"'},
            json={"templates_gcs_uri": "gs://templates-bucket/team", "filename": "QBR_old.pptx"},
        )
        assert r_tpl_del_off.status_code == 403


def test_register_action_hub_soft_fails_without_blocking_extension(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import bin.register_integration as reg

    def fake_codemode(**_kwargs: object) -> object:
        raise RuntimeError("SocketError: name or service not known")

    monkeypatch.setattr(reg, "_run_lkr_codemode", fake_codemode)
    monkeypatch.setattr(
        reg,
        "deploy_extension_to_new_project",
        lambda *_args, **_kwargs: "action-pptx::action_pptx_template_manager",
    )

    ext_id, warnings = reg.register_action_hub_and_extension(
        service_url="https://www.lkr.dev/apps/action-pptx-test",
        api_key_token="raw-secret",
        templates_gcs_uri="gs://b/templates",
        output_gcs_uri="gs://b/output",
        extension_mode="new",
        connection_name="bq_conn",
    )
    assert ext_id == "action-pptx::action_pptx_template_manager"
    assert len(warnings) == 1
    assert "Action Hub registration failed" in warnings[0]


def test_node_modules_in_sync(tmp_path: Path) -> None:
    from bin.register_integration import node_modules_in_sync

    ext = tmp_path / "extension"
    nm = ext / "node_modules"
    nm.mkdir(parents=True)
    react = {"version": "18.3.1", "integrity": "sha512-aaa"}
    embed = {"version": "2.0.10", "integrity": "sha512-bbb"}

    (ext / "package-lock.json").write_text(
        json.dumps({"packages": {"": {"version": "0.1.0"}, "node_modules/react": react}}),
        encoding="utf-8",
    )
    assert node_modules_in_sync(ext) is False

    (nm / ".package-lock.json").write_text(
        json.dumps({"packages": {"node_modules/react": react}}),
        encoding="utf-8",
    )
    assert node_modules_in_sync(ext) is True

    # Git dependency resolved SHA bump with unchanged version and no integrity -> out of sync
    (ext / "package-lock.json").write_text(
        json.dumps(
            {
                "packages": {
                    "": {"version": "0.1.0"},
                    "node_modules/react": {**react, "resolved": "git+https://example.com/r#abc"},
                }
            }
        ),
        encoding="utf-8",
    )
    assert node_modules_in_sync(ext) is False

    (ext / "package-lock.json").write_text(
        json.dumps(
            {
                "packages": {
                    "": {"version": "0.1.0"},
                    "node_modules/react": react,
                    "node_modules/@looker/embed-sdk": embed,
                }
            }
        ),
        encoding="utf-8",
    )
    assert node_modules_in_sync(ext) is False

    (nm / ".package-lock.json").write_text("{broken", encoding="utf-8")
    assert node_modules_in_sync(ext) is False
