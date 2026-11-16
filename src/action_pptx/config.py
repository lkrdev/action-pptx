from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import re
import secrets
import sys
import urllib.parse
import urllib.request
from dataclasses import dataclass

from google.auth.transport import requests as google_requests
from google.oauth2 import id_token

_TOKEN_RE = re.compile(r'^[Tt]oken\s+(?:token="([^"]+)"|(\S+))$')
_BEARER_RE = re.compile(r"^[Bb]earer\s+(\S+)$")
GOOGLE_AUTH_ENDPOINT = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_TOKEN_ENDPOINT = "https://oauth2.googleapis.com/token"
OAUTH_SCOPES_PASSTHROUGH = "openid email https://www.googleapis.com/auth/devstorage.read_write"
OAUTH_SCOPES_IDENTITY = "openid email"


@dataclass(frozen=True, slots=True)
class Settings:
    action_hub_secret: str
    public_base_url: str
    action_hub_label: str
    allowed_gcs_buckets: frozenset[str]
    allowed_oauth_domains: frozenset[str]
    oauth_gcs_passthrough: bool
    google_oauth_client_id: str = ""
    google_oauth_client_secret: str = ""
    cipher_master: str = ""
    encrypt_payload: bool = False


def _split_csv_env(name: str) -> frozenset[str]:
    return frozenset(p.strip().lower() for p in os.environ.get(name, "").split(",") if p.strip())


def load_settings() -> Settings:
    base_url = (os.environ.get("PUBLIC_BASE_URL") or "http://localhost:8080").rstrip("/")
    secret = os.environ.get("ACTION_HUB_SECRET", "")
    client_id = os.environ.get("GOOGLE_OAUTH_CLIENT_ID", "").strip()
    client_secret = os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET", "").strip()
    cipher_master = (os.environ.get("CIPHER_MASTER") or secret or "action-pptx-default-key").strip()
    return Settings(
        action_hub_secret=secret,
        public_base_url=base_url,
        action_hub_label=os.environ.get("ACTION_HUB_LABEL", "Looker PowerPoint Action Hub"),
        allowed_gcs_buckets=_split_csv_env("ALLOWED_GCS_BUCKETS"),
        allowed_oauth_domains=_split_csv_env("ALLOWED_OAUTH_DOMAINS"),
        oauth_gcs_passthrough=os.environ.get("OAUTH_GCS_PASSTHROUGH", "").lower() == "true",
        google_oauth_client_id=client_id,
        google_oauth_client_secret=client_secret,
        cipher_master=cipher_master,
        encrypt_payload=os.environ.get("ENCRYPT_PAYLOAD", "").lower() == "true",
    )


def _keystream_xor(data: bytes, key: bytes, nonce: bytes) -> bytes:
    """XOR data with a SHA-256 counter keystream derived from key and nonce."""
    out = bytearray(len(data))
    offset = 0
    counter = 0
    while offset < len(data):
        block = hashlib.sha256(key + nonce + counter.to_bytes(4, "big")).digest()
        chunk = data[offset : offset + 32]
        for i, b in enumerate(chunk):
            out[offset + i] = b ^ block[i]
        offset += 32
        counter += 1
    return bytes(out)


def cipher_id(secret: str) -> str:
    return hashlib.sha256(secret.encode("utf-8")).hexdigest()[:16]


def encrypt_oauth_state(payload: dict[str, object], secret: str) -> str:
    """Encrypt and MAC a JSON state dictionary into a URL-safe token."""
    key = hashlib.sha256(secret.encode("utf-8")).digest()
    nonce = secrets.token_bytes(16)
    plaintext = json.dumps(payload, separators=(",", ":")).encode("utf-8")
    ciphertext = _keystream_xor(plaintext, key, nonce)
    mac = hmac.new(key, nonce + ciphertext, hashlib.sha256).digest()
    return base64.urlsafe_b64encode(nonce + mac + ciphertext).decode("ascii").rstrip("=")


def decrypt_oauth_state(token: str, secret: str) -> dict[str, object]:
    """Verify MAC and decrypt a URL-safe state token back into a dictionary."""
    cleaned = (token or "").strip()
    pad = "=" * ((4 - len(cleaned) % 4) % 4)
    raw = base64.urlsafe_b64decode(cleaned + pad)
    if len(raw) < 48:
        raise ValueError("Invalid encrypted state payload.")
    nonce, mac, ciphertext = raw[:16], raw[16:48], raw[48:]
    key = hashlib.sha256(secret.encode("utf-8")).digest()
    expected_mac = hmac.new(key, nonce + ciphertext, hashlib.sha256).digest()
    if not hmac.compare_digest(mac, expected_mac):
        raise ValueError("State signature verification failed.")
    plaintext = _keystream_xor(ciphertext, key, nonce)
    parsed = json.loads(plaintext.decode("utf-8"))
    if not isinstance(parsed, dict):
        raise ValueError("Decrypted state must be a JSON object.")
    return dict(parsed)


def build_google_oauth_url(
    client_id: str,
    redirect_uri: str,
    state: str,
    *,
    passthrough: bool = True,
) -> str:
    """Build Google OAuth 2.0 authorization URL requesting an offline refresh token."""
    params = {
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": OAUTH_SCOPES_PASSTHROUGH if passthrough else OAUTH_SCOPES_IDENTITY,
        "access_type": "offline",
        "prompt": "consent",
        "state": state,
    }
    return f"{GOOGLE_AUTH_ENDPOINT}?{urllib.parse.urlencode(params)}"


def _post_oauth_token(params: dict[str, str]) -> dict[str, object]:
    req = urllib.request.Request(
        GOOGLE_TOKEN_ENDPOINT,
        data=urllib.parse.urlencode(params).encode("utf-8"),
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=10) as resp:
        return dict(json.loads(resp.read().decode("utf-8")))


def exchange_oauth_code(
    code: str,
    redirect_uri: str,
    settings: Settings,
) -> dict[str, object]:
    """Exchange an authorization code for access_token and offline refresh_token."""
    return _post_oauth_token(
        {
            "code": code,
            "client_id": settings.google_oauth_client_id,
            "client_secret": settings.google_oauth_client_secret,
            "redirect_uri": redirect_uri,
            "grant_type": "authorization_code",
        }
    )


def refresh_offline_access_token(
    refresh_token: str,
    settings: Settings,
) -> dict[str, object]:
    """Use an offline refresh_token to mint a fresh Google OAuth access_token."""
    return _post_oauth_token(
        {
            "refresh_token": refresh_token,
            "client_id": settings.google_oauth_client_id,
            "client_secret": settings.google_oauth_client_secret,
            "grant_type": "refresh_token",
        }
    )


def extract_oauth_state_tokens(
    state_json: str | None,
    settings: Settings,
) -> dict[str, object] | None:
    """Extract OAuth tokens dict from Looker's state_json (plain or encrypted)."""
    cleaned = (state_json or "").strip()
    if not cleaned:
        return None
    try:
        parsed = json.loads(cleaned)
    except json.JSONDecodeError:
        return None
    if not isinstance(parsed, dict):
        return None
    if "cid" in parsed and "payload" in parsed:
        try:
            parsed = decrypt_oauth_state(str(parsed["payload"]), settings.cipher_master)
        except Exception:
            return None
    tokens = parsed.get("tokens")
    if isinstance(tokens, dict):
        return dict(tokens)
    if "access_token" in parsed or "refresh_token" in parsed:
        return dict(parsed)
    return None


def resolve_action_oauth_token(
    state_json: str | None,
    settings: Settings,
) -> str:
    """Resolve, refresh (if offline refresh_token is present), and verify an action OAuth token."""
    tokens = extract_oauth_state_tokens(state_json, settings)
    if not tokens:
        raise PermissionError("No OAuth credentials found in state_json.")

    access_token = str(tokens.get("access_token") or "").strip()
    refresh_token = str(tokens.get("refresh_token") or "").strip()

    if access_token:
        try:
            return verify_oauth_bearer(f"Bearer {access_token}", settings.allowed_oauth_domains)[1]
        except PermissionError:
            if not refresh_token:
                raise

    if not refresh_token:
        raise PermissionError("OAuth state did not contain a valid access_token or refresh_token.")

    new_access = str(
        refresh_offline_access_token(refresh_token, settings).get("access_token") or ""
    ).strip()
    if not new_access:
        raise PermissionError("OAuth state did not contain a valid access_token or refresh_token.")
    return verify_oauth_bearer(f"Bearer {new_access}", settings.allowed_oauth_domains)[1]


def verify_action_token(auth_header: str | None, secret: str) -> bool:
    """Verify Looker Authorization: Token token='...' (raw secret or nonce/digest HMAC-SHA512)."""
    if not secret or not auth_header or not (match := _TOKEN_RE.match(auth_header.strip())):
        return False
    provided = match.group(1) or match.group(2) or ""
    if hmac.compare_digest(provided.encode("utf-8"), secret.encode("utf-8")):
        return True
    if "/" in provided:
        nonce, provided_digest = provided.split("/", 1)
        if nonce and provided_digest:
            expected = hmac.new(
                secret.encode("utf-8"), nonce.encode("utf-8"), hashlib.sha512
            ).hexdigest()
            return hmac.compare_digest(provided_digest.encode("utf-8"), expected.encode("utf-8"))
    return False


def verify_oauth_token_payload(token: str) -> dict[str, object]:
    """Verify a Google OAuth2 ID token or access token and return claims dict containing email."""
    if token.count(".") == 2:
        return dict(id_token.verify_oauth2_token(token, google_requests.Request()))
    query = urllib.parse.urlencode({"access_token": token})
    with urllib.request.urlopen(f"https://oauth2.googleapis.com/tokeninfo?{query}", timeout=5) as r:
        return dict(json.loads(r.read().decode("utf-8")))


def verify_oauth_bearer(
    auth_header: str | None,
    allowed_domains: frozenset[str],
) -> tuple[str, str]:
    """Validate Authorization: Bearer <token> and return (email, token)."""
    if not auth_header or not (match := _BEARER_RE.match(auth_header.strip())):
        raise PermissionError("Missing or invalid Bearer Authorization header.")
    token = match.group(1)
    try:
        claims = verify_oauth_token_payload(token)
    except Exception as exc:
        raise PermissionError(f"OAuth token verification failed: {exc}") from exc

    email = str(claims.get("email") or "").strip().lower()
    email_verified = str(claims.get("email_verified", "true")).lower() == "true"
    if not email or "@" not in email or not email_verified:
        raise PermissionError("OAuth token does not contain a verified email claim.")

    domain = email.rsplit("@", 1)[-1]
    if allowed_domains and domain not in allowed_domains:
        raise PermissionError(f"Email domain '{domain}' is not in ALLOWED_OAUTH_DOMAINS.")
    return email, token


def parse_and_validate_gcs_uri(
    uri: str,
    allowed_buckets: frozenset[str],
) -> tuple[str, str]:
    """Validate a gs://bucket[/prefix] URI against path traversal and ALLOWED_GCS_BUCKETS."""
    cleaned = (uri or "").strip()
    if not cleaned.startswith("gs://") or not (rest := cleaned[5:].strip("/")):
        raise ValueError(f"Invalid GCS URI: {uri!r}")

    parts = rest.split("/", 1)
    bucket, prefix = parts[0].strip(), (parts[1].strip("/") if len(parts) > 1 else "")

    if not bucket or ".." in bucket or "\\" in bucket:
        raise ValueError(f"Invalid GCS bucket name: {bucket!r}")

    if prefix and any(seg in ("", ".", "..") or "\\" in seg for seg in prefix.split("/")):
        raise ValueError(f"Invalid GCS path prefix: {prefix!r}")

    if allowed_buckets and bucket.lower() not in allowed_buckets:
        raise ValueError(f"Bucket '{bucket}' is not in ALLOWED_GCS_BUCKETS.")

    return bucket, prefix


def sanitize_relative_pptx_name(filename: str) -> str:
    """Validate a relative template path (e.g. 'deck.pptx' or 'team/deck.pptx')."""
    cleaned = (filename or "").strip().lstrip("/")
    if not cleaned or "\\" in cleaned:
        raise ValueError(f"Invalid template filename: {filename!r}")
    if any(seg in ("", ".", "..") for seg in cleaned.split("/")):
        raise ValueError(f"Path traversal rejected in filename: {filename!r}")
    return cleaned if cleaned.lower().endswith(".pptx") else f"{cleaned}.pptx"


def log_structured(
    event: str,
    *,
    trace_header: str | None = None,
    severity: str = "INFO",
    **fields: object,
) -> None:
    """Emit a single-line Cloud Logging JSON record to stdout without sensitive cell data."""
    record: dict[str, object] = {"severity": severity, "message": event, **fields}
    if trace_header and (trace_id := trace_header.split("/", 1)[0].strip()):
        record["logging.googleapis.com/trace"] = trace_id
    print(json.dumps(record, default=str), file=sys.stdout, flush=True)
