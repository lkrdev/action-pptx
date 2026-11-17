from __future__ import annotations

import re
from datetime import UTC, datetime

from google.cloud import storage
from google.oauth2.credentials import Credentials

from action_pptx.config import parse_and_validate_gcs_uri, sanitize_relative_pptx_name

PPTX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
_SAFE_TOKEN_RE = re.compile(r"[^a-zA-Z0-9_-]+")


def _get_storage_client(oauth_token: str | None = None) -> storage.Client:
    return (
        storage.Client(credentials=Credentials(token=oauth_token))
        if oauth_token
        else storage.Client()
    )


def _resolve_blob(
    gcs_uri: str,
    rel_name: str,
    allowed_buckets: frozenset[str],
    oauth_token: str | None = None,
) -> tuple[str, str, storage.Blob]:
    bucket_name, prefix = parse_and_validate_gcs_uri(gcs_uri, allowed_buckets)
    blob_path = f"{prefix}/{rel_name}" if prefix else rel_name
    blob = _get_storage_client(oauth_token).bucket(bucket_name).blob(blob_path)
    return bucket_name, blob_path, blob


def _iter_pptx_blobs(
    gcs_uri: str,
    allowed_buckets: frozenset[str],
    oauth_token: str | None = None,
) -> list[tuple[str, object]]:
    bucket_name, prefix = parse_and_validate_gcs_uri(gcs_uri, allowed_buckets)
    search_prefix = f"{prefix}/" if prefix else ""
    blobs = _get_storage_client(oauth_token).list_blobs(bucket_name, prefix=search_prefix)

    matched: list[tuple[str, object]] = []
    for blob in blobs:
        name = str(blob.name)
        if not name.lower().endswith(".pptx"):
            continue
        rel = (
            name[len(search_prefix) :] if search_prefix and name.startswith(search_prefix) else name
        )
        if rel and not rel.endswith("/"):
            matched.append((rel, blob))
    return matched


def list_pptx_templates(
    gcs_uri: str,
    allowed_buckets: frozenset[str],
    oauth_token: str | None = None,
) -> list[str]:
    """List relative *.pptx object names under gcs_uri."""
    return sorted(rel for rel, _ in _iter_pptx_blobs(gcs_uri, allowed_buckets, oauth_token))


def list_pptx_outputs(
    gcs_uri: str,
    allowed_buckets: frozenset[str],
    oauth_token: str | None = None,
) -> list[dict[str, object]]:
    """List relative *.pptx objects with size/updated metadata under gcs_uri, newest first."""
    items = [
        {
            "name": rel,
            "size": int(getattr(blob, "size", 0) or 0),
            "updated": u.isoformat() if (u := getattr(blob, "updated", None)) else "",
        }
        for rel, blob in _iter_pptx_blobs(gcs_uri, allowed_buckets, oauth_token)
    ]
    items.sort(key=lambda item: (str(item["updated"]), str(item["name"])), reverse=True)
    return items


def download_template_bytes(
    gcs_uri: str,
    filename: str,
    allowed_buckets: frozenset[str],
    oauth_token: str | None = None,
) -> bytes:
    """Download a .pptx template from gcs_uri."""
    _, _, blob = _resolve_blob(
        gcs_uri, sanitize_relative_pptx_name(filename), allowed_buckets, oauth_token
    )
    return bytes(blob.download_as_bytes())


def upload_template_bytes(
    gcs_uri: str,
    filename: str,
    data: bytes,
    allowed_buckets: frozenset[str],
    oauth_token: str | None = None,
) -> str:
    """Upload a validated .pptx template to gcs_uri and return its gs:// URI."""
    bucket_name, blob_path, blob = _resolve_blob(
        gcs_uri, sanitize_relative_pptx_name(filename), allowed_buckets, oauth_token
    )
    blob.upload_from_string(data, content_type=PPTX_CONTENT_TYPE)
    return f"gs://{bucket_name}/{blob_path}"


def _sanitize_output_prefix(prefix: str) -> str:
    raw = re.sub(r"\.pptx$", "", (prefix or "").strip(), flags=re.IGNORECASE)
    parts: list[str] = []
    for seg in raw.split("/"):
        trimmed = seg.strip()
        if trimmed in (".", ".."):
            raise ValueError(f"Path traversal rejected in output prefix: {prefix!r}")
        cleaned = _SAFE_TOKEN_RE.sub("_", trimmed).strip("_")
        if cleaned:
            parts.append(cleaned)
    return "/".join(parts) or "presentation"


def write_timestamped_output(
    gcs_uri: str,
    prefix: str,
    data: bytes,
    allowed_buckets: frozenset[str],
    oauth_token: str | None = None,
) -> str:
    """Write a generated .pptx to a timestamped object path under gcs_uri."""
    safe_prefix = _sanitize_output_prefix(prefix)
    ts = datetime.now(UTC).strftime("%Y%m%dT%H%M%SZ")
    filename = f"{safe_prefix}_{ts}.pptx"

    bucket_name, blob_path, blob = _resolve_blob(gcs_uri, filename, allowed_buckets, oauth_token)
    blob.upload_from_string(data, content_type=PPTX_CONTENT_TYPE)
    return f"gs://{bucket_name}/{blob_path}"
