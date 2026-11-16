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
import secrets
import time

import looker_sdk
from looker_sdk import models40


def main() -> None:
    sdk = looker_sdk.init40()
    url = os.environ.get("SERVICE_URL")
    token = os.environ.get("API_KEY_TOKEN")
    if not url or not token:
        raise RuntimeError("SERVICE_URL and API_KEY_TOKEN environment variables are required.")

    if "/" not in token:
        nonce = secrets.token_bytes(32).hex()
        digest = hmac.new(token.encode("utf-8"), nonce.encode("utf-8"), hashlib.sha512).hexdigest()
        token = f"{nonce}/{digest}"

    print(f"Registering Integration Hub {url} in Looker...")
    body = models40.WriteIntegrationHub(url=url, authorization_token=token)
    try:
        hub = sdk.create_integration_hub(body=body)
        print(f"Successfully registered Integration Hub (ID: {hub.id})")
    except Exception as exc:
        existing = next((h for h in sdk.all_integration_hubs() if h.url == url and h.id), None)
        if not existing or not existing.id:
            raise RuntimeError(f"Failed to find or create the Integration Hub: {exc}") from exc
        hub = sdk.update_integration_hub(integration_hub_id=existing.id, body=body)
        print(f"Updated existing Integration Hub (ID: {hub.id})")

    if hub.id and hub.legal_agreement_required and not hub.legal_agreement_signed:
        hub = sdk.accept_integration_hub_legal_agreement(integration_hub_id=hub.id)

    print("Enabling action-pptx integration...")
    integration = None
    for attempt in range(6):
        integration = next(
            (
                i
                for i in sdk.all_integrations(integration_hub_id=hub.id or "")
                if i.id and i.id.split("::")[-1] == "action-pptx"
            ),
            None,
        )
        if integration or attempt == 5:
            break
        print(f"Integration not found on hub {hub.id} yet; retrying in 3s ({attempt + 1}/6)...")
        time.sleep(3)

    if not integration or not integration.id:
        raise RuntimeError("Could not find action-pptx integration on the registered hub.")

    params = [
        models40.WriteIntegrationParam(name=k, value=v)
        for k, env_var in (
            ("templates_gcs_uri", "TEMPLATES_GCS_URI"),
            ("output_gcs_uri", "OUTPUT_GCS_URI"),
        )
        if (v := os.environ.get(env_var, ""))
    ]
    sdk.update_integration(
        integration_id=integration.id,
        body=models40.WriteIntegration(enabled=True, params=params or None),
    )
    print("Successfully enabled and configured action-pptx integration!")


if __name__ == "__main__":
    main()
