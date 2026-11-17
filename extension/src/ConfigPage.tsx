import React, { useEffect, useState } from "react";

export interface ExtensionConfigData {
  cloudRunUrl: string;
  templatesGcsUri: string;
  outputGcsUri?: string;
  actionIntegrationId?: string;
  templatesUserAttributeName?: string;
  useOAuth: boolean;
  oauthClientId?: string;
  oauthGcsPassthrough?: boolean;
}

export function buildSecretDomainAllowlist(rawUrl?: string): string {
  const cleaned = (rawUrl || "").trim().replace(/\/+$/, "");
  if (!cleaned) return "";
  try {
    const parsed = new URL(cleaned);
    const base = `${parsed.origin}${parsed.pathname.replace(/\/+$/, "")}`;
    return `${base},${base}/,${base}/*`;
  } catch {
    return "";
  }
}

export async function copyToClipboardSafe(extensionSDK: any, text: string): Promise<void> {
  try {
    if (typeof extensionSDK?.clipboardWrite === "function") {
      await extensionSDK.clipboardWrite(text);
      return;
    }
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return;
    }
  } catch {
    // fall through to textarea copy inside sandboxed iframe
  }
  const ta = document.createElement("textarea");
  ta.value = text;
  ta.style.position = "fixed";
  ta.style.opacity = "0";
  document.body.appendChild(ta);
  ta.focus();
  ta.select();
  try {
    document.execCommand("copy");
  } finally {
    document.body.removeChild(ta);
  }
}

export interface LookerOAuthRedirectUris {
  actionRedirectUri: string;
  extensionUri: string;
  spartanUri: string;
}

export function resolveOAuthRedirectUris(
  cloudRunUrl?: string,
  lookerHostUrl?: string
): LookerOAuthRedirectUris {
  const winOrigin =
    typeof window !== "undefined" &&
    window.location?.origin &&
    window.location.origin !== "null"
      ? window.location.origin
      : "";
  const lookerBase = (lookerHostUrl?.trim() || winOrigin || "https://looker.example.com").replace(
    /\/+$/,
    ""
  );
  const cloudRunBase = (cloudRunUrl?.trim() || "https://action-pptx.a.run.app").replace(
    /\/+$/,
    ""
  );
  return {
    actionRedirectUri: `${cloudRunBase}/actions/pptx/oauth_redirect`,
    extensionUri: `${lookerBase}/extensions/oauth2_redirect`,
    spartanUri: `${lookerBase}/spartan/oauth2_redirect`,
  };
}

export const Svg: React.FC<{ size?: number; children: React.ReactNode }> = ({
  size = 15,
  children,
}) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2.2"
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden="true"
  >
    {children}
  </svg>
);

const IconArrowLeft = () => (
  <Svg>
    <path d="M19 12H5m7 7-7-7 7-7" />
  </Svg>
);
const IconSparkles = () => (
  <Svg>
    <path d="m12 3-1.9 5.8a2 2 0 0 1-1.3 1.3L3 12l5.8 1.9a2 2 0 0 1 1.3 1.3L12 21l1.9-5.8a2 2 0 0 1 1.3-1.3L21 12l-5.8-1.9a2 2 0 0 1-1.3-1.3Z" />
  </Svg>
);
const IconCopy = () => (
  <Svg size={13}>
    <rect x="9" y="9" width="13" height="13" rx="2" />
    <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
  </Svg>
);
const IconCheck = () => (
  <Svg size={14}>
    <path d="M20 6 9 17l-5-5" />
  </Svg>
);

interface ConfigPageProps {
  core40SDK: any;
  extensionSDK: any;
  initialConfig: ExtensionConfigData;
  onSave: (cfg: ExtensionConfigData) => void;
  onBack: () => void;
}

export const ConfigPage: React.FC<ConfigPageProps> = ({
  core40SDK,
  extensionSDK,
  initialConfig,
  onSave,
  onBack,
}) => {
  const [cfg, setCfg] = useState<ExtensionConfigData>({
    cloudRunUrl: initialConfig.cloudRunUrl || "",
    templatesGcsUri: initialConfig.templatesGcsUri || "",
    outputGcsUri: initialConfig.outputGcsUri || "",
    actionIntegrationId: initialConfig.actionIntegrationId || "",
    templatesUserAttributeName: initialConfig.templatesUserAttributeName || "",
    useOAuth: Boolean(initialConfig.useOAuth),
    oauthClientId: initialConfig.oauthClientId || "",
    oauthGcsPassthrough:
      initialConfig.oauthGcsPassthrough !== undefined
        ? Boolean(initialConfig.oauthGcsPassthrough)
        : true,
  });
  const [actionHubSecret, setActionHubSecret] = useState("");
  const [savedSecretDomain, setSavedSecretDomain] = useState("");
  const [copiedKey, setCopiedKey] = useState<string>("");
  const [status, setStatus] = useState<string>("");
  const patch = (next: Partial<ExtensionConfigData>) =>
    setCfg((prev) => ({ ...prev, ...next }));

  const hostUrl = extensionSDK?.lookerHostData?.hostUrl;
  const redirectUris = resolveOAuthRedirectUris(cfg.cloudRunUrl, hostUrl);
  const secretKeyName = String(
    extensionSDK?.createSecretKeyTag?.("action_hub_secret") || "action_hub_secret"
  ).replace(/[{}]/g, "");

  const expectedAllowlist = buildSecretDomainAllowlist(cfg.cloudRunUrl);
  const normalizeWl = (s: string) => s.replace(/\s+/g, "");
  const secretIsRequired =
    !cfg.useOAuth &&
    (!savedSecretDomain ||
      (Boolean(expectedAllowlist) &&
        normalizeWl(savedSecretDomain) !== normalizeWl(expectedAllowlist)));

  function copyUri(key: string, text: string) {
    void copyToClipboardSafe(extensionSDK, text);
    setCopiedKey(key);
    setTimeout(() => setCopiedKey((prev) => (prev === key ? "" : prev)), 1800);
  }

  async function autoDiscoverFromActionHub() {
    setStatus("Inspecting Looker Action Hub integrations...");
    try {
      const [integrations, hubs]: [any[], any[]] = await Promise.all([
        core40SDK.ok(core40SDK.all_integrations({})),
        core40SDK.ok(core40SDK.all_integration_hubs()).catch(() => []),
      ]);
      const hubUrlFor = (item: any): string => {
        const hid = String(item?.integration_hub_id || String(item?.id || "").split("::")[0] || "");
        const h = hubs.find((x) => String(x?.id || "") === hid);
        return String(h?.url || item?.url || "").trim();
      };
      const candidates = integrations.filter(
        (i) => i.id?.includes("action-pptx") || i.label?.toLowerCase().includes("action-pptx")
      );
      const preferredId = (cfg.actionIntegrationId || "").trim();
      const match =
        (preferredId && candidates.find((i) => String(i.id || "") === preferredId)) ||
        candidates.find((i) => i.enabled && !hubUrlFor(i).includes("trycloudflare")) ||
        candidates.find((i) => i.enabled) ||
        candidates[0];
      if (!match) {
        setStatus("No registered 'action-pptx' integration found on this Looker instance.");
        return;
      }
      const tmplParam = Array.isArray(match.params)
        ? match.params.find((p: any) => p.name === "templates_gcs_uri")
        : null;
      const outParam = Array.isArray(match.params)
        ? match.params.find((p: any) => p.name === "output_gcs_uri")
        : null;
      const rawUrl = hubUrlFor(match);
      const discoveredOrigin = rawUrl
        .replace(/\/actions\/pptx\/execute\/?$/, "")
        .replace(/\/+$/, "");
      let discoveredClientId = String(match.oauth_client_id || "").trim();
      let discoveredUsesOAuth = Boolean(match.uses_oauth || discoveredClientId);
      let discoveredPassthrough = false;

      if (discoveredOrigin && !discoveredClientId) {
        try {
          const healthResp = await fetch(`${discoveredOrigin}/health`);
          if (healthResp.ok) {
            const healthJson = await healthResp.json();
            if (healthJson?.oauth_client_id) {
              discoveredClientId = String(healthJson.oauth_client_id).trim();
              discoveredUsesOAuth = true;
            }
            if (typeof healthJson?.oauth_gcs_passthrough === "boolean") {
              discoveredPassthrough = healthJson.oauth_gcs_passthrough;
            }
          }
        } catch {
          // ignore health probe failure
        }
      }

      patch({
        ...(match.id ? { actionIntegrationId: String(match.id) } : {}),
        ...(discoveredOrigin ? { cloudRunUrl: discoveredOrigin } : {}),
        ...(tmplParam?.value ? { templatesGcsUri: String(tmplParam.value) } : {}),
        ...(outParam?.value ? { outputGcsUri: String(outParam.value) } : {}),
        ...(tmplParam?.user_attribute_name
          ? { templatesUserAttributeName: String(tmplParam.user_attribute_name) }
          : {}),
        useOAuth: discoveredUsesOAuth,
        oauthGcsPassthrough: discoveredPassthrough,
        oauthClientId: discoveredClientId,
      });
      const actionDesc = match.id ? `Action "${match.id}"` : "registered Action Hub";
      setStatus(
        discoveredClientId
          ? `Pre-filled Cloud Run URL, templates_gcs_uri, and shared Google OAuth Client ID from ${actionDesc}.`
          : `Pre-filled Cloud Run URL and templates_gcs_uri from ${actionDesc}.`
      );
    } catch (err: any) {
      setStatus(`Auto-discovery failed (Admin permissions required): ${err?.message || err}`);
    }
  }

  useEffect(() => {
    void core40SDK
      .ok(core40SDK.all_user_attributes({}))
      .then((attrs: any[]) => {
        const existing = (Array.isArray(attrs) ? attrs : []).find((a) => a?.name === secretKeyName);
        if (existing?.value_is_hidden && existing?.hidden_value_domain_whitelist) {
          setSavedSecretDomain(String(existing.hidden_value_domain_whitelist));
        }
      })
      .catch(() => {});

    if (!initialConfig.cloudRunUrl) {
      void autoDiscoverFromActionHub();
    }
  }, []);

  async function handleSave(e: React.FormEvent) {
    e.preventDefault();
    setStatus("Saving configuration...");
    try {
      const nextConfig: ExtensionConfigData = {
        cloudRunUrl: cfg.cloudRunUrl.trim().replace(/\/+$/, ""),
        templatesGcsUri: cfg.templatesGcsUri.trim(),
        outputGcsUri: cfg.outputGcsUri?.trim() || undefined,
        actionIntegrationId: cfg.actionIntegrationId?.trim() || undefined,
        templatesUserAttributeName: cfg.templatesUserAttributeName?.trim() || undefined,
        useOAuth: cfg.useOAuth,
        oauthClientId: cfg.oauthClientId?.trim() || undefined,
        oauthGcsPassthrough: cfg.oauthGcsPassthrough,
      };
      if (!nextConfig.cloudRunUrl) {
        throw new Error("Cloud Run Service Base URL is required.");
      }
      if (!nextConfig.templatesGcsUri && !nextConfig.templatesUserAttributeName) {
        throw new Error("Templates GCS URI is required.");
      }
      const domainAllowlist = buildSecretDomainAllowlist(nextConfig.cloudRunUrl);
      if (!domainAllowlist) {
        throw new Error("Cloud Run Service Base URL must be a valid http(s) URL.");
      }
      if (secretIsRequired && !actionHubSecret.trim()) {
        throw new Error(
          savedSecretDomain && normalizeWl(savedSecretDomain) !== normalizeWl(domainAllowlist)
            ? `Service URL changed (${savedSecretDomain} → ${domainAllowlist}); re-enter Action Hub Secret to recreate the hidden attribute.`
            : "Action Hub Secret is required when Google OAuth is disabled."
        );
      }

      if (actionHubSecret.trim()) {
        const attrs: any[] = await core40SDK.ok(core40SDK.all_user_attributes({}));
        const existing = attrs.find((a) => a.name === secretKeyName);
        if (existing?.id) {
          await core40SDK.ok(core40SDK.delete_user_attribute(String(existing.id)));
        }
        await core40SDK.ok(
          core40SDK.create_user_attribute({
            name: secretKeyName,
            label: `${extensionSDK?.lookerHostData?.extensionId || "action-pptx"} Action Hub Secret`,
            type: "string",
            value_is_hidden: true,
            hidden_value_domain_whitelist: domainAllowlist,
            user_can_edit: false,
            user_can_view: true,
            default_value: actionHubSecret.trim(),
          })
        );
        setSavedSecretDomain(domainAllowlist);
        setActionHubSecret("");
      }

      await extensionSDK.saveContextData(nextConfig);

      onSave(nextConfig);
      setStatus("Configuration saved to Extension Context Data.");
    } catch (err: any) {
      setStatus(`Failed to save configuration: ${err?.message || err}`);
    }
  }

  const inputStyle: React.CSSProperties = {
    width: "100%",
    padding: "9px 12px",
    borderRadius: 8,
    background: "#0f172a",
    color: "#f8fafc",
    border: "1px solid #334155",
    fontSize: 13,
    boxSizing: "border-box",
  };

  return (
    <div
      style={{
        minHeight: "100vh",
        background: "#0b1120",
        color: "#f8fafc",
        fontFamily: "Inter, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
        padding: "28px 32px",
      }}
    >
      <div style={{ maxWidth: 780, margin: "0 auto" }}>
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            borderBottom: "1px solid #1e293b",
            paddingBottom: 18,
            marginBottom: 20,
          }}
        >
          <div>
            <div
              style={{
                display: "inline-block",
                fontSize: 11,
                fontWeight: 700,
                letterSpacing: "0.08em",
                color: "#38bdf8",
                background: "#1e293b",
                padding: "4px 10px",
                borderRadius: 999,
                marginBottom: 8,
              }}
            >
              ADMIN SETTINGS  •  /CONFIG
            </div>
            <h2 style={{ margin: 0, fontSize: 22, fontWeight: 700 }}>
              action-pptx Admin Configuration (/config)
            </h2>
          </div>
          {initialConfig.cloudRunUrl && (
            <button
              id="btn-back"
              type="button"
              onClick={onBack}
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 6,
                background: "#1e293b",
                color: "#e2e8f0",
                border: "1px solid #334155",
                borderRadius: 8,
                padding: "8px 14px",
                fontSize: 13,
                fontWeight: 600,
                cursor: "pointer",
              }}
            >
              <IconArrowLeft />
              <span>Back</span>
            </button>
          )}
        </div>

        <p style={{ color: "#94a3b8", fontSize: 13, lineHeight: 1.5, marginBottom: 16 }}>
          Configure the Cloud Run service URL, default <code>templates_gcs_uri</code>, and shared
          Google Web OAuth 2.0 Client ID so both the Looker Custom Action (with offline refresh
          tokens &amp; GCS passthrough) and this Extension share a single OAuth client.
        </p>

        <button
          id="btn-autodiscover"
          type="button"
          onClick={autoDiscoverFromActionHub}
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 8,
            marginBottom: 18,
            background: "#0284c7",
            color: "#ffffff",
            border: "none",
            borderRadius: 8,
            padding: "9px 16px",
            fontSize: 13,
            fontWeight: 600,
            cursor: "pointer",
          }}
        >
          <IconSparkles />
          <span>Auto-Discover from Registered Action Hub</span>
        </button>

        <form
          onSubmit={handleSave}
          style={{
            display: "grid",
            gap: 14,
            background: "#111827",
            border: "1px solid #1e293b",
            borderRadius: 12,
            padding: 20,
          }}
        >
          <label>
            <div style={{ fontSize: 12, fontWeight: 600, color: "#cbd5e1", marginBottom: 6 }}>
              <span style={{ color: "#ef4444", fontWeight: 700, marginRight: 4 }}>*</span>
              Cloud Run Service Base URL
            </div>
            <input
              type="url"
              required
              value={cfg.cloudRunUrl}
              onChange={(e) => patch({ cloudRunUrl: e.target.value })}
              placeholder="https://action-pptx-xyz.a.run.app"
              style={inputStyle}
            />
          </label>

          <label>
            <div style={{ fontSize: 12, fontWeight: 600, color: "#cbd5e1", marginBottom: 6 }}>
              <span style={{ color: "#ef4444", fontWeight: 700, marginRight: 4 }}>*</span>
              Templates GCS URI (gs://bucket/prefix)
            </div>
            <input
              type="text"
              required={!cfg.templatesUserAttributeName?.trim()}
              value={cfg.templatesGcsUri}
              onChange={(e) => patch({ templatesGcsUri: e.target.value })}
              placeholder="gs://my-templates-bucket/templates"
              style={inputStyle}
            />
          </label>

          <label>
            <div style={{ fontSize: 12, fontWeight: 600, color: "#cbd5e1", marginBottom: 6 }}>
              Output GCS URI (gs://bucket/prefix — enables Report Viewer)
            </div>
            <input
              id="input-output-gcs-uri"
              type="text"
              value={cfg.outputGcsUri || ""}
              onChange={(e) => patch({ outputGcsUri: e.target.value })}
              placeholder="gs://my-output-bucket/output"
              style={inputStyle}
            />
          </label>

          <label>
            <div style={{ fontSize: 12, fontWeight: 600, color: "#cbd5e1", marginBottom: 6 }}>
              Action Integration ID (auto-discovered, e.g. <code>28::action-pptx</code>)
            </div>
            <input
              id="input-action-integration-id"
              type="text"
              value={cfg.actionIntegrationId || ""}
              onChange={(e) => patch({ actionIntegrationId: e.target.value })}
              placeholder="28::action-pptx"
              style={inputStyle}
            />
          </label>

          <label>
            <div style={{ fontSize: 12, fontWeight: 600, color: "#cbd5e1", marginBottom: 6 }}>
              Or Looker User Attribute Name for Templates GCS URI (optional per-team scoping)
            </div>
            <input
              type="text"
              value={cfg.templatesUserAttributeName || ""}
              onChange={(e) => patch({ templatesUserAttributeName: e.target.value })}
              placeholder="templates_gcs_uri"
              style={inputStyle}
            />
          </label>

          <label>
            <div style={{ fontSize: 12, fontWeight: 600, color: "#cbd5e1", marginBottom: 6 }}>
              {!cfg.useOAuth && (
                <span style={{ color: "#ef4444", fontWeight: 700, marginRight: 4 }}>*</span>
              )}
              Action Hub Secret (hidden user attribute <code>{secretKeyName}</code>){" "}
              <span style={{ color: secretIsRequired ? "#f87171" : "#34d399" }}>
                {secretIsRequired
                  ? savedSecretDomain
                    ? "• Required (URL changed — recreates hidden attribute)"
                    : "• Required"
                  : `• Saved (${savedSecretDomain})`}
              </span>
            </div>
            <input
              type="password"
              required={secretIsRequired}
              value={actionHubSecret}
              onChange={(e) => setActionHubSecret(e.target.value)}
              placeholder={
                secretIsRequired
                  ? "Enter Action Hub Secret (stored hidden & domain-allowlisted)"
                  : "Leave blank to keep existing secret, or enter a new one to replace it"
              }
              style={inputStyle}
            />
          </label>

          <label style={{ display: "flex", gap: 10, alignItems: "center", fontSize: 13 }}>
            <input
              id="chk-use-oauth"
              type="checkbox"
              checked={cfg.useOAuth}
              onChange={(e) => patch({ useOAuth: e.target.checked })}
            />
            <span style={{ color: "#e2e8f0", fontWeight: 600 }}>
              Use Shared Google OAuth 2.0 (Browser Sign-In + GCS Token Passthrough)
            </span>
          </label>

          {cfg.useOAuth && (
            <div
              style={{
                display: "grid",
                gap: 12,
                background: "#0f172a",
                border: "1px solid #334155",
                borderRadius: 10,
                padding: 16,
              }}
            >
              <label>
                <div style={{ fontSize: 12, fontWeight: 600, color: "#38bdf8", marginBottom: 4 }}>
                  Shared Google Web OAuth 2.0 Client ID (same as GOOGLE_OAUTH_CLIENT_ID on Cloud Run)
                </div>
                <input
                  id="input-oauth-client-id"
                  type="text"
                  value={cfg.oauthClientId || ""}
                  onChange={(e) => patch({ oauthClientId: e.target.value })}
                  placeholder="123456789-abc.apps.googleusercontent.com"
                  style={{ ...inputStyle, background: "#111827" }}
                />
              </label>

              <label style={{ display: "flex", gap: 10, alignItems: "center", fontSize: 13 }}>
                <input
                  id="chk-oauth-passthrough"
                  type="checkbox"
                  checked={Boolean(cfg.oauthGcsPassthrough)}
                  onChange={(e) => patch({ oauthGcsPassthrough: e.target.checked })}
                />
                <span style={{ color: "#e2e8f0" }}>
                  Pass through user OAuth token to GCS (<code>devstorage.read_write</code> scope)
                </span>
              </label>

              <div
                style={{
                  background: "#111827",
                  border: "1px solid #1e293b",
                  borderRadius: 8,
                  padding: 12,
                  display: "grid",
                  gap: 8,
                  fontSize: 12,
                }}
              >
                <div style={{ fontWeight: 700, color: "#f8fafc" }}>
                  Authorized Redirect URIs for Shared GCP Web OAuth 2.0 Client
                </div>
                {[
                  {
                    key: "action",
                    label: "1. Custom Action (Offline + Passthrough):",
                    uri: redirectUris.actionRedirectUri,
                  },
                  {
                    key: "ext",
                    label: "2. Extension (Standard):",
                    uri: redirectUris.extensionUri,
                  },
                  {
                    key: "spartan",
                    label: "3. Extension (Full Screen / Spartan):",
                    uri: redirectUris.spartanUri,
                  },
                ].map((item) => (
                  <div
                    key={item.key}
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                      gap: 8,
                      background: "#0b1120",
                      padding: "6px 10px",
                      borderRadius: 6,
                      border: "1px solid #1e293b",
                    }}
                  >
                    <span style={{ color: "#cbd5e1", overflow: "hidden", textOverflow: "ellipsis" }}>
                      <strong>{item.label}</strong>{" "}
                      <code style={{ color: "#38bdf8" }}>{item.uri}</code>
                    </span>
                    <button
                      type="button"
                      onClick={() => copyUri(item.key, item.uri)}
                      style={{
                        display: "inline-flex",
                        alignItems: "center",
                        gap: 5,
                        background: "#1e293b",
                        color: "#e2e8f0",
                        border: "1px solid #334155",
                        borderRadius: 6,
                        padding: "4px 8px",
                        fontSize: 11,
                        fontWeight: 600,
                        cursor: "pointer",
                        whiteSpace: "nowrap",
                      }}
                    >
                      {copiedKey === item.key ? <IconCheck /> : <IconCopy />}
                      <span>{copiedKey === item.key ? "Copied" : "Copy URI"}</span>
                    </button>
                  </div>
                ))}
                <div style={{ fontSize: 11, color: "#94a3b8" }}>
                  Add all 3 URIs under <strong>Authorized redirect URIs</strong> on your single GCP
                  Web OAuth 2.0 Client ID so both the Custom Action (offline <code>refresh_token</code>
                  ) and the Extension (<code>oauth2Authenticate</code>) work without{" "}
                  <code>redirect_uri_mismatch</code>.
                </div>
              </div>
            </div>
          )}

          <button
            type="submit"
            style={{
              display: "inline-flex",
              alignItems: "center",
              justifyContent: "center",
              gap: 8,
              padding: "10px 16px",
              background: "#059669",
              color: "#ffffff",
              border: "none",
              borderRadius: 8,
              fontSize: 13,
              fontWeight: 600,
              cursor: "pointer",
            }}
          >
            <IconCheck />
            <span>Save Configuration</span>
          </button>
        </form>

        {status && (
          <div
            style={{
              marginTop: 14,
              background: "#0f172a",
              borderLeft: "4px solid #38bdf8",
              padding: "10px 14px",
              borderRadius: 8,
              color: "#38bdf8",
              fontSize: 13,
              fontWeight: 600,
            }}
          >
            {status}
          </div>
        )}
      </div>
    </div>
  );
};
