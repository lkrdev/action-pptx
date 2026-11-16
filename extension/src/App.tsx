import React, { useEffect, useRef, useState } from "react";
import { init as initPptxPreview } from "pptx-preview";
import { ConfigPage, ExtensionConfigData, Svg } from "./ConfigPage";
import {
  DashboardTileSchema,
  fetchDashboardTileCsvs,
  fetchDashboardTileSchemas,
} from "./dashboardSchema";

interface AppProps {
  core40SDK: any;
  extensionSDK: any;
}

export const OAUTH_STORAGE_KEY = "action_pptx_oauth_session";
export const BRAND_OPTIONS = ["Levi's", "Calvin Klein", "Columbia", "Allegra K", "Default"] as const;

export function getEmbedThemeName(isDark: boolean, brand?: string): string {
  const suffix = isDark ? "Dark" : "Light";
  const cleaned = (brand || "").trim();
  if (!cleaned || cleaned === "Default" || cleaned === "%" || cleaned.toLowerCase() === "all") {
    return `Embed_Demo_${suffix}`;
  }
  const sanitized = cleaned.replace(/[^a-zA-Z0-9\s]/g, "").trim().replace(/\s+/g, "_");
  return `${sanitized}_${suffix}`;
}

interface UiPalette {
  bg: string;
  cardBg: string;
  inputBg: string;
  border: string;
  textPrimary: string;
  textSecondary: string;
  textMuted: string;
  accent: string;
  accentContrast: string;
  canvasBg: string;
}

function getUiPalette(brand: string, isDark: boolean): UiPalette {
  const brandAccents: Record<string, { dark: string; light: string }> = {
    "Levi's": { dark: "#e11d48", light: "#c41230" },
    "Calvin Klein": { dark: "#94a3b8", light: "#334155" },
    Columbia: { dark: "#00a896", light: "#0066b2" },
    "Allegra K": { dark: "#f43f5e", light: "#be123c" },
    Default: { dark: "#38bdf8", light: "#0b57d0" },
  };
  const pair = brandAccents[brand] || brandAccents.Default;
  const accent = isDark ? pair.dark : pair.light;

  if (isDark) {
    return {
      bg: "#0b1120",
      cardBg: "#111827",
      inputBg: "#0f172a",
      border: "#1e293b",
      textPrimary: "#f8fafc",
      textSecondary: "#e2e8f0",
      textMuted: "#94a3b8",
      accent,
      accentContrast: "#ffffff",
      canvasBg: "#090d16",
    };
  }
  return {
    bg: "#f0f4f9",
    cardBg: "#ffffff",
    inputBg: "#f8fafc",
    border: "#cbd5e1",
    textPrimary: "#0f172a",
    textSecondary: "#334155",
    textMuted: "#475569",
    accent,
    accentContrast: "#ffffff",
    canvasBg: "#e2e8f0",
  };
}

const IconSlides: React.FC<{ size?: number }> = ({ size = 24 }) => (
  <Svg size={size}>
    <rect x="2" y="3" width="20" height="14" rx="2" />
    <path d="M8 21h8m-4-4v4M7 12V8m5 4v-2m5 2V6" />
  </Svg>
);
const IconGoogle = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
    <path d="M21.8 12.2c0-.7-.1-1.4-.2-2H12v3.8h5.5a4.7 4.7 0 0 1-2 3.1v2.6h3.3c1.9-1.8 3-4.4 3-7.5ZM12 22c2.7 0 5-.9 6.7-2.4l-3.3-2.6c-.9.6-2.1 1-3.4 1-2.6 0-4.8-1.8-5.6-4.1H3v2.6A10 10 0 0 0 12 22ZM6.4 13.9a6 6 0 0 1 0-3.8V7.5H3a10 10 0 0 0 0 9l3.4-2.6ZM12 5.9c1.5 0 2.8.5 3.8 1.5l2.9-2.9A10 10 0 0 0 3 7.5l3.4 2.6c.8-2.3 3-4.2 5.6-4.2Z" />
  </svg>
);
const IconCheckCircle = () => (
  <Svg size={14}>
    <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
    <path d="m9 11 3 3L22 4" />
  </Svg>
);
const IconAlert = () => (
  <Svg size={14}>
    <circle cx="12" cy="12" r="10" />
    <path d="M12 8v4m0 4h.01" />
  </Svg>
);
const IconLogOut = () => (
  <Svg size={14}>
    <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4m7 14 5-5-5-5m5 5H9" />
  </Svg>
);
const IconPalette = () => (
  <Svg size={14}>
    <circle cx="13.5" cy="6.5" r=".5" fill="currentColor" />
    <circle cx="17.5" cy="10.5" r=".5" fill="currentColor" />
    <circle cx="8.5" cy="7.5" r=".5" fill="currentColor" />
    <circle cx="6.5" cy="12.5" r=".5" fill="currentColor" />
    <path d="M12 2C6.5 2 2 6.5 2 12s4.5 10 10 10c.93 0 1.65-.75 1.65-1.69 0-.43-.18-.83-.44-1.12-.29-.29-.44-.65-.44-1.13a1.64 1.64 0 0 1 1.67-1.67h1.99c3.05 0 5.55-2.5 5.55-5.55C21.97 6.01 17.46 2 12 2Z" />
  </Svg>
);
const IconSun = () => (
  <Svg>
    <circle cx="12" cy="12" r="4" />
    <path d="M12 2v2m0 16v2M4.93 4.93l1.41 1.41m11.32 11.32 1.41 1.41M2 12h2m16 0h2M6.34 17.66l-1.41 1.41M19.07 4.93l-1.41 1.41" />
  </Svg>
);
const IconMoon = () => (
  <Svg>
    <path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z" />
  </Svg>
);
const IconPlay = () => (
  <Svg>
    <polygon points="6 3 20 12 6 21 6 3" />
  </Svg>
);
const IconCheck = () => (
  <Svg>
    <path d="M20 6 9 17l-5-5" />
  </Svg>
);
const IconDownload = () => (
  <Svg>
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4m4-5 5 5 5-5m-5 5V3" />
  </Svg>
);
const IconUpload = () => (
  <Svg>
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4m14-7-5-5-5 5m5-5v12" />
  </Svg>
);
const IconClose = () => (
  <Svg>
    <path d="M18 6 6 18M6 6l12 12" />
  </Svg>
);
const IconArrowLeft = () => (
  <Svg size={16}>
    <path d="m15 18-6-6 6-6" />
  </Svg>
);
const IconArrowRight = () => (
  <Svg size={16}>
    <path d="m9 18 6-6-6-6" />
  </Svg>
);
const IconSkipEnd = () => (
  <Svg size={16}>
    <polygon points="5 4 15 12 5 20 5 4" />
    <line x1="19" y1="5" x2="19" y2="19" />
  </Svg>
);

interface StoredOAuthSession {
  token: string;
  expiresAt: number;
  clientId?: string;
}

function parseStoredOAuthSession(
  raw: string | null | undefined,
  expectedClientId?: string
): StoredOAuthSession | null {
  if (!raw?.trim()) return null;
  try {
    const parsed = JSON.parse(raw) as Partial<StoredOAuthSession> | null;
    const token = typeof parsed?.token === "string" ? parsed.token.trim() : "";
    const expiresAt = Number(parsed?.expiresAt);
    if (!token || !Number.isFinite(expiresAt) || expiresAt <= Date.now()) return null;
    if (expectedClientId && parsed?.clientId && parsed.clientId !== expectedClientId) {
      return null;
    }
    return { token, expiresAt, clientId: parsed?.clientId };
  } catch {
    return null;
  }
}

function persistOAuthSession(extensionSDK: any, session: StoredOAuthSession): void {
  const serialized = JSON.stringify(session);
  try {
    window.localStorage?.setItem(OAUTH_STORAGE_KEY, serialized);
  } catch {
    // ignore storage errors
  }
  try {
    void extensionSDK?.localStorageSetItem?.(OAUTH_STORAGE_KEY, serialized);
  } catch {
    // ignore extension storage errors
  }
}

function clearOAuthSession(extensionSDK: any): void {
  try {
    window.localStorage?.removeItem(OAUTH_STORAGE_KEY);
  } catch {
    // ignore
  }
  try {
    void extensionSDK?.localStorageRemoveItem?.(OAUTH_STORAGE_KEY);
    void extensionSDK?.localStorageSetItem?.(OAUTH_STORAGE_KEY, "");
  } catch {
    // ignore
  }
}

function loginWithGoogleOAuth(
  extensionSDK: any,
  clientId: string,
  passthrough: boolean
): Promise<string> {
  if (!extensionSDK?.oauth2Authenticate) {
    return Promise.reject(new Error("Looker Extension SDK oauth2Authenticate is not available."));
  }
  const scope = passthrough
    ? "openid email https://www.googleapis.com/auth/devstorage.read_write"
    : "openid email";
  return Promise.resolve(
    extensionSDK.oauth2Authenticate(
      "https://accounts.google.com/o/oauth2/v2/auth",
      { client_id: clientId.trim(), scope, response_type: "token" },
      "GET"
    )
  ).then((res: any) => {
    const rawToken = String(res?.access_token || "").trim();
    if (!rawToken) {
      throw new Error("Google OAuth response did not include an access_token.");
    }
    const expiresAt = Date.now() + Math.max(60, (Number(res?.expires_in) || 3600) - 60) * 1000;
    persistOAuthSession(extensionSDK, {
      token: rawToken,
      expiresAt,
      clientId: clientId.trim(),
    });
    return rawToken;
  });
}

function readFileAsBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || "").split(",")[1] || "");
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function base64ToArrayBuffer(b64: string): ArrayBuffer {
  const binary = window.atob(b64.replace(/\s+/g, ""));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

export const App: React.FC<AppProps> = ({ core40SDK, extensionSDK }) => {
  const [route, setRoute] = useState<string>(() => {
    const hash = window.location.hash.replace(/^#/, "");
    if (hash) return hash;
    return window.location.pathname.endsWith("/config") ? "/config" : "/";
  });
  const [config, setConfig] = useState<ExtensionConfigData>(() => {
    const ctx = extensionSDK?.getContextData?.() as ExtensionConfigData | undefined;
    return {
      cloudRunUrl: ctx?.cloudRunUrl || "",
      templatesGcsUri: ctx?.templatesGcsUri || "",
      templatesUserAttributeName: ctx?.templatesUserAttributeName,
      useOAuth: Boolean(ctx?.useOAuth),
      oauthClientId: ctx?.oauthClientId,
      oauthGcsPassthrough: ctx?.oauthGcsPassthrough ?? true,
    };
  });
  const [brand, setBrand] = useState<string>("Levi's");
  const [isDark, setIsDark] = useState<boolean>(true);
  const [dashboards, setDashboards] = useState<Array<{ id: string; title: string }>>([]);
  const [selectedDashboardId, setSelectedDashboardId] = useState<string>("");
  const [tiles, setTiles] = useState<DashboardTileSchema[]>([]);
  const [fileName, setFileName] = useState<string>("");
  const [pptxBase64, setPptxBase64] = useState<string>("");
  const [report, setReport] = useState<any>(null);
  const [mergedPptxBase64, setMergedPptxBase64] = useState<string>("");
  const [currentSlide, setCurrentSlide] = useState<number>(0);
  const [totalSlides, setTotalSlides] = useState<number>(0);
  const [status, setStatus] = useState<string>("");
  const [oauthToken, setOauthToken] = useState<string>("");
  const [oauthSigningIn, setOauthSigningIn] = useState<boolean>(false);
  const previewContainerRef = useRef<HTMLDivElement | null>(null);
  const previewerRef = useRef<any>(null);
  const desiredSlideRef = useRef<number>(0);

  const ui = getUiPalette(brand, isDark);
  const activeEmbedTheme = getEmbedThemeName(isDark, brand);

  useEffect(() => {
    const initialClientId = config.oauthClientId || "";
    const localSession = parseStoredOAuthSession(
      window.localStorage?.getItem(OAUTH_STORAGE_KEY),
      initialClientId
    );
    if (localSession?.token) {
      setOauthToken(localSession.token);
    } else {
      void Promise.resolve(extensionSDK?.localStorageGetItem?.(OAUTH_STORAGE_KEY))
        .then((extRaw: string | null) => {
          const extSession = parseStoredOAuthSession(extRaw, initialClientId);
          if (extSession?.token) setOauthToken(extSession.token);
        })
        .catch(() => {});
    }

    // Automatically detect user's Looker `brand` user attribute (like looker-embed-demo)
    void Promise.resolve(extensionSDK?.userAttributeGetItem?.("brand"))
      .then((attrBrand: string | null) => {
        const cleaned = (attrBrand || "").trim();
        if (cleaned && BRAND_OPTIONS.includes(cleaned as any)) {
          setBrand(cleaned);
        }
      })
      .catch(() => {});

    const onStorage = (e: StorageEvent) => {
      if (e.key && e.key !== OAUTH_STORAGE_KEY) return;
      const next = parseStoredOAuthSession(e.newValue, initialClientId);
      setOauthToken(next?.token || "");
    };
    const onHashChange = () => {
      setRoute(window.location.hash.replace(/^#/, "") || "/");
    };
    window.addEventListener("storage", onStorage);
    window.addEventListener("hashchange", onHashChange);

    void (async () => {
      try {
        const list = await core40SDK.ok(core40SDK.search_dashboards({ limit: 200 }));
        setDashboards(
          (Array.isArray(list) ? list : [])
            .filter((d: any) => d?.id && d?.title)
            .map((d: any) => ({ id: String(d.id), title: String(d.title) }))
        );
      } catch (err: any) {
        setStatus(`Could not load dashboards: ${err?.message || err}`);
      }
    })();

    return () => {
      window.removeEventListener("storage", onStorage);
      window.removeEventListener("hashchange", onHashChange);
    };
  }, []);

  useEffect(() => {
    if (!mergedPptxBase64 || !previewContainerRef.current) return;
    const container = previewContainerRef.current;
    container.innerHTML = "";
    const previewer = initPptxPreview(container, {
      width: 860,
      height: 484,
      mode: "list",
    });
    previewerRef.current = previewer;
    void previewer.load(base64ToArrayBuffer(mergedPptxBase64)).then(() => {
      const count = Number(previewer.slideCount) || 1;
      setTotalSlides(count);
      const targetIdx = Math.max(0, Math.min(count - 1, desiredSlideRef.current));
      setCurrentSlide(targetIdx);
      previewer.renderSingleSlide(targetIdx);
    });
  }, [mergedPptxBase64]);

  function goToSlide(index: number) {
    if (!previewerRef.current || totalSlides <= 0) return;
    const clamped = Math.max(0, Math.min(totalSlides - 1, index));
    desiredSlideRef.current = clamped;
    previewerRef.current.renderSingleSlide(clamped);
    setCurrentSlide(clamped);
  }

  function handleEndPreview() {
    if (previewerRef.current?.destroy) {
      previewerRef.current.destroy();
    }
    previewerRef.current = null;
    desiredSlideRef.current = 0;
    setMergedPptxBase64("");
    setCurrentSlide(0);
    setTotalSlides(0);
  }

  async function handleOAuthLogin(): Promise<string> {
    if (!config.oauthClientId) {
      setRoute("/config");
      throw new Error("OAuth Client ID is not configured.");
    }
    setOauthSigningIn(true);
    setStatus("");
    try {
      const token = await loginWithGoogleOAuth(
        extensionSDK,
        config.oauthClientId,
        config.oauthGcsPassthrough !== false
      );
      setOauthToken(token);
      setStatus("");
      return token;
    } catch (err: any) {
      setStatus(`Sign-in failed: ${err?.message || err}`);
      throw err;
    } finally {
      setOauthSigningIn(false);
    }
  }

  function handleOAuthLogout() {
    clearOAuthSession(extensionSDK);
    setOauthToken("");
    setStatus("");
  }

  async function handleSelectDashboard(dashId: string) {
    setSelectedDashboardId(dashId);
    setReport(null);
    handleEndPreview();
    if (!dashId) {
      setTiles([]);
      return;
    }
    setStatus("Loading dashboard tiles...");
    try {
      const schema = await fetchDashboardTileSchemas(core40SDK, dashId);
      setTiles(schema);
      setStatus("");
    } catch (err: any) {
      setStatus(`Could not load dashboard tiles: ${err?.message || err}`);
    }
  }

  async function handleFileSelect(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    setPptxBase64(await readFileAsBase64(file));
    setReport(null);
    handleEndPreview();
  }

  async function callCloudRun(path: string, body: Record<string, any>): Promise<any> {
    if (!config.cloudRunUrl) {
      setRoute("/config");
      throw new Error("Service is not configured yet.");
    }
    const url = `${config.cloudRunUrl}${path}`;
    if (config.useOAuth) {
      let token = oauthToken;
      if (!token) {
        token = await handleOAuthLogin();
      }
      const resp = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
        body: JSON.stringify(body),
      });
      if (resp.status === 401 || resp.status === 403) {
        clearOAuthSession(extensionSDK);
        setOauthToken("");
        const errBody = await resp.json().catch(() => ({}));
        throw new Error(errBody?.detail || `Session expired (${resp.status}). Please sign in again.`);
      }
      return await resp.json();
    }

    const secretTag = extensionSDK.createSecretKeyTag("action_hub_secret");
    const resp = await extensionSDK.serverProxy(url, {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Token token="${secretTag}"` },
      body: JSON.stringify(body),
    });
    return resp.body;
  }

  async function handleValidate() {
    if (!pptxBase64) return;
    setStatus("Checking template placeholders...");
    try {
      const res = await callCloudRun("/templates/validate", {
        filename: fileName,
        pptx_base64: pptxBase64,
        dashboard_tiles: tiles.length > 0 ? tiles : undefined,
      });
      setReport(res);
      setStatus(
        res?.valid
          ? `Template check passed! All ${res.bindings?.length || 0} items across ${res.slide_count} slides match.`
          : `Template check found ${res?.issues?.length || 0} item(s) to review.`
      );
    } catch (err: any) {
      setStatus(`Template check failed: ${err?.message || err}`);
    }
  }

  async function runPreviewWithTheme(overrideThemeName?: string) {
    if (!pptxBase64 || !selectedDashboardId) return;
    const targetTheme = overrideThemeName || activeEmbedTheme;
    setStatus(`Filling your PowerPoint slides with live data (${targetTheme})...`);
    try {
      const valRes =
        report ||
        (await callCloudRun("/templates/validate", {
          filename: fileName,
          pptx_base64: pptxBase64,
          dashboard_tiles: tiles.length > 0 ? tiles : undefined,
        }));
      setReport(valRes);

      const reqSlugs = new Set<string>(
        (valRes?.bindings || [])
          .map((b: any) => String(b.tile_slug || ""))
          .filter((s: string) => s && s !== "_plan")
      );
      const tileCsvs = await fetchDashboardTileCsvs(core40SDK, selectedDashboardId, reqSlugs);
      const selectedDash = dashboards.find((d) => d.id === selectedDashboardId);

      const mergeRes = await callCloudRun("/templates/preview", {
        pptx_base64: pptxBase64,
        tile_csvs: tileCsvs,
        plan_title:
          brand && brand !== "Default"
            ? `${brand} — ${selectedDash?.title || "Brand Overview"}`
            : selectedDash?.title || "Executive Business Review",
        plan_url: `https://looker.example.com/dashboards/${selectedDashboardId}`,
        theme: targetTheme,
      });
      if (mergeRes?.error) {
        setStatus(`Could not generate preview: ${mergeRes.error}`);
        return;
      }
      setMergedPptxBase64(String(mergeRes.pptx_base64 || ""));
      setStatus("");
    } catch (err: any) {
      setStatus(`Preview failed: ${err?.message || err}`);
    }
  }

  async function handleBrandOrModeChange(nextBrand: string, nextIsDark: boolean) {
    setBrand(nextBrand);
    setIsDark(nextIsDark);
    if (mergedPptxBase64 && pptxBase64 && selectedDashboardId) {
      desiredSlideRef.current = currentSlide;
      await runPreviewWithTheme(getEmbedThemeName(nextIsDark, nextBrand));
    }
  }

  function handleDownloadFilled() {
    if (!mergedPptxBase64) return;
    const outName = fileName.replace(/\.pptx$/i, "") + `_${activeEmbedTheme}.pptx`;
    const link = document.createElement("a");
    link.href = `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,${mergedPptxBase64}`;
    link.download = outName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  async function handleUpload() {
    if (!pptxBase64) return;
    setStatus("Saving template to your shared library...");
    try {
      const targetUri =
        config.templatesGcsUri ||
        (config.templatesUserAttributeName
          ? await extensionSDK.userAttributeGetItem(config.templatesUserAttributeName)
          : "");
      const res = await callCloudRun("/templates/upload", {
        filename: fileName,
        pptx_base64: pptxBase64,
        dashboard_tiles: tiles.length > 0 ? tiles : undefined,
        templates_gcs_uri: targetUri,
      });
      setReport(res);
      setStatus(
        res?.gcs_uri
          ? `Saved "${fileName}" to your shared template library.`
          : "Could not save template until all template issues are resolved."
      );
    } catch (err: any) {
      setStatus(`Save failed: ${err?.message || err}`);
    }
  }

  const isConfigured = Boolean(
    config.cloudRunUrl && (!config.useOAuth || config.oauthClientId)
  );

  // Show /config automatically if not configured yet, or when navigated to /config via route/hash
  if (!isConfigured || route === "/config") {
    return (
      <ConfigPage
        core40SDK={core40SDK}
        extensionSDK={extensionSDK}
        initialConfig={config}
        onSave={(next) => {
          setConfig(next);
          if (window.location.hash === "#/config") {
            window.location.hash = "";
          }
          setRoute("/");
        }}
        onBack={() => {
          if (window.location.hash === "#/config") {
            window.location.hash = "";
          }
          setRoute("/");
        }}
      />
    );
  }

  // Centered Sign-In screen when OAuth is enabled and the user is not logged in
  if (config.useOAuth && !oauthToken) {
    return (
      <div
        style={{
          minHeight: "100vh",
          background: ui.bg,
          color: ui.textPrimary,
          fontFamily: "Inter, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          padding: 24,
          boxSizing: "border-box",
        }}
      >
        <div
          style={{
            width: "100%",
            maxWidth: 460,
            background: ui.cardBg,
            border: `1px solid ${ui.border}`,
            borderRadius: 16,
            padding: "42px 36px",
            textAlign: "center",
            boxShadow: "0 20px 50px rgba(0, 0, 0, 0.35)",
          }}
        >
          <div
            style={{
              width: 56,
              height: 56,
              borderRadius: 14,
              background: "rgba(56, 189, 248, 0.12)",
              border: `1px solid ${ui.accent}`,
              color: ui.accent,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              margin: "0 auto 18px",
            }}
          >
            <IconSlides size={26} />
          </div>
          <h1 style={{ margin: "0 0 10px", fontSize: 24, fontWeight: 700, color: ui.textPrimary }}>
            PowerPoint Report Builder
          </h1>
          <p
            style={{
              margin: "0 0 28px",
              fontSize: 14,
              lineHeight: 1.6,
              color: ui.textMuted,
            }}
          >
            Sign in with your Google account to preview PowerPoint templates with live Looker
            dashboard data and save templates to your shared library.
          </p>
          <button
            id="btn-oauth-toggle"
            data-testid="center-oauth-signin-btn"
            type="button"
            onClick={() => void handleOAuthLogin()}
            disabled={oauthSigningIn}
            style={{
              width: "100%",
              display: "inline-flex",
              alignItems: "center",
              justifyContent: "center",
              gap: 10,
              background: ui.accent,
              color: ui.accentContrast,
              border: "none",
              borderRadius: 10,
              padding: "13px 20px",
              fontSize: 15,
              fontWeight: 700,
              cursor: oauthSigningIn ? "wait" : "pointer",
              boxShadow: "0 4px 14px rgba(0, 0, 0, 0.25)",
            }}
          >
            <IconGoogle />
            <span>{oauthSigningIn ? "Signing in..." : "Sign in with Google"}</span>
          </button>
          {status && (
            <div
              style={{
                marginTop: 16,
                fontSize: 13,
                color: "#fca5a5",
                background: "#450a0a",
                border: "1px solid #dc2626",
                borderRadius: 8,
                padding: "10px 12px",
              }}
            >
              {status}
            </div>
          )}
        </div>
      </div>
    );
  }

  const selectedDashTitle =
    dashboards.find((d) => d.id === selectedDashboardId)?.title || "Selected Dashboard";

  return (
    <div
      style={{
        minHeight: "100vh",
        background: ui.bg,
        color: ui.textPrimary,
        fontFamily: "Inter, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
        padding: "28px 32px",
        transition: "background 0.2s ease, color 0.2s ease",
      }}
    >
      <div style={{ maxWidth: 960, margin: "0 auto" }}>
        {/* Header with Dynamic Brand & Light/Dark Theme Switcher */}
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            flexWrap: "wrap",
            gap: 14,
            borderBottom: `1px solid ${ui.border}`,
            paddingBottom: 18,
            marginBottom: 24,
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div
              style={{
                width: 40,
                height: 40,
                borderRadius: 10,
                background: ui.cardBg,
                border: `1px solid ${ui.accent}`,
                color: ui.accent,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
              }}
            >
              <IconSlides size={20} />
            </div>
            <div>
              <h1 style={{ margin: 0, fontSize: 22, fontWeight: 700, color: ui.textPrimary }}>
                PowerPoint Report Builder
              </h1>
              <p style={{ margin: "3px 0 0", fontSize: 13, color: ui.textMuted }}>
                Preview your PowerPoint template with live Looker dashboard data or save it to your
                shared library.
              </p>
            </div>
          </div>

          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
            {/* Brand Theme Selector (matching looker-embed-demo) */}
            <div
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 6,
                background: ui.cardBg,
                border: `1px solid ${ui.border}`,
                borderRadius: 8,
                padding: "5px 10px",
                color: ui.accent,
              }}
            >
              <IconPalette />
              <select
                id="brand-theme-select"
                aria-label="Brand Theme"
                value={brand}
                onChange={(e) => void handleBrandOrModeChange(e.target.value, isDark)}
                style={{
                  background: "transparent",
                  color: ui.textPrimary,
                  border: "none",
                  fontSize: 12,
                  fontWeight: 600,
                  cursor: "pointer",
                  outline: "none",
                }}
              >
                {BRAND_OPTIONS.map((b) => (
                  <option
                    key={b}
                    value={b}
                    style={{ background: ui.cardBg, color: ui.textPrimary }}
                  >
                    {b === "Default" ? "Default Theme" : `${b} Theme`}
                  </option>
                ))}
              </select>
            </div>

            {/* Light / Dark Mode Toggle Button */}
            <button
              id="btn-theme-mode"
              type="button"
              onClick={() => void handleBrandOrModeChange(brand, !isDark)}
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 6,
                background: ui.cardBg,
                color: ui.textPrimary,
                border: `1px solid ${ui.border}`,
                borderRadius: 8,
                padding: "7px 12px",
                fontSize: 12,
                fontWeight: 600,
                cursor: "pointer",
              }}
            >
              {isDark ? <IconSun /> : <IconMoon />}
              <span>{isDark ? "Light Mode" : "Dark Mode"}</span>
            </button>

            {config.useOAuth && oauthToken && (
              <>
                <span
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 5,
                    fontSize: 12,
                    fontWeight: 600,
                    color: isDark ? "#34d399" : "#047857",
                    background: isDark ? "#064e3b" : "#d1fae5",
                    border: `1px solid ${isDark ? "#059669" : "#10b981"}`,
                    padding: "6px 11px",
                    borderRadius: 999,
                  }}
                >
                  <IconCheckCircle />
                  <span>Signed in</span>
                </span>
                <button
                  id="btn-oauth-signout"
                  type="button"
                  onClick={handleOAuthLogout}
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 6,
                    background: ui.cardBg,
                    color: ui.textSecondary,
                    border: `1px solid ${ui.border}`,
                    borderRadius: 8,
                    padding: "7px 12px",
                    fontSize: 12,
                    fontWeight: 600,
                    cursor: "pointer",
                  }}
                >
                  <IconLogOut />
                  <span>Sign Out</span>
                </button>
              </>
            )}
          </div>
        </div>

        {/* Step 1 & Step 2 Cards */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 18 }}>
          <div
            style={{
              background: ui.cardBg,
              border: `1px solid ${ui.border}`,
              borderRadius: 12,
              padding: 18,
            }}
          >
            <div style={{ fontSize: 11, fontWeight: 700, color: ui.accent, marginBottom: 4 }}>
              STEP 1
            </div>
            <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 10 }}>
              Choose a Dashboard
            </div>
            <select
              id="dashboard-select"
              value={selectedDashboardId}
              onChange={(e) => void handleSelectDashboard(e.target.value)}
              style={{
                width: "100%",
                padding: "10px 12px",
                borderRadius: 8,
                background: ui.inputBg,
                color: ui.textPrimary,
                border: `1px solid ${ui.border}`,
                fontSize: 13,
              }}
            >
              <option value="">-- Select a Looker Dashboard --</option>
              {dashboards.map((d) => (
                <option key={d.id} value={d.id}>
                  {d.title}
                </option>
              ))}
            </select>
            {tiles.length > 0 && (
              <div style={{ marginTop: 10, display: "flex", flexWrap: "wrap", gap: 6 }}>
                {tiles.slice(0, 9).map((t) => (
                  <span
                    key={t.slug}
                    style={{
                      fontSize: 11,
                      background: ui.inputBg,
                      color: ui.textSecondary,
                      border: `1px solid ${ui.border}`,
                      padding: "3px 9px",
                      borderRadius: 999,
                      fontWeight: 500,
                    }}
                  >
                    {t.title}
                  </span>
                ))}
                {tiles.length > 9 && (
                  <span
                    style={{
                      fontSize: 11,
                      background: ui.inputBg,
                      color: ui.accent,
                      border: `1px solid ${ui.border}`,
                      padding: "3px 9px",
                      borderRadius: 999,
                      fontWeight: 700,
                    }}
                  >
                    +{tiles.length - 9} more
                  </span>
                )}
              </div>
            )}
          </div>

          <div
            style={{
              background: ui.cardBg,
              border: `1px solid ${ui.border}`,
              borderRadius: 12,
              padding: 18,
            }}
          >
            <div style={{ fontSize: 11, fontWeight: 700, color: ui.accent, marginBottom: 4 }}>
              STEP 2
            </div>
            <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 10 }}>
              Select PowerPoint Template (.pptx)
            </div>
            <input
              id="pptx-file-input"
              type="file"
              accept=".pptx"
              onChange={(e) => void handleFileSelect(e)}
              style={{
                width: "100%",
                padding: "8px 10px",
                borderRadius: 8,
                background: ui.inputBg,
                color: ui.textSecondary,
                border: `1px solid ${ui.border}`,
                fontSize: 13,
                boxSizing: "border-box",
              }}
            />
            {fileName && (
              <div style={{ marginTop: 10, fontSize: 12, color: ui.textMuted }}>
                Selected template:{" "}
                <strong style={{ color: ui.textPrimary }}>{fileName}</strong>
              </div>
            )}
          </div>
        </div>

        {/* Action Toolbar */}
        <div
          style={{
            display: "flex",
            flexWrap: "wrap",
            gap: 10,
            alignItems: "center",
            background: ui.cardBg,
            border: `1px solid ${ui.border}`,
            borderRadius: 12,
            padding: "14px 18px",
            marginBottom: 18,
          }}
        >
          <button
            id="btn-preview"
            type="button"
            disabled={!pptxBase64 || !selectedDashboardId}
            onClick={() => void runPreviewWithTheme()}
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 8,
              background: ui.accent,
              color: ui.accentContrast,
              border: "none",
              borderRadius: 8,
              padding: "10px 18px",
              fontSize: 13,
              fontWeight: 700,
              cursor: pptxBase64 && selectedDashboardId ? "pointer" : "not-allowed",
              opacity: pptxBase64 && selectedDashboardId ? 1 : 0.5,
            }}
          >
            <IconPlay />
            <span>Live Preview</span>
          </button>

          <button
            id="btn-validate"
            type="button"
            disabled={!pptxBase64}
            onClick={() => void handleValidate()}
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 8,
              background: ui.inputBg,
              color: ui.textPrimary,
              border: `1px solid ${ui.border}`,
              borderRadius: 8,
              padding: "10px 16px",
              fontSize: 13,
              fontWeight: 600,
              cursor: pptxBase64 ? "pointer" : "not-allowed",
              opacity: pptxBase64 ? 1 : 0.5,
            }}
          >
            <IconCheck />
            <span>Check Template</span>
          </button>

          {mergedPptxBase64 && (
            <button
              id="btn-download"
              type="button"
              onClick={handleDownloadFilled}
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 8,
                background: ui.inputBg,
                color: ui.accent,
                border: `1px solid ${ui.accent}`,
                borderRadius: 8,
                padding: "10px 16px",
                fontSize: 13,
                fontWeight: 600,
                cursor: "pointer",
              }}
            >
              <IconDownload />
              <span>Download Filled .pptx</span>
            </button>
          )}

          <button
            id="btn-upload"
            type="button"
            disabled={!pptxBase64 || (report && !report.valid)}
            onClick={() => void handleUpload()}
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 8,
              background: "#059669",
              color: "#ffffff",
              border: "none",
              borderRadius: 8,
              padding: "10px 18px",
              fontSize: 13,
              fontWeight: 700,
              marginLeft: "auto",
              cursor: pptxBase64 && (!report || report.valid) ? "pointer" : "not-allowed",
              opacity: pptxBase64 && (!report || report.valid) ? 1 : 0.5,
            }}
          >
            <IconUpload />
            <span>Save to Template Library</span>
          </button>
        </div>

        {/* Status Banner */}
        {status && (
          <div
            style={{
              background: ui.cardBg,
              borderLeft: `4px solid ${ui.accent}`,
              border: `1px solid ${ui.border}`,
              padding: "12px 16px",
              borderRadius: 8,
              marginBottom: 18,
              fontSize: 13,
              fontWeight: 600,
              color: ui.textPrimary,
            }}
          >
            {status}
          </div>
        )}

        {/* Live Filled-In PPTX Preview Panel with Prominent SVG Icon Next / Back / End Controls */}
        {mergedPptxBase64 && (
          <div
            style={{
              background: ui.cardBg,
              border: `1px solid ${ui.border}`,
              borderRadius: 12,
              padding: 20,
              marginBottom: 20,
            }}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: 14,
              }}
            >
              <div>
                <div style={{ fontSize: 17, fontWeight: 700, color: ui.textPrimary }}>
                  Live Preview — {selectedDashTitle}
                </div>
                <div style={{ fontSize: 12, color: ui.textMuted, marginTop: 2 }}>
                  Viewing Slide {currentSlide + 1} of {totalSlides || 1}  •  Theme:{" "}
                  <strong style={{ color: ui.accent }}>{activeEmbedTheme}</strong>
                </div>
              </div>
              <button
                id="btn-close-preview"
                type="button"
                onClick={handleEndPreview}
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: 6,
                  background: isDark ? "#1e293b" : "#fee2e2",
                  color: isDark ? "#fda4af" : "#b91c1c",
                  border: `1px solid ${isDark ? "#881337" : "#fca5a5"}`,
                  borderRadius: 8,
                  padding: "8px 14px",
                  fontSize: 13,
                  fontWeight: 700,
                  cursor: "pointer",
                }}
              >
                <IconClose />
                <span>Close Preview</span>
              </button>
            </div>

            {/* Slide Canvas */}
            <div
              ref={previewContainerRef}
              style={{
                display: "flex",
                justifyContent: "center",
                background: ui.canvasBg,
                borderRadius: 10,
                padding: 16,
                border: `1px solid ${ui.border}`,
              }}
            />

            {/* Obvious Next / Back / End Slide Navigation Bar with SVG Icons */}
            {(() => {
              const navBtnStyle = (enabled: boolean, primary = false): React.CSSProperties => ({
                display: "inline-flex",
                alignItems: "center",
                gap: 6,
                background: !enabled ? ui.inputBg : primary ? ui.accent : ui.cardBg,
                color: !enabled ? ui.textMuted : primary ? ui.accentContrast : ui.textPrimary,
                border: `1px solid ${enabled && primary ? ui.accent : ui.border}`,
                borderRadius: 8,
                padding: primary ? "10px 22px" : "10px 18px",
                fontSize: 14,
                fontWeight: 700,
                cursor: enabled ? "pointer" : "not-allowed",
                opacity: enabled ? 1 : 0.5,
              });
              return (
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    marginTop: 16,
                    padding: "12px 16px",
                    background: ui.inputBg,
                    border: `1px solid ${ui.border}`,
                    borderRadius: 10,
                  }}
                >
                  <button
                    id="btn-slide-prev"
                    type="button"
                    disabled={currentSlide <= 0}
                    onClick={() => goToSlide(currentSlide - 1)}
                    style={navBtnStyle(currentSlide > 0)}
                  >
                    <IconArrowLeft />
                    <span>Back</span>
                  </button>

                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    {Array.from({ length: totalSlides }, (_, idx) => {
                      const active = idx === currentSlide;
                      return (
                        <button
                          key={idx}
                          id={`btn-slide-${idx + 1}`}
                          type="button"
                          onClick={() => goToSlide(idx)}
                          style={{
                            background: active ? ui.accent : ui.cardBg,
                            color: active ? ui.accentContrast : ui.textSecondary,
                            border: `1px solid ${active ? ui.accent : ui.border}`,
                            borderRadius: 8,
                            padding: "8px 14px",
                            fontSize: 13,
                            fontWeight: 700,
                            cursor: "pointer",
                          }}
                        >
                          Slide {idx + 1}
                        </button>
                      );
                    })}
                  </div>

                  <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                    <button
                      id="btn-slide-next"
                      type="button"
                      disabled={currentSlide >= totalSlides - 1}
                      onClick={() => goToSlide(currentSlide + 1)}
                      style={navBtnStyle(currentSlide < totalSlides - 1, true)}
                    >
                      <span>Next</span>
                      <IconArrowRight />
                    </button>
                    <button
                      id="btn-slide-end"
                      type="button"
                      disabled={currentSlide >= totalSlides - 1}
                      onClick={() => goToSlide(totalSlides - 1)}
                      style={navBtnStyle(currentSlide < totalSlides - 1)}
                    >
                      <span>End</span>
                      <IconSkipEnd />
                    </button>
                  </div>
                </div>
              );
            })()}
          </div>
        )}

        {/* Simplified End-User Template Check Card */}
        {report && (
          <div
            style={{
              background: ui.cardBg,
              border: `1px solid ${ui.border}`,
              borderRadius: 12,
              padding: "16px 20px",
            }}
          >
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div>
                <div style={{ fontSize: 15, fontWeight: 700, color: ui.textPrimary }}>
                  {report.valid ? "Template Ready" : "Template Needs Attention"}
                </div>
                <div style={{ fontSize: 13, color: ui.textMuted, marginTop: 2 }}>
                  {report.valid
                    ? `Matched ${report.bindings?.length || 0} charts, tables, and text placeholders across ${report.slide_count} slides.`
                    : `Found ${report.issues?.length || 0} item(s) that did not match the selected dashboard.`}
                </div>
              </div>
              <span
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: 6,
                  fontSize: 12,
                  fontWeight: 700,
                  padding: "6px 14px",
                  borderRadius: 999,
                  background: report.valid
                    ? isDark
                      ? "#064e3b"
                      : "#d1fae5"
                    : isDark
                      ? "#7f1d1d"
                      : "#fee2e2",
                  color: report.valid
                    ? isDark
                      ? "#34d399"
                      : "#047857"
                    : isDark
                      ? "#fca5a5"
                      : "#b91c1c",
                }}
              >
                {report.valid ? <IconCheckCircle /> : <IconAlert />}
                <span>{report.valid ? "Ready to Use" : `${report.issues?.length || 0} Issue(s)`}</span>
              </span>
            </div>

            {Array.isArray(report.issues) && report.issues.length > 0 && (
              <div
                style={{
                  background: isDark ? "#450a0a" : "#fef2f2",
                  border: "1px solid #dc2626",
                  borderRadius: 8,
                  padding: 12,
                  marginTop: 12,
                }}
              >
                {report.issues.map((iss: any, idx: number) => (
                  <div
                    key={idx}
                    style={{
                      fontSize: 13,
                      color: isDark ? "#fecaca" : "#991b1b",
                      marginBottom: 4,
                    }}
                  >
                    Slide {iss.slide_index || "?"}: {iss.message}
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};
