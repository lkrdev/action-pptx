import React, { useEffect, useMemo, useRef, useState } from "react";
import { init as initPptxPreview } from "pptx-preview";
import { ConfigPage, copyToClipboardSafe, ExtensionConfigData, Svg } from "./ConfigPage";
import {
  DashboardPicker,
  DashboardSummary,
  FALLBACK_LOOKER_THEMES,
  hexLuminance,
  loadLookerThemeOptions,
  LookerThemeOption,
  LookerThemePicker,
  TEMPLATE_ORIGINAL_THEME,
  useDismissOnOutsideOrEscape,
} from "./DashboardPicker";
import {
  DashboardTileSchema,
  fetchDashboardTileCsvs,
  fetchDashboardTileSchemas,
} from "./dashboardSchema";
import { SchedulerSeed, SchedulerView } from "./SchedulerView";
import { buildTemplateLinksArtifactKey, upsertTemplateLinkArtifact } from "./schedulerUtils";

interface AppProps {
  core40SDK: any;
  extensionSDK: any;
  initialRoute?: string;
  onHostRouteChange?: (listener: (route: string) => void) => () => void;
}

type ViewerMode = "outputs" | "templates";

function parseExtensionRoute(raw?: string, fallbackSearch = "") {
  const trimmed = (raw || "").trim();
  const qIdx = trimmed.indexOf("?");
  const cleaned = (qIdx >= 0 ? trimmed.slice(0, qIdx) : trimmed).replace(/\/+$/, "");
  const params = new URLSearchParams(qIdx >= 0 ? trimmed.slice(qIdx + 1) : fallbackSearch);
  const path: "/" | "/viewer" | "/scheduler" | "/config" =
    cleaned === "config" || cleaned.endsWith("/config")
      ? "/config"
      : cleaned === "viewer" || cleaned.endsWith("/viewer")
        ? "/viewer"
        : cleaned === "scheduler" || cleaned.endsWith("/scheduler")
          ? "/scheduler"
          : "/";
  const mode = params.get("mode");
  return {
    path,
    filename: (params.get("filename") || "").trim(),
    mode: mode === "templates" || mode === "outputs" ? (mode as ViewerMode) : undefined,
    q: (params.get("q") || "").trim(),
    userId: (params.get("user_id") || "").trim(),
  };
}

function buildViewerClientRoute(filename: string, mode: ViewerMode, query = ""): string {
  const params = new URLSearchParams();
  if (mode === "templates") params.set("mode", "templates");
  if (filename) params.set("filename", filename);
  if (query.trim()) params.set("q", query.trim());
  const qs = params.toString();
  return qs ? `/viewer?${qs}` : "/viewer";
}

function buildSchedulerClientRoute(userId = ""): string {
  return userId.trim() ? `/scheduler?user_id=${encodeURIComponent(userId.trim())}` : "/scheduler";
}

export const OAUTH_STORAGE_KEY = "action_pptx_oauth_session";

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

function getUiPalette(accentHex: string | undefined, isDark: boolean): UiPalette {
  let accent = accentHex || (isDark ? "#38bdf8" : "#0b57d0");
  const rawLum = hexLuminance(accent);
  if (isDark && rawLum < 35) accent = "#f8fafc";
  if (!isDark && rawLum > 215) accent = "#0f172a";
  const accentContrast = hexLuminance(accent) > 145 ? "#0f172a" : "#ffffff";

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
      accentContrast,
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
    accentContrast,
    canvasBg: "#e2e8f0",
  };
}

/** Option 1 (Fluent Slide Deck) SVG Logo matching the Looker Action Hub icon */
const ActionPptxDeckLogo: React.FC<{ size?: number; pulse?: boolean }> = ({
  size = 28,
  pulse = false,
}) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 64 64"
    fill="none"
    aria-hidden="true"
    style={
      pulse
        ? {
            animation: "actionPptxPulse 1.15s ease-in-out infinite",
            filter: "drop-shadow(0 8px 20px rgba(217, 79, 43, 0.42))",
          }
        : undefined
    }
  >
    <defs>
      <linearGradient id="apBgGrad" x1="6" y1="6" x2="58" y2="58" gradientUnits="userSpaceOnUse">
        <stop offset="0%" stopColor="#FF7A4D" />
        <stop offset="52%" stopColor="#D94F2B" />
        <stop offset="100%" stopColor="#A62C14" />
      </linearGradient>
    </defs>
    <rect x="4" y="4" width="56" height="56" rx="14" fill="url(#apBgGrad)" />
    <rect
      x="17"
      y="13"
      width="35"
      height="24"
      rx="4"
      fill="#FFFFFF"
      fillOpacity="0.24"
      transform="rotate(4 34 25)"
    />
    <rect x="11" y="18" width="38" height="27" rx="4.5" fill="#FFFFFF" />
    <rect x="16" y="23" width="16" height="3" rx="1.5" fill="#D94F2B" />
    <rect x="17" y="34" width="5" height="7" rx="1.2" fill="#F59E0B" />
    <rect x="25" y="30" width="5" height="11" rx="1.2" fill="#D94F2B" />
    <rect x="33" y="26" width="5" height="15" rx="1.2" fill="#10B981" />
    <circle cx="48" cy="45" r="9.5" fill="#0F172A" />
    <path
      d="M44.5 45.2L47 47.7L52 42.5"
      stroke="#38BDF8"
      strokeWidth="2.6"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
  </svg>
);

const IconGoogle = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
    <path d="M21.8 12.2c0-.7-.1-1.4-.2-2H12v3.8h5.5a4.7 4.7 0 0 1-2 3.1v2.6h3.3c1.9-1.8 3-4.4 3-7.5ZM12 22c2.7 0 5-.9 6.7-2.4l-3.3-2.6c-.9.6-2.1 1-3.4 1-2.6 0-4.8-1.8-5.6-4.1H3v2.6A10 10 0 0 0 12 22ZM6.4 13.9a6 6 0 0 1 0-3.8V7.5H3a10 10 0 0 0 0 9l3.4-2.6ZM12 5.9c1.5 0 2.8.5 3.8 1.5l2.9-2.9A10 10 0 0 0 3 7.5l3.4 2.6c.8-2.3 3-4.2 5.6-4.2Z" />
  </svg>
);
const IconCheckCircle = () => (
  <Svg size={16}>
    <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
    <path d="m9 11 3 3L22 4" />
  </Svg>
);
const IconXCircle = () => (
  <Svg size={16}>
    <circle cx="12" cy="12" r="10" />
    <path d="m15 9-6 6M9 9l6 6" />
  </Svg>
);
const IconLogOut = () => (
  <Svg size={14}>
    <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4m7 14 5-5-5-5m5 5H9" />
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
  <Svg size={14}>
    <path d="M18 6 6 18M6 6l12 12" />
  </Svg>
);
const IconExternal = () => (
  <Svg size={13}>
    <path d="M15 3h6v6M10 14 21 3M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6" />
  </Svg>
);
const IconCopy = () => (
  <Svg size={13}>
    <rect x="9" y="9" width="13" height="13" rx="2" />
    <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
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
const IconSearch = () => (
  <Svg size={14}>
    <circle cx="11" cy="11" r="8" />
    <path d="m21 21-4.3-4.3" />
  </Svg>
);
const IconRefresh = () => (
  <Svg size={14}>
    <path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8" />
    <path d="M21 3v5h-5" />
  </Svg>
);

interface OutputReportItem {
  name: string;
  size: number;
  updated: string;
}

const VIEWER_MODE_META: Record<
  ViewerMode,
  {
    tabLabel: string;
    searchPlaceholder: string;
    dateColumn: string;
    emptyText: string;
    missingUriText: string;
    listPath: string;
    downloadPath: string;
    uriField: string;
  }
> = {
  outputs: {
    tabLabel: "Rendered PowerPoints",
    searchPlaceholder: "Search rendered PowerPoints by filename...",
    dateColumn: "Generated",
    emptyText: "No matching .pptx reports found.",
    missingUriText: "Configure an Output GCS URI in /config to view rendered PowerPoints.",
    listPath: "/outputs/list",
    downloadPath: "/outputs/download",
    uriField: "output_gcs_uri",
  },
  templates: {
    tabLabel: "Templates",
    searchPlaceholder: "Search templates by filename...",
    dateColumn: "Updated",
    emptyText: "No matching .pptx templates found.",
    missingUriText: "Configure a Templates GCS URI in /config to view templates.",
    listPath: "/templates/list",
    downloadPath: "/templates/download",
    uriField: "templates_gcs_uri",
  },
};

function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 B";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function base64ToArrayBuffer(b64: string): ArrayBuffer {
  const binary = window.atob(b64.replace(/\s+/g, ""));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

function downloadPptxBase64(b64: string, outName: string): void {
  const blob = new Blob([base64ToArrayBuffer(b64)], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = outName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

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
    void extensionSDK?.localStorageSetItem?.(OAUTH_STORAGE_KEY, serialized);
  } catch {
    // ignore extension storage errors
  }
}

function clearOAuthSession(extensionSDK: any): void {
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

export const App: React.FC<AppProps> = ({
  core40SDK,
  extensionSDK,
  initialRoute,
  onHostRouteChange,
}) => {
  const initialParsed = useMemo(
    () =>
      parseExtensionRoute(
        initialRoute ?? extensionSDK?.lookerHostData?.route ?? window.location.pathname,
        window.location.search
      ),
    []
  );
  const [route, setRoute] = useState<string>(initialParsed.path);
  const [config, setConfig] = useState<ExtensionConfigData>(() => {
    const ctx = extensionSDK?.getContextData?.() as ExtensionConfigData | undefined;
    return {
      cloudRunUrl: ctx?.cloudRunUrl || "",
      templatesGcsUri: ctx?.templatesGcsUri || "",
      outputGcsUri: ctx?.outputGcsUri || "",
      actionIntegrationId: ctx?.actionIntegrationId,
      templatesUserAttributeName: ctx?.templatesUserAttributeName,
      useOAuth: Boolean(ctx?.useOAuth),
      oauthClientId: ctx?.oauthClientId,
      oauthGcsPassthrough: ctx?.oauthGcsPassthrough ?? true,
    };
  });
  const [serverViewerEnabled, setServerViewerEnabled] = useState<boolean>(false);
  const [viewerMode, setViewerMode] = useState<ViewerMode>(initialParsed.mode || "outputs");
  const [requestedViewerFile, setRequestedViewerFile] = useState<string>(initialParsed.filename);
  const [viewerSourceUri, setViewerSourceUri] = useState<string>("");
  const [viewerFiles, setViewerFiles] = useState<OutputReportItem[]>([]);
  const [viewerSearch, setViewerSearch] = useState<string>(initialParsed.q);
  const [schedulerUserId, setSchedulerUserId] = useState<string>(initialParsed.userId);
  const [schedulerSeed, setSchedulerSeed] = useState<SchedulerSeed | null>(null);
  const [viewerSortKey, setViewerSortKey] = useState<"name" | "size" | "updated">("updated");
  const [viewerSortAsc, setViewerSortAsc] = useState<boolean>(false);
  const [viewerLoadingList, setViewerLoadingList] = useState<boolean>(false);
  const [viewerLoadingFile, setViewerLoadingFile] = useState<string>("");
  const [viewerFileName, setViewerFileName] = useState<string>("");
  const [viewerPptxBase64, setViewerPptxBase64] = useState<string>("");
  const [confirmDeleteViewerFile, setConfirmDeleteViewerFile] = useState<string>("");
  const [viewerDeletingFile, setViewerDeletingFile] = useState<string>("");

  const artifactNamespace = useMemo(
    () =>
      String(
        extensionSDK?.lookerHostData?.extensionId || "action-pptx::action_pptx_template_manager"
      ),
    [extensionSDK]
  );

  const [lookerThemes, setLookerThemes] = useState<LookerThemeOption[]>(FALLBACK_LOOKER_THEMES);
  const [selectedThemeName, setSelectedThemeName] = useState<string>("Levis_Dark");
  const [isDark, setIsDark] = useState<boolean>(true);

  const [selectedDashboardId, setSelectedDashboardId] = useState<string>("");
  const [selectedDashboardTitle, setSelectedDashboardTitle] = useState<string>("");
  const [tiles, setTiles] = useState<DashboardTileSchema[]>([]);

  const [fileName, setFileName] = useState<string>("");
  const [pptxBase64, setPptxBase64] = useState<string>("");
  const [templateFingerprint, setTemplateFingerprint] = useState<string>("");

  const [report, setReport] = useState<any>(null);
  const [actionError, setActionError] = useState<string>("");
  const [savedNotice, setSavedNotice] = useState<string>("");
  const [busyLabel, setBusyLabel] = useState<string>("");
  const [checkPopoverOpen, setCheckPopoverOpen] = useState<boolean>(false);
  const [copiedErrors, setCopiedErrors] = useState<boolean>(false);

  const [mergedPptxBase64, setMergedPptxBase64] = useState<string>("");
  const [previewLoading, setPreviewLoading] = useState<boolean>(false);
  const [previewLoadingLabel, setPreviewLoadingLabel] = useState<string>("");
  const [currentSlide, setCurrentSlide] = useState<number>(0);
  const [totalSlides, setTotalSlides] = useState<number>(0);

  const [oauthToken, setOauthToken] = useState<string>("");
  const [oauthSigningIn, setOauthSigningIn] = useState<boolean>(false);

  const previewContainerRef = useRef<HTMLDivElement | null>(null);
  const slideStripRef = useRef<HTMLDivElement | null>(null);
  const previewerRef = useRef<any>(null);
  const desiredSlideRef = useRef<number>(0);
  const checkPopoverRef = useRef<HTMLDivElement | null>(null);

  // Monotonic sequence guard + SWR-style in-memory caches for instant/race-safe theme switching
  const previewSeqRef = useRef<number>(0);
  const viewerLoadSeqRef = useRef<number>(0);
  const tileCsvCacheRef = useRef<Map<string, Record<string, string>>>(new Map());
  const previewPptxCacheRef = useRef<Map<string, string>>(new Map());

  const activeThemeOption = useMemo(
    () =>
      lookerThemes.find((t) => t.name === selectedThemeName) ||
      lookerThemes[1] ||
      FALLBACK_LOOKER_THEMES[1],
    [lookerThemes, selectedThemeName]
  );

  const ui = getUiPalette(activeThemeOption?.palette?.accent, isDark);
  const canOpenViewer = Boolean(config.outputGcsUri?.trim()) && serverViewerEnabled;
  const activePreviewBase64 = route === "/viewer" ? viewerPptxBase64 : mergedPptxBase64;

  function navigateTo(raw?: string, fallbackSearch = "") {
    const parsed = parseExtensionRoute(raw, fallbackSearch);
    const hasExplicitQuery = Boolean((raw && raw.includes("?")) || fallbackSearch);
    setActionError("");
    setRoute(parsed.path);
    if (parsed.mode) {
      setViewerMode(parsed.mode);
    }
    if (parsed.path === "/viewer") {
      if (hasExplicitQuery) {
        setViewerSearch(parsed.q);
        if (!parsed.mode && parsed.q) setViewerMode("outputs");
      }
      setRequestedViewerFile(parsed.filename);
      if (!parsed.filename) {
        viewerLoadSeqRef.current += 1;
        setViewerLoadingFile("");
        setViewerPptxBase64("");
        setViewerFileName("");
      }
    } else if (parsed.path === "/scheduler" && hasExplicitQuery) {
      setSchedulerUserId(parsed.userId);
    }
  }

  useEffect(() => {
    const initialClientId = config.oauthClientId || "";
    void Promise.resolve(extensionSDK?.localStorageGetItem?.(OAUTH_STORAGE_KEY))
      .then((extRaw: string | null) => {
        const extSession = parseStoredOAuthSession(extRaw, initialClientId);
        if (extSession?.token) setOauthToken(extSession.token);
      })
      .catch(() => {});

    void Promise.all([
      loadLookerThemeOptions(core40SDK),
      Promise.resolve(extensionSDK?.userAttributeGetItem?.("brand")).catch(() => null),
    ]).then(([loadedThemes, attrBrand]) => {
      setLookerThemes(loadedThemes);
      const preferred = getEmbedThemeName(true, String(attrBrand || ""));
      const match =
        loadedThemes.find((t) => t.name.toLowerCase() === preferred.toLowerCase()) ||
        loadedThemes.find((t) => t.name === "Embed_Demo_Dark") ||
        loadedThemes.find((t) => t.name !== TEMPLATE_ORIGINAL_THEME);
      if (match) {
        setSelectedThemeName(match.name);
        setIsDark(match.isDark);
      }
    });

    const onPopState = () => {
      navigateTo(window.location.pathname, window.location.search);
    };
    const onLookerRoute = (e: Event) => {
      navigateTo(String((e as CustomEvent).detail || "/"));
    };
    const unsubscribeHostRoute = onHostRouteChange?.((nextRoute) => {
      navigateTo(nextRoute);
    });
    window.addEventListener("popstate", onPopState);
    window.addEventListener("looker-route", onLookerRoute);

    return () => {
      window.removeEventListener("popstate", onPopState);
      window.removeEventListener("looker-route", onLookerRoute);
      unsubscribeHostRoute?.();
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    if (!config.cloudRunUrl || !config.outputGcsUri?.trim()) {
      setServerViewerEnabled(false);
      return;
    }
    void fetch(`${config.cloudRunUrl}/health`)
      .then((r) => (r.ok ? r.json() : null))
      .then((data) => {
        if (!cancelled) setServerViewerEnabled(Boolean(data?.report_viewer_enabled));
      })
      .catch(() => {
        if (!cancelled) setServerViewerEnabled(false);
      });
    return () => {
      cancelled = true;
    };
  }, [config.cloudRunUrl, config.outputGcsUri]);

  useEffect(() => {
    if (route === "/viewer" && canOpenViewer && (!config.useOAuth || oauthToken)) {
      void loadViewerFiles(viewerMode);
    }
  }, [
    route,
    viewerMode,
    canOpenViewer,
    config.outputGcsUri,
    config.templatesGcsUri,
    config.templatesUserAttributeName,
    oauthToken,
  ]);

  useEffect(() => {
    if (
      route === "/viewer" &&
      requestedViewerFile &&
      canOpenViewer &&
      (!config.useOAuth || oauthToken) &&
      viewerFileName !== requestedViewerFile &&
      viewerLoadingFile !== requestedViewerFile
    ) {
      void handleLoadViewerReport(requestedViewerFile, viewerMode);
    }
  }, [
    route,
    requestedViewerFile,
    viewerMode,
    canOpenViewer,
    config.useOAuth,
    oauthToken,
    viewerFileName,
    viewerLoadingFile,
    handleLoadViewerReport,
  ]);

  useEffect(() => {
    const activeBtn = slideStripRef.current?.querySelector<HTMLElement>('[aria-current="true"]');
    activeBtn?.scrollIntoView({ block: "nearest", inline: "center", behavior: "smooth" });
  }, [currentSlide, totalSlides]);

  // Sticky check/error popover: stays open until user clicks outside or presses Escape
  useDismissOnOutsideOrEscape(checkPopoverOpen, checkPopoverRef, () => setCheckPopoverOpen(false));

  useEffect(() => {
    if (!activePreviewBase64 || !previewContainerRef.current) return;
    const container = previewContainerRef.current;
    container.innerHTML = "";
    const previewer = initPptxPreview(container, {
      width: 860,
      height: 484,
      mode: "list",
    });
    previewerRef.current = previewer;
    void previewer.load(base64ToArrayBuffer(activePreviewBase64)).then(() => {
      const count = Number(previewer.slideCount) || 1;
      setTotalSlides(count);
      const targetIdx = Math.max(0, Math.min(count - 1, desiredSlideRef.current));
      setCurrentSlide(targetIdx);
      previewer.renderSingleSlide(targetIdx);
    });
  }, [activePreviewBase64, route]);

  function goToSlide(index: number) {
    if (!previewerRef.current || totalSlides <= 0) return;
    const clamped = Math.max(0, Math.min(totalSlides - 1, index));
    desiredSlideRef.current = clamped;
    previewerRef.current.renderSingleSlide(clamped);
    setCurrentSlide(clamped);
  }

  function handleEndPreview() {
    previewSeqRef.current += 1;
    if (previewerRef.current?.destroy) {
      previewerRef.current.destroy();
    }
    previewerRef.current = null;
    desiredSlideRef.current = 0;
    setPreviewLoading(false);
    if (route === "/viewer") {
      viewerLoadSeqRef.current += 1;
      setViewerLoadingFile("");
      setRequestedViewerFile("");
      setViewerPptxBase64("");
      setViewerFileName("");
    } else {
      setMergedPptxBase64("");
    }
    setCurrentSlide(0);
    setTotalSlides(0);
  }

  async function handleOAuthLogin(): Promise<string> {
    if (!config.oauthClientId) {
      navigateTo("/config");
      throw new Error("OAuth Client ID is not configured.");
    }
    setOauthSigningIn(true);
    setActionError("");
    try {
      const token = await loginWithGoogleOAuth(
        extensionSDK,
        config.oauthClientId,
        config.oauthGcsPassthrough !== false
      );
      setOauthToken(token);
      return token;
    } catch (err: any) {
      setActionError(`Sign-in failed: ${err?.message || err}`);
      throw err;
    } finally {
      setOauthSigningIn(false);
    }
  }

  function handleOAuthLogout() {
    clearOAuthSession(extensionSDK);
    setOauthToken("");
    setActionError("");
  }

  async function handleSelectDashboard(dashId: string, summary?: DashboardSummary) {
    setSelectedDashboardId(dashId);
    setSelectedDashboardTitle(summary?.title || dashId);
    setReport(null);
    setActionError("");
    setSavedNotice("");
    setCheckPopoverOpen(false);
    handleEndPreview();
    if (!dashId) {
      setTiles([]);
      return;
    }
    setBusyLabel("Loading tiles...");
    try {
      const schema = await fetchDashboardTileSchemas(core40SDK, dashId);
      setTiles(schema);
    } catch (err: any) {
      setActionError(`Could not load dashboard tiles: ${err?.message || err}`);
      setCheckPopoverOpen(true);
    } finally {
      setBusyLabel("");
    }
  }

  async function handleFileSelect(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    const b64 = await readFileAsBase64(file);
    setFileName(file.name);
    setPptxBase64(b64);
    setTemplateFingerprint(`${file.name}:${file.size}:${file.lastModified}:${b64.slice(-32)}`);
    setReport(null);
    setActionError("");
    setSavedNotice("");
    setCheckPopoverOpen(false);
    handleEndPreview();
  }

  async function callCloudRun(path: string, body: Record<string, any>): Promise<any> {
    if (!config.cloudRunUrl) {
      navigateTo("/config");
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
      if (!resp.ok) {
        const errBody = await resp.json().catch(() => ({}));
        const detail = String(errBody?.detail || "");
        if (resp.status === 401 || (resp.status === 403 && !detail.includes("disabled by server"))) {
          clearOAuthSession(extensionSDK);
          setOauthToken("");
          throw new Error(detail || `Session expired (${resp.status}). Please sign in again.`);
        }
        throw new Error(detail || `Request failed (${resp.status}).`);
      }
      return await resp.json();
    }

    const secretTag = extensionSDK.createSecretKeyTag("action_hub_secret");
    const resp = await extensionSDK.serverProxy(url, {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Token token="${secretTag}"` },
      body: JSON.stringify(body),
    });
    if (resp.status >= 400) {
      throw new Error(resp.body?.detail || resp.body || `Request failed (${resp.status})`);
    }
    return resp.body;
  }

  async function resolveTemplatesGcsUri(): Promise<string> {
    const direct = config.templatesGcsUri?.trim();
    if (direct) return direct;
    if (config.templatesUserAttributeName && extensionSDK?.userAttributeGetItem) {
      const val = await extensionSDK.userAttributeGetItem(config.templatesUserAttributeName);
      return String(val || "").trim();
    }
    return "";
  }

  async function resolveViewerBucketUri(mode: ViewerMode = viewerMode): Promise<string> {
    return mode === "templates" ? resolveTemplatesGcsUri() : (config.outputGcsUri?.trim() || "");
  }

  function handleSwitchViewerMode(nextMode: ViewerMode) {
    if (nextMode === viewerMode) return;
    handleEndPreview();
    setConfirmDeleteViewerFile("");
    setViewerFiles([]);
    setViewerSourceUri("");
    setActionError("");
    setViewerMode(nextMode);
  }

  async function loadViewerFiles(mode: ViewerMode = viewerMode) {
    setConfirmDeleteViewerFile("");
    const bucketUri = await resolveViewerBucketUri(mode);
    setViewerSourceUri(bucketUri);
    if (!bucketUri) {
      setViewerFiles([]);
      return;
    }
    const meta = VIEWER_MODE_META[mode];
    setViewerLoadingList(true);
    setActionError("");
    try {
      const res = await callCloudRun(meta.listPath, {
        [meta.uriField]: bucketUri,
      });
      setViewerFiles(Array.isArray(res?.files) ? res.files : []);
    } catch (err: any) {
      setActionError(`Could not list files: ${err?.message || err}`);
    } finally {
      setViewerLoadingList(false);
    }
  }

  async function fetchViewerReport(
    name: string,
    mode: ViewerMode = viewerMode
  ): Promise<{ filename: string; pptxBase64: string }> {
    const bucketUri =
      (mode === viewerMode && viewerSourceUri) || (await resolveViewerBucketUri(mode));
    const meta = VIEWER_MODE_META[mode];
    if (!bucketUri) throw new Error(meta.missingUriText);
    const res = await callCloudRun(meta.downloadPath, {
      [meta.uriField]: bucketUri,
      filename: name,
    });
    if (!res?.pptx_base64) {
      throw new Error("Server response did not contain presentation data.");
    }
    return { filename: String(res.filename || name), pptxBase64: String(res.pptx_base64) };
  }

  async function handleLoadViewerReport(name: string, mode: ViewerMode = viewerMode) {
    const seq = ++viewerLoadSeqRef.current;
    setRequestedViewerFile(name);
    setViewerLoadingFile(name);
    setPreviewLoadingLabel("Loading presentation...");
    setPreviewLoading(true);
    setActionError("");
    try {
      const { filename, pptxBase64: b64 } = await fetchViewerReport(name, mode);
      if (seq !== viewerLoadSeqRef.current) return;
      desiredSlideRef.current = 0;
      setRequestedViewerFile(filename);
      setViewerFileName(filename);
      setViewerPptxBase64(b64);
    } catch (err: any) {
      if (seq === viewerLoadSeqRef.current) {
        setRequestedViewerFile("");
        setActionError(`Could not load "${name}": ${err?.message || err}`);
      }
    } finally {
      if (seq === viewerLoadSeqRef.current) {
        setPreviewLoading(false);
        setViewerLoadingFile("");
      }
    }
  }

  async function handleDownloadViewerReport(name: string) {
    if (viewerFileName === name && viewerPptxBase64) {
      downloadPptxBase64(viewerPptxBase64, name);
      return;
    }
    setViewerLoadingFile(name);
    setActionError("");
    try {
      const { filename, pptxBase64: b64 } = await fetchViewerReport(name);
      downloadPptxBase64(b64, filename);
    } catch (err: any) {
      setActionError(`Could not download "${name}": ${err?.message || err}`);
    } finally {
      setViewerLoadingFile("");
    }
  }

  async function handleDeleteViewerTemplate(name: string) {
    if (viewerMode !== "templates") return;
    const bucketUri = viewerSourceUri || (await resolveViewerBucketUri("templates"));
    if (!bucketUri) {
      setActionError(VIEWER_MODE_META.templates.missingUriText);
      return;
    }
    setViewerDeletingFile(name);
    setActionError("");
    try {
      await callCloudRun("/templates/delete", {
        templates_gcs_uri: bucketUri,
        filename: name,
      });
      if (typeof core40SDK?.delete_artifact === "function") {
        await core40SDK
          .ok(core40SDK.delete_artifact(artifactNamespace, buildTemplateLinksArtifactKey(name)))
          .catch(() => {});
      }
      setConfirmDeleteViewerFile("");
      if (viewerFileName === name) {
        handleEndPreview();
      }
      setViewerFiles((prev) => prev.filter((item) => item.name !== name));
    } catch (err: any) {
      setActionError(`Could not delete "${name}": ${err?.message || err}`);
    } finally {
      setViewerDeletingFile("");
    }
  }

  function toggleViewerSort(key: "name" | "size" | "updated") {
    if (viewerSortKey === key) {
      setViewerSortAsc((prev) => !prev);
    } else {
      setViewerSortKey(key);
      setViewerSortAsc(key === "name");
    }
  }

  async function handleValidate() {
    if (!pptxBase64) return;
    setBusyLabel("Checking template...");
    setActionError("");
    setSavedNotice("");
    try {
      const res = await callCloudRun("/templates/validate", {
        filename: fileName,
        pptx_base64: pptxBase64,
        dashboard_tiles: tiles.length > 0 ? tiles : undefined,
      });
      setReport(res);
      if (!res?.valid) {
        setCheckPopoverOpen(true);
      }
    } catch (err: any) {
      setActionError(`Template check failed: ${err?.message || err}`);
      setCheckPopoverOpen(true);
    } finally {
      setBusyLabel("");
    }
  }

  async function runPreview(overrideTheme?: LookerThemeOption) {
    if (!pptxBase64 || !selectedDashboardId) return;
    const targetOpt = overrideTheme || activeThemeOption;
    const seq = ++previewSeqRef.current;
    const fp = templateFingerprint || `${fileName}:${pptxBase64.length}:${pptxBase64.slice(-32)}`;
    const cacheKey = `${selectedDashboardId}::${fp}::${targetOpt.name}`;

    setActionError("");

    // SWR-style instant cache hit when switching back to an already-rendered theme on the same upload
    const cachedMerged = previewPptxCacheRef.current.get(cacheKey);
    if (cachedMerged) {
      setPreviewLoading(false);
      setMergedPptxBase64(cachedMerged);
      return;
    }

    // Immediately show the animated Slide Deck loading logo overlay on the Live Preview
    setPreviewLoadingLabel(
      mergedPptxBase64 ? `Applying ${targetOpt.label}...` : "Building live slide preview..."
    );
    setPreviewLoading(true);

    try {
      let valRes = report;
      if (!valRes) {
        valRes = await callCloudRun("/templates/validate", {
          filename: fileName,
          pptx_base64: pptxBase64,
          dashboard_tiles: tiles.length > 0 ? tiles : undefined,
        });
        if (seq !== previewSeqRef.current) return;
        setReport(valRes);
        if (valRes && !valRes.valid) {
          setCheckPopoverOpen(true);
        }
      }

      const reqSlugsArr = Array.from(
        new Set<string>(
          (valRes?.bindings || [])
            .map((b: any) => String(b.tile_slug || ""))
            .filter((s: string) => s && s !== "_plan")
        )
      ).sort();
      const csvCacheKey = `${selectedDashboardId}::${reqSlugsArr.join(",")}`;
      let tileCsvs = tileCsvCacheRef.current.get(csvCacheKey);
      if (!tileCsvs) {
        tileCsvs = await fetchDashboardTileCsvs(
          core40SDK,
          selectedDashboardId,
          new Set(reqSlugsArr)
        );
        if (seq !== previewSeqRef.current) return;
        tileCsvCacheRef.current.set(csvCacheKey, tileCsvs);
      }

      const hostBase = (
        extensionSDK?.lookerHostData?.hostUrl || "https://looker.example.com"
      ).replace(/\/+$/, "");

      const isOriginal = targetOpt.name === TEMPLATE_ORIGINAL_THEME;
      const mergeRes = await callCloudRun("/templates/preview", {
        pptx_base64: pptxBase64,
        tile_csvs: tileCsvs,
        plan_title: selectedDashboardTitle || "Executive Business Review",
        plan_url: `${hostBase}/dashboards/${selectedDashboardId}`,
        theme: isOriginal ? undefined : targetOpt.name,
        theme_palette: isOriginal ? undefined : targetOpt.palette,
      });

      // Discard late-arriving response if user clicked another theme in the meantime
      if (seq !== previewSeqRef.current) return;

      if (mergeRes?.error) {
        setPreviewLoading(false);
        setActionError(`Could not generate preview: ${mergeRes.error}`);
        setCheckPopoverOpen(true);
        return;
      }

      const outB64 = String(mergeRes?.pptx_base64 || "");
      if (outB64) {
        previewPptxCacheRef.current.set(cacheKey, outB64);
        setMergedPptxBase64(outB64);
      }
      setPreviewLoading(false);
    } catch (err: any) {
      if (seq !== previewSeqRef.current) return;
      setPreviewLoading(false);
      setActionError(`Preview failed: ${err?.message || err}`);
      setCheckPopoverOpen(true);
    }
  }

  function handleSelectLookerTheme(opt: LookerThemeOption) {
    setSelectedThemeName(opt.name);
    if (opt.name !== TEMPLATE_ORIGINAL_THEME) {
      setIsDark(opt.isDark);
    }
    if ((mergedPptxBase64 || previewLoading) && pptxBase64 && selectedDashboardId) {
      desiredSlideRef.current = currentSlide;
      void runPreview(opt);
    }
  }

  function handleDownloadFilled() {
    if (!mergedPptxBase64) return;
    const suffix =
      selectedThemeName === TEMPLATE_ORIGINAL_THEME ? "filled" : selectedThemeName;
    const outName = fileName.replace(/\.pptx$/i, "") + `_${suffix}.pptx`;
    downloadPptxBase64(mergedPptxBase64, outName);
  }

  async function handleUpload() {
    if (!pptxBase64 || !report?.valid) return;
    setBusyLabel("Saving template...");
    setActionError("");
    setSavedNotice("");
    try {
      const targetUri = await resolveTemplatesGcsUri();
      const res = await callCloudRun("/templates/upload", {
        filename: fileName,
        pptx_base64: pptxBase64,
        dashboard_tiles: tiles.length > 0 ? tiles : undefined,
        templates_gcs_uri: targetUri,
      });
      setReport(res);
      if (res?.gcs_uri) {
        if (selectedDashboardId) {
          void upsertTemplateLinkArtifact(core40SDK, artifactNamespace, fileName, {
            id: selectedDashboardId,
            title: selectedDashboardTitle || selectedDashboardId,
          });
        }
        setSavedNotice(`Saved "${fileName}"`);
      } else {
        setCheckPopoverOpen(true);
      }
    } catch (err: any) {
      setActionError(`Save failed: ${err?.message || err}`);
      setCheckPopoverOpen(true);
    } finally {
      setBusyLabel("");
    }
  }

  const formattedErrorLines = useMemo(() => {
    const lines: string[] = [];
    if (actionError) lines.push(actionError);
    if (Array.isArray(report?.issues)) {
      for (const iss of report.issues) {
        lines.push(`Slide ${iss.slide_index || "?"}: ${iss.message}`);
      }
    }
    return lines;
  }, [actionError, report]);

  async function handleCopyErrors() {
    if (formattedErrorLines.length === 0) return;
    await copyToClipboardSafe(extensionSDK, formattedErrorLines.join("\n"));
    setCopiedErrors(true);
    setTimeout(() => setCopiedErrors(false), 1800);
  }

  const isConfigured = Boolean(
    config.cloudRunUrl && (!config.useOAuth || config.oauthClientId)
  );
  const effectiveRoute = !isConfigured ? "/config" : route;
  const effectiveClientRoute =
    effectiveRoute === "/viewer"
      ? buildViewerClientRoute(viewerFileName || requestedViewerFile, viewerMode, viewerSearch)
      : effectiveRoute === "/scheduler"
        ? buildSchedulerClientRoute(schedulerUserId)
        : effectiveRoute;

  useEffect(() => {
    try {
      extensionSDK?.clientRouteChanged?.(effectiveClientRoute);
    } catch {
      // ignore host route sync errors in standalone mock harness
    }
  }, [effectiveClientRoute]);

  if (effectiveRoute === "/config") {
    return (
      <ConfigPage
        core40SDK={core40SDK}
        extensionSDK={extensionSDK}
        initialConfig={config}
        onSave={(next) => {
          setConfig(next);
          navigateTo("/");
        }}
        onBack={() => {
          navigateTo("/");
        }}
      />
    );
  }

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
          <div style={{ display: "flex", justifyContent: "center", marginBottom: 18 }}>
            <ActionPptxDeckLogo size={52} />
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
          {actionError && (
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
              {actionError}
            </div>
          )}
        </div>
      </div>
    );
  }

  const selectedDashTitle = selectedDashboardTitle || "Open Dashboard";
  const isViewerRoute = effectiveRoute === "/viewer";
  const isSchedulerRoute = effectiveRoute === "/scheduler";
  const hasErrors = Boolean(actionError) || Boolean(report && !report.valid);
  const isCheckPassed = Boolean(report?.valid) && !actionError;
  const showStatusIcon = hasErrors || isCheckPassed;
  const canSaveTemplate = Boolean(pptxBase64 && report?.valid && !busyLabel);
  const isPreviewOpen = Boolean(mergedPptxBase64 || previewLoading);

  const checkSummaryTooltip = hasErrors
    ? formattedErrorLines[0] || "Template check found issues — click for details"
    : isCheckPassed
      ? `Template Ready — Matched ${report.bindings?.length || 0} items across ${report.slide_count} slides (click for details)`
      : "";

  const viewerQuery = viewerSearch.trim().toLowerCase();
  const filteredViewerFiles = viewerFiles
    .filter((f) => !viewerQuery || f.name.toLowerCase().includes(viewerQuery))
    .slice()
    .sort((a, b) => {
      let cmp = 0;
      if (viewerSortKey === "size") cmp = (a.size || 0) - (b.size || 0);
      else if (viewerSortKey === "updated") cmp = (a.updated || "").localeCompare(b.updated || "");
      else cmp = a.name.localeCompare(b.name);
      return viewerSortAsc ? cmp : -cmp;
    });

  function renderSlideCanvasAndNav(minHeight = 460) {
    return (
      <>
        <div
          style={{
            position: "relative",
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            minHeight,
            background: ui.canvasBg,
            borderRadius: 8,
            padding: 12,
            border: `1px solid ${ui.border}`,
            overflow: "hidden",
          }}
        >
          <div ref={previewContainerRef} style={{ display: "flex", justifyContent: "center" }} />

          {previewLoading && (
            <div
              data-testid="preview-loading-overlay"
              style={{
                position: "absolute",
                inset: 0,
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                justifyContent: "center",
                gap: 10,
                background: isDark ? "rgba(9, 13, 22, 0.74)" : "rgba(240, 244, 249, 0.78)",
                backdropFilter: "blur(3px)",
                zIndex: 20,
              }}
            >
              <ActionPptxDeckLogo size={48} pulse />
              <div
                style={{
                  fontSize: 12.5,
                  fontWeight: 700,
                  color: ui.textPrimary,
                  background: ui.cardBg,
                  border: `1px solid ${ui.border}`,
                  padding: "5px 12px",
                  borderRadius: 999,
                }}
              >
                {previewLoadingLabel || "Rendering preview..."}
              </div>
            </div>
          )}
        </div>

        {totalSlides > 0 &&
          (() => {
            const navBtnStyle = (enabled: boolean, primary = false): React.CSSProperties => ({
              display: "inline-flex",
              alignItems: "center",
              gap: 5,
              background: !enabled ? ui.inputBg : primary ? ui.accent : ui.cardBg,
              color: !enabled ? ui.textMuted : primary ? ui.accentContrast : ui.textPrimary,
              border: `1px solid ${enabled && primary ? ui.accent : ui.border}`,
              borderRadius: 6,
              padding: primary ? "6px 14px" : "6px 12px",
              fontSize: 12.5,
              fontWeight: 700,
              whiteSpace: "nowrap",
              cursor: enabled ? "pointer" : "not-allowed",
              opacity: enabled ? 1 : 0.5,
            });
            return (
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  flexWrap: "nowrap",
                  gap: 8,
                  marginTop: 10,
                  padding: "7px 10px",
                  background: ui.inputBg,
                  border: `1px solid ${ui.border}`,
                  borderRadius: 8,
                }}
              >
                <button
                  id="btn-slide-prev"
                  type="button"
                  disabled={currentSlide <= 0}
                  onClick={() => goToSlide(currentSlide - 1)}
                  style={{ ...navBtnStyle(currentSlide > 0), flexShrink: 0 }}
                >
                  <IconArrowLeft />
                  <span>Back</span>
                </button>

                <div
                  ref={slideStripRef}
                  style={{
                    flex: "1 1 0%",
                    minWidth: 0,
                    display: "flex",
                    alignItems: "center",
                    gap: 4,
                    overflowX: "auto",
                    scrollbarWidth: "thin",
                    padding: "2px 0",
                  }}
                >
                  {Array.from({ length: totalSlides }, (_, idx) => {
                    const active = idx === currentSlide;
                    return (
                      <button
                        key={idx}
                        id={`btn-slide-${idx + 1}`}
                        type="button"
                        aria-label={`Slide ${idx + 1} of ${totalSlides}`}
                        aria-current={active ? "true" : undefined}
                        onClick={() => goToSlide(idx)}
                        style={{
                          flexShrink: 0,
                          marginLeft: idx === 0 ? "auto" : 0,
                          marginRight: idx === totalSlides - 1 ? "auto" : 0,
                          minWidth: 30,
                          height: 28,
                          display: "inline-flex",
                          alignItems: "center",
                          justifyContent: "center",
                          background: active ? ui.accent : ui.cardBg,
                          color: active ? ui.accentContrast : ui.textSecondary,
                          border: `1px solid ${active ? ui.accent : ui.border}`,
                          borderRadius: 6,
                          padding: "0 8px",
                          fontSize: 12,
                          fontWeight: 700,
                          whiteSpace: "nowrap",
                          cursor: "pointer",
                        }}
                      >
                        {idx + 1}
                      </button>
                    );
                  })}
                </div>

                <div style={{ display: "flex", gap: 8, alignItems: "center", flexShrink: 0 }}>
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
      </>
    );
  }

  return (
    <div
      style={{
        minHeight: "100vh",
        background: ui.bg,
        color: ui.textPrimary,
        fontFamily: "Inter, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
        padding: "16px 20px",
        transition: "background 0.2s ease, color 0.2s ease",
      }}
    >
      <style>{`
        @keyframes actionPptxPulse {
          0%, 100% { transform: scale(1) translateY(0); opacity: 1; }
          50% { transform: scale(1.08) translateY(-3px); opacity: 0.88; }
        }
      `}</style>

      <div style={{ maxWidth: 960, margin: "0 auto" }}>
        {/* Header */}
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            gap: 12,
            borderBottom: `1px solid ${ui.border}`,
            paddingBottom: 12,
            marginBottom: 14,
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 10, flex: "1 1 0%", minWidth: 0 }}>
            <ActionPptxDeckLogo size={34} />
            <div style={{ minWidth: 0 }}>
              <h1 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: ui.textPrimary }}>
                {isViewerRoute
                  ? "PowerPoint Report Viewer"
                  : isSchedulerRoute
                    ? "PowerPoint Batch Scheduler"
                    : "PowerPoint Report Builder"}
              </h1>
              <p style={{ margin: "2px 0 0", fontSize: 12, color: ui.textMuted }}>
                {isViewerRoute
                  ? "Browse and inspect rendered PowerPoint reports and saved templates."
                  : isSchedulerRoute
                    ? "Run a PowerPoint template across dashboard filter batches on a shared Looker schedule."
                    : "Preview your PowerPoint template with live Looker dashboard data or save it to your shared library."}
              </p>
            </div>
          </div>

          <div
            style={{
              display: "flex",
              gap: 6,
              alignItems: "center",
              flexShrink: 0,
              marginLeft: "auto",
            }}
          >
            <div
              role="navigation"
              aria-label="Extension views"
              style={{
                display: "inline-flex",
                background: ui.inputBg,
                border: `1px solid ${ui.border}`,
                borderRadius: 7,
                padding: 2,
                gap: 2,
              }}
            >
              {(
                [
                  { id: "btn-nav-builder", path: "/", label: "Builder", show: true },
                  {
                    id: "btn-nav-viewer",
                    path: "/viewer",
                    label: "Viewer",
                    show: canOpenViewer || isViewerRoute,
                  },
                  { id: "btn-nav-scheduler", path: "/scheduler", label: "Batches", show: true },
                ] as const
              )
                .filter((item) => item.show)
                .map((item) => {
                  const active = effectiveRoute === item.path;
                  return (
                    <button
                      key={item.path}
                      id={item.id}
                      type="button"
                      aria-current={active ? "page" : undefined}
                      onClick={() =>
                        navigateTo(
                          item.path === "/scheduler"
                            ? buildSchedulerClientRoute(schedulerUserId)
                            : item.path
                        )
                      }
                      style={{
                        border: "none",
                        borderRadius: 5,
                        padding: "5px 10px",
                        fontSize: 12,
                        fontWeight: 700,
                        cursor: "pointer",
                        whiteSpace: "nowrap",
                        background: active ? ui.accent : "transparent",
                        color: active ? ui.accentContrast : ui.textSecondary,
                      }}
                    >
                      {item.label}
                    </button>
                  );
                })}
            </div>

            {!isViewerRoute && !isSchedulerRoute && (
              <LookerThemePicker
                themes={lookerThemes}
                selectedThemeName={selectedThemeName}
                onSelect={handleSelectLookerTheme}
                ui={ui}
                isDark={isDark}
              />
            )}

            {config.useOAuth && oauthToken && (
              <>
                <span
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 4,
                    fontSize: 11.5,
                    fontWeight: 600,
                    color: isDark ? "#34d399" : "#047857",
                    background: isDark ? "#064e3b" : "#d1fae5",
                    border: `1px solid ${isDark ? "#059669" : "#10b981"}`,
                    padding: "4px 9px",
                    borderRadius: 999,
                    whiteSpace: "nowrap",
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
                    gap: 5,
                    background: ui.cardBg,
                    color: ui.textSecondary,
                    border: `1px solid ${ui.border}`,
                    borderRadius: 6,
                    padding: "5px 10px",
                    fontSize: 12,
                    fontWeight: 600,
                    cursor: "pointer",
                    whiteSpace: "nowrap",
                  }}
                >
                  <IconLogOut />
                  <span>Sign Out</span>
                </button>
              </>
            )}
          </div>
        </div>

        {isViewerRoute ? (
          !canOpenViewer ? (
            <div
              style={{
                background: ui.cardBg,
                border: `1px solid ${ui.border}`,
                borderRadius: 10,
                padding: 20,
              }}
            >
              <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 6 }}>
                Report Viewer Unavailable
              </div>
              <p style={{ margin: "0 0 12px", fontSize: 12.5, color: ui.textMuted }}>
                {!serverViewerEnabled
                  ? "The PowerPoint Report Viewer has been disabled by the server configuration."
                  : "Configure an Output GCS Bucket URI in Settings (/config) to enable the Report Viewer."}
              </p>
              <button
                type="button"
                onClick={() => navigateTo("/")}
                style={{
                  background: ui.accent,
                  color: ui.accentContrast,
                  border: "none",
                  borderRadius: 6,
                  padding: "6px 12px",
                  fontSize: 12,
                  fontWeight: 700,
                  cursor: "pointer",
                }}
              >
                Back to Builder
              </button>
            </div>
          ) : (
            <div style={{ display: "grid", gap: 12 }}>
              <div
                style={{
                  background: ui.cardBg,
                  border: `1px solid ${ui.border}`,
                  borderRadius: 10,
                  padding: "10px 12px",
                }}
              >
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 8,
                    marginBottom: 10,
                    flexWrap: "wrap",
                  }}
                >
                  <div
                    role="tablist"
                    aria-label="Viewer source mode"
                    style={{
                      display: "inline-flex",
                      background: ui.inputBg,
                      border: `1px solid ${ui.border}`,
                      borderRadius: 6,
                      padding: 2,
                      gap: 2,
                      flexShrink: 0,
                    }}
                  >
                    {(["outputs", "templates"] as ViewerMode[]).map((mode) => {
                      const active = viewerMode === mode;
                      return (
                        <button
                          key={mode}
                          id={`btn-viewer-mode-${mode}`}
                          type="button"
                          role="tab"
                          aria-selected={active}
                          onClick={() => handleSwitchViewerMode(mode)}
                          style={{
                            border: "none",
                            borderRadius: 4,
                            padding: "5px 10px",
                            fontSize: 12,
                            fontWeight: 700,
                            whiteSpace: "nowrap",
                            cursor: "pointer",
                            background: active ? ui.accent : "transparent",
                            color: active ? ui.accentContrast : ui.textSecondary,
                          }}
                        >
                          {VIEWER_MODE_META[mode].tabLabel}
                        </button>
                      );
                    })}
                  </div>

                  <div
                    style={{
                      flex: "1 1 200px",
                      minWidth: 160,
                      display: "flex",
                      alignItems: "center",
                      gap: 6,
                      background: ui.inputBg,
                      border: `1px solid ${ui.border}`,
                      borderRadius: 6,
                      padding: "5px 10px",
                    }}
                  >
                    <IconSearch />
                    <input
                      id="viewer-search-input"
                      type="text"
                      value={viewerSearch}
                      onChange={(e) => setViewerSearch(e.target.value)}
                      placeholder={VIEWER_MODE_META[viewerMode].searchPlaceholder}
                      style={{
                        flex: 1,
                        border: "none",
                        outline: "none",
                        background: "transparent",
                        color: ui.textPrimary,
                        fontSize: 12.5,
                      }}
                    />
                  </div>

                  <button
                    id="btn-viewer-refresh"
                    type="button"
                    disabled={viewerLoadingList}
                    onClick={() => void loadViewerFiles()}
                    style={{
                      display: "inline-flex",
                      alignItems: "center",
                      gap: 5,
                      background: ui.inputBg,
                      color: ui.textPrimary,
                      border: `1px solid ${ui.border}`,
                      borderRadius: 6,
                      padding: "6px 12px",
                      fontSize: 12,
                      fontWeight: 600,
                      whiteSpace: "nowrap",
                      flexShrink: 0,
                      cursor: viewerLoadingList ? "wait" : "pointer",
                    }}
                  >
                    <IconRefresh />
                    <span>{viewerLoadingList ? "Refreshing..." : "Refresh"}</span>
                  </button>
                </div>

                {actionError && (
                  <div
                    style={{
                      marginBottom: 10,
                      background: isDark ? "#450a0a" : "#fef2f2",
                      border: "1px solid #dc2626",
                      color: isDark ? "#fecaca" : "#991b1b",
                      borderRadius: 6,
                      padding: "7px 10px",
                      fontSize: 12,
                    }}
                  >
                    {actionError}
                  </div>
                )}

                <div
                  style={{
                    border: `1px solid ${ui.border}`,
                    borderRadius: 8,
                    overflow: "auto",
                    maxHeight: 340,
                  }}
                >
                  <table
                    id="viewer-reports-table"
                    style={{
                      width: "100%",
                      borderCollapse: "collapse",
                      fontSize: 12,
                    }}
                  >
                    <thead>
                      <tr
                        style={{
                          background: ui.inputBg,
                          borderBottom: `1px solid ${ui.border}`,
                          textAlign: "left",
                          position: "sticky",
                          top: 0,
                          zIndex: 1,
                        }}
                      >
                        {(
                          [
                            ["name", "Name"],
                            ["size", "Size"],
                            ["updated", VIEWER_MODE_META[viewerMode].dateColumn],
                          ] as const
                        ).map(([key, label]) => (
                          <th key={key} style={{ padding: "7px 10px", fontWeight: 700 }}>
                            <button
                              type="button"
                              onClick={() => toggleViewerSort(key)}
                              style={{
                                background: "none",
                                border: "none",
                                padding: 0,
                                color: ui.textPrimary,
                                font: "inherit",
                                fontWeight: 700,
                                cursor: "pointer",
                              }}
                            >
                              {label}
                              {viewerSortKey === key ? (viewerSortAsc ? " ▲" : " ▼") : ""}
                            </button>
                          </th>
                        ))}
                        <th style={{ padding: "7px 10px", textAlign: "right", fontWeight: 700 }}>
                          Actions
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredViewerFiles.length === 0 ? (
                        <tr>
                          <td
                            colSpan={4}
                            style={{
                              padding: "16px 10px",
                              textAlign: "center",
                              color: ui.textMuted,
                            }}
                          >
                            {viewerLoadingList
                              ? "Loading files..."
                              : viewerQuery
                                ? `No presentations matching "${viewerSearch}".`
                                : VIEWER_MODE_META[viewerMode].emptyText}
                          </td>
                        </tr>
                      ) : (
                        filteredViewerFiles.map((f) => {
                          const isSelected = viewerFileName === f.name;
                          const isLoadingRow = viewerLoadingFile === f.name;
                          const viewerBtnBase: React.CSSProperties = {
                            background: ui.inputBg,
                            border: `1px solid ${ui.border}`,
                            borderRadius: 5,
                            padding: "4px 9px",
                            fontSize: 11.5,
                            fontWeight: 600,
                            cursor: viewerLoadingFile || viewerDeletingFile ? "wait" : "pointer",
                          };
                          return (
                            <tr
                              key={f.name}
                              style={{
                                borderBottom: `1px solid ${ui.border}`,
                                background: isSelected
                                  ? isDark
                                    ? "rgba(59, 130, 246, 0.14)"
                                    : "rgba(37, 99, 235, 0.08)"
                                  : "transparent",
                              }}
                            >
                              <td style={{ padding: "6px 10px", fontWeight: isSelected ? 700 : 500 }}>
                                {f.name}
                              </td>
                              <td style={{ padding: "6px 10px", color: ui.textSecondary, whiteSpace: "nowrap" }}>
                                {formatBytes(f.size)}
                              </td>
                              <td style={{ padding: "6px 10px", color: ui.textSecondary, whiteSpace: "nowrap" }}>
                                {f.updated ? new Date(f.updated).toLocaleString() : "—"}
                              </td>
                              <td style={{ padding: "6px 10px", textAlign: "right", whiteSpace: "nowrap" }}>
                                <div style={{ display: "inline-flex", gap: 6 }}>
                                  <button
                                    type="button"
                                    disabled={Boolean(viewerLoadingFile)}
                                    onClick={() => void handleLoadViewerReport(f.name)}
                                    style={{
                                      ...viewerBtnBase,
                                      minWidth: 68,
                                      display: "inline-flex",
                                      alignItems: "center",
                                      justifyContent: "center",
                                      background: ui.accent,
                                      color: ui.accentContrast,
                                      border: "none",
                                      fontWeight: 700,
                                    }}
                                  >
                                    {isLoadingRow ? "Loading..." : "View"}
                                  </button>
                                  <button
                                    type="button"
                                    disabled={Boolean(viewerLoadingFile)}
                                    onClick={() => void handleDownloadViewerReport(f.name)}
                                    style={{ ...viewerBtnBase, color: ui.textPrimary }}
                                  >
                                    Download
                                  </button>
                                  {viewerMode === "templates" &&
                                    (confirmDeleteViewerFile === f.name ? (
                                      <>
                                        <button
                                          type="button"
                                          data-testid={`btn-confirm-delete-template-${f.name}`}
                                          disabled={Boolean(viewerDeletingFile)}
                                          onClick={() => void handleDeleteViewerTemplate(f.name)}
                                          style={{
                                            ...viewerBtnBase,
                                            background: "#dc2626",
                                            color: "#ffffff",
                                            border: "none",
                                            fontWeight: 700,
                                          }}
                                        >
                                          {viewerDeletingFile === f.name ? "Deleting..." : "Confirm"}
                                        </button>
                                        <button
                                          type="button"
                                          disabled={Boolean(viewerDeletingFile)}
                                          onClick={() => setConfirmDeleteViewerFile("")}
                                          style={{ ...viewerBtnBase, color: ui.textSecondary }}
                                        >
                                          Cancel
                                        </button>
                                      </>
                                    ) : (
                                      <button
                                        type="button"
                                        data-testid={`btn-delete-template-${f.name}`}
                                        disabled={Boolean(viewerLoadingFile || viewerDeletingFile)}
                                        onClick={() => setConfirmDeleteViewerFile(f.name)}
                                        style={{ ...viewerBtnBase, color: "#ef4444" }}
                                      >
                                        Delete
                                      </button>
                                    ))}
                                </div>
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>
              </div>

              {(viewerPptxBase64 || previewLoading) && (
                <div
                  style={{
                    background: ui.cardBg,
                    border: `1px solid ${ui.border}`,
                    borderRadius: 10,
                    padding: "10px 12px",
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                      flexWrap: "wrap",
                      gap: 8,
                      marginBottom: 10,
                    }}
                  >
                    <div style={{ fontSize: 13, fontWeight: 700, color: ui.textPrimary }}>
                      {viewerFileName || "Presentation Preview"}
                    </div>
                    <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                      {viewerPptxBase64 && (
                        <button
                          id="btn-viewer-download-active"
                          type="button"
                          onClick={() => downloadPptxBase64(viewerPptxBase64, viewerFileName || "report.pptx")}
                          style={{
                            display: "inline-flex",
                            alignItems: "center",
                            gap: 5,
                            background: ui.inputBg,
                            color: ui.accent,
                            border: `1px solid ${ui.accent}`,
                            borderRadius: 6,
                            padding: "5px 10px",
                            fontSize: 12,
                            fontWeight: 600,
                            cursor: "pointer",
                          }}
                        >
                          <IconDownload />
                          <span>Download .pptx</span>
                        </button>
                      )}
                      <button
                        id="btn-viewer-close"
                        type="button"
                        onClick={handleEndPreview}
                        style={{
                          display: "inline-flex",
                          alignItems: "center",
                          gap: 5,
                          background: isDark ? "#1e293b" : "#fee2e2",
                          color: isDark ? "#fda4af" : "#b91c1c",
                          border: `1px solid ${isDark ? "#881337" : "#fca5a5"}`,
                          borderRadius: 6,
                          padding: "5px 10px",
                          fontSize: 12,
                          fontWeight: 700,
                          cursor: "pointer",
                        }}
                      >
                        <IconClose />
                        <span>Close</span>
                      </button>
                    </div>
                  </div>
                  {renderSlideCanvasAndNav(440)}
                </div>
              )}
            </div>
          )
        ) : isSchedulerRoute ? (
          <SchedulerView
            core40SDK={core40SDK}
            extensionSDK={extensionSDK}
            ui={ui}
            isDark={isDark}
            lookerThemes={lookerThemes}
            actionIntegrationId={config.actionIntegrationId}
            targetUserIdOverride={schedulerUserId || undefined}
            seed={schedulerSeed}
            onClearSeed={() => setSchedulerSeed(null)}
            callCloudRun={callCloudRun}
            resolveTemplatesGcsUri={resolveTemplatesGcsUri}
            canOpenViewer={canOpenViewer}
            onOpenViewerSearch={(query) => {
              navigateTo(`/viewer?q=${encodeURIComponent(query)}`);
            }}
          />
        ) : (
          <>

        {/* Step 1 & Step 2 Cards */}
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1fr)",
            gap: 10,
            marginBottom: 10,
          }}
        >
          <div
            style={{
              background: ui.cardBg,
              border: `1px solid ${ui.border}`,
              borderRadius: 10,
              padding: 12,
              minWidth: 0,
            }}
          >
            <div style={{ fontSize: 10.5, fontWeight: 700, color: ui.accent, marginBottom: 2 }}>
              STEP 1
            </div>
            <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 8 }}>
              Choose a Dashboard
            </div>
            <DashboardPicker
              core40SDK={core40SDK}
              extensionSDK={extensionSDK}
              selectedDashboardId={selectedDashboardId}
              onSelect={(id, summary) => void handleSelectDashboard(id, summary)}
              ui={ui}
              isDark={isDark}
            />
            {tiles.length > 0 && (
              <div style={{ marginTop: 8, display: "flex", flexWrap: "wrap", gap: 5 }}>
                {tiles.slice(0, 9).map((t) => (
                  <span
                    key={t.slug}
                    style={{
                      fontSize: 10.5,
                      background: ui.inputBg,
                      color: ui.textSecondary,
                      border: `1px solid ${ui.border}`,
                      padding: "2px 7px",
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
                      fontSize: 10.5,
                      background: ui.inputBg,
                      color: ui.accent,
                      border: `1px solid ${ui.border}`,
                      padding: "2px 7px",
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
              borderRadius: 10,
              padding: 12,
              minWidth: 0,
            }}
          >
            <div style={{ fontSize: 10.5, fontWeight: 700, color: ui.accent, marginBottom: 2 }}>
              STEP 2
            </div>
            <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 8 }}>
              Select PowerPoint Template (.pptx)
            </div>
            <input
              id="pptx-file-input"
              type="file"
              accept=".pptx"
              onChange={(e) => void handleFileSelect(e)}
              style={{
                width: "100%",
                padding: "6px 8px",
                borderRadius: 6,
                background: ui.inputBg,
                color: ui.textSecondary,
                border: `1px solid ${ui.border}`,
                fontSize: 12.5,
                boxSizing: "border-box",
              }}
            />
            {fileName && (
              <div style={{ marginTop: 8, fontSize: 11.5, color: ui.textMuted }}>
                Selected template:{" "}
                <strong style={{ color: ui.textPrimary }}>{fileName}</strong>
              </div>
            )}
          </div>
        </div>

        {/* Single Unified Workspace Card: Action Toolbar + Sticky Check/Error Icon + Live Preview */}
        <div
          style={{
            background: ui.cardBg,
            border: `1px solid ${ui.border}`,
            borderRadius: 10,
            padding: "10px 12px",
            marginBottom: 14,
          }}
        >
          {/* Unified Top Action Bar */}
          <div
            style={{
              display: "flex",
              flexWrap: "wrap",
              gap: 8,
              alignItems: "center",
              justifyContent: "space-between",
            }}
          >
            {/* Left Action Cluster */}
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6, alignItems: "center" }}>
              <button
                id="btn-preview"
                type="button"
                disabled={!pptxBase64 || !selectedDashboardId || previewLoading}
                onClick={() => void runPreview()}
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: 5,
                  background: ui.accent,
                  color: ui.accentContrast,
                  border: "none",
                  borderRadius: 6,
                  padding: "6px 12px",
                  fontSize: 12,
                  fontWeight: 700,
                  whiteSpace: "nowrap",
                  cursor:
                    pptxBase64 && selectedDashboardId && !previewLoading
                      ? "pointer"
                      : "not-allowed",
                  opacity: pptxBase64 && selectedDashboardId ? 1 : 0.5,
                }}
              >
                <IconPlay />
                <span>Live Preview</span>
              </button>

              <button
                id="btn-validate"
                type="button"
                disabled={!pptxBase64 || Boolean(busyLabel)}
                onClick={() => void handleValidate()}
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: 5,
                  background: ui.inputBg,
                  color: ui.textPrimary,
                  border: `1px solid ${ui.border}`,
                  borderRadius: 6,
                  padding: "6px 11px",
                  fontSize: 12,
                  fontWeight: 600,
                  whiteSpace: "nowrap",
                  cursor: pptxBase64 && !busyLabel ? "pointer" : "not-allowed",
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
                    gap: 5,
                    background: ui.inputBg,
                    color: ui.accent,
                    border: `1px solid ${ui.accent}`,
                    borderRadius: 6,
                    padding: "6px 11px",
                    fontSize: 12,
                    fontWeight: 600,
                    whiteSpace: "nowrap",
                    cursor: "pointer",
                  }}
                >
                  <IconDownload />
                  <span>Download Filled .pptx</span>
                </button>
              )}

              {isPreviewOpen && selectedDashboardId && (
                <button
                  type="button"
                  data-testid="preview-dashboard-link"
                  onClick={() =>
                    extensionSDK?.openBrowserWindow?.(
                      `/dashboards/${encodeURIComponent(selectedDashboardId)}`,
                      "_blank"
                    )
                  }
                  title="Open source dashboard in Looker"
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 5,
                    background: "transparent",
                    border: "none",
                    color: ui.accent,
                    fontSize: 13,
                    fontWeight: 700,
                    padding: "4px 6px",
                    cursor: "pointer",
                    whiteSpace: "nowrap",
                  }}
                >
                  <span>{selectedDashTitle}</span>
                  <IconExternal />
                </button>
              )}
            </div>

            {/* Right Status + Save + Close Cluster */}
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 9,
                marginLeft: "auto",
              }}
            >
              {busyLabel && (
                <span
                  style={{
                    fontSize: 12,
                    fontWeight: 600,
                    color: ui.textMuted,
                    whiteSpace: "nowrap",
                  }}
                >
                  {busyLabel}
                </span>
              )}

              {savedNotice && !busyLabel && (
                <span
                  style={{
                    fontSize: 12,
                    fontWeight: 600,
                    color: isDark ? "#34d399" : "#047857",
                    whiteSpace: "nowrap",
                  }}
                >
                  {savedNotice}
                </span>
              )}

              {/* Clickable Green Check / Red X Icon with Hover Tooltip & Sticky Click-Away Popover */}
              {showStatusIcon && (
                <div ref={checkPopoverRef} style={{ position: "relative", display: "inline-flex" }}>
                  <button
                    id="btn-validation-status"
                    data-testid="btn-validation-status"
                    type="button"
                    title={checkSummaryTooltip}
                    aria-label={checkSummaryTooltip}
                    aria-expanded={checkPopoverOpen}
                    onClick={() => setCheckPopoverOpen((v) => !v)}
                    style={{
                      display: "inline-flex",
                      alignItems: "center",
                      justifyContent: "center",
                      gap: 5,
                      height: 36,
                      minWidth: 36,
                      padding: formattedErrorLines.length > 1 ? "0 10px" : "0 9px",
                      borderRadius: 999,
                      cursor: "pointer",
                      background: hasErrors
                        ? isDark
                          ? "#450a0a"
                          : "#fee2e2"
                        : isDark
                          ? "#064e3b"
                          : "#d1fae5",
                      color: hasErrors
                        ? isDark
                          ? "#fca5a5"
                          : "#b91c1c"
                        : isDark
                          ? "#34d399"
                          : "#047857",
                      border: `1px solid ${
                        hasErrors
                          ? isDark
                            ? "#dc2626"
                            : "#f87171"
                          : isDark
                            ? "#059669"
                            : "#10b981"
                      }`,
                      boxShadow: checkPopoverOpen
                        ? `0 0 0 2px ${hasErrors ? "#ef444455" : "#10b98155"}`
                        : "none",
                    }}
                  >
                    {hasErrors ? <IconXCircle /> : <IconCheckCircle />}
                    {hasErrors && formattedErrorLines.length > 0 && (
                      <span style={{ fontSize: 11.5, fontWeight: 700 }}>
                        {formattedErrorLines.length}
                      </span>
                    )}
                  </button>

                  {checkPopoverOpen && (
                    <div
                      role="dialog"
                      data-testid="validation-sticky-popover"
                      style={{
                        position: "absolute",
                        top: "calc(100% + 8px)",
                        right: 0,
                        width: "min(440px, 88vw)",
                        zIndex: 80,
                        background: ui.cardBg,
                        border: `1px solid ${hasErrors ? "#dc2626" : ui.border}`,
                        borderRadius: 12,
                        padding: 14,
                        boxShadow: isDark
                          ? "0 22px 48px rgba(0, 0, 0, 0.68)"
                          : "0 18px 40px rgba(15, 23, 42, 0.2)",
                      }}
                    >
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "space-between",
                          gap: 8,
                          marginBottom: hasErrors ? 10 : 4,
                        }}
                      >
                        <div
                          style={{
                            display: "inline-flex",
                            alignItems: "center",
                            gap: 7,
                            fontSize: 13.5,
                            fontWeight: 700,
                            color: hasErrors
                              ? isDark
                                ? "#fca5a5"
                                : "#b91c1c"
                              : isDark
                                ? "#34d399"
                                : "#047857",
                          }}
                        >
                          {hasErrors ? <IconXCircle /> : <IconCheckCircle />}
                          <span>
                            {hasErrors
                              ? `Template Check (${formattedErrorLines.length} issue${
                                  formattedErrorLines.length === 1 ? "" : "s"
                                })`
                              : "Template Ready"}
                          </span>
                        </div>

                        <div style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
                          {hasErrors && formattedErrorLines.length > 0 && (
                            <button
                              type="button"
                              data-testid="btn-copy-errors"
                              onClick={() => void handleCopyErrors()}
                              style={{
                                display: "inline-flex",
                                alignItems: "center",
                                gap: 5,
                                padding: "4px 9px",
                                borderRadius: 6,
                                fontSize: 11.5,
                                fontWeight: 600,
                                cursor: "pointer",
                                background: ui.inputBg,
                                color: ui.textPrimary,
                                border: `1px solid ${ui.border}`,
                                whiteSpace: "nowrap",
                              }}
                            >
                              {copiedErrors ? <IconCheck /> : <IconCopy />}
                              <span>{copiedErrors ? "Copied" : "Copy"}</span>
                            </button>
                          )}
                          <button
                            type="button"
                            aria-label="Close details"
                            onClick={() => setCheckPopoverOpen(false)}
                            style={{
                              background: "transparent",
                              border: "none",
                              color: ui.textMuted,
                              cursor: "pointer",
                              padding: 3,
                              display: "inline-flex",
                            }}
                          >
                            <IconClose />
                          </button>
                        </div>
                      </div>

                      {!hasErrors && report?.valid && (
                        <div style={{ fontSize: 12.5, color: ui.textSecondary, lineHeight: 1.5 }}>
                          Matched <strong>{report.bindings?.length || 0}</strong> charts, tables,
                          and text placeholders across <strong>{report.slide_count}</strong>{" "}
                          slide{report.slide_count === 1 ? "" : "s"}. Ready to save to your
                          shared Template Library.
                        </div>
                      )}

                      {hasErrors && (
                        <div
                          style={{
                            maxHeight: 220,
                            overflowY: "auto",
                            background: isDark ? "#450a0a" : "#fef2f2",
                            border: "1px solid #dc2626",
                            borderRadius: 8,
                            padding: "9px 11px",
                            display: "grid",
                            gap: 6,
                          }}
                        >
                          {formattedErrorLines.map((line, idx) => (
                            <div
                              key={idx}
                              style={{
                                fontSize: 12.5,
                                lineHeight: 1.45,
                                color: isDark ? "#fecaca" : "#991b1b",
                                wordBreak: "break-word",
                              }}
                            >
                              {line}
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                </div>
              )}

              <button
                id="btn-upload"
                type="button"
                disabled={!canSaveTemplate}
                title={
                  !pptxBase64
                    ? "Select a .pptx template first"
                    : !report
                      ? "Run Check Template (or Live Preview) before saving"
                      : !report.valid
                        ? "Resolve template check errors before saving"
                        : "Save validated template to shared library"
                }
                onClick={() => void handleUpload()}
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: 5,
                  background: "#059669",
                  color: "#ffffff",
                  border: "none",
                  borderRadius: 6,
                  padding: "6px 12px",
                  fontSize: 12,
                  fontWeight: 700,
                  whiteSpace: "nowrap",
                  cursor: canSaveTemplate ? "pointer" : "not-allowed",
                  opacity: canSaveTemplate ? 1 : 0.45,
                }}
              >
                <IconUpload />
                <span>Save to Template Library</span>
              </button>

              {(savedNotice || (fileName && report?.valid)) && (
                <button
                  id="btn-schedule-template"
                  type="button"
                  onClick={() => {
                    setSchedulerSeed({
                      dashboardId: selectedDashboardId || undefined,
                      dashboardTitle: selectedDashboardTitle || undefined,
                      templateName: fileName || undefined,
                      themeName: selectedThemeName || undefined,
                    });
                    navigateTo(buildSchedulerClientRoute(schedulerUserId));
                  }}
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 5,
                    background: ui.inputBg,
                    color: ui.accent,
                    border: `1px solid ${ui.accent}`,
                    borderRadius: 6,
                    padding: "6px 11px",
                    fontSize: 12,
                    fontWeight: 700,
                    whiteSpace: "nowrap",
                    cursor: "pointer",
                  }}
                >
                  <span>Schedule Batch →</span>
                </button>
              )}

              {isPreviewOpen && (
                <button
                  id="btn-close-preview"
                  type="button"
                  onClick={handleEndPreview}
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 5,
                    background: isDark ? "#1e293b" : "#fee2e2",
                    color: isDark ? "#fda4af" : "#b91c1c",
                    border: `1px solid ${isDark ? "#881337" : "#fca5a5"}`,
                    borderRadius: 6,
                    padding: "6px 10px",
                    fontSize: 12,
                    fontWeight: 700,
                    whiteSpace: "nowrap",
                    cursor: "pointer",
                  }}
                >
                  <IconClose />
                  <span>Close Preview</span>
                </button>
              )}
            </div>
          </div>

          {/* Embedded Live Preview inside the Same Unified Card */}
          {isPreviewOpen && (
            <div
              style={{
                borderTop: `1px solid ${ui.border}`,
                marginTop: 10,
                paddingTop: 12,
              }}
            >
              {renderSlideCanvasAndNav(440)}
            </div>
          )}
        </div>
          </>
        )}
      </div>
    </div>
  );
};
