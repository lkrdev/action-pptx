import React, { useEffect, useMemo, useRef, useState } from "react";
import { Svg } from "./ConfigPage";

export interface DashboardSummary {
  id: string;
  title: string;
  isLookml: boolean;
  folderId?: string;
  folderName?: string;
  isPersonal?: boolean;
  modelId?: string;
  modelLabel?: string;
  userId?: string;
  lastViewedAt?: string;
  viewCount?: number;
  isFavorite?: boolean;
  description?: string;
}

export interface PickerPalette {
  bg: string;
  cardBg: string;
  inputBg: string;
  border: string;
  textPrimary: string;
  textSecondary: string;
  textMuted: string;
  accent: string;
  accentContrast: string;
}

type FilterScope = "recent" | "personal" | "owned" | "shared" | "lookml" | "all";

const DASH_SEARCH_FIELDS =
  "id,title,folder(id,name,is_personal,is_personal_descendant),user_id,last_viewed_at,view_count,content_favorite_id,description";
const ALL_DASH_FIELDS = "id,title,folder(id,name),model(id,label),description";

const IconSearch = () => (
  <Svg size={15}>
    <circle cx="11" cy="11" r="8" />
    <path d="m21 21-4.3-4.3" />
  </Svg>
);
const IconClock = () => (
  <Svg size={13}>
    <circle cx="12" cy="12" r="9" />
    <path d="M12 7v5l3 2" />
  </Svg>
);
const IconUser = () => (
  <Svg size={13}>
    <circle cx="12" cy="8" r="4" />
    <path d="M20 21a8 8 0 0 0-16 0" />
  </Svg>
);
const IconFolder = () => (
  <Svg size={13}>
    <path d="M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z" />
  </Svg>
);
const IconCode = () => (
  <Svg size={13}>
    <polyline points="16 18 22 12 16 6" />
    <polyline points="8 6 2 12 8 18" />
  </Svg>
);
const IconSpark = () => (
  <Svg size={13}>
    <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2" />
  </Svg>
);
const IconExternal = () => (
  <Svg size={14}>
    <path d="M15 3h6v6M10 14 21 3M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6" />
  </Svg>
);
const IconChevronDown = () => (
  <Svg size={14}>
    <path d="m6 9 6 6 6-6" />
  </Svg>
);
const IconX = () => (
  <Svg size={13}>
    <path d="M18 6 6 18M6 6l12 12" />
  </Svg>
);
const IconCheckSmall = () => (
  <Svg size={14}>
    <path d="M20 6 9 17l-5-5" />
  </Svg>
);

function normalizeDashboard(raw: any, existing?: DashboardSummary): DashboardSummary | null {
  const id = String(raw?.id ?? existing?.id ?? "").trim();
  const title = String(raw?.title ?? existing?.title ?? "").trim();
  if (!id || !title) return null;
  const isLookml = id.includes("::") || raw?.folder?.id === "lookml";
  const folderObj = raw?.folder;
  const modelObj = raw?.model;
  return {
    id,
    title,
    isLookml,
    folderId: String(folderObj?.id ?? existing?.folderId ?? "") || undefined,
    folderName: String(folderObj?.name ?? existing?.folderName ?? "").trim() || undefined,
    isPersonal:
      typeof folderObj?.is_personal === "boolean" ||
      typeof folderObj?.is_personal_descendant === "boolean"
        ? Boolean(folderObj?.is_personal || folderObj?.is_personal_descendant)
        : existing?.isPersonal,
    modelId:
      String(modelObj?.id ?? (isLookml ? id.split("::")[0] : "") ?? existing?.modelId ?? "") ||
      undefined,
    modelLabel:
      String(
        modelObj?.label ?? (isLookml ? id.split("::")[0] : "") ?? existing?.modelLabel ?? ""
      ).trim() || undefined,
    userId: raw?.user_id != null ? String(raw.user_id) : existing?.userId,
    lastViewedAt: raw?.last_viewed_at ? String(raw.last_viewed_at) : existing?.lastViewedAt,
    viewCount:
      typeof raw?.view_count === "number" ? raw.view_count : existing?.viewCount,
    isFavorite:
      raw?.content_favorite_id != null ? Boolean(raw.content_favorite_id) : existing?.isFavorite,
    description: String(raw?.description ?? existing?.description ?? "").trim() || undefined,
  };
}

function extractPastedDashboardId(input: string): string | null {
  const trimmed = input.trim();
  if (!trimmed) return null;
  const urlMatch = trimmed.match(/\/dashboards(?:-next)?\/([^/?#\s]+)/i);
  if (urlMatch?.[1]) return decodeURIComponent(urlMatch[1]);
  if (/^#\d+$/.test(trimmed)) return trimmed.slice(1);
  if (/^[a-z0-9_-]+::[a-z0-9_-]+$/i.test(trimmed)) return trimmed;
  return null;
}

function formatShortDate(iso?: string): string {
  if (!iso) return "";
  const d = new Date(iso.replace(" ", "T"));
  if (Number.isNaN(d.getTime())) return "";
  return d.toLocaleDateString(undefined, { month: "short", day: "numeric" });
}

export function useDismissOnOutsideOrEscape(
  open: boolean,
  ref: React.RefObject<HTMLElement | null>,
  onClose: () => void
): void {
  useEffect(() => {
    if (!open) return;
    const onDocDown = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) onClose();
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
    };
    document.addEventListener("mousedown", onDocDown);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDocDown);
      document.removeEventListener("keydown", onKey);
    };
  }, [open]);
}

interface DashboardPickerProps {
  core40SDK: any;
  extensionSDK: any;
  selectedDashboardId: string;
  onSelect: (id: string, summary?: DashboardSummary) => void;
  ui: PickerPalette;
  isDark: boolean;
}

export const DashboardPicker: React.FC<DashboardPickerProps> = ({
  core40SDK,
  extensionSDK,
  selectedDashboardId,
  onSelect,
  ui,
  isDark,
}) => {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [scope, setScope] = useState<FilterScope>("recent");
  const [loading, setLoading] = useState(true);
  const [searchingRemote, setSearchingRemote] = useState(false);
  const [meInfo, setMeInfo] = useState<{ id: string; personalFolderId: string }>({
    id: "",
    personalFolderId: "",
  });
  const [catalog, setCatalog] = useState<Map<string, DashboardSummary>>(new Map());
  const [activeIdx, setActiveIdx] = useState(0);
  const wrapRef = useRef<HTMLDivElement | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      setLoading(true);
      try {
        const [meRes, recentRes, allRes] = await Promise.all([
          typeof core40SDK.me === "function"
            ? core40SDK.ok(core40SDK.me("id,personal_folder_id")).catch(() => null)
            : Promise.resolve(null),
          typeof core40SDK.search_dashboards === "function"
            ? core40SDK
                .ok(
                  core40SDK.search_dashboards({
                    sorts: "last_viewed_at desc",
                    limit: 150,
                    fields: DASH_SEARCH_FIELDS,
                  })
                )
                .catch(() => [])
            : Promise.resolve([]),
          typeof core40SDK.all_dashboards === "function"
            ? core40SDK.ok(core40SDK.all_dashboards(ALL_DASH_FIELDS)).catch(() => [])
            : Promise.resolve([]),
        ]);
        if (cancelled) return;
        if (meRes) {
          setMeInfo({
            id: String(meRes.id || ""),
            personalFolderId: String(meRes.personal_folder_id || ""),
          });
        }
        const merged = new Map<string, DashboardSummary>();
        for (const item of Array.isArray(allRes) ? allRes : []) {
          const norm = normalizeDashboard(item, merged.get(String(item?.id || "")));
          if (norm) merged.set(norm.id, norm);
        }
        for (const item of Array.isArray(recentRes) ? recentRes : []) {
          const norm = normalizeDashboard(item, merged.get(String(item?.id || "")));
          if (norm) merged.set(norm.id, norm);
        }
        setCatalog(merged);
        const hasRecent = Array.from(merged.values()).some((d) => Boolean(d.lastViewedAt));
        if (!hasRecent) setScope("all");
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  // Debounced server-side search_dashboards when typing >= 2 chars
  useEffect(() => {
    const q = query.trim();
    if (q.length < 2 || extractPastedDashboardId(q) || typeof core40SDK.search_dashboards !== "function") {
      setSearchingRemote(false);
      return;
    }
    let cancelled = false;
    setSearchingRemote(true);
    const timer = setTimeout(() => {
      void core40SDK
        .ok(
          core40SDK.search_dashboards({
            title: `%${q.replace(/\s+/g, "%")}%`,
            limit: 50,
            fields: DASH_SEARCH_FIELDS,
          })
        )
        .then((rows: any[]) => {
          if (cancelled || !Array.isArray(rows)) return;
          setCatalog((prev) => {
            const next = new Map(prev);
            for (const r of rows) {
              const norm = normalizeDashboard(r, next.get(String(r?.id || "")));
              if (norm) next.set(norm.id, norm);
            }
            return next;
          });
        })
        .catch(() => {})
        .finally(() => {
          if (!cancelled) setSearchingRemote(false);
        });
    }, 260);
    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, [query]);

  useDismissOnOutsideOrEscape(open, wrapRef, () => {
    setOpen(false);
    setQuery("");
  });

  const allItems = useMemo(() => Array.from(catalog.values()), [catalog]);
  const selectedItem = selectedDashboardId ? catalog.get(selectedDashboardId) : undefined;

  const counts = useMemo(() => {
    let recent = 0;
    let personal = 0;
    let owned = 0;
    let shared = 0;
    let lookml = 0;
    for (const d of allItems) {
      if (d.lastViewedAt) recent++;
      if (d.isLookml) {
        lookml++;
      } else {
        if (
          (meInfo.personalFolderId && d.folderId === meInfo.personalFolderId) ||
          d.isPersonal === true
        ) {
          personal++;
        } else {
          shared++;
        }
        if (meInfo.id && d.userId === meInfo.id) owned++;
      }
    }
    return { recent, personal, owned, shared, lookml, all: allItems.length };
  }, [allItems, meInfo]);

  const effectiveScope: FilterScope =
    query.trim().length > 0 && scope === "recent" ? "all" : scope;

  const filteredItems = useMemo(() => {
    const q = query.trim().toLowerCase();
    const pastedId = extractPastedDashboardId(query);

    const byScope = allItems.filter((d) => {
      if (pastedId && d.id.toLowerCase() === pastedId.toLowerCase()) return true;
      switch (effectiveScope) {
        case "recent":
          return Boolean(d.lastViewedAt) || d.isLookml;
        case "personal":
          return (
            !d.isLookml &&
            ((Boolean(meInfo.personalFolderId) && d.folderId === meInfo.personalFolderId) ||
              d.isPersonal === true)
          );
        case "owned":
          return !d.isLookml && Boolean(meInfo.id) && d.userId === meInfo.id;
        case "shared":
          return (
            !d.isLookml &&
            d.isPersonal !== true &&
            (!meInfo.personalFolderId || d.folderId !== meInfo.personalFolderId)
          );
        case "lookml":
          return d.isLookml;
        case "all":
        default:
          return true;
      }
    });

    const matched = q
      ? byScope.filter((d) => {
          if (pastedId && d.id.toLowerCase() === pastedId.toLowerCase()) return true;
          const hay =
            `${d.title} ${d.id} ${d.folderName || ""} ${d.modelId || ""} ${d.modelLabel || ""} ${d.description || ""} ${d.isLookml ? "lookml" : ""}`.toLowerCase();
          return q.split(/\s+/).every((tok) => hay.includes(tok));
        })
      : byScope;

    const sorted = [...matched];
    if (effectiveScope === "recent" && !q) {
      sorted.sort((a, b) => {
        if (a.isLookml !== b.isLookml) return a.isLookml ? 1 : -1;
        if (!a.isLookml) {
          return String(b.lastViewedAt || "").localeCompare(String(a.lastViewedAt || ""));
        }
        const mCmp = (a.modelLabel || a.id).localeCompare(b.modelLabel || b.id);
        return mCmp !== 0 ? mCmp : a.title.localeCompare(b.title);
      });
    } else {
      sorted.sort((a, b) => {
        const groupA = a.isLookml ? a.modelLabel || "LookML" : a.folderName || "Dashboards";
        const groupB = b.isLookml ? b.modelLabel || "LookML" : b.folderName || "Dashboards";
        if (effectiveScope === "shared" || effectiveScope === "lookml") {
          const gCmp = groupA.localeCompare(groupB);
          if (gCmp !== 0) return gCmp;
        }
        return a.title.localeCompare(b.title);
      });
    }
    return sorted.slice(0, 120);
  }, [allItems, effectiveScope, meInfo, query]);

  const pastedDirectId = useMemo(() => {
    const pid = extractPastedDashboardId(query);
    if (!pid) return null;
    return catalog.has(pid) ? null : pid;
  }, [query, catalog]);

  useEffect(() => {
    setActiveIdx(0);
  }, [query, effectiveScope]);

  function chooseDashboard(id: string, summary?: DashboardSummary) {
    onSelect(id, summary);
    setOpen(false);
    setQuery("");
  }

  function handleKeyDown(e: React.KeyboardEvent) {
    if (e.key === "ArrowDown") {
      e.preventDefault();
      if (!open) {
        setOpen(true);
        return;
      }
      setActiveIdx((i) => Math.min(Math.max(filteredItems.length - 1, 0), i + 1));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      setActiveIdx((i) => Math.max(0, i - 1));
    } else if (e.key === "Enter" && open) {
      e.preventDefault();
      if (pastedDirectId) {
        chooseDashboard(pastedDirectId);
      } else if (filteredItems[activeIdx]) {
        chooseDashboard(filteredItems[activeIdx].id, filteredItems[activeIdx]);
      }
    }
  }

  const tabs: Array<{ id: FilterScope; label: string; count: number; icon: React.ReactNode }> = [
    { id: "recent", label: "Recent", count: counts.recent, icon: <IconClock /> },
    { id: "personal", label: "My Folder", count: counts.personal, icon: <IconFolder /> },
    { id: "owned", label: "Owned by Me", count: counts.owned, icon: <IconUser /> },
    { id: "shared", label: "Shared", count: counts.shared, icon: <IconSpark /> },
    { id: "lookml", label: "LookML", count: counts.lookml, icon: <IconCode /> },
    { id: "all", label: "All", count: counts.all, icon: null },
  ];

  const selectedBadge = selectedItem
    ? selectedItem.isLookml
      ? `LookML · ${selectedItem.modelLabel || selectedItem.id.split("::")[0]}`
      : `${selectedItem.folderName || "Folder"} · #${selectedItem.id}`
    : selectedDashboardId
      ? selectedDashboardId
      : "";

  return (
    <div ref={wrapRef} style={{ position: "relative", width: "100%", minWidth: 0 }}>
      {/* Search Bar / Selected Dashboard Trigger */}
      <div style={{ display: "flex", gap: 6, alignItems: "center", width: "100%", minWidth: 0 }}>
        <div
          onClick={() => {
            setOpen(true);
            setTimeout(() => inputRef.current?.focus(), 0);
          }}
          style={{
            flex: 1,
            minWidth: 0,
            display: "flex",
            alignItems: "center",
            gap: 8,
            padding: "8px 11px",
            borderRadius: 9,
            background: ui.inputBg,
            border: `1px solid ${open ? ui.accent : ui.border}`,
            boxShadow: open ? `0 0 0 2px ${ui.accent}33` : "none",
            cursor: "text",
            transition: "border-color 0.15s ease, box-shadow 0.15s ease",
          }}
        >
          <span style={{ color: open ? ui.accent : ui.textMuted, display: "inline-flex", flexShrink: 0 }}>
            <IconSearch />
          </span>

          {open || !selectedDashboardId ? (
            <input
              ref={inputRef}
              data-testid="dashboard-search-input"
              type="text"
              value={query}
              onFocus={() => setOpen(true)}
              onChange={(e) => {
                setQuery(e.target.value);
                if (!open) setOpen(true);
              }}
              onKeyDown={handleKeyDown}
              placeholder={
                loading
                  ? "Loading Looker dashboards..."
                  : selectedItem
                    ? `${selectedItem.title} — type to search or switch...`
                    : "Search by title, folder, LookML model, or paste URL..."
              }
              style={{
                flex: 1,
                minWidth: 0,
                background: "transparent",
                border: "none",
                outline: "none",
                color: ui.textPrimary,
                fontSize: 13,
                fontFamily: "inherit",
              }}
            />
          ) : (
            <div
              style={{
                flex: 1,
                minWidth: 0,
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                gap: 8,
              }}
            >
              <span
                style={{
                  fontSize: 13,
                  fontWeight: 600,
                  color: ui.textPrimary,
                  overflow: "hidden",
                  textOverflow: "ellipsis",
                  whiteSpace: "nowrap",
                }}
              >
                {selectedItem?.title || selectedDashboardId}
              </span>
              {selectedBadge && (
                <span
                  style={{
                    flexShrink: 0,
                    fontSize: 11,
                    fontWeight: 600,
                    padding: "2px 8px",
                    borderRadius: 999,
                    background: isDark ? "rgba(148, 163, 184, 0.14)" : "rgba(15, 23, 42, 0.06)",
                    color: ui.textMuted,
                    maxWidth: 170,
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                    whiteSpace: "nowrap",
                  }}
                >
                  {selectedBadge}
                </span>
              )}
            </div>
          )}

          {query ? (
            <button
              type="button"
              aria-label="Clear search"
              onClick={(e) => {
                e.stopPropagation();
                setQuery("");
                inputRef.current?.focus();
              }}
              style={{
                background: "transparent",
                border: "none",
                color: ui.textMuted,
                cursor: "pointer",
                padding: 2,
                display: "inline-flex",
              }}
            >
              <IconX />
            </button>
          ) : (
            <span style={{ color: ui.textMuted, display: "inline-flex", flexShrink: 0 }}>
              <IconChevronDown />
            </span>
          )}
        </div>

        {selectedDashboardId && (
          <button
            type="button"
            title={`Open dashboard ${selectedDashboardId} in Looker`}
            aria-label="Open selected dashboard in Looker"
            onClick={() =>
              extensionSDK?.openBrowserWindow?.(
                `/dashboards/${encodeURIComponent(selectedDashboardId)}`,
                "_blank"
              )
            }
            style={{
              flexShrink: 0,
              width: 36,
              height: 36,
              borderRadius: 9,
              background: ui.inputBg,
              color: ui.textSecondary,
              border: `1px solid ${ui.border}`,
              display: "inline-flex",
              alignItems: "center",
              justifyContent: "center",
              cursor: "pointer",
            }}
          >
            <IconExternal />
          </button>
        )}
      </div>

      {/* Rich Filterable Popover */}
      {open && (
        <div
          role="listbox"
          data-testid="dashboard-picker-popover"
          style={{
            position: "absolute",
            top: "calc(100% + 6px)",
            left: 0,
            width: "min(560px, 92vw)",
            zIndex: 60,
            background: ui.cardBg,
            border: `1px solid ${ui.border}`,
            borderRadius: 12,
            boxShadow: isDark
              ? "0 22px 48px rgba(0, 0, 0, 0.65)"
              : "0 18px 40px rgba(15, 23, 42, 0.18)",
            overflow: "hidden",
          }}
        >
          {/* Scope Filter Pills */}
          <div
            style={{
              display: "flex",
              gap: 5,
              flexWrap: "wrap",
              padding: "9px 10px",
              borderBottom: `1px solid ${ui.border}`,
              background: ui.inputBg,
            }}
          >
            {tabs.map((t) => {
              const active = effectiveScope === t.id;
              return (
                <button
                  key={t.id}
                  type="button"
                  data-testid={`dash-scope-${t.id}`}
                  onClick={() => setScope(t.id)}
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 5,
                    padding: "4px 9px",
                    borderRadius: 7,
                    fontSize: 11.5,
                    fontWeight: 600,
                    cursor: "pointer",
                    border: `1px solid ${active ? ui.accent : ui.border}`,
                    background: active ? ui.accent : ui.cardBg,
                    color: active ? ui.accentContrast : ui.textSecondary,
                    whiteSpace: "nowrap",
                  }}
                >
                  {t.icon}
                  <span>{t.label}</span>
                  <span style={{ opacity: active ? 0.9 : 0.65, fontSize: 10.5 }}>{t.count}</span>
                </button>
              );
            })}
          </div>

          {/* Subtle Loading / Remote Search Bar */}
          {(loading || searchingRemote) && (
            <div
              style={{
                padding: "4px 12px",
                fontSize: 11,
                color: ui.accent,
                borderBottom: `1px solid ${ui.border}`,
                background: ui.inputBg,
              }}
            >
              {loading ? "Loading Looker dashboards..." : "Searching Looker instance..."}
            </div>
          )}

          {/* Results List */}
          <div style={{ maxHeight: 330, overflowY: "auto", padding: 6 }}>
            {pastedDirectId && (
              <button
                type="button"
                onClick={() => chooseDashboard(pastedDirectId)}
                style={{
                  width: "100%",
                  textAlign: "left",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  gap: 10,
                  padding: "9px 10px",
                  borderRadius: 8,
                  border: `1px dashed ${ui.accent}`,
                  background: ui.inputBg,
                  color: ui.textPrimary,
                  cursor: "pointer",
                  marginBottom: 4,
                }}
              >
                <span style={{ fontSize: 12.5, fontWeight: 600 }}>
                  Load dashboard ID <code>{pastedDirectId}</code> directly
                </span>
                <span style={{ fontSize: 11, color: ui.accent, fontWeight: 700 }}>Enter ↵</span>
              </button>
            )}

            {filteredItems.length === 0 && !pastedDirectId ? (
              <div style={{ padding: "22px 14px", textAlign: "center", fontSize: 12.5, color: ui.textMuted }}>
                <div>
                  No dashboards in <strong>{effectiveScope}</strong>
                  {query ? ` matching "${query}"` : ""}.
                </div>
                {effectiveScope !== "all" && (
                  <button
                    type="button"
                    onClick={() => setScope("all")}
                    style={{
                      marginTop: 8,
                      padding: "5px 11px",
                      borderRadius: 7,
                      border: `1px solid ${ui.accent}`,
                      background: "transparent",
                      color: ui.accent,
                      fontSize: 12,
                      fontWeight: 600,
                      cursor: "pointer",
                    }}
                  >
                    Search across all {counts.all} dashboards
                  </button>
                )}
              </div>
            ) : (
              filteredItems.map((d, idx) => {
                const isSelected = d.id === selectedDashboardId;
                const isHighlighted = idx === activeIdx;
                const groupLabel =
                  effectiveScope === "recent" && !query.trim()
                    ? d.isLookml
                      ? "LookML Dashboards"
                      : "Recently Viewed"
                    : d.isLookml
                      ? `LookML Model · ${d.modelLabel || d.id.split("::")[0]}`
                      : d.folderName || "Dashboards";
                const prev = idx > 0 ? filteredItems[idx - 1] : null;
                const prevGroup = prev
                  ? effectiveScope === "recent" && !query.trim()
                    ? prev.isLookml
                      ? "LookML Dashboards"
                      : "Recently Viewed"
                    : prev.isLookml
                      ? `LookML Model · ${prev.modelLabel || prev.id.split("::")[0]}`
                      : prev.folderName || "Dashboards"
                  : null;
                const showGroupHeader =
                  ((effectiveScope === "recent" && !query.trim()) ||
                    effectiveScope === "shared" ||
                    effectiveScope === "lookml") &&
                  groupLabel !== prevGroup;

                const subMeta = d.isLookml
                  ? d.id
                  : [
                      d.lastViewedAt ? `Viewed ${formatShortDate(d.lastViewedAt)}` : "",
                      d.viewCount ? `${d.viewCount} views` : "",
                      d.description || "",
                    ]
                      .filter(Boolean)
                      .join(" · ");

                const rightBadge = d.isLookml
                  ? d.modelLabel || d.id.split("::")[0]
                  : `${d.folderName || "Folder"} · #${d.id}`;

                return (
                  <React.Fragment key={d.id}>
                    {showGroupHeader && (
                      <div
                        style={{
                          padding: "7px 10px 3px",
                          fontSize: 10.5,
                          fontWeight: 700,
                          letterSpacing: "0.05em",
                          textTransform: "uppercase",
                          color: ui.textMuted,
                        }}
                      >
                        {groupLabel}
                      </div>
                    )}
                    <button
                      type="button"
                      role="option"
                      aria-selected={isSelected}
                      data-testid={`dash-option-${d.id}`}
                      onMouseEnter={() => setActiveIdx(idx)}
                      onClick={() => chooseDashboard(d.id, d)}
                      style={{
                        width: "100%",
                        textAlign: "left",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "space-between",
                        gap: 10,
                        padding: "8px 10px",
                        borderRadius: 8,
                        border: "none",
                        cursor: "pointer",
                        background: isSelected
                          ? `${ui.accent}22`
                          : isHighlighted
                            ? ui.inputBg
                            : "transparent",
                        color: ui.textPrimary,
                      }}
                    >
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 9,
                          minWidth: 0,
                          flex: 1,
                        }}
                      >
                        <span
                          style={{
                            flexShrink: 0,
                            color: d.isLookml ? ui.accent : ui.textMuted,
                            display: "inline-flex",
                          }}
                        >
                          {d.isLookml ? <IconCode /> : <IconFolder />}
                        </span>
                        <div style={{ minWidth: 0, flex: 1 }}>
                          <div
                            style={{
                              fontSize: 13,
                              fontWeight: isSelected ? 700 : 600,
                              color: ui.textPrimary,
                              overflow: "hidden",
                              textOverflow: "ellipsis",
                              whiteSpace: "nowrap",
                            }}
                          >
                            {d.title}
                          </div>
                          {subMeta && (
                            <div
                              style={{
                                fontSize: 11,
                                color: ui.textMuted,
                                overflow: "hidden",
                                textOverflow: "ellipsis",
                                whiteSpace: "nowrap",
                                marginTop: 1,
                              }}
                            >
                              {subMeta}
                            </div>
                          )}
                        </div>
                      </div>

                      <div
                        style={{
                          display: "inline-flex",
                          alignItems: "center",
                          gap: 6,
                          flexShrink: 0,
                        }}
                      >
                        <span
                          style={{
                            fontSize: 11,
                            fontWeight: 600,
                            padding: "2px 7px",
                            borderRadius: 999,
                            background: d.isLookml
                              ? `${ui.accent}1f`
                              : isDark
                                ? "rgba(148, 163, 184, 0.12)"
                                : "rgba(15, 23, 42, 0.06)",
                            color: d.isLookml ? ui.accent : ui.textMuted,
                            maxWidth: 165,
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                          }}
                        >
                          {rightBadge}
                        </span>
                        {isSelected && (
                          <span style={{ color: ui.accent, display: "inline-flex" }}>
                            <IconCheckSmall />
                          </span>
                        )}
                      </div>
                    </button>
                  </React.Fragment>
                );
              })
            )}
          </div>
        </div>
      )}
    </div>
  );
};

export const TEMPLATE_ORIGINAL_THEME = "__template_original__";

export interface LookerThemePalettePayload {
  bg_color: string;
  card_bg: string;
  title_color: string;
  text_color: string;
  accent: string;
  series_colors: string[];
}

export interface LookerThemeOption {
  name: string;
  label: string;
  isDark: boolean;
  palette?: LookerThemePalettePayload;
}

export function expandHex6(raw: unknown, fallback = "#0b57d0"): string {
  const cleaned = String(raw || "")
    .trim()
    .replace(/^#/, "");
  if (/^[0-9a-fA-F]{3,4}$/.test(cleaned)) {
    return (
      "#" +
      cleaned
        .slice(0, 3)
        .split("")
        .map((c) => c + c)
        .join("")
        .toLowerCase()
    );
  }
  if (/^[0-9a-fA-F]{6,8}$/.test(cleaned)) {
    return "#" + cleaned.slice(0, 6).toLowerCase();
  }
  return fallback;
}

export function hexLuminance(hex: string): number {
  const h = expandHex6(hex, "#090d16").slice(1);
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return 0.299 * r + 0.587 * g + 0.114 * b;
}

export function isDarkHex(hex: string): boolean {
  return hexLuminance(hex) < 128;
}

function prettyThemeLabel(name: string): string {
  if (name === TEMPLATE_ORIGINAL_THEME) return "Template Original";
  const m = name.match(/^(.*)_(Dark|Light)$/i);
  if (m) {
    const rawBase = m[1].replace(/_/g, " ").trim();
    const base = rawBase === "Levis" ? "Levi's" : rawBase;
    const mode = m[2][0].toUpperCase() + m[2].slice(1).toLowerCase();
    return `${base} · ${mode}`;
  }
  return name.replace(/_/g, " ");
}

const FALLBACK_BRAND_ROWS: Array<[string, string[], string[]]> = [
  [
    "Embed_Demo",
    ["#38bdf8", "#818cf8", "#34d399", "#fbbf24", "#fb7185"],
    ["#0b57d0", "#4f46e5", "#059669", "#d97706", "#e11d48"],
  ],
  [
    "Levis",
    ["#e11d48", "#3b82f6", "#2dd4bf", "#f59e0b", "#a855f7"],
    ["#c41230", "#0052cc", "#0d9488", "#d97706", "#7c3aed"],
  ],
  [
    "Calvin_Klein",
    ["#94a3b8", "#38bdf8", "#cbd5e1", "#64748b", "#a1a1aa"],
    ["#334155", "#0284c7", "#64748b", "#475569", "#0f766e"],
  ],
  [
    "Columbia",
    ["#00a896", "#38bdf8", "#f4a261", "#34d399", "#e76f51"],
    ["#0066b2", "#00a896", "#d97706", "#059669", "#dc2626"],
  ],
  [
    "Allegra_K",
    ["#f43f5e", "#c084fc", "#ec4899", "#fbbf24", "#2dd4bf"],
    ["#be123c", "#9333ea", "#db2777", "#d97706", "#0d9488"],
  ],
];

export const FALLBACK_LOOKER_THEMES: LookerThemeOption[] = [
  { name: TEMPLATE_ORIGINAL_THEME, label: "Template Original", isDark: true },
  ...FALLBACK_BRAND_ROWS.flatMap(([key, darkSeries, lightSeries]): LookerThemeOption[] => [
    {
      name: `${key}_Dark`,
      label: prettyThemeLabel(`${key}_Dark`),
      isDark: true,
      palette: {
        bg_color: "#090d16",
        card_bg: "#111827",
        title_color: "#f8fafc",
        text_color: "#e2e8f0",
        accent: darkSeries[0],
        series_colors: darkSeries,
      },
    },
    {
      name: `${key}_Light`,
      label: prettyThemeLabel(`${key}_Light`),
      isDark: false,
      palette: {
        bg_color: "#f0f4f9",
        card_bg: "#ffffff",
        title_color: "#0f172a",
        text_color: "#334155",
        accent: lightSeries[0],
        series_colors: lightSeries,
      },
    },
  ]),
];

export async function loadLookerThemeOptions(core40SDK: any): Promise<LookerThemeOption[]> {
  if (typeof core40SDK?.all_themes !== "function") return FALLBACK_LOOKER_THEMES;
  try {
    const [themesRaw, collectionsRaw] = await Promise.all([
      core40SDK.ok(core40SDK.all_themes()),
      typeof core40SDK.all_color_collections === "function"
        ? core40SDK.ok(core40SDK.all_color_collections()).catch(() => [])
        : Promise.resolve([]),
    ]);
    if (!Array.isArray(themesRaw) || themesRaw.length === 0) return FALLBACK_LOOKER_THEMES;

    const collectionMap = new Map<string, string[]>();
    for (const col of Array.isArray(collectionsRaw) ? collectionsRaw : []) {
      const cid = String(col?.id || "");
      const palettes = Array.isArray(col?.categoricalPalettes) ? col.categoricalPalettes : [];
      const rawStops = Array.isArray(palettes[0]?.stops) ? palettes[0].stops : [];
      const colors = rawStops
        .map((s: any) => expandHex6(s?.color ?? s, ""))
        .filter((c: string) => c.startsWith("#"));
      if (cid && colors.length > 0) collectionMap.set(cid, colors);
    }

    const built: LookerThemeOption[] = [];
    for (const t of themesRaw) {
      const name = String(t?.name || "").trim();
      if (!name) continue;
      const s = t?.settings || {};
      const bg = expandHex6(s.background_color, "#f6f8fa");
      const dark = isDarkHex(bg);
      const cardBg = expandHex6(s.tile_background_color, dark ? "#111827" : "#ffffff");
      const titleColor = expandHex6(
        s.tile_title_color || s.font_color,
        dark ? "#f8fafc" : "#0f172a"
      );
      const textColor = expandHex6(
        s.tile_text_color || s.font_color,
        dark ? "#e2e8f0" : "#334155"
      );
      const colColors = s.color_collection_id
        ? collectionMap.get(String(s.color_collection_id)) || []
        : [];
      const accent = expandHex6(
        s.primary_button_color || colColors[0],
        dark ? "#38bdf8" : "#0b57d0"
      );
      const series = (colColors.length > 0 ? colColors : [accent]).slice(0, 6);

      built.push({
        name,
        label: prettyThemeLabel(name),
        isDark: dark,
        palette: {
          bg_color: bg,
          card_bg: cardBg,
          title_color: titleColor,
          text_color: textColor,
          accent,
          series_colors: series,
        },
      });
    }

    built.sort((a, b) => a.label.localeCompare(b.label));
    return [{ name: TEMPLATE_ORIGINAL_THEME, label: "Template Original", isDark: true }, ...built];
  } catch {
    return FALLBACK_LOOKER_THEMES;
  }
}

interface LookerThemePickerProps {
  themes: LookerThemeOption[];
  selectedThemeName: string;
  onSelect: (option: LookerThemeOption) => void;
  ui: PickerPalette;
  isDark: boolean;
}

export const LookerThemePicker: React.FC<LookerThemePickerProps> = ({
  themes,
  selectedThemeName,
  onSelect,
  ui,
  isDark,
}) => {
  const [open, setOpen] = useState(false);
  const [filter, setFilter] = useState("");
  const wrapRef = useRef<HTMLDivElement | null>(null);

  useDismissOnOutsideOrEscape(open, wrapRef, () => {
    setOpen(false);
    setFilter("");
  });

  const current =
    themes.find((t) => t.name === selectedThemeName) ||
    themes[1] ||
    FALLBACK_LOOKER_THEMES[1];

  const visible = useMemo(() => {
    const q = filter.trim().toLowerCase();
    if (!q) return themes;
    return themes.filter(
      (t) => t.label.toLowerCase().includes(q) || t.name.toLowerCase().includes(q)
    );
  }, [themes, filter]);

  const renderSwatches = (opt: LookerThemeOption) => {
    if (!opt.palette) return null;
    const dots = [
      opt.palette.bg_color,
      opt.palette.accent,
      opt.palette.series_colors[1] || opt.palette.series_colors[0] || opt.palette.accent,
    ];
    return (
      <span style={{ display: "inline-flex", alignItems: "center", gap: 3, flexShrink: 0 }}>
        {dots.map((clr, idx) => (
          <span
            key={idx}
            style={{
              width: 10,
              height: 10,
              borderRadius: "50%",
              background: clr,
              border: "1px solid rgba(148, 163, 184, 0.45)",
              display: "inline-block",
            }}
          />
        ))}
      </span>
    );
  };

  return (
    <div ref={wrapRef} style={{ position: "relative" }}>
      <button
        id="brand-theme-select"
        type="button"
        aria-label="Looker Slide Theme"
        onClick={() => setOpen((v) => !v)}
        style={{
          display: "inline-flex",
          alignItems: "center",
          gap: 8,
          background: ui.cardBg,
          color: ui.textPrimary,
          border: `1px solid ${open ? ui.accent : ui.border}`,
          borderRadius: 8,
          padding: "7px 11px",
          fontSize: 12,
          fontWeight: 600,
          cursor: "pointer",
          whiteSpace: "nowrap",
        }}
      >
        {renderSwatches(current)}
        <span>{current.label}</span>
        <span style={{ color: ui.textMuted, display: "inline-flex" }}>
          <IconChevronDown />
        </span>
      </button>

      {open && (
        <div
          role="listbox"
          data-testid="looker-theme-popover"
          style={{
            position: "absolute",
            top: "calc(100% + 6px)",
            right: 0,
            width: 280,
            zIndex: 70,
            background: ui.cardBg,
            border: `1px solid ${ui.border}`,
            borderRadius: 11,
            boxShadow: isDark
              ? "0 20px 44px rgba(0, 0, 0, 0.6)"
              : "0 16px 36px rgba(15, 23, 42, 0.18)",
            overflow: "hidden",
          }}
        >
          <div
            style={{
              padding: "8px 10px",
              borderBottom: `1px solid ${ui.border}`,
              background: ui.inputBg,
              display: "flex",
              alignItems: "center",
              gap: 6,
            }}
          >
            <span style={{ color: ui.textMuted, display: "inline-flex" }}>
              <IconSearch />
            </span>
            <input
              type="text"
              value={filter}
              onChange={(e) => setFilter(e.target.value)}
              placeholder={`Filter ${themes.length - 1} Looker themes...`}
              style={{
                flex: 1,
                background: "transparent",
                border: "none",
                outline: "none",
                color: ui.textPrimary,
                fontSize: 12,
                fontFamily: "inherit",
              }}
            />
          </div>
          <div style={{ maxHeight: 280, overflowY: "auto", padding: 5 }}>
            {visible.map((opt) => {
              const active = opt.name === selectedThemeName;
              return (
                <button
                  key={opt.name}
                  type="button"
                  role="option"
                  aria-selected={active}
                  onClick={() => {
                    onSelect(opt);
                    setOpen(false);
                    setFilter("");
                  }}
                  style={{
                    width: "100%",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "space-between",
                    gap: 8,
                    padding: "7px 9px",
                    borderRadius: 7,
                    border: "none",
                    background: active ? `${ui.accent}22` : "transparent",
                    color: ui.textPrimary,
                    fontSize: 12.5,
                    fontWeight: active ? 700 : 500,
                    cursor: "pointer",
                    textAlign: "left",
                  }}
                >
                  <span style={{ display: "inline-flex", alignItems: "center", gap: 8, minWidth: 0 }}>
                    {renderSwatches(opt)}
                    <span
                      style={{
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        whiteSpace: "nowrap",
                      }}
                    >
                      {opt.label}
                    </span>
                  </span>
                  {active && (
                    <span style={{ color: ui.accent, display: "inline-flex", flexShrink: 0 }}>
                      <IconCheckSmall />
                    </span>
                  )}
                </button>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
};
