import React, { useEffect, useMemo, useRef, useState } from "react";
import { Svg } from "./ConfigPage";
import {
  DashboardPicker,
  DashboardSummary,
  LookerThemeOption,
  LookerThemePicker,
  PickerPalette,
  TEMPLATE_ORIGINAL_THEME,
} from "./DashboardPicker";
import {
  buildFanoutArtifactKey,
  buildFanoutSearchPattern,
  buildScheduledPlanWriteBody,
  buildSliceOutputPrefix,
  CadenceMode,
  CadenceSpec,
  describeCadence,
  FanoutSlice,
  isSlicePlanDrifted,
  MasterFanoutSpec,
  sanitizeOutputFolder,
  serializeFiltersString,
  slugifySegment,
} from "./schedulerUtils";

export interface SchedulerSeed {
  dashboardId?: string;
  dashboardTitle?: string;
  templateName?: string;
  themeName?: string;
}

interface SchedulerViewProps {
  core40SDK: any;
  extensionSDK: any;
  ui: PickerPalette;
  isDark: boolean;
  lookerThemes: LookerThemeOption[];
  actionIntegrationId?: string;
  targetUserIdOverride?: string;
  seed?: SchedulerSeed | null;
  onClearSeed?: () => void;
  callCloudRun: (path: string, body: Record<string, any>) => Promise<any>;
  resolveTemplatesGcsUri: () => Promise<string>;
  canOpenViewer: boolean;
  onOpenViewerSearch: (query: string) => void;
}

interface StoredFanoutRecord {
  artifactKey: string;
  version?: number;
  spec: MasterFanoutSpec;
}

interface DashboardFilterMeta {
  name: string;
  title: string;
  model: string;
  explore: string;
  dimension: string;
  defaultValue: string;
}

interface ExploreSortOption {
  name: string;
  label: string;
  kind: "dimension" | "measure";
}

interface LoadedFilterValue {
  value: string;
  metric?: string;
}

const COMMON_TIMEZONES = [
  "America/Los_Angeles",
  "America/Denver",
  "America/Chicago",
  "America/New_York",
  "UTC",
  "Europe/London",
  "Europe/Berlin",
  "Asia/Tokyo",
  "Asia/Singapore",
  "Australia/Sydney",
];

const WEEKDAYS: Array<{ value: number; label: string }> = [
  { value: 1, label: "Monday" },
  { value: 2, label: "Tuesday" },
  { value: 3, label: "Wednesday" },
  { value: 4, label: "Thursday" },
  { value: 5, label: "Friday" },
  { value: 6, label: "Saturday" },
  { value: 0, label: "Sunday" },
];

function defaultTimezone(): string {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || "America/Los_Angeles";
  } catch {
    return "America/Los_Angeles";
  }
}

const makeShortId = (): string => crypto.randomUUID().slice(0, 8);

function createBlankSpec(ownerUserId: string, seed?: SchedulerSeed | null): MasterFanoutSpec {
  const cadence: CadenceSpec = {
    mode: "weekly",
    hour: 8,
    minute: 0,
    dayOfWeek: 1,
    dayOfMonth: 1,
    crontab: "0 8 * * 1",
    timezone: defaultTimezone(),
  };
  const dashTitle = seed?.dashboardTitle || "";
  return {
    id: makeShortId(),
    title: dashTitle ? `${dashTitle} Fanout` : "",
    owner_user_id: ownerUserId,
    dashboard_id: seed?.dashboardId || "",
    dashboard_title: dashTitle,
    template: seed?.templateName || "",
    theme: seed?.themeName || TEMPLATE_ORIGINAL_THEME,
    output_folder: dashTitle ? slugifySegment(dashTitle, "fanout") : "fanout",
    cadence,
    enabled: true,
    base_filters: {},
    slices: [],
    updated_at: new Date().toISOString(),
  };
}

export const SchedulerView: React.FC<SchedulerViewProps> = ({
  core40SDK,
  extensionSDK,
  ui,
  isDark,
  lookerThemes,
  actionIntegrationId,
  targetUserIdOverride,
  seed,
  onClearSeed,
  callCloudRun,
  resolveTemplatesGcsUri,
  canOpenViewer,
  onOpenViewerSearch,
}) => {
  const artifactNamespace = useMemo(
    () =>
      String(
        extensionSDK?.lookerHostData?.extensionId || "action-pptx::action_pptx_template_manager"
      ),
    [extensionSDK]
  );

  const [meId, setMeId] = useState<string>("");
  const [destinationType, setDestinationType] = useState<string>(
    actionIntegrationId ? `looker-integration://${actionIntegrationId}` : ""
  );
  const [filterDashId, setFilterDashId] = useState<string>("");

  const [records, setRecords] = useState<StoredFanoutRecord[]>([]);
  const [livePlansById, setLivePlansById] = useState<Record<string, any>>({});
  const [loading, setLoading] = useState<boolean>(true);
  const [busyMessage, setBusyMessage] = useState<string>("");
  const [notice, setNotice] = useState<string>("");
  const [error, setError] = useState<string>("");
  const [confirmDeleteId, setConfirmDeleteId] = useState<string>("");

  // Editor state
  const [editingPrevKey, setEditingPrevKey] = useState<string | null>(null);
  const [draft, setDraft] = useState<MasterFanoutSpec | null>(null);
  const [templatesList, setTemplatesList] = useState<string[]>([]);
  const [loadingTemplates, setLoadingTemplates] = useState<boolean>(false);
  const [dashFilters, setDashFilters] = useState<DashboardFilterMeta[]>([]);
  const [showBaseFilters, setShowBaseFilters] = useState<boolean>(false);

  // Fanout value loader + slice table filter/sort state
  const [activeFilterName, setActiveFilterName] = useState<string>("");
  const [exploreSortOptions, setExploreSortOptions] = useState<ExploreSortOption[]>([]);
  const [queryFilterExpr, setQueryFilterExpr] = useState<string>("");
  const [querySortField, setQuerySortField] = useState<string>("");
  const [querySortDir, setQuerySortDir] = useState<"asc" | "desc">("asc");
  const [loadingValues, setLoadingValues] = useState<boolean>(false);
  const [loadedValues, setLoadedValues] = useState<LoadedFilterValue[]>([]);
  const [selectedValues, setSelectedValues] = useState<Set<string>>(new Set());
  const [valueSearch, setValueSearch] = useState<string>("");
  const [valueClientSort, setValueClientSort] = useState<"query" | "asc" | "desc">("query");
  const [pasteOpen, setPasteOpen] = useState<boolean>(false);
  const [pasteText, setPasteText] = useState<string>("");

  const [sliceSearch, setSliceSearch] = useState<string>("");
  const [sliceSortKey, setSliceSortKey] = useState<"label" | "slug" | "value" | "status">("label");
  const [sliceSortAsc, setSliceSortAsc] = useState<boolean>(true);

  const exploreCacheRef = useRef<Map<string, ExploreSortOption[]>>(new Map());
  const effectiveUserId = (targetUserIdOverride || meId || "").trim();
  const isAdminInspectingOther = Boolean(
    targetUserIdOverride && meId && targetUserIdOverride.trim() !== meId
  );

  async function ensureIdentityAndDestination(): Promise<{ uid: string; dest: string }> {
    let uid = effectiveUserId;
    if (!uid) {
      const me = await core40SDK.ok(core40SDK.me());
      uid = String(me?.id || "");
      setMeId(uid);
      if (targetUserIdOverride?.trim()) uid = targetUserIdOverride.trim();
    }
    let dest = destinationType;
    if (!dest) {
      if (actionIntegrationId) {
        dest = `looker-integration://${actionIntegrationId}`;
      } else {
        try {
          const integrations: any[] = await core40SDK.ok(core40SDK.all_integrations({}));
          const candidates = (Array.isArray(integrations) ? integrations : []).filter(
            (i) => i?.id?.includes("action-pptx") || i?.label?.toLowerCase().includes("action-pptx")
          );
          const match = candidates.find((i) => i.enabled) || candidates[0];
          if (match?.id) dest = `looker-integration://${match.id}`;
        } catch {
          // Non-admin users without manage_integrations fall back to inspecting existing plans
        }
      }
      if (dest) setDestinationType(dest);
    }
    return { uid, dest };
  }

  async function refreshSchedules() {
    setLoading(true);
    setError("");
    try {
      const { uid, dest } = await ensureIdentityAndDestination();
      if (!uid) throw new Error("Unable to resolve current Looker user ID.");

      const keyPattern = buildFanoutSearchPattern(uid, filterDashId);
      const planQuery: Record<string, unknown> = {
        user_id: uid,
        all_states: true,
        ...(dest ? { destination_type: dest } : {}),
        ...(filterDashId
          ? filterDashId.includes("::")
            ? { lookml_dashboard_id: filterDashId }
            : { dashboard_id: filterDashId }
          : {}),
      };

      const [rawArtifacts, rawPlans]: [any[], any[]] = await Promise.all([
        core40SDK
          .ok(core40SDK.search_artifacts({ namespace: artifactNamespace, key: keyPattern }))
          .catch(() => []),
        core40SDK.ok(core40SDK.search_scheduled_plans(planQuery)).catch(() => []),
      ]);

      const parsedRecords: StoredFanoutRecord[] = [];
      for (const art of Array.isArray(rawArtifacts) ? rawArtifacts : []) {
        try {
          const spec = JSON.parse(String(art.value || "")) as MasterFanoutSpec;
          if (spec && spec.id && Array.isArray(spec.slices)) {
            parsedRecords.push({
              artifactKey: String(art.key),
              version: typeof art.version === "number" ? art.version : undefined,
              spec,
            });
          }
        } catch {
          // Ignore malformed artifact entries
        }
      }
      parsedRecords.sort((a, b) =>
        String(b.spec.updated_at || "").localeCompare(String(a.spec.updated_at || ""))
      );

      const planMap: Record<string, any> = {};
      let inferredDest = dest;
      for (const p of Array.isArray(rawPlans) ? rawPlans : []) {
        if (p?.id != null) planMap[String(p.id)] = p;
        if (!inferredDest && Array.isArray(p?.scheduled_plan_destination)) {
          const found = p.scheduled_plan_destination.find((d: any) =>
            String(d?.type || "").includes("action-pptx")
          );
          if (found?.type) inferredDest = String(found.type);
        }
      }
      if (!dest && inferredDest) setDestinationType(inferredDest);

      setRecords(parsedRecords);
      setLivePlansById(planMap);
    } catch (err: any) {
      setError(`Could not load schedules: ${err?.message || err}`);
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    void refreshSchedules();
  }, [targetUserIdOverride, filterDashId]);

  useEffect(() => {
    if (!seed) return;
    void ensureIdentityAndDestination().then(({ uid }) => {
      openEditor(createBlankSpec(uid || "me", seed), null);
      onClearSeed?.();
    });
  }, [seed]);

  async function loadTemplateOptions(preselect?: string) {
    setLoadingTemplates(true);
    try {
      const uri = await resolveTemplatesGcsUri();
      if (!uri) return;
      const res = await callCloudRun("/templates/list", { templates_gcs_uri: uri });
      const names = (Array.isArray(res?.files) ? res.files : [])
        .map((f: any) => String(typeof f === "string" ? f : f?.name || ""))
        .filter(Boolean)
        .sort();
      setTemplatesList(names);
      if (!preselect && names.length > 0) {
        setDraft((prev) => (prev && !prev.template ? { ...prev, template: names[0] } : prev));
      }
    } catch (err: any) {
      setError(`Could not list GCS templates: ${err?.message || err}`);
    } finally {
      setLoadingTemplates(false);
    }
  }

  async function loadDashboardFilterMeta(dashId: string, existingSlices?: FanoutSlice[]) {
    if (!dashId) {
      setDashFilters([]);
      setActiveFilterName("");
      return;
    }
    try {
      const dash = await core40SDK.ok(core40SDK.dashboard(dashId));
      const rawFilters: any[] = Array.isArray(dash?.dashboard_filters)
        ? dash.dashboard_filters
        : [];
      const mapped: DashboardFilterMeta[] = rawFilters
        .map((f) => ({
          name: String(f?.name || f?.title || "").trim(),
          title: String(f?.title || f?.name || "").trim(),
          model: String(f?.model || "").trim(),
          explore: String(f?.explore || "").trim(),
          dimension: String(f?.dimension || "").trim(),
          defaultValue: f?.default_value != null ? String(f.default_value) : "",
        }))
        .filter((f) => Boolean(f.name));
      setDashFilters(mapped);

      const usedKey = existingSlices?.flatMap((s) => Object.keys(s.filters || {}))[0];
      const nextFilter =
        (usedKey && mapped.find((m) => m.name === usedKey)?.name) || mapped[0]?.name || usedKey || "";
      setActiveFilterName(nextFilter);
      if (dash?.title) {
        setDraft((prev) =>
          prev
            ? {
                ...prev,
                dashboard_title: prev.dashboard_title || String(dash.title),
                title: prev.title || `${dash.title} Fanout`,
                output_folder:
                  prev.output_folder && prev.output_folder !== "fanout"
                    ? prev.output_folder
                    : slugifySegment(String(dash.title), "fanout"),
              }
            : prev
        );
      }
    } catch (err: any) {
      setError(`Could not inspect dashboard filters: ${err?.message || err}`);
    }
  }

  const selectedDashFilter = useMemo(
    () => dashFilters.find((f) => f.name === activeFilterName) || null,
    [dashFilters, activeFilterName]
  );

  useEffect(() => {
    setLoadedValues([]);
    setSelectedValues(new Set());
    setQueryFilterExpr("");
    if (!selectedDashFilter?.model || !selectedDashFilter?.explore) {
      setExploreSortOptions([]);
      setQuerySortField("");
      return;
    }
    const defaultDim = selectedDashFilter.dimension;
    setQuerySortField(defaultDim);
    setQuerySortDir("asc");
    const cacheKey = `${selectedDashFilter.model}::${selectedDashFilter.explore}`;
    const cached = exploreCacheRef.current.get(cacheKey);
    if (cached) {
      setExploreSortOptions(cached);
      return;
    }
    void core40SDK
      .ok(
        core40SDK.lookml_model_explore(
          selectedDashFilter.model,
          selectedDashFilter.explore,
          "fields(dimensions(name,label_short,label),measures(name,label_short,label))"
        )
      )
      .then((exp: any) => {
        const dims: ExploreSortOption[] = (exp?.fields?.dimensions || [])
          .map((d: any) => ({
            name: String(d?.name || ""),
            label: `${d?.label_short || d?.label || d?.name} (dim)`,
            kind: "dimension" as const,
          }))
          .filter((d: ExploreSortOption) => Boolean(d.name));
        const meas: ExploreSortOption[] = (exp?.fields?.measures || [])
          .map((m: any) => ({
            name: String(m?.name || ""),
            label: `${m?.label_short || m?.label || m?.name} (measure)`,
            kind: "measure" as const,
          }))
          .filter((m: ExploreSortOption) => Boolean(m.name));
        const opts = [...dims, ...meas];
        exploreCacheRef.current.set(cacheKey, opts);
        setExploreSortOptions(opts);
      })
      .catch(() => setExploreSortOptions([]));
  }, [selectedDashFilter?.name, selectedDashFilter?.model, selectedDashFilter?.explore]);

  function openEditor(spec: MasterFanoutSpec, prevArtifactKey: string | null) {
    setError("");
    setNotice("");
    setEditingPrevKey(prevArtifactKey);
    const cloned: MasterFanoutSpec = JSON.parse(JSON.stringify(spec));
    const themeOpt = lookerThemes.find((t) => t.name === cloned.theme);
    if (themeOpt?.palette && !cloned.theme_palette) {
      cloned.theme_palette = themeOpt.palette as unknown as Record<string, unknown>;
    }
    setDraft(cloned);
    setLoadedValues([]);
    setSelectedValues(new Set());
    setPasteOpen(false);
    setPasteText("");
    setSliceSearch("");
    void loadTemplateOptions(cloned.template);
    if (cloned.dashboard_id) {
      void loadDashboardFilterMeta(cloned.dashboard_id, cloned.slices);
    } else {
      setDashFilters([]);
    }
  }

  async function handleLoadDimensionValues() {
    if (!selectedDashFilter?.model || !selectedDashFilter?.explore || !selectedDashFilter?.dimension) {
      setError("Selected filter has no linked LookML model/explore/dimension; use Paste Values or + Add Row.");
      return;
    }
    setLoadingValues(true);
    setError("");
    try {
      const dim = selectedDashFilter.dimension;
      const sortBy = querySortField || dim;
      const fields = sortBy !== dim ? [dim, sortBy] : [dim];
      const filters: Record<string, string> = {};
      for (const f of dashFilters) {
        const baseVal = draft?.base_filters?.[f.name]?.trim();
        if (
          baseVal &&
          f.model === selectedDashFilter.model &&
          f.explore === selectedDashFilter.explore &&
          f.dimension
        ) {
          filters[f.dimension] = baseVal;
        }
      }
      if (queryFilterExpr.trim()) {
        filters[dim] = queryFilterExpr.trim();
      }
      const raw = await core40SDK.ok(
        core40SDK.run_inline_query({
          result_format: "json",
          body: {
            model: selectedDashFilter.model,
            view: selectedDashFilter.explore,
            fields,
            ...(Object.keys(filters).length > 0 ? { filters } : {}),
            sorts: [`${sortBy} ${querySortDir}`],
            // ponytail: Looker inline picker capped at 500 rows; use query filter or Paste Values above 500
            limit: "500",
          },
        })
      );
      const rows: any[] = typeof raw === "string" ? JSON.parse(raw) : Array.isArray(raw) ? raw : [];
      const items: LoadedFilterValue[] = [];
      const seen = new Set<string>();
      for (const r of rows) {
        const rawVal = r?.[dim];
        if (rawVal === null || rawVal === undefined) continue;
        const val = String(rawVal).trim();
        if (!val || seen.has(val)) continue;
        seen.add(val);
        items.push({
          value: val,
          metric: sortBy !== dim && r?.[sortBy] != null ? String(r[sortBy]) : undefined,
        });
      }
      setLoadedValues(items);
      setSelectedValues(new Set());
    } catch (err: any) {
      setError(`Failed to load filter values: ${err?.message || err}`);
    } finally {
      setLoadingValues(false);
    }
  }

  const visibleLoadedValues = useMemo(() => {
    const q = valueSearch.trim().toLowerCase();
    const filtered = loadedValues.filter(
      (item) => !q || item.value.toLowerCase().includes(q) || (item.metric || "").toLowerCase().includes(q)
    );
    if (valueClientSort === "query") return filtered;
    return filtered.slice().sort((a, b) => {
      const cmp = a.value.localeCompare(b.value, undefined, { numeric: true, sensitivity: "base" });
      return valueClientSort === "asc" ? cmp : -cmp;
    });
  }, [loadedValues, valueSearch, valueClientSort]);

  function appendValuesAsSlices(values: string[]) {
    if (!draft) return;
    const filterKey = activeFilterName.trim() || dashFilters[0]?.name || "Filter";
    const existingVals = new Set(
      draft.slices.map((s) => String(s.filters?.[filterKey] || "").trim().toLowerCase())
    );
    const additions: FanoutSlice[] = [];
    for (const raw of values) {
      const cleaned = raw.trim();
      if (!cleaned || existingVals.has(cleaned.toLowerCase())) continue;
      existingVals.add(cleaned.toLowerCase());
      additions.push({
        id: makeShortId(),
        label: cleaned,
        slug: slugifySegment(cleaned, "slice"),
        filters: { [filterKey]: cleaned },
      });
    }
    if (additions.length > 0) {
      setDraft({ ...draft, slices: [...draft.slices, ...additions] });
    }
  }

  async function saveFanoutArtifact(
    key: string,
    spec: MasterFanoutSpec,
    fallbackVer?: number
  ): Promise<void> {
    const existing: any[] = await core40SDK
      .ok(core40SDK.artifact({ namespace: artifactNamespace, key }))
      .catch(() => []);
    const ver =
      Array.isArray(existing) && typeof existing[0]?.version === "number"
        ? existing[0].version
        : fallbackVer;
    await core40SDK.ok(
      core40SDK.update_artifacts(artifactNamespace, [
        {
          key,
          value: JSON.stringify(spec),
          content_type: "application/json",
          ...(ver !== undefined ? { version: ver } : {}),
        },
      ])
    );
  }

  async function syncAndSaveSpec(
    nextSpec: MasterFanoutSpec,
    prevKey: string | null,
    prevSlices?: FanoutSlice[]
  ) {
    const { uid, dest } = await ensureIdentityAndDestination();
    if (!dest) {
      throw new Error(
        "Could not resolve 'looker-integration://<id>::action-pptx'. Open /config once as Admin to auto-discover."
      );
    }
    const ownerId = nextSpec.owner_user_id || uid;
    const normalized: MasterFanoutSpec = {
      ...nextSpec,
      owner_user_id: ownerId,
      output_folder: sanitizeOutputFolder(nextSpec.output_folder, "fanout"),
      updated_at: new Date().toISOString(),
    };

    const nextPlanIds = new Set(normalized.slices.map((s) => s.plan_id).filter(Boolean));
    await Promise.all(
      (prevSlices || [])
        .filter((oldSlice) => oldSlice.plan_id && !nextPlanIds.has(oldSlice.plan_id))
        .map((oldSlice) =>
          core40SDK.ok(core40SDK.delete_scheduled_plan(String(oldSlice.plan_id))).catch(() => {})
        )
    );

    const syncedSlices: FanoutSlice[] = await Promise.all(
      normalized.slices.map(async (slice) => {
        const body = buildScheduledPlanWriteBody(normalized, slice, dest);
        if (slice.plan_id && livePlansById[String(slice.plan_id)]) {
          const updated = await core40SDK.ok(
            core40SDK.update_scheduled_plan(String(slice.plan_id), body)
          );
          return { ...slice, plan_id: String(updated?.id || slice.plan_id) };
        }
        const created = await core40SDK.ok(core40SDK.create_scheduled_plan(body));
        return { ...slice, plan_id: String(created?.id || "") };
      })
    );

    const finalSpec: MasterFanoutSpec = { ...normalized, slices: syncedSlices };
    const nextKey = buildFanoutArtifactKey(ownerId, finalSpec.dashboard_id, finalSpec.id);
    await saveFanoutArtifact(nextKey, finalSpec);

    if (prevKey && prevKey !== nextKey) {
      await core40SDK.ok(core40SDK.delete_artifact(artifactNamespace, prevKey)).catch(() => {});
    }
  }

  async function handleEditorSave(e: React.FormEvent) {
    e.preventDefault();
    if (!draft) return;
    if (!draft.title.trim()) {
      setError("Schedule title is required.");
      return;
    }
    if (!draft.dashboard_id.trim()) {
      setError("Choose a Looker dashboard.");
      return;
    }
    if (!draft.template.trim()) {
      setError("Select or enter a .pptx template filename.");
      return;
    }
    if (draft.cadence.mode === "datagroup" && !String(draft.cadence.datagroup || "").trim()) {
      setError("Enter a LookML datagroup name for datagroup-triggered schedules.");
      return;
    }
    if (draft.slices.length === 0) {
      setError("Add at least one fanout slice row.");
      return;
    }

    setBusyMessage(`Syncing ${draft.slices.length} Looker scheduled plan(s)...`);
    setError("");
    try {
      const prevRec = editingPrevKey
        ? records.find((r) => r.artifactKey === editingPrevKey)
        : undefined;
      await syncAndSaveSpec(draft, editingPrevKey, prevRec?.spec.slices);
      setDraft(null);
      setEditingPrevKey(null);
      setNotice(`Saved "${draft.title}" (${draft.slices.length} slice plans synced).`);
      await refreshSchedules();
    } catch (err: any) {
      setError(`Save & Sync failed: ${err?.message || err}`);
    } finally {
      setBusyMessage("");
    }
  }

  async function handleCardSync(rec: StoredFanoutRecord) {
    setBusyMessage(`Reconciling "${rec.spec.title}"...`);
    setError("");
    try {
      await syncAndSaveSpec(rec.spec, rec.artifactKey, rec.spec.slices);
      setNotice(`Synced all ${rec.spec.slices.length} slice plans for "${rec.spec.title}".`);
      await refreshSchedules();
    } catch (err: any) {
      setError(`Sync failed: ${err?.message || err}`);
    } finally {
      setBusyMessage("");
    }
  }

  async function handleRunAllNow(rec: StoredFanoutRecord) {
    const liveIds = rec.spec.slices
      .map((s) => s.plan_id)
      .filter((pid): pid is string => Boolean(pid && livePlansById[String(pid)]));
    if (liveIds.length === 0) {
      setError("No synced Looker plans found on this Master Schedule — click Sync Plans first.");
      return;
    }
    setBusyMessage(`Queueing ${liveIds.length} run(s) in Looker...`);
    setError("");
    try {
      await Promise.all(
        liveIds.map((pid) => core40SDK.ok(core40SDK.scheduled_plan_run_once_by_id(String(pid))))
      );
      setNotice(`Queued ${liveIds.length} immediate run(s) for "${rec.spec.title}".`);
    } catch (err: any) {
      setError(`Run All Now failed: ${err?.message || err}`);
    } finally {
      setBusyMessage("");
    }
  }

  async function handleToggleEnabled(rec: StoredFanoutRecord) {
    const nextEnabled = !rec.spec.enabled;
    setBusyMessage(`${nextEnabled ? "Resuming" : "Pausing"} "${rec.spec.title}"...`);
    setError("");
    try {
      const liveIds = rec.spec.slices
        .map((s) => s.plan_id)
        .filter((pid): pid is string => Boolean(pid && livePlansById[String(pid)]));
      await Promise.all(
        liveIds.map((pid) =>
          core40SDK.ok(core40SDK.update_scheduled_plan(String(pid), { enabled: nextEnabled }))
        )
      );
      const updatedSpec: MasterFanoutSpec = {
        ...rec.spec,
        enabled: nextEnabled,
        updated_at: new Date().toISOString(),
      };
      await saveFanoutArtifact(rec.artifactKey, updatedSpec, rec.version);
      setNotice(`${nextEnabled ? "Resumed" : "Paused"} "${rec.spec.title}".`);
      await refreshSchedules();
    } catch (err: any) {
      setError(`Could not update schedule state: ${err?.message || err}`);
    } finally {
      setBusyMessage("");
    }
  }

  async function handleDeleteMaster(rec: StoredFanoutRecord) {
    setBusyMessage(`Deleting "${rec.spec.title}"...`);
    setError("");
    try {
      const planIds = rec.spec.slices.map((s) => s.plan_id).filter(Boolean);
      await Promise.all(
        planIds.map((pid) =>
          core40SDK.ok(core40SDK.delete_scheduled_plan(String(pid))).catch(() => {})
        )
      );
      await core40SDK.ok(core40SDK.delete_artifact(artifactNamespace, rec.artifactKey));
      setConfirmDeleteId("");
      setNotice(`Deleted "${rec.spec.title}" and ${planIds.length} Looker plan(s).`);
      await refreshSchedules();
    } catch (err: any) {
      setError(`Delete failed: ${err?.message || err}`);
    } finally {
      setBusyMessage("");
    }
  }

  const inputStyle: React.CSSProperties = {
    width: "100%",
    padding: "6px 9px",
    borderRadius: 6,
    background: ui.inputBg,
    color: ui.textPrimary,
    border: `1px solid ${ui.border}`,
    fontSize: 12.5,
    boxSizing: "border-box",
  };

  const btnSecondary: React.CSSProperties = {
    display: "inline-flex",
    alignItems: "center",
    gap: 5,
    background: ui.inputBg,
    color: ui.textPrimary,
    border: `1px solid ${ui.border}`,
    borderRadius: 6,
    padding: "5px 10px",
    fontSize: 12,
    fontWeight: 600,
    cursor: busyMessage ? "wait" : "pointer",
    whiteSpace: "nowrap",
  };

  const primarySliceFilterKey = useMemo(() => {
    if (activeFilterName) return activeFilterName;
    const fromSlice = draft?.slices.flatMap((s) => Object.keys(s.filters || {}))[0];
    return fromSlice || dashFilters[0]?.name || "Filter";
  }, [activeFilterName, draft?.slices, dashFilters]);

  const displayedDraftSlices = useMemo(() => {
    if (!draft) return [];
    const q = sliceSearch.trim().toLowerCase();
    const filtered = draft.slices.filter((s) => {
      if (!q) return true;
      const filterVals = Object.entries(s.filters || {})
        .map(([k, v]) => `${k}=${v}`)
        .join(" ");
      return (
        s.label.toLowerCase().includes(q) ||
        s.slug.toLowerCase().includes(q) ||
        filterVals.toLowerCase().includes(q)
      );
    });
    return filtered.slice().sort((a, b) => {
      let cmp = 0;
      if (sliceSortKey === "slug") {
        cmp = a.slug.localeCompare(b.slug, undefined, { numeric: true });
      } else if (sliceSortKey === "value") {
        cmp = String(a.filters?.[primarySliceFilterKey] || "").localeCompare(
          String(b.filters?.[primarySliceFilterKey] || ""),
          undefined,
          { numeric: true }
        );
      } else if (sliceSortKey === "status") {
        const sa = a.plan_id && livePlansById[a.plan_id] ? "1" : "0";
        const sb = b.plan_id && livePlansById[b.plan_id] ? "1" : "0";
        cmp = sa.localeCompare(sb);
      } else {
        cmp = a.label.localeCompare(b.label, undefined, { numeric: true });
      }
      return sliceSortAsc ? cmp : -cmp;
    });
  }, [draft, sliceSearch, sliceSortKey, sliceSortAsc, primarySliceFilterKey, livePlansById]);

  function toggleSliceSort(key: "label" | "slug" | "value" | "status") {
    if (sliceSortKey === key) setSliceSortAsc((v) => !v);
    else {
      setSliceSortKey(key);
      setSliceSortAsc(true);
    }
  }

  return (
    <div style={{ display: "grid", gap: 12 }}>
      {isAdminInspectingOther && (
        <div
          data-testid="scheduler-impersonation-banner"
          style={{
            background: isDark ? "#1e293b" : "#eff6ff",
            border: `1px solid ${ui.accent}`,
            color: ui.textPrimary,
            borderRadius: 8,
            padding: "8px 12px",
            fontSize: 12,
            fontWeight: 600,
          }}
        >
          Admin Inspection Mode: viewing &amp; managing Master Fanout schedules for Looker{" "}
          <code>user_id={targetUserIdOverride}</code>.
        </div>
      )}

      {error && (
        <div
          data-testid="scheduler-error-banner"
          style={{
            background: isDark ? "#450a0a" : "#fef2f2",
            border: "1px solid #dc2626",
            color: isDark ? "#fecaca" : "#991b1b",
            borderRadius: 8,
            padding: "8px 12px",
            fontSize: 12.5,
          }}
        >
          {error}
        </div>
      )}

      {(notice || busyMessage) && (
        <div
          style={{
            background: isDark ? "#064e3b" : "#ecfdf5",
            border: `1px solid ${isDark ? "#059669" : "#10b981"}`,
            color: isDark ? "#6ee7b7" : "#065f46",
            borderRadius: 8,
            padding: "8px 12px",
            fontSize: 12.5,
            fontWeight: 600,
          }}
        >
          {busyMessage || notice}
        </div>
      )}

      {draft ? (
        <form
          onSubmit={(e) => void handleEditorSave(e)}
          data-testid="scheduler-editor-form"
          style={{
            background: ui.cardBg,
            border: `1px solid ${ui.border}`,
            borderRadius: 10,
            padding: 14,
            display: "grid",
            gap: 12,
          }}
        >
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              gap: 8,
              borderBottom: `1px solid ${ui.border}`,
              paddingBottom: 10,
            }}
          >
            <div>
              <div style={{ fontSize: 15, fontWeight: 700 }}>
                {editingPrevKey ? "Edit Master Fanout Schedule" : "New Master Fanout Schedule"}
              </div>
              <div style={{ fontSize: 11.5, color: ui.textMuted }}>
                1 dashboard + 1 template + 1 cadence → N Looker ScheduledPlans writing to{" "}
                <code>{sanitizeOutputFolder(draft.output_folder, "fanout")}/&lt;slice&gt;_&lt;ts&gt;.pptx</code>
              </div>
            </div>
            <div style={{ display: "flex", gap: 6 }}>
              <button
                type="button"
                onClick={() => {
                  setDraft(null);
                  setEditingPrevKey(null);
                }}
                style={btnSecondary}
              >
                Cancel
              </button>
              <button
                id="btn-save-fanout"
                type="submit"
                disabled={Boolean(busyMessage)}
                style={{
                  ...btnSecondary,
                  background: "#059669",
                  color: "#ffffff",
                  border: "none",
                  fontWeight: 700,
                }}
              >
                Save &amp; Sync ({draft.slices.length})
              </button>
            </div>
          </div>

          {/* Core Master Metadata */}
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
              gap: 10,
            }}
          >
            <label>
              <div style={{ fontSize: 11.5, fontWeight: 700, marginBottom: 4 }}>Master Title</div>
              <input
                id="input-fanout-title"
                type="text"
                required
                value={draft.title}
                onChange={(e) => setDraft({ ...draft, title: e.target.value })}
                placeholder="Quarterly Brand EBR Fanout"
                style={inputStyle}
              />
            </label>

            <label>
              <div style={{ fontSize: 11.5, fontWeight: 700, marginBottom: 4 }}>
                Output Subfolder Prefix
              </div>
              <input
                id="input-fanout-folder"
                type="text"
                value={draft.output_folder}
                onChange={(e) => setDraft({ ...draft, output_folder: e.target.value })}
                placeholder="Q4_EBR"
                style={inputStyle}
              />
            </label>
          </div>

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "minmax(0, 1.2fr) minmax(0, 1fr) auto",
              gap: 10,
              alignItems: "end",
            }}
          >
            <div>
              <div style={{ fontSize: 11.5, fontWeight: 700, marginBottom: 4 }}>
                Source Dashboard
              </div>
              <DashboardPicker
                core40SDK={core40SDK}
                extensionSDK={extensionSDK}
                selectedDashboardId={draft.dashboard_id}
                onSelect={(id: string, summary?: DashboardSummary) => {
                  const nextTitle = summary?.title || id;
                  setDraft((prev) =>
                    prev
                      ? {
                          ...prev,
                          dashboard_id: id,
                          dashboard_title: nextTitle,
                          title: prev.title || (nextTitle ? `${nextTitle} Fanout` : ""),
                          output_folder:
                            prev.output_folder && prev.output_folder !== "fanout"
                              ? prev.output_folder
                              : slugifySegment(nextTitle, "fanout"),
                        }
                      : prev
                  );
                  void loadDashboardFilterMeta(id, draft.slices);
                }}
                ui={ui}
                isDark={isDark}
              />
            </div>

            <div>
              <div style={{ fontSize: 11.5, fontWeight: 700, marginBottom: 4 }}>
                PowerPoint Template (.pptx){" "}
                {loadingTemplates && (
                  <span style={{ fontWeight: 400, color: ui.textMuted }}>(loading...)</span>
                )}
              </div>
              {templatesList.length > 0 ? (
                <select
                  id="select-fanout-template"
                  value={draft.template}
                  onChange={(e) => setDraft({ ...draft, template: e.target.value })}
                  style={inputStyle}
                >
                  {!templatesList.includes(draft.template) && draft.template && (
                    <option value={draft.template}>{draft.template}</option>
                  )}
                  <option value="" disabled>
                    Select a saved template...
                  </option>
                  {templatesList.map((t) => (
                    <option key={t} value={t}>
                      {t}
                    </option>
                  ))}
                </select>
              ) : (
                <input
                  id="input-fanout-template"
                  type="text"
                  value={draft.template}
                  onChange={(e) => setDraft({ ...draft, template: e.target.value })}
                  placeholder="simple_template.pptx"
                  style={inputStyle}
                />
              )}
            </div>

            <div>
              <div style={{ fontSize: 11.5, fontWeight: 700, marginBottom: 4 }}>Deck Theme</div>
              <LookerThemePicker
                themes={lookerThemes}
                selectedThemeName={draft.theme || TEMPLATE_ORIGINAL_THEME}
                onSelect={(opt) =>
                  setDraft({
                    ...draft,
                    theme: opt.name,
                    theme_palette: opt.palette
                      ? (opt.palette as unknown as Record<string, unknown>)
                      : undefined,
                  })
                }
                ui={ui}
                isDark={isDark}
              />
            </div>
          </div>

          {/* Cadence Controls */}
          <div
            style={{
              background: ui.inputBg,
              border: `1px solid ${ui.border}`,
              borderRadius: 8,
              padding: 10,
              display: "grid",
              gap: 8,
            }}
          >
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span style={{ fontSize: 12, fontWeight: 700 }}>
                Schedule Cadence ·{" "}
                <span style={{ color: ui.accent }}>{describeCadence(draft.cadence)}</span>
              </span>
              <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12 }}>
                <input
                  type="checkbox"
                  checked={draft.enabled}
                  onChange={(e) => setDraft({ ...draft, enabled: e.target.checked })}
                />
                <span>Enabled</span>
              </label>
            </div>

            <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "end" }}>
              <label style={{ minWidth: 130 }}>
                <div style={{ fontSize: 11, color: ui.textMuted, marginBottom: 3 }}>Trigger</div>
                <select
                  id="select-cadence-mode"
                  value={draft.cadence.mode}
                  onChange={(e) =>
                    setDraft({
                      ...draft,
                      cadence: { ...draft.cadence, mode: e.target.value as CadenceMode },
                    })
                  }
                  style={inputStyle}
                >
                  <option value="daily">Daily</option>
                  <option value="weekly">Weekly</option>
                  <option value="monthly">Monthly</option>
                  <option value="cron">Custom Cron</option>
                  <option value="datagroup">Datagroup</option>
                </select>
              </label>

              {draft.cadence.mode === "weekly" && (
                <label style={{ minWidth: 125 }}>
                  <div style={{ fontSize: 11, color: ui.textMuted, marginBottom: 3 }}>Day</div>
                  <select
                    value={draft.cadence.dayOfWeek}
                    onChange={(e) =>
                      setDraft({
                        ...draft,
                        cadence: { ...draft.cadence, dayOfWeek: Number(e.target.value) },
                      })
                    }
                    style={inputStyle}
                  >
                    {WEEKDAYS.map((d) => (
                      <option key={d.value} value={d.value}>
                        {d.label}
                      </option>
                    ))}
                  </select>
                </label>
              )}

              {draft.cadence.mode === "monthly" && (
                <label style={{ width: 95 }}>
                  <div style={{ fontSize: 11, color: ui.textMuted, marginBottom: 3 }}>
                    Day (1–28)
                  </div>
                  <input
                    type="number"
                    min={1}
                    max={28}
                    value={draft.cadence.dayOfMonth}
                    onChange={(e) =>
                      setDraft({
                        ...draft,
                        cadence: { ...draft.cadence, dayOfMonth: Number(e.target.value) || 1 },
                      })
                    }
                    style={inputStyle}
                  />
                </label>
              )}

              {(draft.cadence.mode === "daily" ||
                draft.cadence.mode === "weekly" ||
                draft.cadence.mode === "monthly") && (
                <>
                  <label style={{ width: 82 }}>
                    <div style={{ fontSize: 11, color: ui.textMuted, marginBottom: 3 }}>
                      Hour (0–23)
                    </div>
                    <input
                      type="number"
                      min={0}
                      max={23}
                      value={draft.cadence.hour}
                      onChange={(e) =>
                        setDraft({
                          ...draft,
                          cadence: { ...draft.cadence, hour: Number(e.target.value) || 0 },
                        })
                      }
                      style={inputStyle}
                    />
                  </label>
                  <label style={{ width: 82 }}>
                    <div style={{ fontSize: 11, color: ui.textMuted, marginBottom: 3 }}>
                      Min (0–59)
                    </div>
                    <input
                      type="number"
                      min={0}
                      max={59}
                      value={draft.cadence.minute}
                      onChange={(e) =>
                        setDraft({
                          ...draft,
                          cadence: { ...draft.cadence, minute: Number(e.target.value) || 0 },
                        })
                      }
                      style={inputStyle}
                    />
                  </label>
                </>
              )}

              {draft.cadence.mode === "cron" && (
                <label style={{ flex: "1 1 180px" }}>
                  <div style={{ fontSize: 11, color: ui.textMuted, marginBottom: 3 }}>
                    Crontab (min hour dom mon dow)
                  </div>
                  <input
                    type="text"
                    value={draft.cadence.crontab}
                    onChange={(e) =>
                      setDraft({
                        ...draft,
                        cadence: { ...draft.cadence, crontab: e.target.value },
                      })
                    }
                    placeholder="0 8 * * 1-5"
                    style={inputStyle}
                  />
                </label>
              )}

              {draft.cadence.mode === "datagroup" && (
                <label style={{ flex: "1 1 200px" }}>
                  <div style={{ fontSize: 11, color: ui.textMuted, marginBottom: 3 }}>
                    LookML Datagroup Name
                  </div>
                  <input
                    type="text"
                    value={draft.cadence.datagroup || ""}
                    onChange={(e) =>
                      setDraft({
                        ...draft,
                        cadence: { ...draft.cadence, datagroup: e.target.value },
                      })
                    }
                    placeholder="orders_daily_datagroup"
                    style={inputStyle}
                  />
                </label>
              )}

              {draft.cadence.mode !== "datagroup" && (
                <label style={{ minWidth: 175 }}>
                  <div style={{ fontSize: 11, color: ui.textMuted, marginBottom: 3 }}>Timezone</div>
                  <select
                    value={draft.cadence.timezone}
                    onChange={(e) =>
                      setDraft({
                        ...draft,
                        cadence: { ...draft.cadence, timezone: e.target.value },
                      })
                    }
                    style={inputStyle}
                  >
                    {!COMMON_TIMEZONES.includes(draft.cadence.timezone) && (
                      <option value={draft.cadence.timezone}>{draft.cadence.timezone}</option>
                    )}
                    {COMMON_TIMEZONES.map((tz) => (
                      <option key={tz} value={tz}>
                        {tz}
                      </option>
                    ))}
                  </select>
                </label>
              )}
            </div>
          </div>

          {/* Shared Baseline Filters */}
          {dashFilters.length > 0 && (
            <div
              style={{
                border: `1px solid ${ui.border}`,
                borderRadius: 8,
                padding: "8px 10px",
              }}
            >
              <button
                type="button"
                onClick={() => setShowBaseFilters((v) => !v)}
                style={{
                  background: "none",
                  border: "none",
                  padding: 0,
                  color: ui.textPrimary,
                  fontSize: 12,
                  fontWeight: 700,
                  cursor: "pointer",
                }}
              >
                {showBaseFilters ? "▼" : "▶"} Shared Baseline Dashboard Filters (
                {Object.values(draft.base_filters || {}).filter((v) => String(v).trim()).length}{" "}
                active)
              </button>
              {showBaseFilters && (
                <div
                  style={{
                    marginTop: 8,
                    display: "grid",
                    gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))",
                    gap: 8,
                  }}
                >
                  {dashFilters.map((f) => (
                    <label key={f.name}>
                      <div style={{ fontSize: 11, color: ui.textMuted, marginBottom: 2 }}>
                        {f.title}
                      </div>
                      <input
                        type="text"
                        value={draft.base_filters[f.name] ?? ""}
                        onChange={(e) =>
                          setDraft({
                            ...draft,
                            base_filters: { ...draft.base_filters, [f.name]: e.target.value },
                          })
                        }
                        placeholder={f.defaultValue || "Leave blank for dashboard default"}
                        style={inputStyle}
                      />
                    </label>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* Fanout Slice Generator (Load Values limit 500 + Filter & Sort + Paste + Manual) */}
          <div
            style={{
              border: `1px solid ${ui.border}`,
              borderRadius: 8,
              padding: 10,
              display: "grid",
              gap: 10,
            }}
          >
            <div
              style={{
                display: "flex",
                flexWrap: "wrap",
                justifyContent: "space-between",
                alignItems: "center",
                gap: 8,
              }}
            >
              <div style={{ fontSize: 12.5, fontWeight: 700 }}>
                Fanout Slices ({draft.slices.length})
              </div>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                <button
                  type="button"
                  onClick={() => setPasteOpen((v) => !v)}
                  style={btnSecondary}
                >
                  {pasteOpen ? "Hide Paste" : "Paste Values"}
                </button>
                <button
                  id="btn-add-slice-row"
                  type="button"
                  onClick={() =>
                    setDraft({
                      ...draft,
                      slices: [
                        ...draft.slices,
                        {
                          id: makeShortId(),
                          label: `Slice ${draft.slices.length + 1}`,
                          slug: `slice_${draft.slices.length + 1}`,
                          filters: { [primarySliceFilterKey]: "" },
                        },
                      ],
                    })
                  }
                  style={btnSecondary}
                >
                  + Add Slice Row
                </button>
              </div>
            </div>

            {/* Query-time Filter & Sort Value Loader */}
            <div
              style={{
                display: "flex",
                flexWrap: "wrap",
                gap: 8,
                alignItems: "end",
                background: ui.inputBg,
                padding: 8,
                borderRadius: 6,
                border: `1px solid ${ui.border}`,
              }}
            >
              <label style={{ minWidth: 160, flex: "1 1 160px" }}>
                <div style={{ fontSize: 11, color: ui.textMuted, marginBottom: 3 }}>
                  Fanout Dashboard Filter
                </div>
                {dashFilters.length > 0 ? (
                  <select
                    id="select-fanout-filter"
                    value={activeFilterName}
                    onChange={(e) => setActiveFilterName(e.target.value)}
                    style={inputStyle}
                  >
                    {dashFilters.map((f) => (
                      <option key={f.name} value={f.name}>
                        {f.title} ({f.name})
                      </option>
                    ))}
                  </select>
                ) : (
                  <input
                    type="text"
                    value={activeFilterName}
                    onChange={(e) => setActiveFilterName(e.target.value)}
                    placeholder="Dashboard Filter Name (e.g. Brand)"
                    style={inputStyle}
                  />
                )}
              </label>

              {selectedDashFilter?.dimension && (
                <>
                  <label style={{ flex: "1 1 150px", minWidth: 140 }}>
                    <div style={{ fontSize: 11, color: ui.textMuted, marginBottom: 3 }}>
                      Query Filter (Looker expr)
                    </div>
                    <input
                      id="input-query-filter"
                      type="text"
                      value={queryFilterExpr}
                      onChange={(e) => setQueryFilterExpr(e.target.value)}
                      placeholder="e.g. %Apparel%, -NULL"
                      style={inputStyle}
                    />
                  </label>

                  <label style={{ flex: "1 1 180px", minWidth: 150 }}>
                    <div style={{ fontSize: 11, color: ui.textMuted, marginBottom: 3 }}>
                      Sort Query By
                    </div>
                    <select
                      id="select-query-sort-field"
                      value={querySortField || selectedDashFilter.dimension}
                      onChange={(e) => {
                        const nextField = e.target.value;
                        setQuerySortField(nextField);
                        const opt = exploreSortOptions.find((o) => o.name === nextField);
                        if (opt?.kind === "measure") setQuerySortDir("desc");
                      }}
                      style={inputStyle}
                    >
                      <option value={selectedDashFilter.dimension}>
                        {selectedDashFilter.dimension} (filter dimension)
                      </option>
                      {exploreSortOptions
                        .filter((o) => o.name !== selectedDashFilter.dimension)
                        .map((o) => (
                          <option key={o.name} value={o.name}>
                            {o.label}
                          </option>
                        ))}
                    </select>
                  </label>

                  <label style={{ width: 92 }}>
                    <div style={{ fontSize: 11, color: ui.textMuted, marginBottom: 3 }}>Order</div>
                    <select
                      id="select-query-sort-dir"
                      value={querySortDir}
                      onChange={(e) => setQuerySortDir(e.target.value as "asc" | "desc")}
                      style={inputStyle}
                    >
                      <option value="asc">ASC ▲</option>
                      <option value="desc">DESC ▼</option>
                    </select>
                  </label>

                  <button
                    id="btn-load-filter-values"
                    type="button"
                    disabled={loadingValues}
                    onClick={() => void handleLoadDimensionValues()}
                    style={{
                      ...btnSecondary,
                      background: ui.accent,
                      color: ui.accentContrast,
                      border: "none",
                      fontWeight: 700,
                      padding: "6px 12px",
                    }}
                  >
                    {loadingValues ? "Loading..." : "Load Values (max 500)"}
                  </button>
                </>
              )}
            </div>

            {pasteOpen && (
              <div
                style={{
                  background: ui.inputBg,
                  border: `1px solid ${ui.border}`,
                  borderRadius: 6,
                  padding: 8,
                  display: "grid",
                  gap: 6,
                }}
              >
                <textarea
                  rows={4}
                  value={pasteText}
                  onChange={(e) => setPasteText(e.target.value)}
                  placeholder={"Paste one filter value per line...\nLevi's\nCalvin Klein\nDockers"}
                  style={{ ...inputStyle, fontFamily: "monospace" }}
                />
                <div style={{ display: "flex", justifyContent: "flex-end", gap: 6 }}>
                  <button
                    type="button"
                    onClick={() => {
                      appendValuesAsSlices(pasteText.split(/\r?\n/));
                      setPasteText("");
                      setPasteOpen(false);
                    }}
                    style={{
                      ...btnSecondary,
                      background: ui.accent,
                      color: ui.accentContrast,
                      border: "none",
                    }}
                  >
                    Add Pasted Lines as Slices
                  </button>
                </div>
              </div>
            )}

            {/* Loaded Values Picker with Client-Side Filter & Sort */}
            {loadedValues.length > 0 && (
              <div
                style={{
                  background: ui.inputBg,
                  border: `1px solid ${ui.border}`,
                  borderRadius: 6,
                  padding: 8,
                  display: "grid",
                  gap: 6,
                }}
              >
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6, alignItems: "center" }}>
                  <input
                    type="text"
                    value={valueSearch}
                    onChange={(e) => setValueSearch(e.target.value)}
                    placeholder={`Filter ${loadedValues.length} loaded value(s)...`}
                    style={{ ...inputStyle, flex: "1 1 180px" }}
                  />
                  <select
                    value={valueClientSort}
                    onChange={(e) =>
                      setValueClientSort(e.target.value as "query" | "asc" | "desc")
                    }
                    style={{ ...inputStyle, width: "auto" }}
                  >
                    <option value="query">Sort: Query Order</option>
                    <option value="asc">Sort: A → Z</option>
                    <option value="desc">Sort: Z → A</option>
                  </select>
                  <button
                    type="button"
                    onClick={() =>
                      setSelectedValues(new Set(visibleLoadedValues.map((v) => v.value)))
                    }
                    style={btnSecondary}
                  >
                    Select All ({visibleLoadedValues.length})
                  </button>
                  <button
                    type="button"
                    disabled={selectedValues.size === 0}
                    onClick={() => {
                      appendValuesAsSlices(Array.from(selectedValues));
                      setSelectedValues(new Set());
                    }}
                    style={{
                      ...btnSecondary,
                      background: ui.accent,
                      color: ui.accentContrast,
                      border: "none",
                      opacity: selectedValues.size > 0 ? 1 : 0.5,
                    }}
                  >
                    + Add {selectedValues.size} Selected
                  </button>
                  <button
                    type="button"
                    onClick={() => setLoadedValues([])}
                    style={btnSecondary}
                  >
                    Close
                  </button>
                </div>
                <div
                  style={{
                    maxHeight: 170,
                    overflowY: "auto",
                    display: "flex",
                    flexWrap: "wrap",
                    gap: 5,
                    paddingTop: 4,
                  }}
                >
                  {visibleLoadedValues.map((item) => {
                    const checked = selectedValues.has(item.value);
                    return (
                      <button
                        key={item.value}
                        type="button"
                        onClick={() => {
                          const next = new Set(selectedValues);
                          if (checked) next.delete(item.value);
                          else next.add(item.value);
                          setSelectedValues(next);
                        }}
                        style={{
                          padding: "3px 8px",
                          borderRadius: 999,
                          fontSize: 11.5,
                          cursor: "pointer",
                          border: `1px solid ${checked ? ui.accent : ui.border}`,
                          background: checked ? ui.accent : ui.cardBg,
                          color: checked ? ui.accentContrast : ui.textPrimary,
                        }}
                      >
                        {item.value}
                        {item.metric !== undefined ? ` (${item.metric})` : ""}
                      </button>
                    );
                  })}
                </div>
              </div>
            )}

            {/* Slice Table Filter & Sort Controls */}
            {draft.slices.length > 0 && (
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input
                  id="input-slice-table-filter"
                  type="text"
                  value={sliceSearch}
                  onChange={(e) => setSliceSearch(e.target.value)}
                  placeholder="Filter slice rows by label, output slug, or filter value..."
                  style={{ ...inputStyle, maxWidth: 340 }}
                />
                <span style={{ fontSize: 11.5, color: ui.textMuted }}>
                  Showing {displayedDraftSlices.length} of {draft.slices.length}
                </span>
              </div>
            )}

            <div
              style={{
                border: `1px solid ${ui.border}`,
                borderRadius: 6,
                overflowX: "auto",
                maxHeight: 320,
              }}
            >
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr
                    style={{
                      background: ui.inputBg,
                      borderBottom: `1px solid ${ui.border}`,
                      textAlign: "left",
                    }}
                  >
                    {(
                      [
                        ["label", "Slice Label"],
                        ["slug", "Output Subpath"],
                        ["value", `Filter: ${primarySliceFilterKey}`],
                        ["status", "Plan Status"],
                      ] as const
                    ).map(([key, col]) => (
                      <th key={key} style={{ padding: "6px 8px", fontWeight: 700 }}>
                        <button
                          type="button"
                          onClick={() => toggleSliceSort(key)}
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
                          {col}
                          {sliceSortKey === key ? (sliceSortAsc ? " ▲" : " ▼") : ""}
                        </button>
                      </th>
                    ))}
                    <th style={{ padding: "6px 8px", width: 40 }} />
                  </tr>
                </thead>
                <tbody>
                  {displayedDraftSlices.length === 0 ? (
                    <tr>
                      <td
                        colSpan={5}
                        style={{ padding: 14, textAlign: "center", color: ui.textMuted }}
                      >
                        {draft.slices.length === 0
                          ? "Load filter values above, paste lines, or click + Add Slice Row."
                          : "No slice rows match your filter."}
                      </td>
                    </tr>
                  ) : (
                    displayedDraftSlices.map((s) => {
                      const live = s.plan_id ? livePlansById[s.plan_id] : undefined;
                      const outPrefix = buildSliceOutputPrefix(draft.output_folder, s.slug);
                      return (
                        <tr key={s.id} style={{ borderBottom: `1px solid ${ui.border}` }}>
                          <td style={{ padding: "5px 8px" }}>
                            <input
                              type="text"
                              value={s.label}
                              onChange={(e) =>
                                setDraft({
                                  ...draft,
                                  slices: draft.slices.map((row) =>
                                    row.id === s.id ? { ...row, label: e.target.value } : row
                                  ),
                                })
                              }
                              style={inputStyle}
                            />
                          </td>
                          <td style={{ padding: "5px 8px" }}>
                            <input
                              type="text"
                              value={s.slug}
                              onChange={(e) =>
                                setDraft({
                                  ...draft,
                                  slices: draft.slices.map((row) =>
                                    row.id === s.id ? { ...row, slug: e.target.value } : row
                                  ),
                                })
                              }
                              style={inputStyle}
                            />
                            <div style={{ fontSize: 10.5, color: ui.textMuted, marginTop: 2 }}>
                              <code>{outPrefix}_&lt;ts&gt;.pptx</code>
                            </div>
                          </td>
                          <td style={{ padding: "5px 8px" }}>
                            <input
                              type="text"
                              value={s.filters?.[primarySliceFilterKey] ?? ""}
                              onChange={(e) =>
                                setDraft({
                                  ...draft,
                                  slices: draft.slices.map((row) =>
                                    row.id === s.id
                                      ? {
                                          ...row,
                                          filters: {
                                            ...row.filters,
                                            [primarySliceFilterKey]: e.target.value,
                                          },
                                        }
                                      : row
                                  ),
                                })
                              }
                              placeholder="Filter value..."
                              style={inputStyle}
                            />
                          </td>
                          <td style={{ padding: "5px 8px", whiteSpace: "nowrap", fontSize: 11.5 }}>
                            {live ? (
                              <span style={{ color: isDark ? "#34d399" : "#047857" }}>
                                Plan #{s.plan_id}
                              </span>
                            ) : s.plan_id ? (
                              <span style={{ color: "#f59e0b" }}>Missing in Looker</span>
                            ) : (
                              <span style={{ color: ui.textMuted }}>New (unsynced)</span>
                            )}
                          </td>
                          <td style={{ padding: "5px 8px", textAlign: "right" }}>
                            <button
                              type="button"
                              aria-label={`Remove slice ${s.label}`}
                              onClick={() =>
                                setDraft({
                                  ...draft,
                                  slices: draft.slices.filter((row) => row.id !== s.id),
                                })
                              }
                              style={{
                                background: "none",
                                border: "none",
                                color: "#ef4444",
                                cursor: "pointer",
                                fontWeight: 700,
                                fontSize: 13,
                              }}
                            >
                              ✕
                            </button>
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </form>
      ) : (
        <>
          {/* Fleet Toolbar: Filter by Dashboard + Refresh + New Master Schedule */}
          <div
            style={{
              background: ui.cardBg,
              border: `1px solid ${ui.border}`,
              borderRadius: 10,
              padding: "10px 12px",
              display: "flex",
              flexWrap: "wrap",
              gap: 8,
              alignItems: "center",
              justifyContent: "space-between",
            }}
          >
            <div
              style={{
                display: "flex",
                flexWrap: "wrap",
                gap: 8,
                alignItems: "center",
                flex: "1 1 320px",
              }}
            >
              <div style={{ flex: "1 1 260px", maxWidth: 420 }}>
                <DashboardPicker
                  core40SDK={core40SDK}
                  extensionSDK={extensionSDK}
                  selectedDashboardId={filterDashId}
                  onSelect={(id: string) => setFilterDashId(id)}
                  ui={ui}
                  isDark={isDark}
                />
              </div>
              {filterDashId && (
                <button
                  type="button"
                  onClick={() => setFilterDashId("")}
                  style={btnSecondary}
                >
                  All Dashboards
                </button>
              )}
            </div>

            <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
              <button
                id="btn-scheduler-refresh"
                type="button"
                disabled={loading || Boolean(busyMessage)}
                onClick={() => void refreshSchedules()}
                style={btnSecondary}
              >
                <Svg size={13}>
                  <path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8" />
                  <path d="M21 3v5h-5" />
                </Svg>
                <span>{loading ? "Loading..." : "Refresh"}</span>
              </button>
              <button
                id="btn-new-fanout"
                type="button"
                disabled={Boolean(busyMessage)}
                onClick={() =>
                  openEditor(
                    createBlankSpec(effectiveUserId || "me", {
                      dashboardId: filterDashId || undefined,
                    }),
                    null
                  )
                }
                style={{
                  ...btnSecondary,
                  background: ui.accent,
                  color: ui.accentContrast,
                  border: "none",
                  fontWeight: 700,
                }}
              >
                + New Master Schedule
              </button>
            </div>
          </div>

          {/* Master Fanout Schedule Cards */}
          {loading ? (
            <div
              style={{
                background: ui.cardBg,
                border: `1px solid ${ui.border}`,
                borderRadius: 10,
                padding: 22,
                textAlign: "center",
                fontSize: 13,
                color: ui.textMuted,
              }}
            >
              Loading Master Fanout schedules...
            </div>
          ) : records.length === 0 ? (
            <div
              data-testid="scheduler-empty-state"
              style={{
                background: ui.cardBg,
                border: `1px solid ${ui.border}`,
                borderRadius: 10,
                padding: 24,
                textAlign: "center",
              }}
            >
              <div style={{ fontSize: 14.5, fontWeight: 700, marginBottom: 6 }}>
                No Master Fanout Schedules{filterDashId ? " for this Dashboard" : ""}
              </div>
              <p style={{ margin: "0 0 14px", fontSize: 12.5, color: ui.textMuted }}>
                Create one Master Schedule to fan out a PowerPoint template across N dashboard
                filter slices on a single cadence.
              </p>
              <button
                type="button"
                onClick={() =>
                  openEditor(
                    createBlankSpec(effectiveUserId || "me", {
                      dashboardId: filterDashId || undefined,
                    }),
                    null
                  )
                }
                style={{
                  ...btnSecondary,
                  background: ui.accent,
                  color: ui.accentContrast,
                  border: "none",
                  fontWeight: 700,
                }}
              >
                + Create First Master Schedule
              </button>
            </div>
          ) : (
            records.map((rec) => {
              const { spec } = rec;
              const driftedCount = spec.slices.filter((s) =>
                isSlicePlanDrifted(spec, s, s.plan_id ? livePlansById[s.plan_id] : undefined)
              ).length;
              const liveCount = spec.slices.filter(
                (s) => s.plan_id && livePlansById[s.plan_id]
              ).length;
              const folderPrefix = `${sanitizeOutputFolder(spec.output_folder, "fanout")}/`;
              const samplePlan = spec.slices
                .map((s) => (s.plan_id ? livePlansById[s.plan_id] : null))
                .find(Boolean);

              return (
                <div
                  key={rec.artifactKey}
                  data-testid={`fanout-card-${spec.id}`}
                  style={{
                    background: ui.cardBg,
                    border: `1px solid ${driftedCount > 0 ? "#f59e0b" : ui.border}`,
                    borderRadius: 10,
                    padding: 12,
                    display: "grid",
                    gap: 10,
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      flexWrap: "wrap",
                      justifyContent: "space-between",
                      alignItems: "flex-start",
                      gap: 8,
                    }}
                  >
                    <div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                        <span style={{ fontSize: 15, fontWeight: 700 }}>{spec.title}</span>
                        <span
                          style={{
                            fontSize: 11,
                            fontWeight: 700,
                            padding: "2px 8px",
                            borderRadius: 999,
                            background: !spec.enabled
                              ? ui.inputBg
                              : driftedCount > 0
                                ? isDark
                                  ? "#451a03"
                                  : "#fef3c7"
                                : isDark
                                  ? "#064e3b"
                                  : "#d1fae5",
                            color: !spec.enabled
                              ? ui.textMuted
                              : driftedCount > 0
                                ? "#f59e0b"
                                : isDark
                                  ? "#34d399"
                                  : "#047857",
                          }}
                        >
                          {!spec.enabled
                            ? "Paused"
                            : driftedCount > 0
                              ? `${driftedCount} slice(s) need sync`
                              : `Active (${liveCount}/${spec.slices.length} synced)`}
                        </span>
                      </div>
                      <div style={{ fontSize: 12, color: ui.textSecondary, marginTop: 3 }}>
                        <strong>{spec.dashboard_title || spec.dashboard_id}</strong> · Template{" "}
                        <code>{spec.template}</code> · {describeCadence(spec.cadence)}
                        {samplePlan?.next_run_at && spec.enabled
                          ? ` · Next: ${new Date(samplePlan.next_run_at).toLocaleString()}`
                          : ""}
                      </div>
                    </div>

                    <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                      {driftedCount > 0 && (
                        <button
                          type="button"
                          disabled={Boolean(busyMessage)}
                          onClick={() => void handleCardSync(rec)}
                          style={{
                            ...btnSecondary,
                            background: "#f59e0b",
                            color: "#0f172a",
                            border: "none",
                            fontWeight: 700,
                          }}
                        >
                          Sync Plans ({driftedCount})
                        </button>
                      )}
                      <button
                        type="button"
                        disabled={Boolean(busyMessage)}
                        onClick={() => void handleRunAllNow(rec)}
                        style={{
                          ...btnSecondary,
                          borderColor: ui.accent,
                          color: ui.accent,
                          fontWeight: 700,
                        }}
                      >
                        Run All Now
                      </button>
                      <button
                        type="button"
                        disabled={Boolean(busyMessage)}
                        onClick={() => void handleToggleEnabled(rec)}
                        style={btnSecondary}
                      >
                        {spec.enabled ? "Pause All" : "Resume All"}
                      </button>
                      {canOpenViewer && (
                        <button
                          type="button"
                          onClick={() => onOpenViewerSearch(folderPrefix)}
                          style={btnSecondary}
                        >
                          View Outputs →
                        </button>
                      )}
                      <button
                        type="button"
                        onClick={() => openEditor(spec, rec.artifactKey)}
                        style={btnSecondary}
                      >
                        Edit
                      </button>
                      {confirmDeleteId === spec.id ? (
                        <>
                          <button
                            type="button"
                            disabled={Boolean(busyMessage)}
                            onClick={() => void handleDeleteMaster(rec)}
                            style={{
                              ...btnSecondary,
                              background: "#dc2626",
                              color: "#ffffff",
                              border: "none",
                              fontWeight: 700,
                            }}
                          >
                            Confirm Delete
                          </button>
                          <button
                            type="button"
                            onClick={() => setConfirmDeleteId("")}
                            style={btnSecondary}
                          >
                            Cancel
                          </button>
                        </>
                      ) : (
                        <button
                          type="button"
                          onClick={() => setConfirmDeleteId(spec.id)}
                          style={{ ...btnSecondary, color: "#ef4444" }}
                        >
                          Delete
                        </button>
                      )}
                    </div>
                  </div>

                  {/* Slice Fleet Summary Chips / Mini Table */}
                  <div
                    style={{
                      border: `1px solid ${ui.border}`,
                      borderRadius: 8,
                      overflowX: "auto",
                    }}
                  >
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11.5 }}>
                      <thead>
                        <tr
                          style={{
                            background: ui.inputBg,
                            borderBottom: `1px solid ${ui.border}`,
                            textAlign: "left",
                          }}
                        >
                          <th style={{ padding: "5px 8px" }}>Slice</th>
                          <th style={{ padding: "5px 8px" }}>Filters</th>
                          <th style={{ padding: "5px 8px" }}>Output Prefix</th>
                          <th style={{ padding: "5px 8px" }}>Looker Plan</th>
                          <th style={{ padding: "5px 8px", textAlign: "right" }}>Outputs</th>
                        </tr>
                      </thead>
                      <tbody>
                        {spec.slices.map((s) => {
                          const live = s.plan_id ? livePlansById[s.plan_id] : undefined;
                          const drifted = isSlicePlanDrifted(spec, s, live);
                          const subPrefix = buildSliceOutputPrefix(
                            spec.output_folder,
                            s.slug || s.label
                          );
                          return (
                            <tr key={s.id} style={{ borderBottom: `1px solid ${ui.border}` }}>
                              <td style={{ padding: "5px 8px", fontWeight: 600 }}>{s.label}</td>
                              <td style={{ padding: "5px 8px", color: ui.textSecondary }}>
                                <code>
                                  {serializeFiltersString(spec.base_filters, s.filters) || "(none)"}
                                </code>
                              </td>
                              <td style={{ padding: "5px 8px", color: ui.textSecondary }}>
                                <code>{subPrefix}_&lt;ts&gt;.pptx</code>
                              </td>
                              <td style={{ padding: "5px 8px" }}>
                                {!live ? (
                                  <span style={{ color: "#f59e0b", fontWeight: 600 }}>
                                    Missing — Sync to recreate
                                  </span>
                                ) : drifted ? (
                                  <span style={{ color: "#f59e0b" }}>Drifted (#{s.plan_id})</span>
                                ) : (
                                  <span style={{ color: isDark ? "#34d399" : "#047857" }}>
                                    #{s.plan_id} {live.enabled === false ? "(paused)" : "(synced)"}
                                    {live.last_run_at
                                      ? ` · Last: ${new Date(live.last_run_at).toLocaleDateString()}`
                                      : ""}
                                  </span>
                                )}
                              </td>
                              <td style={{ padding: "5px 8px", textAlign: "right" }}>
                                {canOpenViewer && (
                                  <button
                                    type="button"
                                    onClick={() => onOpenViewerSearch(subPrefix)}
                                    style={{
                                      background: "none",
                                      border: "none",
                                      color: ui.accent,
                                      fontSize: 11.5,
                                      fontWeight: 700,
                                      cursor: "pointer",
                                      padding: 0,
                                    }}
                                  >
                                    View →
                                  </button>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              );
            })
          )}
        </>
      )}
    </div>
  );
};
