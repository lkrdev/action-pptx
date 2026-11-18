export type CadenceMode = "daily" | "weekly" | "monthly" | "cron" | "datagroup";

export interface CadenceSpec {
  mode: CadenceMode;
  hour: number;
  minute: number;
  dayOfWeek: number;
  dayOfMonth: number;
  crontab: string;
  datagroup?: string;
  timezone: string;
}

export interface FanoutSlice {
  id: string;
  label: string;
  slug: string;
  filters: Record<string, string>;
  plan_id?: string;
}

export interface MasterFanoutSpec {
  id: string;
  title: string;
  owner_user_id: string;
  dashboard_id: string;
  dashboard_title: string;
  template: string;
  theme: string;
  theme_palette?: Record<string, unknown>;
  output_folder: string;
  cadence: CadenceSpec;
  enabled: boolean;
  base_filters: Record<string, string>;
  slices: FanoutSlice[];
  updated_at: string;
}

export interface TemplateDashboardLink {
  id: string;
  title: string;
}

export interface TemplateLinksDoc {
  template: string;
  dashboards: TemplateDashboardLink[];
  updated_at: string;
}

export const TEMPLATE_LINKS_SEARCH_PATTERN = "tpl_links.%";

const WEEKDAY_NAMES = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"] as const;

/**
 * Hex-escape non-alphanumerics using `-xx` so Looker `search_artifacts`
 * SQL `LIKE` wildcards (`%` and `_`) and `.` segment delimiters never collide.
 */
export function encodeDashKey(raw: string): string {
  return (
    String(raw || "")
      .trim()
      .replace(
        /[^a-zA-Z0-9]/gu,
        (ch) => `-${(ch.codePointAt(0) ?? 0).toString(16).padStart(2, "0")}`
      ) || "0"
  );
}

// ponytail: keep "fanout.u-*" artifact key prefix for zero-migration backward compatibility with existing saved batches
export function buildFanoutArtifactKey(
  userId: string,
  dashboardId: string,
  fanoutId: string
): string {
  return `fanout.u-${encodeDashKey(userId)}.d-${encodeDashKey(dashboardId)}.${encodeDashKey(fanoutId)}`;
}

export function buildFanoutSearchPattern(userId: string, dashboardId?: string): string {
  const dash = String(dashboardId || "").trim();
  if (dash) {
    return `fanout.u-${encodeDashKey(userId)}.d-${encodeDashKey(dash)}.%`;
  }
  return `fanout.u-${encodeDashKey(userId)}.%`;
}

export function buildTemplateLinksArtifactKey(templateName: string): string {
  return `tpl_links.${encodeDashKey(templateName)}`;
}

export function mergeTemplateDashboardLink(
  existing: TemplateLinksDoc | null | undefined,
  templateName: string,
  dash: TemplateDashboardLink,
  nowIso = new Date().toISOString()
): TemplateLinksDoc {
  const cleanTpl = String(templateName || "").trim();
  const cleanId = String(dash.id || "").trim();
  const cleanTitle = String(dash.title || cleanId).trim() || cleanId;
  const prior = Array.isArray(existing?.dashboards) ? existing.dashboards : [];
  const rest = prior
    .map((d) => ({ id: String(d?.id || "").trim(), title: String(d?.title || d?.id || "").trim() }))
    .filter((d) => d.id && d.id !== cleanId);
  return {
    template: cleanTpl,
    dashboards: cleanId ? [{ id: cleanId, title: cleanTitle }, ...rest] : rest,
    updated_at: nowIso,
  };
}

export function parseTemplateLinksArtifacts(
  rawArtifacts: any[]
): Record<string, TemplateDashboardLink[]> {
  const out: Record<string, TemplateDashboardLink[]> = {};
  for (const art of Array.isArray(rawArtifacts) ? rawArtifacts : []) {
    try {
      const parsed = JSON.parse(String(art?.value || "")) as Partial<TemplateLinksDoc>;
      const tpl = String(parsed?.template || "").trim();
      if (!tpl || !Array.isArray(parsed?.dashboards)) continue;
      const links = parsed.dashboards
        .map((d) => ({
          id: String(d?.id || "").trim(),
          title: String(d?.title || d?.id || "").trim(),
        }))
        .filter((d) => Boolean(d.id));
      if (links.length > 0) out[tpl] = links;
    } catch {
      // Ignore malformed artifact entries
    }
  }
  return out;
}

export function findMatchingLinkedTemplate(
  templatesList: string[],
  linksMap: Record<string, TemplateDashboardLink[]>,
  dashboardId: string,
  dashboardTitle?: string
): string {
  const cleanId = String(dashboardId || "").trim();
  const cleanTitle = String(dashboardTitle || "").trim().toLowerCase();
  if (!cleanId && !cleanTitle) return "";
  const candidates = templatesList.length > 0 ? templatesList : Object.keys(linksMap).sort();
  if (cleanId) {
    const byId = candidates.find((t) => (linksMap[t] || []).some((d) => d.id === cleanId));
    if (byId) return byId;
  }
  if (cleanTitle) {
    const byTitle = candidates.find((t) =>
      (linksMap[t] || []).some((d) => d.title.trim().toLowerCase() === cleanTitle)
    );
    if (byTitle) return byTitle;
  }
  return "";
}

export async function upsertTemplateLinkArtifact(
  core40SDK: any,
  namespace: string,
  templateName: string,
  dash: TemplateDashboardLink
): Promise<TemplateLinksDoc | null> {
  const cleanTpl = String(templateName || "").trim();
  const cleanId = String(dash?.id || "").trim();
  if (!cleanTpl || !cleanId || typeof core40SDK?.update_artifacts !== "function") return null;
  const key = buildTemplateLinksArtifactKey(cleanTpl);
  try {
    const existing: any[] =
      typeof core40SDK.artifact === "function"
        ? await core40SDK.ok(core40SDK.artifact({ namespace, key })).catch(() => [])
        : [];
    let prior: TemplateLinksDoc | null = null;
    let ver: number | undefined;
    if (Array.isArray(existing) && existing[0]) {
      if (typeof existing[0].version === "number") ver = existing[0].version;
      try {
        prior = JSON.parse(String(existing[0].value || "")) as TemplateLinksDoc;
      } catch {
        prior = null;
      }
    }
    const merged = mergeTemplateDashboardLink(prior, cleanTpl, {
      id: cleanId,
      title: String(dash.title || cleanId),
    });
    await core40SDK.ok(
      core40SDK.update_artifacts(namespace, [
        {
          key,
          value: JSON.stringify(merged),
          content_type: "application/json",
          ...(ver !== undefined ? { version: ver } : {}),
        },
      ])
    );
    return merged;
  } catch {
    return null;
  }
}

export function extractDashboardModels(dash: any): string[] {
  const models = new Set<string>();
  const add = (val: unknown) => {
    const s = String(val || "").trim();
    if (s) models.add(s);
  };
  const dashId = String(dash?.id || "").trim();
  if (dashId.includes("::")) add(dashId.split("::")[0]);
  add(dash?.model?.id);
  for (const f of Array.isArray(dash?.dashboard_filters) ? dash.dashboard_filters : []) {
    add(f?.model);
  }
  for (const el of Array.isArray(dash?.dashboard_elements) ? dash.dashboard_elements : []) {
    add(el?.query?.model);
    add(el?.look?.query?.model);
    add(el?.result_maker?.query?.model);
    for (const fl of Array.isArray(el?.result_maker?.filterables)
      ? el.result_maker.filterables
      : []) {
      add(fl?.model);
    }
  }
  return Array.from(models).sort();
}

export function formatDatagroupValue(modelName: string, name: string): string {
  return `${String(modelName || "").trim()}::${String(name || "").trim()}`;
}

export function slugifySegment(raw: string, fallback = "delivery"): string {
  const cleaned = String(raw || "")
    .trim()
    .replace(/[^a-zA-Z0-9_-]+/g, "_")
    .replace(/^_+|_+$/g, "");
  return cleaned || fallback;
}

export function sanitizeOutputFolder(raw: string, fallback = "batch"): string {
  const stripped = String(raw || "")
    .replace(/\.pptx$/i, "")
    .trim();
  const parts = stripped
    .split("/")
    .map((s) => s.trim())
    .filter((s) => s && s !== "." && s !== "..")
    .map((s) => slugifySegment(s, ""))
    .filter(Boolean);
  return parts.join("/") || fallback;
}

export function buildSliceOutputPrefix(outputFolder: string, sliceSlug: string): string {
  return `${sanitizeOutputFolder(outputFolder, "batch")}/${slugifySegment(sliceSlug, "delivery")}`;
}

export function mergeEffectiveFilters(
  baseFilters: Record<string, string>,
  sliceFilters: Record<string, string>
): Record<string, string> {
  const merged: Record<string, string> = {};
  for (const src of [baseFilters, sliceFilters]) {
    for (const [k, v] of Object.entries(src || {})) {
      const key = k.trim();
      const val = String(v ?? "").trim();
      if (key && val) merged[key] = val;
    }
  }
  return merged;
}

export function buildEmbedFilterParams(
  baseFilters: Record<string, string>,
  sliceFilters: Record<string, string>,
  dashboardFilterDefaults: Array<{ name: string; defaultValue?: string }>
): Record<string, string> {
  const effective = mergeEffectiveFilters(baseFilters, sliceFilters);
  const out: Record<string, string> = {};
  for (const f of dashboardFilterDefaults) {
    const name = String(f.name || "").trim();
    if (!name) continue;
    out[name] = effective[name] ?? String(f.defaultValue ?? "");
  }
  for (const [k, v] of Object.entries(effective)) {
    out[k] = v;
  }
  return out;
}

export function serializeFiltersString(
  baseFilters: Record<string, string>,
  sliceFilters: Record<string, string>
): string {
  const merged = mergeEffectiveFilters(baseFilters, sliceFilters);
  const params = new URLSearchParams();
  for (const key of Object.keys(merged).sort()) {
    params.set(key, merged[key]);
  }
  const qs = params.toString();
  return qs ? `?${qs}` : "";
}

export function resolveCrontab(cadence: CadenceSpec): string {
  if (cadence.mode === "datagroup") return "";
  if (cadence.mode === "cron") return String(cadence.crontab || "").trim() || "0 8 * * 1";
  const m = Math.max(0, Math.min(59, Math.floor(Number(cadence.minute) || 0)));
  const h = Math.max(0, Math.min(23, Math.floor(Number(cadence.hour) || 0)));
  if (cadence.mode === "daily") return `${m} ${h} * * *`;
  if (cadence.mode === "weekly") {
    const dow = Math.max(0, Math.min(6, Math.floor(Number(cadence.dayOfWeek) || 0)));
    return `${m} ${h} * * ${dow}`;
  }
  const dom = Math.max(1, Math.min(28, Math.floor(Number(cadence.dayOfMonth) || 1)));
  return `${m} ${h} ${dom} * *`;
}

export function describeCadence(cadence: CadenceSpec): string {
  const pad2 = (n: number) => String(Math.max(0, Math.floor(Number(n) || 0))).padStart(2, "0");
  const tz = cadence.timezone || "UTC";
  if (cadence.mode === "datagroup") {
    return `Datagroup · ${cadence.datagroup || "(unset)"}`;
  }
  if (cadence.mode === "cron") {
    return `Cron · ${resolveCrontab(cadence)} (${tz})`;
  }
  const hhmm = `${pad2(cadence.hour)}:${pad2(cadence.minute)}`;
  if (cadence.mode === "daily") return `Daily · ${hhmm} (${tz})`;
  if (cadence.mode === "weekly") {
    const dow = WEEKDAY_NAMES[Math.max(0, Math.min(6, Number(cadence.dayOfWeek) || 0))];
    return `Weekly · ${dow} ${hhmm} (${tz})`;
  }
  return `Monthly · Day ${Math.max(1, Math.min(28, Number(cadence.dayOfMonth) || 1))} ${hhmm} (${tz})`;
}

export function buildScheduledPlanWriteBody(
  spec: MasterFanoutSpec,
  slice: FanoutSlice,
  destinationType: string
): Record<string, unknown> {
  const isLookml = spec.dashboard_id.includes("::");
  const crontab = spec.cadence.mode === "datagroup" ? undefined : resolveCrontab(spec.cadence);
  const datagroup =
    spec.cadence.mode === "datagroup" ? String(spec.cadence.datagroup || "").trim() : undefined;
  const formParams: Record<string, string> = {
    template: spec.template,
    filename_prefix: buildSliceOutputPrefix(spec.output_folder, slice.slug || slice.label),
  };
  if (spec.theme && spec.theme !== "__template_original__" && spec.theme !== "auto") {
    formParams.theme = spec.theme_palette
      ? JSON.stringify({ name: spec.theme, ...spec.theme_palette })
      : spec.theme;
  }

  return {
    name: `${spec.title} — ${slice.label || slice.slug}`,
    ...(spec.owner_user_id ? { user_id: spec.owner_user_id } : {}),
    ...(isLookml
      ? { lookml_dashboard_id: spec.dashboard_id }
      : { dashboard_id: String(spec.dashboard_id) }),
    enabled: Boolean(spec.enabled),
    require_results: false,
    require_no_results: false,
    require_change: false,
    send_all_results: false,
    timezone: spec.cadence.timezone || "UTC",
    ...(datagroup ? { datagroup } : { crontab }),
    filters_string: serializeFiltersString(spec.base_filters, slice.filters),
    scheduled_plan_destination: [
      {
        type: destinationType,
        format: "csv_zip",
        apply_formatting: false,
        apply_vis: true,
        address: "",
        parameters: JSON.stringify(formParams),
      },
    ],
  };
}

export function isSlicePlanDrifted(
  spec: MasterFanoutSpec,
  slice: FanoutSlice,
  livePlan: Record<string, any> | undefined
): boolean {
  if (!slice.plan_id || !livePlan) return true;
  if (Boolean(livePlan.enabled) !== Boolean(spec.enabled)) return true;
  const expectedFilters = serializeFiltersString(spec.base_filters, slice.filters);
  if (String(livePlan.filters_string || "") !== expectedFilters) return true;
  if (spec.cadence.mode === "datagroup") {
    if (String(livePlan.datagroup || "") !== String(spec.cadence.datagroup || "").trim()) {
      return true;
    }
  } else if (String(livePlan.crontab || "").trim() !== resolveCrontab(spec.cadence)) {
    return true;
  }
  const dest = Array.isArray(livePlan.scheduled_plan_destination)
    ? livePlan.scheduled_plan_destination[0]
    : null;
  if (!dest) return true;
  try {
    const params = JSON.parse(String(dest.parameters || "{}"));
    if (String(params.template || "") !== spec.template) return true;
    if (
      String(params.filename_prefix || "") !==
      buildSliceOutputPrefix(spec.output_folder, slice.slug || slice.label)
    ) {
      return true;
    }
  } catch {
    return true;
  }
  return false;
}
