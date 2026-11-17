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

export function slugifySegment(raw: string, fallback = "slice"): string {
  const cleaned = String(raw || "")
    .trim()
    .replace(/[^a-zA-Z0-9_-]+/g, "_")
    .replace(/^_+|_+$/g, "");
  return cleaned || fallback;
}

export function sanitizeOutputFolder(raw: string, fallback = "fanout"): string {
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
  return `${sanitizeOutputFolder(outputFolder, "fanout")}/${slugifySegment(sliceSlug, "slice")}`;
}

export function serializeFiltersString(
  baseFilters: Record<string, string>,
  sliceFilters: Record<string, string>
): string {
  const merged: Record<string, string> = {};
  for (const src of [baseFilters, sliceFilters]) {
    for (const [k, v] of Object.entries(src || {})) {
      const key = k.trim();
      const val = String(v ?? "").trim();
      if (key && val) merged[key] = val;
    }
  }
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
