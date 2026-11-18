export interface DashboardTileSchema {
  title: string;
  slug: string;
  columns: string[];
}

const SLUG_RE = /(^\.+$)|[/?*\\:|"<>\s%]/g;

/**
 * Matches Looker's DashboardZIPProducer filename sanitizer:
 * gsub(/(^\.+$)|[\/?*\\:|\"><\s%]/, '_').toLowerCase()
 */
export function sanitizeTileSlug(title: string): string {
  return title.trim().replace(SLUG_RE, "_").toLowerCase();
}

interface ExploreFieldMeta {
  name?: string | null;
  label?: string | null;
  label_short?: string | null;
}

interface QueryLike {
  model?: string | null;
  view?: string | null;
  fields?: string[] | null;
  dynamic_fields?: string | null;
  vis_config?: Record<string, any> | null;
}

function buildFieldLabelMap(explore: any): Map<string, { label: string; labelShort: string }> {
  const map = new Map<string, { label: string; labelShort: string }>();
  const fieldsObj = explore?.fields;
  if (!fieldsObj) return map;

  for (const group of [fieldsObj.dimensions, fieldsObj.measures, fieldsObj.filters, fieldsObj.parameters]) {
    if (!Array.isArray(group)) continue;
    for (const f of group as ExploreFieldMeta[]) {
      if (f?.name) {
        const label = f.label || f.name;
        const labelShort = f.label_short || label;
        map.set(f.name, { label, labelShort });
      }
    }
  }
  return map;
}

function parseDynamicFields(
  dynamicFieldsJson: string | null | undefined
): Map<string, { label: string; isHidden: boolean }> {
  const result = new Map<string, { label: string; isHidden: boolean }>();
  if (!dynamicFieldsJson) return result;
  try {
    const parsed = JSON.parse(dynamicFieldsJson);
    if (Array.isArray(parsed)) {
      for (const item of parsed) {
        const key = item?.table_calculation || item?.dimension || item?.measure;
        if (key) {
          result.set(String(key), {
            label: String(item.label || key),
            isHidden: Boolean(item.is_disabled),
          });
        }
      }
    }
  } catch {
    // Ignore malformed dynamic_fields JSON
  }
  return result;
}

function resolveColumnsForQuery(
  query: QueryLike,
  fieldMap: Map<string, { label: string; labelShort: string }>
): string[] {
  const visConfig = query.vis_config || {};
  const showViewNames = visConfig.show_view_names ?? true;
  const seriesLabels: Record<string, string> = visConfig.series_labels || {};
  const hiddenFields = new Set<string>([
    ...(Array.isArray(visConfig.hidden_fields) ? visConfig.hidden_fields : []),
    ...(Array.isArray(visConfig.hidden_points_if_no) ? visConfig.hidden_points_if_no : []),
  ]);

  const dynMap = parseDynamicFields(query.dynamic_fields);
  const queryFields = Array.isArray(query.fields) ? query.fields : [];
  const allFieldNames = [...queryFields];
  for (const [dynKey, dynMeta] of dynMap.entries()) {
    if (!dynMeta.isHidden && !allFieldNames.includes(dynKey)) {
      allFieldNames.push(dynKey);
    }
  }

  const columns: string[] = [];
  for (const fieldName of allFieldNames) {
    if (hiddenFields.has(fieldName)) continue;
    if (seriesLabels[fieldName]) {
      columns.push(seriesLabels[fieldName]);
      continue;
    }
    const dyn = dynMap.get(fieldName);
    if (dyn) {
      columns.push(dyn.label);
      continue;
    }
    const meta = fieldMap.get(fieldName);
    if (meta) {
      columns.push(showViewNames ? meta.label : meta.labelShort);
    } else {
      columns.push(fieldName);
    }
  }
  return columns;
}

/**
 * Loads a Looker Dashboard's tile titles, slugs, and csv_zip column headers
 * using cached lookml_model_explore metadata (zero warehouse SQL execution).
 */
export async function fetchDashboardTileSchemas(
  core40SDK: any,
  dashboardId: string,
  preloadedDash?: any
): Promise<DashboardTileSchema[]> {
  const dash = preloadedDash ?? (await core40SDK.ok(core40SDK.dashboard(dashboardId)));
  const elements: any[] = Array.isArray(dash?.dashboard_elements) ? dash.dashboard_elements : [];

  const exploreCache = new Map<string, Map<string, { label: string; labelShort: string }>>();

  async function getExploreFieldMap(
    model: string | null | undefined,
    view: string | null | undefined
  ): Promise<Map<string, { label: string; labelShort: string }>> {
    if (!model || !view) return new Map();
    const key = `${model}::${view}`;
    const cached = exploreCache.get(key);
    if (cached) return cached;
    try {
      const explore = await core40SDK.ok(
        core40SDK.lookml_model_explore({
          lookml_model_name: model,
          explore_name: view,
          fields: "fields(dimensions(name,label,label_short),measures(name,label,label_short),filters(name,label,label_short),parameters(name,label,label_short))",
        })
      );
      const map = buildFieldLabelMap(explore);
      exploreCache.set(key, map);
      return map;
    } catch {
      const empty = new Map<string, { label: string; labelShort: string }>();
      exploreCache.set(key, empty);
      return empty;
    }
  }

  const results: DashboardTileSchema[] = [];
  for (const el of elements) {
    const title = String(el?.title || el?.title_text || "").trim();
    if (!title) continue;

    let query: QueryLike | null = el?.query || el?.result_maker?.query || null;
    if (!query && el?.merge_result_id) {
      try {
        const mq = await core40SDK.ok(core40SDK.merge_query(el.merge_result_id));
        const firstSource = Array.isArray(mq?.source_queries) ? mq.source_queries[0] : null;
        query = {
          model: firstSource?.model,
          view: firstSource?.view,
          fields: firstSource?.fields || [],
          dynamic_fields: mq?.dynamic_fields,
          vis_config: mq?.vis_config || el?.result_maker?.vis_config,
        };
      } catch {
        query = null;
      }
    }

    if (!query) continue;
    const fieldMap = await getExploreFieldMap(query.model, query.view);
    const columns = resolveColumnsForQuery(query, fieldMap);
    results.push({
      title,
      slug: sanitizeTileSlug(title),
      columns,
    });
  }

  return results;
}

/**
 * Executes Looker dashboard tile queries as unformatted CSVs (with vis labels applied)
 * so the Extension can preview a populated .pptx deck live in the browser.
 */
export async function fetchDashboardTileCsvs(
  core40SDK: any,
  dashboardId: string,
  requiredSlugs?: Set<string>
): Promise<Record<string, string>> {
  const dash = await core40SDK.ok(core40SDK.dashboard(dashboardId));
  const elements: any[] = Array.isArray(dash?.dashboard_elements) ? dash.dashboard_elements : [];
  const csvMap: Record<string, string> = {};

  for (const el of elements) {
    const title = String(el?.title || el?.title_text || "").trim();
    if (!title) continue;
    const slug = sanitizeTileSlug(title);
    if (requiredSlugs && requiredSlugs.size > 0 && !requiredSlugs.has(slug)) {
      continue;
    }

    const queryObj = el?.query || el?.result_maker?.query || null;
    const queryId = el?.query_id || queryObj?.id;
    if (queryId) {
      const csv = await core40SDK.ok(
        core40SDK.run_query({
          query_id: queryId,
          result_format: "csv",
          apply_formatting: false,
          apply_vis: true,
        })
      );
      csvMap[slug] = String(csv || "");
    } else if (queryObj?.model && queryObj?.view) {
      const csv = await core40SDK.ok(
        core40SDK.run_inline_query({
          result_format: "csv",
          apply_formatting: false,
          apply_vis: true,
          body: {
            model: queryObj.model,
            view: queryObj.view,
            fields: queryObj.fields || [],
            filters: queryObj.filters || {},
            sorts: queryObj.sorts || [],
            limit: String(queryObj.limit || "500"),
            dynamic_fields: queryObj.dynamic_fields,
            vis_config: queryObj.vis_config,
          },
        })
      );
      csvMap[slug] = String(csv || "");
    }
  }

  return csvMap;
}
