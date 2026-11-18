import assert from "node:assert/strict";
import {
  buildEmbedFilterParams,
  buildFanoutArtifactKey,
  buildFanoutSearchPattern,
  buildScheduledPlanWriteBody,
  buildSliceOutputPrefix,
  buildTemplateLinksArtifactKey,
  encodeDashKey,
  extractDashboardModels,
  findMatchingLinkedTemplate,
  formatDatagroupValue,
  isSlicePlanDrifted,
  mergeEffectiveFilters,
  mergeTemplateDashboardLink,
  parseTemplateLinksArtifacts,
  resolveCrontab,
  sanitizeOutputFolder,
  serializeFiltersString,
  type MasterFanoutSpec,
} from "../src/schedulerUtils.ts";

// 1. Artifact key encoding never emits SQL LIKE wildcards (`_` or `%`) or `.` segment separators
const encodedLookml = encodeDashKey("action-pptx::ebr_1%test");
assert.equal(encodedLookml.includes("_"), false);
assert.equal(encodedLookml.includes("%"), false);
assert.equal(encodedLookml.includes("."), false);
assert.notEqual(encodeDashKey("a-b"), encodeDashKey("a+b"));

const key = buildFanoutArtifactKey("21", "embed_demo::brand_overview", "f1");
const dashPattern = buildFanoutSearchPattern("21", "embed_demo::brand_overview");
const userPattern = buildFanoutSearchPattern("21");
assert.equal(key.startsWith(dashPattern.slice(0, -1)), true);
assert.equal(key.startsWith(userPattern.slice(0, -1)), true);
// Prefix boundary: dashboard 155 pattern must not match dashboard 1550
assert.equal(
  buildFanoutArtifactKey("21", "1550", "f1").startsWith(
    buildFanoutSearchPattern("21", "155").slice(0, -1)
  ),
  false
);

// 2. Output folder & slice subpath sanitization rejects traversal and matches backend gcs.py
assert.equal(sanitizeOutputFolder("/Q4 EBR/../Brand Decks.pptx"), "Q4_EBR/Brand_Decks");
assert.equal(buildSliceOutputPrefix("Q4 EBR", "Levi's"), "Q4_EBR/Levi_s");

// 3. Deterministic Looker `filters_string` serialization (slice overrides baseline)
assert.equal(
  serializeFiltersString({ Country: "USA", Brand: "Default" }, { Brand: "Levi's" }),
  "?Brand=Levi%27s&Country=USA"
);
assert.equal(serializeFiltersString({}, {}), "");

// 4. Cron presets + WriteScheduledPlan body + drift detection
const spec: MasterFanoutSpec = {
  id: "f1",
  title: "Weekly Brand Fanout",
  owner_user_id: "21",
  dashboard_id: "embed_demo::brand_overview",
  dashboard_title: "Brand Overview",
  template: "embed_demo_template.pptx",
  theme: "Levis_Dark",
  output_folder: "Q4_EBR",
  cadence: {
    mode: "weekly",
    hour: 8,
    minute: 30,
    dayOfWeek: 1,
    dayOfMonth: 1,
    crontab: "",
    timezone: "America/Los_Angeles",
  },
  enabled: true,
  base_filters: { Country: "USA" },
  slices: [{ id: "s1", label: "Levi's", slug: "Levis", filters: { Brand: "Levi's" }, plan_id: "99" }],
  updated_at: "2026-11-17T00:00:00.000Z",
};

assert.equal(resolveCrontab(spec.cadence), "30 8 * * 1");
const planBody = buildScheduledPlanWriteBody(
  spec,
  spec.slices[0],
  "looker-integration://66::action-pptx"
);
assert.equal(planBody.lookml_dashboard_id, "embed_demo::brand_overview");
assert.equal(planBody.filters_string, "?Brand=Levi%27s&Country=USA");
assert.equal(isSlicePlanDrifted(spec, spec.slices[0], { ...planBody, id: "99" }), false);
assert.equal(
  isSlicePlanDrifted(spec, spec.slices[0], { ...planBody, id: "99", enabled: false }),
  true
);

// 5. Model extraction + Looker `model::datagroup` identifier formatting
assert.deepEqual(
  extractDashboardModels({
    id: "3965",
    dashboard_filters: [{ model: "action_pptx" }],
    dashboard_elements: [
      { query: { model: "action_pptx" } },
      { result_maker: { filterables: [{ model: "embed_demo" }] } },
    ],
  }),
  ["action_pptx", "embed_demo"]
);
assert.deepEqual(extractDashboardModels({ id: "action_pptx::executive_business_review" }), [
  "action_pptx",
]);
assert.equal(
  formatDatagroupValue("action_pptx", "ebr_daily_datagroup"),
  "action_pptx::ebr_daily_datagroup"
);
const dgSpec: MasterFanoutSpec = {
  ...spec,
  cadence: { ...spec.cadence, mode: "datagroup", datagroup: "action_pptx::ebr_daily_datagroup" },
};
const dgBody = buildScheduledPlanWriteBody(
  dgSpec,
  dgSpec.slices[0],
  "looker-integration://66::action-pptx"
);
assert.equal(dgBody.datagroup, "action_pptx::ebr_daily_datagroup");
assert.equal("crontab" in dgBody, false);

// 6. 1-to-N Template <-> Dashboard artifact links + Python parity key + UDD title fallback
assert.equal(
  buildTemplateLinksArtifactKey("simple_template.pptx"),
  "tpl_links.simple-5ftemplate-2epptx"
);
const doc1 = mergeTemplateDashboardLink(
  null,
  "simple_template.pptx",
  { id: "action_pptx::executive_business_review", title: "Executive Business Review" },
  "2026-11-17T00:00:00Z"
);
const doc2 = mergeTemplateDashboardLink(
  doc1,
  "simple_template.pptx",
  { id: "3965", title: "Executive Business Review UDD" },
  "2026-11-17T01:00:00Z"
);
// Re-saving an existing dashboard promotes it without duplicating
const doc3 = mergeTemplateDashboardLink(
  doc2,
  "simple_template.pptx",
  { id: "action_pptx::executive_business_review", title: "Executive Business Review" },
  "2026-11-17T02:00:00Z"
);
assert.deepEqual(
  doc3.dashboards.map((d) => d.id),
  ["action_pptx::executive_business_review", "3965"]
);
const linksMap = parseTemplateLinksArtifacts([
  { key: "tpl_links.simple-5ftemplate-2epptx", value: JSON.stringify(doc1) },
  {
    key: "tpl_links.embed-5fdemo-5ftemplate-2epptx",
    value: JSON.stringify(
      mergeTemplateDashboardLink(null, "embed_demo_template.pptx", {
        id: "embed_demo::brand_overview",
        title: "Brand Overview",
      })
    ),
  },
]);
assert.equal(
  findMatchingLinkedTemplate(
    ["embed_demo_template.pptx", "simple_template.pptx"],
    linksMap,
    "action_pptx::executive_business_review"
  ),
  "simple_template.pptx"
);
// UDD copy with different ID ("3965") but matching title ("Executive Business Review") resolves to simple_template.pptx
assert.equal(
  findMatchingLinkedTemplate(
    ["embed_demo_template.pptx", "simple_template.pptx"],
    linksMap,
    "3965",
    "Executive Business Review"
  ),
  "simple_template.pptx"
);

// 7. Embed SDK filter payloads reset cleared slice filters back to dashboard defaults
assert.deepEqual(mergeEffectiveFilters({ Country: "USA" }, { Region: "APAC" }), {
  Country: "USA",
  Region: "APAC",
});
assert.deepEqual(
  buildEmbedFilterParams(
    { Country: "USA" },
    { Region: "" },
    [
      { name: "Country", defaultValue: "All" },
      { name: "Region", defaultValue: "North America" },
    ]
  ),
  { Country: "USA", Region: "North America" }
);

console.log("scheduler_check: all assertions passed");
