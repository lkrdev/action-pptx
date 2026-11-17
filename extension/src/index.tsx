import React from "react";
import { createRoot } from "react-dom/client";
import { connectExtensionHost, LookerExtensionSDK } from "@looker/extension-sdk";
import { App } from "./App";

interface MockTileDef {
  title: string;
  id: string;
  fields: string[];
  csv: string;
}

const EBR_TILE_DEFS: MockTileDef[] = [
  {
    title: "KPI Summary",
    id: "q_kpi",
    fields: ["Total Spend", "Trend Rate", "Avg Order Value"],
    csv: "Total Spend,Trend Rate,Avg Order Value\n34855339.456,0.0842,438.92\n",
  },
  {
    title: "Mixed Metrics",
    id: "q_mixed",
    fields: ["Metric", "Prior", "Current", "Delta"],
    csv:
      "Metric,Prior,Current,Delta\n" +
      "Customers,3190,3378,0.0589\n" +
      "Orders per Customer,65.1,62.44,-0.0409\n" +
      "Gross Margin,0.5312,0.5568,0.0256\n" +
      "Avg Order Value,404.80,438.92,0.0842\n",
  },
  {
    title: "Spend by Category",
    id: "q_cat",
    fields: ["Product Category", "Total Spend"],
    csv:
      "Product Category,Total Spend\n" +
      "Apparel,11250000.0\n" +
      "Footwear,9840000.0\n" +
      "Accessories,6415339.46\n" +
      "Outerwear,7350000.0\n",
  },
  {
    title: "Monthly Trend",
    id: "q_monthly",
    fields: ["Quarter", "Direct", "Marketplace"],
    csv:
      "Quarter,Direct,Marketplace\n" +
      "2026-Q1,6450000.5,1620000.25\n" +
      "2026-Q2,6810000.0,1740500.75\n" +
      "2026-Q3,6590000.0,\n" +
      "2026-Q4,7655339.0,1989499.0\n",
  },
  {
    title: "AOV Trend",
    id: "q_aov",
    fields: ["Quarter", "Prior Year AOV", "Current Year AOV", "Target Benchmark"],
    csv:
      "Quarter,Prior Year AOV,Current Year AOV,Target Benchmark\n" +
      "2026-Q1,392.10,418.40,405.00\n" +
      "2026-Q2,401.50,432.15,410.00\n" +
      "2026-Q3,408.20,439.80,414.50\n" +
      "2026-Q4,417.40,465.30,420.00\n",
  },
  {
    title: "Cost Bridge",
    id: "q_bridge",
    fields: ["Bridge Step", "Base", "Impact"],
    csv:
      "Bridge Step,Base,Impact\n" +
      "Prior AOV,0.0,404.80\n" +
      "Volume Mix,404.80,14.20\n" +
      "Unit Price,419.00,11.60\n" +
      "Premium Mix,430.60,8.32\n" +
      "Current AOV,0.0,438.92\n",
  },
  {
    title: "Regional Customers",
    id: "q_region",
    fields: ["Region", "Prior Customers", "Current Customers"],
    csv:
      "Region,Prior Customers,Current Customers\n" +
      "North America,710,745\n" +
      "EMEA,820,890\n" +
      "APAC,940,988\n" +
      "LATAM,720,755\n",
  },
  {
    title: "Top Brands",
    id: "q_brands",
    fields: ["Brand Name", "Category", "Units Sold", "Net Revenue", "YoY Trend"],
    csv:
      "Brand Name,Category,Units Sold,Net Revenue,YoY Trend\n" +
      "Apex Goods,Outerwear,120,450200.4,0.124\n" +
      "Nova Apparel,Activewear,340,289100.2,0.382\n" +
      "Vertex Active,Footwear,890,195000.0,0.065\n" +
      "Solstice Co,Accessories,64,182400.0,0.091\n",
  },
  {
    title: "Top Accounts",
    id: "q_accounts",
    fields: ["Account Tier", "Primary Segment", "Orders", "Total Revenue"],
    csv:
      "Account Tier,Primary Segment,Orders,Total Revenue\n" +
      'Account #1 (>$500k),Enterprise Retail,142,"$684,210"\n' +
      'Account #2 ($250k-$500k),Direct Wholesale,88,"$412,900"\n' +
      'Account #3 ($250k-$500k),Global Partner,215,"$298,450"\n',
  },
];

const EMBED_DEMO_TILE_DEFS: MockTileDef[] = [
  {
    title: "Revenue",
    id: "ed_rev",
    fields: ["Current Period", "P/P Change"],
    csv: "Current Period,P/P Change\n1542800.50,0.142\n",
  },
  {
    title: "Orders",
    id: "ed_ord",
    fields: ["Current Period", "P/P Change"],
    csv: "Current Period,P/P Change\n16920,0.098\n",
  },
  {
    title: "Average Order Value",
    id: "ed_aov",
    fields: ["Current AOV", "P/P Change"],
    csv: "Current AOV,P/P Change\n91.18,0.041\n",
  },
  {
    title: "Units Sold",
    id: "ed_units",
    fields: ["Current Period", "P/P Change"],
    csv: "Current Period,P/P Change\n20040,0.116\n",
  },
  {
    title: "Top Categories by Revenue & Volume",
    id: "ed_top_cat",
    fields: ["Category", "Total Sale Price", "Purchased Items"],
    csv:
      "Category,Total Sale Price,Purchased Items\n" +
      "Jeans,482500.0,5420\n" +
      "Outerwear & Coats,391200.0,3180\n" +
      "Sweaters,284900.0,3890\n" +
      "Tops & Tees,219400.0,4610\n" +
      "Accessories,164800.0,2940\n",
  },
  {
    title: "Revenue Share by Country",
    id: "ed_country",
    fields: ["Country", "Total Sale Price"],
    csv:
      "Country,Total Sale Price\n" +
      "United States,742900.0\n" +
      "United Kingdom,318400.0\n" +
      "China,268100.0\n" +
      "Brasil,213400.0\n",
  },
  {
    title: "Total Active Customers",
    id: "ed_cust",
    fields: ["Current Period", "P/P Change"],
    csv: "Current Period,P/P Change\n11880,0.128\n",
  },
  {
    title: "Customer Average Spend",
    id: "ed_spend",
    fields: ["Current Spend", "P/P Change"],
    csv: "Current Spend,P/P Change\n129.87,0.064\n",
  },
  {
    title: "Orders per Customer",
    id: "ed_opc",
    fields: ["Current Period", "P/P Change"],
    csv: "Current Period,P/P Change\n1.42,0.032\n",
  },
  {
    title: "New Customer Acquisitions",
    id: "ed_new_cust",
    fields: ["Current Period", "P/P Change"],
    csv: "Current Period,P/P Change\n3490,0.185\n",
  },
  {
    title: "Monthly Revenue & AOV Trajectory",
    id: "ed_traj",
    fields: ["Created Month", "Total Sale Price", "Average Sale Price"],
    csv:
      "Created Month,Total Sale Price,Average Sale Price\n" +
      "2026-07,342000.0,86.4\n" +
      "2026-08,368500.0,88.9\n" +
      "2026-09,394200.0,91.2\n" +
      "2026-10,438100.0,94.8\n",
  },
  {
    title: "Customer Traffic Source Composition",
    id: "ed_traffic",
    fields: ["Traffic Source", "Users Count"],
    csv:
      "Traffic Source,Users Count\n" +
      "Search,4820\n" +
      "Organic,2910\n" +
      "Email,1840\n" +
      "Social,1420\n" +
      "Display,890\n",
  },
  {
    title: "Top Geographic Footprint by Location",
    id: "ed_geo",
    fields: ["State", "Users Count"],
    csv:
      "State,Users Count\n" +
      "California,2410\n" +
      "Texas,1680\n" +
      "New York,1490\n" +
      "England,1320\n" +
      "Florida,1140\n",
  },
  {
    title: "Shipped Orders",
    id: "ed_shipped",
    fields: ["Current Period", "P/P Change"],
    csv: "Current Period,P/P Change\n14310,0.114\n",
  },
  {
    title: "Processing Orders",
    id: "ed_proc",
    fields: ["Current Period", "P/P Change"],
    csv: "Current Period,P/P Change\n2140,-0.024\n",
  },
  {
    title: "Returned Orders",
    id: "ed_ret",
    fields: ["Current Period", "P/P Change"],
    csv: "Current Period,P/P Change\n1180,-0.038\n",
  },
  {
    title: "Cancelled Orders",
    id: "ed_canc",
    fields: ["Current Period", "P/P Change"],
    csv: "Current Period,P/P Change\n790,-0.051\n",
  },
  {
    title: "Fulfillment Pipeline Status Breakdown",
    id: "ed_pipe",
    fields: ["Status", "Order Items Count"],
    csv:
      "Status,Order Items Count\n" +
      "Complete,8420\n" +
      "Shipped,5890\n" +
      "Processing,2140\n" +
      "Returned,1180\n" +
      "Cancelled,790\n",
  },
  {
    title: "Return & Cancellation Hotspots by Category",
    id: "ed_churn",
    fields: ["Category", "Count Returned", "Count Cancelled"],
    csv:
      "Category,Count Returned,Count Cancelled\n" +
      "Jeans,380,210\n" +
      "Outerwear & Coats,290,175\n" +
      "Sweaters,220,160\n" +
      "Tops & Tees,175,140\n" +
      "Accessories,115,105\n",
  },
];

const ALL_MOCK_TILES: MockTileDef[] = [...EMBED_DEMO_TILE_DEFS, ...EBR_TILE_DEFS];

function createMockHarness(apiBase: string, secret: string, unconfigured: boolean) {
  let ctxData: Record<string, any> = unconfigured
    ? {}
    : {
        cloudRunUrl: apiBase,
        templatesGcsUri: "gs://fifth-inkwell-452619-v2-action-pptx/templates",
        outputGcsUri: "gs://fifth-inkwell-452619-v2-action-pptx/output",
        useOAuth: true,
        oauthClientId: "104928374651-shared-action-pptx.apps.googleusercontent.com",
        oauthGcsPassthrough: true,
      };

  const extensionSDK = {
    lookerHostData: { hostUrl: "https://looker.example.com" },
    getContextData: () => ctxData,
    saveContextData: async (next: Record<string, any>) => {
      ctxData = next;
    },
    clientRouteChanged: (nextRoute: string) => {
      const target = nextRoute.startsWith("/") ? nextRoute : `/${nextRoute}`;
      const qIdx = target.indexOf("?");
      const pathPart = qIdx >= 0 ? target.slice(0, qIdx) : target;
      const merged = new URLSearchParams(window.location.search);
      for (const k of ["filename", "mode", "q", "user_id"]) {
        merged.delete(k);
      }
      if (qIdx >= 0) {
        for (const [k, v] of new URLSearchParams(target.slice(qIdx + 1))) merged.set(k, v);
      }
      const qs = merged.toString();
      window.history.replaceState(null, "", qs ? `${pathPart}?${qs}` : pathPart);
    },
    localStorageGetItem: async (k: string) => window.localStorage.getItem(k),
    localStorageSetItem: async (k: string, v: string) => window.localStorage.setItem(k, v),
    localStorageRemoveItem: async (k: string) => window.localStorage.removeItem(k),
    oauth2Authenticate: async () => ({
      access_token: "mock-shared-oauth-token",
      expires_in: 3600,
    }),
    createSecretKeyTag: (key: string) => key,
    openBrowserWindow: (url: string, target?: string) => window.open(url, target || "_blank"),
    userAttributeGetItem: async (attr: string) =>
      attr === "brand" ? "Levi's" : "gs://fifth-inkwell-452619-v2-action-pptx/templates",
    serverProxy: async (url: string, init: any) => {
      const resp = await fetch(url, {
        method: init.method || "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Token token="${secret}"`,
        },
        body: init.body,
      });
      return { status: resp.status, body: await resp.json() };
    },
  };

  const artifactsStore = new Map<string, Record<string, any>>();
  const plansStore = new Map<string, Record<string, any>>();
  let nextPlanId = 500;

  const sqlLikeToRegex = (pattern: string): RegExp => {
    const escaped = pattern.replace(/[.+^${}()|[\]\\]/g, "\\$&");
    return new RegExp("^" + escaped.replace(/%/g, ".*").replace(/_/g, ".") + "$");
  };

  const core40SDK = {
    ok: async (val: any) => await val,
    me: async () => ({ id: "21", personal_folder_id: "43", display_name: "Demo User" }),
    all_dashboards: async () => [
      {
        id: "embed_demo::brand_overview",
        title: "Brand Overview",
        model: { id: "embed_demo", label: "Embed Demo" },
        folder: { id: "lookml", name: "LookML Dashboards" },
        description: "Multi-brand executive pulse across sales, customers, and fulfillment",
      },
      {
        id: "action_pptx_extension::executive_business_review",
        title: "Executive Business Review",
        model: { id: "action_pptx_extension", label: "Action PPTX" },
        folder: { id: "lookml", name: "LookML Dashboards" },
        description: "9-tile quarterly business review showcase",
      },
      {
        id: "155",
        title: "Brand Overview",
        folder: { id: "12542", name: "Embed Demo", is_personal: false },
      },
      {
        id: "170",
        title: "Q4 Merchandising Scorecard",
        folder: { id: "43", name: "Personal Folder", is_personal: true },
      },
    ],
    search_dashboards: async () => [
      {
        id: "170",
        title: "Q4 Merchandising Scorecard",
        user_id: "21",
        last_viewed_at: "2026-11-15 18:20:00+00:00",
        view_count: 19,
        folder: { id: "43", name: "Personal Folder", is_personal: true },
      },
      {
        id: "155",
        title: "Brand Overview",
        user_id: "21",
        last_viewed_at: "2026-11-14 09:10:00+00:00",
        view_count: 42,
        folder: { id: "12542", name: "Embed Demo", is_personal: false },
      },
    ],
    dashboard: async (id: string) => {
      const isEmbedDemo =
        id === "155" || id.includes("embed_demo") || id.includes("brand_overview");
      const defs = isEmbedDemo ? EMBED_DEMO_TILE_DEFS : EBR_TILE_DEFS;
      return {
        id,
        title: isEmbedDemo
          ? "Brand Overview (Looker Embed Demo)"
          : "Executive Business Review",
        dashboard_filters: isEmbedDemo
          ? [
              {
                name: "Brand",
                title: "Brand",
                model: "embed_demo",
                explore: "order_items",
                dimension: "products.brand",
                default_value: "",
              },
              {
                name: "Country",
                title: "Country",
                model: "embed_demo",
                explore: "order_items",
                dimension: "users.country",
                default_value: "USA",
              },
            ]
          : [
              {
                name: "Region",
                title: "Region",
                model: "action_pptx",
                explore: "inventory_items",
                dimension: "inventory_items.region",
                default_value: "",
              },
            ],
        dashboard_elements: defs.map((t) => ({
          title: t.title,
          query_id: t.id,
          query: {
            id: t.id,
            model: isEmbedDemo ? "embed_demo" : "action_pptx",
            view: isEmbedDemo ? "order_items" : "inventory_items",
            fields: t.fields,
            vis_config: { show_view_names: false },
          },
        })),
      };
    },
    lookml_model_explore: async () => ({
      fields: {
        dimensions: [
          { name: "products.brand", label: "Products Brand", label_short: "Brand" },
          { name: "users.country", label: "Users Country", label_short: "Country" },
          { name: "inventory_items.region", label: "Region", label_short: "Region" },
          ...ALL_MOCK_TILES.flatMap((t) =>
            t.fields.map((f) => ({ name: f, label: f, label_short: f }))
          ),
        ],
        measures: [
          {
            name: "order_items.total_sale_price",
            label: "Order Items Total Sale Price",
            label_short: "Total Sale Price",
          },
          { name: "order_items.count", label: "Order Items Count", label_short: "Orders" },
        ],
        filters: [],
        parameters: [],
      },
    }),
    run_query: async ({ query_id }: { query_id: string }) => {
      const found = ALL_MOCK_TILES.find((t) => t.id === query_id);
      return found ? found.csv : "";
    },
    run_inline_query: async ({ body }: { body: Record<string, any> }) => {
      const dim = String((body?.fields || [])[0] || "products.brand");
      const sampleValues = dim.includes("region")
        ? [
            { [dim]: "North America", "order_items.total_sale_price": 745000 },
            { [dim]: "EMEA", "order_items.total_sale_price": 890000 },
            { [dim]: "APAC", "order_items.total_sale_price": 988000 },
            { [dim]: "LATAM", "order_items.total_sale_price": 412000 },
          ]
        : [
            { [dim]: "Levi's", "order_items.total_sale_price": 1542800 },
            { [dim]: "Calvin Klein", "order_items.total_sale_price": 984200 },
            { [dim]: "Carhartt", "order_items.total_sale_price": 842100 },
            { [dim]: "Tommy Hilfiger", "order_items.total_sale_price": 712400 },
            { [dim]: "Columbia", "order_items.total_sale_price": 654900 },
          ];
      const filterExpr = String(body?.filters?.[dim] || "").replace(/%/g, "").trim().toLowerCase();
      const filtered = filterExpr
        ? sampleValues.filter((r) => String(r[dim]).toLowerCase().includes(filterExpr))
        : sampleValues.slice();
      const sortSpec = String((body?.sorts || [])[0] || `${dim} asc`);
      const [sortField, dirRaw] = sortSpec.split(/\s+/);
      const desc = dirRaw?.toLowerCase() === "desc";
      filtered.sort((a: any, b: any) => {
        const av = a[sortField] ?? a[dim];
        const bv = b[sortField] ?? b[dim];
        const cmp =
          typeof av === "number" && typeof bv === "number"
            ? av - bv
            : String(av).localeCompare(String(bv));
        return desc ? -cmp : cmp;
      });
      return filtered.slice(0, Number(body?.limit) || 500);
    },
    all_integrations: async () => [
      {
        id: "28::action-pptx",
        enabled: true,
        label: "PowerPoint Template Merge (action-pptx)",
        url: `${apiBase}/actions/pptx/execute`,
        uses_oauth: true,
        oauth_client_id: "104928374651-shared-action-pptx.apps.googleusercontent.com",
        params: [
          {
            name: "templates_gcs_uri",
            value: "gs://fifth-inkwell-452619-v2-action-pptx/templates",
          },
          {
            name: "output_gcs_uri",
            value: "gs://fifth-inkwell-452619-v2-action-pptx/output",
          },
        ],
      },
    ],
    artifact: async ({ namespace, key }: { namespace: string; key: string }) => {
      const hit = artifactsStore.get(`${namespace}::${key}`);
      return hit ? [hit] : [];
    },
    search_artifacts: async ({ namespace, key }: { namespace: string; key?: string }) => {
      const rx = key ? sqlLikeToRegex(key) : null;
      const out: Record<string, any>[] = [];
      for (const item of artifactsStore.values()) {
        if (item.namespace !== namespace) continue;
        if (rx && !rx.test(item.key)) continue;
        out.push(item);
      }
      return out;
    },
    update_artifacts: async (namespace: string, body: Array<Record<string, any>>) => {
      const saved: Record<string, any>[] = [];
      for (const item of body || []) {
        const k = `${namespace}::${item.key}`;
        const prev = artifactsStore.get(k);
        if (prev && item.version !== prev.version) {
          throw new Error(`Artifact version conflict for ${item.key}`);
        }
        const next = {
          namespace,
          key: item.key,
          value: item.value,
          content_type: item.content_type || "application/json",
          version: prev ? prev.version + 1 : 1,
          updated_at: new Date().toISOString(),
        };
        artifactsStore.set(k, next);
        saved.push(next);
      }
      return saved;
    },
    delete_artifact: async (namespace: string, key: string) => {
      artifactsStore.delete(`${namespace}::${key}`);
    },
    search_scheduled_plans: async (req: Record<string, any>) => {
      const out: Record<string, any>[] = [];
      for (const p of plansStore.values()) {
        if (req?.user_id && String(p.user_id) !== String(req.user_id)) continue;
        if (!req?.all_states && p.enabled === false) continue;
        if (req?.dashboard_id && String(p.dashboard_id || "") !== String(req.dashboard_id)) {
          continue;
        }
        if (
          req?.lookml_dashboard_id &&
          String(p.lookml_dashboard_id || "") !== String(req.lookml_dashboard_id)
        ) {
          continue;
        }
        if (req?.destination_type) {
          const dests = Array.isArray(p.scheduled_plan_destination)
            ? p.scheduled_plan_destination
            : [];
          if (!dests.some((d: any) => d?.type === req.destination_type)) continue;
        }
        out.push(p);
      }
      return out;
    },
    create_scheduled_plan: async (body: Record<string, any>) => {
      const id = String(++nextPlanId);
      const created = {
        ...body,
        id,
        user_id: String(body.user_id || "21"),
        next_run_at: new Date(Date.now() + 86400_000).toISOString(),
      };
      plansStore.set(id, created);
      return created;
    },
    update_scheduled_plan: async (id: string, body: Record<string, any>) => {
      const prev = plansStore.get(String(id));
      if (!prev) throw new Error(`ScheduledPlan ${id} not found`);
      const next = { ...prev, ...body, id: String(id) };
      plansStore.set(String(id), next);
      return next;
    },
    delete_scheduled_plan: async (id: string) => {
      plansStore.delete(String(id));
      return "";
    },
    scheduled_plan_run_once_by_id: async (id: string) => {
      const prev = plansStore.get(String(id));
      if (!prev) throw new Error(`ScheduledPlan ${id} not found`);
      const next = { ...prev, last_run_at: new Date().toISOString() };
      plansStore.set(String(id), next);
      return next;
    },
    all_user_attributes: async () => [],
    delete_user_attribute: async () => "",
    create_user_attribute: async () => ({}),
    all_themes: async () => [],
    all_color_collections: async () => [],
  };

  return { extensionSDK, core40SDK };
}

async function mount() {
  const container = document.createElement("div");
  document.body.style.margin = "0";
  document.body.appendChild(container);
  const root = createRoot(container);

  const params = new URLSearchParams(window.location.search);
  if (params.get("mock") === "1" || window === window.parent) {
    const apiBase = params.get("api") || "http://localhost:9119";
    const secret = params.get("secret") || "";
    const unconfigured = params.get("unconfigured") === "1";
    const { extensionSDK, core40SDK } = createMockHarness(apiBase, secret, unconfigured);
    root.render(<App core40SDK={core40SDK} extensionSDK={extensionSDK} />);
    return;
  }

  let routeListener: ((route: string) => void) | null = null;
  const onHostRouteChange = (fn: (route: string) => void) => {
    routeListener = fn;
    return () => {
      if (routeListener === fn) routeListener = null;
    };
  };
  let initialRoute = window.location.pathname + window.location.search;
  const extensionSDK = await connectExtensionHost({
    chattyTimeout: -1,
    setInitialRoute: (route: string) => {
      initialRoute = route || "/";
    },
    hostChangedRoute: (route: string) => {
      routeListener?.(route || "/");
    },
  });
  const core40SDK = LookerExtensionSDK.createClient(extensionSDK);
  root.render(
    <App
      core40SDK={core40SDK}
      extensionSDK={extensionSDK}
      initialRoute={initialRoute}
      onHostRouteChange={onHostRouteChange}
    />
  );
}

void mount();
