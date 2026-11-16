- dashboard: executive_business_review
  title: Executive Business Review
  layout: newspaper
  preferred_viewer: dashboards-next
  description: "End-to-end 9-tile showcase dashboard for action-pptx (Executive Business Review)"
  elements:
  - title: KPI Summary
    name: KPI Summary
    model: action_pptx
    explore: inventory_items
    type: looker_grid
    fields: [inventory_items.total_spend, inventory_items.trend_rate, inventory_items.avg_order_value]
    sorts: [inventory_items.total_spend desc]
    limit: 1
    show_view_names: false
    row: 0
    col: 0
    width: 8
    height: 4

  - title: Mixed Metrics
    name: Mixed Metrics
    model: action_pptx
    explore: inventory_items
    type: looker_grid
    fields: [inventory_items.metric_sort, inventory_items.metric, inventory_items.prior, inventory_items.current, inventory_items.delta]
    sorts: [inventory_items.metric_sort asc]
    limit: 10
    show_view_names: false
    row: 0
    col: 8
    width: 8
    height: 4

  - title: Spend by Category
    name: Spend by Category
    model: action_pptx
    explore: inventory_items
    type: looker_pie
    fields: [inventory_items.cat_sort, inventory_items.product_category, inventory_items.cat_total_spend]
    sorts: [inventory_items.cat_sort asc]
    limit: 10
    show_view_names: false
    row: 0
    col: 16
    width: 8
    height: 4

  - title: Monthly Trend
    name: Monthly Trend
    model: action_pptx
    explore: inventory_items
    type: looker_column
    fields: [inventory_items.quarter, inventory_items.direct_spend, inventory_items.marketplace_spend]
    sorts: [inventory_items.quarter asc]
    limit: 10
    show_view_names: false
    row: 4
    col: 0
    width: 8
    height: 6

  - title: AOV Trend
    name: AOV Trend
    model: action_pptx
    explore: inventory_items
    type: looker_line
    fields: [inventory_items.quarter, inventory_items.prior_year_aov, inventory_items.current_year_aov, inventory_items.target_benchmark]
    sorts: [inventory_items.quarter asc]
    limit: 10
    show_view_names: false
    row: 4
    col: 8
    width: 8
    height: 6

  - title: Cost Bridge
    name: Cost Bridge
    model: action_pptx
    explore: inventory_items
    type: looker_column
    fields: [inventory_items.bridge_sort, inventory_items.bridge_step, inventory_items.bridge_base, inventory_items.bridge_impact]
    sorts: [inventory_items.bridge_sort asc]
    limit: 10
    show_view_names: false
    row: 4
    col: 16
    width: 8
    height: 6

  - title: Regional Customers
    name: Regional Customers
    model: action_pptx
    explore: inventory_items
    type: looker_bar
    fields: [inventory_items.region_sort, inventory_items.region, inventory_items.prior_customers, inventory_items.current_customers]
    sorts: [inventory_items.region_sort asc]
    limit: 10
    show_view_names: false
    row: 10
    col: 0
    width: 8
    height: 6

  - title: Top Brands
    name: Top Brands
    model: action_pptx
    explore: inventory_items
    type: looker_grid
    fields: [inventory_items.brand_sort, inventory_items.brand_name, inventory_items.brand_category, inventory_items.units_sold, inventory_items.net_revenue, inventory_items.yoy_trend]
    sorts: [inventory_items.brand_sort asc]
    limit: 10
    show_view_names: false
    row: 10
    col: 8
    width: 8
    height: 6

  - title: Top Accounts
    name: Top Accounts
    model: action_pptx
    explore: inventory_items
    type: looker_grid
    fields: [inventory_items.account_sort, inventory_items.account_tier, inventory_items.primary_segment, inventory_items.orders, inventory_items.total_revenue]
    sorts: [inventory_items.account_sort asc]
    limit: 10
    show_view_names: false
    row: 10
    col: 16
    width: 8
    height: 6
