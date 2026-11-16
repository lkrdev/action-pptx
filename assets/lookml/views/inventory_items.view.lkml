view: inventory_items {
  sql_table_name: `bigquery-public-data.thelook_ecommerce.inventory_items` ;;

  dimension: row_num {
    type: number
    hidden: yes
    sql: MOD(CAST(${TABLE}.id AS INT64), 4) ;;
  }

  # 1. KPI Summary
  measure: total_spend {
    label: "Total Spend"
    type: number
    sql: 34855339.456 ;;
  }
  measure: trend_rate {
    label: "Trend Rate"
    type: number
    sql: 0.0842 ;;
  }
  measure: avg_order_value {
    label: "Avg Order Value"
    type: number
    sql: 438.92 ;;
  }

  # 2. Mixed Metrics
  dimension: metric_sort {
    type: number
    hidden: yes
    sql: ${row_num} ;;
  }
  dimension: metric {
    label: "Metric"
    type: string
    sql: CASE ${row_num}
      WHEN 0 THEN 'Customers'
      WHEN 1 THEN 'Orders per Customer'
      WHEN 2 THEN 'Gross Margin'
      ELSE 'Avg Order Value' END ;;
  }
  measure: prior {
    label: "Prior"
    type: max
    sql: CASE ${row_num}
      WHEN 0 THEN 3190
      WHEN 1 THEN 65.1
      WHEN 2 THEN 0.5312
      ELSE 404.80 END ;;
  }
  measure: current {
    label: "Current"
    type: max
    sql: CASE ${row_num}
      WHEN 0 THEN 3378
      WHEN 1 THEN 62.44
      WHEN 2 THEN 0.5568
      ELSE 438.92 END ;;
  }
  measure: delta {
    label: "Delta"
    type: max
    sql: CASE ${row_num}
      WHEN 0 THEN 0.0589
      WHEN 1 THEN -0.0409
      WHEN 2 THEN 0.0256
      ELSE 0.0842 END ;;
  }

  # 3. Spend by Category
  dimension: cat_sort {
    type: number
    hidden: yes
    sql: ${row_num} ;;
  }
  dimension: product_category {
    label: "Product Category"
    type: string
    sql: CASE ${row_num}
      WHEN 0 THEN 'Apparel'
      WHEN 1 THEN 'Footwear'
      WHEN 2 THEN 'Accessories'
      ELSE 'Outerwear' END ;;
  }
  measure: cat_total_spend {
    label: "Total Spend"
    type: max
    sql: CASE ${row_num}
      WHEN 0 THEN 11250000.0
      WHEN 1 THEN 9840000.0
      WHEN 2 THEN 6415339.46
      ELSE 7350000.0 END ;;
  }

  # 4 & 5. Quarterly Trends (Monthly Trend + AOV Trend)
  dimension: quarter {
    label: "Quarter"
    type: string
    sql: CASE ${row_num}
      WHEN 0 THEN '2026-Q1'
      WHEN 1 THEN '2026-Q2'
      WHEN 2 THEN '2026-Q3'
      ELSE '2026-Q4' END ;;
  }
  measure: direct_spend {
    label: "Direct"
    type: max
    sql: CASE ${row_num}
      WHEN 0 THEN 6450000.5
      WHEN 1 THEN 6810000.0
      WHEN 2 THEN 6590000.0
      ELSE 7655339.0 END ;;
  }
  measure: marketplace_spend {
    label: "Marketplace"
    type: max
    sql: CASE ${row_num}
      WHEN 0 THEN 1620000.25
      WHEN 1 THEN 1740500.75
      WHEN 2 THEN NULL
      ELSE 1989499.0 END ;;
  }
  measure: prior_year_aov {
    label: "Prior Year AOV"
    type: max
    sql: CASE ${row_num}
      WHEN 0 THEN 392.10
      WHEN 1 THEN 401.50
      WHEN 2 THEN 408.20
      ELSE 417.40 END ;;
  }
  measure: current_year_aov {
    label: "Current Year AOV"
    type: max
    sql: CASE ${row_num}
      WHEN 0 THEN 418.40
      WHEN 1 THEN 432.15
      WHEN 2 THEN 439.80
      ELSE 465.30 END ;;
  }
  measure: target_benchmark {
    label: "Target Benchmark"
    type: max
    sql: CASE ${row_num}
      WHEN 0 THEN 405.00
      WHEN 1 THEN 410.00
      WHEN 2 THEN 414.50
      ELSE 420.00 END ;;
  }

  # 6. Cost Bridge (5 rows)
  dimension: bridge_sort {
    type: number
    hidden: yes
    sql: MOD(CAST(${TABLE}.id AS INT64), 5) ;;
  }
  dimension: bridge_step {
    label: "Bridge Step"
    type: string
    sql: CASE ${bridge_sort}
      WHEN 0 THEN 'Prior AOV'
      WHEN 1 THEN 'Volume Mix'
      WHEN 2 THEN 'Unit Price'
      WHEN 3 THEN 'Premium Mix'
      ELSE 'Current AOV' END ;;
  }
  measure: bridge_base {
    label: "Base"
    type: max
    sql: CASE ${bridge_sort}
      WHEN 0 THEN 0.0
      WHEN 1 THEN 404.80
      WHEN 2 THEN 419.00
      WHEN 3 THEN 430.60
      ELSE 0.0 END ;;
  }
  measure: bridge_impact {
    label: "Impact"
    type: max
    sql: CASE ${bridge_sort}
      WHEN 0 THEN 404.80
      WHEN 1 THEN 14.20
      WHEN 2 THEN 11.60
      WHEN 3 THEN 8.32
      ELSE 438.92 END ;;
  }

  # 7. Regional Customers
  dimension: region_sort {
    type: number
    hidden: yes
    sql: ${row_num} ;;
  }
  dimension: region {
    label: "Region"
    type: string
    sql: CASE ${row_num}
      WHEN 0 THEN 'North America'
      WHEN 1 THEN 'EMEA'
      WHEN 2 THEN 'APAC'
      ELSE 'LATAM' END ;;
  }
  measure: prior_customers {
    label: "Prior Customers"
    type: max
    sql: CASE ${row_num}
      WHEN 0 THEN 710
      WHEN 1 THEN 820
      WHEN 2 THEN 940
      ELSE 720 END ;;
  }
  measure: current_customers {
    label: "Current Customers"
    type: max
    sql: CASE ${row_num}
      WHEN 0 THEN 745
      WHEN 1 THEN 890
      WHEN 2 THEN 988
      ELSE 755 END ;;
  }

  # 8. Top Brands
  dimension: brand_sort {
    type: number
    hidden: yes
    sql: ${row_num} ;;
  }
  dimension: brand_name {
    label: "Brand Name"
    type: string
    sql: CASE ${row_num}
      WHEN 0 THEN 'Apex Goods'
      WHEN 1 THEN 'Nova Apparel'
      WHEN 2 THEN 'Vertex Active'
      ELSE 'Solstice Co' END ;;
  }
  dimension: brand_category {
    label: "Category"
    type: string
    sql: CASE ${row_num}
      WHEN 0 THEN 'Outerwear'
      WHEN 1 THEN 'Activewear'
      WHEN 2 THEN 'Footwear'
      ELSE 'Accessories' END ;;
  }
  measure: units_sold {
    label: "Units Sold"
    type: max
    sql: CASE ${row_num}
      WHEN 0 THEN 120
      WHEN 1 THEN 340
      WHEN 2 THEN 890
      ELSE 64 END ;;
  }
  measure: net_revenue {
    label: "Net Revenue"
    type: max
    sql: CASE ${row_num}
      WHEN 0 THEN 450200.4
      WHEN 1 THEN 289100.2
      WHEN 2 THEN 195000.0
      ELSE 182400.0 END ;;
  }
  measure: yoy_trend {
    label: "YoY Trend"
    type: max
    sql: CASE ${row_num}
      WHEN 0 THEN 0.124
      WHEN 1 THEN 0.382
      WHEN 2 THEN 0.065
      ELSE 0.091 END ;;
  }

  # 9. Top Accounts (3 rows)
  dimension: account_sort {
    type: number
    hidden: yes
    sql: MOD(CAST(${TABLE}.id AS INT64), 3) ;;
  }
  dimension: account_tier {
    label: "Account Tier"
    type: string
    sql: CASE ${account_sort}
      WHEN 0 THEN 'Account #1 (>$500k)'
      WHEN 1 THEN 'Account #2 ($250k-$500k)'
      ELSE 'Account #3 ($250k-$500k)' END ;;
  }
  dimension: primary_segment {
    label: "Primary Segment"
    type: string
    sql: CASE ${account_sort}
      WHEN 0 THEN 'Enterprise Retail'
      WHEN 1 THEN 'Direct Wholesale'
      ELSE 'Global Partner' END ;;
  }
  measure: orders {
    label: "Orders"
    type: max
    sql: CASE ${account_sort}
      WHEN 0 THEN 142
      WHEN 1 THEN 88
      ELSE 215 END ;;
  }
  measure: total_revenue {
    label: "Total Revenue"
    type: string
    sql: MAX(CASE ${account_sort}
      WHEN 0 THEN '$684,210'
      WHEN 1 THEN '$412,900'
      ELSE '$298,450' END) ;;
  }
}
