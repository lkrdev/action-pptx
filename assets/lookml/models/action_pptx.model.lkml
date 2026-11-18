connection: "looker-private-demo"

include: "/views/*.view.lkml"
include: "/dashboards/*.dashboard.lookml"

datagroup: ebr_daily_datagroup {
  sql_trigger: SELECT CURRENT_DATE() ;;
  max_cache_age: "24 hours"
}

persist_with: ebr_daily_datagroup

explore: inventory_items {
  label: "Inventory Items"
}
