connection: "looker-private-demo"

include: "/views/*.view.lkml"
include: "/dashboards/*.dashboard.lookml"

explore: inventory_items {
  label: "Inventory Items"
  view_label: ""
}
