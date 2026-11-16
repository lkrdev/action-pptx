application: action_pptx_template_manager {
  label: "PowerPoint Template Manager (action-pptx)"
  url: "https://cdn.lkr.dev/apps/action-pptx/latest/bundle.js"
  mount_points: {
    standalone: yes
  }
  entitlements: {
    core_api_methods: [
      "me",
      "search_dashboards",
      "dashboard",
      "lookml_model_explore",
      "merge_query",
      "run_query",
      "run_inline_query",
      "all_integration_hubs",
      "all_integrations",
      "all_user_attributes",
      "create_user_attribute",
      "set_user_attribute_group_values"
    ]
    scoped_user_attributes: ["action_hub_secret"]
    external_api_urls: [
      "https://*.a.run.app",
      "https://*.trycloudflare.com",
      "https://oauth2.googleapis.com"
    ]
    oauth2_urls: [
      "https://accounts.google.com/o/oauth2/v2/auth",
      "https://oauth2.googleapis.com/token"
    ]
    local_storage: yes
    use_form_submit: yes
  }
}
