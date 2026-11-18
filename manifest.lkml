application: action_pptx_template_manager {
  label: "PowerPoint Template Manager (action-pptx)"
  url: "https://cdn.lkr.dev/apps/action-pptx/latest/bundle.js"
  mount_points: {
    standalone: yes
  }
  entitlements: {
    core_api_methods: [
      "me",
      "all_dashboards",
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
      "delete_user_attribute",
      "all_themes",
      "all_color_collections",
      "all_datagroups",
      "search_scheduled_plans",
      "create_scheduled_plan",
      "update_scheduled_plan",
      "delete_scheduled_plan",
      "scheduled_plan_run_once_by_id",
      "artifact",
      "search_artifacts",
      "update_artifacts",
      "delete_artifact"
    ]
    global_user_attributes: ["brand"]
    external_api_urls: [
      # Replace with your deployed Cloud Run URL:
      "https://*.a.run.app"
    ]
    # only required for OAuth2 use in the Action and Extension
    oauth2_urls: [
      "https://accounts.google.com/o/oauth2/v2/auth",
      "https://oauth2.googleapis.com/token"
    ]
    local_storage: yes
    new_window: yes
    use_downloads: yes
    use_embeds: yes
    use_form_submit: yes
  }
}
