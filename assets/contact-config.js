window.CONTACT_FORM = {
  supabaseUrl: "https://ecbhaeiwnygplawlbnio.supabase.co",
  supabaseAnonKey: "sb_publishable_LDBZuLzFMUo0xsugiKRpZA_qKeqeXeW",
  table: "public_visit_requests",
};

window.WEBSITE_CONTENT = {
  supabaseUrl: window.CONTACT_FORM.supabaseUrl,
  supabaseAnonKey: window.CONTACT_FORM.supabaseAnonKey,
  settingsTables: ["website_settings", "site_settings", "company_profile"],
  categoriesTables: ["website_categories", "categories"],
  galleryItemsTables: ["website_gallery_items", "gallery_items", "portfolio_items"],
};

window.MEDIA_STUDIO = {
  apiBaseUrl: "",
  productionApiBaseUrl: "",
  localApiBaseUrl: "http://127.0.0.1:4000",
  authStorageKey: "nolimit_gallery_control_admin_token",
  apiUrlStorageKey: "nolimit_gallery_control_api_url",
};
