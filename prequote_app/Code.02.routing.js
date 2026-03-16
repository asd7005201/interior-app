// ============================================================
//  ROUTING — doGet
// ============================================================

function doGet(e) {
  ensureSpreadsheetId_();
  // Support both ?page=admin and /admin path formats
  var p = String((e && e.parameter && e.parameter.page) || "").toLowerCase().trim();
  if (!p && e && e.pathInfo) {
    p = String(e.pathInfo).toLowerCase().replace(/^\/+/, "").replace(/\/+$/, "").trim();
  }
  try {
    if (!p || p === "survey" || p === "survey_embed") return renderSurveyPage_(e, p || "survey");
    // Admin: keep default X-Frame-Options policy.
    if (p === "admin") return renderAdminPage_(e);
    // Result: keep default X-Frame-Options policy.
    if (p === "result") return renderResultPage_(e);
    // Build info
    if (p === "buildinfo") return jsonResponse_({ version: APP_VERSION_, build: BUILD_TAG_ });
    // Fallback to public survey only.
    return renderSurveyPage_(e, "survey");
  } catch (err) {
    return HtmlService.createHtmlOutput(
      "<div style='font-family:sans-serif;padding:20px'><b>Error</b><br>" + escapeHtml_(String(err)) + "</div>"
    ).setTitle("Error");
  }
}

function renderSurveyPage_(e, pageKey) {
  var tpl = HtmlService.createTemplateFromFile("survey");
  tpl.app_url = getAppUrl_();
  tpl.settings_json = JSON.stringify(getPublicSettings_());
  tpl.bootstrap_config_json = JSON.stringify(getSurveyConfigBootstrapPayload_() || null);
  tpl.is_embed = isSurveyEmbedRequest_(e, pageKey);
  return tpl.evaluate()
    .setTitle("가견적 시작하기")
    // Only the public survey page may be embedded on external sites.
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no");
}

function renderAdminPage_(e) {
  var cred = String((e && e.parameter && e.parameter.token) || "").trim();
  var tpl = HtmlService.createTemplateFromFile("admin");
  tpl.app_url = getAppUrl_();
  tpl.prefill_token = cred;
  tpl.settings_json = JSON.stringify(getPublicSettings_());
  tpl.ui_content_json = JSON.stringify(getUiContentRows_());
  return tpl.evaluate()
    .setTitle("가견적 관리자")
    // Admin must not be iframe-embedded.
    .addMetaTag("viewport", "width=device-width,initial-scale=1");
}

function renderResultPage_(e) {
  var requestId = String((e && e.parameter && e.parameter.id) || "").trim();
  var shareToken = String((e && e.parameter && e.parameter.token) || "").trim();
  var tpl = HtmlService.createTemplateFromFile("result");
  tpl.app_url = getAppUrl_();
  tpl.request_id = requestId;
  tpl.share_token = shareToken;
  tpl.settings_json = JSON.stringify(getPublicSettings_());
  tpl.ui_content_json = JSON.stringify(getUiContentRows_());
  tpl.bootstrap_result_json = JSON.stringify(buildResultBootstrapForPage_(requestId, shareToken) || null);
  return tpl.evaluate()
    .setTitle("가견적 결과")
    // Result pages are served by direct URL access, not iframe embedding.
    .addMetaTag("viewport", "width=device-width,initial-scale=1");
}

function isSurveyEmbedRequest_(e, pageKey) {
  var route = String(pageKey || "").trim().toLowerCase();
  if (route === "survey_embed") return true;
  var raw = String((e && e.parameter && e.parameter.embed) || "").trim().toLowerCase();
  return raw === "1" || raw === "y" || raw === "yes" || raw === "true";
}

function redirectHtml_(url) {
  var target = String(url || "").trim();
  if (!target) throw new Error("Redirect URL is missing");
  var safeUrl = safeJsonForInlineScript_(target);
  var html = [
    "<!DOCTYPE html>",
    "<html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>",
    "<meta http-equiv='refresh' content='0; url=" + escapeHtml_(target) + "'>",
    "</head><body>",
    "<script>",
    "(function(){",
    "var target=" + safeUrl + ";",
    "try { if (window.top && window.top.location) window.top.location.replace(target); else window.location.replace(target); } catch (e) { window.location.href = target; }",
    "})();",
    "</script>",
    "<div style='font-family:sans-serif;padding:20px'>외부 공개 설문 페이지로 이동합니다. <a href='" + escapeHtml_(target) + "'>계속</a></div>",
    "</body></html>"
  ].join("");
  return HtmlService.createHtmlOutput(html).setTitle("Redirect");
}

function buildSurveySubmitBridgeResponse_(result, origin) {
  var targetOrigin = sanitizePostMessageOrigin_(origin);
  var payloadJson = safeJsonForInlineScript_(result || {});
  var html = [
    "<!DOCTYPE html>",
    "<html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'></head><body>",
    "<script>",
    "(function(){",
    "var message=" + payloadJson + ";",
    "var targetOrigin=" + safeJsonForInlineScript_(targetOrigin) + ";",
    "try {",
    "  if (window.parent && window.parent !== window && window.parent.postMessage) window.parent.postMessage(message, targetOrigin);",
    "  if (window.top && window.top !== window.parent && window.top.postMessage) window.top.postMessage(message, targetOrigin);",
    "} catch (e) {}",
    "if (message && message.ok && message.result_url) {",
    "  document.body.innerHTML='<div style=\"font-family:sans-serif;padding:16px\"><a href=\"'+message.result_url+'\" target=\"_top\">결과 확인하기</a></div>';",
    "} else if (message && message.error) {",
    "  document.body.innerHTML='<div style=\"font-family:sans-serif;padding:16px;color:#b82f24\">'+String(message.error).replace(/</g,'&lt;')+'</div>';",
    "}",
    "})();",
    "</script></body></html>"
  ].join("");

  return HtmlService.createHtmlOutput(html)
    // Public survey submission bridge must load inside an external iframe.
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAppUrl_() {
  try {
    var s = getSettings_();
    if (s.base_url) return normalizeUrl_(s.base_url);
    return normalizeUrl_(ScriptApp.getService().getUrl());
  } catch (e) {
    return "";
  }
}

function normalizeUrl_(url) {
  var u = String(url || "").trim();
  if (u && u.charAt(u.length - 1) === "/") u = u.slice(0, -1);
  return u;
}

function sha256Hex_(value) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value || ""), Utilities.Charset.UTF_8);
  var out = [];
  for (var i = 0; i < bytes.length; i++) {
    var v = (bytes[i] + 256) % 256;
    out.push((v < 16 ? "0" : "") + v.toString(16));
  }
  return out.join("");
}

function formatMoneyAmount_(value) {
  var amount = Number(value || 0);
  if (!isFinite(amount)) amount = 0;
  return Math.round(amount / 10000).toLocaleString("ko-KR") + "만원";
}

function formatMoneyRange(minValue, maxValue) {
  return formatMoneyAmount_(minValue) + " ~ " + formatMoneyAmount_(maxValue);
}

function getPublicSettings_() {
  var s = getSettings_();
  return {
    hero_title: s.hero_title || "",
    hero_subtitle: s.hero_subtitle || "",
    ui_primary_hex: s.ui_primary_hex || "#2F6B5E",
    ui_radius: Number(s.ui_radius || 18),
    ui_density: s.ui_density || "comfortable",
    auto_advance_single_select: s.auto_advance_single_select === "Y",
    auto_advance_delay_ms: Number(s.auto_advance_delay_ms || 180),
    file_upload_max_count: Number(s.file_upload_max_count || 10),
    file_upload_max_mb: Number(s.file_upload_max_mb || 20),
    admin_header_compact_mode: s.admin_header_compact_mode || "Y",
    admin_hide_header_description: s.admin_hide_header_description || "Y",
    admin_kpi_cards_mode: s.admin_kpi_cards_mode || "STRIP",
    admin_workspace_default_tab: s.admin_workspace_default_tab || "requests",
    admin_detail_view_mode: s.admin_detail_view_mode || "DRAWER",
    admin_detail_single_flight: s.admin_detail_single_flight || "Y",
    admin_disable_actions_while_pending: s.admin_disable_actions_while_pending || "Y",
    admin_loading_state_guard: s.admin_loading_state_guard || "Y",
    admin_material_actions_layout: s.admin_material_actions_layout || "STACKED",
    admin_recommendation_allow_delete_auto: s.admin_recommendation_allow_delete_auto || "Y",
    admin_quote_open_mode: s.admin_quote_open_mode || "DIRECT_EDIT",
    admin_quote_open_hint_enabled: s.admin_quote_open_hint_enabled || "Y",
    survey_submit_async_postprocess: s.survey_submit_async_postprocess || "Y",
    result_data_mode: s.result_data_mode || "STORED_FIRST",
    result_recommendations_mode: s.result_recommendations_mode || "LAZY_FALLBACK",
    result_cache_ttl_sec: Number(s.result_cache_ttl_sec || 300),
    perf_request_detail_cache_ttl_sec: Number(s.perf_request_detail_cache_ttl_sec || 30),
    perf_material_search_debounce_ms: Number(s.perf_material_search_debounce_ms || 240),
    perf_material_search_min_query_length: Number(s.perf_material_search_min_query_length || 2),
    perf_result_first_paint_mode: s.perf_result_first_paint_mode || "SUMMARY_FIRST",
    app_version: s.app_version || APP_VERSION_,
    survey_version: s.survey_version || "SURVEY_V6"
  };
}

function getUiContentRows_() {
  return readStaticRowsCached_("UiContent").filter(function(row) {
    return ynToBool_(row.is_active, true);
  });
}

function getUiContentText_(contentKey, fallback) {
  var target = String(contentKey || "").trim();
  if (!target) return String(fallback || "");
  var rows = getUiContentRows_();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].content_key || "").trim() !== target) continue;
    return String(rows[i].body || rows[i].title || rows[i].content_value || fallback || "").trim();
  }
  return String(fallback || "");
}
