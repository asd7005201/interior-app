/**
 * ============================================================
 * PREQUOTE APP — Code.js
 * 가견적 설문 + 자동 견적 + 관리자 대시보드
 * ============================================================
 * File structure:
 *   Code.js    — this file (routing, survey engine, estimate engine, admin API)
 *   utils.js   — shared utilities (spreadsheet/settings helpers)
 *   BaseLib.gs — bridge to quote master app (material/template lookup)
 *
 * Sheet structure (prequote_DB):
 *   Settings, SurveyQuestions, SurveyOptions, SurveyLogic,
 *   SurveyTagRules, SurveyProfiles, QuestionScopeMap,
 *   QuickStartBundles, EstimateRules,
 *   Requests, RequestAnswers, RequestRecommendations,
 *   RequestFiles, RequestTimeline,
 *   MaterialsCache, TemplatesCache, SyncLog,
 *   UiContent, AdminUsers
 * ============================================================
 */

var APP_VERSION_ = "PREQUOTE_V6";
var BUILD_TAG_ = "2026-03-12T00:00:00Z";
var __HEADERS_CACHE_ = Object.create(null);
var __COLMAP_CACHE_ = Object.create(null);
var __REQUEST_ROW_INDEX_CACHE_ = Object.create(null);
var __SHEET_META_MEMO_ = Object.create(null);
var __SHEET_ROWS_MEMO_ = Object.create(null);
var __ROW_OBJECT_MEMO_ = Object.create(null);
var __COLUMN_MATCH_MEMO_ = Object.create(null);
var __REQUEST_DUPLICATE_LOOKUP_MEMO_ = Object.create(null);
var __ENSURED_SHEET_SCHEMA_MEMO_ = Object.create(null);
var STATIC_ROWS_CACHE_PREFIX_ = "PQ_STATIC_ROWS_V1";
var STATIC_ROWS_CACHE_TTL_SEC_ = 600;
var REQUEST_ROW_INDEX_SHEET_ = "RequestRowIndex";
var SYNC_STATE_SHEET_ = "SyncState";
var PERF_METRICS_SHEET_ = "PerfMetrics";
var REQUEST_POSTPROCESS_QUEUE_SHEET_ = "RequestPostprocessQueue";
var REQUEST_ROW_INDEX_VERSION_ = "2026-03-13-v1";
var SHEET_SCHEMA_CACHE_PREFIX_ = "PQ_SCHEMA_V2";
var SHEET_SCHEMA_CACHE_TTL_SEC_ = 600;
var SURVEY_CONFIG_CACHE_PREFIX_ = "PQ_SURVEY_CONFIG_V2";
var SURVEY_CONFIG_CACHE_TTL_SEC_ = 3600;
var RESULT_BOOTSTRAP_CACHE_PREFIX_ = "PQ_RESULT_BOOTSTRAP_V1";
var RESULT_BOOTSTRAP_CACHE_TTL_SEC_ = 900;
var SURVEY_SUBMIT_POSTPROCESS_TRIGGER_HANDLER_ = "runSurveySubmitPostprocessQueue_";
var SURVEY_SUBMIT_POSTPROCESS_TRIGGER_PREFIX_ = "PQ_SUBMIT_POSTPROCESS_TRIGGER";
var SURVEY_SUBMIT_POSTPROCESS_TRIGGER_DELAY_MS_ = 10000;
var SURVEY_SUBMIT_POSTPROCESS_BATCH_SIZE_ = 4;
var SURVEY_SUBMIT_POSTPROCESS_MAX_ATTEMPTS_ = 3;
var SYNC_KEY_MATERIALS_CACHE_ = "materials_cache";
var SYNC_KEY_TEMPLATES_CACHE_ = "templates_cache";
var SURVEY_CLIENT_QUESTION_FIELDS_ = [
  "question_id",
  "question_code",
  "question_type",
  "title",
  "description",
  "placeholder",
  "helper_text",
  "step_no",
  "sort_order",
  "is_required",
  "visible_if_expr",
  "exposure_scope"
];
var SURVEY_CLIENT_OPTION_FIELDS_ = [
  "question_id",
  "option_code",
  "option_label",
  "option_description",
  "sort_order"
];
var SURVEY_CLIENT_SCOPE_FIELDS_ = [
  "trigger_question_id",
  "trigger_values_json",
  "target_question_ids_json"
];
var SURVEY_CLIENT_BUNDLE_FIELDS_ = [
  "bundle_id",
  "bundle_label",
  "project_type",
  "sort_order",
  "default_answers_json",
  "locked_question_codes_json",
  "recommended_tags_json",
  "recommended_template_filters_json"
];
var SURVEY_CLIENT_UI_CONTENT_FIELDS_ = [
  "content_key",
  "content_type",
  "title",
  "body",
  "content_value",
  "sort_order"
];
var REQUEST_STATUS_OPTIONS_ = ["NEW", "CONTACTED", "SCHEDULED", "CONVERTED", "CLOSED"];
var ADMIN_REVIEW_STATUS_OPTIONS_ = ["AUTO", "IN_REVIEW", "OVERRIDDEN"];
var REQUEST_PRIORITY_OPTIONS_ = ["LOW", "NORMAL", "HIGH", "URGENT"];
var REQUEST_WORK_STATE_OPTIONS_ = ["NEW", "REVIEWING", "WAITING_CUSTOMER", "READY_FOR_QUOTE", "DONE"];
var REQUEST_ADMIN_REVIEW_SHEET_ = "RequestAdminReviews";
var REQUEST_ADMIN_REVIEW_HEADERS_ = [
  "review_id",
  "request_id",
  "edited_at",
  "actor_type",
  "actor_id",
  "review_status",
  "override_estimate_min",
  "override_estimate_max",
  "override_recommendations_json",
  "override_answers_json",
  "estimate_reason",
  "review_note",
  "review_summary",
  "final_estimate_min",
  "final_estimate_max",
  "final_recommendation_count",
  "source_label",
  "is_current",
  "linked_quote_id",
  "linked_quote_status",
  "slack_notified_at",
  "created_at",
  "updated_at",
  "payload_json"
];
var REQUEST_RECOMMENDATION_HEADERS_ = [
  "request_id",
  "rec_id",
  "rec_type",
  "rank",
  "template_id",
  "material_id",
  "trade_code",
  "material_type",
  "title",
  "subtitle",
  "reason_text",
  "score",
  "price_hint_min",
  "price_hint_max",
  "matched_tags_json",
  "snapshot_json",
  "is_selected",
  "created_at",
  "is_visible",
  "is_deleted",
  "is_manual",
  "image_url",
  "image_file_id",
  "brand",
  "spec",
  "source_type",
  "source_ref_id",
  "source_ref_key",
  "review_id",
  "updated_at",
  "hidden_at",
  "deleted_at",
  "note"
];
var REQUEST_ASSIGNMENT_HEADERS_ = [
  "assignment_id",
  "request_id",
  "assigned_to_name",
  "assigned_to_email",
  "assigned_to_slack_id",
  "assigned_by",
  "assigned_at",
  "is_active",
  "note",
  "payload_json"
];
var REQUEST_TASK_HEADERS_ = [
  "task_id",
  "request_id",
  "task_type",
  "task_label",
  "task_status",
  "priority",
  "assignee_name",
  "assignee_email",
  "assignee_slack_id",
  "due_at",
  "remind_at",
  "created_at",
  "updated_at",
  "completed_at",
  "note",
  "payload_json"
];
var REQUEST_CHECKLIST_TEMPLATE_HEADERS_ = [
  "template_id",
  "project_type",
  "checklist_code",
  "checklist_label",
  "is_required",
  "sort_order",
  "note"
];
var REQUEST_CHECKLIST_ITEM_HEADERS_ = [
  "item_id",
  "request_id",
  "template_id",
  "checklist_code",
  "checklist_label",
  "is_required",
  "is_completed",
  "completed_at",
  "completed_by",
  "note",
  "sort_order",
  "payload_json"
];
var REQUEST_NOTIFICATION_LOG_HEADERS_ = [
  "notification_id",
  "request_id",
  "channel",
  "event_type",
  "dedup_key",
  "recipient_key",
  "recipient_value",
  "status",
  "created_at",
  "sent_at",
  "last_error",
  "payload_json"
];
var REQUEST_QUOTE_DRAFT_LINK_HEADERS_ = [
  "link_id",
  "request_id",
  "review_id",
  "quote_id",
  "status",
  "created_at",
  "updated_at",
  "source_app",
  "target_app",
  "open_url",
  "note",
  "payload_json"
];
var DASHBOARD_SAVED_VIEW_HEADERS_ = [
  "view_id",
  "user_key",
  "view_name",
  "filter_json",
  "sort_json",
  "is_default",
  "created_at",
  "updated_at",
  "note"
];
var REQUEST_ROW_INDEX_HEADERS_ = [
  "request_id",
  "requests_row_no",
  "answers_row_nos_json",
  "timeline_row_nos_json",
  "files_row_nos_json",
  "admin_review_row_nos_json",
  "assignments_row_nos_json",
  "tasks_row_nos_json",
  "checklist_item_row_nos_json",
  "quote_link_row_nos_json",
  "latest_review_id",
  "latest_review_row_no",
  "updated_at",
  "index_version"
];
var SYNC_STATE_HEADERS_ = [
  "sync_key",
  "source_app",
  "source_spreadsheet_id",
  "shared_master_version",
  "last_synced_at",
  "last_status",
  "read_count",
  "write_count",
  "duration_ms",
  "note"
];
var PERF_METRICS_HEADERS_ = [
  "metric_id",
  "measured_at",
  "metric_name",
  "duration_ms",
  "request_id",
  "actor_type",
  "actor_id",
  "context_json",
  "success_yn",
  "note"
];
var QUOTE_DB_PERF_METRICS_HEADERS_ = [
  "metric_id",
  "measured_at",
  "metric_name",
  "duration_ms",
  "entity_id",
  "actor_type",
  "actor_id",
  "context_json",
  "success_yn",
  "note"
];
var REQUEST_POSTPROCESS_QUEUE_HEADERS_ = [
  "job_id",
  "request_id",
  "job_type",
  "status",
  "attempt_count",
  "created_at",
  "scheduled_at",
  "started_at",
  "processed_at",
  "last_error",
  "payload_json"
];
var REQUEST_REQUIRED_COLUMNS_ = [
  "priority",
  "consultant_email",
  "memo",
  "internal_memo",
  "lead_score",
  "quote_conversion_status",
  "linked_quote_id",
  "assignee_name",
  "assignee_email",
  "assignee_slack_id",
  "next_action_type",
  "next_action_note",
  "next_action_due_at",
  "reminder_at",
  "work_state",
  "latest_review_id",
  "latest_review_at",
  "final_estimate_min",
  "final_estimate_max",
  "final_estimate_source",
  "final_recommendation_count",
  "duplicate_phone_key",
  "duplicate_email_key",
  "duplicate_customer_count",
  "review_status_label",
  "quote_draft_status",
  "answers_snapshot_json"
];
var REQUEST_TIMELINE_REQUIRED_COLUMNS_ = [
  "event_label",
  "dedup_key",
  "notify_channel",
  "notify_status",
  "linked_ref_id"
];
var REQUEST_FILE_REQUIRED_COLUMNS_ = [
  "thumbnail_url",
  "caption",
  "is_primary"
];
var QUOTE_DB_REQUIRED_QUOTE_COLUMNS_ = [
  "source_request_id",
  "source_review_id",
  "source_app",
  "assignee_name",
  "assignee_slack_id",
  "draft_stage",
  "prequote_summary_json",
  "recommended_material_ids_json",
  "recommended_materials_snapshot_json"
];
var QUOTE_DB_REQUIRED_ITEM_COLUMNS_ = [
  "source_request_id",
  "source_review_id",
  "source_request_rec_id",
  "imported_from_prequote",
  "source_material_id",
  "source_material_image_url"
];
var QUOTE_DRAFT_QUEUE_HEADERS_ = [
  "queue_id",
  "request_id",
  "review_id",
  "created_at",
  "status",
  "customer_name",
  "contact_phone",
  "project_type",
  "assignee_name",
  "payload_json",
  "linked_quote_id",
  "processed_at",
  "last_error"
];
var QUOTE_DRAFT_ITEM_HEADERS_ = [
  "queue_id",
  "request_id",
  "review_id",
  "seq",
  "material_id",
  "name",
  "brand",
  "spec",
  "image_file_id",
  "image_url",
  "qty",
  "unit",
  "unit_price",
  "amount",
  "note",
  "payload_json"
];
var QUOTE_IMPORT_LOG_HEADERS_ = [
  "log_id",
  "created_at",
  "request_id",
  "review_id",
  "quote_id",
  "status",
  "message",
  "payload_json"
];
var STATUS_LABEL_MAP_ = {
  NEW: "신규 접수",
  CONTACTED: "1차 연락",
  SCHEDULED: "상담 예정",
  CONVERTED: "실견적 전환",
  CLOSED: "종료"
};
var REVIEW_STATUS_LABEL_MAP_ = {
  AUTO: "자동 산출",
  IN_REVIEW: "관리자 검토중",
  OVERRIDDEN: "관리자 수정완료"
};
var EVENT_TYPE_LABEL_MAP_ = {
  STATUS_CHANGE: "상태 변경",
  ADMIN_NOTE: "관리자 메모",
  REVIEW_SAVE: "검토 저장",
  RECALCULATE: "자동 재계산",
  ASSIGNMENT_UPDATE: "담당자 지정",
  TASK_UPDATE: "다음 액션 갱신",
  CHECKLIST_SAVE: "체크리스트 저장",
  BULK_UPDATE: "일괄 처리",
  QUOTE_DRAFT_CREATED: "견적 초안 생성",
  SLACK_NOTIFY: "슬랙 알림",
  RESULT_VIEWED: "결과 확인"
};
var SEVERITY_LABEL_MAP_ = {
  HIGH: "높음",
  MEDIUM: "중간",
  LOW: "낮음"
};
var PRIORITY_LABEL_MAP_ = {
  LOW: "낮음",
  NORMAL: "보통",
  HIGH: "높음",
  URGENT: "긴급"
};
var WORK_STATE_LABEL_MAP_ = {
  NEW: "신규",
  REVIEWING: "검토 중",
  WAITING_CUSTOMER: "고객 회신 대기",
  READY_FOR_QUOTE: "실견적 전환 대기",
  DONE: "처리 완료"
};
var TASK_STATUS_LABEL_MAP_ = {
  OPEN: "대기",
  IN_PROGRESS: "진행 중",
  DONE: "완료",
  CANCELLED: "취소"
};
var ANSWER_GROUP_LABEL_MAP_ = {
  ENTRY: "프로젝트 설정",
  BASIC: "기본 정보",
  SCOPE: "공간 정보",
  TRADE: "공정 범위",
  STYLE: "취향",
  LIFE: "생활 정보",
  RISK: "현장 리스크",
  BUDGET: "예산",
  SCHEDULE: "일정",
  CONTACT: "연락 선호",
  OTHER: "기타"
};
var ANSWER_GROUP_ORDER_MAP_ = {
  ENTRY: 10,
  BASIC: 20,
  SCOPE: 30,
  TRADE: 40,
  STYLE: 50,
  LIFE: 60,
  RISK: 70,
  BUDGET: 80,
  SCHEDULE: 90,
  CONTACT: 100,
  OTHER: 110
};

// ============================================================
//  INITIALIZATION
//  에디터에서 최초 1회 실행: initializeApp()
// ============================================================

/**
 * 최초 배포 전 에디터에서 1회 실행해주세요.
 * 컨테이너 바운드(시트에 붙은) 스크립트라면 자동 감지합니다.
 * 독립 스크립트라면 initializeAppManual('스프레드시트ID') 를 사용하세요.
 */
function initializeApp() {
  var id = setupSpreadsheetId();
  Logger.log("SPREADSHEET_ID set: " + id);
  Logger.log("Settings loaded: " + JSON.stringify(getSettings_()));
  return { spreadsheet_id: id, status: "OK" };
}

function initializeAppManual(spreadsheetId) {
  var id = setupSpreadsheetIdFromManual(spreadsheetId);
  Logger.log("SPREADSHEET_ID set (manual): " + id);
  return { spreadsheet_id: id, status: "OK" };
}

/**
 * 모든 진입점에서 호출. Script Properties에 ID가 없으면 자동 감지 시도.
 * getSpreadsheetId_() 자체에 fallback이 있으므로 여기서는 한 번 호출만 해줌.
 */
function ensureSpreadsheetId_() {
  try {
    getSpreadsheetId_();
  } catch (e) {
    // 컨테이너 바운드라면 자동 설정 시도
    try {
      setupSpreadsheetId();
    } catch (e2) {
      throw new Error(
        "스프레드시트 ID를 찾을 수 없습니다.\n" +
        "에디터에서 initializeApp() 을 먼저 실행해주세요.\n" +
        "원본 오류: " + e.message
      );
    }
  }
}

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

// ============================================================
//  HELPERS
// ============================================================

function escapeHtml_(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function javascriptResponse_(source) {
  return ContentService.createTextOutput(String(source || "")).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function sanitizeJsonpCallback_(callbackName) {
  var name = String(callbackName || "").trim();
  if (!name) return "";
  if (!/^[A-Za-z_$][0-9A-Za-z_$\.]{0,127}$/.test(name)) throw new Error("Invalid callback");
  return name;
}

function jsonOrJsonpResponse_(e, obj) {
  var callback = sanitizeJsonpCallback_((e && e.parameter && e.parameter.callback) || "");
  if (!callback) return jsonResponse_(obj);
  var payload = JSON.stringify(obj).replace(/<\//g, "<\\/");
  return javascriptResponse_(callback + "(" + payload + ");");
}

function sanitizePostMessageOrigin_(origin) {
  var raw = String(origin || "").trim();
  if (!raw || raw === "null") return "*";
  if (/^https?:\/\/[^\/]+$/i.test(raw)) return raw;
  return "*";
}

function safeJsonForInlineScript_(value) {
  return JSON.stringify(value).replace(/<\//g, "<\\/");
}

function safeJsonParse_(s, fallback) {
  try { return JSON.parse(String(s || "")); } catch (e) { return fallback !== undefined ? fallback : null; }
}

function toNumber_(value, fallback) {
  var num = Number(value);
  if (isFinite(num)) return num;
  return fallback !== undefined ? fallback : 0;
}

function ynToBool_(value, defaultValue) {
  if (value === true || value === false) return value;
  var raw = String(value === undefined || value === null ? "" : value).trim().toUpperCase();
  if (!raw) return !!defaultValue;
  if (raw === "Y" || raw === "TRUE" || raw === "1") return true;
  if (raw === "N" || raw === "FALSE" || raw === "0") return false;
  return !!defaultValue;
}

function pushUnique_(arr, value) {
  var out = arr || [];
  var item = String(value || "").trim();
  if (!item) return out;
  if (out.indexOf(item) < 0) out.push(item);
  return out;
}

function toUniqueStringArray_(values) {
  var out = [];
  if (!Array.isArray(values)) return out;
  for (var i = 0; i < values.length; i++) pushUnique_(out, values[i]);
  return out;
}

function normalizeStringList_(value) {
  if (Array.isArray(value)) return toUniqueStringArray_(value);
  var raw = String(value || "").trim();
  if (!raw) return [];
  return toUniqueStringArray_(raw.split(","));
}

function cacheJsonGet_(key) {
  var cacheKey = String(key || "").trim();
  if (!cacheKey) return null;
  try {
    var raw = CacheService.getScriptCache().get(cacheKey);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function cacheJsonPut_(key, obj, ttlSec) {
  var cacheKey = String(key || "").trim();
  if (!cacheKey) return;
  try {
    CacheService.getScriptCache().put(cacheKey, JSON.stringify(obj), Math.max(Number(ttlSec || 60), 30));
  } catch (e) {}
}

function sheetSchemaCacheKey_(sheetName, headers, overrideId) {
  return [
    SHEET_SCHEMA_CACHE_PREFIX_,
    getSpreadsheetId_(overrideId),
    String(sheetName || "").trim(),
    sha256Hex_((Array.isArray(headers) ? headers : []).join("|")).slice(0, 16)
  ].join("_");
}

function markSheetSchemaEnsured_(cacheKey) {
  var key = String(cacheKey || "").trim();
  if (!key) return;
  __ENSURED_SHEET_SCHEMA_MEMO_[key] = true;
  try {
    CacheService.getScriptCache().put(key, "1", SHEET_SCHEMA_CACHE_TTL_SEC_);
  } catch (e) {}
}

function isSheetSchemaEnsured_(cacheKey) {
  var key = String(cacheKey || "").trim();
  if (!key) return false;
  if (__ENSURED_SHEET_SCHEMA_MEMO_[key]) return true;
  try {
    var cached = CacheService.getScriptCache().get(key);
    if (!cached) return false;
    __ENSURED_SHEET_SCHEMA_MEMO_[key] = true;
    return true;
  } catch (e) {
    return false;
  }
}

function sheetMetaCacheKey_(sheetName, overrideId) {
  return [getSpreadsheetId_(overrideId), String(sheetName || "").trim()].join("|");
}

function sheetRowMemoKey_(sheetName, overrideId) {
  return sheetMetaCacheKey_(sheetName, overrideId);
}

function rowMemoKey_(sheetName, rowNo, overrideId) {
  return [sheetRowMemoKey_(sheetName, overrideId), Number(rowNo || 0)].join("|");
}

function columnMatchMemoKey_(sheetName, colName, value, overrideId) {
  return [
    sheetRowMemoKey_(sheetName, overrideId),
    String(colName || "").trim(),
    String(value || "").trim()
  ].join("|");
}

function cacheRowObjectMemo_(sheetName, rowObj, overrideId) {
  var rowNo = Number(rowObj && rowObj._rowNo || 0);
  if (rowNo < 2) return;
  __ROW_OBJECT_MEMO_[rowMemoKey_(sheetName, rowNo, overrideId)] = rowObj;
}

function getCachedRowObjectMemo_(sheetName, rowNo, overrideId) {
  return __ROW_OBJECT_MEMO_[rowMemoKey_(sheetName, rowNo, overrideId)] || null;
}

function invalidateSheetReadCache_(sheetName, overrideId) {
  var sheetKey = sheetRowMemoKey_(sheetName, overrideId);
  delete __SHEET_ROWS_MEMO_[sheetKey];

  for (var rowKey in __ROW_OBJECT_MEMO_) {
    if (!Object.prototype.hasOwnProperty.call(__ROW_OBJECT_MEMO_, rowKey)) continue;
    if (rowKey.indexOf(sheetKey + "|") !== 0) continue;
    delete __ROW_OBJECT_MEMO_[rowKey];
  }
  for (var matchKey in __COLUMN_MATCH_MEMO_) {
    if (!Object.prototype.hasOwnProperty.call(__COLUMN_MATCH_MEMO_, matchKey)) continue;
    if (matchKey.indexOf(sheetKey + "|") !== 0) continue;
    delete __COLUMN_MATCH_MEMO_[matchKey];
  }

  var sheetNameKey = String(sheetName || "").trim();
  if (sheetNameKey === "Requests") {
    delete __REQUEST_DUPLICATE_LOOKUP_MEMO_[String(getSpreadsheetId_(overrideId) || "").trim()];
  }
  if (sheetNameKey === REQUEST_ROW_INDEX_SHEET_) {
    __REQUEST_ROW_INDEX_CACHE_ = Object.create(null);
  }
}

function getHeaders_(sheetName, overrideId) {
  var key = sheetMetaCacheKey_(sheetName, overrideId);
  if (__HEADERS_CACHE_[key]) return __HEADERS_CACHE_[key].slice();
  var sh = getSheet_(sheetName, overrideId);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  __HEADERS_CACHE_[key] = headers;
  delete __COLMAP_CACHE_[key];
  return headers.slice();
}

function getColMap_(sheetName, overrideId) {
  var key = sheetMetaCacheKey_(sheetName, overrideId);
  if (__COLMAP_CACHE_[key]) return __COLMAP_CACHE_[key];
  var headers = getHeaders_(sheetName, overrideId);
  var cols = Object.create(null);
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] && !Object.prototype.hasOwnProperty.call(cols, headers[i])) cols[headers[i]] = i;
  }
  __COLMAP_CACHE_[key] = cols;
  return cols;
}

function invalidateSheetMetaCache_(sheetName, overrideId) {
  var key = sheetMetaCacheKey_(sheetName, overrideId);
  delete __HEADERS_CACHE_[key];
  delete __COLMAP_CACHE_[key];
  delete __SHEET_META_MEMO_[key];
}

function staticRowsCacheKey_(sheetName, overrideId) {
  return [
    STATIC_ROWS_CACHE_PREFIX_,
    getSpreadsheetId_(overrideId),
    String(sheetName || "").trim()
  ].join("_");
}

function readStaticRowsCached_(sheetName, ttlSec, overrideId) {
  var key = staticRowsCacheKey_(sheetName, overrideId);
  var cached = cacheJsonGet_(key);
  if (cached && Array.isArray(cached)) return cached;
  var rows = readAllRows_(sheetName, overrideId);
  cacheJsonPut_(key, rows, ttlSec || STATIC_ROWS_CACHE_TTL_SEC_);
  return rows;
}

function sheetMeta_(sh) {
  var sheet = sh;
  var sheetName = String(sheet.getName() || "").trim();
  var ss = sheet.getParent ? sheet.getParent() : null;
  var ssId = ss && ss.getId ? String(ss.getId() || "").trim() : "";
  var key = sheetMetaCacheKey_(sheetName, ssId);
  if (__SHEET_META_MEMO_[key]) return __SHEET_META_MEMO_[key];
  __SHEET_META_MEMO_[key] = {
    headers: getHeaders_(sheetName, ssId),
    cols: getColMap_(sheetName, ssId)
  };
  return __SHEET_META_MEMO_[key];
}

function rowToObj_(headers, row, rowNo) {
  var obj = rowNo ? { _rowNo: Number(rowNo || 0) } : {};
  for (var i = 0; i < headers.length; i++) {
    if (headers[i]) obj[headers[i]] = row[i] !== undefined && row[i] !== null ? row[i] : "";
  }
  return obj;
}

function readAllRows_(sheetName, overrideId) {
  var memoKey = sheetRowMemoKey_(sheetName, overrideId);
  if (__SHEET_ROWS_MEMO_[memoKey]) return __SHEET_ROWS_MEMO_[memoKey];
  var sh = getSheet_(sheetName, overrideId);
  var lr = sh.getLastRow();
  if (lr < 2) return [];
  var meta = {
    headers: getHeaders_(sheetName, overrideId),
    cols: getColMap_(sheetName, overrideId)
  };
  var data = sh.getRange(2, 1, lr - 1, meta.headers.length).getValues();
  var rows = data.map(function(row, idx) {
    var rowObj = rowToObj_(meta.headers, row, idx + 2);
    cacheRowObjectMemo_(sheetName, rowObj, overrideId);
    return rowObj;
  });
  __SHEET_ROWS_MEMO_[memoKey] = rows;
  return rows;
}

function appendRow_(sheetName, obj, overrideId) {
  var rowNos = appendRowsWithRowNos_(sheetName, [obj], overrideId);
  return rowNos.length ? rowNos[0] : 0;
}

function appendRows_(sheetName, rows, overrideId) {
  var list = Array.isArray(rows) ? rows.filter(Boolean) : [];
  if (!list.length) return 0;
  appendRowsWithRowNos_(sheetName, list, overrideId);
  return list.length;
}

function clearSheetBody_(sheetName, overrideId) {
  var sh = getSheet_(sheetName, overrideId);
  var meta = {
    headers: getHeaders_(sheetName, overrideId),
    cols: getColMap_(sheetName, overrideId)
  };
  var lr = sh.getLastRow();
  if (lr > 1) sh.getRange(2, 1, lr - 1, meta.headers.length).clearContent();
  invalidateSheetReadCache_(sheetName, overrideId);
}

function findRowByCol_(sheetName, colName, value, overrideId) {
  var rowNos = findRowNosByCol_(sheetName, colName, value, overrideId);
  if (!rowNos.length) return null;
  var sh = getSheet_(sheetName, overrideId);
  var meta = sheetMeta_(sh);
  var rows = readRowsByNumbers_(sh, meta, [rowNos[0]]);
  if (!rows.length) return null;
  return { rowNo: rowNos[0], data: rows[0], meta: meta, sheet: sh };
}

function readRowsByNumbers_(sh, meta, rowNos) {
  if (!Array.isArray(rowNos) || !rowNos.length) return [];
  var sheetName = String(sh.getName() || "").trim();
  var ss = sh.getParent ? sh.getParent() : null;
  var overrideId = ss && ss.getId ? String(ss.getId() || "").trim() : "";
  var seen = Object.create(null);
  var sorted = rowNos.filter(function(rowNo) {
    var num = Number(rowNo || 0);
    if (num < 2) return false;
    if (seen[num]) return false;
    seen[num] = 1;
    return true;
  }).sort(function(a, b) { return a - b; });
  if (!sorted.length) return [];

  var cachedRows = [];
  var missing = [];
  for (var i = 0; i < sorted.length; i++) {
    var cached = getCachedRowObjectMemo_(sheetName, sorted[i], overrideId);
    if (cached) cachedRows.push(cached);
    else missing.push(sorted[i]);
  }
  if (!missing.length) return cachedRows.sort(function(a, b) {
    return Number(a._rowNo || 0) - Number(b._rowNo || 0);
  });

  var blocks = [];
  var start = missing[0];
  var prev = missing[0];

  for (var m = 1; m < missing.length; m++) {
    if (missing[m] === prev + 1) {
      prev = missing[m];
      continue;
    }
    blocks.push({ rowNo: start, count: prev - start + 1 });
    start = missing[m];
    prev = missing[m];
  }
  blocks.push({ rowNo: start, count: prev - start + 1 });

  var out = [];
  for (var b = 0; b < blocks.length; b++) {
    var block = blocks[b];
    var values = sh.getRange(block.rowNo, 1, block.count, meta.headers.length).getValues();
    for (var r = 0; r < values.length; r++) {
      var rowObj = rowToObj_(meta.headers, values[r], block.rowNo + r);
      cacheRowObjectMemo_(sheetName, rowObj, overrideId);
      out.push(rowObj);
    }
  }
  return cachedRows.concat(out).sort(function(a, b) {
    return Number(a._rowNo || 0) - Number(b._rowNo || 0);
  });
}

function findRowsByCol_(sheetName, colName, value, overrideId) {
  var rowNos = findRowNosByCol_(sheetName, colName, value, overrideId);
  if (!rowNos.length) return [];
  var sh = getSheet_(sheetName, overrideId);
  var meta = sheetMeta_(sh);
  return readRowsByNumbers_(sh, meta, rowNos);
}

function updateRowFields_(sh, rowNo, meta, fields) {
  var row = sh.getRange(rowNo, 1, 1, meta.headers.length).getValues()[0];
  var changed = false;
  for (var key in fields) {
    if (!Object.prototype.hasOwnProperty.call(fields, key)) continue;
    var ci = meta.cols[key];
    if (ci === undefined) continue;
    row[ci] = fields[key];
    changed = true;
  }
  if (changed) {
    sh.getRange(rowNo, 1, 1, meta.headers.length).setValues([row]);
    invalidateSheetReadCache_(sh.getName(), sh.getParent && sh.getParent().getId ? sh.getParent().getId() : "");
  }
}

function ensureSheetWithHeaders_(sheetName, headers, overrideId) {
  var ss = getSpreadsheet_(overrideId);
  var ssId = String(ss.getId() || overrideId || "").trim();
  var schemaKey = sheetSchemaCacheKey_(sheetName, headers, ssId);
  if (isSheetSchemaEnsured_(schemaKey)) return getSheet_(sheetName, ssId);
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    invalidateSheetMetaCache_(sheetName, ssId);
    invalidateSheetCache_(sheetName, ssId);
    markSheetSchemaEnsured_(schemaKey);
    return getSheet_(sheetName, ssId);
  }

  var lastColumn = Math.max(sh.getLastColumn(), 1);
  var currentHeaders = sh.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(v) {
    return String(v || "").trim();
  });
  var changed = false;

  if (!currentHeaders.some(Boolean)) {
    currentHeaders = headers.slice();
    changed = true;
  } else {
    for (var i = 0; i < headers.length; i++) {
      if (currentHeaders.indexOf(headers[i]) >= 0) continue;
      currentHeaders.push(headers[i]);
      changed = true;
    }
  }

  if (changed) sh.getRange(1, 1, 1, currentHeaders.length).setValues([currentHeaders]);
  invalidateSheetMetaCache_(sheetName, ssId);
  invalidateSheetCache_(sheetName, ssId);
  markSheetSchemaEnsured_(schemaKey);
  return getSheet_(sheetName, ssId);
}

function ensurePrequoteOperationalSchema_() {
  ensureSheetWithHeaders_("Requests", REQUEST_REQUIRED_COLUMNS_);
  ensureSheetWithHeaders_("RequestRecommendations", REQUEST_RECOMMENDATION_HEADERS_);
  ensureSheetWithHeaders_("RequestTimeline", REQUEST_TIMELINE_REQUIRED_COLUMNS_);
  ensureSheetWithHeaders_("RequestFiles", REQUEST_FILE_REQUIRED_COLUMNS_);
  ensureSheetWithHeaders_(REQUEST_ADMIN_REVIEW_SHEET_, REQUEST_ADMIN_REVIEW_HEADERS_);
  ensureSheetWithHeaders_("RequestAssignments", REQUEST_ASSIGNMENT_HEADERS_);
  ensureSheetWithHeaders_("RequestTasks", REQUEST_TASK_HEADERS_);
  ensureSheetWithHeaders_("RequestChecklistTemplates", REQUEST_CHECKLIST_TEMPLATE_HEADERS_);
  ensureSheetWithHeaders_("RequestChecklistItems", REQUEST_CHECKLIST_ITEM_HEADERS_);
  ensureSheetWithHeaders_("RequestNotificationLog", REQUEST_NOTIFICATION_LOG_HEADERS_);
  ensureSheetWithHeaders_("RequestQuoteDraftLinks", REQUEST_QUOTE_DRAFT_LINK_HEADERS_);
  ensureSheetWithHeaders_("DashboardSavedViews", DASHBOARD_SAVED_VIEW_HEADERS_);
  ensureSheetWithHeaders_(REQUEST_ROW_INDEX_SHEET_, REQUEST_ROW_INDEX_HEADERS_);
  ensureSheetWithHeaders_(SYNC_STATE_SHEET_, SYNC_STATE_HEADERS_);
  ensureSheetWithHeaders_(PERF_METRICS_SHEET_, PERF_METRICS_HEADERS_);
  ensureSheetWithHeaders_(REQUEST_POSTPROCESS_QUEUE_SHEET_, REQUEST_POSTPROCESS_QUEUE_HEADERS_);
}

function ensureRequestRowIndexSheet_() {
  return ensureSheetWithHeaders_(REQUEST_ROW_INDEX_SHEET_, REQUEST_ROW_INDEX_HEADERS_);
}

function ensurePerfMetricsSheet_(overrideId) {
  return ensureSheetWithHeaders_(PERF_METRICS_SHEET_, PERF_METRICS_HEADERS_, overrideId);
}

function ensureRequestNotificationLogSheet_() {
  return ensureSheetWithHeaders_("RequestNotificationLog", REQUEST_NOTIFICATION_LOG_HEADERS_);
}

function ensureSurveySubmitStorageSchema_() {
  ensureSheetWithHeaders_("Requests", REQUEST_REQUIRED_COLUMNS_);
  ensureSheetWithHeaders_("RequestAnswers", [
    "request_id",
    "answer_id",
    "question_id",
    "question_code",
    "answer_type",
    "answer_value_text",
    "answer_value_number",
    "answer_value_json",
    "normalized_answer",
    "tags_applied_json",
    "score_delta",
    "created_at"
  ]);
  ensureSheetWithHeaders_("RequestTimeline", REQUEST_TIMELINE_REQUIRED_COLUMNS_);
  ensureSheetWithHeaders_("RequestFiles", REQUEST_FILE_REQUIRED_COLUMNS_);
  ensureRequestRowIndexSheet_();
  ensureSheetWithHeaders_(REQUEST_POSTPROCESS_QUEUE_SHEET_, REQUEST_POSTPROCESS_QUEUE_HEADERS_);
}

function ensureSurveySubmitPostprocessSchema_() {
  ensureSheetWithHeaders_("RequestChecklistItems", REQUEST_CHECKLIST_ITEM_HEADERS_);
  ensureSheetWithHeaders_("RequestChecklistTemplates", REQUEST_CHECKLIST_TEMPLATE_HEADERS_);
  ensureRequestNotificationLogSheet_();
  ensureSheetWithHeaders_(REQUEST_POSTPROCESS_QUEUE_SHEET_, REQUEST_POSTPROCESS_QUEUE_HEADERS_);
}

function perfFlagEnabled_(key, defaultValue) {
  return ynToBool_(getSettings_()[key], defaultValue);
}

function requestRowIndexEnabled_() {
  return perfFlagEnabled_("perf_use_request_row_index", true);
}

function perfMetricsEnabled_() {
  return perfFlagEnabled_("perf_metrics_enabled", false);
}

function partialAdminRefreshEnabled_() {
  return perfFlagEnabled_("perf_admin_partial_refresh", true);
}

function syncUseMasterVersion_() {
  return perfFlagEnabled_("perf_sync_use_master_version", true);
}

function requestRowIndexJsonColumnForSheet_(sheetName) {
  switch (String(sheetName || "").trim()) {
    case "RequestAnswers": return "answers_row_nos_json";
    case "RequestTimeline": return "timeline_row_nos_json";
    case "RequestFiles": return "files_row_nos_json";
    case REQUEST_ADMIN_REVIEW_SHEET_: return "admin_review_row_nos_json";
    case "RequestAssignments": return "assignments_row_nos_json";
    case "RequestTasks": return "tasks_row_nos_json";
    case "RequestChecklistItems": return "checklist_item_row_nos_json";
    case "RequestQuoteDraftLinks": return "quote_link_row_nos_json";
    default: return "";
  }
}

function uniqueSortedRowNos_(rowNos) {
  var seen = Object.create(null);
  var out = [];
  var list = Array.isArray(rowNos) ? rowNos : [];
  for (var i = 0; i < list.length; i++) {
    var rowNo = Number(list[i] || 0);
    if (rowNo < 2 || seen[rowNo]) continue;
    seen[rowNo] = 1;
    out.push(rowNo);
  }
  out.sort(function(a, b) { return a - b; });
  return out;
}

function parseRowNosJson_(raw) {
  var parsed = safeJsonParse_(raw, []);
  return uniqueSortedRowNos_(Array.isArray(parsed) ? parsed : []);
}

function cloneRequestRowIndexRecord_(record) {
  return record ? JSON.parse(JSON.stringify(record)) : null;
}

function cacheRequestRowIndexRecord_(record) {
  var rid = String(record && record.request_id || "").trim();
  if (!rid) return null;
  __REQUEST_ROW_INDEX_CACHE_[rid] = JSON.stringify(record);
  return cloneRequestRowIndexRecord_(record);
}

function getCachedRequestRowIndexRecord_(requestId) {
  var rid = String(requestId || "").trim();
  var raw = __REQUEST_ROW_INDEX_CACHE_[rid];
  if (!raw) return null;
  return safeJsonParse_(raw, null);
}

function emptyRequestRowIndexRecord_(requestId) {
  return {
    request_id: String(requestId || "").trim(),
    requests_row_no: 0,
    answers_row_nos: [],
    timeline_row_nos: [],
    files_row_nos: [],
    admin_review_row_nos: [],
    assignments_row_nos: [],
    tasks_row_nos: [],
    checklist_item_row_nos: [],
    quote_link_row_nos: [],
    latest_review_id: "",
    latest_review_row_no: 0,
    updated_at: nowIso_(),
    index_version: REQUEST_ROW_INDEX_VERSION_
  };
}

function requestRowIndexRecordFromRow_(row, rowNo) {
  var base = emptyRequestRowIndexRecord_(row && row.request_id || "");
  base.requests_row_no = Number((row || {}).requests_row_no || 0);
  base.answers_row_nos = parseRowNosJson_((row || {}).answers_row_nos_json);
  base.timeline_row_nos = parseRowNosJson_((row || {}).timeline_row_nos_json);
  base.files_row_nos = parseRowNosJson_((row || {}).files_row_nos_json);
  base.admin_review_row_nos = parseRowNosJson_((row || {}).admin_review_row_nos_json);
  base.assignments_row_nos = parseRowNosJson_((row || {}).assignments_row_nos_json);
  base.tasks_row_nos = parseRowNosJson_((row || {}).tasks_row_nos_json);
  base.checklist_item_row_nos = parseRowNosJson_((row || {}).checklist_item_row_nos_json);
  base.quote_link_row_nos = parseRowNosJson_((row || {}).quote_link_row_nos_json);
  base.latest_review_id = String((row || {}).latest_review_id || "").trim();
  base.latest_review_row_no = Number((row || {}).latest_review_row_no || 0);
  base.updated_at = String((row || {}).updated_at || "").trim() || nowIso_();
  base.index_version = String((row || {}).index_version || "").trim() || REQUEST_ROW_INDEX_VERSION_;
  if (rowNo) base._index_row_no = Number(rowNo || 0);
  return base;
}

function requestRowIndexRowObject_(record) {
  var row = emptyRequestRowIndexRecord_(record && record.request_id || "");
  var merged = Object.assign(row, record || {});
  return {
    request_id: merged.request_id,
    requests_row_no: Number(merged.requests_row_no || 0) || "",
    answers_row_nos_json: JSON.stringify(uniqueSortedRowNos_(merged.answers_row_nos)),
    timeline_row_nos_json: JSON.stringify(uniqueSortedRowNos_(merged.timeline_row_nos)),
    files_row_nos_json: JSON.stringify(uniqueSortedRowNos_(merged.files_row_nos)),
    admin_review_row_nos_json: JSON.stringify(uniqueSortedRowNos_(merged.admin_review_row_nos)),
    assignments_row_nos_json: JSON.stringify(uniqueSortedRowNos_(merged.assignments_row_nos)),
    tasks_row_nos_json: JSON.stringify(uniqueSortedRowNos_(merged.tasks_row_nos)),
    checklist_item_row_nos_json: JSON.stringify(uniqueSortedRowNos_(merged.checklist_item_row_nos)),
    quote_link_row_nos_json: JSON.stringify(uniqueSortedRowNos_(merged.quote_link_row_nos)),
    latest_review_id: String(merged.latest_review_id || "").trim(),
    latest_review_row_no: Number(merged.latest_review_row_no || 0) || "",
    updated_at: String(merged.updated_at || nowIso_()).trim(),
    index_version: REQUEST_ROW_INDEX_VERSION_
  };
}

function readRequestRowIndexRecord_(requestId) {
  var rid = String(requestId || "").trim();
  if (!rid || !requestRowIndexEnabled_()) return null;
  var cached = getCachedRequestRowIndexRecord_(rid);
  if (cached) return cached;
  var found = findRowByCol_(REQUEST_ROW_INDEX_SHEET_, "request_id", rid);
  if (!found) return null;
  return cacheRequestRowIndexRecord_(requestRowIndexRecordFromRow_(found.data, found.rowNo));
}

function upsertRequestRowIndexRecord_(requestId, patch) {
  var rid = String(requestId || "").trim();
  if (!rid || !requestRowIndexEnabled_()) return null;
  ensureRequestRowIndexSheet_();
  var current = readRequestRowIndexRecord_(rid) || emptyRequestRowIndexRecord_(rid);
  var next = Object.assign({}, current, patch || {}, {
    request_id: rid,
    updated_at: String((patch && patch.updated_at) || nowIso_()).trim(),
    index_version: REQUEST_ROW_INDEX_VERSION_
  });
  next.answers_row_nos = uniqueSortedRowNos_(next.answers_row_nos);
  next.timeline_row_nos = uniqueSortedRowNos_(next.timeline_row_nos);
  next.files_row_nos = uniqueSortedRowNos_(next.files_row_nos);
  next.admin_review_row_nos = uniqueSortedRowNos_(next.admin_review_row_nos);
  next.assignments_row_nos = uniqueSortedRowNos_(next.assignments_row_nos);
  next.tasks_row_nos = uniqueSortedRowNos_(next.tasks_row_nos);
  next.checklist_item_row_nos = uniqueSortedRowNos_(next.checklist_item_row_nos);
  next.quote_link_row_nos = uniqueSortedRowNos_(next.quote_link_row_nos);
  var rowObj = requestRowIndexRowObject_(next);
  if (current._index_row_no) {
    var sh = getSheet_(REQUEST_ROW_INDEX_SHEET_);
    var meta = sheetMeta_(sh);
    updateRowFields_(sh, current._index_row_no, meta, rowObj);
    next._index_row_no = current._index_row_no;
  } else {
    next._index_row_no = appendRow_(REQUEST_ROW_INDEX_SHEET_, rowObj);
  }
  return cacheRequestRowIndexRecord_(next);
}

function getRequestRowNosFromIndex_(record, sheetName) {
  var index = record || null;
  switch (String(sheetName || "").trim()) {
    case "Requests":
      return Number(index && index.requests_row_no || 0) >= 2 ? [Number(index.requests_row_no)] : [];
    case "RequestAnswers":
      return uniqueSortedRowNos_(index && index.answers_row_nos);
    case "RequestTimeline":
      return uniqueSortedRowNos_(index && index.timeline_row_nos);
    case "RequestFiles":
      return uniqueSortedRowNos_(index && index.files_row_nos);
    case REQUEST_ADMIN_REVIEW_SHEET_:
      return uniqueSortedRowNos_(index && index.admin_review_row_nos);
    case "RequestAssignments":
      return uniqueSortedRowNos_(index && index.assignments_row_nos);
    case "RequestTasks":
      return uniqueSortedRowNos_(index && index.tasks_row_nos);
    case "RequestChecklistItems":
      return uniqueSortedRowNos_(index && index.checklist_item_row_nos);
    case "RequestQuoteDraftLinks":
      return uniqueSortedRowNos_(index && index.quote_link_row_nos);
    default:
      return [];
  }
}

function rebuildRequestRowIndexRecord_(requestId) {
  var rid = String(requestId || "").trim();
  if (!rid || !requestRowIndexEnabled_()) return null;
  ensureRequestRowIndexSheet_();
  var reviewRowNos = findRowNosByCol_(REQUEST_ADMIN_REVIEW_SHEET_, "request_id", rid);
  var record = emptyRequestRowIndexRecord_(rid);
  record.requests_row_no = Number(findRowNosByCol_("Requests", "request_id", rid)[0] || 0);
  record.answers_row_nos = findRowNosByCol_("RequestAnswers", "request_id", rid);
  record.timeline_row_nos = findRowNosByCol_("RequestTimeline", "request_id", rid);
  record.files_row_nos = findRowNosByCol_("RequestFiles", "request_id", rid);
  record.admin_review_row_nos = reviewRowNos;
  record.assignments_row_nos = findRowNosByCol_("RequestAssignments", "request_id", rid);
  record.tasks_row_nos = findRowNosByCol_("RequestTasks", "request_id", rid);
  record.checklist_item_row_nos = findRowNosByCol_("RequestChecklistItems", "request_id", rid);
  record.quote_link_row_nos = findRowNosByCol_("RequestQuoteDraftLinks", "request_id", rid);
  if (reviewRowNos.length) {
    var sh = getSheet_(REQUEST_ADMIN_REVIEW_SHEET_);
    var meta = sheetMeta_(sh);
    var rows = readRowsByNumbers_(sh, meta, reviewRowNos);
    rows.sort(function(a, b) {
      return String(b.edited_at || "").localeCompare(String(a.edited_at || ""));
    });
    if (rows.length) {
      record.latest_review_id = String(rows[0].review_id || "").trim();
      record.latest_review_row_no = Number(rows[0]._rowNo || 0);
    }
  }
  return upsertRequestRowIndexRecord_(rid, record);
}

function readRequestRowsIndexed_(sheetName, requestId, options) {
  var rid = String(requestId || "").trim();
  if (!rid) return [];
  if (!requestRowIndexEnabled_()) return findRowsByCol_(sheetName, "request_id", rid);
  var opts = options || {};
  var record = opts.record || readRequestRowIndexRecord_(rid);
  var rowNos = getRequestRowNosFromIndex_(record, sheetName);
  if (rowNos.length) {
    var sh = getSheet_(sheetName);
    var meta = sheetMeta_(sh);
    var rows = readRowsByNumbers_(sh, meta, rowNos).filter(function(row) {
      return String(row.request_id || "").trim() === rid;
    });
    if (rows.length === rowNos.length) return rows;
  }
  var fallback = findRowsByCol_(sheetName, "request_id", rid);
  if (opts.allow_rebuild !== false) rebuildRequestRowIndexRecord_(rid);
  return fallback;
}

function findRequestRowById_(requestId) {
  var rid = String(requestId || "").trim();
  if (!rid) return null;
  if (requestRowIndexEnabled_()) {
    var record = readRequestRowIndexRecord_(rid);
    var rowNo = Number(record && record.requests_row_no || 0);
    if (rowNo >= 2) {
      var sh = getSheet_("Requests");
      var meta = sheetMeta_(sh);
      var rows = readRowsByNumbers_(sh, meta, [rowNo]);
      if (rows.length && String(rows[0].request_id || "").trim() === rid) {
        return { rowNo: rowNo, data: rows[0], meta: meta, sheet: sh };
      }
    }
  }
  var found = findRowByCol_("Requests", "request_id", rid);
  if (found && requestRowIndexEnabled_()) {
    upsertRequestRowIndexRecord_(rid, { requests_row_no: found.rowNo });
  }
  return found;
}

function afterAppendRows_(sheetName, rows, rowNos, overrideId) {
  if (!requestRowIndexEnabled_()) return;
  var primaryId = "";
  try { primaryId = getSpreadsheetId_(); } catch (e) { primaryId = ""; }
  var targetId = String(overrideId || primaryId || "").trim();
  if (!primaryId || !targetId || targetId !== primaryId) return;
  var columnName = requestRowIndexJsonColumnForSheet_(sheetName);
  var shouldTrack = sheetName === "Requests" || !!columnName || sheetName === REQUEST_ADMIN_REVIEW_SHEET_;
  if (!shouldTrack) return;
  var now = nowIso_();
  var grouped = Object.create(null);
  for (var i = 0; i < rows.length && i < rowNos.length; i++) {
    var row = rows[i] || {};
    var rid = String(row.request_id || "").trim();
    if (!rid) continue;
    if (!grouped[rid]) grouped[rid] = { request_id: rid, rowNos: [] };
    grouped[rid].rowNos.push(Number(rowNos[i] || 0));
    if (sheetName === "Requests") grouped[rid].requests_row_no = Number(rowNos[i] || 0);
    if (sheetName === REQUEST_ADMIN_REVIEW_SHEET_) {
      grouped[rid].latest_review_id = String(row.review_id || "").trim();
      grouped[rid].latest_review_row_no = Number(rowNos[i] || 0);
    }
  }
  for (var rid in grouped) {
    if (!Object.prototype.hasOwnProperty.call(grouped, rid)) continue;
    var current = readRequestRowIndexRecord_(rid) || emptyRequestRowIndexRecord_(rid);
    var patch = { updated_at: now, index_version: REQUEST_ROW_INDEX_VERSION_ };
    if (grouped[rid].requests_row_no) patch.requests_row_no = grouped[rid].requests_row_no;
    if (columnName) {
      var currentRows = getRequestRowNosFromIndex_(current, sheetName);
      patch[columnName.replace("_json", "")] = uniqueSortedRowNos_(currentRows.concat(grouped[rid].rowNos));
    }
    if (sheetName === REQUEST_ADMIN_REVIEW_SHEET_) {
      patch.latest_review_id = grouped[rid].latest_review_id;
      patch.latest_review_row_no = grouped[rid].latest_review_row_no;
    }
    upsertRequestRowIndexRecord_(rid, patch);
  }
}

function reviewStatusCodeFromLabel_(label) {
  var raw = String(label || "").trim();
  if (!raw) return "AUTO";
  for (var code in REVIEW_STATUS_LABEL_MAP_) {
    if (!Object.prototype.hasOwnProperty.call(REVIEW_STATUS_LABEL_MAP_, code)) continue;
    if (String(REVIEW_STATUS_LABEL_MAP_[code] || "").trim() === raw) return code;
  }
  return "AUTO";
}

function buildPerfMetricRow_(headers, metricName, durationMs, entityKey, context, success, note, actorType, actorId) {
  var row = {};
  for (var i = 0; i < headers.length; i++) row[headers[i]] = "";
  row.metric_id = uuid_();
  row.measured_at = nowIso_();
  row.metric_name = String(metricName || "").trim();
  row.duration_ms = Math.max(Number(durationMs || 0), 0);
  if (headers.indexOf("request_id") >= 0) row.request_id = String(entityKey || "").trim();
  if (headers.indexOf("entity_id") >= 0) row.entity_id = String(entityKey || "").trim();
  row.actor_type = String(actorType || "SYSTEM").trim();
  row.actor_id = String(actorId || "").trim();
  row.context_json = JSON.stringify(context || {});
  row.success_yn = success === false ? "N" : "Y";
  row.note = String(note || "").trim();
  return row;
}

function logPerfMetricSafe_(metricName, durationMs, requestId, context, success, note, actorType, actorId, overrideId) {
  if (!perfMetricsEnabled_()) return;
  try {
    ensurePerfMetricsSheet_(overrideId);
    appendRow_(PERF_METRICS_SHEET_, buildPerfMetricRow_(
      PERF_METRICS_HEADERS_,
      metricName,
      durationMs,
      requestId,
      context,
      success,
      note,
      actorType,
      actorId
    ), overrideId);
  } catch (e) {}
}

function logQuotePerfMetricSafe_(quoteDbId, metricName, entityId, durationMs, context, success, note, actorType, actorId) {
  try {
    ensureSheetWithHeaders_("PerfMetrics", QUOTE_DB_PERF_METRICS_HEADERS_, quoteDbId);
    appendRow_("PerfMetrics", buildPerfMetricRow_(
      QUOTE_DB_PERF_METRICS_HEADERS_,
      metricName,
      durationMs,
      entityId,
      context,
      success,
      note,
      actorType,
      actorId
    ), quoteDbId);
  } catch (e) {}
}

function readSyncStateRow_(syncKey) {
  var key = String(syncKey || "").trim();
  if (!key) return null;
  ensurePrequoteOperationalSchema_();
  var found = findRowByCol_(SYNC_STATE_SHEET_, "sync_key", key);
  if (!found) return null;
  return {
    rowNo: found.rowNo,
    sync_key: String(found.data.sync_key || "").trim(),
    source_app: String(found.data.source_app || "").trim(),
    source_spreadsheet_id: String(found.data.source_spreadsheet_id || "").trim(),
    shared_master_version: String(found.data.shared_master_version || "").trim(),
    last_synced_at: String(found.data.last_synced_at || "").trim(),
    last_status: String(found.data.last_status || "").trim(),
    read_count: toNumber_(found.data.read_count, 0),
    write_count: toNumber_(found.data.write_count, 0),
    duration_ms: toNumber_(found.data.duration_ms, 0),
    note: String(found.data.note || "").trim()
  };
}

function upsertSyncState_(syncKey, patch) {
  var key = String(syncKey || "").trim();
  if (!key) return null;
  ensurePrequoteOperationalSchema_();
  var current = readSyncStateRow_(key);
  var row = Object.assign({
    sync_key: key,
    source_app: "quote_app",
    source_spreadsheet_id: "",
    shared_master_version: "",
    last_synced_at: "",
    last_status: "",
    read_count: 0,
    write_count: 0,
    duration_ms: 0,
    note: ""
  }, current || {}, patch || {}, { sync_key: key });
  if (current && current.rowNo) {
    updateRowFields_(getSheet_(SYNC_STATE_SHEET_), current.rowNo, sheetMeta_(getSheet_(SYNC_STATE_SHEET_)), row);
  } else {
    appendRow_(SYNC_STATE_SHEET_, row);
  }
  return row;
}

function hasSheetBodyData_(sheetName, overrideId) {
  try {
    return getSheet_(sheetName, overrideId).getLastRow() >= 2;
  } catch (e) {
    return false;
  }
}

function ensureQuoteDraftImportSchema_() {
  var quoteDbId = getQuoteBaseDbId_();
  ensureSheetWithHeaders_("Quotes", QUOTE_DB_REQUIRED_QUOTE_COLUMNS_, quoteDbId);
  ensureSheetWithHeaders_("Items", QUOTE_DB_REQUIRED_ITEM_COLUMNS_, quoteDbId);
  ensureSheetWithHeaders_("PrequoteDraftQueue", QUOTE_DRAFT_QUEUE_HEADERS_, quoteDbId);
  ensureSheetWithHeaders_("PrequoteDraftItems", QUOTE_DRAFT_ITEM_HEADERS_, quoteDbId);
  ensureSheetWithHeaders_("PrequoteImportLog", QUOTE_IMPORT_LOG_HEADERS_, quoteDbId);
}

function findRowNosByCol_(sheetName, colName, value, overrideId) {
  var needle = String(value || "").trim();
  if (!needle) return [];
  var memoKey = columnMatchMemoKey_(sheetName, colName, needle, overrideId);
  if (__COLUMN_MATCH_MEMO_[memoKey]) return __COLUMN_MATCH_MEMO_[memoKey].slice();
  var sh = getSheet_(sheetName, overrideId);
  var meta = sheetMeta_(sh);
  var colIdx = meta.cols[colName];
  if (colIdx === undefined) return [];
  var rows = __SHEET_ROWS_MEMO_[sheetRowMemoKey_(sheetName, overrideId)];
  var rowNos = [];
  if (rows && rows.length) {
    rowNos = rows.filter(function(row) {
      return String(row[colName] || "").trim() === needle;
    }).map(function(row) {
      return Number(row._rowNo || 0);
    }).filter(Boolean).sort(function(a, b) { return a - b; });
  } else {
    var lr = sh.getLastRow();
    if (lr < 2) return [];
    var range = sh.getRange(2, colIdx + 1, lr - 1, 1);
    var cells = range.createTextFinder(needle).matchEntireCell(true).findAll();
    if (!cells || !cells.length) return [];
    rowNos = cells.map(function(cell) { return Number(cell.getRow() || 0); }).filter(Boolean).sort(function(a, b) { return a - b; });
  }
  __COLUMN_MATCH_MEMO_[memoKey] = rowNos.slice();
  return rowNos;
}

function clearRowsByRowNos_(sh, rowNos, width) {
  var rows = Array.isArray(rowNos) ? rowNos : [];
  var targetWidth = Math.max(Number(width || 0), 1);
  for (var i = 0; i < rows.length; i++) {
    if (Number(rows[i] || 0) < 2) continue;
    sh.getRange(rows[i], 1, 1, targetWidth).clearContent();
  }
  invalidateSheetReadCache_(sh.getName(), sh.getParent && sh.getParent().getId ? sh.getParent().getId() : "");
}

function writeRowsByRowNos_(sh, rowNos, values) {
  var rows = Array.isArray(rowNos) ? rowNos : [];
  var list = Array.isArray(values) ? values : [];
  for (var i = 0; i < rows.length && i < list.length; i++) {
    if (Number(rows[i] || 0) < 2) continue;
    sh.getRange(rows[i], 1, 1, list[i].length).setValues([list[i]]);
  }
  invalidateSheetReadCache_(sh.getName(), sh.getParent && sh.getParent().getId ? sh.getParent().getId() : "");
}

function appendRowsWithRowNos_(sheetName, rows, overrideId) {
  var list = Array.isArray(rows) ? rows.filter(Boolean) : [];
  if (!list.length) return [];
  var sh = getSheet_(sheetName, overrideId);
  var meta = {
    headers: getHeaders_(sheetName, overrideId),
    cols: getColMap_(sheetName, overrideId)
  };
  var startRow = sh.getLastRow() + 1;
  var values = list.map(function(obj) {
    return meta.headers.map(function(h) {
      return obj[h] !== undefined ? obj[h] : "";
    });
  });
  sh.getRange(startRow, 1, values.length, meta.headers.length).setValues(values);
  invalidateSheetReadCache_(sheetName, overrideId);
  var rowNos = [];
  for (var i = 0; i < values.length; i++) rowNos.push(startRow + i);
  afterAppendRows_(sheetName, list, rowNos, overrideId);
  return rowNos;
}

function replaceRowsByKey_(sheetName, keyCol, keyValue, rows, overrideId) {
  var list = Array.isArray(rows) ? rows.filter(Boolean) : [];
  var sh = getSheet_(sheetName, overrideId);
  var meta = sheetMeta_(sh);
  var existingRowNos = findRowNosByCol_(sheetName, keyCol, keyValue, overrideId);
  var values = list.map(function(obj) {
    return meta.headers.map(function(h) {
      return obj[h] !== undefined ? obj[h] : "";
    });
  });
  var reuseCount = Math.min(existingRowNos.length, values.length);
  if (reuseCount > 0) {
    writeRowsByRowNos_(sh, existingRowNos.slice(0, reuseCount), values.slice(0, reuseCount));
  }
  if (existingRowNos.length > reuseCount) {
    clearRowsByRowNos_(sh, existingRowNos.slice(reuseCount), meta.headers.length);
  }
  if (values.length > reuseCount) {
    var appendValues = values.slice(reuseCount);
    var startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, appendValues.length, meta.headers.length).setValues(appendValues);
  }
  invalidateSheetReadCache_(sheetName, overrideId);
}

// ============================================================
//  SURVEY DATA API (called from survey.html via google.script.run)
// ============================================================

function projectRowsForClient_(rows, fields) {
  var list = Array.isArray(rows) ? rows : [];
  var pickedFields = Array.isArray(fields) ? fields : [];
  return list.map(function(row) {
    var out = {};
    for (var i = 0; i < pickedFields.length; i++) {
      var key = pickedFields[i];
      out[key] = row && row[key] !== undefined ? row[key] : "";
    }
    return out;
  });
}

function surveyConfigCacheKey_() {
  return [
    SURVEY_CONFIG_CACHE_PREFIX_,
    getSpreadsheetId_(),
    BUILD_TAG_
  ].join("_");
}

function getCachedSurveyConfigBootstrap_() {
  return cacheJsonGet_(surveyConfigCacheKey_());
}

function getSurveyConfigBootstrapPayload_() {
  return getCachedSurveyConfigBootstrap_();
}

function resultBootstrapCacheKey_(requestId, shareToken) {
  var rid = String(requestId || "").trim();
  var tok = String(shareToken || "").trim();
  if (!rid || !tok) return "";
  return [
    RESULT_BOOTSTRAP_CACHE_PREFIX_,
    rid,
    tok
  ].join("_");
}

function buildResultBootstrapRequestView_(requestRow, optionIndex) {
  var row = requestRow || {};
  var projectType = String(row.project_type || "").trim().toUpperCase();
  var housingQuestionId = projectType === "COMMERCIAL" ? "C001_BIZ_TYPE" : "R001_HOUSING_TYPE";
  var areaQuestionId = projectType === "COMMERCIAL" ? "C002_AREA" : "R002_AREA";
  var scopeQuestionId = projectType === "COMMERCIAL" ? "C011_SCOPE_LEVEL" : "R011_SCOPE_LEVEL";
  return {
    request_id: row.request_id || "",
    customer_name: row.customer_name || "",
    project_type: row.project_type || "",
    project_type_label: lookupSurveyOptionLabel_("Q000_PROJECT_TYPE", row.project_type, optionIndex) || row.project_type || "",
    housing_type: row.housing_type || "",
    housing_type_label: lookupSurveyOptionLabel_(housingQuestionId, row.housing_type, optionIndex) || row.housing_type || "",
    area_py: row.area_py || "",
    area_label: lookupSurveyOptionLabel_(areaQuestionId, row.area_py, optionIndex) || row.area_py || "",
    scope_level: row.scope_level || "",
    scope_level_label: lookupSurveyOptionLabel_(scopeQuestionId, row.scope_level, optionIndex) || row.scope_level || "",
    flow_mode: row.flow_mode || "",
    flow_mode_label: lookupSurveyOptionLabel_("Q000_FLOW_MODE", row.flow_mode, optionIndex) || row.flow_mode || "",
    created_at: row.created_at || nowIso_(),
    status: row.status || "NEW",
    schedule_start_label: row.schedule_start || "",
    schedule_end_label: row.schedule_end || "",
    address_label: row.address_text || "",
    customer_note_label: row.customer_note || ""
  };
}

function buildResultBootstrapPayload_(requestRow, estimate) {
  var row = requestRow || {};
  var est = estimate || buildStoredEstimateFromRequest_(row);
  var requestView = buildResultBootstrapRequestView_(row, getSurveyOptionLabelIndex_());
  return {
    _bootstrap: true,
    request: requestView,
    estimate: Object.assign({}, est || {}, { source: "AUTO", source_label: "AUTO" }),
    flags: Array.isArray(est && est.flags) ? est.flags : [],
    recommendations: [],
    answers: [],
    answer_groups: [],
    notice_text: "",
    recommendation_status: "PENDING",
    recommendation_count: 0,
    first_paint: {
      range_label: formatMoneyRange(
        toNumber_(est && est.min, 0),
        toNumber_(est && est.max, 0)
      ),
      request_summary: [
        requestView.project_type_label || requestView.project_type || "",
        requestView.housing_type_label || requestView.housing_type || "",
        requestView.area_label || requestView.area_py || "",
        requestView.scope_level_label || requestView.scope_level || ""
      ].filter(Boolean).join(" / "),
      created_at_label: formatAdminDateTime_(row.created_at)
    }
  };
}

function cacheResultBootstrap_(requestRow, estimate) {
  var row = requestRow || {};
  var cacheKey = resultBootstrapCacheKey_(row.request_id, row.share_token);
  if (!cacheKey) return null;
  var payload = buildResultBootstrapPayload_(row, estimate);
  cacheJsonPut_(cacheKey, payload, RESULT_BOOTSTRAP_CACHE_TTL_SEC_);
  return payload;
}

function getCachedResultBootstrap_(requestId, shareToken) {
  var cacheKey = resultBootstrapCacheKey_(requestId, shareToken);
  if (!cacheKey) return null;
  return cacheJsonGet_(cacheKey);
}

function buildResultBootstrapForPage_(requestId, shareToken) {
  var cached = getCachedResultBootstrap_(requestId, shareToken);
  if (cached) return cached;

  var rid = String(requestId || "").trim();
  var tok = String(shareToken || "").trim();
  if (!rid || !tok) return null;

  var found = findRequestRowById_(rid);
  if (!found) return null;
  if (String(found.data.share_token || "").trim() !== tok) return null;
  return cacheResultBootstrap_(found.data, buildStoredEstimateFromRequest_(found.data));
}

function buildSurveyConfigPayload_() {
  var questions = projectRowsForClient_(
    readStaticRowsCached_("SurveyQuestions").filter(function(q) {
      return ynToBool_(q.is_active);
    }),
    SURVEY_CLIENT_QUESTION_FIELDS_
  );
  var options = projectRowsForClient_(
    readStaticRowsCached_("SurveyOptions").filter(function(o) {
      return ynToBool_(o.is_active);
    }),
    SURVEY_CLIENT_OPTION_FIELDS_
  );
  var profiles = [];
  var scopeMap = projectRowsForClient_(
    readStaticRowsCached_("QuestionScopeMap"),
    SURVEY_CLIENT_SCOPE_FIELDS_
  );
  var bundles = projectRowsForClient_(
    readStaticRowsCached_("QuickStartBundles").filter(function(b) {
      return ynToBool_(b.is_active);
    }),
    SURVEY_CLIENT_BUNDLE_FIELDS_
  );
  var uiContent = projectRowsForClient_(
    readStaticRowsCached_("UiContent").filter(function(u) {
      return ynToBool_(u.is_active);
    }),
    SURVEY_CLIENT_UI_CONTENT_FIELDS_
  );

  bundles.forEach(function(b) {
    b.default_answers_json = safeJsonParse_(b.default_answers_json, {});
    b.locked_question_codes_json = safeJsonParse_(b.locked_question_codes_json, []);
    b.recommended_tags_json = safeJsonParse_(b.recommended_tags_json, {});
    b.recommended_template_filters_json = safeJsonParse_(b.recommended_template_filters_json, {});
  });
  scopeMap.forEach(function(m) {
    m.trigger_values_json = safeJsonParse_(m.trigger_values_json, []);
    m.target_question_ids_json = safeJsonParse_(m.target_question_ids_json, []);
  });

  return {
    questions: questions,
    options: options,
    profiles: profiles,
    scopeMap: scopeMap,
    bundles: bundles,
    uiContent: uiContent,
    settings: getPublicSettings_(),
    app_url: getAppUrl_()
  };
}

/**
 * Load all survey config for the client in one call.
 * Returns: { questions, options, profiles, scopeMap, bundles, uiContent }
 */
function loadSurveyConfig() {
  ensureSpreadsheetId_();
  var cacheKey = surveyConfigCacheKey_();
  var cached = cacheJsonGet_(cacheKey);
  if (cached) return cached;
  var payload = buildSurveyConfigPayload_();
  cacheJsonPut_(cacheKey, payload, SURVEY_CONFIG_CACHE_TTL_SEC_);
  return payload;
}

function buildSurveySubmitPostprocessOptions_(settings) {
  var s = settings || {};
  return {
    request_id: "",
    update_duplicate_count: true,
    ensure_checklist: ynToBool_(s.enable_request_tasks, true),
    notify_on_submit: ynToBool_(s.slack_notify_on_submit, true)
  };
}

function buildSurveySubmitPostprocessPayload_(requestObj, estimate, settings) {
  var payload = buildSurveySubmitPostprocessOptions_(settings);
  payload.request_id = String(requestObj && requestObj.request_id || "").trim();
  payload.estimate = {
    min: toNumber_(estimate && estimate.min, 0),
    max: toNumber_(estimate && estimate.max, 0),
    flags: Array.isArray(estimate && estimate.flags) ? estimate.flags : []
  };
  return payload;
}

function surveySubmitPostprocessTriggerCacheKey_() {
  return [SURVEY_SUBMIT_POSTPROCESS_TRIGGER_PREFIX_, getSpreadsheetId_()].join("_");
}

function scheduleSurveySubmitPostprocess_() {
  var cacheKey = surveySubmitPostprocessTriggerCacheKey_();
  var cache = CacheService.getScriptCache();
  try {
    if (cache.get(cacheKey)) return false;
  } catch (e) {}
  try {
    cache.put(
      cacheKey,
      nowIso_(),
      Math.max(Math.ceil(SURVEY_SUBMIT_POSTPROCESS_TRIGGER_DELAY_MS_ / 1000), 30)
    );
  } catch (e2) {}
  try {
    ScriptApp.newTrigger(SURVEY_SUBMIT_POSTPROCESS_TRIGGER_HANDLER_)
      .timeBased()
      .after(SURVEY_SUBMIT_POSTPROCESS_TRIGGER_DELAY_MS_)
      .create();
    return true;
  } catch (e3) {
    try { cache.remove(cacheKey); } catch (e4) {}
    throw e3;
  }
}

function enqueueSurveySubmitPostprocess_(requestObj, estimate, settings) {
  ensureSurveySubmitPostprocessSchema_();
  var now = nowIso_();
  var payload = buildSurveySubmitPostprocessPayload_(requestObj, estimate, settings);
  scheduleSurveySubmitPostprocess_();
  appendRow_(REQUEST_POSTPROCESS_QUEUE_SHEET_, {
    job_id: "SPQ_" + token_().slice(0, 20).toUpperCase(),
    request_id: payload.request_id,
    job_type: "SURVEY_SUBMIT",
    status: "QUEUED",
    attempt_count: 0,
    created_at: now,
    scheduled_at: now,
    started_at: "",
    processed_at: "",
    last_error: "",
    payload_json: JSON.stringify(payload)
  });
  return payload;
}

function runSurveySubmitPostprocess_(requestId, estimate, options) {
  var rid = String(requestId || "").trim();
  if (!rid) return { success: false, skipped: true, reason: "missing_request_id" };

  var found = findRequestRowById_(rid);
  if (!found) return { success: false, skipped: true, reason: "request_not_found" };

  var opts = options || {};
  var requestRow = found.data || {};
  var patch = { updated_at: nowIso_() };
  var hasPatch = false;

  if (opts.update_duplicate_count !== false) {
    patch.duplicate_customer_count = countDuplicateRequests_(
      String(requestRow.duplicate_phone_key || "").trim() || normalizePhoneKey_(requestRow.contact_phone),
      String(requestRow.duplicate_email_key || "").trim() || normalizeEmailKey_(requestRow.contact_email),
      requestRow.customer_name,
      rid
    );
    requestRow.duplicate_customer_count = patch.duplicate_customer_count;
    hasPatch = true;
  }

  if (hasPatch) {
    updateRowFields_(found.sheet, found.rowNo, found.meta, patch);
    requestRow.updated_at = patch.updated_at;
  }

  if (opts.ensure_checklist) {
    ensureChecklistItemsForRequest_(requestRow);
  }
  if (opts.notify_on_submit) {
    notifyNewRequestSlack_(requestRow, estimate || buildStoredEstimateFromRequest_(requestRow));
  }

  return {
    success: true,
    request_id: rid,
    duplicate_customer_count: requestRow.duplicate_customer_count || 0
  };
}

function getQueuedSurveySubmitPostprocessJobs_(limit) {
  return readAllRows_(REQUEST_POSTPROCESS_QUEUE_SHEET_).filter(function(row) {
    var status = String(row.status || "QUEUED").trim().toUpperCase();
    return String(row.job_type || "").trim().toUpperCase() === "SURVEY_SUBMIT"
      && (status === "QUEUED" || status === "RETRY");
  }).sort(function(a, b) {
    return String(a.created_at || "").localeCompare(String(b.created_at || ""))
      || Number(a._rowNo || 0) - Number(b._rowNo || 0);
  }).slice(0, Math.max(Number(limit || SURVEY_SUBMIT_POSTPROCESS_BATCH_SIZE_), 1));
}

function runSurveySubmitPostprocessQueue_() {
  ensureSpreadsheetId_();
  try {
    CacheService.getScriptCache().remove(surveySubmitPostprocessTriggerCacheKey_());
  } catch (e0) {}

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    scheduleSurveySubmitPostprocess_();
    return { success: false, locked: true };
  }

  var processedCount = 0;
  var errorCount = 0;
  try {
    ensureSurveySubmitPostprocessSchema_();
    var jobs = getQueuedSurveySubmitPostprocessJobs_(SURVEY_SUBMIT_POSTPROCESS_BATCH_SIZE_);
    var queueSheet = getSheet_(REQUEST_POSTPROCESS_QUEUE_SHEET_);
    var queueMeta = sheetMeta_(queueSheet);

    for (var i = 0; i < jobs.length; i++) {
      var job = jobs[i] || {};
      var rowNo = Number(job._rowNo || 0);
      if (rowNo < 2) continue;
      var attemptCount = Math.max(toNumber_(job.attempt_count, 0), 0) + 1;
      updateRowFields_(queueSheet, rowNo, queueMeta, {
        status: "RUNNING",
        attempt_count: attemptCount,
        started_at: nowIso_(),
        last_error: ""
      });

      try {
        var payload = safeJsonParse_(job.payload_json, {}) || {};
        runSurveySubmitPostprocess_(payload.request_id || job.request_id, payload.estimate || null, payload);
        updateRowFields_(queueSheet, rowNo, queueMeta, {
          status: "DONE",
          processed_at: nowIso_(),
          last_error: ""
        });
        processedCount++;
      } catch (jobError) {
        errorCount++;
        updateRowFields_(queueSheet, rowNo, queueMeta, {
          status: attemptCount >= SURVEY_SUBMIT_POSTPROCESS_MAX_ATTEMPTS_ ? "FAILED" : "RETRY",
          processed_at: attemptCount >= SURVEY_SUBMIT_POSTPROCESS_MAX_ATTEMPTS_ ? nowIso_() : "",
          last_error: String(jobError && jobError.message ? jobError.message : jobError)
        });
      }
    }

    if (getQueuedSurveySubmitPostprocessJobs_(1).length) {
      scheduleSurveySubmitPostprocess_();
    }

    return { success: errorCount === 0, processed_count: processedCount, error_count: errorCount };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
//  SURVEY SUBMISSION + ESTIMATE CALCULATION
// ============================================================

/**
 * Submit completed survey and calculate pre-quote estimate.
 * @param {Object} payload - { answers: {code: value}, contact: {name, phone, email, method, note}, files: [] }
 * @returns {Object} { request_id, share_token, estimate, recommendations, flags }
 */
function submitSurvey(payload) {
  ensureSpreadsheetId_();
  ensureSurveySubmitStorageSchema_();
  var startMs = Date.now();
  var error = null;
  var requestId = "";
  try {
    if (!payload || typeof payload !== "object") throw new Error("Invalid payload");
    var answers = payload.answers || {};
    var contact = payload.contact || {};
    var files = Array.isArray(payload.files) ? payload.files : [];
    var settings = getSettings_();

    var name = String(contact.name || "").trim();
    var phone = String(contact.phone || "").trim();
    if (!name) throw new Error("이름을 입력해주세요.");
    if (!phone) throw new Error("연락처를 입력해주세요.");

    requestId = "PQ_" + uuid_().replace(/-/g, "").slice(0, 16).toUpperCase();
    var shareToken = token_().slice(0, 32);
    var now = nowIso_();
    var projectType = String(answers.Q000_PROJECT_TYPE || "").toUpperCase();
    var flowMode = String(answers.Q000_FLOW_MODE || "").toUpperCase();
    var tags = calculateTags_(answers);
    var estimate = calculateEstimate_(answers, projectType);
    var recommendations = [];
    var asyncPostprocess = ynToBool_(settings.survey_submit_async_postprocess, true);
    var postprocessOptions = buildSurveySubmitPostprocessOptions_(settings);
    var duplicatePhoneKey = normalizePhoneKey_(phone);
    var duplicateEmailKey = normalizeEmailKey_(contact.email);
    var requestObj = {
      request_id: requestId,
      created_at: now,
      updated_at: now,
      source_channel: "WEB_SURVEY",
      customer_name: name,
      contact_phone: phone,
      contact_email: String(contact.email || "").trim(),
      preferred_contact_method: String(contact.method || settings.default_contact_method || "KAKAO").toUpperCase(),
      project_type: projectType,
      housing_type: answers.R001_HOUSING_TYPE || answers.C001_BIZ_TYPE || "",
      property_status: "",
      address_text: String(contact.address || answers.Q906_ADDRESS || "").trim(),
      area_py: answers.R002_AREA || answers.C002_AREA || "",
      area_m2: "",
      scope_level: answers.R011_SCOPE_LEVEL || answers.C011_SCOPE_LEVEL || "",
      flow_mode: flowMode,
      preset_bundle_id: answers.Q000_PRESET_RESI || answers.Q000_PRESET_COMM || "",
      tags_json: JSON.stringify(tags),
      estimate_min: estimate.min || 0,
      estimate_max: estimate.max || 0,
      estimate_json: JSON.stringify(estimate),
      risk_flags_json: JSON.stringify(estimate.flags || []),
      recommendation_count: recommendations.length,
      final_estimate_min: estimate.min || 0,
      final_estimate_max: estimate.max || 0,
      final_estimate_source: "AUTO",
      final_recommendation_count: recommendations.length,
      status: "NEW",
      priority: "NORMAL",
      work_state: "NEW",
      survey_version: settings.survey_version || "SURVEY_V6",
      share_token: shareToken,
      customer_note: String(contact.note || "").trim(),
      submitted_at: now,
      schedule_start: answers.R320_START || answers.C320_START || "",
      schedule_end: answers.R321_END || answers.C321_END || "",
      schedule_visit: answers.R322_VISIT || answers.C322_VISIT || "",
      duplicate_phone_key: duplicatePhoneKey,
      duplicate_email_key: duplicateEmailKey,
      duplicate_customer_count: asyncPostprocess ? "" : countDuplicateRequests_(duplicatePhoneKey, duplicateEmailKey, name, requestId),
      review_status_label: REVIEW_STATUS_LABEL_MAP_.AUTO,
      quote_draft_status: "PENDING",
      answers_snapshot_json: JSON.stringify(answers)
    };
    appendRow_("Requests", requestObj);

    var questionCodes = Object.keys(answers);
    if (questionCodes.length > 0) {
      var answerRows = [];
      for (var i = 0; i < questionCodes.length; i++) {
        var code = questionCodes[i];
        var val = answers[code];
        answerRows.push({
          request_id: requestId,
          answer_id: requestId + "_A" + String(i + 1).padStart(3, "0"),
          question_id: code,
          question_code: code,
          answer_type: Array.isArray(val) ? "MULTI" : "SINGLE",
          answer_value_text: Array.isArray(val) ? val.join(",") : String(val || ""),
          answer_value_number: Array.isArray(val) ? "" : (isFinite(Number(val)) ? Number(val) : ""),
          answer_value_json: JSON.stringify(val === undefined ? "" : val),
          normalized_answer: Array.isArray(val) ? val.join(",") : String(val || ""),
          tags_applied_json: "",
          score_delta: "",
          created_at: now
        });
      }
      appendRows_("RequestAnswers", answerRows);
    }

    if (files.length > 0) {
      var fileLimit = Math.max(Number(settings.file_upload_max_count || 10), 1);
      var fileRows = [];
      for (var f = 0; f < Math.min(files.length, fileLimit); f++) {
        fileRows.push(buildRequestFileRow_(requestId, files[f], f, now));
      }
      appendRows_("RequestFiles", fileRows);
    }

    appendRow_("RequestTimeline", {
      event_id: uuid_(),
      request_id: requestId,
      event_at: now,
      actor_type: "CUSTOMER",
      actor_id: name,
      event_type: "SUBMITTED",
      from_status: "",
      to_status: "NEW",
      message: "설문 접수 완료",
      payload_json: "",
      event_label: "설문 접수",
      dedup_key: "request:" + requestId + ":submitted"
    });

    if (asyncPostprocess) {
      try {
        enqueueSurveySubmitPostprocess_(requestObj, estimate, settings);
      } catch (queueError) {
        Logger.log("Survey submit postprocess queue fallback: " + queueError);
        runSurveySubmitPostprocess_(requestId, estimate, postprocessOptions);
      }
    } else {
      runSurveySubmitPostprocess_(requestId, estimate, postprocessOptions);
    }

    cacheResultBootstrap_(requestObj, estimate);

    return {
      request_id: requestId,
      share_token: shareToken,
      estimate: estimate,
      recommendations: recommendations.slice(0, 12),
      flags: estimate.flags || []
    };
  } catch (e) {
    error = e;
    throw e;
  } finally {
    logPerfMetricSafe_(
      "submit_survey",
      Date.now() - startMs,
      requestId,
      {},
      !error,
      error ? String(error) : "",
      "CUSTOMER",
      requestId
    );
  }
}

// ============================================================
//  TAG CALCULATION ENGINE
// ============================================================

function calculateTags_(answers) {
  var rules = readStaticRowsCached_("SurveyTagRules").filter(function(r) {
    return ynToBool_(r.is_active);
  });

  var tagMap = {}; // { "style:MODERN": { type, value, totalWeight } }

  for (var i = 0; i < rules.length; i++) {
    var rule = rules[i];
    var qCode = String(rule.question_id || "").trim();
    var optCode = String(rule.option_code || "").trim();
    var answerVal = answers[qCode];
    if (!answerVal) continue;

    var match = false;
    if (Array.isArray(answerVal)) {
      match = answerVal.indexOf(optCode) >= 0;
    } else {
      match = String(answerVal) === optCode;
    }

    if (match) {
      var tagKey = String(rule.tag_type || "") + ":" + String(rule.tag_value || "");
      if (!tagMap[tagKey]) {
        tagMap[tagKey] = { type: rule.tag_type, value: rule.tag_value, weight: 0 };
      }
      tagMap[tagKey].weight += Number(rule.weight || 1);
    }
  }

  // Convert to array sorted by weight desc
  var tags = [];
  for (var k in tagMap) {
    if (Object.prototype.hasOwnProperty.call(tagMap, k)) {
      tags.push(tagMap[k]);
    }
  }
  tags.sort(function(a, b) { return b.weight - a.weight; });
  return tags;
}

// ============================================================
//  ESTIMATE CALCULATION ENGINE
// ============================================================

function calculateEstimate_(answers, projectType) {
  var rules = readStaticRowsCached_("EstimateRules").filter(function(r) {
    return ynToBool_(r.is_active);
  });

  // Sort by priority
  rules.sort(function(a, b) { return Number(a.priority || 0) - Number(b.priority || 0); });

  var result = { min: 0, max: 0, base_min: 0, base_max: 0, adjustments: [], flags: [] };

  // Phase 1: BASE_RANGE — find matching base
  var baseRules = rules.filter(function(r) { return r.rule_type === "BASE_RANGE"; });
  var bestBase = null;
  var bestBaseScore = -1;

  for (var i = 0; i < baseRules.length; i++) {
    var br = baseRules[i];
    var score = scoreBaseRule_(br, answers, projectType);
    if (score > bestBaseScore) {
      bestBaseScore = score;
      bestBase = br;
    }
  }

  if (bestBase) {
    var baseVal = safeJsonParse_(bestBase.adjustment_value, {});
    result.base_min = Number(baseVal.min || 0);
    result.base_max = Number(baseVal.max || 0);
    result.min = result.base_min;
    result.max = result.base_max;
    result.adjustments.push({
      rule_id: bestBase.rule_id,
      label: bestBase.display_label || "기본 범위",
      type: "BASE",
      min: result.base_min,
      max: result.base_max
    });
  }

  // Phase 2: RANGE_FACTOR — multiply
  var factorRules = rules.filter(function(r) { return r.rule_type === "RANGE_FACTOR"; });
  for (var f = 0; f < factorRules.length; f++) {
    var fr = factorRules[f];
    if (!matchCondition_(fr.condition_json, answers)) continue;
    var fv = safeJsonParse_(fr.adjustment_value, {});
    var minF = Number(fv.min_factor || 1);
    var maxF = Number(fv.max_factor || 1);
    result.min = Math.round(result.min * minF);
    result.max = Math.round(result.max * maxF);
    result.adjustments.push({
      rule_id: fr.rule_id,
      label: fr.display_label || "",
      type: "FACTOR",
      min_factor: minF,
      max_factor: maxF
    });
  }

  // Phase 3: RANGE_ADD — add flat amounts
  var addRules = rules.filter(function(r) { return r.rule_type === "RANGE_ADD"; });
  for (var a = 0; a < addRules.length; a++) {
    var ar = addRules[a];
    if (!matchCondition_(ar.condition_json, answers)) continue;
    var av = safeJsonParse_(ar.adjustment_value, {});
    var addMin = Number(av.min || 0);
    var addMax = Number(av.max || 0);
    result.min += addMin;
    result.max += addMax;
    result.adjustments.push({
      rule_id: ar.rule_id,
      label: ar.display_label || "",
      type: "ADD",
      min: addMin,
      max: addMax
    });
  }

  // Phase 4: REVIEW_FLAG — collect flags
  var flagRules = rules.filter(function(r) { return r.rule_type === "REVIEW_FLAG"; });
  for (var fl = 0; fl < flagRules.length; fl++) {
    var flr = flagRules[fl];
    if (!matchCondition_(flr.condition_json, answers)) continue;
    var flagVal = safeJsonParse_(flr.adjustment_value, {});
    result.flags.push({
      rule_id: flr.rule_id,
      flag: flagVal.flag || "",
      severity: flagVal.severity || "medium",
      label: flr.display_label || ""
    });
  }

  return result;
}

function scoreBaseRule_(rule, answers, projectType) {
  // Match housing_type
  var rHousing = String(rule.housing_type || "").toUpperCase().trim();
  var aHousing = String(answers.R001_HOUSING_TYPE || answers.C001_BIZ_TYPE || "").toUpperCase().trim();
  if (projectType === "COMMERCIAL") aHousing = String(answers.C001_BIZ_TYPE || "COMMERCIAL").toUpperCase();

  if (rHousing && aHousing && rHousing !== aHousing) return -1;

  // Match area_band
  var rArea = String(rule.area_band || "").toUpperCase().trim();
  var aArea = String(answers.R002_AREA || answers.C002_AREA || "").toUpperCase().trim();
  if (rArea && aArea && rArea !== aArea) return -1;

  // Match condition_json
  var cond = safeJsonParse_(rule.condition_json, {});
  var score = 0;
  for (var key in cond) {
    if (!Object.prototype.hasOwnProperty.call(cond, key)) continue;
    var expected = cond[key];
    var actual = answers[key];
    if (!actual) return -1;
    var actualArr = Array.isArray(actual) ? actual : [String(actual)];
    var expectedArr = Array.isArray(expected) ? expected : [String(expected)];
    var found = false;
    for (var e = 0; e < expectedArr.length; e++) {
      if (actualArr.indexOf(expectedArr[e]) >= 0) { found = true; break; }
    }
    if (!found) return -1;
    score++;
  }

  // Bonus for specificity
  if (rHousing && rHousing === aHousing) score += 2;
  if (rArea && rArea === aArea) score += 2;
  return score;
}

function matchCondition_(conditionJson, answers) {
  var cond = safeJsonParse_(conditionJson, {});
  if (!cond || typeof cond !== "object") return false;
  var keys = Object.keys(cond);
  if (keys.length === 0) return false;

  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var expected = cond[key];
    var actual = answers[key];
    if (!actual && actual !== 0) return false;

    var actualArr = Array.isArray(actual) ? actual : [String(actual)];
    var expectedArr = Array.isArray(expected) ? expected : [String(expected)];

    // ANY match (OR within same key)
    var matched = false;
    for (var e = 0; e < expectedArr.length; e++) {
      if (actualArr.indexOf(expectedArr[e]) >= 0) { matched = true; break; }
    }
    if (!matched) return false;
  }
  return true; // All keys matched (AND across keys)
}

// ============================================================
//  RECOMMENDATION ENGINE
// ============================================================

var SURVEY_TRADE_TO_QUOTE_TRADE_MAP_ = {
  R_FLOOR: ["FLOOR"],
  C_FLOOR: ["FLOOR"],
  R_WALL: ["WALL"],
  R_TILE: ["WALL", "FLOOR"],
  R_BATH: ["WALL", "FLOOR"],
  R_KITCHEN: ["WALL", "FLOOR"],
  C_KITCHEN: ["WALL", "FLOOR"],
  C_RESTROOM: ["WALL", "FLOOR"]
};

function getQuoteBaseDbId_() {
  var settings = getSettings_();
  var props = PropertiesService.getScriptProperties();
  var id = String(settings.base_db_spreadsheet_id || props.getProperty("BASE_DB_SPREADSHEET_ID") || "").trim();
  if (!id) {
    throw new Error("Quote DB spreadsheet ID is missing. Set Settings.base_db_spreadsheet_id or ScriptProperties.BASE_DB_SPREADSHEET_ID.");
  }
  return id;
}

function getQuoteMasterBaseUrl_() {
  var settings = getSettings_();
  return normalizeUrl_(settings.quote_master_app_url || "");
}

function buildQuoteImageUrl_(fileId) {
  var id = String(fileId || "").trim();
  if (!id) return "";
  var base = getQuoteMasterBaseUrl_();
  if (!base) return "";
  return base + "?page=img&fileId=" + encodeURIComponent(id);
}

function parseTagSummaryMap_(summary) {
  var out = Object.create(null);
  var raw = String(summary || "").trim();
  if (!raw) return out;
  raw.split(",").forEach(function(token) {
    var chunk = String(token || "").trim();
    if (!chunk) return;
    var idx = chunk.indexOf(":");
    if (idx < 0) return;
    var type = String(chunk.slice(0, idx) || "").trim().toLowerCase();
    var value = String(chunk.slice(idx + 1) || "").trim();
    if (!type || !value) return;
    if (!out[type]) out[type] = [];
    pushUnique_(out[type], value);
  });
  return out;
}

function mergeTagMaps_(primary, secondary) {
  var out = Object.create(null);
  var sources = [primary || {}, secondary || {}];
  for (var s = 0; s < sources.length; s++) {
    var src = sources[s];
    for (var key in src) {
      if (!Object.prototype.hasOwnProperty.call(src, key)) continue;
      if (!out[key]) out[key] = [];
      var values = Array.isArray(src[key]) ? src[key] : [src[key]];
      for (var i = 0; i < values.length; i++) pushUnique_(out[key], values[i]);
    }
  }
  return out;
}

function splitTemplateSummaryTags_(raw) {
  var source = String(raw || "").trim();
  if (!source) return [];
  return toUniqueStringArray_(source.split(","));
}

function buildRequestRecommendationRow_(requestId, rec, rank, createdAt) {
  var item = rec || {};
  var snapshot = JSON.parse(JSON.stringify(item || {}));
  var now = createdAt || nowIso_();
  var sourceRefId = item.source_ref_id || item.material_id || item.template_id || "";
  var sourceRefKey = item.source_ref_key || item.trade_code || item.material_type || "";
  return {
    request_id: requestId,
    rec_id: item.rec_id || (requestId + "_REC" + String(rank).padStart(3, "0")),
    rec_type: item.type || item.rec_type || "MATERIAL",
    rank: rank,
    template_id: item.template_id || "",
    material_id: item.material_id || "",
    trade_code: item.trade_code || "",
    material_type: item.material_type || "",
    title: item.title || "",
    subtitle: item.subtitle || "",
    reason_text: item.reason_text || item.reason || "",
    score: toNumber_(item.score, 0),
    price_hint_min: toNumber_(item.price_min !== undefined ? item.price_min : item.price_hint_min, 0),
    price_hint_max: toNumber_(item.price_max !== undefined ? item.price_max : item.price_hint_max, 0),
    matched_tags_json: JSON.stringify(item.matched_tags || {}),
    snapshot_json: JSON.stringify(snapshot),
    is_selected: "",
    created_at: now,
    is_visible: item.is_visible === false ? "N" : "Y",
    is_deleted: item.is_deleted ? "Y" : "N",
    is_manual: item.is_manual ? "Y" : "N",
    image_url: item.image_url || "",
    image_file_id: item.image_file_id || "",
    brand: item.brand || "",
    spec: item.spec || "",
    source_type: item.source_type || item.source || (item.is_manual ? "MANUAL" : "AUTO"),
    source_ref_id: sourceRefId,
    source_ref_key: sourceRefKey,
    review_id: item.review_id || "",
    updated_at: item.updated_at || now,
    hidden_at: item.is_visible === false ? (item.hidden_at || now) : "",
    deleted_at: item.is_deleted ? (item.deleted_at || now) : "",
    note: item.note || ""
  };
}

function normalizeRequestRecommendationRow_(row) {
  var snapshot = safeJsonParse_(row.snapshot_json, {});
  return {
    request_id: row.request_id,
    rec_id: row.rec_id,
    rec_type: row.rec_type || snapshot.type || snapshot.rec_type || "",
    rank: toNumber_(row.rank || snapshot.rank, 0),
    template_id: row.template_id || snapshot.template_id || "",
    material_id: row.material_id || snapshot.material_id || "",
    trade_code: row.trade_code || snapshot.trade_code || "",
    material_type: row.material_type || snapshot.material_type || "",
    title: row.title || snapshot.title || "",
    subtitle: row.subtitle || snapshot.subtitle || "",
    reason_text: row.reason_text || snapshot.reason_text || snapshot.reason || "",
    score: toNumber_(row.score || snapshot.score, 0),
    price_hint_min: toNumber_(row.price_hint_min || snapshot.price_hint_min || snapshot.price_min, 0),
    price_hint_max: toNumber_(row.price_hint_max || snapshot.price_hint_max || snapshot.price_max, 0),
    matched_tags: safeJsonParse_(row.matched_tags_json, snapshot.matched_tags || {}),
    image_url: row.image_url || snapshot.image_url || "",
    image_file_id: row.image_file_id || snapshot.image_file_id || "",
    brand: row.brand || snapshot.brand || "",
    spec: row.spec || snapshot.spec || "",
    source_type: String(row.source_type || snapshot.source_type || snapshot.source || "").trim().toUpperCase(),
    source_ref_id: row.source_ref_id || snapshot.source_ref_id || "",
    source_ref_key: row.source_ref_key || snapshot.source_ref_key || "",
    is_visible: ynToBool_(row.is_visible, true),
    is_deleted: ynToBool_(row.is_deleted, false),
    is_manual: ynToBool_(row.is_manual, false),
    review_id: row.review_id || snapshot.review_id || "",
    updated_at: row.updated_at || snapshot.updated_at || row.created_at || snapshot.created_at || "",
    hidden_at: row.hidden_at || snapshot.hidden_at || "",
    deleted_at: row.deleted_at || snapshot.deleted_at || "",
    note: row.note || snapshot.note || "",
    created_at: row.created_at || snapshot.created_at || ""
  };
}

function buildRequestFileRow_(requestId, file, index, createdAt) {
  var item = file || {};
  var sizeValue = item.file_size !== undefined ? item.file_size : item.size;
  return {
    request_id: requestId,
    file_id: requestId + "_FILE" + String(index + 1).padStart(3, "0"),
    file_name: String(item.file_name || item.name || "").trim(),
    file_mime: String(item.file_mime || item.type || "").trim(),
    file_size: isFinite(Number(sizeValue)) ? Number(sizeValue) : "",
    drive_file_id: String(item.drive_file_id || "").trim(),
    file_url: String(item.file_url || "").trim(),
    upload_type: String(item.upload_type || "CLIENT_METADATA_ONLY").trim(),
    sort_order: index + 1,
    created_at: createdAt || nowIso_(),
    note: String(item.note || "").trim()
  };
}

function normalizeMaterialRecommendationRecord_(row, sourceType) {
  var tagsByType = mergeTagMaps_(
    safeJsonParse_(row.tags_by_type_json, {}),
    parseTagSummaryMap_(row.tags_summary)
  );
  var tradeCodes = [];
  pushUnique_(tradeCodes, row.trade_code);
  if (tagsByType.trade) tradeCodes = toUniqueStringArray_(tradeCodes.concat(tagsByType.trade));
  return {
    source_type: sourceType || "CACHE",
    material_id: String(row.material_id || "").trim(),
    name: String(row.name || "").trim(),
    brand: String(row.brand || "").trim(),
    spec: String(row.spec || "").trim(),
    unit: String(row.unit || "").trim(),
    unit_price: toNumber_(row.unit_price, 0),
    image_file_id: String(row.image_file_id || "").trim(),
    image_url: String(row.image_url || "").trim() || buildQuoteImageUrl_(row.image_file_id),
    is_active: ynToBool_(row.is_active, true),
    expose_to_prequote: ynToBool_(row.expose_to_prequote, true),
    is_representative: ynToBool_(row.is_representative, false),
    material_type: String(row.material_type || "").trim(),
    trade_code: String(row.trade_code || "").trim(),
    space_type: String(row.space_type || "").trim(),
    effective_trade_codes: tradeCodes,
    price_band: String(row.price_band || "").trim(),
    recommendation_score_base: toNumber_(row.recommendation_score_base, 0),
    sort_order: toNumber_(row.sort_order, 0),
    tags_summary: String(row.tags_summary || "").trim(),
    tags_by_type: tagsByType,
    legacy_group_keys: safeJsonParse_(row.legacy_group_keys_json, []),
    recommendation_note: String(row.recommendation_note || "").trim()
  };
}

function normalizeTemplateRecommendationRecord_(row, sourceType) {
  return {
    source_type: sourceType || "CACHE",
    template_id: String(row.template_id || "").trim(),
    category: String(row.category || "").trim(),
    template_name: String(row.template_name || "").trim(),
    latest_version: row.latest_version || "",
    latest_item_count: toNumber_(row.latest_item_count, 0),
    is_active: ynToBool_(row.is_active, true),
    expose_to_prequote: ynToBool_(row.expose_to_prequote, true),
    template_type: String(row.template_type || "").trim(),
    housing_type: String(row.housing_type || "").trim(),
    area_band: String(row.area_band || "").trim(),
    budget_band: String(row.budget_band || "").trim(),
    style_tags_summary: String(row.style_tags_summary || "").trim(),
    tone_tags_summary: String(row.tone_tags_summary || "").trim(),
    trade_scope_summary: String(row.trade_scope_summary || "").trim(),
    prequote_priority: toNumber_(row.prequote_priority, 0),
    sort_order: toNumber_(row.sort_order, 0),
    recommendation_note: String(row.recommendation_note || "").trim(),
    target_customer_summary: String(row.target_customer_summary || "").trim(),
    recommended_for_summary: String(row.recommended_for_summary || "").trim(),
    is_featured_prequote: ynToBool_(row.is_featured_prequote, false)
  };
}

function loadLocalMaterialCache_() {
  return readAllRows_("MaterialsCache")
    .filter(function(row) { return String(row.material_id || "").trim(); })
    .map(function(row) { return normalizeMaterialRecommendationRecord_(row, "CACHE"); });
}

function loadLocalTemplateCache_() {
  return readAllRows_("TemplatesCache")
    .filter(function(row) { return String(row.template_id || "").trim(); })
    .map(function(row) { return normalizeTemplateRecommendationRecord_(row, "CACHE"); });
}

function readQuoteMaterialsForRecommendation_() {
  var ctx = getQuoteBaseContext_();
  if (typeof BaseLib !== "undefined" && BaseLib) {
    try {
      if (typeof BaseLib.getRepresentativeMaterialsSnapshotForPrequote === "function") {
        var snapshot = BaseLib.getRepresentativeMaterialsSnapshotForPrequote(ctx, {}, {
          include_inactive: false,
          include_unexposed: true,
          limit: 50000
        });
        if (snapshot && Array.isArray(snapshot.items)) return snapshot.items;
      } else if (typeof BaseLib.listRepresentativeMaterialsForPrequote === "function") {
        var list = BaseLib.listRepresentativeMaterialsForPrequote(ctx, {
          include_inactive: false,
          include_unexposed: true,
          limit: 50000
        }, {
          include_inactive: false,
          include_unexposed: true,
          limit: 50000
        });
        if (Array.isArray(list) && list.length) return list;
      }
    } catch (e0) {
      Logger.log("Material snapshot/list fallback to raw: " + e0);
    }
  }
  var baseDbId = getQuoteBaseDbId_();
  var materials = readAllRows_("Materials", baseDbId);
  var tags = [];
  var groups = [];
  try { tags = readAllRows_("MaterialTags", baseDbId); } catch (e) { tags = []; }
  try { groups = readAllRows_("MaterialGroups", baseDbId); } catch (e2) { groups = []; }

  var tagsByMaterial = Object.create(null);
  for (var i = 0; i < tags.length; i++) {
    var tagRow = tags[i];
    if (!ynToBool_(tagRow.is_active, true)) continue;
    var materialId = String(tagRow.material_id || "").trim();
    if (!materialId) continue;
    if (!tagsByMaterial[materialId]) {
      tagsByMaterial[materialId] = { tags_by_type: Object.create(null) };
    }
    var tagType = String(tagRow.tag_type || "").trim().toLowerCase();
    var tagValue = String(tagRow.tag_value || "").trim();
    if (!tagType || !tagValue) continue;
    if (!tagsByMaterial[materialId].tags_by_type[tagType]) tagsByMaterial[materialId].tags_by_type[tagType] = [];
    pushUnique_(tagsByMaterial[materialId].tags_by_type[tagType], tagValue);
  }

  var groupsByMaterial = Object.create(null);
  for (var g = 0; g < groups.length; g++) {
    var groupRow = groups[g];
    if (!ynToBool_(groupRow.is_active, true)) continue;
    var groupMaterialId = String(groupRow.material_id || "").trim();
    var groupKey = String(groupRow.group_key || "").trim();
    if (!groupMaterialId || !groupKey) continue;
    if (!groupsByMaterial[groupMaterialId]) groupsByMaterial[groupMaterialId] = [];
    pushUnique_(groupsByMaterial[groupMaterialId], groupKey);
  }

  var out = [];
  for (var m = 0; m < materials.length; m++) {
    var row = materials[m];
    if (!ynToBool_(row.is_active, true)) continue;
    var summaryTags = parseTagSummaryMap_(row.tags_summary);
    var materialTags = tagsByMaterial[row.material_id] ? tagsByMaterial[row.material_id].tags_by_type : {};
    var mergedTags = mergeTagMaps_(materialTags, summaryTags);
    var tradeCodes = [];
    pushUnique_(tradeCodes, row.trade_code);
    if (mergedTags.trade) tradeCodes = toUniqueStringArray_(tradeCodes.concat(mergedTags.trade));
    out.push({
      material_id: String(row.material_id || "").trim(),
      name: String(row.name || "").trim(),
      brand: String(row.brand || "").trim(),
      spec: String(row.spec || "").trim(),
      unit: String(row.unit || "").trim(),
      unit_price: toNumber_(row.unit_price, 0),
      image_file_id: String(row.image_file_id || "").trim(),
      image_url: buildQuoteImageUrl_(row.image_file_id),
      is_active: ynToBool_(row.is_active, true),
      expose_to_prequote: ynToBool_(row.expose_to_prequote, true),
      is_representative: ynToBool_(row.is_representative, false),
      material_type: String(row.material_type || "").trim(),
      trade_code: String(row.trade_code || "").trim(),
      space_type: String(row.space_type || "").trim(),
      effective_trade_codes: tradeCodes,
      price_band: String(row.price_band || "").trim(),
      recommendation_score_base: toNumber_(row.recommendation_score_base, 0),
      sort_order: toNumber_(row.sort_order, 0),
      tags_summary: String(row.tags_summary || "").trim(),
      tags_by_type: mergedTags,
      legacy_group_keys: groupsByMaterial[row.material_id] || [],
      recommendation_note: String(row.recommendation_note || "").trim()
    });
  }
  return out;
}

function readQuoteTemplatesForRecommendation_() {
  var ctx = getQuoteBaseContext_();
  if (typeof BaseLib !== "undefined" && BaseLib) {
    try {
      if (typeof BaseLib.getTemplateCatalogSnapshotForPrequote === "function") {
        var snapshot = BaseLib.getTemplateCatalogSnapshotForPrequote(ctx, {}, {
          include_inactive: true,
          include_unexposed: true,
          limit: 50000
        });
        if (snapshot && Array.isArray(snapshot.items)) return snapshot.items;
      } else if (typeof BaseLib.listTemplateCatalogForPrequote === "function") {
        var list = BaseLib.listTemplateCatalogForPrequote(ctx, {
          include_inactive: true,
          include_unexposed: true,
          limit: 50000
        }, {
          include_inactive: true,
          include_unexposed: true,
          limit: 50000
        });
        if (Array.isArray(list) && list.length) return list;
      }
    } catch (e0) {
      Logger.log("Template snapshot/list fallback to raw: " + e0);
    }
  }
  var baseDbId = getQuoteBaseDbId_();
  var rows = [];
  try {
    rows = readAllRows_("TemplateCatalog", baseDbId);
  } catch (e) {
    rows = [];
  }
  return rows
    .filter(function(row) { return String(row.template_id || "").trim(); })
    .map(function(row) { return normalizeTemplateRecommendationRecord_(row, "QUOTE_DB"); });
}

function loadMaterialsForRecommendations_() {
  var cacheRows = loadLocalMaterialCache_();
  if (cacheRows.length) return cacheRows;
  try {
    return readQuoteMaterialsForRecommendation_();
  } catch (e) {
    Logger.log("Material recommendation source unavailable: " + e);
    return [];
  }
}

function loadTemplatesForRecommendations_() {
  var cacheRows = loadLocalTemplateCache_();
  if (cacheRows.length) return cacheRows;
  try {
    return readQuoteTemplatesForRecommendation_();
  } catch (e) {
    Logger.log("Template recommendation source unavailable: " + e);
    return [];
  }
}

function mapSurveyTradeToQuoteTradeCodes_(tradeCode) {
  var code = String(tradeCode || "").trim().toUpperCase();
  if (!code) return [];
  if (SURVEY_TRADE_TO_QUOTE_TRADE_MAP_[code]) return SURVEY_TRADE_TO_QUOTE_TRADE_MAP_[code].slice();
  if (code.indexOf("FLOOR") >= 0) return ["FLOOR"];
  if (code.indexOf("WALL") >= 0 || code.indexOf("TILE") >= 0) return ["WALL", "FLOOR"];
  return [];
}

function collectRequestedQuoteTradeCodes_(answers) {
  var rawTrades = answers.R012_TRADES || answers.C012_TRADES || [];
  var trades = Array.isArray(rawTrades) ? rawTrades : [rawTrades];
  var out = [];
  for (var i = 0; i < trades.length; i++) {
    var mapped = mapSurveyTradeToQuoteTradeCodes_(trades[i]);
    for (var j = 0; j < mapped.length; j++) pushUnique_(out, mapped[j]);
  }
  return out;
}

function hasAnyIntersection_(left, right) {
  if (!Array.isArray(left) || !Array.isArray(right) || !left.length || !right.length) return false;
  for (var i = 0; i < left.length; i++) {
    if (right.indexOf(left[i]) >= 0) return true;
  }
  return false;
}

function materialHasTag_(material, type, value) {
  var tagsByType = (material && material.tags_by_type) || {};
  var values = Array.isArray(tagsByType[type]) ? tagsByType[type] : [];
  return values.indexOf(String(value || "").trim()) >= 0;
}

function scoreMaterialRecommendation_(material, requestedTradeCodes, styleTags, toneTags) {
  var score = toNumber_(material.recommendation_score_base, 0);
  var matchedTags = { trade: [], style: [], tone: [] };
  var materialTrades = Array.isArray(material.effective_trade_codes) ? material.effective_trade_codes : [];

  if (requestedTradeCodes.length) {
    for (var t = 0; t < materialTrades.length; t++) {
      if (requestedTradeCodes.indexOf(materialTrades[t]) >= 0) {
        pushUnique_(matchedTags.trade, materialTrades[t]);
      }
    }
    if (matchedTags.trade.length) score += 8 + matchedTags.trade.length * 2;
  }

  for (var i = 0; i < styleTags.length; i++) {
    if (!materialHasTag_(material, "style", styleTags[i])) continue;
    pushUnique_(matchedTags.style, styleTags[i]);
    score += 4;
  }
  for (var j = 0; j < toneTags.length; j++) {
    if (!materialHasTag_(material, "tone", toneTags[j])) continue;
    pushUnique_(matchedTags.tone, toneTags[j]);
    score += 3;
  }

  if (material.is_representative) score += 2;
  if (!matchedTags.trade.length && !matchedTags.style.length && !matchedTags.tone.length) score -= 1;

  return {
    score: score,
    matched_tags: matchedTags
  };
}

function scoreTemplateRecommendation_(templateRow, answers, styleTags, toneTags) {
  var score = toNumber_(templateRow.prequote_priority, 0);
  var housingType = String(answers.R001_HOUSING_TYPE || answers.C001_BIZ_TYPE || "").trim().toUpperCase();
  var areaBand = String(answers.R002_AREA || answers.C002_AREA || "").trim().toUpperCase();
  var templateHousing = String(templateRow.housing_type || "").trim().toUpperCase();
  var templateArea = String(templateRow.area_band || "").trim().toUpperCase();
  var styleSummary = splitTemplateSummaryTags_(templateRow.style_tags_summary);
  var toneSummary = splitTemplateSummaryTags_(templateRow.tone_tags_summary);
  var matchedTags = { style: [], tone: [] };

  if (templateHousing && templateHousing === housingType) score += 5;
  if (templateArea && templateArea === areaBand) score += 4;
  if (templateRow.is_featured_prequote) score += 3;

  for (var i = 0; i < styleTags.length; i++) {
    if (styleSummary.indexOf(styleTags[i]) < 0) continue;
    pushUnique_(matchedTags.style, styleTags[i]);
    score += 3;
  }
  for (var j = 0; j < toneTags.length; j++) {
    if (toneSummary.indexOf(toneTags[j]) < 0) continue;
    pushUnique_(matchedTags.tone, toneTags[j]);
    score += 2;
  }

  return {
    score: score,
    matched_tags: matchedTags
  };
}

function getRecommendations_(tags, answers, projectType) {
  var styleTags = (tags || []).filter(function(tag) { return String(tag.type || "").toLowerCase() === "style"; }).map(function(tag) { return tag.value; });
  var toneTags = (tags || []).filter(function(tag) { return String(tag.type || "").toLowerCase() === "tone"; }).map(function(tag) { return tag.value; });
  var requestedTradeCodes = collectRequestedQuoteTradeCodes_(answers || {});
  var materials = loadMaterialsForRecommendations_();
  var templates = loadTemplatesForRecommendations_();
  var materialCandidates = [];
  var templateCandidates = [];

  for (var i = 0; i < materials.length; i++) {
    var material = materials[i];
    if (!material.is_active || !material.expose_to_prequote) continue;
    var scorePack = scoreMaterialRecommendation_(material, requestedTradeCodes, styleTags, toneTags);
    materialCandidates.push({
      type: "MATERIAL",
      material_id: material.material_id,
      trade_code: (scorePack.matched_tags.trade || []).join(","),
      material_type: material.material_type,
      title: material.name,
      subtitle: [material.brand, material.spec].filter(Boolean).join(" / "),
      brand: material.brand,
      spec: material.spec,
      reason: scorePack.matched_tags.style.length || scorePack.matched_tags.tone.length
        ? "스타일/톤 기반 추천"
        : (scorePack.matched_tags.trade.length ? "공종 기반 추천" : "노출 자재 기본 추천"),
      score: scorePack.score,
      price_min: material.unit_price,
      price_max: material.unit_price,
      image_url: material.image_url,
      image_file_id: material.image_file_id,
      source_type: material.source_type || "AUTO",
      matched_tags: scorePack.matched_tags
    });
  }

  materialCandidates.sort(function(a, b) {
    if (Number(a.score || 0) !== Number(b.score || 0)) return Number(b.score || 0) - Number(a.score || 0);
    return String(a.material_id || "").localeCompare(String(b.material_id || ""));
  });

  var seenMaterialIds = Object.create(null);
  var selectedMaterials = [];
  for (var m = 0; m < materialCandidates.length && selectedMaterials.length < 6; m++) {
    if (seenMaterialIds[materialCandidates[m].material_id]) continue;
    seenMaterialIds[materialCandidates[m].material_id] = 1;
    selectedMaterials.push(materialCandidates[m]);
  }

  for (var t = 0; t < templates.length; t++) {
    var templateRow = templates[t];
    if (!templateRow.is_active || !templateRow.expose_to_prequote) continue;
    var templateScore = scoreTemplateRecommendation_(templateRow, answers || {}, styleTags, toneTags);
    templateCandidates.push({
      type: "TEMPLATE",
      template_id: templateRow.template_id,
      title: templateRow.template_name,
      subtitle: templateRow.category,
      reason: "유사 조건 템플릿",
      score: templateScore.score,
      source_type: templateRow.source_type || "AUTO",
      matched_tags: templateScore.matched_tags
    });
  }

  templateCandidates.sort(function(a, b) {
    if (Number(a.score || 0) !== Number(b.score || 0)) return Number(b.score || 0) - Number(a.score || 0);
    return String(a.template_id || "").localeCompare(String(b.template_id || ""));
  });

  return selectedMaterials.concat(templateCandidates.slice(0, 3));
}

// ============================================================
//  RESULT PAGE DATA
// ============================================================

function buildAnswerMapFromRows_(answerRows) {
  var out = {};
  var rows = Array.isArray(answerRows) ? answerRows : [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    out[row.question_code] = parseStoredAnswerValue_(row);
  }
  return out;
}

function buildStoredAnswerRowsFromSnapshot_(requestRow) {
  var row = requestRow || {};
  var rid = String(row.request_id || "").trim();
  if (!rid) return [];
  var snapshot = safeJsonParse_(row.answers_snapshot_json, null);
  if (!snapshot || typeof snapshot !== "object") return [];

  var codes = Object.keys(snapshot);
  var createdAt = row.submitted_at || row.created_at || nowIso_();
  var rows = [];
  for (var i = 0; i < codes.length; i++) {
    var code = String(codes[i] || "").trim();
    if (!code) continue;
    var value = snapshot[code];
    rows.push({
      request_id: rid,
      answer_id: rid + "_S" + String(i + 1).padStart(3, "0"),
      question_id: code,
      question_code: code,
      answer_type: Array.isArray(value) ? "MULTI" : "SINGLE",
      answer_value_text: Array.isArray(value) ? value.join(",") : String(value === undefined || value === null ? "" : value),
      answer_value_number: Array.isArray(value) ? "" : (value === "" || value === null || value === undefined || !isFinite(Number(value)) ? "" : Number(value)),
      answer_value_json: JSON.stringify(value === undefined ? "" : value),
      normalized_answer: Array.isArray(value) ? value.join(",") : String(value === undefined || value === null ? "" : value),
      tags_applied_json: "",
      score_delta: "",
      created_at: createdAt
    });
  }
  return rows;
}

function getRequestAnswerRows_(requestRow, options) {
  var row = requestRow || {};
  var snapshotRows = buildStoredAnswerRowsFromSnapshot_(row);
  if (snapshotRows.length) return snapshotRows;
  var rid = String(row.request_id || (options && options.request_id) || "").trim();
  if (!rid) return [];
  return readRequestRowsIndexed_("RequestAnswers", rid, options).sort(function(a, b) {
    return String(a.answer_id || "").localeCompare(String(b.answer_id || ""));
  });
}

function buildStoredEstimateFromRequest_(requestRow) {
  var estimate = safeJsonParse_((requestRow || {}).estimate_json, {});
  if (!estimate || typeof estimate !== "object") estimate = {};
  if (estimate.min === undefined) estimate.min = toNumber_((requestRow || {}).estimate_min, 0);
  if (estimate.max === undefined) estimate.max = toNumber_((requestRow || {}).estimate_max, 0);
  if (!Array.isArray(estimate.adjustments)) estimate.adjustments = [];
  if (!Array.isArray(estimate.flags)) estimate.flags = safeJsonParse_((requestRow || {}).risk_flags_json, []);
  return estimate;
}

function generateAutoRecommendationsForRequest_(requestRow, answerRows) {
  var source = requestRow || {};
  var tags = safeJsonParse_(source.tags_json, []);
  var ansMap = buildAnswerMapFromRows_(answerRows || []);
  return getRecommendations_(tags, ansMap, source.project_type || "").map(function(item, index) {
    return normalizeAdminRecommendationItem_(item, index + 1, item.source_type || "AUTO");
  });
}

function saveRequestRecommendationState_(requestFound, recommendations, reviewId) {
  var found = requestFound || {};
  var requestRow = found.data || {};
  var rid = String(requestRow.request_id || "").trim();
  if (!rid) return [];
  ensurePrequoteOperationalSchema_();

  var now = nowIso_();
  var list = sanitizeReviewRecommendations_(recommendations || []);
  var rows = list.map(function(item, index) {
    var normalized = normalizeAdminRecommendationItem_(item, index + 1, item.source || item.source_type || "ADMIN");
    normalized.review_id = reviewId || normalized.review_id || "";
    if (normalized.is_visible === false && !normalized.hidden_at) normalized.hidden_at = now;
    if (normalized.is_deleted && !normalized.deleted_at) normalized.deleted_at = now;
    return buildRequestRecommendationRow_(rid, normalized, index + 1, normalized.created_at || now);
  });
  replaceRowsByKey_("RequestRecommendations", "request_id", rid, rows);
  return rows.map(normalizeRequestRecommendationRow_);
}

function ensureAutoRecommendationsSaved_(requestFound, answerRows) {
  var found = requestFound || {};
  var requestRow = found.data || {};
  var rid = String(requestRow.request_id || "").trim();
  if (!rid) return [];

  var recs = findRowsByCol_("RequestRecommendations", "request_id", rid)
    .map(normalizeRequestRecommendationRow_)
    .sort(function(a, b) { return Number(a.rank || 0) - Number(b.rank || 0); });
  if (recs.length > 0) return recs;

  try {
    var generatedRecs = generateAutoRecommendationsForRequest_(requestRow, answerRows);
    if (!generatedRecs.length) return [];

    var recRows = saveRequestRecommendationState_(found, generatedRecs, "");
    if (found.sheet && found.rowNo && found.meta) {
      updateRowFields_(found.sheet, found.rowNo, found.meta, {
        recommendation_count: generatedRecs.length,
        final_recommendation_count: generatedRecs.length,
        updated_at: nowIso_()
      });
    }
    return recRows;
  } catch (e) {
    Logger.log("Lazy recommendation error: " + e);
    return [];
  }
}

function loadRequestRecommendationState_(found, answerRows, options) {
  var opts = options || {};
  var rid = String(found && found.data && found.data.request_id || "").trim();
  if (!rid) return [];
  var rows = opts.allow_generate === false
    ? findRowsByCol_("RequestRecommendations", "request_id", rid)
        .map(normalizeRequestRecommendationRow_)
        .sort(function(a, b) { return Number(a.rank || 0) - Number(b.rank || 0); })
    : ensureAutoRecommendationsSaved_(found, answerRows);
  var list = Array.isArray(rows) ? rows.slice() : [];
  if (!opts.include_deleted) {
    list = list.filter(function(item) { return !item.is_deleted; });
  }
  if (!opts.include_hidden) {
    list = list.filter(function(item) { return item.is_visible !== false; });
  }
  list.sort(function(a, b) {
    return Number(a.rank || 0) - Number(b.rank || 0);
  });
  return list;
}

function loadOrGenerateRequestRecommendations_(found, answerRows) {
  return loadRequestRecommendationState_(found, answerRows, { include_hidden: false, include_deleted: false });
}

function buildPublicResultRequestView_(found, requestView) {
  var requestRow = found && found.data ? found.data : {};
  var view = requestView || {};
  return {
    request_id: requestRow.request_id,
    customer_name: requestRow.customer_name,
    project_type: requestRow.project_type,
    project_type_label: view.project_type_label,
    housing_type: requestRow.housing_type,
    housing_type_label: view.housing_type_label,
    area_py: requestRow.area_py,
    area_label: view.area_label,
    scope_level: requestRow.scope_level,
    scope_level_label: view.scope_level_label,
    flow_mode: requestRow.flow_mode,
    flow_mode_label: view.flow_mode_label,
    created_at: requestRow.created_at,
    status: requestRow.status,
    schedule_start_label: view.schedule_start_label,
    schedule_end_label: view.schedule_end_label,
    address_label: view.address_label,
    customer_note_label: view.customer_note_label
  };
}

function getResultDataCacheTtlSec_() {
  return Math.max(Number(getSettings_().result_cache_ttl_sec || 300), 0);
}

function buildResultDataCacheKey_(requestRow) {
  var row = requestRow || {};
  var rid = String(row.request_id || "").trim();
  var shareToken = String(row.share_token || "").trim();
  if (!rid || !shareToken) return "";
  return [
    "result-data",
    rid,
    shareToken,
    String(row.updated_at || row.submitted_at || row.created_at || ""),
    String(row.estimate_min || ""),
    String(row.estimate_max || ""),
    String(row.final_estimate_min || ""),
    String(row.final_estimate_max || ""),
    String(row.recommendation_count || ""),
    String(row.final_recommendation_count || "")
  ].join("|");
}

function resultRecommendationsMode_() {
  return String(getSettings_().result_recommendations_mode || "LAZY_FALLBACK").trim().toUpperCase();
}

function shouldHydrateResultRecommendations_(storedRecommendations) {
  var mode = resultRecommendationsMode_();
  if (mode === "OFF" || mode === "NONE" || mode === "DISABLED" || mode === "STORED_ONLY") return false;
  return !(Array.isArray(storedRecommendations) && storedRecommendations.length);
}

function hasTimelineEventDedup_(requestId, eventType) {
  var rid = String(requestId || "").trim();
  var code = String(eventType || "").trim().toUpperCase();
  if (!rid || !code) return false;
  var timelineRows = readRequestRowsIndexed_("RequestTimeline", rid, { allow_rebuild: false });
  for (var i = 0; i < timelineRows.length; i++) {
    if (String(timelineRows[i].event_type || "").trim().toUpperCase() === code) return true;
  }
  return false;
}

function hasRecentNotificationEvent_(requestId, eventType, dedupHours) {
  var rid = String(requestId || "").trim();
  var eventCode = String(eventType || "").trim().toUpperCase();
  if (!rid || !eventCode) return false;
  var hours = Math.max(Number(dedupHours || 0), 0);
  var thresholdMs = hours > 0 ? (Date.now() - hours * 3600000) : 0;
  var rows = findRowsByCol_("RequestNotificationLog", "request_id", rid);
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    if (String(row.event_type || "").trim().toUpperCase() !== eventCode) continue;
    var sentAtRaw = String(row.sent_at || row.created_at || "").trim();
    if (!sentAtRaw) return true;
    var sentMs = new Date(sentAtRaw).getTime();
    if (!hours || isNaN(sentMs) || sentMs >= thresholdMs) return true;
  }
  return false;
}

function maybeNotifyResultViewedSlack_(requestRow, requestView, estimate) {
  var row = requestRow || {};
  var rid = String(row.request_id || "").trim();
  if (!rid) return { ok: false, skipped: true, reason: "missing_request_id" };
  var settings = getSettings_();
  if (!ynToBool_(settings.slack_notify_on_result_view, true)) {
    return { ok: true, skipped: true, reason: "disabled" };
  }
  if (hasTimelineEventDedup_(rid, "RESULT_VIEWED")) {
    return { ok: true, skipped: true, reason: "already_marked" };
  }
  var dedupHours = Math.max(Number(settings.slack_result_view_dedup_hours || 24), 0);
  if (hasRecentNotificationEvent_(rid, "RESULT_VIEWED", dedupHours)) {
    return { ok: true, skipped: true, reason: "recent_notification" };
  }

  var extra = {
    project_label: requestView && requestView.project_type_label,
    area_label: requestView && requestView.area_label,
    summary: "고객이 결과 페이지를 확인했습니다."
  };
  var result = sendSlackWebhook_(rid, "RESULT_VIEWED", "request:" + rid + ":result-viewed:first", {
    text: buildSlackRequestText_(Object.assign({}, row, {
      final_estimate_min: estimate && estimate.min,
      final_estimate_max: estimate && estimate.max
    }), "RESULT_VIEWED", extra)
  });
  appendRow_("RequestTimeline", buildTimelineEvent_(
    rid,
    "CUSTOMER",
    "RESULT_VIEWED",
    "",
    "",
    "고객이 결과 페이지를 확인했습니다.",
    JSON.stringify({
      notify_ok: !!(result && result.ok),
      notify_skipped: !!(result && result.skipped)
    }),
    nowIso_()
  ));
  return result;
}

function getResultData(requestId, shareToken) {
  ensureSpreadsheetId_();
  var startMs = Date.now();
  var error = null;
  var cached = null;
  var rid = String(requestId || "").trim();
  var tok = String(shareToken || "").trim();
  try {
    if (!rid || !tok) throw new Error("Invalid request");

    var found = findRequestRowById_(rid);
    if (!found) throw new Error("Request not found");
    if (String(found.data.share_token || "") !== tok) throw new Error("Invalid share token");

    var settings = getSettings_();
    var cacheTtlSec = getResultDataCacheTtlSec_();
    var cacheKey = cacheTtlSec ? buildResultDataCacheKey_(found.data) : "";
    cached = cacheKey ? cacheJsonGet_(cacheKey) : null;
    if (cached) {
      try {
        maybeNotifyResultViewedSlack_(found.data, cached.request || {}, cached.estimate || {});
      } catch (notifyCachedError) {
        Logger.log("Result view notify error: " + notifyCachedError);
      }
      return cached;
    }

    var answerRows = getRequestAnswerRows_(found.data, { request_id: rid });
    var autoEstimate = buildStoredEstimateFromRequest_(found.data);
    var flags = safeJsonParse_(found.data.risk_flags_json, []);
    var recs = loadRequestRecommendationState_(found, answerRows, {
      include_hidden: false,
      include_deleted: false,
      allow_generate: false
    });
    var optionIndex = getSurveyOptionLabelIndex_();
    var questionMetaMap = getSurveyQuestionMetaMap_();
    var enrichedAnswers = answerRows.map(function(row) {
      return enrichAdminAnswerRow_(row, {}, questionMetaMap, optionIndex);
    });
    var answerGroups = groupAdminAnswers_(enrichedAnswers);
    var answerIndex = indexAdminAnswersByCode_(enrichedAnswers);
    var requestView = buildRequestDisplayForAdmin_(found.data, answerIndex, null, autoEstimate, autoEstimate, recs, optionIndex);

    try {
      maybeNotifyResultViewedSlack_(found.data, requestView, autoEstimate);
    } catch (notifyError) {
      Logger.log("Result view notify error: " + notifyError);
    }

    var payload = {
      request: buildPublicResultRequestView_(found, requestView),
      estimate: Object.assign({}, autoEstimate, { source: "AUTO", source_label: "자동 산출값" }),
      flags: flags,
      recommendations: recs,
      answers: enrichedAnswers,
      answer_groups: answerGroups,
      notice_text: String(settings.result_notice_review_pending || "").trim(),
      recommendation_status: recs.length ? "READY" : (shouldHydrateResultRecommendations_(recs) ? "PENDING" : "EMPTY"),
      recommendation_count: recs.length,
      first_paint: {
        range_label: formatMoneyRange(autoEstimate.min || 0, autoEstimate.max || 0),
        request_summary: [requestView.project_type_label, requestView.housing_type_label, requestView.area_label, requestView.scope_level_label].filter(Boolean).join(" / "),
        created_at_label: formatAdminDateTime_(found.data.created_at)
      }
    };
    if (cacheKey && cacheTtlSec) cacheJsonPut_(cacheKey, payload, cacheTtlSec);
    return payload;
  } catch (e) {
    error = e;
    throw e;
  } finally {
    logPerfMetricSafe_(
      "result_data_load",
      Date.now() - startMs,
      rid,
      { cache_hit: !!cached, used_snapshot: !!(found && found.data && found.data.answers_snapshot_json) },
      !error,
      error ? String(error) : "",
      "CUSTOMER",
      rid
    );
  }
}

function getResultRecommendations(requestId, shareToken) {
  ensureSpreadsheetId_();
  var rid = String(requestId || "").trim();
  var tok = String(shareToken || "").trim();
  if (!rid || !tok) throw new Error("Invalid request");
  var found = findRequestRowById_(rid);
  if (!found) throw new Error("접수 내역을 찾을 수 없습니다.");
  if (String(found.data.share_token || "") !== tok) throw new Error("접근 권한이 없습니다.");
  var answerRows = getRequestAnswerRows_(found.data, { request_id: rid });
  return {
    request_id: rid,
    recommendations: loadRequestRecommendationState_(found, answerRows, {
      include_hidden: false,
      include_deleted: false,
      allow_generate: true
    }),
    status: "READY"
  };
}

// ============================================================
//  ADMIN API
// ============================================================

function ensureAdminReviewSheet_() {
  return ensureSheetWithHeaders_(REQUEST_ADMIN_REVIEW_SHEET_, REQUEST_ADMIN_REVIEW_HEADERS_);
}

function normalizeAdminReviewStatus_(value, fallback) {
  var code = String(value || fallback || "AUTO").trim().toUpperCase();
  if (ADMIN_REVIEW_STATUS_OPTIONS_.indexOf(code) >= 0) return code;
  return String(fallback || "AUTO").trim().toUpperCase();
}

function statusLabel_(status) {
  var code = String(status || "").trim().toUpperCase();
  return STATUS_LABEL_MAP_[code] || (code || "-");
}

function reviewStatusLabel_(status) {
  var code = normalizeAdminReviewStatus_(status, "AUTO");
  return REVIEW_STATUS_LABEL_MAP_[code] || code;
}

function eventTypeLabel_(eventType) {
  var code = String(eventType || "").trim().toUpperCase();
  return EVENT_TYPE_LABEL_MAP_[code] || (code || "이력");
}

function severityLabel_(severity) {
  var code = String(severity || "").trim().toUpperCase();
  return SEVERITY_LABEL_MAP_[code] || (code || "기타");
}

function priorityLabel_(priority) {
  var code = String(priority || "NORMAL").trim().toUpperCase();
  return PRIORITY_LABEL_MAP_[code] || PRIORITY_LABEL_MAP_.NORMAL;
}

function workStateLabel_(workState) {
  var code = String(workState || "NEW").trim().toUpperCase();
  return WORK_STATE_LABEL_MAP_[code] || WORK_STATE_LABEL_MAP_.NEW;
}

function taskStatusLabel_(taskStatus) {
  var code = String(taskStatus || "OPEN").trim().toUpperCase();
  return TASK_STATUS_LABEL_MAP_[code] || TASK_STATUS_LABEL_MAP_.OPEN;
}

function formatAdminDateTime_(value) {
  var raw = String(value || "").trim();
  if (!raw) return "";
  var dt = new Date(raw);
  if (isNaN(dt.getTime())) return raw;
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), "yyyy.MM.dd HH:mm");
}

function buildNextActionSummary_(requestRow) {
  var row = requestRow || {};
  var parts = [];
  var actionType = String(row.next_action_type || "").trim();
  var actionNote = String(row.next_action_note || "").trim();
  var dueAt = formatAdminDateTime_(row.next_action_due_at);
  if (actionType) parts.push(actionType);
  if (actionNote) parts.push(actionNote);
  if (dueAt) parts.push(dueAt);
  return parts.length ? parts.join(" · ") : "다음 액션 미지정";
}

function mapQuestionSectionToAdminGroup_(sectionId) {
  var section = String(sectionId || "").trim().toUpperCase();
  if (!section) return "OTHER";
  if (section === "COMMON_CONTACT") return "CONTACT";
  if (section === "ENTRY") return "ENTRY";
  if (/_BASIC$/.test(section)) return "BASIC";
  if (/_SCOPE$/.test(section)) return "SCOPE";
  if (/_TRADE_/.test(section)) return "TRADE";
  if (/_STYLE$/.test(section)) return "STYLE";
  if (/_LIFE$/.test(section)) return "LIFE";
  if (/_RISK$/.test(section)) return "RISK";
  if (/_BUDGET$/.test(section)) return "BUDGET";
  if (/_SCHEDULE$/.test(section)) return "SCHEDULE";
  return "OTHER";
}

function getAnswerGroupLabel_(groupKey) {
  var code = String(groupKey || "OTHER").trim().toUpperCase();
  return ANSWER_GROUP_LABEL_MAP_[code] || ANSWER_GROUP_LABEL_MAP_.OTHER;
}

function getAnswerGroupOrder_(groupKey) {
  var code = String(groupKey || "OTHER").trim().toUpperCase();
  return Number(ANSWER_GROUP_ORDER_MAP_[code] || ANSWER_GROUP_ORDER_MAP_.OTHER || 999);
}

function getSurveyQuestionMetaMap_() {
  var rows = readStaticRowsCached_("SurveyQuestions").filter(function(q) {
    return ynToBool_(q.is_active, true);
  });
  var out = Object.create(null);
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var code = String(row.question_code || row.question_id || "").trim();
    if (!code) continue;
    var groupKey = mapQuestionSectionToAdminGroup_(row.section_id);
    var meta = {
      question_id: String(row.question_id || code).trim(),
      question_code: code,
      question_label: String(row.title || code).trim(),
      group_key: groupKey,
      group_label: getAnswerGroupLabel_(groupKey),
      group_order: getAnswerGroupOrder_(groupKey),
      section_id: String(row.section_id || "").trim(),
      question_order: Number(row.sort_order || row.step_no || i + 1)
    };
    out[code] = meta;
    out[meta.question_id] = meta;
  }
  return out;
}

function getSurveyOptionLabelIndex_() {
  var rows = readStaticRowsCached_("SurveyOptions").filter(function(o) {
    return ynToBool_(o.is_active, true);
  });
  var out = Object.create(null);
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var questionId = String(row.question_id || "").trim();
    if (!questionId) continue;
    if (!out[questionId]) out[questionId] = Object.create(null);
    var label = String(row.option_label || row.option_value || row.option_code || "").trim();
    var keys = toUniqueStringArray_([
      row.option_code,
      row.option_value,
      row.numeric_value !== undefined && row.numeric_value !== null && row.numeric_value !== ""
        ? String(row.numeric_value)
        : ""
    ]);
    for (var k = 0; k < keys.length; k++) {
      out[questionId][keys[k]] = label || keys[k];
      out[questionId][String(keys[k]).toUpperCase()] = label || keys[k];
    }
  }
  return out;
}

function lookupSurveyOptionLabel_(questionId, value, optionIndex) {
  if (value === undefined || value === null || value === "") return "";
  if (value === true) return "예";
  if (value === false) return "아니오";

  var raw = String(value).trim();
  if (!raw) return "";
  var questionKey = String(questionId || "").trim();
  var questionMap = (optionIndex && optionIndex[questionKey]) || {};
  if (Object.prototype.hasOwnProperty.call(questionMap, raw)) {
    return formatDisplayValueByQuestion_(questionKey, raw, questionMap[raw]);
  }
  if (Object.prototype.hasOwnProperty.call(questionMap, raw.toUpperCase())) {
    return formatDisplayValueByQuestion_(questionKey, raw, questionMap[raw.toUpperCase()]);
  }

  var upper = raw.toUpperCase();
  if (upper === "Y") return "예";
  if (upper === "N") return "아니오";
  return formatDisplayValueByQuestion_(questionKey, raw, raw);
}

function isBudgetQuestion_(questionId) {
  return /BUDGET/.test(String(questionId || "").trim().toUpperCase());
}

function formatBudgetDisplayValue_(rawValue, label) {
  var raw = String(rawValue || "").trim().toUpperCase();
  var base = String(label || rawValue || "").trim();
  var codeMap = {
    B_LT10: "1000만원 미만",
    B10: "1000만원 이하",
    B10_20: "1000~2000만원",
    B20_30: "2000~3000만원",
    B30_40: "3000~4000만원",
    B30_50: "3000~5000만원",
    B40_50: "4000~5000만원",
    B50_60: "5000~6000만원",
    B50_70: "5000~7000만원",
    B60_70: "6000~7000만원",
    B70_80: "7000~8000만원",
    B70_100: "7000만원~1억",
    B80_100: "8000만원~1억",
    GE100: "1억 이상",
    B100_UP: "1억 이상",
    CONSULT: "상담 후 결정"
  };
  if (codeMap[raw]) return codeMap[raw];
  if (!base) return "";
  if (/^\d+천\s*미만$/.test(base)) {
    return base.replace(/^(\d+)천\s*미만$/, function(_, v) { return String(Number(v) * 1000) + "만원 미만"; });
  }
  if (/^\d+천만원 이하$/.test(base)) {
    return base.replace(/^(\d+)천만원 이하$/, function(_, v) { return String(Number(v) * 1000) + "만원 이하"; });
  }
  if (/^\d+천\s*~\s*\d+천$/.test(base)) {
    return base.replace(/(\d+)천\s*~\s*(\d+)천/, function(_, a, b) {
      return String(Number(a) * 1000) + "~" + String(Number(b) * 1000) + "만원";
    });
  }
  if (/^\d+천\s*~\s*1억$/.test(base)) {
    return base.replace(/(\d+)천\s*~\s*1억/, function(_, a) {
      return String(Number(a) * 1000) + "만원~1억";
    });
  }
  return base;
}

function formatDisplayValueByQuestion_(questionId, rawValue, fallbackLabel) {
  var questionKey = String(questionId || "").trim();
  if (isBudgetQuestion_(questionKey)) {
    return formatBudgetDisplayValue_(rawValue, fallbackLabel);
  }
  return String(fallbackLabel || rawValue || "").trim();
}

function parseStoredAnswerValue_(row) {
  var parsed = safeJsonParse_((row || {}).answer_value_json, null);
  if (parsed !== null && parsed !== undefined) return parsed;
  if (String((row || {}).answer_type || "").trim().toUpperCase() === "MULTI") {
    return normalizeStringList_((row || {}).answer_value_text);
  }
  if (row && row.answer_value_number !== "" && row.answer_value_number !== null && row.answer_value_number !== undefined) {
    var num = Number(row.answer_value_number);
    if (isFinite(num)) return num;
  }
  return row ? row.answer_value_text : "";
}

function buildAnswerDisplayValues_(questionCode, rawValue, answerType, optionIndex) {
  var values = [];
  if (Array.isArray(rawValue)) {
    values = rawValue.slice();
  } else if (String(answerType || "").trim().toUpperCase() === "MULTI") {
    values = normalizeStringList_(rawValue);
  } else if (rawValue !== undefined && rawValue !== null && String(rawValue).trim() !== "") {
    values = [rawValue];
  }

  var displayValues = [];
  for (var i = 0; i < values.length; i++) {
    var displayValue = lookupSurveyOptionLabel_(questionCode, values[i], optionIndex) || String(values[i] || "").trim();
    if (!displayValue) continue;
    displayValues.push(displayValue);
  }
  return displayValues;
}

function enrichAdminAnswerRow_(row, overrideAnswers, questionMetaMap, optionIndex) {
  var answerRow = row || {};
  var code = String(answerRow.question_code || answerRow.question_id || "").trim();
  var meta = questionMetaMap[code] || {
    question_id: String(answerRow.question_id || code).trim(),
    question_code: code,
    question_label: code || "미정 질문",
    group_key: "OTHER",
    group_label: getAnswerGroupLabel_("OTHER"),
    group_order: getAnswerGroupOrder_("OTHER"),
    section_id: "",
    question_order: 9999
  };
  var originalValue = parseStoredAnswerValue_(answerRow);
  var hasOverride = !!(overrideAnswers && Object.prototype.hasOwnProperty.call(overrideAnswers, code));
  var finalValue = hasOverride ? overrideAnswers[code] : originalValue;
  var originalDisplayValues = buildAnswerDisplayValues_(meta.question_id, originalValue, answerRow.answer_type, optionIndex);
  var finalDisplayValues = buildAnswerDisplayValues_(meta.question_id, finalValue, answerRow.answer_type, optionIndex);

  return {
    request_id: answerRow.request_id,
    answer_id: answerRow.answer_id,
    question_id: meta.question_id,
    question_code: meta.question_code,
    question_label: meta.question_label,
    answer_type: answerRow.answer_type,
    answer_group: meta.group_key,
    group_label: meta.group_label,
    group_order: meta.group_order,
    section_id: meta.section_id,
    question_sort_order: meta.question_order,
    has_override: hasOverride,
    original_value: originalValue,
    final_value: finalValue,
    original_display_values: originalDisplayValues,
    original_display_text: originalDisplayValues.length ? originalDisplayValues.join(", ") : "미입력",
    answer_display_values: finalDisplayValues,
    answer_display_text: finalDisplayValues.length ? finalDisplayValues.join(", ") : "미입력"
  };
}

function groupAdminAnswers_(answerRows) {
  var groups = [];
  var map = Object.create(null);
  var rows = Array.isArray(answerRows) ? answerRows : [];
  rows.sort(function(a, b) {
    if (Number(a.group_order || 0) !== Number(b.group_order || 0)) {
      return Number(a.group_order || 0) - Number(b.group_order || 0);
    }
    if (Number(a.question_sort_order || 0) !== Number(b.question_sort_order || 0)) {
      return Number(a.question_sort_order || 0) - Number(b.question_sort_order || 0);
    }
    return String(a.question_code || "").localeCompare(String(b.question_code || ""));
  });

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var key = row.answer_group || "OTHER";
    if (!map[key]) {
      map[key] = {
        group_key: key,
        group_label: row.group_label || getAnswerGroupLabel_(key),
        group_order: row.group_order || getAnswerGroupOrder_(key),
        items: []
      };
      groups.push(map[key]);
    }
    map[key].items.push(row);
  }

  groups.sort(function(a, b) {
    return Number(a.group_order || 0) - Number(b.group_order || 0);
  });
  return groups;
}

function indexAdminAnswersByCode_(answerRows) {
  var out = Object.create(null);
  var rows = Array.isArray(answerRows) ? answerRows : [];
  for (var i = 0; i < rows.length; i++) out[rows[i].question_code] = rows[i];
  return out;
}

function getAnswerDisplayByCodes_(answerIndex, codes, fallback) {
  var list = Array.isArray(codes) ? codes : [codes];
  for (var i = 0; i < list.length; i++) {
    var key = String(list[i] || "").trim();
    if (!key || !answerIndex[key]) continue;
    if (answerIndex[key].answer_display_text) return answerIndex[key].answer_display_text;
  }
  var raw = String(fallback || "").trim();
  return raw || "미입력";
}

function questionCodeForProject_(projectType, residentialCode, commercialCode) {
  return String(projectType || "").trim().toUpperCase() === "COMMERCIAL" ? commercialCode : residentialCode;
}

function normalizeRiskFlagsForAdmin_(flags) {
  var rows = Array.isArray(flags) ? flags : [];
  return rows.map(function(flag) {
    var severity = String((flag || {}).severity || "").trim().toUpperCase();
    return {
      label: String((flag || {}).label || "").trim(),
      severity: severity || "MEDIUM",
      severity_label: severityLabel_(severity || "MEDIUM")
    };
  });
}

function normalizeAdminRecommendationItem_(rec, rank, sourceType) {
  var item = rec || {};
  return {
    rec_id: String(item.rec_id || item.review_item_id || item.material_id || item.template_id || ("REC_" + String(rank || 0))).trim(),
    rec_type: String(item.rec_type || item.type || "MATERIAL").trim().toUpperCase(),
    rank: Number(rank || item.rank || 0),
    template_id: String(item.template_id || "").trim(),
    material_id: String(item.material_id || "").trim(),
    trade_code: String(item.trade_code || "").trim(),
    material_type: String(item.material_type || "").trim(),
    title: String(item.title || "").trim(),
    subtitle: String(item.subtitle || "").trim(),
    reason_text: String(item.reason_text || item.reason || "").trim(),
    score: toNumber_(item.score, 0),
    price_hint_min: toNumber_(item.price_hint_min !== undefined ? item.price_hint_min : item.price_min, 0),
    price_hint_max: toNumber_(item.price_hint_max !== undefined ? item.price_hint_max : item.price_max, 0),
    image_url: String(item.image_url || "").trim(),
    image_file_id: String(item.image_file_id || "").trim(),
    brand: String(item.brand || "").trim(),
    spec: String(item.spec || "").trim(),
    matched_tags: item.matched_tags && typeof item.matched_tags === "object" ? item.matched_tags : {},
    is_visible: item.is_visible !== false,
    is_deleted: !!item.is_deleted,
    is_manual: !!item.is_manual,
    source: String(sourceType || item.source || item.source_type || "AUTO").trim().toUpperCase(),
    source_type: String(item.source_type || item.source || sourceType || "AUTO").trim().toUpperCase(),
    source_ref_id: String(item.source_ref_id || item.material_id || item.template_id || "").trim(),
    source_ref_key: String(item.source_ref_key || item.trade_code || item.material_type || "").trim(),
    review_id: String(item.review_id || "").trim(),
    hidden_at: String(item.hidden_at || "").trim(),
    deleted_at: String(item.deleted_at || "").trim(),
    note: String(item.note || "").trim()
  };
}

function sanitizeReviewRecommendations_(rows) {
  var list = Array.isArray(rows) ? rows : [];
  var out = [];
  for (var i = 0; i < list.length; i++) {
    var item = normalizeAdminRecommendationItem_(list[i], i + 1, (list[i] || {}).source || ((list[i] || {}).is_manual ? "MANUAL" : "ADMIN"));
    if (!item.title && !item.material_id && !item.template_id) continue;
    item.rank = out.length + 1;
    if (item.is_visible === false && !item.hidden_at) item.hidden_at = nowIso_();
    if (item.is_deleted && !item.deleted_at) item.deleted_at = nowIso_();
    out.push(item);
  }
  return out;
}

function buildComparableRecommendationList_(rows) {
  return sanitizeReviewRecommendations_(rows).map(function(item, index) {
    return {
      rec_type: item.rec_type,
      template_id: item.template_id,
      material_id: item.material_id,
      trade_code: item.trade_code,
      material_type: item.material_type,
      title: item.title,
      subtitle: item.subtitle,
      reason_text: item.reason_text,
      score: item.score,
      price_hint_min: item.price_hint_min,
      price_hint_max: item.price_hint_max,
      is_visible: item.is_visible !== false,
      is_deleted: !!item.is_deleted,
      is_manual: !!item.is_manual,
      rank: index + 1
    };
  });
}

function recommendationListsEqual_(left, right) {
  return JSON.stringify(buildComparableRecommendationList_(left)) === JSON.stringify(buildComparableRecommendationList_(right));
}

function sanitizeReviewAnswers_(overrideAnswers, baseAnswerMap) {
  var raw = overrideAnswers && typeof overrideAnswers === "object" ? overrideAnswers : {};
  var base = baseAnswerMap || {};
  var out = {};
  for (var key in raw) {
    if (!Object.prototype.hasOwnProperty.call(raw, key)) continue;
    var questionCode = String(key || "").trim();
    if (!questionCode) continue;
    var value = raw[key];
    var normalized = Array.isArray(value)
      ? value.map(function(v) { return String(v || "").trim(); }).filter(Boolean)
      : value;
    if (JSON.stringify(normalized) === JSON.stringify(base[questionCode])) continue;
    out[questionCode] = normalized;
  }
  return out;
}

function buildFinalEstimateForAdmin_(autoEstimate, review) {
  var base = autoEstimate && typeof autoEstimate === "object"
    ? JSON.parse(JSON.stringify(autoEstimate))
    : { min: 0, max: 0, adjustments: [], flags: [] };
  if (!review || !review.has_estimate_override) {
    base.source = "AUTO";
    base.source_label = "자동 산출값";
    return base;
  }
  base.min = toNumber_(review.override_estimate_min, toNumber_(base.min, 0));
  base.max = toNumber_(review.override_estimate_max, toNumber_(base.max, 0));
  base.source = "ADMIN";
  base.source_label = "관리자 수정값";
  return base;
}

function buildFinalRecommendationsForAdmin_(autoRecommendations, review) {
  var baseList = sanitizeReviewRecommendations_((autoRecommendations || []).map(function(item) {
    return normalizeAdminRecommendationItem_(item, item.rank || 0, "AUTO");
  }));
  if (!review || !review.has_recommendation_override) return baseList.filter(function(item) { return !item.is_deleted; });
  return sanitizeReviewRecommendations_(review.override_recommendations || []).filter(function(item) { return !item.is_deleted; });
}

function normalizeAdminReviewRow_(row) {
  var payload = safeJsonParse_((row || {}).payload_json, {});
  var overrideRecommendationsRaw = String((row || {}).override_recommendations_json || "").trim();
  var overrideAnswersRaw = String((row || {}).override_answers_json || "").trim();
  var overrideRecommendations = overrideRecommendationsRaw ? sanitizeReviewRecommendations_(safeJsonParse_(overrideRecommendationsRaw, [])) : [];
  var overrideAnswers = overrideAnswersRaw ? safeJsonParse_(overrideAnswersRaw, {}) : {};
  var reviewStatus = normalizeAdminReviewStatus_((row || {}).review_status, "AUTO");
  var overrideEstimateMin = (row || {}).override_estimate_min === "" || (row || {}).override_estimate_min === null || (row || {}).override_estimate_min === undefined
    ? ""
    : Number((row || {}).override_estimate_min);
  var overrideEstimateMax = (row || {}).override_estimate_max === "" || (row || {}).override_estimate_max === null || (row || {}).override_estimate_max === undefined
    ? ""
    : Number((row || {}).override_estimate_max);

  return {
    _rowNo: Number(row._rowNo || 0),
    review_id: row.review_id,
    request_id: row.request_id,
    edited_at: row.edited_at,
    actor_type: row.actor_type || "ADMIN",
    actor_id: row.actor_id || "",
    review_status: reviewStatus,
    review_status_label: reviewStatusLabel_(reviewStatus),
    override_estimate_min: isFinite(overrideEstimateMin) ? overrideEstimateMin : "",
    override_estimate_max: isFinite(overrideEstimateMax) ? overrideEstimateMax : "",
    override_recommendations: overrideRecommendations,
    override_answers: overrideAnswers && typeof overrideAnswers === "object" ? overrideAnswers : {},
    estimate_reason: String((row || {}).estimate_reason || (payload || {}).estimate_reason || "").trim(),
    review_note: String(row.review_note || "").trim(),
    review_summary: String(row.review_summary || (payload || {}).review_summary || "").trim(),
    final_estimate_min: toNumber_((row || {}).final_estimate_min, overrideEstimateMin !== "" ? overrideEstimateMin : 0),
    final_estimate_max: toNumber_((row || {}).final_estimate_max, overrideEstimateMax !== "" ? overrideEstimateMax : 0),
    final_recommendation_count: toNumber_((row || {}).final_recommendation_count, 0),
    source_label: String((row || {}).source_label || (payload || {}).source_label || "").trim(),
    linked_quote_id: String((row || {}).linked_quote_id || "").trim(),
    linked_quote_status: String((row || {}).linked_quote_status || "").trim(),
    is_current: ynToBool_((row || {}).is_current, true),
    created_at: String((row || {}).created_at || (row || {}).edited_at || "").trim(),
    updated_at: String((row || {}).updated_at || (row || {}).edited_at || "").trim(),
    payload: payload && typeof payload === "object" ? payload : {},
    has_estimate_override: overrideEstimateMin !== "" && overrideEstimateMax !== "",
    has_recommendation_override: !!overrideRecommendationsRaw,
    has_answer_override: !!overrideAnswersRaw && Object.keys(overrideAnswers || {}).length > 0,
    has_manual_override: (overrideEstimateMin !== "" && overrideEstimateMax !== "") || !!overrideRecommendationsRaw || (!!overrideAnswersRaw && Object.keys(overrideAnswers || {}).length > 0)
  };
}

function readAllAdminReviewRows_() {
  ensureAdminReviewSheet_();
  return readAllRows_(REQUEST_ADMIN_REVIEW_SHEET_).map(normalizeAdminReviewRow_);
}

function getLatestAdminReviewMap_() {
  var rows = readAllAdminReviewRows_();
  var map = Object.create(null);
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var rid = String(row.request_id || "").trim();
    if (!rid) continue;
    if (!map[rid] || String(row.edited_at || "").localeCompare(String(map[rid].edited_at || "")) > 0) {
      map[rid] = row;
    }
  }
  return map;
}

function getLatestAdminReviewForRequest_(requestId) {
  ensureAdminReviewSheet_();
  var rid = String(requestId || "").trim();
  var record = readRequestRowIndexRecord_(rid);
  var rowNo = Number(record && record.latest_review_row_no || 0);
  if (rowNo >= 2) {
    var sh = getSheet_(REQUEST_ADMIN_REVIEW_SHEET_);
    var meta = sheetMeta_(sh);
    var row = sh.getRange(rowNo, 1, 1, meta.headers.length).getValues()[0];
    if (String(row[meta.cols.request_id] || "").trim() === rid) {
      return normalizeAdminReviewRow_(rowToObj_(meta.headers, row, rowNo));
    }
  }
  var rows = readRequestRowsIndexed_(REQUEST_ADMIN_REVIEW_SHEET_, rid)
    .map(normalizeAdminReviewRow_)
    .sort(function(a, b) {
      return String(b.edited_at || "").localeCompare(String(a.edited_at || ""));
    });
  if (rows.length && requestRowIndexEnabled_()) {
    upsertRequestRowIndexRecord_(rid, {
      latest_review_id: String(rows[0].review_id || "").trim(),
      latest_review_row_no: Number(rows[0]._rowNo || 0)
    });
  }
  return rows.length ? rows[0] : null;
}

function buildAdminReviewPayloadSnapshot_(reviewRow, autoEstimate, finalEstimate, finalRecommendations) {
  return {
    review_status: reviewRow.review_status,
    review_status_label: reviewRow.review_status_label,
    has_estimate_override: reviewRow.has_estimate_override,
    has_recommendation_override: reviewRow.has_recommendation_override,
    has_answer_override: reviewRow.has_answer_override,
    estimate_reason: reviewRow.payload.estimate_reason || "",
    auto_estimate: {
      min: toNumber_((autoEstimate || {}).min, 0),
      max: toNumber_((autoEstimate || {}).max, 0)
    },
    final_estimate: {
      min: toNumber_((finalEstimate || {}).min, 0),
      max: toNumber_((finalEstimate || {}).max, 0)
    },
    final_recommendation_count: (finalRecommendations || []).filter(function(item) {
      return item.is_visible !== false && !item.is_deleted;
    }).length,
    saved_at: reviewRow.edited_at
  };
}

function summarizeAdminReviewSave_(reviewRow) {
  var summary = [];
  if (reviewRow.has_estimate_override) summary.push("가견적 범위 수정");
  if (reviewRow.has_recommendation_override) summary.push("추천 항목 조정");
  if (reviewRow.has_answer_override) summary.push("운영 보정값 반영");
  if (reviewRow.review_note) summary.push("검토 메모 저장");
  if (!summary.length && reviewRow.review_status === "IN_REVIEW") summary.push("검토 상태 업데이트");
  return summary.length ? summary.join(" / ") : "관리자 검토 결과 저장";
}

function buildAdminReviewRow_(requestId, payload) {
  var reviewInput = payload || {};
  var row = {
    review_id: uuid_(),
    request_id: String(requestId || "").trim(),
    edited_at: reviewInput.edited_at || nowIso_(),
    actor_type: reviewInput.actor_type || "ADMIN",
    actor_id: reviewInput.actor_id || "",
    review_status: normalizeAdminReviewStatus_(reviewInput.review_status, "AUTO"),
    override_estimate_min: reviewInput.override_estimate_min !== "" && reviewInput.override_estimate_min !== null && reviewInput.override_estimate_min !== undefined
      ? Number(reviewInput.override_estimate_min)
      : "",
    override_estimate_max: reviewInput.override_estimate_max !== "" && reviewInput.override_estimate_max !== null && reviewInput.override_estimate_max !== undefined
      ? Number(reviewInput.override_estimate_max)
      : "",
    override_recommendations_json: reviewInput.has_recommendation_override ? JSON.stringify(reviewInput.override_recommendations || []) : "",
    override_answers_json: reviewInput.has_answer_override ? JSON.stringify(reviewInput.override_answers || {}) : "",
    review_note: String(reviewInput.review_note || "").trim(),
    payload_json: JSON.stringify(reviewInput.payload || {})
  };
  return normalizeAdminReviewRow_(row);
}

function buildAdminReviewViewModel_(review, autoEstimate, finalEstimate, recommendationEditorItems) {
  var current = review || {
    review_id: "",
    edited_at: "",
    review_status: "AUTO",
    review_status_label: reviewStatusLabel_("AUTO"),
    review_note: "",
    override_estimate_min: "",
    override_estimate_max: "",
    override_recommendations: [],
    override_answers: {},
    has_estimate_override: false,
    has_recommendation_override: false,
    has_answer_override: false,
    has_manual_override: false,
    payload: {}
  };
  var sourceCode = current.has_manual_override ? "ADMIN" : (current.review_status === "IN_REVIEW" ? "IN_REVIEW" : "AUTO");
  var sourceLabel = String(current.source_label || "").trim() || (sourceCode === "ADMIN" ? "관리자 수정값" : (sourceCode === "IN_REVIEW" ? "관리자 검토중" : "자동 산출값"));
  return {
    review_id: current.review_id,
    edited_at: current.edited_at,
    review_status: current.review_status,
    review_status_label: current.review_status_label,
    review_note: current.review_note || "",
    override_estimate_min: current.has_estimate_override ? current.override_estimate_min : "",
    override_estimate_max: current.has_estimate_override ? current.override_estimate_max : "",
    override_answers: current.override_answers || {},
    estimate_reason: String((current.payload || {}).estimate_reason || "").trim(),
    auto_estimate_min: toNumber_((autoEstimate || {}).min, 0),
    auto_estimate_max: toNumber_((autoEstimate || {}).max, 0),
    final_estimate_min: toNumber_((finalEstimate || {}).min, 0),
    final_estimate_max: toNumber_((finalEstimate || {}).max, 0),
    recommendations: recommendationEditorItems || [],
    has_estimate_override: !!current.has_estimate_override,
    has_recommendation_override: !!current.has_recommendation_override,
    has_answer_override: !!current.has_answer_override,
    has_manual_override: !!current.has_manual_override,
    source_code: sourceCode,
    source_label: sourceLabel
  };
}

function buildTimelineDisplayMessage_(row) {
  var eventType = String((row || {}).event_type || "").trim().toUpperCase();
  var message = String((row || {}).message || "").trim();
  if (eventType === "STATUS_CHANGE") {
    var fromLabel = statusLabel_(row.from_status);
    var toLabel = statusLabel_(row.to_status);
    return fromLabel + " → " + toLabel + (message ? " · " + message : "");
  }
  if (eventType === "REVIEW_SAVE" || eventType === "RECALCULATE") {
    return message || eventTypeLabel_(eventType);
  }
  return message || eventTypeLabel_(eventType);
}

function buildAdminTimelineRows_(rows) {
  var timelineRows = Array.isArray(rows) ? rows : [];
  return timelineRows.map(function(row) {
    return {
      event_id: row.event_id,
      request_id: row.request_id,
      event_at: row.event_at,
      actor_type: row.actor_type,
      actor_id: row.actor_id,
      event_type: row.event_type,
      event_type_label: eventTypeLabel_(row.event_type),
      from_status: row.from_status,
      from_status_label: row.from_status ? statusLabel_(row.from_status) : "",
      to_status: row.to_status,
      to_status_label: row.to_status ? statusLabel_(row.to_status) : "",
      message: row.message,
      message_display: buildTimelineDisplayMessage_(row),
      payload: safeJsonParse_(row.payload_json, {})
    };
  });
}

function buildRequestDisplayForAdmin_(requestRow, answerIndex, review, autoEstimate, finalEstimate, visibleRecommendations, optionIndex) {
  var request = requestRow || {};
  var projectType = String(request.project_type || "").trim().toUpperCase();
  var housingQuestion = questionCodeForProject_(projectType, "R001_HOUSING_TYPE", "C001_BIZ_TYPE");
  var areaQuestion = questionCodeForProject_(projectType, "R002_AREA", "C002_AREA");
  var scopeQuestion = questionCodeForProject_(projectType, "R011_SCOPE_LEVEL", "C011_SCOPE_LEVEL");
  var startQuestion = questionCodeForProject_(projectType, "R320_START", "C320_START");
  var endQuestion = questionCodeForProject_(projectType, "R321_END", "C321_END");

  var reviewStatus = review ? review.review_status : "AUTO";
  var estimateSourceCode = String(request.final_estimate_source || (review && review.has_estimate_override ? "ADMIN" : (reviewStatus === "IN_REVIEW" ? "IN_REVIEW" : "AUTO"))).trim().toUpperCase();
  var estimateSourceLabel = estimateSourceCode === "ADMIN"
    ? "관리자 수정값"
    : (estimateSourceCode === "IN_REVIEW" ? "관리자 검토중" : "자동 산출값");
  var recommendationSourceLabel = review && review.has_recommendation_override
    ? "관리자 수정값"
    : "자동 추천값";

  return Object.assign({}, request, {
    status_label: statusLabel_(request.status),
    review_status: reviewStatus,
    review_status_label: reviewStatusLabel_(reviewStatus),
    project_type_label: getAnswerDisplayByCodes_(answerIndex, ["Q000_PROJECT_TYPE"], lookupSurveyOptionLabel_("Q000_PROJECT_TYPE", request.project_type, optionIndex) || (projectType === "COMMERCIAL" ? "상업" : "주거")),
    flow_mode_label: getAnswerDisplayByCodes_(answerIndex, ["Q000_FLOW_MODE"], lookupSurveyOptionLabel_("Q000_FLOW_MODE", request.flow_mode, optionIndex) || request.flow_mode),
    housing_type_label: getAnswerDisplayByCodes_(answerIndex, [housingQuestion], lookupSurveyOptionLabel_(housingQuestion, request.housing_type, optionIndex) || request.housing_type),
    area_label: getAnswerDisplayByCodes_(answerIndex, [areaQuestion], lookupSurveyOptionLabel_(areaQuestion, request.area_py, optionIndex) || request.area_py),
    scope_level_label: getAnswerDisplayByCodes_(answerIndex, [scopeQuestion], lookupSurveyOptionLabel_(scopeQuestion, request.scope_level, optionIndex) || request.scope_level),
    preferred_contact_method_label: getAnswerDisplayByCodes_(answerIndex, ["Q903_CONTACT_METHOD"], lookupSurveyOptionLabel_("Q903_CONTACT_METHOD", request.preferred_contact_method, optionIndex) || request.preferred_contact_method),
    schedule_start_label: getAnswerDisplayByCodes_(answerIndex, [startQuestion], request.schedule_start),
    schedule_end_label: getAnswerDisplayByCodes_(answerIndex, [endQuestion], request.schedule_end),
    customer_note_label: getAnswerDisplayByCodes_(answerIndex, ["Q904_NOTE"], request.customer_note),
    address_label: getAnswerDisplayByCodes_(answerIndex, ["Q906_ADDRESS"], request.address_text),
    priority_label: priorityLabel_(request.priority),
    work_state_label: workStateLabel_(request.work_state),
    estimate_source_label: estimateSourceLabel,
    estimate_source_code: estimateSourceCode,
    recommendation_source_label: recommendationSourceLabel,
    final_estimate_min: toNumber_(request.final_estimate_min, toNumber_((finalEstimate || {}).min, toNumber_(request.estimate_min, 0))),
    final_estimate_max: toNumber_(request.final_estimate_max, toNumber_((finalEstimate || {}).max, toNumber_(request.estimate_max, 0))),
    auto_estimate_min: toNumber_((autoEstimate || {}).min, toNumber_(request.estimate_min, 0)),
    auto_estimate_max: toNumber_((autoEstimate || {}).max, toNumber_(request.estimate_max, 0)),
    visible_recommendation_count: toNumber_(request.final_recommendation_count, (visibleRecommendations || []).length),
    is_manual_estimate_override: estimateSourceCode === "ADMIN",
    next_action_summary: buildNextActionSummary_(request),
    reminder_at_label: formatAdminDateTime_(request.reminder_at),
    latest_review_at_label: formatAdminDateTime_(request.latest_review_at)
  });
}

function buildDashboardRequestSummary_(requestRow, review, optionIndex) {
  var baseEstimate = buildStoredEstimateFromRequest_(requestRow);
  var finalEstimate = buildFinalEstimateForAdmin_(baseEstimate, review);
  var projectType = String((requestRow || {}).project_type || "").trim().toUpperCase();
  var areaQuestion = questionCodeForProject_(projectType, "R002_AREA", "C002_AREA");
  var scopeQuestion = questionCodeForProject_(projectType, "R011_SCOPE_LEVEL", "C011_SCOPE_LEVEL");
  var housingQuestion = questionCodeForProject_(projectType, "R001_HOUSING_TYPE", "C001_BIZ_TYPE");
  var reviewStatus = review ? review.review_status : reviewStatusCodeFromLabel_((requestRow || {}).review_status_label);
  var estimateSourceCode = String(
    (requestRow || {}).final_estimate_source ||
    (review && review.has_estimate_override ? "ADMIN" : (reviewStatus === "IN_REVIEW" ? "IN_REVIEW" : "AUTO"))
  ).trim().toUpperCase() || "AUTO";

  return {
    request_id: requestRow.request_id,
    customer_name: requestRow.customer_name,
    contact_phone: requestRow.contact_phone,
    project_type: requestRow.project_type,
    project_type_label: lookupSurveyOptionLabel_("Q000_PROJECT_TYPE", requestRow.project_type, optionIndex) || (projectType === "COMMERCIAL" ? "상업" : "주거"),
    housing_type: requestRow.housing_type,
    housing_type_label: lookupSurveyOptionLabel_(housingQuestion, requestRow.housing_type, optionIndex) || String(requestRow.housing_type || "").trim() || "-",
    area_py: requestRow.area_py,
    area_label: lookupSurveyOptionLabel_(areaQuestion, requestRow.area_py, optionIndex) || String(requestRow.area_py || "").trim() || "-",
    scope_level: requestRow.scope_level,
    scope_level_label: lookupSurveyOptionLabel_(scopeQuestion, requestRow.scope_level, optionIndex) || String(requestRow.scope_level || "").trim() || "-",
    flow_mode: requestRow.flow_mode,
    flow_mode_label: lookupSurveyOptionLabel_("Q000_FLOW_MODE", requestRow.flow_mode, optionIndex) || String(requestRow.flow_mode || "").trim() || "-",
    preferred_contact_method: requestRow.preferred_contact_method,
    preferred_contact_method_label: lookupSurveyOptionLabel_("Q903_CONTACT_METHOD", requestRow.preferred_contact_method, optionIndex) || String(requestRow.preferred_contact_method || "").trim() || "-",
    estimate_min: toNumber_(requestRow.final_estimate_min, toNumber_(finalEstimate.min, toNumber_(requestRow.estimate_min, 0))),
    estimate_max: toNumber_(requestRow.final_estimate_max, toNumber_(finalEstimate.max, toNumber_(requestRow.estimate_max, 0))),
    auto_estimate_min: toNumber_(baseEstimate.min, toNumber_(requestRow.estimate_min, 0)),
    auto_estimate_max: toNumber_(baseEstimate.max, toNumber_(requestRow.estimate_max, 0)),
    status: requestRow.status,
    status_label: statusLabel_(requestRow.status),
    review_status: reviewStatus,
    review_status_label: reviewStatusLabel_(reviewStatus),
    estimate_source_label: estimateSourceCode === "ADMIN"
      ? "관리자 수정값"
      : (estimateSourceCode === "IN_REVIEW" ? "관리자 검토중" : "자동 산출값"),
    estimate_source_code: estimateSourceCode,
    is_manual_override: estimateSourceCode === "ADMIN",
    priority: requestRow.priority || "NORMAL",
    priority_label: priorityLabel_(requestRow.priority),
    assignee_name: String(requestRow.assignee_name || "").trim(),
    assignee_email: String(requestRow.assignee_email || "").trim(),
    next_action_type: String(requestRow.next_action_type || "").trim(),
    next_action_note: String(requestRow.next_action_note || "").trim(),
    next_action_due_at: String(requestRow.next_action_due_at || "").trim(),
    next_action_summary: buildNextActionSummary_(requestRow),
    reminder_at: String(requestRow.reminder_at || "").trim(),
    work_state: String(requestRow.work_state || "NEW").trim().toUpperCase(),
    work_state_label: workStateLabel_(requestRow.work_state),
    duplicate_customer_count: toNumber_(requestRow.duplicate_customer_count, 0),
    quote_draft_status: String(requestRow.quote_draft_status || "").trim(),
    created_at: requestRow.created_at,
    schedule_start: requestRow.schedule_start,
    latest_review_at: requestRow.latest_review_at || "",
    latest_review_id: requestRow.latest_review_id || ""
  };
}

function buildAdminRequestDetailPayload_(found) {
  var requestFound = found || {};
  var requestRow = requestFound.data || {};
  var rid = String(requestRow.request_id || "").trim();
  if (!rid) throw new Error("Request not found");

  var answerRows = readRequestRowsIndexed_("RequestAnswers", rid).sort(function(a, b) {
    return String(a.answer_id || "").localeCompare(String(b.answer_id || ""));
  });
  var storedRecommendations = loadRequestRecommendationState_(requestFound, answerRows, {
    include_hidden: true,
    include_deleted: true
  });
  var autoRecommendations = storedRecommendations.length
    ? storedRecommendations.slice()
    : generateAutoRecommendationsForRequest_(requestRow, answerRows);
  var timelineRows = readRequestRowsIndexed_("RequestTimeline", rid).sort(function(a, b) {
    return String(b.event_at || "").localeCompare(String(a.event_at || ""));
  });
  var files = readRequestRowsIndexed_("RequestFiles", rid).sort(function(a, b) {
    return Number(a.sort_order || 0) - Number(b.sort_order || 0);
  });
  var latestReview = getLatestAdminReviewForRequest_(rid);
  var questionMetaMap = getSurveyQuestionMetaMap_();
  var optionIndex = getSurveyOptionLabelIndex_();
  var enrichedAnswers = answerRows
    .filter(function(row) { return String(row.question_code || "").trim() !== "Q905_FILES"; })
    .map(function(row) {
      return enrichAdminAnswerRow_(row, latestReview ? latestReview.override_answers : {}, questionMetaMap, optionIndex);
    });
  var answerGroups = groupAdminAnswers_(enrichedAnswers);
  var answerIndex = indexAdminAnswersByCode_(enrichedAnswers);
  var autoEstimate = buildStoredEstimateFromRequest_(requestRow);
  var finalEstimate = buildFinalEstimateForAdmin_(autoEstimate, latestReview);
  var recommendationEditorItems = latestReview && latestReview.has_recommendation_override
    ? sanitizeReviewRecommendations_(latestReview.override_recommendations || [])
    : sanitizeReviewRecommendations_(storedRecommendations.length ? storedRecommendations : autoRecommendations);
  var visibleRecommendations = recommendationEditorItems.filter(function(item) {
    return item.is_visible !== false && !item.is_deleted;
  });
  var requestView = buildRequestDisplayForAdmin_(requestRow, answerIndex, latestReview, autoEstimate, finalEstimate, visibleRecommendations, optionIndex);
  var assignment = buildActiveAssignmentViewModel_(requestRow);
  var tasks = listRequestTasksForAdmin_(rid);
  var checklist = buildChecklistItemsForAdmin_(requestRow);
  var duplicates = collectDuplicateRequestHistory_(requestRow, 8);
  var latestDraftLink = getLatestQuoteDraftLinkForRequest_(rid);
  var consultationSummary = buildConsultationSummaryPayload_({
    request: requestView,
    answers: enrichedAnswers,
    answer_groups: answerGroups,
    recommendations: visibleRecommendations,
    flags: normalizeRiskFlagsForAdmin_(safeJsonParse_(requestRow.risk_flags_json, [])),
    assignment: assignment,
    tasks: tasks,
    checklist: checklist
  });

  return {
    request: requestView,
    estimate: finalEstimate,
    auto_estimate: autoEstimate,
    flags: normalizeRiskFlagsForAdmin_(safeJsonParse_(requestRow.risk_flags_json, [])),
    tags: safeJsonParse_(requestRow.tags_json, []),
    answers: enrichedAnswers,
    answer_groups: answerGroups,
    recommendations: visibleRecommendations,
    auto_recommendations: autoRecommendations.map(function(item) {
      return normalizeAdminRecommendationItem_(item, item.rank || 0, "AUTO");
    }),
    recommendation_editor_items: recommendationEditorItems,
    timeline: buildAdminTimelineRows_(timelineRows),
    files: files,
    review: buildAdminReviewViewModel_(latestReview, autoEstimate, finalEstimate, recommendationEditorItems),
    assignment: assignment,
    tasks: tasks,
    checklist: checklist,
    duplicate_history: duplicates,
    consultation_summary: consultationSummary,
    latest_quote_draft: latestDraftLink
  };
}

function getRequestDetailCacheTtlSec_() {
  return Math.max(Number(getSettings_().perf_request_detail_cache_ttl_sec || 0), 0);
}

function buildRequestDetailCacheKey_(found) {
  var requestFound = found || {};
  var requestRow = requestFound.data || {};
  var rid = String(requestRow.request_id || "").trim();
  if (!rid) return "";
  var record = readRequestRowIndexRecord_(rid) || {};
  return [
    "admin-detail",
    rid,
    String(requestRow.updated_at || requestRow.created_at || ""),
    String(requestRow.latest_review_id || ""),
    String(record.updated_at || ""),
    String(getRequestRowNosFromIndex_(record, "RequestTimeline").length),
    String(getRequestRowNosFromIndex_(record, "RequestChecklistItems").length),
    String(getRequestRowNosFromIndex_(record, "RequestTasks").length),
    String(getRequestRowNosFromIndex_(record, "RequestFiles").length),
    String(getRequestRowNosFromIndex_(record, REQUEST_ADMIN_REVIEW_SHEET_).length)
  ].join("|");
}


function buildTimelineEvent_(requestId, actorType, eventType, fromStatus, toStatus, message, payloadJson, eventAt) {
  return {
    event_id: uuid_(),
    request_id: requestId,
    event_at: eventAt || nowIso_(),
    actor_type: actorType || "",
    actor_id: "",
    event_type: eventType || "",
    from_status: fromStatus || "",
    to_status: toStatus || "",
    message: String(message || "").trim(),
    payload_json: payloadJson || "",
    event_label: eventTypeLabel_(eventType || "")
  };
}

function adminLogin(password) {
  ensureSpreadsheetId_();
  assertAdmin_(password);
  var tok = createAdminSessionToken_();
  try {
    CacheService.getScriptCache().put(getAdminSessionCacheKey_(tok), "1", ADMIN_SESSION_TTL_SEC_);
  } catch (e) {}
  return { token: tok };
}

function adminLogout(credential) {
  ensureSpreadsheetId_();
  assertAdminCredential_(credential);
  clearAdminSession_(credential);
  return { success: true };
}

function adminGetDashboard(credential) {
  ensureSpreadsheetId_();
  ensurePrequoteOperationalSchema_();
  var credInfo = assertAdminCredential_(credential);
  var startMs = Date.now();
  var requests = [];
  var result = null;
  var error = null;
  try {
    requests = readAllRows_("Requests");
    var settings = getSettings_();
    var optionIndex = getSurveyOptionLabelIndex_();

    var now = new Date();
    var recentDays = Number(settings.dashboard_recent_days || 30);
    var cutoff = new Date(now.getTime() - recentDays * 86400000).toISOString();

    var recent = requests.filter(function(r) { return String(r.created_at || "") >= cutoff; });
    var byStatus = {};
    recent.forEach(function(r) {
      var st = String(r.status || "NEW");
      byStatus[st] = (byStatus[st] || 0) + 1;
    });

    requests.sort(function(a, b) {
      return String(b.created_at || "").localeCompare(String(a.created_at || ""));
    });
    var queue = buildDashboardQueueSummary_(requests);

    result = {
      total_requests: requests.length,
      recent_count: recent.length,
      recent_days: recentDays,
      by_status: byStatus,
      queue: queue,
      requests: requests.slice(0, 200).map(function(r) {
        return buildDashboardRequestSummary_(r, null, optionIndex);
      })
    };
    return result;
  } catch (e) {
    error = e;
    throw e;
  } finally {
    logPerfMetricSafe_(
      "dashboard_load",
      Date.now() - startMs,
      "",
      { request_count: requests.length, returned_count: result && result.requests ? result.requests.length : 0 },
      !error,
      error ? String(error) : "",
      "ADMIN",
      credInfo.type
    );
  }
}

function adminGetRequestDetail(credential, requestId) {
  ensureSpreadsheetId_();
  ensurePrequoteOperationalSchema_();
  var credInfo = assertAdminCredential_(credential);
  var rid = String(requestId || "").trim();
  var startMs = Date.now();
  var error = null;
  var cacheHit = false;
  try {
    var found = findRequestRowById_(rid);
    if (!found) throw new Error("Request not found");
    var cacheTtl = getRequestDetailCacheTtlSec_();
    var cacheKey = cacheTtl ? buildRequestDetailCacheKey_(found) : "";
    if (cacheKey) {
      var cached = cacheJsonGet_(cacheKey);
      if (cached) {
        cacheHit = true;
        return cached;
      }
    }
    var detail = buildAdminRequestDetailPayload_(found);
    if (cacheKey && cacheTtl) cacheJsonPut_(cacheKey, detail, cacheTtl);
    return detail;
  } catch (e) {
    error = e;
    throw e;
  } finally {
    logPerfMetricSafe_(
      "request_detail_load",
      Date.now() - startMs,
      rid,
      { cache_hit: cacheHit },
      !error,
      error ? String(error) : "",
      "ADMIN",
      credInfo.type
    );
  }
}

function adminSaveRequestReview(credential, requestId, payload) {
  ensureSpreadsheetId_();
  ensurePrequoteOperationalSchema_();
  var credInfo = assertAdminCredential_(credential);
  ensureAdminReviewSheet_();
  var startMs = Date.now();
  var error = null;
  var result = null;

  try {
    var found = findRequestRowById_(String(requestId || "").trim());
    if (!found) throw new Error("Request not found");

    var reviewInput = payload && typeof payload === "object" ? payload : {};
    var returnDetail = reviewInput.return_detail !== false;
    var previousRequestRow = JSON.parse(JSON.stringify(found.data || {}));
    var rid = found.data.request_id;
    var answerRows = readRequestRowsIndexed_("RequestAnswers", rid).sort(function(a, b) {
      return String(a.answer_id || "").localeCompare(String(b.answer_id || ""));
    });
    var baseAnswerMap = buildAnswerMapFromRows_(answerRows);
    var autoEstimate = buildStoredEstimateFromRequest_(found.data);
    var autoRecommendationSource = loadRequestRecommendationState_(found, answerRows, {
      include_hidden: true,
      include_deleted: true,
      allow_generate: true
    });
    var autoRecommendations = autoRecommendationSource.map(function(item) {
      return normalizeAdminRecommendationItem_(item, item.rank || 0, "AUTO");
    });
    var currentRecommendationState = loadRequestRecommendationState_(found, answerRows, {
      include_hidden: true,
      include_deleted: true
    }).map(function(item) {
      return normalizeAdminRecommendationItem_(item, item.rank || 0, item.source || item.source_type || "ADMIN");
    });

  var overrideEstimateMin = reviewInput.override_estimate_min;
  var overrideEstimateMax = reviewInput.override_estimate_max;
  var hasEstimateInput = !(overrideEstimateMin === "" || overrideEstimateMin === null || overrideEstimateMin === undefined)
    || !(overrideEstimateMax === "" || overrideEstimateMax === null || overrideEstimateMax === undefined);
  var hasEstimateOverride = false;
  if (hasEstimateInput) {
    if (overrideEstimateMin === "" || overrideEstimateMin === null || overrideEstimateMin === undefined ||
        overrideEstimateMax === "" || overrideEstimateMax === null || overrideEstimateMax === undefined) {
      throw new Error("가견적 최소/최대 금액을 모두 입력해주세요.");
    }
    overrideEstimateMin = Number(overrideEstimateMin);
    overrideEstimateMax = Number(overrideEstimateMax);
    if (!isFinite(overrideEstimateMin) || !isFinite(overrideEstimateMax)) {
      throw new Error("가견적 금액은 숫자로 입력해주세요.");
    }
    if (overrideEstimateMin > overrideEstimateMax) {
      throw new Error("최소 금액이 최대 금액보다 클 수 없습니다.");
    }
    hasEstimateOverride = overrideEstimateMin !== toNumber_(autoEstimate.min, 0) || overrideEstimateMax !== toNumber_(autoEstimate.max, 0);
  }

  var submittedRecommendations = Array.isArray(reviewInput.recommendations)
    ? reviewInput.recommendations
    : (Array.isArray(reviewInput.override_recommendations) ? reviewInput.override_recommendations : null);
  var sanitizedRecommendations = submittedRecommendations
    ? sanitizeReviewRecommendations_(submittedRecommendations)
    : currentRecommendationState;
  var hasRecommendationOverride = submittedRecommendations !== null && !recommendationListsEqual_(sanitizedRecommendations, autoRecommendations);
  var sanitizedAnswerOverrides = sanitizeReviewAnswers_(reviewInput.override_answers, baseAnswerMap);
  var hasAnswerOverride = Object.keys(sanitizedAnswerOverrides).length > 0;
  var opsPayload = normalizeRequestOpsPayload_(reviewInput, found.data);
  var requestedStatus = String(reviewInput.status || (reviewInput.ops && reviewInput.ops.status) || "").trim().toUpperCase();
  if (requestedStatus && REQUEST_STATUS_OPTIONS_.indexOf(requestedStatus) < 0) {
    throw new Error("Invalid status");
  }
  var targetStatus = requestedStatus || String(found.data.status || "NEW").trim().toUpperCase();

  var defaultReviewStatus = hasEstimateOverride || hasRecommendationOverride || hasAnswerOverride
    ? "OVERRIDDEN"
    : (String(reviewInput.review_note || "").trim() ? "IN_REVIEW" : "AUTO");
  var reviewStatus = normalizeAdminReviewStatus_(reviewInput.review_status, defaultReviewStatus);
  if (reviewStatus === "AUTO" && (hasEstimateOverride || hasRecommendationOverride || hasAnswerOverride)) {
    reviewStatus = "OVERRIDDEN";
  }

  var draftReview = buildAdminReviewRow_(rid, {
    edited_at: nowIso_(),
    review_status: reviewStatus,
    override_estimate_min: hasEstimateOverride ? overrideEstimateMin : "",
    override_estimate_max: hasEstimateOverride ? overrideEstimateMax : "",
    override_recommendations: hasRecommendationOverride ? sanitizedRecommendations : [],
    override_answers: sanitizedAnswerOverrides,
    has_recommendation_override: hasRecommendationOverride,
    has_answer_override: hasAnswerOverride,
    review_note: String(reviewInput.review_note || "").trim(),
    payload: {
      estimate_reason: String(reviewInput.estimate_reason || "").trim(),
      saved_from: "adminSaveRequestReview"
    }
  });
  var finalEstimate = buildFinalEstimateForAdmin_(autoEstimate, draftReview);
  var recommendationEditorItems = hasRecommendationOverride
    ? sanitizeReviewRecommendations_(sanitizedRecommendations)
    : sanitizeReviewRecommendations_(currentRecommendationState.length ? currentRecommendationState : autoRecommendations);
  var visibleRecommendations = recommendationEditorItems.filter(function(item) {
    return item.is_visible !== false && !item.is_deleted;
  });
  draftReview.payload = buildAdminReviewPayloadSnapshot_(draftReview, autoEstimate, finalEstimate, recommendationEditorItems);

  var reviewRow = buildAdminReviewRow_(rid, {
    edited_at: draftReview.edited_at,
    review_status: draftReview.review_status,
    override_estimate_min: draftReview.has_estimate_override ? draftReview.override_estimate_min : "",
    override_estimate_max: draftReview.has_estimate_override ? draftReview.override_estimate_max : "",
    override_recommendations: draftReview.has_recommendation_override ? draftReview.override_recommendations : [],
    override_answers: draftReview.override_answers,
    has_recommendation_override: draftReview.has_recommendation_override,
    has_answer_override: draftReview.has_answer_override,
    review_note: draftReview.review_note,
    payload: Object.assign({}, draftReview.payload, {
      estimate_reason: String(reviewInput.estimate_reason || "").trim(),
      saved_from: "adminSaveRequestReview"
    })
  });
  reviewRow.review_summary = summarizeAdminReviewSave_(reviewRow);
  reviewRow.final_estimate_min = toNumber_(finalEstimate.min, 0);
  reviewRow.final_estimate_max = toNumber_(finalEstimate.max, 0);
  reviewRow.final_recommendation_count = visibleRecommendations.length;
  reviewRow.source_label = reviewRow.has_estimate_override ? "관리자 수정값" : (reviewRow.review_status === "IN_REVIEW" ? "관리자 검토중" : "자동 산출값");
  reviewRow.linked_quote_id = String(found.data.linked_quote_id || "").trim();
  reviewRow.linked_quote_status = String(found.data.quote_draft_status || "").trim();
  reviewRow.created_at = reviewRow.created_at || reviewRow.edited_at;
  reviewRow.updated_at = reviewRow.edited_at;
  reviewRow.is_current = true;

  var now = reviewRow.edited_at;
  markAdminReviewsAsNotCurrent_(rid);
  saveRequestRecommendationState_(found, recommendationEditorItems, reviewRow.review_id);
  upsertRequestAssignment_(rid, opsPayload, now);
  upsertRequestNextActionTask_(rid, opsPayload, now);
  saveChecklistItems_(rid, reviewInput.checklist_items, now);

  appendRow_(REQUEST_ADMIN_REVIEW_SHEET_, {
    review_id: reviewRow.review_id,
    request_id: reviewRow.request_id,
    edited_at: reviewRow.edited_at,
    actor_type: reviewRow.actor_type,
    actor_id: reviewRow.actor_id,
    review_status: reviewRow.review_status,
    override_estimate_min: reviewRow.override_estimate_min,
    override_estimate_max: reviewRow.override_estimate_max,
    override_recommendations_json: reviewRow.has_recommendation_override ? JSON.stringify(reviewRow.override_recommendations) : "",
    override_answers_json: reviewRow.has_answer_override ? JSON.stringify(reviewRow.override_answers) : "",
    estimate_reason: reviewRow.estimate_reason || String((reviewRow.payload || {}).estimate_reason || "").trim(),
    review_note: reviewRow.review_note,
    review_summary: reviewRow.review_summary,
    final_estimate_min: reviewRow.final_estimate_min,
    final_estimate_max: reviewRow.final_estimate_max,
    final_recommendation_count: reviewRow.final_recommendation_count,
    source_label: reviewRow.source_label,
    is_current: "Y",
    linked_quote_id: reviewRow.linked_quote_id,
    linked_quote_status: reviewRow.linked_quote_status,
    slack_notified_at: "",
    created_at: reviewRow.created_at,
    updated_at: reviewRow.updated_at,
    payload_json: JSON.stringify(Object.assign({}, reviewRow.payload || {}, {
      ops: opsPayload,
      review_summary: reviewRow.review_summary,
      source_label: reviewRow.source_label
    }))
  });

  var requestPatch = buildRequestReviewPatch_(found.data, reviewRow, finalEstimate, visibleRecommendations, opsPayload, now);
  requestPatch.status = targetStatus;
  updateRowFields_(found.sheet, found.rowNo, found.meta, requestPatch);
  for (var patchKey in requestPatch) {
    if (!Object.prototype.hasOwnProperty.call(requestPatch, patchKey)) continue;
    found.data[patchKey] = requestPatch[patchKey];
  }

  var statusTimelineEvent = null;
  var timelineRowsToAppend = [];
  if (targetStatus !== String(previousRequestRow.status || "NEW").trim().toUpperCase()) {
    statusTimelineEvent = buildTimelineEvent_(rid, "ADMIN", "STATUS_CHANGE", previousRequestRow.status, targetStatus, "", "", now);
    statusTimelineEvent.event_label = eventTypeLabel_("STATUS_CHANGE");
    timelineRowsToAppend.push(statusTimelineEvent);
  }
  var timelineEvent = buildTimelineEvent_(rid, "ADMIN", "REVIEW_SAVE", "", "", reviewRow.review_summary, JSON.stringify(reviewRow.payload || {}), now);
  timelineEvent.event_label = eventTypeLabel_("REVIEW_SAVE");
  timelineRowsToAppend.push(timelineEvent);
  appendRows_("RequestTimeline", timelineRowsToAppend);
  maybeNotifyAssigneeChange_(previousRequestRow, found.data, reviewRow);

    result = {
      success: true,
      saved_review: reviewRow,
      timeline_event: timelineEvent,
      status_timeline_event: statusTimelineEvent,
      dashboard_patch: buildDashboardRequestSummary_(found.data, reviewRow, getSurveyOptionLabelIndex_()),
      detail: returnDetail ? buildAdminRequestDetailPayload_(found) : null
    };
    return result;
  } catch (e) {
    error = e;
    throw e;
  } finally {
    logPerfMetricSafe_(
      "request_review_save",
      Date.now() - startMs,
      String(requestId || "").trim(),
      {
        success: !error,
        has_status_change: !!(result && result.status_timeline_event),
        recommendation_count: result && result.detail && result.detail.recommendations ? result.detail.recommendations.length : 0
      },
      !error,
      error ? String(error) : "",
      "ADMIN",
      credInfo.type
    );
  }
}

function adminRecalculateRequest(credential, requestId, payload) {
  ensureSpreadsheetId_();
  ensurePrequoteOperationalSchema_();
  assertAdminCredential_(credential);
  var found = findRequestRowById_(String(requestId || "").trim());
  if (!found) throw new Error("Request not found");

  var reviewInput = payload && typeof payload === "object" ? payload : {};
  var rid = found.data.request_id;
  var answerRows = readRequestRowsIndexed_("RequestAnswers", rid).sort(function(a, b) {
    return String(a.answer_id || "").localeCompare(String(b.answer_id || ""));
  });
  var baseAnswerMap = buildAnswerMapFromRows_(answerRows);
  var overrideAnswers = sanitizeReviewAnswers_(reviewInput.override_answers, baseAnswerMap);
  var mergedAnswers = {};
  for (var key in baseAnswerMap) {
    if (Object.prototype.hasOwnProperty.call(baseAnswerMap, key)) mergedAnswers[key] = baseAnswerMap[key];
  }
  for (var overrideKey in overrideAnswers) {
    if (Object.prototype.hasOwnProperty.call(overrideAnswers, overrideKey)) mergedAnswers[overrideKey] = overrideAnswers[overrideKey];
  }

  var projectType = String(found.data.project_type || mergedAnswers.Q000_PROJECT_TYPE || "").trim().toUpperCase();
  var tags = calculateTags_(mergedAnswers);
  var estimate = calculateEstimate_(mergedAnswers, projectType);
  var recommendations = getRecommendations_(tags, mergedAnswers, projectType).map(function(item, index) {
    return normalizeAdminRecommendationItem_(item, index + 1, "AUTO");
  });
  var now = nowIso_();
  updateRowFields_(found.sheet, found.rowNo, found.meta, {
    updated_at: now
  });
  found.data.updated_at = now;
  var timelinePayload = {
    override_answer_keys: Object.keys(overrideAnswers),
    recommendation_count: recommendations.length,
    estimate_min: estimate.min,
    estimate_max: estimate.max
  };
  var timelineEvent = buildTimelineEvent_(rid, "ADMIN", "RECALCULATE", "", "", "자동 산출값을 다시 계산했습니다.", JSON.stringify(timelinePayload), now);
  appendRow_("RequestTimeline", timelineEvent);

  return {
    success: true,
    estimate: Object.assign({}, estimate, { source: "AUTO", source_label: "자동 산출값" }),
    recommendations: recommendations,
    override_answers: overrideAnswers,
    timeline_event: timelineEvent
  };
}

function adminUpdateRequestStatus(credential, requestId, newStatus, note) {
  ensureSpreadsheetId_();
  ensurePrequoteOperationalSchema_();
  assertAdminCredential_(credential);
  var found = findRequestRowById_(String(requestId || "").trim());
  if (!found) throw new Error("Request not found");

  var targetStatus = String(newStatus || "").trim().toUpperCase();
  if (REQUEST_STATUS_OPTIONS_.indexOf(targetStatus) < 0) throw new Error("Invalid status");
  var oldStatus = String(found.data.status || "NEW");
  var now = nowIso_();
  var noteText = String(note || "").trim();

  updateRowFields_(found.sheet, found.rowNo, found.meta, {
    status: targetStatus,
    updated_at: now
  });
  found.data.status = targetStatus;
  found.data.updated_at = now;

  var event = buildTimelineEvent_(found.data.request_id, "ADMIN", "STATUS_CHANGE", oldStatus, targetStatus, noteText, "", now);
  appendRow_("RequestTimeline", event);

  return {
    success: true,
    old_status: oldStatus,
    new_status: targetStatus,
    updated_at: now,
    timeline_event: event,
    dashboard_patch: buildDashboardRequestSummary_(found.data, null, getSurveyOptionLabelIndex_()),
    detail: buildAdminRequestDetailPayload_(found)
  };
}

function adminAddNote(credential, requestId, noteText, options) {
  ensureSpreadsheetId_();
  ensurePrequoteOperationalSchema_();
  assertAdminCredential_(credential);
  var found = findRequestRowById_(String(requestId || "").trim());
  if (!found) throw new Error("Request not found");
  var opts = options && typeof options === "object" ? options : {};
  var returnDetail = opts.return_detail !== false;
  var text = String(noteText || "").trim();
  if (!text) throw new Error("Note text required");
  var now = nowIso_();
  updateRowFields_(found.sheet, found.rowNo, found.meta, {
    updated_at: now
  });
  found.data.updated_at = now;
  var event = buildTimelineEvent_(found.data.request_id, "ADMIN", "ADMIN_NOTE", "", "", text, "", now);
  appendRow_("RequestTimeline", event);
  return {
    success: true,
    updated_at: now,
    timeline_event: event,
    dashboard_patch: buildDashboardRequestSummary_(found.data, null, getSurveyOptionLabelIndex_()),
    detail: returnDetail ? buildAdminRequestDetailPayload_(found) : null
  };
}

function normalizePriority_(value, fallback) {
  var code = String(value || fallback || "NORMAL").trim().toUpperCase();
  return REQUEST_PRIORITY_OPTIONS_.indexOf(code) >= 0 ? code : String(fallback || "NORMAL").trim().toUpperCase();
}

function normalizeWorkState_(value, fallback) {
  var code = String(value || fallback || "NEW").trim().toUpperCase();
  return REQUEST_WORK_STATE_OPTIONS_.indexOf(code) >= 0 ? code : String(fallback || "NEW").trim().toUpperCase();
}

function normalizeTaskStatus_(value, fallback) {
  var code = String(value || fallback || "OPEN").trim().toUpperCase();
  if (TASK_STATUS_LABEL_MAP_[code]) return code;
  return String(fallback || "OPEN").trim().toUpperCase();
}

function normalizePhoneKey_(value) {
  var digits = String(value || "").replace(/\D+/g, "");
  if (digits.length < 7) return "";
  return digits;
}

function normalizeEmailKey_(value) {
  return String(value || "").trim().toLowerCase();
}

function normalizeNameKey_(value) {
  return String(value || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function isDuplicateMatch_(requestRow, phoneKey, emailKey, nameKey, selfRequestId) {
  var row = requestRow || {};
  if (selfRequestId && String(row.request_id || "") === String(selfRequestId || "")) return false;
  var rowPhoneKey = String(row.duplicate_phone_key || "").trim() || normalizePhoneKey_(row.contact_phone);
  var rowEmailKey = String(row.duplicate_email_key || "").trim() || normalizeEmailKey_(row.contact_email);
  var rowNameKey = normalizeNameKey_(row.customer_name);
  if (phoneKey && rowPhoneKey && phoneKey === rowPhoneKey) return true;
  if (emailKey && rowEmailKey && emailKey === rowEmailKey) return true;
  return !!(nameKey && rowNameKey && nameKey === rowNameKey);
}

function getRequestDuplicateLookup_(overrideId) {
  var ssId = String(getSpreadsheetId_(overrideId) || "").trim();
  if (!ssId) return { rows: [], phone: Object.create(null), email: Object.create(null), name: Object.create(null) };
  if (__REQUEST_DUPLICATE_LOOKUP_MEMO_[ssId]) return __REQUEST_DUPLICATE_LOOKUP_MEMO_[ssId];
  var rows = readAllRows_("Requests", overrideId);
  var lookup = {
    rows: rows,
    phone: Object.create(null),
    email: Object.create(null),
    name: Object.create(null)
  };
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    var phoneKey = String(row.duplicate_phone_key || "").trim() || normalizePhoneKey_(row.contact_phone);
    var emailKey = String(row.duplicate_email_key || "").trim() || normalizeEmailKey_(row.contact_email);
    var nameKey = normalizeNameKey_(row.customer_name);
    if (phoneKey) {
      if (!lookup.phone[phoneKey]) lookup.phone[phoneKey] = [];
      lookup.phone[phoneKey].push(row);
    }
    if (emailKey) {
      if (!lookup.email[emailKey]) lookup.email[emailKey] = [];
      lookup.email[emailKey].push(row);
    }
    if (nameKey) {
      if (!lookup.name[nameKey]) lookup.name[nameKey] = [];
      lookup.name[nameKey].push(row);
    }
  }
  __REQUEST_DUPLICATE_LOOKUP_MEMO_[ssId] = lookup;
  return lookup;
}

function findDuplicateRequests_(phoneKey, emailKey, customerName, selfRequestId, overrideId) {
  var nameKey = normalizeNameKey_(customerName);
  if (!phoneKey && !emailKey && !nameKey) return [];
  var lookup = getRequestDuplicateLookup_(overrideId);
  var out = [];
  var seen = Object.create(null);
  var buckets = [];
  if (phoneKey && lookup.phone[phoneKey]) buckets.push(lookup.phone[phoneKey]);
  if (emailKey && lookup.email[emailKey]) buckets.push(lookup.email[emailKey]);
  if (nameKey && lookup.name[nameKey]) buckets.push(lookup.name[nameKey]);
  for (var i = 0; i < buckets.length; i++) {
    var rows = buckets[i];
    for (var j = 0; j < rows.length; j++) {
      var row = rows[j] || {};
      var requestId = String(row.request_id || "").trim();
      var dedupRowKey = requestId || ("row:" + String(row._rowNo || ""));
      if (!dedupRowKey || seen[dedupRowKey]) continue;
      if (!isDuplicateMatch_(row, phoneKey, emailKey, nameKey, selfRequestId)) continue;
      seen[dedupRowKey] = 1;
      out.push(row);
    }
  }
  return out;
}

function countDuplicateRequests_(phoneKey, emailKey, customerName, selfRequestId) {
  return findDuplicateRequests_(phoneKey, emailKey, customerName, selfRequestId).length;
}

function normalizeRequestOpsPayload_(input, requestRow) {
  var row = requestRow || {};
  var raw = input && typeof input === "object" ? input : {};
  var ops = raw.ops && typeof raw.ops === "object" ? raw.ops : raw;
  var assignment = ops.assignment && typeof ops.assignment === "object" ? ops.assignment : {};
  var slackMap = getSlackUserMap_();
  var assigneeName = String(assignment.assignee_name || ops.assignee_name || row.assignee_name || "").trim();
  var assigneeEmail = String(assignment.assignee_email || ops.assignee_email || row.assignee_email || "").trim();
  var assigneeSlackId = String(assignment.assignee_slack_id || ops.assignee_slack_id || row.assignee_slack_id || "").trim();
  if (!assigneeSlackId && assigneeName && slackMap[assigneeName]) assigneeSlackId = String(slackMap[assigneeName] || "").trim();
  return {
    priority: normalizePriority_(ops.priority, row.priority || "NORMAL"),
    work_state: normalizeWorkState_(
      ops.work_state || ((raw.review_status === "IN_REVIEW" && !ops.work_state) ? "REVIEWING" : ""),
      row.work_state || "NEW"
    ),
    assignee_name: assigneeName,
    assignee_email: assigneeEmail,
    assignee_slack_id: assigneeSlackId,
    next_action_type: String(ops.next_action_type || row.next_action_type || "").trim(),
    next_action_note: String(ops.next_action_note || row.next_action_note || "").trim(),
    next_action_due_at: String(ops.next_action_due_at || row.next_action_due_at || "").trim(),
    reminder_at: String(ops.reminder_at || row.reminder_at || "").trim(),
    task_status: normalizeTaskStatus_(ops.task_status, row.work_state === "DONE" ? "DONE" : "OPEN")
  };
}

function buildRequestReviewPatch_(requestRow, reviewRow, finalEstimate, visibleRecommendations, opsPayload, updatedAt) {
  var row = requestRow || {};
  var ops = opsPayload || {};
  var phoneKey = String(row.duplicate_phone_key || "").trim() || normalizePhoneKey_(row.contact_phone);
  var emailKey = String(row.duplicate_email_key || "").trim() || normalizeEmailKey_(row.contact_email);
  return {
    updated_at: updatedAt || nowIso_(),
    priority: ops.priority,
    assignee_name: ops.assignee_name,
    assignee_email: ops.assignee_email,
    assignee_slack_id: ops.assignee_slack_id,
    next_action_type: ops.next_action_type,
    next_action_note: ops.next_action_note,
    next_action_due_at: ops.next_action_due_at,
    reminder_at: ops.reminder_at,
    work_state: ops.work_state,
    latest_review_id: reviewRow.review_id,
    latest_review_at: reviewRow.edited_at,
    final_estimate_min: toNumber_((finalEstimate || {}).min, 0),
    final_estimate_max: toNumber_((finalEstimate || {}).max, 0),
    final_estimate_source: reviewRow.has_estimate_override ? "ADMIN" : (reviewRow.review_status === "IN_REVIEW" ? "IN_REVIEW" : "AUTO"),
    final_recommendation_count: (visibleRecommendations || []).length,
    review_status_label: reviewStatusLabel_(reviewRow.review_status),
    duplicate_phone_key: phoneKey,
    duplicate_email_key: emailKey,
    duplicate_customer_count: countDuplicateRequests_(phoneKey, emailKey, row.customer_name, row.request_id),
    quote_draft_status: String(row.quote_draft_status || "PENDING").trim()
  };
}

function markAdminReviewsAsNotCurrent_(requestId) {
  var target = String(requestId || "").trim();
  if (!target) return;
  var rowNos = getRequestRowNosFromIndex_(readRequestRowIndexRecord_(target), REQUEST_ADMIN_REVIEW_SHEET_);
  if (!rowNos.length) rowNos = findRowNosByCol_(REQUEST_ADMIN_REVIEW_SHEET_, "request_id", target);
  if (!rowNos.length) return;
  var sh = getSheet_(REQUEST_ADMIN_REVIEW_SHEET_);
  var meta = sheetMeta_(sh);
  for (var i = 0; i < rowNos.length; i++) {
    updateRowFields_(sh, rowNos[i], meta, {
      is_current: "N",
      updated_at: nowIso_()
    });
  }
}

function getLatestActiveAssignmentRow_(requestId) {
  var rows = readRequestRowsIndexed_("RequestAssignments", String(requestId || "").trim()).filter(function(row) {
    return ynToBool_(row.is_active, true);
  }).sort(function(a, b) {
    return String(b.assigned_at || "").localeCompare(String(a.assigned_at || ""));
  });
  return rows.length ? rows[0] : null;
}

function buildActiveAssignmentViewModel_(requestRow) {
  var request = requestRow || {};
  var active = getLatestActiveAssignmentRow_(request.request_id) || {};
  return {
    assignment_id: String(active.assignment_id || "").trim(),
    assignee_name: String(active.assigned_to_name || request.assignee_name || "").trim(),
    assignee_email: String(active.assigned_to_email || request.assignee_email || "").trim(),
    assignee_slack_id: String(active.assigned_to_slack_id || request.assignee_slack_id || "").trim(),
    assigned_at: String(active.assigned_at || request.updated_at || request.created_at || "").trim(),
    note: String(active.note || "").trim()
  };
}

function upsertRequestAssignment_(requestId, opsPayload, assignedAtIso) {
  var rid = String(requestId || "").trim();
  if (!rid) return null;
  var ops = opsPayload || {};
  var latest = getLatestActiveAssignmentRow_(rid);
  var sameAssignee = latest
    && String(latest.assigned_to_name || "").trim() === String(ops.assignee_name || "").trim()
    && String(latest.assigned_to_email || "").trim() === String(ops.assignee_email || "").trim()
    && String(latest.assigned_to_slack_id || "").trim() === String(ops.assignee_slack_id || "").trim();
  var rowNos = getRequestRowNosFromIndex_(readRequestRowIndexRecord_(rid), "RequestAssignments");
  if (!rowNos.length) rowNos = findRowNosByCol_("RequestAssignments", "request_id", rid);
  if (rowNos.length) {
    var sh = getSheet_("RequestAssignments");
    var meta = sheetMeta_(sh);
    var rows = readRowsByNumbers_(sh, meta, rowNos);
    for (var i = 0; i < rowNos.length; i++) {
      if (!ynToBool_(rows[i].is_active, false)) continue;
      updateRowFields_(sh, rowNos[i], meta, { is_active: "N" });
    }
  }
  if (sameAssignee || (!ops.assignee_name && !ops.assignee_email && !ops.assignee_slack_id)) return latest;
  var assignment = {
    assignment_id: uuid_(),
    request_id: rid,
    assigned_to_name: ops.assignee_name,
    assigned_to_email: ops.assignee_email,
    assigned_to_slack_id: ops.assignee_slack_id,
    assigned_by: "ADMIN",
    assigned_at: assignedAtIso || nowIso_(),
    is_active: "Y",
    note: buildNextActionSummary_(ops),
    payload_json: JSON.stringify(ops)
  };
  appendRow_("RequestAssignments", assignment);
  return assignment;
}

function listRequestTasksForAdmin_(requestId) {
  return readRequestRowsIndexed_("RequestTasks", String(requestId || "").trim()).map(function(row) {
    var dueAt = String(row.due_at || "").trim();
    var remindAt = String(row.remind_at || "").trim();
    var isDone = normalizeTaskStatus_(row.task_status, "OPEN") === "DONE";
    return {
      task_id: row.task_id,
      task_type: String(row.task_type || "").trim(),
      task_label: String(row.task_label || "").trim(),
      task_status: normalizeTaskStatus_(row.task_status, "OPEN"),
      task_status_label: taskStatusLabel_(row.task_status),
      priority: normalizePriority_(row.priority, "NORMAL"),
      priority_label: priorityLabel_(row.priority),
      assignee_name: String(row.assignee_name || "").trim(),
      assignee_email: String(row.assignee_email || "").trim(),
      assignee_slack_id: String(row.assignee_slack_id || "").trim(),
      due_at: dueAt,
      due_at_label: formatAdminDateTime_(dueAt),
      remind_at: remindAt,
      remind_at_label: formatAdminDateTime_(remindAt),
      note: String(row.note || "").trim(),
      created_at: String(row.created_at || "").trim(),
      updated_at: String(row.updated_at || "").trim(),
      completed_at: String(row.completed_at || "").trim(),
      is_overdue: !isDone && !!remindAt && new Date(remindAt).getTime() < Date.now()
    };
  }).sort(function(a, b) {
    return String(b.updated_at || b.created_at || "").localeCompare(String(a.updated_at || a.created_at || ""));
  });
}

function upsertRequestNextActionTask_(requestId, opsPayload, nowIsoValue) {
  var rid = String(requestId || "").trim();
  if (!rid) return null;
  var ops = opsPayload || {};
  if (!ops.next_action_type && !ops.next_action_note && !ops.next_action_due_at && !ops.reminder_at) return null;
  var latest = listRequestTasksForAdmin_(rid).filter(function(task) {
    return task.task_type === "NEXT_ACTION";
  })[0];
  if (latest
    && latest.task_label === String(ops.next_action_type || "").trim()
    && latest.note === String(ops.next_action_note || "").trim()
    && latest.due_at === String(ops.next_action_due_at || "").trim()
    && latest.remind_at === String(ops.reminder_at || "").trim()
    && latest.task_status === String(ops.task_status || "OPEN").trim().toUpperCase()) {
    return latest;
  }
  var now = nowIsoValue || nowIso_();
  var task = {
    task_id: uuid_(),
    request_id: rid,
    task_type: "NEXT_ACTION",
    task_label: String(ops.next_action_type || "다음 액션").trim(),
    task_status: normalizeTaskStatus_(ops.task_status, "OPEN"),
    priority: normalizePriority_(ops.priority, "NORMAL"),
    assignee_name: ops.assignee_name,
    assignee_email: ops.assignee_email,
    assignee_slack_id: ops.assignee_slack_id,
    due_at: ops.next_action_due_at,
    remind_at: ops.reminder_at,
    created_at: now,
    updated_at: now,
    completed_at: normalizeTaskStatus_(ops.task_status, "OPEN") === "DONE" ? now : "",
    note: String(ops.next_action_note || "").trim(),
    payload_json: JSON.stringify(ops)
  };
  appendRow_("RequestTasks", task);
  return task;
}

function getDefaultChecklistTemplates_(projectType) {
  var project = String(projectType || "").trim().toUpperCase();
  if (project === "COMMERCIAL") {
    return [
      { template_id: "COMMERCIAL_BIZ", checklist_code: "biz_type", checklist_label: "업종 확인", is_required: "Y", sort_order: 10 },
      { template_id: "COMMERCIAL_OPEN", checklist_code: "open_schedule", checklist_label: "오픈 일정", is_required: "Y", sort_order: 20 },
      { template_id: "COMMERCIAL_EQUIP", checklist_code: "equipment_exhaust", checklist_label: "설비/배기", is_required: "Y", sort_order: 30 },
      { template_id: "COMMERCIAL_SIGN", checklist_code: "signage", checklist_label: "간판", is_required: "N", sort_order: 40 },
      { template_id: "COMMERCIAL_FIRE", checklist_code: "fire_legal", checklist_label: "소방/법규", is_required: "Y", sort_order: 50 },
      { template_id: "COMMERCIAL_MEASURE", checklist_code: "site_measure", checklist_label: "현장 실측 필요 여부", is_required: "Y", sort_order: 60 }
    ];
  }
  return [
    { template_id: "RESIDENTIAL_LIVING", checklist_code: "living_in_status", checklist_label: "거주중 여부", is_required: "Y", sort_order: 10 },
    { template_id: "RESIDENTIAL_DEMOLITION", checklist_code: "demolition_scope", checklist_label: "철거 범위", is_required: "Y", sort_order: 20 },
    { template_id: "RESIDENTIAL_WET", checklist_code: "wet_area_scope", checklist_label: "욕실/주방 포함 여부", is_required: "Y", sort_order: 30 },
    { template_id: "RESIDENTIAL_MOVE", checklist_code: "move_storage", checklist_label: "이사/짐보관", is_required: "N", sort_order: 40 },
    { template_id: "RESIDENTIAL_MOVEIN", checklist_code: "target_move_in", checklist_label: "입주 희망일", is_required: "Y", sort_order: 50 }
  ];
}

function getChecklistTemplatesForProject_(projectType) {
  var project = String(projectType || "").trim().toUpperCase();
  var rows = readAllRows_("RequestChecklistTemplates").filter(function(row) {
    var target = String(row.project_type || "").trim().toUpperCase();
    return !target || target === "ALL" || target === project;
  });
  if (!rows.length) return getDefaultChecklistTemplates_(project);
  return rows;
}

function normalizeChecklistItemRow_(row) {
  return {
    item_id: String(row.item_id || "").trim(),
    request_id: String(row.request_id || "").trim(),
    template_id: String(row.template_id || "").trim(),
    checklist_code: String(row.checklist_code || "").trim(),
    checklist_label: String(row.checklist_label || "").trim(),
    is_required: ynToBool_(row.is_required, true),
    is_completed: ynToBool_(row.is_completed, false),
    completed_at: String(row.completed_at || "").trim(),
    completed_by: String(row.completed_by || "").trim(),
    note: String(row.note || "").trim(),
    sort_order: toNumber_(row.sort_order, 0),
    payload: safeJsonParse_(row.payload_json, {})
  };
}

function ensureChecklistItemsForRequest_(requestRow) {
  var request = requestRow || {};
  var rid = String(request.request_id || "").trim();
  if (!rid) return [];
  var existing = readRequestRowsIndexed_("RequestChecklistItems", rid);
  if (existing.length) return existing.map(normalizeChecklistItemRow_);
  var templates = getChecklistTemplatesForProject_(request.project_type);
  var now = nowIso_();
  var rows = templates.map(function(tpl, index) {
    return {
      item_id: uuid_(),
      request_id: rid,
      template_id: String(tpl.template_id || "").trim() || ("TPL_" + String(index + 1)),
      checklist_code: String(tpl.checklist_code || "").trim(),
      checklist_label: String(tpl.checklist_label || "").trim(),
      is_required: ynToBool_(tpl.is_required, true) ? "Y" : "N",
      is_completed: "N",
      completed_at: "",
      completed_by: "",
      note: "",
      sort_order: toNumber_(tpl.sort_order, (index + 1) * 10),
      payload_json: JSON.stringify({ created_at: now, source: "DEFAULT_TEMPLATE" })
    };
  });
  appendRows_("RequestChecklistItems", rows);
  return rows.map(normalizeChecklistItemRow_);
}

function buildChecklistItemsForAdmin_(requestRow) {
  var list = ensureChecklistItemsForRequest_(requestRow).map(normalizeChecklistItemRow_);
  list.sort(function(a, b) {
    return Number(a.sort_order || 0) - Number(b.sort_order || 0);
  });
  return list;
}

function saveChecklistItems_(requestId, items, updatedAtIso) {
  var rid = String(requestId || "").trim();
  if (!rid || !Array.isArray(items)) return buildChecklistItemsForAdmin_({ request_id: rid });
  ensurePrequoteOperationalSchema_();
  var now = updatedAtIso || nowIso_();
  for (var i = 0; i < items.length; i++) {
    var item = items[i] || {};
    var itemId = String(item.item_id || "").trim();
    var found = itemId ? findRowByCol_("RequestChecklistItems", "item_id", itemId) : null;
    var completed = !!item.is_completed;
    var payload = {
      request_id: rid,
      template_id: String(item.template_id || "").trim(),
      checklist_code: String(item.checklist_code || "").trim(),
      checklist_label: String(item.checklist_label || "").trim(),
      is_required: item.is_required ? "Y" : "N",
      is_completed: completed ? "Y" : "N",
      completed_at: completed ? (String(item.completed_at || "").trim() || now) : "",
      completed_by: completed ? String(item.completed_by || "ADMIN").trim() : "",
      note: String(item.note || "").trim(),
      sort_order: toNumber_(item.sort_order, (i + 1) * 10),
      payload_json: JSON.stringify(item.payload || {})
    };
    if (found) {
      updateRowFields_(found.sheet, found.rowNo, found.meta, payload);
    } else {
      appendRow_("RequestChecklistItems", Object.assign({ item_id: itemId || uuid_() }, payload));
    }
  }
  return buildChecklistItemsForAdmin_({ request_id: rid, project_type: "" });
}

function collectDuplicateRequestHistory_(requestRow, limit) {
  var row = requestRow || {};
  var phoneKey = String(row.duplicate_phone_key || "").trim() || normalizePhoneKey_(row.contact_phone);
  var emailKey = String(row.duplicate_email_key || "").trim() || normalizeEmailKey_(row.contact_email);
  var optionIndex = getSurveyOptionLabelIndex_();
  return findDuplicateRequests_(phoneKey, emailKey, row.customer_name, row.request_id).sort(function(a, b) {
    return String(b.created_at || "").localeCompare(String(a.created_at || ""));
  }).slice(0, Math.max(Number(limit || 8), 1)).map(function(item) {
    var projectType = String(item.project_type || "").trim().toUpperCase();
    var areaQuestion = questionCodeForProject_(projectType, "R002_AREA", "C002_AREA");
    return {
      request_id: item.request_id,
      customer_name: item.customer_name,
      contact_phone: item.contact_phone,
      created_at: item.created_at,
      created_at_label: formatAdminDateTime_(item.created_at),
      status: item.status,
      status_label: statusLabel_(item.status),
      project_type_label: lookupSurveyOptionLabel_("Q000_PROJECT_TYPE", item.project_type, optionIndex) || item.project_type,
      area_label: lookupSurveyOptionLabel_(areaQuestion, item.area_py, optionIndex) || String(item.area_py || "").trim(),
      estimate_min: toNumber_(item.final_estimate_min || item.estimate_min, 0),
      estimate_max: toNumber_(item.final_estimate_max || item.estimate_max, 0),
      address_text: String(item.address_text || "").trim()
    };
  });
}

function parseIsoDateMs_(value) {
  var raw = String(value || "").trim();
  if (!raw) return NaN;
  var dt = new Date(raw);
  return dt.getTime();
}

function buildDashboardQueueSummary_(requests) {
  var rows = Array.isArray(requests) ? requests : [];
  var nowMs = Date.now();
  var todayEnd = new Date();
  todayEnd.setHours(23, 59, 59, 999);
  var todayEndMs = todayEnd.getTime();
  var out = {
    today_due: 0,
    overdue_reminder: 0,
    in_review: 0,
    new_requests: 0,
    quote_waiting: 0
  };
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var status = String(row.status || "NEW").trim().toUpperCase();
    var workState = normalizeWorkState_(row.work_state, "NEW");
    var reviewLabel = String(row.review_status_label || "").trim();
    var dueMs = parseIsoDateMs_(row.next_action_due_at);
    var remindMs = parseIsoDateMs_(row.reminder_at);
    if (status === "NEW") out.new_requests++;
    if (workState === "REVIEWING" || reviewLabel === REVIEW_STATUS_LABEL_MAP_.IN_REVIEW) out.in_review++;
    if (workState === "READY_FOR_QUOTE" || (status === "CONVERTED" && String(row.quote_draft_status || "").trim().toUpperCase() !== "CREATED")) {
      out.quote_waiting++;
    }
    if (!isNaN(dueMs) && dueMs <= todayEndMs && workState !== "DONE" && status !== "CLOSED") out.today_due++;
    if (!isNaN(remindMs) && remindMs < nowMs && workState !== "DONE" && status !== "CLOSED") out.overdue_reminder++;
  }
  return out;
}

function buildConsultationSummaryPayload_(context) {
  var ctx = context || {};
  var req = ctx.request || {};
  var flags = Array.isArray(ctx.flags) ? ctx.flags : [];
  var recommendations = Array.isArray(ctx.recommendations) ? ctx.recommendations : [];
  var checklist = Array.isArray(ctx.checklist) ? ctx.checklist : [];
  var incompleteChecklist = checklist.filter(function(item) {
    return item.is_required && !item.is_completed;
  }).map(function(item) {
    return item.checklist_label;
  });
  var topRecommendationTitles = recommendations.filter(function(item) {
    return item.is_visible !== false && !item.is_deleted;
  }).slice(0, 3).map(function(item) {
    return item.title;
  });
  var riskLabels = flags.slice(0, 3).map(function(flag) {
    return flag.label;
  });
  var internalLines = [
    "고객: " + [req.customer_name || "-", req.contact_phone || "-"].join(" / "),
    "프로젝트: " + [req.project_type_label || "-", req.housing_type_label || "-", req.area_label || "-", req.scope_level_label || "-"].join(" / "),
    "예산/가견적: " + formatMoneyRange(req.final_estimate_min || req.auto_estimate_min || 0, req.final_estimate_max || req.auto_estimate_max || 0),
    "일정: " + [req.schedule_start_label || "미정", req.schedule_end_label || "미정"].join(" ~ "),
    "핵심 요구: " + (req.customer_note_label || "추가 메모 없음"),
    "리스크: " + (riskLabels.length ? riskLabels.join(", ") : "특이 리스크 없음"),
    "추천 포인트: " + (topRecommendationTitles.length ? topRecommendationTitles.join(", ") : "추천 자재 미선정"),
    "다음 액션: " + (req.next_action_summary || "다음 액션 미지정")
  ];
  if (incompleteChecklist.length) internalLines.push("체크리스트 미완료: " + incompleteChecklist.join(", "));
  var shortParts = [
    req.customer_name || "고객",
    req.project_type_label || req.project_type || "",
    req.area_label || "",
    formatMoneyRange(req.final_estimate_min || req.auto_estimate_min || 0, req.final_estimate_max || req.auto_estimate_max || 0),
    req.next_action_summary || ""
  ].filter(Boolean);
  return {
    internal_text: internalLines.join("\n"),
    short_text: shortParts.join(" / "),
    incomplete_required_items: incompleteChecklist
  };
}

function getLatestQuoteDraftLinkForRequest_(requestId) {
  var rows = readRequestRowsIndexed_("RequestQuoteDraftLinks", String(requestId || "").trim()).sort(function(a, b) {
    return String(b.updated_at || b.created_at || "").localeCompare(String(a.updated_at || a.created_at || ""));
  });
  if (!rows.length) return null;
  var row = rows[0];
  return {
    link_id: row.link_id,
    request_id: row.request_id,
    review_id: row.review_id,
    quote_id: row.quote_id,
    status: row.status,
    created_at: row.created_at,
    updated_at: row.updated_at,
    open_url: row.open_url,
    open_hint: buildQuoteDraftOpenHint_(row.quote_id, { request_id: row.request_id, latest_review_id: row.review_id }),
    note: row.note,
    payload: safeJsonParse_(row.payload_json, {})
  };
}

function getSlackRuntimeConfig_() {
  var props = PropertiesService.getScriptProperties();
  var settings = getSettings_();
  return {
    webhook_url: String(props.getProperty("SLACK_WEBHOOK_URL") || "").trim(),
    channel: String(props.getProperty("SLACK_CHANNEL") || settings.slack_channel || "").trim(),
    user_map: safeJsonParse_(props.getProperty("SLACK_USER_MAP_JSON") || settings.slack_user_map_json || "{}", {})
  };
}

function getSlackUserMap_() {
  var cfg = getSlackRuntimeConfig_();
  return cfg.user_map && typeof cfg.user_map === "object" ? cfg.user_map : {};
}

function hasLoggedNotificationDedup_(dedupKey) {
  if (!dedupKey) return false;
  return findRowsByCol_("RequestNotificationLog", "dedup_key", dedupKey).some(function(row) {
    var status = String(row.status || "").trim().toUpperCase();
    return status === "SENT" || status === "SKIPPED" || status === "FAILED";
  });
}

function appendNotificationLog_(entry) {
  appendRow_("RequestNotificationLog", {
    notification_id: uuid_(),
    request_id: entry.request_id || "",
    channel: entry.channel || "SLACK",
    event_type: entry.event_type || "",
    dedup_key: entry.dedup_key || "",
    recipient_key: entry.recipient_key || "",
    recipient_value: entry.recipient_value || "",
    status: entry.status || "",
    created_at: entry.created_at || nowIso_(),
    sent_at: entry.sent_at || "",
    last_error: entry.last_error || "",
    payload_json: JSON.stringify(entry.payload || {})
  });
}

function sendSlackWebhook_(requestId, eventType, dedupKey, payloadObj) {
  ensureRequestNotificationLogSheet_();
  if (hasLoggedNotificationDedup_(dedupKey)) {
    return { ok: true, skipped: true, reason: "dedup" };
  }
  var cfg = getSlackRuntimeConfig_();
  var payload = payloadObj || {};
  var createdAt = nowIso_();
  if (!cfg.webhook_url) {
    appendNotificationLog_({
      request_id: requestId,
      channel: "SLACK",
      event_type: eventType,
      dedup_key: dedupKey,
      recipient_key: "channel",
      recipient_value: cfg.channel || "",
      status: "SKIPPED",
      created_at: createdAt,
      last_error: "SLACK_WEBHOOK_URL missing",
      payload: payload
    });
    return { ok: false, skipped: true, reason: "missing_webhook" };
  }
  var body = Object.assign({}, payload);
  if (cfg.channel) body.channel = cfg.channel;
  var response = null;
  var code = 0;
  var errorText = "";
  try {
    response = UrlFetchApp.fetch(cfg.webhook_url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    code = Number(response.getResponseCode() || 0);
    if (code < 200 || code >= 300) {
      errorText = response.getContentText() || ("HTTP " + code);
    }
  } catch (e) {
    errorText = String(e && e.message ? e.message : e);
  }
  appendNotificationLog_({
    request_id: requestId,
    channel: "SLACK",
    event_type: eventType,
    dedup_key: dedupKey,
    recipient_key: "channel",
    recipient_value: cfg.channel || "",
    status: errorText ? "FAILED" : "SENT",
    created_at: createdAt,
    sent_at: errorText ? "" : nowIso_(),
    last_error: errorText,
    payload: body
  });
  return { ok: !errorText, skipped: false, error: errorText, http_code: code };
}

function buildSlackRequestText_(requestRow, eventType, extras) {
  var row = requestRow || {};
  var extra = extras || {};
  var titleMap = {
    NEW_REQUEST: "새 가견적 요청",
    REMINDER_OVERDUE: "리마인드 기한 경과",
    QUOTE_READY: "실견적 전환 대상",
    ASSIGNEE_CHANGED: "담당자 변경",
    RESULT_VIEWED: "고객이 결과 확인"
  };
  var lines = [
    "[" + (titleMap[eventType] || eventType) + "]",
    row.customer_name + " / " + (row.contact_phone || "-"),
    (extra.project_label || row.project_type || "-") + " / " + (extra.area_label || row.area_py || "-"),
    "상태: " + statusLabel_(row.status) + " / 검토: " + String(row.review_status_label || row.final_estimate_source || "AUTO"),
    "가견적: " + formatMoneyRange(row.final_estimate_min || row.estimate_min || 0, row.final_estimate_max || row.estimate_max || 0)
  ];
  if (extra.summary) lines.push(extra.summary);
  if (extra.assignee_mention) lines.push("담당: " + extra.assignee_mention);
  return lines.join("\n");
}

function notifyNewRequestSlack_(requestRow, estimate) {
  var row = requestRow || {};
  var dedupKey = "request:" + String(row.request_id || "") + ":new";
  var text = buildSlackRequestText_(Object.assign({}, row, {
    final_estimate_min: estimate && estimate.min,
    final_estimate_max: estimate && estimate.max
  }), "NEW_REQUEST", {
    summary: "자동 계산 참고 범위가 생성되었습니다."
  });
  sendSlackWebhook_(row.request_id, "NEW_REQUEST", dedupKey, { text: text });
}

function maybeNotifyAssigneeChange_(previousRow, currentRow, reviewRow) {
  var prev = previousRow || {};
  var next = currentRow || {};
  if (String(prev.assignee_name || "").trim() === String(next.assignee_name || "").trim()
    && String(prev.assignee_email || "").trim() === String(next.assignee_email || "").trim()
    && String(prev.assignee_slack_id || "").trim() === String(next.assignee_slack_id || "").trim()) {
    return;
  }
  if (!next.assignee_name && !next.assignee_slack_id) return;
  var dedupKey = "request:" + String(next.request_id || "") + ":assignee:" + sha256Hex_([
    next.assignee_name, next.assignee_email, next.assignee_slack_id, reviewRow.review_id
  ].join("|")).slice(0, 16);
  var mention = next.assignee_slack_id ? ("<@" + next.assignee_slack_id + ">") : (next.assignee_name || "");
  sendSlackWebhook_(next.request_id, "ASSIGNEE_CHANGED", dedupKey, {
    text: buildSlackRequestText_(next, "ASSIGNEE_CHANGED", {
      summary: "담당자가 " + (prev.assignee_name || "미지정") + " → " + (next.assignee_name || "미지정") + " 로 변경되었습니다.",
      assignee_mention: mention
    })
  });
}

function processOperationalNotifications_(requests) {
  var rows = Array.isArray(requests) ? requests : readAllRows_("Requests");
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var requestId = String(row.request_id || "").trim();
    if (!requestId) continue;
    var workState = normalizeWorkState_(row.work_state, "NEW");
    var remindAt = String(row.reminder_at || "").trim();
    if (remindAt && workState !== "DONE" && String(row.status || "").trim().toUpperCase() !== "CLOSED") {
      var remindMs = parseIsoDateMs_(remindAt);
      if (!isNaN(remindMs) && remindMs < Date.now()) {
        sendSlackWebhook_(requestId, "REMINDER_OVERDUE", "request:" + requestId + ":reminder:" + remindAt, {
          text: buildSlackRequestText_(row, "REMINDER_OVERDUE", {
            summary: "리마인드 예정 시각 " + formatAdminDateTime_(remindAt) + " 이 지났습니다."
          })
        });
      }
    }
    if (workState === "READY_FOR_QUOTE" && String(row.quote_draft_status || "").trim().toUpperCase() !== "CREATED") {
      sendSlackWebhook_(requestId, "QUOTE_READY", "request:" + requestId + ":quote-ready:" + String(row.latest_review_id || row.latest_review_at || ""), {
        text: buildSlackRequestText_(row, "QUOTE_READY", {
          summary: "견적앱 초안 생성 대기 상태입니다."
        })
      });
    }
  }
}

function runOperationalNotifications_() {
  ensureSpreadsheetId_();
  ensurePrequoteOperationalSchema_();
  var mode = String(getSettings_().perf_dashboard_notifications_mode || "TRIGGER").trim().toUpperCase();
  if (mode === "DISABLED" || mode === "OFF" || mode === "NONE") {
    return { success: true, skipped: true, mode: mode };
  }
  var rows = readAllRows_("Requests");
  processOperationalNotifications_(rows);
  return { success: true, processed_count: rows.length, mode: mode };
}

function adminSearchMaterials(credential, query, options) {
  ensureSpreadsheetId_();
  var credInfo = assertAdminCredential_(credential);
  var opts = options || {};
  var limit = Math.min(Math.max(Number(opts.limit || 15), 1), 50);
  var ctx = {
    baseDbId: getQuoteBaseDbId_(),
    baseUrl: getQuoteMasterBaseUrl_()
  };
  var list = [];
  var startMs = Date.now();
  var error = null;
  try {
    try {
      list = BaseLib.searchMaterials(ctx, query, { limit: limit });
    } catch (e) {
      var mats = loadLocalMaterialCache_();
      if (!mats.length) mats = readQuoteMaterialsForRecommendation_();
      var q = String(query || "").trim().toLowerCase();
      list = mats.filter(function(item) {
        var hay = [
          item.material_id, item.name, item.brand, item.spec, item.trade_code, item.material_type, item.tags_summary
        ].join(" ").toLowerCase();
        return hay.indexOf(q) >= 0;
      }).slice(0, limit);
    }
    return list.map(function(item) {
      return {
        material_id: String(item.material_id || "").trim(),
        name: String(item.name || item.title || "").trim(),
        title: String(item.name || item.title || "").trim(),
        brand: String(item.brand || "").trim(),
        spec: String(item.spec || "").trim(),
        trade_code: String(item.trade_code || "").trim(),
        material_type: String(item.material_type || "").trim(),
        image_url: String(item.image_url || "").trim(),
        image_file_id: String(item.image_file_id || "").trim(),
        unit_price: toNumber_(item.unit_price, 0),
        subtitle: [item.brand, item.spec].filter(Boolean).join(" / "),
        source_type: String(item.source_type || "QUOTE_DB").trim().toUpperCase()
      };
    });
  } catch (e2) {
    error = e2;
    throw e2;
  } finally {
    logPerfMetricSafe_(
      "material_search",
      Date.now() - startMs,
      "",
      { query: String(query || "").trim(), limit: limit, result_count: list.length },
      !error,
      error ? String(error) : "",
      "ADMIN",
      credInfo.type
    );
  }
}

function adminBulkUpdateRequests(credential, payload) {
  ensureSpreadsheetId_();
  ensurePrequoteOperationalSchema_();
  assertAdminCredential_(credential);
  var input = payload && typeof payload === "object" ? payload : {};
  var ids = Array.isArray(input.request_ids) ? input.request_ids : [];
  if (!ids.length) throw new Error("대상 접수를 선택해주세요.");
  var updated = [];
  var now = nowIso_();
  for (var i = 0; i < ids.length; i++) {
    var found = findRowByCol_("Requests", "request_id", String(ids[i] || "").trim());
    if (!found) continue;
    var patch = {};
    if (input.status) patch.status = String(input.status || "").trim().toUpperCase();
    if (input.priority !== undefined) patch.priority = normalizePriority_(input.priority, found.data.priority || "NORMAL");
    if (input.work_state !== undefined) patch.work_state = normalizeWorkState_(input.work_state, found.data.work_state || "NEW");
    if (input.assignee_name !== undefined) patch.assignee_name = String(input.assignee_name || "").trim();
    if (input.assignee_email !== undefined) patch.assignee_email = String(input.assignee_email || "").trim();
    if (input.assignee_slack_id !== undefined) patch.assignee_slack_id = String(input.assignee_slack_id || "").trim();
    if (input.reminder_at !== undefined) patch.reminder_at = String(input.reminder_at || "").trim();
    patch.updated_at = now;
    updateRowFields_(found.sheet, found.rowNo, found.meta, patch);
    appendRow_("RequestTimeline", Object.assign(
      buildTimelineEvent_(found.data.request_id, "ADMIN", "BULK_UPDATE", "", patch.status || "", "일괄 처리 적용", JSON.stringify(patch), now),
      { event_label: eventTypeLabel_("BULK_UPDATE") }
    ));
    updated.push(Object.assign({}, found.data, patch));
  }
  return {
    success: true,
    updated_count: updated.length,
    updated_requests: updated.map(function(row) {
      return {
        request_id: row.request_id,
        status: row.status,
        priority: row.priority,
        assignee_name: row.assignee_name,
        reminder_at: row.reminder_at,
        work_state: row.work_state
      };
    })
  };
}

function buildQuoteDraftItemsFromRecommendations_(requestRow, detailPayload) {
  var request = requestRow || {};
  var detail = detailPayload || {};
  var list = (detail.recommendations || []).filter(function(item) {
    return String(item.rec_type || "").trim().toUpperCase() === "MATERIAL" && item.is_visible !== false && !item.is_deleted;
  });
  var out = [];
  for (var i = 0; i < list.length; i++) {
    var item = list[i];
    var detailText = [item.brand, item.spec, item.reason_text].filter(Boolean).join(" / ");
    out.push({
      seq: i + 1,
      group_id: slugifyQuoteGroupId_(item.trade_code || item.material_type || "recommendation"),
      group_label: item.trade_code || item.material_type || "추천 자재",
      group_code: item.trade_code || "",
      group_order: i + 1,
      item_order: i + 1,
      name: item.title || item.material_id || ("추천 자재 " + String(i + 1)),
      location: "",
      detail: detailText,
      price: toNumber_(item.price_hint_max || item.price_hint_min, 0),
      price_type: "NORMAL",
      process: item.trade_code || item.material_type || "추천 자재",
      material: item.title || item.material_id || "",
      brand: item.brand || "",
      spec: item.spec || "",
      qty: 1,
      unit: "식",
      unit_price: toNumber_(item.price_hint_max || item.price_hint_min, 0),
      amount: toNumber_(item.price_hint_max || item.price_hint_min, 0),
      note: item.reason_text || "",
      material_ref_id: item.material_id || "",
      material_image_id: item.image_file_id || "",
      material_image_name: "",
      source_request_id: request.request_id,
      source_review_id: request.latest_review_id || "",
      source_request_rec_id: item.rec_id || "",
      imported_from_prequote: "Y",
      source_material_id: item.material_id || "",
      source_material_image_url: item.image_url || ""
    });
  }
  if (out.length) return out;
  return [{
    seq: 1,
    group_id: "prequote-summary",
    group_label: "가견적 요약",
    group_code: "PREQUOTE",
    group_order: 1,
    item_order: 1,
    name: "현장 검토 후 상세 견적 구성",
    location: "",
    detail: String((detail.consultation_summary || {}).short_text || "").trim(),
    price: 0,
    price_type: "NORMAL",
    process: "가견적 요약",
    material: "현장 검토 후 상세 견적 구성",
    spec: "",
    qty: 1,
    unit: "식",
    unit_price: 0,
    amount: 0,
    note: "관리자 검토 내용을 기반으로 생성된 초안입니다.",
    material_ref_id: "",
    material_image_id: "",
    material_image_name: "",
    source_request_id: request.request_id,
    source_review_id: request.latest_review_id || "",
    source_request_rec_id: "",
    imported_from_prequote: "Y",
    source_material_id: "",
    source_material_image_url: ""
  }];
}

function slugifyQuoteGroupId_(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9가-힣]+/g, "-")
    .replace(/^-+|-+$/g, "") || "prequote";
}

function buildQuoteDraftOpenUrl_(quoteId, requestRow) {
  var params = [
    "page=edit",
    "quoteId=" + encodeURIComponent(String(quoteId || "").trim())
  ];
  var request = requestRow || {};
  if (request.request_id) params.push("source=prequote", "sourceRequestId=" + encodeURIComponent(String(request.request_id || "").trim()));
  if (request.latest_review_id) params.push("sourceReviewId=" + encodeURIComponent(String(request.latest_review_id || "").trim()));
  return getQuoteMasterBaseUrl_() + "?" + params.join("&");
}

function buildQuoteDraftOpenHint_(quoteId, requestRow) {
  var request = requestRow || {};
  return [
    "견적앱의 '견적 작성' 화면으로 이동합니다.",
    quoteId ? ("견적 번호: " + String(quoteId)) : "",
    request.request_id ? ("가견적 요청: " + String(request.request_id)) : "",
    "상단 배너에서 가견적에서 가져온 초안 여부를 바로 확인할 수 있습니다."
  ].filter(Boolean).join(" ");
}

function createQuoteDraftInQuoteDb_(requestRow, detailPayload) {
  ensureQuoteDraftImportSchema_();
  var quoteDbId = getQuoteBaseDbId_();
  var request = requestRow || {};
  var detail = detailPayload || {};
  var startMs = Date.now();
  var error = null;
  var result = null;
  var quoteId = "";
  try {
    var quoteSs = getSpreadsheet_(quoteDbId);
    var now = nowIso_();
    var queueId = "PQDQ_" + token_().slice(0, 16).toUpperCase();
    quoteId = uuid_();
    var items = buildQuoteDraftItemsFromRecommendations_(request, detail);
    var subtotal = 0;
    for (var i = 0; i < items.length; i++) subtotal += toNumber_(items[i].price, 0);
    var vatRate = Number(getSettingFromSs_(quoteSs, "vat_rate", 0.1) || 0.1);
    var vatIncluded = String(getSettingFromSs_(quoteSs, "default_vat_included", "N") || "N").trim().toUpperCase();
    var vat = vatIncluded === "Y" ? Math.round(subtotal * vatRate) : 0;
    var total = subtotal + vat;
    var summary = detail.consultation_summary || buildConsultationSummaryPayload_(detail);
    var materialSnapshot = items.filter(function(item) {
      return !!String(item.source_material_id || "").trim();
    }).map(function(item) {
      return {
        material_id: item.source_material_id,
        title: item.material,
        spec: item.spec,
        image_url: item.source_material_image_url,
        price: item.price
      };
    });
    appendRow_("Quotes", {
      quote_id: quoteId,
      created_at: now,
      updated_at: now,
      customer_name: request.customer_name || "",
      site_name: request.address_text || request.address_label || "",
      contact_name: request.customer_name || "",
      contact_phone: request.contact_phone || "",
      memo: summary.internal_text || "",
      subtotal: subtotal,
      vat: vat,
      total: total,
      share_token: "",
      status: "draft",
      shared_at: "",
      expire_at: "",
      view_count: 0,
      last_viewed_at: "",
      next_action_at: request.next_action_due_at || "",
      last_followup_sent_at: "",
      approved_at: "",
      deposit_amount: 0,
      deposit_due_at: "",
      last_customer_note_at: "",
      customer_note_latest: "",
      owner_last_seen_note_at: "",
      vat_included: vatIncluded,
      include_service_in_total: "N",
      include_option_in_total: "N",
      include_included_in_total: "N",
      source_request_id: request.request_id,
      source_review_id: request.latest_review_id || "",
      source_app: "PREQUOTE_APP",
      assignee_name: request.assignee_name || "",
      assignee_slack_id: request.assignee_slack_id || "",
      draft_stage: String(getSettings_().quote_draft_default_status || "DRAFT").trim(),
      prequote_summary_json: JSON.stringify({
        request_id: request.request_id,
        review_id: request.latest_review_id || "",
        summary_internal_text: summary.internal_text || "",
        summary_short_text: summary.short_text || "",
        final_estimate_min: request.final_estimate_min || request.estimate_min || 0,
        final_estimate_max: request.final_estimate_max || request.estimate_max || 0
      }),
      recommended_material_ids_json: JSON.stringify(materialSnapshot.map(function(item) { return item.material_id; })),
      recommended_materials_snapshot_json: JSON.stringify(materialSnapshot)
    }, quoteDbId);
    appendRows_("Items", items.map(function(item) {
      return Object.assign({ quote_id: quoteId }, item);
    }), quoteDbId);
    appendRow_("PrequoteDraftQueue", {
      queue_id: queueId,
      request_id: request.request_id,
      review_id: request.latest_review_id || "",
      created_at: now,
      status: "CREATED",
      customer_name: request.customer_name || "",
      contact_phone: request.contact_phone || "",
      project_type: request.project_type || "",
      assignee_name: request.assignee_name || "",
      payload_json: JSON.stringify({
        quote_id: quoteId,
        summary: summary,
        recommendation_count: (detail.recommendations || []).length
      }),
      linked_quote_id: quoteId,
      processed_at: now,
      last_error: ""
    }, quoteDbId);
    appendRows_("PrequoteDraftItems", items.map(function(item, index) {
      return {
        queue_id: queueId,
        request_id: request.request_id,
        review_id: request.latest_review_id || "",
        seq: index + 1,
        material_id: item.source_material_id || "",
        name: item.material || item.name || "",
        brand: item.brand || "",
        spec: item.spec || "",
        image_file_id: item.material_image_id || "",
        image_url: item.source_material_image_url || "",
        qty: item.qty || 1,
        unit: item.unit || "",
        unit_price: item.unit_price || 0,
        amount: item.amount || item.price || 0,
        note: item.note || "",
        payload_json: JSON.stringify(item)
      };
    }), quoteDbId);
    appendRow_("PrequoteImportLog", {
      log_id: uuid_(),
      created_at: now,
      request_id: request.request_id,
      review_id: request.latest_review_id || "",
      quote_id: quoteId,
      status: "CREATED",
      message: "가견적 검토 내용으로 견적 초안 생성",
      payload_json: JSON.stringify({
        queue_id: queueId,
        item_count: items.length,
        subtotal: subtotal,
        total: total
      })
    }, quoteDbId);
    result = {
      queue_id: queueId,
      quote_id: quoteId,
      open_url: buildQuoteDraftOpenUrl_(quoteId, request),
      open_hint: buildQuoteDraftOpenHint_(quoteId, request)
    };
    return result;
  } catch (e) {
    error = e;
    throw e;
  } finally {
    logQuotePerfMetricSafe_(
      quoteDbId,
      "quote_draft_import_from_prequote",
      quoteId || String(request.request_id || "").trim(),
      Date.now() - startMs,
      {
        request_id: String(request.request_id || "").trim(),
        review_id: String(request.latest_review_id || "").trim(),
        item_count: detail && detail.recommendations ? detail.recommendations.length : 0,
        quote_id: result && result.quote_id ? result.quote_id : quoteId
      },
      !error,
      error ? String(error) : "",
      "SYSTEM",
      "PREQUOTE_APP"
    );
  }
}

function updateAdminReviewQuoteLink_(reviewId, quoteId, quoteStatus) {
  var targetReviewId = String(reviewId || "").trim();
  if (!targetReviewId) return;
  var found = findRowByCol_(REQUEST_ADMIN_REVIEW_SHEET_, "review_id", targetReviewId);
  if (!found) return;
  updateRowFields_(found.sheet, found.rowNo, found.meta, {
    linked_quote_id: quoteId,
    linked_quote_status: quoteStatus,
    updated_at: nowIso_()
  });
}

function adminCreateQuoteDraft(credential, requestId, options) {
  ensureSpreadsheetId_();
  ensurePrequoteOperationalSchema_();
  var credInfo = assertAdminCredential_(credential);
  var opts = options && typeof options === "object" ? options : {};
  var returnDetail = opts.return_detail !== false;
  var startMs = Date.now();
  var error = null;
  var result = null;
  try {
    var rid = String(requestId || "").trim();
    var found = findRequestRowById_(rid);
    if (!found) throw new Error("Request not found");
    if (!String(found.data.latest_review_id || "").trim()) {
      throw new Error("관리자 검토를 먼저 저장한 뒤 견적앱으로 전환해주세요.");
    }
    var detail = buildAdminRequestDetailPayload_(found);
    var incompleteRequired = (detail.checklist || []).filter(function(item) {
      return item.is_required && !item.is_completed;
    });
    if (incompleteRequired.length && !opts.force) {
      result = {
        success: false,
        requires_confirmation: true,
        message: "필수 체크리스트가 완료되지 않았습니다.",
        incomplete_items: incompleteRequired.map(function(item) { return item.checklist_label; }),
        detail: returnDetail ? detail : null,
        dashboard_patch: buildDashboardRequestSummary_(found.data, null, getSurveyOptionLabelIndex_())
      };
      return result;
    }

    var latestLink = getLatestQuoteDraftLinkForRequest_(found.data.request_id);
    if (latestLink && latestLink.quote_id && !opts.force_new) {
      result = {
        success: true,
        reused: true,
        quote_id: latestLink.quote_id,
        open_url: latestLink.open_url,
        open_hint: latestLink.open_hint || buildQuoteDraftOpenHint_(latestLink.quote_id, found.data),
        detail: returnDetail ? detail : null,
        dashboard_patch: buildDashboardRequestSummary_(found.data, null, getSurveyOptionLabelIndex_())
      };
      return result;
    }

    var previousStatus = String(found.data.status || "NEW").trim().toUpperCase();
    var draftResult = createQuoteDraftInQuoteDb_(found.data, detail);
    var now = nowIso_();
    appendRow_("RequestQuoteDraftLinks", {
      link_id: uuid_(),
      request_id: found.data.request_id,
      review_id: found.data.latest_review_id || "",
      quote_id: draftResult.quote_id,
      status: "CREATED",
      created_at: now,
      updated_at: now,
      source_app: "prequote_app",
      target_app: "quote_app",
      open_url: draftResult.open_url,
      note: "견적앱 초안 생성",
      payload_json: JSON.stringify(draftResult)
    });
    updateAdminReviewQuoteLink_(found.data.latest_review_id, draftResult.quote_id, "CREATED");
    var requestPatch = {
      updated_at: now,
      linked_quote_id: draftResult.quote_id,
      quote_conversion_status: "DRAFT_CREATED",
      quote_draft_status: "CREATED",
      status: "CONVERTED"
    };
    updateRowFields_(found.sheet, found.rowNo, found.meta, requestPatch);
    for (var key in requestPatch) {
      if (!Object.prototype.hasOwnProperty.call(requestPatch, key)) continue;
      found.data[key] = requestPatch[key];
    }
    var timelineEvent = Object.assign(
      buildTimelineEvent_(found.data.request_id, "ADMIN", "QUOTE_DRAFT_CREATED", previousStatus, "CONVERTED", "견적앱 초안을 생성했습니다.", JSON.stringify(draftResult), now),
      {
        event_label: eventTypeLabel_("QUOTE_DRAFT_CREATED"),
        linked_ref_id: draftResult.quote_id,
        notify_channel: "QUOTE_APP"
      }
    );
    appendRow_("RequestTimeline", timelineEvent);
    result = {
      success: true,
      quote_id: draftResult.quote_id,
      open_url: draftResult.open_url,
      open_hint: draftResult.open_hint || buildQuoteDraftOpenHint_(draftResult.quote_id, found.data),
      timeline_event: timelineEvent,
      dashboard_patch: buildDashboardRequestSummary_(found.data, null, getSurveyOptionLabelIndex_()),
      detail: returnDetail ? buildAdminRequestDetailPayload_(found) : null
    };
    return result;
  } catch (e) {
    error = e;
    throw e;
  } finally {
    logPerfMetricSafe_(
      "create_quote_draft",
      Date.now() - startMs,
      String(requestId || "").trim(),
      {
        success: !error,
        reused: !!(result && result.reused),
        requires_confirmation: !!(result && result.requires_confirmation),
        quote_id: result && result.quote_id ? result.quote_id : ""
      },
      !error,
      error ? String(error) : "",
      "ADMIN",
      credInfo.type
    );
  }
}

// ============================================================
//  SYNC WITH QUOTE MASTER APP
// ============================================================

function writeSyncLog_(syncType, status, payload) {
  var info = payload || {};
  appendRow_("SyncLog", {
    sync_id: uuid_(),
    synced_at: info.synced_at || nowIso_(),
    sync_type: syncType,
    source_app_url: info.source_app_url || "",
    master_version: info.master_version || "",
    status: status,
    read_count: info.read_count || 0,
    write_count: info.write_count || 0,
    duration_ms: info.duration_ms || 0,
    note: info.note || "",
    error_message: info.error_message || ""
  });
}

function runWithSyncLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function getQuoteBaseContext_() {
  return {
    baseDbId: getQuoteBaseDbId_(),
    baseUrl: getQuoteMasterBaseUrl_()
  };
}

function getQuoteSharedMasterVersion_() {
  var ctx = getQuoteBaseContext_();
  if (typeof BaseLib !== "undefined" && BaseLib && typeof BaseLib.getSharedMaterialMasterVersion === "function") {
    return BaseLib.getSharedMaterialMasterVersion(ctx) || null;
  }
  return null;
}

function shouldSkipSyncByVersion_(syncKey, sharedMasterVersion, cacheSheetName) {
  if (!syncUseMasterVersion_()) return false;
  var version = String(sharedMasterVersion || "").trim();
  if (!version) return false;
  if (!hasSheetBodyData_(cacheSheetName)) return false;
  var state = readSyncStateRow_(syncKey);
  if (!state) return false;
  var lastStatus = String(state.last_status || "").trim().toUpperCase();
  if (lastStatus !== "SUCCESS" && lastStatus !== "SKIPPED") return false;
  return String(state.shared_master_version || "").trim() === version;
}

function readQuoteMaterialSnapshotForSync_(versionInfo) {
  var ctx = getQuoteBaseContext_();
  var info = versionInfo || getQuoteSharedMasterVersion_() || {};
  var snapshot = null;
  if (typeof BaseLib !== "undefined" && BaseLib) {
    if (typeof BaseLib.getRepresentativeMaterialsSnapshotForPrequote === "function") {
      snapshot = BaseLib.getRepresentativeMaterialsSnapshotForPrequote(ctx, {}, {
        include_inactive: false,
        include_unexposed: true,
        limit: 50000
      });
    } else if (typeof BaseLib.listRepresentativeMaterialsForPrequote === "function") {
      snapshot = {
        shared_master_version: String(info.shared_master_version || "").trim(),
        built_at: nowIso_(),
        items: BaseLib.listRepresentativeMaterialsForPrequote(ctx, {
          include_inactive: false,
          include_unexposed: true,
          limit: 50000
        }, {
          include_inactive: false,
          include_unexposed: true,
          limit: 50000
        }) || [],
        source: "list_api"
      };
    }
  }
  if (!snapshot || !Array.isArray(snapshot.items)) {
    var fallbackItems = readQuoteMaterialsForRecommendation_();
    snapshot = {
      shared_master_version: String(info.shared_master_version || nowIso_()).trim(),
      built_at: nowIso_(),
      items: fallbackItems,
      source: "raw_fallback"
    };
  }
  snapshot.items = Array.isArray(snapshot.items) ? snapshot.items : [];
  snapshot.row_count = Math.max(Number(snapshot.row_count || 0), snapshot.items.length);
  snapshot.shared_master_version = String(snapshot.shared_master_version || info.shared_master_version || nowIso_()).trim();
  return snapshot;
}

function readQuoteTemplateSnapshotForSync_(versionInfo) {
  var ctx = getQuoteBaseContext_();
  var info = versionInfo || getQuoteSharedMasterVersion_() || {};
  var snapshot = null;
  if (typeof BaseLib !== "undefined" && BaseLib) {
    if (typeof BaseLib.getTemplateCatalogSnapshotForPrequote === "function") {
      snapshot = BaseLib.getTemplateCatalogSnapshotForPrequote(ctx, {}, {
        include_inactive: true,
        include_unexposed: true,
        limit: 50000
      });
    } else if (typeof BaseLib.listTemplateCatalogForPrequote === "function") {
      snapshot = {
        shared_master_version: String(info.shared_master_version || "").trim(),
        built_at: nowIso_(),
        items: BaseLib.listTemplateCatalogForPrequote(ctx, {
          include_inactive: true,
          include_unexposed: true,
          limit: 50000
        }, {
          include_inactive: true,
          include_unexposed: true,
          limit: 50000
        }) || [],
        source: "list_api"
      };
    }
  }
  if (!snapshot || !Array.isArray(snapshot.items)) {
    var fallbackItems = readQuoteTemplatesForRecommendation_();
    snapshot = {
      shared_master_version: String(info.shared_master_version || nowIso_()).trim(),
      built_at: nowIso_(),
      items: fallbackItems,
      source: "raw_fallback"
    };
  }
  snapshot.items = Array.isArray(snapshot.items) ? snapshot.items : [];
  snapshot.row_count = Math.max(Number(snapshot.row_count || 0), snapshot.items.length);
  snapshot.shared_master_version = String(snapshot.shared_master_version || info.shared_master_version || nowIso_()).trim();
  return snapshot;
}

function buildMaterialsCacheRows_(materials, version, syncedAt) {
  var list = Array.isArray(materials) ? materials : [];
  var rows = [];
  for (var i = 0; i < list.length; i++) {
    var m = list[i] || {};
    rows.push({
      master_version: version,
      synced_at: syncedAt,
      material_id: m.material_id,
      name: m.name,
      brand: m.brand,
      spec: m.spec,
      unit: m.unit,
      unit_price: m.unit_price,
      image_url: m.image_url || buildQuoteImageUrl_(m.image_file_id),
      is_active: ynToBool_(m.is_active, true) ? "Y" : "N",
      is_representative: ynToBool_(m.is_representative, false) ? "Y" : "N",
      material_type: m.material_type,
      trade_code: m.trade_code,
      space_type: m.space_type || "",
      sort_order: m.sort_order,
      expose_to_prequote: ynToBool_(m.expose_to_prequote, true) ? "Y" : "N",
      recommendation_score_base: m.recommendation_score_base,
      price_band: m.price_band,
      tags_summary: m.tags_summary,
      tag_records_json: JSON.stringify(Array.isArray(m.tag_records) ? m.tag_records : []),
      tags_by_type_json: JSON.stringify(m.tags_by_type || {}),
      legacy_group_keys_json: JSON.stringify(m.legacy_group_keys || []),
      recommendation_note: m.recommendation_note || ""
    });
  }
  return rows;
}

function buildTemplatesCacheRows_(templates, version, syncedAt) {
  var list = Array.isArray(templates) ? templates : [];
  var rows = [];
  for (var i = 0; i < list.length; i++) {
    var t = list[i] || {};
    rows.push({
      master_version: version,
      synced_at: syncedAt,
      template_id: t.template_id,
      category: t.category,
      template_name: t.template_name,
      latest_version: t.latest_version,
      latest_item_count: t.latest_item_count,
      is_active: ynToBool_(t.is_active, true) ? "Y" : "N",
      template_type: t.template_type,
      housing_type: t.housing_type,
      area_band: t.area_band,
      budget_band: t.budget_band,
      style_tags_summary: t.style_tags_summary,
      tone_tags_summary: t.tone_tags_summary,
      trade_scope_summary: t.trade_scope_summary,
      expose_to_prequote: ynToBool_(t.expose_to_prequote, true) ? "Y" : "N",
      prequote_priority: t.prequote_priority,
      sort_order: t.sort_order,
      recommendation_note: t.recommendation_note,
      target_customer_summary: t.target_customer_summary,
      recommended_for_summary: t.recommended_for_summary,
      is_featured_prequote: ynToBool_(t.is_featured_prequote, false) ? "Y" : "N",
      summary_json: String(t.summary_json || (t.summary ? JSON.stringify(t.summary) : "")),
      metadata_snapshot_json: String(t.metadata_snapshot_json || (t.metadata_snapshot ? JSON.stringify(t.metadata_snapshot) : ""))
    });
  }
  return rows;
}

function syncMaterialsCacheInternal_(versionInfo) {
  ensureSpreadsheetId_();
  var settings = getSettings_();
  var startMs = Date.now();
  var syncedAt = nowIso_();
  var error = null;
  var result = null;
  var versionData = versionInfo || getQuoteSharedMasterVersion_() || {};
  var sharedVersion = String(versionData.shared_master_version || "").trim();
  try {
    if (shouldSkipSyncByVersion_(SYNC_KEY_MATERIALS_CACHE_, sharedVersion, "MaterialsCache")) {
      var skipDurationMs = Date.now() - startMs;
      result = {
        success: true,
        skipped: true,
        count: readAllRows_("MaterialsCache").length,
        duration_ms: skipDurationMs,
        master_version: sharedVersion
      };
      upsertSyncState_(SYNC_KEY_MATERIALS_CACHE_, {
        source_app: "quote_app",
        source_spreadsheet_id: getQuoteBaseDbId_(),
        shared_master_version: sharedVersion,
        last_synced_at: syncedAt,
        last_status: "SKIPPED",
        read_count: 0,
        write_count: 0,
        duration_ms: skipDurationMs,
        note: "Shared master version unchanged"
      });
      writeSyncLog_("MATERIALS", "SKIPPED", {
        synced_at: syncedAt,
        source_app_url: settings.quote_master_app_url || "",
        master_version: sharedVersion,
        duration_ms: skipDurationMs,
        note: "Shared master version unchanged"
      });
      return result;
    }

    var snapshot = readQuoteMaterialSnapshotForSync_(versionData);
    var version = String(snapshot.shared_master_version || sharedVersion || syncedAt).trim();
    var rows = buildMaterialsCacheRows_(snapshot.items, version, syncedAt);
    clearSheetBody_("MaterialsCache");
    appendRows_("MaterialsCache", rows);

    var durationMs = Date.now() - startMs;
    upsertSyncState_(SYNC_KEY_MATERIALS_CACHE_, {
      source_app: "quote_app",
      source_spreadsheet_id: getQuoteBaseDbId_(),
      shared_master_version: version,
      last_synced_at: syncedAt,
      last_status: "SUCCESS",
      read_count: Number(snapshot.row_count || rows.length),
      write_count: rows.length,
      duration_ms: durationMs,
      note: String(snapshot.source || "snapshot").trim()
    });
    writeSyncLog_("MATERIALS", "SUCCESS", {
      synced_at: syncedAt,
      source_app_url: settings.quote_master_app_url || "",
      master_version: version,
      read_count: Number(snapshot.row_count || rows.length),
      write_count: rows.length,
      duration_ms: durationMs,
      note: String(snapshot.source || "snapshot").trim()
    });
    result = {
      success: true,
      skipped: false,
      count: rows.length,
      duration_ms: durationMs,
      master_version: version,
      source: String(snapshot.source || "snapshot").trim()
    };
    return result;
  } catch (e) {
    error = e;
    var failedDurationMs = Date.now() - startMs;
    upsertSyncState_(SYNC_KEY_MATERIALS_CACHE_, {
      source_app: "quote_app",
      source_spreadsheet_id: getQuoteBaseDbId_(),
      shared_master_version: sharedVersion,
      last_synced_at: syncedAt,
      last_status: "FAILED",
      read_count: 0,
      write_count: 0,
      duration_ms: failedDurationMs,
      note: String(e)
    });
    writeSyncLog_("MATERIALS", "FAILED", {
      source_app_url: settings.quote_master_app_url || "",
      master_version: sharedVersion,
      duration_ms: failedDurationMs,
      error_message: String(e)
    });
    throw e;
  } finally {
    logPerfMetricSafe_(
      "sync_materials_cache",
      Date.now() - startMs,
      "",
      {
        shared_master_version: sharedVersion,
        skipped: !!(result && result.skipped),
        row_count: result && result.count ? result.count : 0,
        source: result && result.source ? result.source : ""
      },
      !error,
      error ? String(error) : "",
      "SYSTEM",
      "SYNC"
    );
  }
}

function syncTemplatesCacheInternal_(versionInfo) {
  ensureSpreadsheetId_();
  var settings = getSettings_();
  var startMs = Date.now();
  var syncedAt = nowIso_();
  var error = null;
  var result = null;
  var versionData = versionInfo || getQuoteSharedMasterVersion_() || {};
  var sharedVersion = String(versionData.shared_master_version || "").trim();
  try {
    if (shouldSkipSyncByVersion_(SYNC_KEY_TEMPLATES_CACHE_, sharedVersion, "TemplatesCache")) {
      var skipDurationMs = Date.now() - startMs;
      result = {
        success: true,
        skipped: true,
        count: readAllRows_("TemplatesCache").length,
        duration_ms: skipDurationMs,
        master_version: sharedVersion
      };
      upsertSyncState_(SYNC_KEY_TEMPLATES_CACHE_, {
        source_app: "quote_app",
        source_spreadsheet_id: getQuoteBaseDbId_(),
        shared_master_version: sharedVersion,
        last_synced_at: syncedAt,
        last_status: "SKIPPED",
        read_count: 0,
        write_count: 0,
        duration_ms: skipDurationMs,
        note: "Shared master version unchanged"
      });
      writeSyncLog_("TEMPLATES", "SKIPPED", {
        synced_at: syncedAt,
        source_app_url: settings.quote_master_app_url || "",
        master_version: sharedVersion,
        duration_ms: skipDurationMs,
        note: "Shared master version unchanged"
      });
      return result;
    }

    var snapshot = readQuoteTemplateSnapshotForSync_(versionData);
    var version = String(snapshot.shared_master_version || sharedVersion || syncedAt).trim();
    var rows = buildTemplatesCacheRows_(snapshot.items, version, syncedAt);
    clearSheetBody_("TemplatesCache");
    appendRows_("TemplatesCache", rows);

    var durationMs = Date.now() - startMs;
    upsertSyncState_(SYNC_KEY_TEMPLATES_CACHE_, {
      source_app: "quote_app",
      source_spreadsheet_id: getQuoteBaseDbId_(),
      shared_master_version: version,
      last_synced_at: syncedAt,
      last_status: "SUCCESS",
      read_count: Number(snapshot.row_count || rows.length),
      write_count: rows.length,
      duration_ms: durationMs,
      note: String(snapshot.source || "snapshot").trim()
    });
    writeSyncLog_("TEMPLATES", "SUCCESS", {
      synced_at: syncedAt,
      source_app_url: settings.quote_master_app_url || "",
      master_version: version,
      read_count: Number(snapshot.row_count || rows.length),
      write_count: rows.length,
      duration_ms: durationMs,
      note: String(snapshot.source || "snapshot").trim()
    });
    result = {
      success: true,
      skipped: false,
      count: rows.length,
      duration_ms: durationMs,
      master_version: version,
      source: String(snapshot.source || "snapshot").trim()
    };
    return result;
  } catch (e) {
    error = e;
    var failedDurationMs = Date.now() - startMs;
    upsertSyncState_(SYNC_KEY_TEMPLATES_CACHE_, {
      source_app: "quote_app",
      source_spreadsheet_id: getQuoteBaseDbId_(),
      shared_master_version: sharedVersion,
      last_synced_at: syncedAt,
      last_status: "FAILED",
      read_count: 0,
      write_count: 0,
      duration_ms: failedDurationMs,
      note: String(e)
    });
    writeSyncLog_("TEMPLATES", "FAILED", {
      source_app_url: settings.quote_master_app_url || "",
      master_version: sharedVersion,
      duration_ms: failedDurationMs,
      error_message: String(e)
    });
    throw e;
  } finally {
    logPerfMetricSafe_(
      "sync_templates_cache",
      Date.now() - startMs,
      "",
      {
        shared_master_version: sharedVersion,
        skipped: !!(result && result.skipped),
        row_count: result && result.count ? result.count : 0,
        source: result && result.source ? result.source : ""
      },
      !error,
      error ? String(error) : "",
      "SYSTEM",
      "SYNC"
    );
  }
}

function syncMaterialsCache(credential) {
  ensureSpreadsheetId_();
  if (credential) assertAdminCredential_(credential);
  else assertEditorAdminExecution_();
  return runWithSyncLock_(syncMaterialsCacheInternal_);
}

function syncTemplatesCache(credential) {
  ensureSpreadsheetId_();
  if (credential) assertAdminCredential_(credential);
  else assertEditorAdminExecution_();
  return runWithSyncLock_(syncTemplatesCacheInternal_);
}

function adminSyncCaches(credential) {
  ensureSpreadsheetId_();
  assertAdminCredential_(credential);
  return runWithSyncLock_(function() {
    var versionInfo = getQuoteSharedMasterVersion_();
    var materials = syncMaterialsCacheInternal_(versionInfo);
    var templates = syncTemplatesCacheInternal_(versionInfo);
    return {
      success: true,
      shared_master_version: String(versionInfo && versionInfo.shared_master_version || "").trim(),
      materials: materials,
      templates: templates
    };
  });
}

// ============================================================
//  OWNER EMAIL TARGETS (for notifications)
// ============================================================

function getOwnerEmailTargets_() {
  var s = getSettings_();
  var raw = String(s.owner_email || "").trim();
  if (!raw) return [];
  return raw.split(",").map(function(v) { return v.trim(); }).filter(Boolean);
}

// ============================================================
//  TIMED TRIGGERS
// ============================================================

function setupSyncTrigger() {
  ensureSpreadsheetId_();
  assertEditorAdminExecution_();
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === "runScheduledSync_") {
      ScriptApp.deleteTrigger(existing[i]);
    }
  }
  var s = getSettings_();
  var mins = Math.max(Number(s.sync_interval_minutes || 30), 10);
  ScriptApp.newTrigger("runScheduledSync_").timeBased().everyMinutes(mins).create();
  return { interval: mins };
}

function setupOperationalNotificationsTrigger() {
  ensureSpreadsheetId_();
  assertEditorAdminExecution_();
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === "runOperationalNotifications_") {
      ScriptApp.deleteTrigger(existing[i]);
    }
  }
  var settings = getSettings_();
  var mode = String(settings.perf_dashboard_notifications_mode || "TRIGGER").trim().toUpperCase();
  if (mode === "DISABLED" || mode === "OFF" || mode === "NONE") {
    return { enabled: false, mode: mode };
  }
  var mins = Math.max(Number(settings.sync_interval_minutes || 30), 10);
  ScriptApp.newTrigger("runOperationalNotifications_").timeBased().everyMinutes(mins).create();
  return { enabled: true, mode: mode, interval: mins };
}

function runScheduledSync_() {
  ensureSpreadsheetId_();
  var s = getSettings_();
  if (String(s.sync_enabled || "").toUpperCase() !== "Y") return;
  try {
    runWithSyncLock_(function() {
      var versionInfo = null;
      try { versionInfo = getQuoteSharedMasterVersion_(); } catch (e0) { Logger.log("Sync version fetch error: " + e0); }
      try { syncMaterialsCacheInternal_(versionInfo); } catch (e) { Logger.log("Sync materials error: " + e); }
      try { syncTemplatesCacheInternal_(versionInfo); } catch (e2) { Logger.log("Sync templates error: " + e2); }
    });
  } catch (e3) {
    Logger.log("Scheduled sync lock error: " + e3);
  }
}
