var CORE_SCHEMA_VERSION_ = "2026-03-09-v3";
var CORE_SCHEMA_FLAG_KEY_ = "CORE_SCHEMA_VERSION";

var __SCHEMA_READY_THIS_EXEC_ = false;
var __HEADERS_CACHE_ = Object.create(null);
var __COLMAP_CACHE_ = Object.create(null);
var __QUOTE_ROWNO_CACHE_ = null;
var __QUOTE_ITEMS_INDEX_ROWNO_CACHE_ = null;
var __QUOTE_ROW_CACHE_ = Object.create(null);

var __UPLOAD_ROOT_FOLDER_CACHE_ = null;
var __MATERIALS_FOLDER_CACHE_ = null;
var __QUOTE_FOLDER_CACHE_ = Object.create(null);

var __MATERIAL_INDEX_MEM_VER_ = "";
var __MATERIAL_INDEX_MEM_MAP_ = null;
var __SHEET_COLUMN_SET_MEM_ = Object.create(null);
var __TEMPLATE_CATALOG_SNAPSHOT_MEM_VER_ = "";
var __TEMPLATE_CATALOG_SNAPSHOT_MEM_ = null;
var __TEMPLATE_CATALOG_CORE_MEM_VER_ = "";
var __TEMPLATE_CATALOG_CORE_MEM_ = null;
var __QUOTE_LIST_CORE_MEM_VER_ = "";
var __QUOTE_LIST_CORE_MEM_ = null;

var EVENT_CHANGE_REQUEST_MEMO_ = "CHANGE_REQUEST_MEMO";
var EVENT_APPROVAL_TOGGLED_ = "APPROVAL_TOGGLED";
var EVENT_OWNER_APPROVAL_EMAIL_ = "OWNER_APPROVAL_EMAIL";
var EVENT_OWNER_NOTE_EMAIL_ = "OWNER_NOTE_EMAIL";
var SLACK_MESSAGE_TEMPLATE_VERSION_ = "KO_v3_2026-02-23";

// Assumption: identical memo notifications are deduped within 2 minutes.
var MEMO_DEDUP_WINDOW_MS_ = 120000;

var __NOTIFICATION_SCHEMA_READY_THIS_EXEC_ = false;
var __SLACK_QUEUE_SCHEMA_READY_THIS_EXEC_ = false;
var __SLACK_QUEUE_TRIGGER_CHECKED_THIS_EXEC_ = false;

var TEMPLATE_CATALOG_SHEET_ = "TemplateCatalog";
var TEMPLATE_VERSIONS_SHEET_ = "TemplateVersions";
var TEMPLATE_ITEMS_CACHE_SHEET_ = "TemplateItemsCache";
var TEMPLATE_LEGACY_SHEET_ = "Templates";
var TEMPLATE_LIST_VERSION_KEY_ = "TEMPLATE_LIST_VER";
var MATERIAL_GROUP_LIST_VERSION_KEY_ = "MATERIAL_GROUP_LIST_VER";
var MATERIAL_TAG_LIST_VERSION_KEY_ = "MATERIAL_TAG_LIST_VER";
var MATERIAL_GROUPS_SHEET_ = "MaterialGroups";
var MATERIAL_TAGS_SHEET_ = "MaterialTags";
// false: base URL opens admin pages directly and auth is handled by in-page login flow.
// true: admin pages require URL credential at doGet-time.
var STRICT_ADMIN_GET_AUTH_ = false;
var BUILD_TAG_ = "2026-03-03T09:25:00Z";

var DEFAULT_PUBLIC_EXEC_BASE_URL_ = "";
var ITEMS_CANONICAL_HEADERS_ = [
  "quote_id", "item_id", "seq",
  "group_id", "group_label", "group_code", "group_order", "item_order",
  "name", "location", "detail", "price", "price_type",
  "process", "material", "spec", "qty", "unit", "unit_price", "amount", "note",
  "material_ref_id", "material_image_id", "material_image_name"
];
var MATERIAL_MASTER_HEADERS_ = [
  "material_id", "name", "brand", "spec", "unit", "unit_price",
  "image_file_id", "image_file_name", "note", "created_at", "updated_at",
  "search_text", "is_active", "is_representative", "material_type", "trade_code",
  "space_type", "sort_order", "expose_to_prequote", "recommendation_score_base",
  "price_band", "tags_summary", "recommendation_note"
];
var TEMPLATE_CATALOG_HEADERS_ = [
  "template_id",
  "category",
  "template_name",
  "latest_version",
  "latest_item_count",
  "created_at",
  "updated_at",
  "is_active",
  "note",
  "template_type",
  "housing_type",
  "area_band",
  "style_tags_summary",
  "tone_tags_summary",
  "trade_scope_summary",
  "expose_to_prequote",
  "prequote_priority",
  "sort_order",
  "recommendation_note",
  "target_customer_summary",
  "budget_band",
  "recommended_for_summary",
  "is_featured_prequote"
];
var TEMPLATE_VERSIONS_HEADERS_ = [
  "template_id",
  "version",
  "created_at",
  "created_by",
  "source_quote_id",
  "items_json",
  "summary_json",
  "item_count",
  "metadata_snapshot_json"
];
var MATERIAL_TAG_TYPES_ = [
  "trade", "space", "mood", "tone", "texture",
  "style", "usage", "budget", "feature", "group_key", "group_label"
];
var MATERIAL_TAG_SCORE_MULTIPLIERS_ = {
  trade: 8,
  space: 6,
  mood: 5,
  tone: 5,
  texture: 4,
  style: 4,
  usage: 3,
  budget: 3,
  feature: 4,
  group_key: 2,
  group_label: 1
};
var PREQUOTE_TEMPLATE_TYPES_ = {
  QUOTE_ONLY: "QUOTE_ONLY",
  QUOTE: "QUOTE_ONLY",
  PREQUOTE_PACKAGE: "PREQUOTE_PACKAGE",
  BOTH: "BOTH"
};
var TEMPLATE_HOUSING_TYPE_OPTIONS_ = [
  "ALL",
  "APARTMENT",
  "HOUSE",
  "VILLA",
  "OFFICETEL",
  "COMMERCIAL",
  "OFFICE"
];
var TEMPLATE_AREA_BAND_OPTIONS_ = [
  "ALL",
  "UNDER_20",
  "20_29",
  "30_39",
  "40_49",
  "50_PLUS"
];
var TEMPLATE_BUDGET_BAND_OPTIONS_ = [
  "ANY",
  "ENTRY",
  "MID",
  "PREMIUM",
  "HIGH_END"
];
var TEMPLATE_META_FIELD_NAMES_ = [
  "template_type",
  "housing_type",
  "area_band",
  "style_tags_summary",
  "tone_tags_summary",
  "trade_scope_summary",
  "expose_to_prequote",
  "prequote_priority",
  "sort_order",
  "recommendation_note",
  "target_customer_summary",
  "budget_band",
  "recommended_for_summary",
  "is_featured_prequote"
];
var TEMPLATE_META_DEFAULTS_ = {
  template_type: PREQUOTE_TEMPLATE_TYPES_.QUOTE_ONLY,
  housing_type: "ALL",
  area_band: "ALL",
  style_tags_summary: "",
  tone_tags_summary: "",
  trade_scope_summary: "",
  expose_to_prequote: false,
  prequote_priority: 0,
  sort_order: 0,
  recommendation_note: "",
  target_customer_summary: "",
  budget_band: "ANY",
  recommended_for_summary: "",
  is_featured_prequote: false
};
var SLACK_EVENT_TYPES_ = {
  CHANGE_REQUEST_MEMO: EVENT_CHANGE_REQUEST_MEMO_,
  APPROVAL_TOGGLED: EVENT_APPROVAL_TOGGLED_,
  OWNER_APPROVAL_EMAIL: EVENT_OWNER_APPROVAL_EMAIL_,
  OWNER_NOTE_EMAIL: EVENT_OWNER_NOTE_EMAIL_
};
var SLACK_EVENT_SCHEMA_DOCS_ = [
  {
    event_type: EVENT_CHANGE_REQUEST_MEMO_,
    channel: "memo",
    payload_keys: ["quoteId", "memo", "customerName", "siteName", "deepLink", "eventTimeIso"]
  },
  {
    event_type: EVENT_APPROVAL_TOGGLED_,
    channel: "approve",
    payload_keys: ["quoteId", "oldStatus", "newStatus", "customerName", "siteName", "deepLink", "eventTimeIso"]
  },
  {
    event_type: EVENT_OWNER_APPROVAL_EMAIL_,
    channel: "approve",
    payload_keys: ["quoteId"]
  },
  {
    event_type: EVENT_OWNER_NOTE_EMAIL_,
    channel: "memo",
    payload_keys: ["quoteId"]
  }
];

// Compatibility fallback for environments where utils.js has not yet been updated.
if (typeof getSheetFromSs_ !== "function") {
  var getSheetFromSs_ = function(ss, name) {
    var key = String(name || "").trim();
    if (!key) throw new Error("Sheet name required");
    if (!ss || typeof ss.getSheetByName !== "function") throw new Error("Spreadsheet instance required");
    var sh = ss.getSheetByName(key);
    if (!sh) throw new Error("Missing sheet: " + key);
    return sh;
  };
}

if (typeof readSettingsMap_ !== "function") {
  var readSettingsMap_ = function(ss) {
    var spreadsheet = ss || getSpreadsheet_();
    var sh = spreadsheet.getSheetByName("Settings");
    if (!sh) {
      sh = spreadsheet.insertSheet("Settings");
      sh.getRange(1, 1, 1, 2).setValues([["key", "value"]]);
    }
    var map = {};
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      var values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
      for (var i = 0; i < values.length; i++) {
        var k = String(values[i][0] || "").trim();
        if (!k) continue;
        map[k] = String(values[i][1] || "").trim();
      }
    }
    return map;
  };
}

if (typeof getSettingFromSs_ !== "function") {
  var getSettingFromSs_ = function(ss, key, defaultValue) {
    var k = String(key || "").trim();
    if (!k) return defaultValue;
    var map = readSettingsMap_(ss);
    if (!Object.prototype.hasOwnProperty.call(map, k)) return defaultValue;
    var value = map[k];
    if (value === undefined || value === null || String(value).trim() === "") return defaultValue;
    return value;
  };
}

function doGet(e) {
  var page = String((e && e.parameter && e.parameter.page) || "edit").toLowerCase();
  if (page === "admin") page = "edit";

  try {
    if (page === "buildinfo") return renderBuildInfo_();
    // Only public image serving is anonymous by policy.
    if (page === "img") return servePublicImage_(e);
    if (page === "image") return serveQuoteImage_(e);

    if (isAdminPage_(page)) {
      var adminCred = getAdminCredentialFromRequest_(e);
      var auth = null;
      var authErr = null;
      try {
        if (adminCred) auth = assertAdminCredential_(adminCred);
      } catch (eAuth) {
        authErr = eAuth;
      }
      if (STRICT_ADMIN_GET_AUTH_ && !auth) {
        return renderAdminAuthRequiredPage_(page, authErr || new Error("Admin credential required"));
      }

      var appUrl = normalizeWebAppBaseUrl_(getAppUrl_());
      var tplAdmin = createAdminTemplateForPage_(page);
      tplAdmin.error = null;
      tplAdmin.app_url = appUrl;
      tplAdmin.vat_rate = 0.1;
      tplAdmin.default_vat_included = "N";
      tplAdmin.prefill_admin_password = "";
      if (auth && auth.type === "session") {
        tplAdmin.prefill_admin_password = String(auth.token || "");
      } else if (auth && auth.type === "password") {
        try {
          var tok = createAdminSessionToken_();
          CacheService.getScriptCache().put(getAdminSessionCacheKey_(tok), "1", ADMIN_SESSION_TTL_SEC_);
          tplAdmin.prefill_admin_password = tok;
        } catch (eTok) {
          tplAdmin.prefill_admin_password = "";
        }
      }
      tplAdmin.prefill_template_id = String(
        (e && e.parameter && (e.parameter.template_id || e.parameter.templateId)) || ""
      );
      return tplAdmin.evaluate().setTitle("Quote Admin");
    }

    if (page === "view" || page === "print") {
      var tpl = HtmlService.createTemplateFromFile(page);
      tpl.error = null;
      tpl.data = null;
      tpl.app_url = "";

      try {
        var quoteId = String((e && e.parameter && e.parameter.quoteId) || "").trim();
        var token = String((e && e.parameter && e.parameter.token) || "").trim();
        if (!quoteId || !token) throw new Error("Missing quoteId/token");
        tpl.data = getQuote(quoteId, token, {
          include_notes: false,
          include_photos: false,
          note_limit: 100
        });
      } catch (errInner) {
        tpl.error = formatErr_(errInner);
      }

      return tpl.evaluate().setTitle(page === "view" ? "Quote View" : "Quote Print");
    }

    var available = ["edit", "catalog", "templates", "templateslist", "dashboard", "materialgroups", "view", "print", "image", "img"];
    var serviceUrl = "";
    try { serviceUrl = String(ScriptApp.getService().getUrl() || ""); } catch (e0) { serviceUrl = ""; }
    var unknownHtml =
      "<div style='font-family:ui-monospace,monospace;white-space:pre-wrap;padding:16px'>" +
      "<b>Unknown page</b>\n" +
      "requested page = " + escapeHtml_(page) + "\n" +
      "available pages = " + escapeHtml_(available.join(", ")) + "\n\n" +
      "service url (runtime) = " + escapeHtml_(serviceUrl) +
      "</div>";
    return HtmlService.createHtmlOutput(unknownHtml).setTitle("Unknown");
  } catch (err) {
    var msg = formatErr_(err);
    return HtmlService.createHtmlOutput(
      "<div style='font-family:ui-monospace,monospace;white-space:pre-wrap;padding:16px'>" +
      "<b>Server Render Error</b>\n\n" + escapeHtml_(msg) +
      "</div>"
    ).setTitle("Error");
  }
}

function renderBuildInfo_() {
  var runtimeUrl = "";
  var appUrl = "";
  var scriptId = "";
  try { runtimeUrl = String(ScriptApp.getService().getUrl() || ""); } catch (e1) { runtimeUrl = ""; }
  try { appUrl = String(getAppUrl_() || ""); } catch (e2) { appUrl = ""; }
  try { scriptId = String(ScriptApp.getScriptId() || ""); } catch (e3) { scriptId = ""; }
  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      build_tag: BUILD_TAG_,
      strict_admin_get_auth: STRICT_ADMIN_GET_AUTH_,
      runtime_url: runtimeUrl,
      app_url: appUrl,
      script_id: scriptId
    }, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

function isAdminPage_(page) {
  var p = String(page || "").trim().toLowerCase();
  return (
    p === "edit" ||
    p === "catalog" ||
    p === "dashboard" ||
    p === "templates" ||
    p === "templateslist" ||
    p === "materialgroups"
  );
}

function getAdminCredentialFromRequest_(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  return String(
    p.admin_auth ||
    p.admin_pw ||
    p.adminPassword ||
    p.admin_password ||
    ""
  ).trim();
}

function renderAdminAuthRequiredPage_(page, err) {
  var p = String(page || "edit").trim().toLowerCase();
  var msg = String((err && err.message) ? err.message : "Admin credential required");
  var html =
    "<!doctype html><html><head><meta charset='UTF-8'><meta name='viewport' content='width=device-width,initial-scale=1'>" +
    "<title>401 Unauthorized</title></head><body style='font-family:Arial,sans-serif;background:#f7f8fb;padding:24px;'>" +
    "<div style='max-width:440px;margin:0 auto;background:#fff;border:1px solid #dde3ee;border-radius:12px;padding:20px;'>" +
    "<h2 style='margin:0 0 10px;'>401 Unauthorized</h2>" +
    "<p style='margin:0 0 12px;color:#586070;'>Admin authentication is required.</p>" +
    "<p style='margin:0 0 16px;color:#8a2b2b;font-size:13px;'>" + escapeHtml_(msg) + "</p>" +
    "<form method='get'>" +
    "<input type='hidden' name='page' value='" + escapeHtml_(p) + "'>" +
    "<label style='display:block;font-size:13px;margin:0 0 6px;'>Admin password or session token</label>" +
    "<input type='password' name='admin_password' autocomplete='current-password' style='width:100%;box-sizing:border-box;padding:10px;border:1px solid #ccd4e0;border-radius:8px;margin:0 0 12px;'>" +
    "<button type='submit' style='padding:10px 14px;border:0;border-radius:8px;background:#1f5eff;color:#fff;'>Continue</button>" +
    "</form></div></body></html>";
  return HtmlService.createHtmlOutput(html).setTitle("401 Unauthorized");
}

function createAdminTemplateForPage_(page) {
  var p = String(page || "").trim();
  try {
    return HtmlService.createTemplateFromFile(p);
  } catch (err) {
    // Graceful fallback when a newly added HTML file is not yet deployed.
    if (p === "templateslist") {
      var fallback = HtmlService.createTemplateFromFile("templates");
      fallback.prefill_missing_page = "templateslist";
      return fallback;
    }
    if (p === "materialgroups") {
      return HtmlService.createTemplateFromFile("dashboard");
    }
    throw err;
  }
}

function formatErr_(err) {
  if (!err) return "Unknown error";
  if (err.stack) return String(err.stack);
  if (err.message) return String(err.message);
  return String(err);
}

function getAppUrl_() {
  var cacheKey = "APP_URL_V2";
  try {
    var cache = CacheService.getScriptCache();
    var cached = String(cache.get(cacheKey) || "").trim();
    if (cached) return cached;

    var props = PropertiesService.getScriptProperties();
    var propUrl = normalizeWebAppBaseUrl_(props.getProperty("WEBAPP_URL_CACHE"));
    if (propUrl) {
      cache.put(cacheKey, propUrl, 21600);
      return propUrl;
    }

    var runtimeUrl = normalizeWebAppBaseUrl_(String(ScriptApp.getService().getUrl() || ""));
    if (runtimeUrl) {
      var forLinks = /\/dev$/i.test(runtimeUrl) ? runtimeUrl.replace(/\/dev$/i, "/exec") : runtimeUrl;
      cache.put(cacheKey, forLinks, 21600);
      try { props.setProperty("WEBAPP_URL_CACHE", forLinks); } catch (e1) {}
      return forLinks;
    }

    var manualUrl = normalizeWebAppBaseUrl_(props.getProperty("BASE_WEBAPP_URL"));
    if (manualUrl) return manualUrl;
  } catch (e) {}
  return normalizeWebAppBaseUrl_(DEFAULT_PUBLIC_EXEC_BASE_URL_) || "";
}

function normalizeBaseUrl_(u) {
  var s = String(u || "").trim();
  if (!s) return "";
  var q = s.indexOf("?");
  if (q >= 0) s = s.slice(0, q);
  var h = s.indexOf("#");
  if (h >= 0) s = s.slice(0, h);
  return s.replace(/\/+$/, "");
}

function normalizeWebAppBaseUrl_(u) {
  var s = normalizeBaseUrl_(u);
  if (!s) return "";
  if (/\/(?:exec|dev)$/i.test(s)) return s;
  if (/\/userCodeAppPanel$/i.test(s)) return s.replace(/\/userCodeAppPanel$/i, "/exec");
  return "";
}

function getConfiguredBaseUrl_(settingsOrUrl) {
  var raw = (settingsOrUrl && typeof settingsOrUrl === "object")
    ? settingsOrUrl.base_url
    : settingsOrUrl;
  var configured = normalizeWebAppBaseUrl_(raw);
  if (configured) return configured;
  return normalizeWebAppBaseUrl_(getAppUrl_());
}

function ensureNotificationSchemaReady_() {
  if (__NOTIFICATION_SCHEMA_READY_THIS_EXEC_) return;
  ensureSheetColumns_("SlackUsers", ["assignee_name", "slack_member_id", "aliases"]);
  ensureSheetColumns_("NotificationDedup", ["quoteId", "event_type", "last_dedup_key", "last_sent_at"]);
  __NOTIFICATION_SCHEMA_READY_THIS_EXEC_ = true;
}

function ensureSlackQueueSchemaReady_() {
  if (__SLACK_QUEUE_SCHEMA_READY_THIS_EXEC_) return;
  ensureSheetColumns_("SlackQueue", [
    "queue_id",
    "created_at",
    "next_attempt_at",
    "status",
    "event_type",
    "quote_id",
    "dedup_key",
    "payload_json",
    "attempt_count",
    "last_attempt_at",
    "last_http_code",
    "last_error",
    "sent_at"
  ]);
  __SLACK_QUEUE_SCHEMA_READY_THIS_EXEC_ = true;
}

function getSlackQueueOptions_() {
  var s = getSettings_();
  var batchSize = Math.min(Math.max(Number(s.slack_queue_batch_size || 30), 1), 200);
  var maxAttempts = Math.min(Math.max(Number(s.slack_retry_max_attempts || 6), 1), 20);
  var baseDelaySec = Math.min(Math.max(Number(s.slack_retry_base_seconds || 60), 5), 3600);
  var maxDelaySec = Math.min(Math.max(Number(s.slack_retry_max_seconds || 1800), baseDelaySec), 86400);
  return {
    batchSize: batchSize,
    maxAttempts: maxAttempts,
    baseDelaySec: baseDelaySec,
    maxDelaySec: maxDelaySec
  };
}

function calcSlackBackoffSeconds_(attemptCount, baseDelaySec, maxDelaySec) {
  var n = Math.max(Number(attemptCount || 1), 1);
  var base = Math.max(Number(baseDelaySec || 60), 1);
  var max = Math.max(Number(maxDelaySec || 1800), base);
  var delay = base * Math.pow(2, n - 1);
  if (delay > max) delay = max;
  return Math.round(delay);
}

function normalizeQueueStatus_(rawStatus) {
  var s = String(rawStatus || "").toLowerCase().trim();
  if (s === "sent" || s === "dead" || s === "retrying" || s === "queued") return s;
  return "queued";
}

function withScriptLock_(waitMs, fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(Math.max(Number(waitMs || 10000), 1000));
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function withUserLock_(waitMs, fn, onBusy) {
  var lock = null;
  try { lock = LockService.getUserLock(); } catch (e0) { lock = null; }
  if (!lock) return fn();
  var timeout = Math.max(Number(waitMs || 0), 0);
  var acquired = false;
  try { acquired = lock.tryLock(timeout); } catch (e1) { acquired = false; }
  if (!acquired) {
    if (typeof onBusy === "function") return onBusy();
    return { ok: true, skipped: true, reason: "user_lock_busy" };
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function normalizeSpaces_(text) {
  return String(text || "").replace(/\s+/g, " ").trim();
}

function normalizeAssigneeName_(name) {
  return normalizeSpaces_(name);
}

function normalizeSlackMemberId_(memberId) {
  var id = String(memberId || "").trim();
  if (!id) return "";
  id = id.replace(/^<@/, "").replace(/>$/, "").trim();
  return /^[UW][A-Z0-9]+$/.test(id) ? id : "";
}

function parseJsonSafe_(raw, fallback) {
  if (!raw) return fallback;
  try { return JSON.parse(raw); } catch (e) { return fallback; }
}

function normalizeMemoForDedup_(memo) {
  return String(memo || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function sha256Hex_(text) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(text || ""),
    Utilities.Charset.UTF_8
  );
  var out = [];
  for (var i = 0; i < bytes.length; i++) {
    var v = bytes[i];
    if (v < 0) v += 256;
    out.push((v < 16 ? "0" : "") + v.toString(16));
  }
  return out.join("");
}

function truncateText_(text, maxLen) {
  var s = String(text || "");
  var max = Math.max(Number(maxLen || 0), 0);
  if (!max || s.length <= max) return s;
  return s.slice(0, max) + "...";
}

function escapeSlackMrkdwn_(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeSlackMrkdwnAllowMention_(text) {
  var raw = String(text || "");
  var kept = [];
  raw = raw.replace(/<@[UW][A-Z0-9]+>/g, function(m) {
    var token = "__MENTION_" + kept.length + "__";
    kept.push({ token: token, mention: m });
    return token;
  });
  var escaped = escapeSlackMrkdwn_(raw);
  for (var i = 0; i < kept.length; i++) {
    escaped = escaped.replace(kept[i].token, kept[i].mention);
  }
  return escaped;
}

function formatEventIso_(dateObj) {
  var d = dateObj instanceof Date ? dateObj : new Date(dateObj || Date.now());
  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd'T'HH:mm:ssXXX"
  );
}

function approvalStateFromQuote_(quoteObj) {
  var q = quoteObj || {};
  var st = String(q.status || "").toLowerCase();
  var approvedAt = String(q.approved_at || "").trim();
  return (st === "approved" || !!approvedAt) ? "approved" : "not_approved";
}

function approvalStateLabel_(state) {
  return String(state || "").toLowerCase() === "approved" ? "approved" : "not_approved";
}

function approvalStateLabelKr_(state) {
  return String(state || "").toLowerCase() === "approved"
    ? "\uD655\uC815"
    : "\uBBF8\uD655\uC815";
}

function normalizeApprovalTarget_(raw, fallbackState) {
  var s = String(raw || "").toLowerCase().trim();
  if (s === "approved" || s === "approve" || s === "y") return "approved";
  if (s === "not_approved" || s === "unapproved" || s === "not-approved" || s === "n") return "not_approved";
  return fallbackState || "";
}

function buildDeepLinkUrl_(quoteId, token) {
  var qid = String(quoteId || "").trim();
  var tok = String(token || "").trim();
  if (!qid || !tok) return "";
  var base = getConfiguredBaseUrl_(getSettings_());
  if (!base) base = normalizeWebAppBaseUrl_(getAppUrl_());
  if (!base) return "";
  return base + "?page=view&quoteId=" + encodeURIComponent(qid) + "&token=" + encodeURIComponent(tok);
}

function pickFirstNonEmpty_() {
  for (var i = 0; i < arguments.length; i++) {
    var s = String(arguments[i] || "").trim();
    if (s) return s;
  }
  return "";
}

function getSlackUsersMap_() {
  ensureNotificationSchemaReady_();
  var cacheKey = "SLACK_USERS_MAP_V1";
  var cached = getCachedJson_(cacheKey);
  if (cached && typeof cached === "object") return cached;

  var out = Object.create(null);
  var sh = getSheet_("SlackUsers");
  var cols = getColMap_("SlackUsers");
  var lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    var headers = getHeaders_("SlackUsers");
    var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (var i = 0; i < values.length; i++) {
      var assignee = normalizeAssigneeName_(values[i][cols.assignee_name]);
      var slackId = normalizeSlackMemberId_(values[i][cols.slack_member_id]);
      if (!slackId) continue;

      if (assignee) out[assignee] = slackId;

      var aliases = String(values[i][cols.aliases] || "").split(",");
      for (var a = 0; a < aliases.length; a++) {
        var alias = normalizeAssigneeName_(aliases[a]);
        if (alias) out[alias] = slackId;
      }
    }
  }

  putCachedJson_(cacheKey, out, 300);
  return out;
}

function getSlackWebhooks_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty("SLACK_WEBHOOK_URL");

  var rawText = String(raw || "").trim();
  var memo = "";
  var approve = "";

  if (rawText) {
    if (/^https?:\/\//i.test(rawText)) {
      memo = rawText;
      approve = rawText;
    } else {
      var parsed = parseJsonSafe_(rawText, null);
      if (typeof parsed === "string") {
        var one = String(parsed || "").trim();
        if (/^https?:\/\//i.test(one)) {
          memo = one;
          approve = one;
        }
      } else if (parsed && typeof parsed === "object") {
        memo = String(
          parsed.memo ||
          parsed.change_request_memo ||
          parsed.changeRequestMemo ||
          parsed.note ||
          parsed.default ||
          parsed.url ||
          ""
        ).trim();
        approve = String(
          parsed.approve ||
          parsed.approval ||
          parsed.approval_toggled ||
          parsed.approvalToggled ||
          parsed.confirm ||
          parsed.default ||
          parsed.url ||
          ""
        ).trim();
      }
    }
  }

  if (!memo) {
    memo = String(
      props.getProperty("SLACK_WEBHOOK_URL_MEMO") ||
      props.getProperty("SLACK_WEBHOOK_MEMO") ||
      props.getProperty("SLACK_WEBHOOK_URL_NOTE") ||
      ""
    ).trim();
  }
  if (!approve) {
    approve = String(
      props.getProperty("SLACK_WEBHOOK_URL_APPROVE") ||
      props.getProperty("SLACK_WEBHOOK_APPROVE") ||
      props.getProperty("SLACK_WEBHOOK_URL_APPROVAL") ||
      ""
    ).trim();
  }
  if ((!memo || !approve) && /^https?:\/\//i.test(rawText)) {
    if (!memo) memo = rawText;
    if (!approve) approve = rawText;
  }

  if (!memo && approve) memo = approve;
  if (!approve && memo) approve = memo;
  return { memo: memo, approve: approve };
}

function resolveSlackMention_(assigneeName) {
  ensureNotificationSchemaReady_();

  var normalized = normalizeAssigneeName_(assigneeName);
  if (!normalized) return "";

  var usersMap = getSlackUsersMap_();
  var memberId = normalizeSlackMemberId_(usersMap[normalized]);
  if (!memberId) {
    var rawFallback = PropertiesService.getScriptProperties().getProperty("SLACK_MEMBER_MAP");
    var fallbackMap = parseJsonSafe_(rawFallback, {});
    memberId = normalizeSlackMemberId_(fallbackMap && fallbackMap[normalized]);
    if (!memberId && fallbackMap && typeof fallbackMap === "object") {
      for (var k in fallbackMap) {
        if (normalizeAssigneeName_(k) !== normalized) continue;
        memberId = normalizeSlackMemberId_(fallbackMap[k]);
        if (memberId) break;
      }
    }
  }

  if (!memberId) return normalized;
  return "<@" + memberId + ">";
}

function buildSlackMessage_(eventType, payload) {
  var p = payload || {};
  var qid = String(p.quoteId || p.quote_id || "").trim();
  var customerName = String(
    p.customerName || p.customer_name || "(\uC774\uB984 \uC5C6\uC74C)"
  ).trim();
  var phone = String(p.customerPhone || p.customer_phone || "-").trim();
  var email = String(p.customerEmail || p.customer_email || "-").trim();
  var siteName = String(p.siteName || p.site_name || "-").trim();
  var deepLink = String(p.deepLink || p.deep_link || "").trim();
  var assignee = String(p.assigneeDisplay || p.assignee || "(\uBBF8\uC9C0\uC815)").trim();
  assignee = assignee
    .replace(/^Assignee:\s*/i, "")
    .replace(/^\uB2F4\uB2F9\uC790:\s*/i, "")
    .trim() || "(\uBBF8\uC9C0\uC815)";
  var total = Number(p.total || 0);
  var eventTimeIso = String(p.eventTimeIso || formatEventIso_(new Date())).trim();

  var headerText = "\uD83D\uDCDD \uACE0\uAC1D \uBA54\uBAA8 \uC811\uC218";
  var statusLine = "-";
  var fallbackText = "[\uACE0\uAC1D \uBA54\uBAA8] " + customerName + " | " + (qid || "-");
  var isMemoEvent = eventType === EVENT_CHANGE_REQUEST_MEMO_;
  var isApprovalEvent = eventType === EVENT_APPROVAL_TOGGLED_;

  if (isApprovalEvent) {
    var oldState = normalizeApprovalTarget_(p.oldStatus, "not_approved");
    var newState = normalizeApprovalTarget_(p.newStatus, "not_approved");
    statusLine = approvalStateLabelKr_(oldState) + " -> " + approvalStateLabelKr_(newState);
    headerText = (newState === "approved")
      ? "\u2705 \uACE0\uAC1D \uD655\uC815 \uCC98\uB9AC"
      : "\uD83D\uDD04 \uACE0\uAC1D \uD655\uC815 \uC0C1\uD0DC \uBCC0\uACBD";
    fallbackText = "[\uD655\uC815 \uC0C1\uD0DC \uBCC0\uACBD] " +
      customerName + " | " + (qid || "-") + " | " + statusLine;
  } else if (isMemoEvent) {
    statusLine = approvalStateLabelKr_(normalizeApprovalTarget_(p.currentStatus, "not_approved"));
  } else {
    headerText = "\uD83D\uDD14 \uC2DC\uC2A4\uD15C \uC774\uBCA4\uD2B8";
    statusLine = String(eventType || "UNKNOWN");
    fallbackText = "[\uC774\uBCA4\uD2B8] " + (eventType || "UNKNOWN") + " | " + (qid || "-");
  }

  var summary = [
    "*\uD83D\uDCCC \uACAC\uC801 ID*",
    "`" + escapeSlackMrkdwn_(qid || "-") + "`",
    "",
    "*\uACE0\uAC1D\uBA85* " + escapeSlackMrkdwn_(customerName),
    "*\uD604\uC7A5\uBA85* " + escapeSlackMrkdwn_(siteName),
    "*\uB2F4\uB2F9\uC790* " + escapeSlackMrkdwnAllowMention_(assignee),
    "*\uC0C1\uD0DC* " + escapeSlackMrkdwn_(statusLine || "-")
  ].join("\n");

  var meta = [
    "\uD83D\uDCDE \uC5F0\uB77D\uCC98: " + escapeSlackMrkdwn_(phone || "-"),
    "\u2709\uFE0F \uC774\uBA54\uC77C: " + escapeSlackMrkdwn_(email || "-"),
    "\uD83D\uDD52 \uBC1C\uC0DD\uC2DC\uAC01: " + escapeSlackMrkdwn_(eventTimeIso || "-")
  ];
  if (!isNaN(total) && total > 0) {
    meta.push("\uD83D\uDCB0 \uACAC\uC801\uAE08\uC561: " +
      escapeSlackMrkdwn_(total.toLocaleString("ko-KR") + " \uC6D0"));
  }
  meta.push("\uD83C\uDFF7\uFE0F \uD15C\uD50C\uB9BF: " + escapeSlackMrkdwn_(SLACK_MESSAGE_TEMPLATE_VERSION_));

  var blocks = [
    {
      type: "header",
      text: { type: "plain_text", text: headerText, emoji: true }
    },
    { type: "section", text: { type: "mrkdwn", text: summary } },
    { type: "context", elements: [{ type: "mrkdwn", text: meta.join(" | ") }] }
  ];

  if (isMemoEvent) {
    var memoRaw = String(p.memo || "");
    var memoSafe = truncateText_(memoRaw, 1500).replace(/```/g, "'''");
    blocks.push({
      type: "section",
      text: {
        type: "mrkdwn",
        text: "*\uD83D\uDCDD \uACE0\uAC1D \uBA54\uBAA8*\n```" +
          escapeSlackMrkdwn_(memoSafe || "(\uB0B4\uC6A9 \uC5C6\uC74C)") + "```"
      }
    });
  }

  blocks.push({ type: "divider" });

  if (deepLink) {
    blocks.push({
      type: "actions",
      elements: [{
        type: "button",
        text: {
          type: "plain_text",
          text: "\uD83D\uDCC4 \uACAC\uC801\uC11C \uC5F4\uAE30",
          emoji: true
        },
        url: deepLink
      }]
    });
  }

  return { text: fallbackText, blocks: blocks };
}

function getSlackEventSchemaCatalog() {
  return SLACK_EVENT_SCHEMA_DOCS_.map(function(row) {
    return {
      event_type: row.event_type,
      channel: row.channel,
      payload_keys: Array.isArray(row.payload_keys) ? row.payload_keys.slice() : []
    };
  });
}

function sendSlack_(webhookUrl, body) {
  var url = String(webhookUrl || "").trim();
  if (!url) throw new Error("Slack webhook URL missing");

  var payload = body || {};
  var maxTry = 3;
  var lastErr = null;
  var lastCode = 0;
  var lastBody = "";

  for (var i = 0; i < maxTry; i++) {
    try {
      var res = UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      lastCode = Number(res.getResponseCode() || 0);
      lastBody = String(res.getContentText() || "");
      if (lastCode >= 200 && lastCode < 300) {
        return { ok: true, httpCode: lastCode };
      }
      lastErr = new Error("Slack HTTP " + lastCode);
    } catch (err) {
      lastErr = err;
    }
    if (i < maxTry - 1) Utilities.sleep(Math.pow(2, i) * 300);
  }

  return {
    ok: false,
    httpCode: lastCode,
    error: lastErr ? String(lastErr.message || lastErr) : "Unknown error",
    responseBody: truncateText_(lastBody, 300)
  };
}

function findDedupRowNo_(quoteId, eventType) {
  ensureNotificationSchemaReady_();

  var qid = String(quoteId || "").trim();
  var ev = String(eventType || "").trim();
  if (!qid || !ev) return 0;

  var sh = getSheet_("NotificationDedup");
  var cols = getColMap_("NotificationDedup");

  var cache = CacheService.getScriptCache();
  var cacheKey = "NOTI_DEDUP_ROWNO_" + qid + "|" + ev;
  var cachedRowNo = Number(cache.get(cacheKey) || 0);
  if (cachedRowNo >= 2) {
    var probe = sh.getRange(cachedRowNo, 1, 1, Math.max(cols.quoteId, cols.event_type) + 1).getValues()[0];
    if (normalizeSpaces_(probe[cols.quoteId]) === qid && normalizeSpaces_(probe[cols.event_type]) === ev) {
      return cachedRowNo;
    }
  }

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  var qCol = cols.quoteId + 1;
  var evCol = cols.event_type + 1;
  var startCol = Math.min(qCol, evCol);
  var width = Math.abs(qCol - evCol) + 1;
  var qPos = qCol - startCol;
  var evPos = evCol - startCol;
  var values = sh.getRange(2, startCol, lastRow - 1, width).getValues();

  for (var i = 0; i < values.length; i++) {
    if (normalizeSpaces_(values[i][qPos]) !== qid) continue;
    if (normalizeSpaces_(values[i][evPos]) !== ev) continue;
    var rowNo = i + 2;
    cache.put(cacheKey, String(rowNo), 21600);
    return rowNo;
  }
  return 0;
}

function dedupShouldSendAndStore_(quoteId, eventType, dedupKey, now) {
  ensureNotificationSchemaReady_();

  var qid = String(quoteId || "").trim();
  var ev = String(eventType || "").trim();
  var key = String(dedupKey || "").trim();
  if (!qid || !ev || !key) return false;

  var nowDate = now instanceof Date ? now : new Date(now || Date.now());
  if (isNaN(nowDate.getTime())) nowDate = new Date();
  var nowIso = nowDate.toISOString();

  var sh = getSheet_("NotificationDedup");
  var headers = getHeaders_("NotificationDedup");
  var cols = getColMap_("NotificationDedup");
  var rowNo = findDedupRowNo_(qid, ev);

  if (rowNo) {
    var row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
    var lastKey = String(row[cols.last_dedup_key] || "").trim();
    var lastSentAt = String(row[cols.last_sent_at] || "").trim();
    if (lastKey === key) {
      if (ev !== EVENT_CHANGE_REQUEST_MEMO_) return false;
      var lastMs = parseIsoMs_(lastSentAt);
      if (!isNaN(lastMs) && (nowDate.getTime() - lastMs) <= MEMO_DEDUP_WINDOW_MS_) {
        return false;
      }
    }

    row[cols.quoteId] = qid;
    row[cols.event_type] = ev;
    row[cols.last_dedup_key] = key;
    row[cols.last_sent_at] = nowIso;
    sh.getRange(rowNo, 1, 1, headers.length).setValues([row]);
    return true;
  }

  var newRow = Array(headers.length).fill("");
  newRow[cols.quoteId] = qid;
  newRow[cols.event_type] = ev;
  newRow[cols.last_dedup_key] = key;
  newRow[cols.last_sent_at] = nowIso;

  var startRow = sh.getLastRow() + 1;
  ensureRowCapacity_(sh, startRow);
  sh.getRange(startRow, 1, 1, headers.length).setValues([newRow]);

  try {
    CacheService.getScriptCache().put(
      "NOTI_DEDUP_ROWNO_" + qid + "|" + ev,
      String(startRow),
      21600
    );
  } catch (e) {}

  return true;
}

function notifySlackEvent_(eventType, payload) {
  var hooks = getSlackWebhooks_();
  var webhookUrl = eventType === EVENT_CHANGE_REQUEST_MEMO_ ? hooks.memo : hooks.approve;
  if (!webhookUrl) {
    Logger.log(JSON.stringify({
      level: "error",
      quoteId: String(payload && payload.quoteId || ""),
      eventType: String(eventType || ""),
      message: "Slack webhook missing in script properties"
    }));
    return { ok: false, skipped: true, reason: "webhook_missing" };
  }

  var body = buildSlackMessage_(eventType, payload || {});
  var result = sendSlack_(webhookUrl, body);
  if (!result.ok) {
    Logger.log(JSON.stringify({
      level: "error",
      quoteId: String(payload && payload.quoteId || ""),
      eventType: String(eventType || ""),
      httpCode: Number(result.httpCode || 0),
      message: String(result.error || "Slack send failed")
    }));
  }
  return result;
}

function resolveSlackWebhookUrl_(eventType) {
  var hooks = getSlackWebhooks_();
  return eventType === EVENT_CHANGE_REQUEST_MEMO_ ? hooks.memo : hooks.approve;
}

function ensureSlackQueueTriggerInstalled_() {
  if (__SLACK_QUEUE_TRIGGER_CHECKED_THIS_EXEC_) return;
  __SLACK_QUEUE_TRIGGER_CHECKED_THIS_EXEC_ = true;

  try {
    var cache = CacheService.getScriptCache();
    if (cache.get("TRG_SLACKQ_OK") === "1") return;

    var triggers = ScriptApp.getProjectTriggers();
    var found = false;
    for (var i = 0; i < triggers.length; i++) {
      var t = triggers[i];
      if (t.getHandlerFunction && t.getHandlerFunction() === "runSlackQueueWorker_") {
        found = true;
        break;
      }
    }
    if (!found) {
      ScriptApp.newTrigger("runSlackQueueWorker_").timeBased().everyMinutes(1).create();
    }
    cache.put("TRG_SLACKQ_OK", "1", 300);
  } catch (e) {
    Logger.log(JSON.stringify({
      level: "error",
      eventType: "SLACK_QUEUE_TRIGGER",
      message: String((e && e.message) || e || "trigger install failed")
    }));
  }
}

function scheduleSlackQueueKick_() {
  try {
    var cache = CacheService.getScriptCache();
    if (cache.get("TRG_SLACKQ_KICK_PENDING") === "1") {
      return { ok: true, scheduled: false, reason: "kick_pending_cache" };
    }

    var triggers = ScriptApp.getProjectTriggers();
    var hasKick = false;
    for (var i = 0; i < triggers.length; i++) {
      var t = triggers[i];
      if (t.getHandlerFunction && t.getHandlerFunction() === "runSlackQueueKick_") {
        hasKick = true;
        break;
      }
    }
    if (hasKick) {
      cache.put("TRG_SLACKQ_KICK_PENDING", "1", 20);
      return { ok: true, scheduled: false, reason: "kick_already_scheduled" };
    }

    ScriptApp.newTrigger("runSlackQueueKick_").timeBased().after(5000).create();
    cache.put("TRG_SLACKQ_KICK_PENDING", "1", 20);
    return { ok: true, scheduled: true };
  } catch (e) {
    Logger.log(JSON.stringify({
      level: "error",
      eventType: "SLACK_QUEUE_KICK",
      message: String((e && e.message) || e || "kick schedule failed")
    }));
    return { ok: false, scheduled: false, reason: "kick_schedule_failed" };
  }
}

function runSlackQueueKick_() {
  try {
    return runSlackQueueWorker_();
  } finally {
    try { CacheService.getScriptCache().remove("TRG_SLACKQ_KICK_PENDING"); } catch (e) {}
  }
}

function enqueueSlackEvent_(eventType, payload, dedupKey, nowDate) {
  ensureSlackQueueSchemaReady_();
  ensureSlackQueueTriggerInstalled_();

  var ev = String(eventType || "").trim();
  if (!ev) return { ok: false, skipped: true, reason: "event_type_required" };

  var qid = String(payload && (payload.quoteId || payload.quote_id) || "").trim();
  var now = (nowDate instanceof Date) ? nowDate : new Date();
  if (isNaN(now.getTime())) now = new Date();
  var nowIso = now.toISOString();
  var dedup = String(dedupKey || "").trim();

  var dedupCacheKey = dedup ? ("SLACKQ_DEDUP_" + sha256Hex_(dedup)) : "";
  if (dedupCacheKey) {
    try {
      if (CacheService.getScriptCache().get(dedupCacheKey)) {
        return { ok: true, queued: false, skipped: true, reason: "dedup_cache" };
      }
    } catch (e) {}
  }

  try {
    var sh = getSheet_("SlackQueue");
    var headers = getHeaders_("SlackQueue");
    var cols = getColMap_("SlackQueue");
    var row = Array(headers.length).fill("");
    var queueId = uuid_();

    row[cols.queue_id] = queueId;
    row[cols.created_at] = nowIso;
    row[cols.next_attempt_at] = nowIso;
    row[cols.status] = "queued";
    row[cols.event_type] = ev;
    row[cols.quote_id] = qid;
    row[cols.dedup_key] = dedup;
    row[cols.payload_json] = JSON.stringify(payload || {});
    row[cols.attempt_count] = 0;
    row[cols.last_attempt_at] = "";
    row[cols.last_http_code] = "";
    row[cols.last_error] = "";
    row[cols.sent_at] = "";

    var startRow = sh.getLastRow() + 1;
    ensureRowCapacity_(sh, startRow);
    sh.getRange(startRow, 1, 1, headers.length).setValues([row]);

    if (dedupCacheKey) {
      var dedupTtl = (ev === EVENT_CHANGE_REQUEST_MEMO_) ? 120 : 3600;
      try { CacheService.getScriptCache().put(dedupCacheKey, "1", dedupTtl); } catch (e2) {}
    }
    try { scheduleSlackQueueKick_(); } catch (kickErr) {}

    return { ok: true, queued: true, queue_id: queueId, row_no: startRow };
  } catch (err) {
    Logger.log(JSON.stringify({
      level: "error",
      quoteId: qid,
      eventType: ev,
      message: String(err && err.message || err || "Slack enqueue failed")
    }));
    return { ok: false, queued: false, reason: "queue_write_failed" };
  }
}

function enqueueOwnerEmailTask_(eventType, quoteId, payload, dedupKey, nowDate) {
  var body = cloneObj_(payload || {});
  body.quoteId = String(quoteId || body.quoteId || "").trim();
  return enqueueSlackEvent_(eventType, body, dedupKey, nowDate);
}

function runSlackQueueWorker_() {
  ensureCoreSchemaReady_();
  ensureSlackQueueSchemaReady_();

  return withUserLock_(1000, function() {
    var opts = getSlackQueueOptions_();
    var sh = getSheet_("SlackQueue");
    var headers = getHeaders_("SlackQueue");
    var cols = getColMap_("SlackQueue");
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, scanned: 0, due: 0, sent: 0, retried: 0, dead: 0 };

    var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var now = new Date();
    var nowMs = now.getTime();
    var due = [];
    var maxDuePerRun = Math.min(Math.max(opts.batchSize * 2, opts.batchSize), 400);

    for (var i = 0; i < values.length && due.length < maxDuePerRun; i++) {
      var row = values[i];
      var status = normalizeQueueStatus_(row[cols.status]);
      if (status === "sent" || status === "dead") continue;

      var nextMs = parseIsoMs_(row[cols.next_attempt_at]);
      if (!isNaN(nextMs) && nextMs > nowMs) continue;

      due.push({ rowNo: i + 2, row: row });
    }

    if (!due.length) {
      return { ok: true, scanned: values.length, due: 0, sent: 0, retried: 0, dead: 0 };
    }

    var updateNos = [];
    var updateRows = [];
    var sent = 0;
    var retried = 0;
    var dead = 0;

    for (var d = 0; d < due.length; d++) {
      var item = due[d];
      var row0 = item.row.slice();
      var eventType = String(row0[cols.event_type] || "").trim();
      var payload = parseJsonSafe_(row0[cols.payload_json], {});
      var quoteId = String(row0[cols.quote_id] || "").trim();
      var dedupKey = String(row0[cols.dedup_key] || "").trim();
      var prevAttempts = Number(row0[cols.attempt_count] || 0);
      if (isNaN(prevAttempts) || prevAttempts < 0) prevAttempts = 0;

      var attempt = prevAttempts + 1;
      var nowIso = nowIso_();
      var sendAllowed = true;
      if (dedupKey && quoteId && eventType) {
        try {
          sendAllowed = dedupShouldSendAndStore_(quoteId, eventType, dedupKey, new Date());
        } catch (dedupErr) {
          sendAllowed = true;
        }
      }

      var sendRes = { ok: true, skipped: true, reason: "dedup_suppressed", httpCode: 0 };
      if (sendAllowed) {
        try {
          if (eventType === EVENT_OWNER_APPROVAL_EMAIL_) {
            notifyOwnerApproval_(String(payload && payload.quoteId || quoteId || "").trim());
            sendRes = { ok: true, httpCode: 200, provider: "owner_email" };
          } else if (eventType === EVENT_OWNER_NOTE_EMAIL_) {
            var qidForNote = String(payload && payload.quoteId || quoteId || "").trim();
            var shortNote = String(payload && payload.shortNote || "").trim();
            notifyOwnerCustomerNote_(qidForNote, shortNote);
            sendRes = { ok: true, httpCode: 200, provider: "owner_email" };
          } else {
            sendRes = notifySlackEvent_(eventType, payload || {});
          }
        } catch (sendErr) {
          sendRes = {
            ok: false,
            httpCode: 0,
            error: String(sendErr && sendErr.message || sendErr || "worker_send_failed")
          };
        }
      }

      row0[cols.attempt_count] = attempt;
      row0[cols.last_attempt_at] = nowIso;
      row0[cols.last_http_code] = Number(sendRes && sendRes.httpCode || 0);

      if (sendRes && sendRes.ok) {
        row0[cols.status] = "sent";
        row0[cols.sent_at] = nowIso;
        row0[cols.next_attempt_at] = "";
        row0[cols.last_error] = "";
        sent++;
      } else {
        if (attempt >= opts.maxAttempts) {
          row0[cols.status] = "dead";
          row0[cols.next_attempt_at] = "";
          row0[cols.last_error] = String(sendRes && sendRes.error || sendRes && sendRes.reason || "send_failed");
          dead++;
        } else {
          var delaySec = calcSlackBackoffSeconds_(attempt, opts.baseDelaySec, opts.maxDelaySec);
          row0[cols.status] = "retrying";
          row0[cols.next_attempt_at] = new Date(Date.now() + delaySec * 1000).toISOString();
          row0[cols.last_error] = String(sendRes && sendRes.error || sendRes && sendRes.reason || "send_failed");
          retried++;
        }
      }

      updateNos.push(item.rowNo);
      updateRows.push(row0);
    }

    writeRowsByRowNos_(sh, updateNos, updateRows, 1, headers.length);
    return {
      ok: true,
      scanned: values.length,
      due: due.length,
      sent: sent,
      retried: retried,
      dead: dead
    };
  }, function() {
    return { ok: true, skipped: true, reason: "worker_lock_busy" };
  });
}

function installSlackQueueTriggers(adminPassword) {
  assertAdminCredential_(adminPassword);
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var t = triggers[i];
    if (!t.getHandlerFunction) continue;
    var fn = t.getHandlerFunction();
    if (fn === "runSlackQueueWorker_" || fn === "runSlackQueueKick_") {
      ScriptApp.deleteTrigger(t);
    }
  }
  ScriptApp.newTrigger("runSlackQueueWorker_").timeBased().everyMinutes(1).create();
  try { CacheService.getScriptCache().put("TRG_SLACKQ_OK", "1", 300); } catch (e) {}
  return { ok: true };
}

function removeSlackQueueTriggers(adminPassword) {
  assertAdminCredential_(adminPassword);
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var t = triggers[i];
    if (!t.getHandlerFunction) continue;
    var fn = t.getHandlerFunction();
    if (fn === "runSlackQueueWorker_" || fn === "runSlackQueueKick_") {
      ScriptApp.deleteTrigger(t);
    }
  }
  try { CacheService.getScriptCache().remove("TRG_SLACKQ_OK"); } catch (e) {}
  try { CacheService.getScriptCache().remove("TRG_SLACKQ_KICK_PENDING"); } catch (e1) {}
  return { ok: true };
}

function getSlackQueueHealth(adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  ensureSlackQueueSchemaReady_();

  var sh = getSheet_("SlackQueue");
  var headers = getHeaders_("SlackQueue");
  var cols = getColMap_("SlackQueue");
  var lastRow = sh.getLastRow();
  var out = {
    ok: true,
    total_rows: Math.max(lastRow - 1, 0),
    status_counts: { queued: 0, retrying: 0, sent: 0, dead: 0, other: 0 },
    recent_errors: [],
    trigger_installed: false,
    webhooks: { memo: false, approve: false },
    oldest_pending_at: "",
    oldest_pending_age_sec: 0
  };

  try {
    var hooks = getSlackWebhooks_();
    out.webhooks.memo = !!String(hooks.memo || "").trim();
    out.webhooks.approve = !!String(hooks.approve || "").trim();
  } catch (e0) {}

  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var t = triggers[i];
    if (t.getHandlerFunction && t.getHandlerFunction() === "runSlackQueueWorker_") {
      out.trigger_installed = true;
      break;
    }
  }

  if (lastRow < 2) return out;

  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var nowMs = Date.now();
  var oldestMs = NaN;
  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    var status = normalizeQueueStatus_(row[cols.status]);
    if (status === "queued" || status === "retrying" || status === "sent" || status === "dead") {
      out.status_counts[status]++;
    } else {
      out.status_counts.other++;
    }
    if (status === "queued" || status === "retrying") {
      var nextMs = parseIsoMs_(row[cols.next_attempt_at]);
      if (!isNaN(nextMs) && (isNaN(oldestMs) || nextMs < oldestMs)) oldestMs = nextMs;
    }
  }
  if (!isNaN(oldestMs)) {
    out.oldest_pending_at = new Date(oldestMs).toISOString();
    out.oldest_pending_age_sec = Math.max(Math.floor((nowMs - oldestMs) / 1000), 0);
  }

  var start = Math.max(2, lastRow - 30 + 1);
  var recent = sh.getRange(start, 1, lastRow - start + 1, headers.length).getValues();
  for (var k = recent.length - 1; k >= 0; k--) {
    var rr = recent[k];
    var err = String(rr[cols.last_error] || "").trim();
    if (!err) continue;
    out.recent_errors.push({
      queue_id: String(rr[cols.queue_id] || ""),
      event_type: String(rr[cols.event_type] || ""),
      quote_id: String(rr[cols.quote_id] || ""),
      status: String(rr[cols.status] || ""),
      attempt_count: Number(rr[cols.attempt_count] || 0),
      last_http_code: Number(rr[cols.last_http_code] || 0),
      last_error: err
    });
    if (out.recent_errors.length >= 10) break;
  }

  return out;
}

function runSlackQueueNow(adminPassword) {
  assertAdminCredential_(adminPassword);
  return runSlackQueueWorker_();
}

function debugSlackDelivery(adminPassword) {
  if (!String(adminPassword || "").trim()) {
    throw new Error("Admin password required. In editor, run debugSlackDelivery_ADMIN().");
  }
  assertAdminCredential_(adminPassword);
  var healthBefore = getSlackQueueHealth(adminPassword);
  var worker = runSlackQueueWorker_();
  var healthAfter = getSlackQueueHealth(adminPassword);
  return {
    ok: true,
    template_version: SLACK_MESSAGE_TEMPLATE_VERSION_,
    before: healthBefore,
    worker: worker,
    after: healthAfter
  };
}

// Editor-run helper. Name ends with "_" so it is not callable from google.script.run.
function debugSlackWebhookConfig_ADMIN_() {
  var pw = getAdminPasswordFromSettingsOrThrow_();
  return debugSlackWebhookConfig(pw);
}

// Editor-run helper. Name ends with "_" so it is not callable from google.script.run.
function debugSlackDelivery_ADMIN_() {
  var pw = getAdminPasswordFromSettingsOrThrow_();
  return debugSlackDelivery(pw);
}

// Editor-run smoke test for webhook connectivity. Does not log secrets.
function smokeTestSlackDelivery_ADMIN_() {
  var pw = getAdminPasswordFromSettingsOrThrow_();
  assertAdminCredential_(pw);

  var hooks = getSlackWebhooks_();
  if (!hooks.memo && !hooks.approve) {
    throw new Error("Slack webhooks are not configured. Set SLACK_WEBHOOK_URL first.");
  }

  var stamp = nowIso_();
  var out = {
    ok: true,
    at: stamp,
    template_version: SLACK_MESSAGE_TEMPLATE_VERSION_,
    memo: { ok: false, httpCode: 0, error: "", response: "" },
    approve: { ok: false, httpCode: 0, error: "", response: "" }
  };

  if (hooks.memo) {
    var memoRes = sendSlack_(hooks.memo, { text: "[smoke] memo webhook check @ " + stamp });
    out.memo.ok = !!memoRes.ok;
    out.memo.httpCode = Number(memoRes.httpCode || 0);
    out.memo.error = truncateText_(String(memoRes.error || ""), 120);
    out.memo.response = truncateText_(String(memoRes.responseBody || ""), 120);
  } else {
    out.memo.error = "memo webhook missing";
  }

  if (hooks.approve) {
    var approveRes = sendSlack_(hooks.approve, { text: "[smoke] approve webhook check @ " + stamp });
    out.approve.ok = !!approveRes.ok;
    out.approve.httpCode = Number(approveRes.httpCode || 0);
    out.approve.error = truncateText_(String(approveRes.error || ""), 120);
    out.approve.response = truncateText_(String(approveRes.responseBody || ""), 120);
  } else {
    out.approve.error = "approve webhook missing";
  }

  Logger.log(JSON.stringify({
    ok: out.ok,
    at: out.at,
    memo: { ok: out.memo.ok, httpCode: out.memo.httpCode, error: out.memo.error, response: out.memo.response },
    approve: { ok: out.approve.ok, httpCode: out.approve.httpCode, error: out.approve.error, response: out.approve.response }
  }));
  return out;
}

// Editor-run smoke test using the real Slack template builder (Korean + emoji).
function smokeTestSlackTemplate_ADMIN_() {
  var pw = getAdminPasswordFromSettingsOrThrow_();
  assertAdminCredential_(pw);

  var hooks = getSlackWebhooks_();
  if (!hooks.memo || !hooks.approve) {
    throw new Error("Both memo/approve webhooks are required for template smoke test.");
  }

  var stamp = formatEventIso_(new Date());
  var qid = "TEMPLATE-" + String(uuid_()).slice(0, 8);
  var basePayload = {
    quoteId: qid,
    customerName: "\uD15C\uD50C\uB9BF \uD14C\uC2A4\uD2B8",
    customerPhone: "010-0000-0000",
    customerEmail: "test@example.com",
    siteName: "\uD14C\uC2A4\uD2B8 \uD604\uC7A5",
    assigneeDisplay: "<@U0AG70EJ492>",
    total: 12345000,
    eventTimeIso: stamp,
    deepLink: "https://example.com/quote?debug=1"
  };

  var memoBody = buildSlackMessage_(EVENT_CHANGE_REQUEST_MEMO_, Object.assign({}, basePayload, {
    currentStatus: "not_approved",
    memo: "\uC548\uBC29 \uC870\uBA85 \uCD94\uAC00 \uBC0F \uBC14\uB2E5 \uC790\uC7AC \uBCC0\uACBD \uC694\uCCAD"
  }));
  var approveBody = buildSlackMessage_(EVENT_APPROVAL_TOGGLED_, Object.assign({}, basePayload, {
    oldStatus: "not_approved",
    newStatus: "approved"
  }));

  var memoRes = sendSlack_(hooks.memo, memoBody);
  var approveRes = sendSlack_(hooks.approve, approveBody);
  return {
    ok: !!memoRes.ok && !!approveRes.ok,
    template_version: SLACK_MESSAGE_TEMPLATE_VERSION_,
    memo: {
      ok: !!memoRes.ok,
      httpCode: Number(memoRes.httpCode || 0),
      error: truncateText_(String(memoRes.error || ""), 120)
    },
    approve: {
      ok: !!approveRes.ok,
      httpCode: Number(approveRes.httpCode || 0),
      error: truncateText_(String(approveRes.error || ""), 120)
    }
  };
}

// Visible launchers for Apps Script Run menu.
// They still enforce editor-admin identity and keep web-app password checks unchanged.
function debugSlackWebhookConfig_ADMIN() {
  assertEditorAdminExecution_();
  return debugSlackWebhookConfig_ADMIN_();
}

function debugSlackDelivery_ADMIN() {
  assertEditorAdminExecution_();
  return debugSlackDelivery_ADMIN_();
}

function smokeTestSlackDelivery_ADMIN() {
  assertEditorAdminExecution_();
  return smokeTestSlackDelivery_ADMIN_();
}

function smokeTestSlackTemplate_ADMIN() {
  assertEditorAdminExecution_();
  return smokeTestSlackTemplate_ADMIN_();
}

function reinstallSlackQueueTriggers_ADMIN() {
  assertEditorAdminExecution_();
  var pw = getAdminPasswordFromSettingsOrThrow_();
  try { removeSlackQueueTriggers(pw); } catch (e) {}
  return installSlackQueueTriggers(pw);
}

function debugSlackWebhookConfig(adminPassword) {
  if (!String(adminPassword || "").trim()) {
    throw new Error("Admin password required. In editor, run debugSlackWebhookConfig_ADMIN().");
  }
  assertAdminCredential_(adminPassword);
  var props = PropertiesService.getScriptProperties();
  var raw = String(props.getProperty("SLACK_WEBHOOK_URL") || "").trim();
  var hooks = getSlackWebhooks_();
  return {
    ok: true,
    template_version: SLACK_MESSAGE_TEMPLATE_VERSION_,
    raw_present: !!raw,
    raw_format: raw ? (/^https?:\/\//i.test(raw) ? "plain_url" : "json_or_other") : "missing",
    memo_configured: !!String(hooks.memo || "").trim(),
    approve_configured: !!String(hooks.approve || "").trim(),
    same_target: !!hooks.memo && hooks.memo === hooks.approve
  };
}

function bootstrapConfigOnce_(options) {
  ensureCoreSchemaReady_();
  ensureNotificationSchemaReady_();

  var opts = options || {};
  var force = !!opts.force;
  var props = PropertiesService.getScriptProperties();

  var webhookRaw = props.getProperty("SLACK_WEBHOOK_URL");
  var memberRaw = props.getProperty("SLACK_MEMBER_MAP");

  var webhookSet = false;
  var memberSet = false;

  if (force || !String(webhookRaw || "").trim()) {
    props.setProperty("SLACK_WEBHOOK_URL", JSON.stringify({
      memo: "",
      approve: ""
    }));
    webhookSet = true;
  }

  if (force || !String(memberRaw || "").trim()) {
    props.setProperty("SLACK_MEMBER_MAP", JSON.stringify({
      "default_assignee": "U0AG70EJ492"
    }));
    memberSet = true;
  }

  // NOTE: delete this bootstrap function after first run if you no longer want secrets in code.
  return {
    ok: true,
    forced: force,
    slackWebhookConfigured: !!props.getProperty("SLACK_WEBHOOK_URL"),
    slackMemberMapConfigured: !!props.getProperty("SLACK_MEMBER_MAP"),
    webhookSet: webhookSet,
    memberSet: memberSet,
    sheets: {
      SlackUsers: !!getSpreadsheet_().getSheetByName("SlackUsers"),
      NotificationDedup: !!getSpreadsheet_().getSheetByName("NotificationDedup"),
      SlackQueue: !!getSpreadsheet_().getSheetByName("SlackQueue")
    }
  };
}

function escapeHtml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

function include(filename) {
  var name = String(filename || "").trim();
  if (!name) return "";
  try {
    var content = HtmlService.createHtmlOutputFromFile(name).getContent();

    // Safety guard: avoid rendering raw source text when an include file is wrong/misaligned.
    if (name === "styles") {
      if (content.indexOf("<style") < 0 && content.indexOf("</style>") < 0) return "";
    }
    if (name === "ui") {
      if (content.indexOf("<script") < 0 && content.indexOf("</script>") < 0) return "";
    }
    return content;
  } catch (err) {
    // Optional partial: avoid hard-fail on stale deployments missing ui.html.
    if (name === "ui") return "";
    if (name === "styles") return "";
    throw err;
  }
}

function debugIncludeHealth_ADMIN() {
  assertEditorAdminExecution_();
  var out = {};
  var names = ["styles", "ui", "edit", "dashboard", "catalog", "templates", "templateslist", "materialgroups"];
  for (var i = 0; i < names.length; i++) {
    var n = names[i];
    try {
      var c = HtmlService.createHtmlOutputFromFile(n).getContent();
      out[n] = {
        ok: true,
        length: c.length,
        has_style_tag: c.indexOf("<style") >= 0,
        has_script_tag: c.indexOf("<script") >= 0,
        starts_with: c.slice(0, 80)
      };
    } catch (e) {
      out[n] = { ok: false, error: String((e && e.message) ? e.message : e) };
    }
  }
  return out;
}

function getPublicAppConfig() {
  ensureCoreSchemaReady_();
  var cacheKey = "PUBLIC_APP_CFG_V2";
  var out = getOrBuildCachedJsonWithLock_(
    cacheKey,
    300,
    function() {
      var s = getSettings_();
      return {
        vat_rate: Number(s.vat_rate || 0.1),
        default_vat_included: String(s.default_vat_included || "N"),
        default_include_service_in_total: String(s.default_include_service_in_total || "N"),
        default_include_option_in_total: String(s.default_include_option_in_total || "N"),
        default_include_included_in_total: String(s.default_include_included_in_total || "N"),
        base_url: getConfiguredBaseUrl_(s),
        versions: {
          material_search: String(getMaterialSearchVersion_()),
          template_list: String(getTemplateListVersion_()),
          quote_list: String(getQuoteListVersion_())
        }
      };
    }
  );
  return (out && typeof out === "object") ? out : {};
}

function loginAdmin(adminPassword) {
  assertAdmin_(adminPassword);
  var token = createAdminSessionToken_();
  var cacheKey = getAdminSessionCacheKey_(token);
  CacheService.getScriptCache().put(cacheKey, "1", ADMIN_SESSION_TTL_SEC_);
  return {
    ok: true,
    token: token,
    expires_in_sec: ADMIN_SESSION_TTL_SEC_
  };
}

function logoutAdmin(token) {
  var tok = String(token || "").trim();
  if (!tok) return { ok: true };
  try {
    CacheService.getScriptCache().remove(getAdminSessionCacheKey_(tok));
  } catch (e) {}
  return { ok: true };
}

function adminInitSchema(adminPassword) {
  assertAdminCredential_(adminPassword);
  forceEnsureCoreSchema_();
  return { ok: true, schema_version: CORE_SCHEMA_VERSION_ };
}

// Admin migration helper.
// After code deploy, update the existing Web App deployment (do not create a new URL)
// and set access to "Anyone, even anonymous" so `page=img` stays reachable.
function initBaseMigrations_ADMIN() {
  assertEditorAdminExecution_();
  forceEnsureCoreSchema_();

  var ss = getSpreadsheet_();
  ensureSheetColumnsInSs_(ss, MATERIAL_GROUPS_SHEET_, materialGroupHeaders_());
  var baseUrlFix = ensureBaseUrlSettingReadyInSs_(ss);

  return {
    ok: true,
    schema_version: CORE_SCHEMA_VERSION_,
    material_groups_sheet: MATERIAL_GROUPS_SHEET_,
    base_url: baseUrlFix.base_url,
    base_url_added: baseUrlFix.added,
    base_url_updated: baseUrlFix.updated
  };
}

function ensureBaseUrlSettingReadyInSs_(ss) {
  ensureSheetColumnsInSs_(ss, "Settings", ["key", "value"]);
  var sh = getSheetFromSs_(ss, "Settings");
  var lastRow = sh.getLastRow();
  var baseUrl = normalizeWebAppBaseUrl_(getAppUrl_());
  var added = false;
  var updated = false;

  if (lastRow < 2) {
    ensureRowCapacity_(sh, 2);
    sh.getRange(2, 1, 1, 2).setValues([["base_url", baseUrl]]);
    added = true;
    invalidateSettingsCache_(ss.getId());
    return { base_url: baseUrl, added: added, updated: updated };
  }

  var keys = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  var rowNo = 0;
  var currentVal = "";
  for (var i = 0; i < keys.length; i++) {
    var key = String(keys[i][0] || "").trim();
    if (key !== "base_url") continue;
    rowNo = i + 2;
    currentVal = normalizeWebAppBaseUrl_(keys[i][1]);
    break;
  }

  if (!rowNo) {
    var appendAt = sh.getLastRow() + 1;
    ensureRowCapacity_(sh, appendAt);
    sh.getRange(appendAt, 1, 1, 2).setValues([["base_url", baseUrl]]);
    added = true;
    invalidateSettingsCache_(ss.getId());
    return { base_url: baseUrl, added: added, updated: updated };
  }

  if (!currentVal && baseUrl) {
    sh.getRange(rowNo, 2, 1, 1).setValue(baseUrl);
    updated = true;
    currentVal = baseUrl;
    invalidateSettingsCache_(ss.getId());
  }
  return { base_url: currentVal || baseUrl || "", added: added, updated: updated };
}

function testAuth() {
  ensureCoreSchemaReady_();
  getSpreadsheet_().getId();
  DriveApp.getRootFolder().getId();
  return "ok";
}

function createQuote(adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var s = getSettings_();
  var quoteId = uuid_();
  var sh = getSheet_("Quotes");
  var headers = getHeaders_("Quotes");
  var cols = getColMap_("Quotes");
  var created = nowIso_();
  var row = Array(headers.length).fill("");

  row[cols.quote_id] = quoteId;
  row[cols.created_at] = created;
  row[cols.updated_at] = created;
  row[cols.customer_name] = "";
  row[cols.site_name] = "";
  row[cols.contact_name] = "";
  row[cols.contact_phone] = "";
  row[cols.memo] = "";
  row[cols.vat_included] = String(s.default_vat_included || "N");
  row[cols.include_service_in_total] = String(s.default_include_service_in_total || "N");
  row[cols.include_option_in_total] = String(s.default_include_option_in_total || "N");
  row[cols.include_included_in_total] = String(s.default_include_included_in_total || "N");
  row[cols.subtotal] = 0;
  row[cols.vat] = 0;
  row[cols.total] = 0;
  row[cols.share_token] = "";
  row[cols.status] = "draft";
  row[cols.shared_at] = "";
  row[cols.expire_at] = "";
  row[cols.view_count] = 0;
  row[cols.last_viewed_at] = "";
  row[cols.next_action_at] = "";
  row[cols.last_followup_sent_at] = "";
  row[cols.approved_at] = "";
  row[cols.deposit_amount] = 0;
  row[cols.deposit_due_at] = "";
  row[cols.last_customer_note_at] = "";
  row[cols.customer_note_latest] = "";
  row[cols.owner_last_seen_note_at] = "";

  var startRow = sh.getLastRow() + 1;
  ensureRowCapacity_(sh, startRow);
  sh.getRange(startRow, 1, 1, headers.length).setValues([row]);
  invalidateQuoteRowMap_();
  bumpQuoteListVersion_();

  return { quoteId: quoteId };
}

function buildGetQuoteCacheKey_(quoteId, token, quoteUpdatedAt, options, noteLimit) {
  var opts = options || {};
  var includePhotos = !!(opts.include_photos || opts.includePhotos);
  var includeNotes = !!(opts.include_notes || opts.includeNotes);
  var maxNotes = Number(noteLimit || 0);
  var flags = [
    includePhotos ? "P1" : "P0",
    includeNotes ? "N1" : "N0",
    includeNotes ? ("L" + maxNotes) : "L0"
  ].join("_");
  var tokenHash = sha256Hex_(String(token || "")).slice(0, 16);
  return "GETQ_V2_" + String(quoteId || "") + "_" + tokenHash + "_" + String(quoteUpdatedAt || "") + "_" + flags;
}

function getQuote(quoteId, token, options) {
  ensureCoreSchemaReady_();
  var opts = options || {};

  var q = findQuoteRow_(quoteId);
  if (!q) throw new Error("Quote not found");

  if (!q.share_token || String(q.share_token) !== String(token)) {
    throw new Error("Invalid link token");
  }

  var expMs = parseIsoMs_(q.expire_at);
  if (!isNaN(expMs) && expMs < Date.now()) {
    throw new Error("This quote link has expired");
  }

  var includeNotes = !!(opts.include_notes || opts.includeNotes);
  var includePhotos = !!(opts.include_photos || opts.includePhotos);
  var noteLimit = Math.min(Math.max(Number(opts.note_limit || opts.noteLimit || 100), 1), 500);
  var cacheKey = buildGetQuoteCacheKey_(quoteId, token, q.updated_at, opts, noteLimit);
  var cached = getCachedJson_(cacheKey);
  if (cached && cached.quote && Array.isArray(cached.items)) {
    return cached;
  }

  var items = getItemsFromCache_(quoteId);
  if (!items) {
    items = findItems_(quoteId);
    upsertQuoteItemsCache_(quoteId, items);
  }
  items = normalizeItems_(items);

  var imageIds = [];
  for (var i = 0; i < items.length; i++) {
    var imgId = String(items[i] && items[i].material_image_id || "").trim();
    if (imgId) imageIds.push(imgId);
  }
  markImageRefsCached_(imageIds, 3600);

  var photos = [];
  if (includePhotos) {
    var photosCacheKey = "QPHOTO_" + String(quoteId || "").trim();
    var cachedPhotos = getCachedJson_(photosCacheKey);
    if (cachedPhotos && Array.isArray(cachedPhotos)) {
      photos = cachedPhotos;
    } else {
      photos = findPhotos_(quoteId);
      putCachedJson_(photosCacheKey, photos, 180);
    }
  }
  var notes = [];
  if (includeNotes) {
    notes = listCustomerNotesByQuoteId_(quoteId, noteLimit);
  }

  var s = getSettings_();
  q.include_service_in_total = normalizeTotalFlag_(q.include_service_in_total, "N");
  q.include_option_in_total = normalizeTotalFlag_(q.include_option_in_total, "N");
  q.include_included_in_total = normalizeTotalFlag_(q.include_included_in_total, "N");

  var out = {
    quote: q,
    items: items,
    photos: photos,
    notes: notes,
    meta: {
      base_url: getConfiguredBaseUrl_(s),
      token: String(token || ""),
      account_info: String(s.account_info || ""),
      total_options: getQuoteTotalOptions_(q)
    }
  };
  putCachedJson_(cacheKey, out, includeNotes ? 45 : 180);
  return out;
}

function saveQuote(payload, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var quoteId = String(payload && payload.quoteId || "").trim();
  if (!quoteId) throw new Error("quoteId required");

  var itemsRaw = (payload && payload.items) || [];
  var items = normalizeItems_(itemsRaw);

  var quotePatch = (payload && payload.quote) ? payload.quote : {};
  var qExisting = findQuoteRow_(quoteId);
  if (!qExisting) throw new Error("Quote not found");

  var vatIncluded = String(
    (quotePatch && quotePatch.vat_included !== undefined)
      ? quotePatch.vat_included
      : (qExisting.vat_included || "N")
  );
  var includeServiceInTotal = normalizeTotalFlag_(
    (quotePatch && quotePatch.include_service_in_total),
    qExisting.include_service_in_total || "N"
  );
  var includeOptionInTotal = normalizeTotalFlag_(
    (quotePatch && quotePatch.include_option_in_total),
    qExisting.include_option_in_total || "N"
  );
  var includeIncludedInTotal = normalizeTotalFlag_(
    (quotePatch && quotePatch.include_included_in_total),
    qExisting.include_included_in_total || "N"
  );

  var vatRate = Number(getSettings_().vat_rate || 0.1);
  var totals = calculateTotalsFromItems_(items, vatIncluded, vatRate, {
    include_service_in_total: includeServiceInTotal,
    include_option_in_total: includeOptionInTotal,
    include_included_in_total: includeIncludedInTotal
  });

  var mergedPatch = cloneObj_(quotePatch);
  mergedPatch.vat_included = normUpperYN_(vatIncluded);
  mergedPatch.include_service_in_total = includeServiceInTotal;
  mergedPatch.include_option_in_total = includeOptionInTotal;
  mergedPatch.include_included_in_total = includeIncludedInTotal;
  mergedPatch.subtotal = totals.subtotal;
  mergedPatch.vat = totals.vat;
  mergedPatch.total = totals.total;

  updateQuote_(quoteId, mergedPatch);
  replaceItems_(quoteId, items);
  upsertQuoteItemsCache_(quoteId, items);

  var imageIds = [];
  for (var i = 0; i < items.length; i++) {
    var id = String(items[i].material_image_id || "").trim();
    if (id) imageIds.push(id);
  }
  markImageRefsCached_(imageIds, 3600);
  bumpQuoteListVersion_();

  return { ok: true };
}

function getQuoteAdminDetail(quoteId, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  var target = String(quoteId || "").trim();
  if (!target) throw new Error("quoteId required");
  var quote = findQuoteRow_(target);
  if (!quote) throw new Error("Quote not found");
  var items = getItemsFromCache_(target);
  if (!items) {
    items = findItems_(target);
    upsertQuoteItemsCache_(target, items);
  }
  return {
    quote: quote,
    items: normalizeItems_(items || []),
    prequote_context: {
      source_request_id: String(quote.source_request_id || "").trim(),
      source_review_id: String(quote.source_review_id || "").trim(),
      source_app: String(quote.source_app || "").trim(),
      imported_from_prequote: String(quote.source_app || "").trim().toUpperCase() === "PREQUOTE_APP" || !!String(quote.source_request_id || "").trim(),
      draft_stage: String(quote.draft_stage || "").trim(),
      summary: safeJsonParse_(quote.prequote_summary_json, {}),
      recommended_materials: safeJsonParse_(quote.recommended_materials_snapshot_json, [])
    }
  };
}

function setupTemplatesSchema(adminPassword) {
  assertAdminCredential_(adminPassword);
  return setupTemplatesSchema_();
}

function setupTemplatesSchema_() {
  ensureCoreSchemaReady_();
  return withScriptLock_(15000, function() {
    ensureTemplatesSchemaReady_();
    var migration = migrateLegacyTemplatesIfSafe_();
    var shCatalog = getSheet_(TEMPLATE_CATALOG_SHEET_);
    var shVersions = getSheet_(TEMPLATE_VERSIONS_SHEET_);
    return {
      ok: true,
      catalog_rows: Math.max(shCatalog.getLastRow() - 1, 0),
      version_rows: Math.max(shVersions.getLastRow() - 1, 0),
      migrated_legacy_rows: Number(migration.migrated_rows || 0),
      migration: migration
    };
  });
}

function ensureTemplatesSchemaReady_() {
  ensureTemplateCatalogSchema_();
  ensureTemplateVersionsSchema_();
  ensureSheetColumns_(TEMPLATE_ITEMS_CACHE_SHEET_, [
    "template_id",
    "version",
    "items_json",
    "updated_at"
  ]);
}

function ensureTemplateCatalogSchema_() {
  ensureSheetColumns_(TEMPLATE_CATALOG_SHEET_, TEMPLATE_CATALOG_HEADERS_);
}

function ensureTemplateVersionsSchema_() {
  ensureSheetColumns_(TEMPLATE_VERSIONS_SHEET_, TEMPLATE_VERSIONS_HEADERS_);
}

function migrateLegacyTemplatesIfSafe_() {
  var shCatalog = getSheet_(TEMPLATE_CATALOG_SHEET_);
  var shVersions = getSheet_(TEMPLATE_VERSIONS_SHEET_);

  if (shCatalog.getLastRow() > 1 || shVersions.getLastRow() > 1) {
    return { migrated_rows: 0, skipped_reason: "target_not_empty" };
  }

  var ss = getSpreadsheet_();
  var shLegacy = ss.getSheetByName(TEMPLATE_LEGACY_SHEET_);
  if (!shLegacy) return { migrated_rows: 0, skipped_reason: "legacy_sheet_missing" };

  var legacyLastRow = shLegacy.getLastRow();
  if (legacyLastRow < 2) return { migrated_rows: 0, skipped_reason: "legacy_empty" };

  var legacyLastCol = Math.max(shLegacy.getLastColumn(), 1);
  var legacyHeaders = shLegacy.getRange(1, 1, 1, legacyLastCol).getValues()[0].map(function(h) {
    return String(h || "").trim().toLowerCase();
  });
  var idx = Object.create(null);
  for (var i = 0; i < legacyHeaders.length; i++) idx[legacyHeaders[i]] = i;

  var hasAnyLegacyCols = (
    Object.prototype.hasOwnProperty.call(idx, "process") ||
    Object.prototype.hasOwnProperty.call(idx, "material") ||
    Object.prototype.hasOwnProperty.call(idx, "spec") ||
    Object.prototype.hasOwnProperty.call(idx, "unit") ||
    Object.prototype.hasOwnProperty.call(idx, "unit_price")
  );
  if (!hasAnyLegacyCols) return { migrated_rows: 0, skipped_reason: "legacy_columns_missing" };

  var legacyValues = shLegacy.getRange(2, 1, legacyLastRow - 1, legacyLastCol).getValues();
  var migratedItems = [];
  for (var r = 0; r < legacyValues.length; r++) {
    var row = legacyValues[r];
    var process = Object.prototype.hasOwnProperty.call(idx, "process") ? String(row[idx.process] || "") : "";
    var material = Object.prototype.hasOwnProperty.call(idx, "material") ? String(row[idx.material] || "") : "";
    var spec = Object.prototype.hasOwnProperty.call(idx, "spec") ? String(row[idx.spec] || "") : "";
    var unit = Object.prototype.hasOwnProperty.call(idx, "unit") ? String(row[idx.unit] || "") : "";
    var unitPriceRaw = Object.prototype.hasOwnProperty.call(idx, "unit_price") ? row[idx.unit_price] : 0;
    var unitPrice = Number(unitPriceRaw || 0);
    if (isNaN(unitPrice)) unitPrice = 0;
    if (!process.trim() && !material.trim() && !spec.trim() && !unit.trim() && !unitPrice) continue;

    migratedItems.push({
      process: process,
      material: material,
      spec: spec,
      qty: 1,
      unit: unit,
      unit_price: unitPrice,
      amount: unitPrice,
      note: ""
    });
  }

  migratedItems = normalizeTemplateItems_(migratedItems);
  if (!migratedItems.length) return { migrated_rows: 0, skipped_reason: "legacy_rows_empty" };

  var nowIso = nowIso_();
  var templateId = uuid_();
  var summary = buildTemplateSummary_(migratedItems);

  var headersC = getHeaders_(TEMPLATE_CATALOG_SHEET_);
  var colsC = getColMap_(TEMPLATE_CATALOG_SHEET_);
  var rowC = Array(headersC.length).fill("");
  setByHeader_(headersC, rowC, "template_id", templateId);
  setByHeader_(headersC, rowC, "category", "Legacy");
  setByHeader_(headersC, rowC, "template_name", "Migrated Legacy Templates");
  setByHeader_(headersC, rowC, "latest_version", 1);
  setByHeader_(headersC, rowC, "latest_item_count", migratedItems.length);
  setByHeader_(headersC, rowC, "created_at", nowIso);
  setByHeader_(headersC, rowC, "updated_at", nowIso);
  setByHeader_(headersC, rowC, "is_active", "Y");
  setByHeader_(headersC, rowC, "note", "Migrated from Templates sheet");
  setByHeader_(headersC, rowC, "template_type", PREQUOTE_TEMPLATE_TYPES_.QUOTE_ONLY);
  setByHeader_(headersC, rowC, "housing_type", "ALL");
  setByHeader_(headersC, rowC, "area_band", "ALL");
  setByHeader_(headersC, rowC, "expose_to_prequote", "N");
  setByHeader_(headersC, rowC, "prequote_priority", 0);
  setByHeader_(headersC, rowC, "sort_order", 0);
  setByHeader_(headersC, rowC, "budget_band", "ANY");
  setByHeader_(headersC, rowC, "is_featured_prequote", "N");
  var cStart = shCatalog.getLastRow() + 1;
  ensureRowCapacity_(shCatalog, cStart);
  shCatalog.getRange(cStart, 1, 1, headersC.length).setValues([rowC]);

  var headersV = getHeaders_(TEMPLATE_VERSIONS_SHEET_);
  var rowV = Array(headersV.length).fill("");
  var migratedMeta = pickTemplateMeta_(rowC);
  setByHeader_(headersV, rowV, "template_id", templateId);
  setByHeader_(headersV, rowV, "version", 1);
  setByHeader_(headersV, rowV, "created_at", nowIso);
  setByHeader_(headersV, rowV, "created_by", "migration");
  setByHeader_(headersV, rowV, "source_quote_id", "");
  setByHeader_(headersV, rowV, "items_json", JSON.stringify(migratedItems));
  setByHeader_(headersV, rowV, "summary_json", JSON.stringify(summary));
  setByHeader_(headersV, rowV, "item_count", migratedItems.length);
  setByHeader_(headersV, rowV, "metadata_snapshot_json", JSON.stringify(migratedMeta));
  var vStart = shVersions.getLastRow() + 1;
  ensureRowCapacity_(shVersions, vStart);
  shVersions.getRange(vStart, 1, 1, headersV.length).setValues([rowV]);

  bumpTemplateListVersion_();
  return {
    migrated_rows: migratedItems.length,
    template_id: templateId,
    version: 1,
    skipped_reason: ""
  };
}

function templateBoolToCell_(flag) {
  return flag ? "Y" : "N";
}

function templateCellToBool_(raw) {
  var s = String(raw || "").trim().toUpperCase();
  if (!s) return true;
  if (s === "N" || s === "FALSE" || s === "0" || s === "INACTIVE") return false;
  return true;
}

function getTemplateListVersion_() {
  // Cache key contract:
  // - TEMPLATE_LIST_VER: invalidates template snapshot/list/category caches.
  var props = PropertiesService.getScriptProperties();
  var v = String(props.getProperty(TEMPLATE_LIST_VERSION_KEY_) || "").trim();
  if (!v) {
    v = "1";
    props.setProperty(TEMPLATE_LIST_VERSION_KEY_, v);
  }
  return v;
}

function bumpTemplateListVersion_() {
  PropertiesService.getScriptProperties().setProperty(TEMPLATE_LIST_VERSION_KEY_, String(Date.now()));
  __TEMPLATE_CATALOG_SNAPSHOT_MEM_VER_ = "";
  __TEMPLATE_CATALOG_SNAPSHOT_MEM_ = null;
  __TEMPLATE_CATALOG_CORE_MEM_VER_ = "";
  __TEMPLATE_CATALOG_CORE_MEM_ = null;
}

function templateCatalogListCacheKey_(category, query, includeInactive, filtersOpt) {
  var raw = [
    getTemplateListVersion_(),
    String(category || "").trim().toLowerCase(),
    String(query || "").trim().toLowerCase(),
    includeInactive ? "1" : "0",
    serializeTemplateAdminFilters_(filtersOpt)
  ].join("|");
  return "TPL_LIST_" + sha256Hex_(raw).slice(0, 24);
}

function templateCategoryListCacheKey_(includeInactive) {
  var raw = [getTemplateListVersion_(), includeInactive ? "1" : "0"].join("|");
  return "TPL_CAT_" + sha256Hex_(raw).slice(0, 24);
}

function templateCatalogSnapshotCacheKey_() {
  return "TPL_SNAP_" + String(getTemplateListVersion_());
}

function templateCatalogCoreSnapshotCacheKey_() {
  return "TPL_CORE_SNAP_" + String(getTemplateListVersion_());
}

function getDefaultTemplateMeta_() {
  return {
    template_type: TEMPLATE_META_DEFAULTS_.template_type,
    housing_type: TEMPLATE_META_DEFAULTS_.housing_type,
    area_band: TEMPLATE_META_DEFAULTS_.area_band,
    style_tags_summary: TEMPLATE_META_DEFAULTS_.style_tags_summary,
    tone_tags_summary: TEMPLATE_META_DEFAULTS_.tone_tags_summary,
    trade_scope_summary: TEMPLATE_META_DEFAULTS_.trade_scope_summary,
    expose_to_prequote: !!TEMPLATE_META_DEFAULTS_.expose_to_prequote,
    prequote_priority: Number(TEMPLATE_META_DEFAULTS_.prequote_priority || 0),
    sort_order: Number(TEMPLATE_META_DEFAULTS_.sort_order || 0),
    recommendation_note: TEMPLATE_META_DEFAULTS_.recommendation_note,
    target_customer_summary: TEMPLATE_META_DEFAULTS_.target_customer_summary,
    budget_band: TEMPLATE_META_DEFAULTS_.budget_band,
    recommended_for_summary: TEMPLATE_META_DEFAULTS_.recommended_for_summary,
    is_featured_prequote: !!TEMPLATE_META_DEFAULTS_.is_featured_prequote
  };
}

function normalizeTemplateTagSummary_(value) {
  return normalizeStringList_(value, normalizeTagValue_).join("|");
}

function normalizeTemplateTextSummary_(value) {
  return normalizeSpaces_(value);
}

function normalizeHousingType_(value) {
  var s = String(value || "").trim().toUpperCase();
  if (s === "APT") s = "APARTMENT";
  if (s === "APTMENT") s = "APARTMENT";
  return TEMPLATE_HOUSING_TYPE_OPTIONS_.indexOf(s) >= 0 ? s : "ALL";
}

function normalizeAreaBand_(value) {
  var s = String(value || "").trim().toUpperCase().replace(/\s+/g, "");
  if (!s) return "ALL";
  if (s === "20PY_30PY" || s === "20PY-30PY" || s === "20TO29") s = "20_29";
  if (s === "30PY_40PY" || s === "30PY-40PY" || s === "30TO39") s = "30_39";
  if (s === "40PY_50PY" || s === "40PY-50PY" || s === "40TO49") s = "40_49";
  if (s === "50PY_PLUS" || s === "50_OVER" || s === "OVER_50") s = "50_PLUS";
  return TEMPLATE_AREA_BAND_OPTIONS_.indexOf(s) >= 0 ? s : "ALL";
}

function normalizeBudgetBand_(value) {
  var s = String(value || "").trim().toUpperCase();
  if (s === "LOW") s = "ENTRY";
  if (s === "STANDARD") s = "MID";
  return TEMPLATE_BUDGET_BAND_OPTIONS_.indexOf(s) >= 0 ? s : "ANY";
}

function normalizeTemplatePriority_(value) {
  var n = Math.floor(normalizeNumber_(value, 0));
  return n > 0 ? n : 0;
}

function pickTemplateMeta_(src) {
  var row = src || {};
  return {
    template_type: normalizeTemplateType_(row.template_type),
    housing_type: normalizeHousingType_(row.housing_type),
    area_band: normalizeAreaBand_(row.area_band),
    style_tags_summary: normalizeTemplateTagSummary_(row.style_tags_summary),
    tone_tags_summary: normalizeTemplateTagSummary_(row.tone_tags_summary),
    trade_scope_summary: normalizeTemplateTagSummary_(row.trade_scope_summary),
    expose_to_prequote: ynToBool_(row.expose_to_prequote, false),
    prequote_priority: normalizeTemplatePriority_(row.prequote_priority),
    sort_order: normalizeTemplateSortOrder_(row.sort_order),
    recommendation_note: normalizeTemplateTextSummary_(row.recommendation_note),
    target_customer_summary: normalizeTemplateTextSummary_(row.target_customer_summary),
    budget_band: normalizeBudgetBand_(row.budget_band),
    recommended_for_summary: normalizeTemplateTextSummary_(row.recommended_for_summary),
    is_featured_prequote: ynToBool_(row.is_featured_prequote, false)
  };
}

function mergeTemplateMeta_(base, patch) {
  var out = pickTemplateMeta_(base || getDefaultTemplateMeta_());
  var raw = patch || {};
  var hasOwn = Object.prototype.hasOwnProperty;

  if (hasOwn.call(raw, "template_type")) out.template_type = normalizeTemplateType_(raw.template_type);
  if (hasOwn.call(raw, "housing_type")) out.housing_type = normalizeHousingType_(raw.housing_type);
  if (hasOwn.call(raw, "area_band")) out.area_band = normalizeAreaBand_(raw.area_band);
  if (hasOwn.call(raw, "style_tags_summary")) out.style_tags_summary = normalizeTemplateTagSummary_(raw.style_tags_summary);
  if (hasOwn.call(raw, "tone_tags_summary")) out.tone_tags_summary = normalizeTemplateTagSummary_(raw.tone_tags_summary);
  if (hasOwn.call(raw, "trade_scope_summary")) out.trade_scope_summary = normalizeTemplateTagSummary_(raw.trade_scope_summary);
  if (hasOwn.call(raw, "expose_to_prequote")) out.expose_to_prequote = ynToBool_(raw.expose_to_prequote, false);
  if (hasOwn.call(raw, "prequote_priority")) out.prequote_priority = normalizeTemplatePriority_(raw.prequote_priority);
  if (hasOwn.call(raw, "sort_order")) out.sort_order = normalizeTemplateSortOrder_(raw.sort_order);
  if (hasOwn.call(raw, "recommendation_note")) out.recommendation_note = normalizeTemplateTextSummary_(raw.recommendation_note);
  if (hasOwn.call(raw, "target_customer_summary")) out.target_customer_summary = normalizeTemplateTextSummary_(raw.target_customer_summary);
  if (hasOwn.call(raw, "budget_band")) out.budget_band = normalizeBudgetBand_(raw.budget_band);
  if (hasOwn.call(raw, "recommended_for_summary")) out.recommended_for_summary = normalizeTemplateTextSummary_(raw.recommended_for_summary);
  if (hasOwn.call(raw, "is_featured_prequote")) out.is_featured_prequote = ynToBool_(raw.is_featured_prequote, false);

  return out;
}

function extractTemplateMetaPatch_(payload) {
  var data = payload || {};
  var metaSrc = (data.template_meta && typeof data.template_meta === "object") ? data.template_meta : {};
  var out = {};
  var hasOwn = Object.prototype.hasOwnProperty;
  for (var i = 0; i < TEMPLATE_META_FIELD_NAMES_.length; i++) {
    var field = TEMPLATE_META_FIELD_NAMES_[i];
    if (hasOwn.call(metaSrc, field)) {
      out[field] = metaSrc[field];
    } else if (hasOwn.call(data, field)) {
      out[field] = data[field];
    }
  }
  return out;
}

function normalizeTemplateAdminFilters_(filtersOpt) {
  var raw = filtersOpt || {};
  var expose = String(raw.expose_to_prequote || "").trim().toUpperCase();
  return {
    template_type: raw.template_type !== undefined && raw.template_type !== null && String(raw.template_type).trim() !== ""
      ? normalizeTemplateType_(raw.template_type)
      : "",
    expose_to_prequote: expose === "Y" || expose === "N" ? expose : "",
    housing_type: raw.housing_type !== undefined && raw.housing_type !== null && String(raw.housing_type).trim() !== ""
      ? normalizeHousingType_(raw.housing_type)
      : "",
    area_band: raw.area_band !== undefined && raw.area_band !== null && String(raw.area_band).trim() !== ""
      ? normalizeAreaBand_(raw.area_band)
      : ""
  };
}

function serializeTemplateAdminFilters_(filtersOpt) {
  var f = normalizeTemplateAdminFilters_(filtersOpt);
  return [
    String(f.template_type || ""),
    String(f.expose_to_prequote || ""),
    String(f.housing_type || ""),
    String(f.area_band || "")
  ].join("|");
}

function cloneTemplateCatalogRow_(row) {
  var src = row || {};
  var meta = pickTemplateMeta_(src);
  return {
    _rowNo: Number(src._rowNo || 0),
    template_id: String(src.template_id || "").trim(),
    category: String(src.category || "").trim() || "General",
    template_name: String(src.template_name || "").trim(),
    latest_version: Number(src.latest_version || 0),
    latest_item_count: Number(src.latest_item_count || 0),
    created_at: String(src.created_at || ""),
    updated_at: String(src.updated_at || ""),
    is_active: !!src.is_active,
    note: String(src.note || ""),
    template_type: meta.template_type,
    housing_type: meta.housing_type,
    area_band: meta.area_band,
    style_tags_summary: meta.style_tags_summary,
    tone_tags_summary: meta.tone_tags_summary,
    trade_scope_summary: meta.trade_scope_summary,
    expose_to_prequote: !!meta.expose_to_prequote,
    prequote_priority: meta.prequote_priority,
    sort_order: meta.sort_order,
    recommendation_note: meta.recommendation_note,
    target_customer_summary: meta.target_customer_summary,
    budget_band: meta.budget_band,
    recommended_for_summary: meta.recommended_for_summary,
    is_featured_prequote: !!meta.is_featured_prequote,
    template_meta: cloneObj_(meta),
    _search_text: String(src._search_text || "")
  };
}

function buildTemplateCatalogCoreSnapshot_() {
  var sh = getSheet_(TEMPLATE_CATALOG_SHEET_);
  var headers = getHeaders_(TEMPLATE_CATALOG_SHEET_);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var rowObj = pickTemplateCatalog_(rowToObj_(headers, values[i], i + 2));
    if (!rowObj.template_id) continue;
    out.push(rowObj);
  }
  return out;
}

function getTemplateCatalogCoreSnapshot_() {
  var ver = getTemplateListVersion_();
  if (__TEMPLATE_CATALOG_CORE_MEM_ && __TEMPLATE_CATALOG_CORE_MEM_VER_ === ver) {
    return __TEMPLATE_CATALOG_CORE_MEM_;
  }
  var cacheKey = templateCatalogCoreSnapshotCacheKey_();
  var snapshot = getOrBuildCachedJsonChunkedWithLock_(
    cacheKey,
    900,
    function() { return buildTemplateCatalogCoreSnapshot_(); },
    { wait_ms: 2200, retry_ms: 100 }
  );
  var list = Array.isArray(snapshot) ? snapshot : [];
  __TEMPLATE_CATALOG_CORE_MEM_VER_ = ver;
  __TEMPLATE_CATALOG_CORE_MEM_ = list;
  return list;
}

function buildTemplateCatalogSnapshot_() {
  var sh = getSheet_(TEMPLATE_CATALOG_SHEET_);
  var headers = getHeaders_(TEMPLATE_CATALOG_SHEET_);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var out = [];
  var versionsByTemplateId = Object.create(null);

  for (var i = 0; i < values.length; i++) {
    var rowObj = pickTemplateCatalog_(rowToObj_(headers, values[i], i + 2));
    if (!rowObj.template_id) continue;
    out.push(rowObj);
    versionsByTemplateId[rowObj.template_id] = Number(rowObj.latest_version || 0);
  }

  var summaries = loadTemplateSummaryMapByVersion_(versionsByTemplateId);
  for (var j = 0; j < out.length; j++) {
    var tpl = out[j];
    tpl._search_text = buildTemplateSearchText_(tpl, summaries[tpl.template_id] || {});
  }
  return out;
}

function getTemplateCatalogSnapshot_() {
  var ver = getTemplateListVersion_();
  if (__TEMPLATE_CATALOG_SNAPSHOT_MEM_ && __TEMPLATE_CATALOG_SNAPSHOT_MEM_VER_ === ver) {
    return __TEMPLATE_CATALOG_SNAPSHOT_MEM_;
  }

  var cacheKey = templateCatalogSnapshotCacheKey_();
  var snapshot = getOrBuildCachedJsonChunkedWithLock_(
    cacheKey,
    900,
    function() { return buildTemplateCatalogSnapshot_(); },
    { wait_ms: 2500, retry_ms: 120 }
  );
  var list = Array.isArray(snapshot) ? snapshot : [];
  __TEMPLATE_CATALOG_SNAPSHOT_MEM_VER_ = ver;
  __TEMPLATE_CATALOG_SNAPSHOT_MEM_ = list;
  return list;
}

function splitTemplateQueryTokens_(query) {
  var parts = String(query || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim()
    .split(" ");
  var out = [];
  var seen = Object.create(null);
  for (var i = 0; i < parts.length; i++) {
    var t = String(parts[i] || "").trim();
    if (!t || seen[t]) continue;
    seen[t] = 1;
    out.push(t);
  }
  return out;
}

function isTemplateItemEmpty_(item) {
  var it = item || {};
  var txt = [
    String(it.group_label || it.groupLabel || "").trim(),
    String(it.group_code || it.groupCode || "").trim(),
    String(it.name || "").trim(),
    String(it.location || "").trim(),
    String(it.detail || "").trim(),
    String(it.process || "").trim(),
    String(it.material || "").trim(),
    String(it.spec || "").trim(),
    String(it.unit || "").trim(),
    String(it.note || "").trim()
  ].join("");
  if (txt) return false;
  if (Number(it.qty || 0)) return false;
  if (Number(it.unit_price || 0)) return false;
  if (Number(it.amount || 0)) return false;
  return true;
}

function normalizeTemplateItems_(items) {
  var list = Array.isArray(items) ? items : [];
  var out = [];
  var groupOrderMap = Object.create(null);
  var nextGroupOrder = 1;
  for (var i = 0; i < list.length; i++) {
    var it = list[i] || {};
    var qty = Number(it.qty || 0);
    var unitPrice = Number(it.unit_price || 0);
    var amount = (it.amount !== undefined && it.amount !== null && String(it.amount) !== "")
      ? Number(it.amount || 0)
      : (qty * unitPrice);
    var price = (it.price !== undefined && it.price !== null && String(it.price) !== "")
      ? Number(it.price || 0)
      : amount;
    if (isNaN(qty)) qty = 0;
    if (isNaN(unitPrice)) unitPrice = 0;
    if (isNaN(amount)) amount = qty * unitPrice;
    if (isNaN(price)) price = amount;

    var groupLabel = String(it.group_label || it.groupLabel || it.process || "").trim();
    if (!groupLabel) groupLabel = "기타";
    var groupId = String(it.group_id || it.groupId || "").trim();
    if (!groupId) groupId = slugifyGroupId_(groupLabel);
    var groupCode = String(it.group_code || it.groupCode || "").trim();
    var groupOrder = Number(it.group_order || it.groupOrder || 0);
    if (isNaN(groupOrder) || groupOrder <= 0) {
      if (!groupOrderMap[groupId]) groupOrderMap[groupId] = nextGroupOrder++;
      groupOrder = groupOrderMap[groupId];
    } else {
      groupOrderMap[groupId] = groupOrder;
      if (groupOrder >= nextGroupOrder) nextGroupOrder = groupOrder + 1;
    }
    var itemOrder = Number(it.item_order || it.itemOrder || (i + 1));
    if (isNaN(itemOrder) || itemOrder <= 0) itemOrder = i + 1;

    var normalized = {
      group_id: groupId,
      group_label: groupLabel,
      group_code: groupCode,
      group_order: groupOrder,
      item_order: itemOrder,
      name: String(it.name || it.material || ""),
      location: String(it.location || ""),
      detail: String(it.detail || it.spec || ""),
      price: price,
      price_type: normalizePriceType_(it.price_type || it.priceType),
      process: String(it.process || ""),
      material: String(it.material || ""),
      spec: String(it.spec || ""),
      qty: qty,
      unit: String(it.unit || ""),
      unit_price: unitPrice,
      amount: amount,
      note: String(it.note || ""),
      material_ref_id: String(it.material_ref_id || ""),
      material_image_id: String(it.material_image_id || ""),
      material_image_name: String(it.material_image_name || "")
    };
    if (!normalized.process) normalized.process = normalized.group_label;
    if (!normalized.material) normalized.material = normalized.name;
    if (!normalized.spec) normalized.spec = normalized.detail;
    if (!normalized.amount) normalized.amount = normalized.price;
    if (!normalized.unit_price) normalized.unit_price = normalized.price;
    if (isTemplateItemEmpty_(normalized)) continue;
    out.push(normalized);
  }
  out.sort(function(a, b) {
    var ga = Number(a.group_order || 0);
    var gb = Number(b.group_order || 0);
    if (ga !== gb) return ga - gb;
    var ia = Number(a.item_order || 0);
    var ib = Number(b.item_order || 0);
    if (ia !== ib) return ia - ib;
    return 0;
  });
  return out;
}

function buildTemplateSummary_(items) {
  var list = normalizeTemplateItems_(items);
  var subtotal = 0;
  var kwSeen = Object.create(null);
  var keywords = [];

  for (var i = 0; i < list.length; i++) {
    var it = list[i];
    subtotal += Number(it.amount || 0);
    var text = [
      it.group_label, it.group_code,
      it.name, it.location, it.detail,
      it.process, it.material, it.spec, it.note
    ].join(" ").toLowerCase();
    var tokens = text.split(/[\s,;|\/]+/);
    for (var t = 0; t < tokens.length; t++) {
      var token = String(tokens[t] || "").trim();
      if (!token || token.length < 2 || kwSeen[token]) continue;
      kwSeen[token] = 1;
      keywords.push(token);
      if (keywords.length >= 40) break;
    }
    if (keywords.length >= 40) break;
  }

  return {
    item_count: list.length,
    subtotal: subtotal,
    keywords: keywords,
    search_text: keywords.join(" ")
  };
}

function buildTemplateSearchText_(templateObj, summaryObj) {
  var tpl = templateObj || {};
  var sum = summaryObj || {};
  var meta = pickTemplateMeta_(tpl);
  var words = [
    String(tpl.category || ""),
    String(tpl.template_name || ""),
    String(tpl.note || ""),
    String(meta.template_type || ""),
    String(meta.housing_type || ""),
    String(meta.area_band || ""),
    String(meta.style_tags_summary || ""),
    String(meta.tone_tags_summary || ""),
    String(meta.trade_scope_summary || ""),
    String(meta.budget_band || ""),
    String(meta.target_customer_summary || ""),
    String(meta.recommended_for_summary || ""),
    String(meta.recommendation_note || "")
  ];
  if (Array.isArray(sum.keywords)) words.push(sum.keywords.join(" "));
  if (sum.search_text) words.push(String(sum.search_text || ""));
  return normalizeSpaces_(words.join(" ").toLowerCase());
}

function getSessionUserEmailSafe_() {
  try { return String(Session.getActiveUser().getEmail() || "").trim(); } catch (e) {}
  return "";
}

function pickTemplateCatalog_(rowObj) {
  var row = rowObj || {};
  var meta = pickTemplateMeta_(row);
  return {
    _rowNo: Number(row._rowNo || 0),
    template_id: String(row.template_id || "").trim(),
    category: String(row.category || "").trim() || "General",
    template_name: String(row.template_name || "").trim(),
    latest_version: Number(row.latest_version || 0),
    latest_item_count: Number(row.latest_item_count || 0),
    created_at: String(row.created_at || ""),
    updated_at: String(row.updated_at || ""),
    is_active: templateCellToBool_(row.is_active),
    note: String(row.note || ""),
    template_type: meta.template_type,
    housing_type: meta.housing_type,
    area_band: meta.area_band,
    style_tags_summary: meta.style_tags_summary,
    tone_tags_summary: meta.tone_tags_summary,
    trade_scope_summary: meta.trade_scope_summary,
    expose_to_prequote: !!meta.expose_to_prequote,
    prequote_priority: meta.prequote_priority,
    sort_order: meta.sort_order,
    recommendation_note: meta.recommendation_note,
    target_customer_summary: meta.target_customer_summary,
    budget_band: meta.budget_band,
    recommended_for_summary: meta.recommended_for_summary,
    is_featured_prequote: !!meta.is_featured_prequote,
    template_meta: cloneObj_(meta)
  };
}

function pickTemplateVersionMeta_(rowObj) {
  var row = rowObj || {};
  var summary = parseJsonSafe_(row.summary_json, {});
  var metaSnapshot = parseJsonSafe_(row.metadata_snapshot_json, null);
  return {
    _rowNo: Number(row._rowNo || 0),
    template_id: String(row.template_id || "").trim(),
    version: Number(row.version || 0),
    created_at: String(row.created_at || ""),
    created_by: String(row.created_by || ""),
    source_quote_id: String(row.source_quote_id || ""),
    item_count: Number(row.item_count || (summary && summary.item_count) || 0),
    summary: (summary && typeof summary === "object") ? summary : {},
    template_meta_snapshot: (metaSnapshot && typeof metaSnapshot === "object") ? pickTemplateMeta_(metaSnapshot) : null
  };
}

function readTemplateCatalogRowById_(templateId) {
  var target = String(templateId || "").trim();
  if (!target) return null;

  var sh = getSheet_(TEMPLATE_CATALOG_SHEET_);
  var headers = getHeaders_(TEMPLATE_CATALOG_SHEET_);
  var cols = getColMap_(TEMPLATE_CATALOG_SHEET_);
  if (!Object.prototype.hasOwnProperty.call(cols, "template_id")) {
    throw new Error("TemplateCatalog missing column: template_id");
  }
  var rowNo = findRowNoByColumnValue_(sh, cols.template_id + 1, target);
  if (rowNo < 2) return null;
  var row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
  if (String(row[cols.template_id] || "").trim() !== target) return null;
  return pickTemplateCatalog_(rowToObj_(headers, row, rowNo));
}

function listTemplateVersionRowsRaw_(templateId) {
  var target = String(templateId || "").trim();
  if (!target) return [];

  var sh = getSheet_(TEMPLATE_VERSIONS_SHEET_);
  var headers = getHeaders_(TEMPLATE_VERSIONS_SHEET_);
  var cols = getColMap_(TEMPLATE_VERSIONS_SHEET_);
  if (!Object.prototype.hasOwnProperty.call(cols, "template_id")) {
    throw new Error("TemplateVersions missing column: template_id");
  }

  var rowNos = findRowNosByColumnValue_(sh, cols.template_id + 1, target);
  if (!rowNos.length) return [];

  var rowsMap = readRowsMapByRowNos_(sh, rowNos, 1, headers.length);
  var out = [];
  for (var i = 0; i < rowNos.length; i++) {
    var rowNo = rowNos[i];
    var row = rowsMap[rowNo];
    if (!row) continue;
    if (String(row[cols.template_id] || "").trim() !== target) continue;
    out.push(rowToObj_(headers, row, rowNo));
  }

  out.sort(function(a, b) {
    var av = Number(a.version || 0);
    var bv = Number(b.version || 0);
    if (av !== bv) return bv - av;
    return Number(b._rowNo || 0) - Number(a._rowNo || 0);
  });
  return out;
}

function getTemplateVersionRow_(templateId, version) {
  var target = String(templateId || "").trim();
  var v = Number(version || 0);
  if (!target || v < 1) return null;

  var rows = listTemplateVersionRowsRaw_(target);
  for (var i = 0; i < rows.length; i++) {
    if (Number(rows[i].version || 0) === v) return rows[i];
  }
  return null;
}

function resolveTemplateVersionNumber_(templateId, version) {
  var target = String(templateId || "").trim();
  if (!target) throw new Error("template_id required");

  var v = Number(version || 0);
  if (v >= 1) return Math.floor(v);

  var tpl = readTemplateCatalogRowById_(target);
  if (!tpl) throw new Error("Template not found: " + target);
  var latest = Number(tpl.latest_version || 0);
  if (latest < 1) throw new Error("Template has no versions: " + target);
  return latest;
}

function loadTemplateSummaryMapByVersion_(versionByTemplateId) {
  var wanted = versionByTemplateId || {};
  var hasWanted = false;
  for (var k in wanted) { hasWanted = true; break; }
  if (!hasWanted) return {};

  var sh = getSheet_(TEMPLATE_VERSIONS_SHEET_);
  var headers = getHeaders_(TEMPLATE_VERSIONS_SHEET_);
  var cols = getColMap_(TEMPLATE_VERSIONS_SHEET_);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var out = Object.create(null);
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var tid = String(row[cols.template_id] || "").trim();
    if (!tid || !Object.prototype.hasOwnProperty.call(wanted, tid)) continue;
    var expectedVer = Number(wanted[tid] || 0);
    var rowVer = Number(row[cols.version] || 0);
    if (expectedVer !== rowVer) continue;
    out[tid] = parseJsonSafe_(row[cols.summary_json], {});
  }
  return out;
}

function listTemplateCategories(adminPassword, includeInactive) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();

  var includeInactives = !!includeInactive;
  var cacheKey = templateCategoryListCacheKey_(includeInactives);
  var out = getOrBuildCachedJsonWithLock_(
    cacheKey,
    300,
    function() {
      var snapshot = getTemplateCatalogCoreSnapshot_();
      var counts = Object.create(null);
      for (var i = 0; i < snapshot.length; i++) {
        var row = snapshot[i] || {};
        if (!includeInactives && !row.is_active) continue;
        var category = String(row.category || "").trim() || "General";
        counts[category] = Number(counts[category] || 0) + 1;
      }
      var list = [];
      for (var c in counts) list.push({ category: c, count: counts[c] });
      list.sort(function(a, b) {
        return String(a.category || "").localeCompare(String(b.category || ""));
      });
      return list;
    },
    { wait_ms: 1800, retry_ms: 80 }
  );
  return Array.isArray(out) ? out : [];
}

function listTemplateMetaPresetsAdmin(adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  return {
    defaults: getDefaultTemplateMeta_(),
    enums: {
      template_type: [PREQUOTE_TEMPLATE_TYPES_.QUOTE_ONLY, PREQUOTE_TEMPLATE_TYPES_.PREQUOTE_PACKAGE, PREQUOTE_TEMPLATE_TYPES_.BOTH],
      housing_type: TEMPLATE_HOUSING_TYPE_OPTIONS_.slice(),
      area_band: TEMPLATE_AREA_BAND_OPTIONS_.slice(),
      budget_band: TEMPLATE_BUDGET_BAND_OPTIONS_.slice(),
      expose_to_prequote: ["Y", "N"],
      is_featured_prequote: ["Y", "N"]
    }
  };
}

function listTemplateFilterOptionsAdmin(adminPassword) {
  return listTemplateMetaPresetsAdmin(adminPassword);
}

function listTemplateCatalog(category, query, includeInactive, adminPassword, filtersOpt) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();

  var filterCategory = String(category || "").trim();
  var q = String(query || "").trim();
  var tokens = splitTemplateQueryTokens_(q);
  var includeInactives = !!includeInactive;
  var metaFilters = normalizeTemplateAdminFilters_(filtersOpt);
  var cacheKey = templateCatalogListCacheKey_(filterCategory, q, includeInactives, metaFilters);
  var out = getOrBuildCachedJsonWithLock_(
    cacheKey,
    300,
    function() {
      var snapshot = tokens.length ? getTemplateCatalogSnapshot_() : getTemplateCatalogCoreSnapshot_();
      var rows = [];
      for (var i = 0; i < snapshot.length; i++) {
        var row = snapshot[i] || {};
        if (!row.template_id) continue;
        if (!includeInactives && !row.is_active) continue;
        if (filterCategory && String(row.category || "") !== filterCategory) continue;
        if (metaFilters.template_type && String(row.template_type || "") !== metaFilters.template_type) continue;
        if (metaFilters.expose_to_prequote && boolToYn_(!!row.expose_to_prequote) !== metaFilters.expose_to_prequote) continue;
        if (metaFilters.housing_type && String(row.housing_type || "") !== metaFilters.housing_type) continue;
        if (metaFilters.area_band && String(row.area_band || "") !== metaFilters.area_band) continue;
        if (tokens.length) {
          var searchText = String(row._search_text || "");
          var hitAll = true;
          for (var n = 0; n < tokens.length; n++) {
            if (searchText.indexOf(tokens[n]) < 0) {
              hitAll = false;
              break;
            }
          }
          if (!hitAll) continue;
        }
        rows.push(cloneTemplateCatalogRow_(row));
      }

      rows.sort(function(a, b) {
        var sa = Number(a.sort_order || 0);
        var sb = Number(b.sort_order || 0);
        if (sa !== sb) return sa - sb;
        var pa = Number(a.prequote_priority || 0);
        var pb = Number(b.prequote_priority || 0);
        if (pa !== pb) return pb - pa;
        var ca = String(a.category || "");
        var cb = String(b.category || "");
        if (ca !== cb) return ca.localeCompare(cb);
        var na = String(a.template_name || "");
        var nb = String(b.template_name || "");
        if (na !== nb) return na.localeCompare(nb);
        return String(b.updated_at || "").localeCompare(String(a.updated_at || ""));
      });

      for (var z = 0; z < rows.length; z++) delete rows[z]._search_text;
      return rows;
    },
    { wait_ms: 2000, retry_ms: 100 }
  );
  return Array.isArray(out) ? out : [];
}

function getTemplateDetail(templateId, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();

  var target = String(templateId || "").trim();
  if (!target) throw new Error("template_id required");

  var tpl = readTemplateCatalogRowById_(target);
  if (!tpl) throw new Error("Template not found: " + target);
  var templateOut = cloneTemplateCatalogRow_(tpl);

  var rawVersions = listTemplateVersionRowsRaw_(target);
  var versions = [];
  for (var i = 0; i < rawVersions.length; i++) {
    versions.push(pickTemplateVersionMeta_(rawVersions[i]));
  }

  return {
    template: templateOut,
    template_meta: cloneObj_(templateOut.template_meta || pickTemplateMeta_(templateOut)),
    versions: versions
  };
}

function templateVersionDetailCacheKey_(templateId, version) {
  var tid = String(templateId || "").trim();
  var v = Number(version || 0);
  return "TPL_VER_DETAIL_" + sha256Hex_([getTemplateListVersion_(), tid, v].join("|")).slice(0, 30);
}

function templateVersionItemsScriptCacheKey_(templateId, version) {
  var tid = String(templateId || "").trim();
  var v = Number(version || 0);
  return "TPL_ITEMS_" + sha256Hex_([getTemplateListVersion_(), tid, v].join("|")).slice(0, 30);
}

function templateVersionItemsSheetCacheKey_(templateId, version) {
  var tid = String(templateId || "").trim();
  var v = Number(version || 0);
  return "TPL_ITEMS_SHEET_" + sha256Hex_([getTemplateListVersion_(), tid, v].join("|")).slice(0, 28);
}

function findTemplateItemsCacheRowNo_(sh, cols, templateId, version) {
  var target = String(templateId || "").trim();
  var v = Number(version || 0);
  if (!target || v < 1) return 0;
  var rowNos = findRowNosByColumnValue_(sh, cols.template_id + 1, target);
  if (!rowNos.length) return 0;
  var rows = readRowsMapByRowNos_(sh, rowNos, cols.version + 1, 1);
  for (var i = 0; i < rowNos.length; i++) {
    var rn = rowNos[i];
    var row = rows[rn];
    if (!row) continue;
    if (Number(row[0] || 0) === v) return rn;
  }
  return 0;
}

function upsertTemplateVersionItemsCache_(templateId, version, items) {
  var target = String(templateId || "").trim();
  var v = Number(version || 0);
  if (!target || v < 1) return;
  var list = normalizeTemplateItems_(Array.isArray(items) ? items : []);

  try {
    CacheService.getScriptCache().put(
      templateVersionItemsScriptCacheKey_(target, v),
      JSON.stringify(list),
      21600
    );
  } catch (e0) {}

  putCachedJson_(templateVersionItemsSheetCacheKey_(target, v), list, 900);

  try {
    var sh = getSheet_(TEMPLATE_ITEMS_CACHE_SHEET_);
    var headers = getHeaders_(TEMPLATE_ITEMS_CACHE_SHEET_);
    var cols = getColMap_(TEMPLATE_ITEMS_CACHE_SHEET_);
    var rowNo = findTemplateItemsCacheRowNo_(sh, cols, target, v);
    var row = Array(headers.length).fill("");
    if (rowNo >= 2) {
      row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
    } else {
      rowNo = sh.getLastRow() + 1;
      ensureRowCapacity_(sh, rowNo);
    }
    row[cols.template_id] = target;
    row[cols.version] = v;
    row[cols.items_json] = JSON.stringify(list);
    row[cols.updated_at] = nowIso_();
    sh.getRange(rowNo, 1, 1, headers.length).setValues([row]);
  } catch (sheetErr) {}
}

function getTemplateVersionItemsFromCache_(templateId, version) {
  var target = String(templateId || "").trim();
  var v = Number(version || 0);
  if (!target || v < 1) return null;

  try {
    var raw = CacheService.getScriptCache().get(templateVersionItemsScriptCacheKey_(target, v));
    if (raw) {
      var parsed0 = parseJsonSafe_(raw, null);
      if (Array.isArray(parsed0)) return parsed0;
    }
  } catch (e0) {}

  var sheetCached = getCachedJson_(templateVersionItemsSheetCacheKey_(target, v));
  if (sheetCached && Array.isArray(sheetCached)) {
    try {
      CacheService.getScriptCache().put(
        templateVersionItemsScriptCacheKey_(target, v),
        JSON.stringify(sheetCached),
        21600
      );
    } catch (e1) {}
    return sheetCached.slice();
  }

  try {
    var sh = getSheet_(TEMPLATE_ITEMS_CACHE_SHEET_);
    var headers = getHeaders_(TEMPLATE_ITEMS_CACHE_SHEET_);
    var cols = getColMap_(TEMPLATE_ITEMS_CACHE_SHEET_);
    var rowNo = findTemplateItemsCacheRowNo_(sh, cols, target, v);
    if (rowNo < 2) return null;
    var row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
    var parsed = parseJsonSafe_(row[cols.items_json], null);
    if (!Array.isArray(parsed)) return null;
    putCachedJson_(templateVersionItemsSheetCacheKey_(target, v), parsed, 900);
    try {
      CacheService.getScriptCache().put(
        templateVersionItemsScriptCacheKey_(target, v),
        JSON.stringify(parsed),
        21600
      );
    } catch (e2) {}
    return parsed;
  } catch (err) {
    return null;
  }
}

function deleteTemplateItemsCacheRows_(templateId) {
  var target = String(templateId || "").trim();
  if (!target) return 0;
  try {
    var sh = getSheet_(TEMPLATE_ITEMS_CACHE_SHEET_);
    var cols = getColMap_(TEMPLATE_ITEMS_CACHE_SHEET_);
    var rowNos = findRowNosByColumnValue_(sh, cols.template_id + 1, target);
    if (!rowNos.length) return 0;
    rowNos.sort(function(a, b) { return b - a; });
    for (var i = 0; i < rowNos.length; i++) sh.deleteRow(rowNos[i]);
    return rowNos.length;
  } catch (e) {
    return 0;
  }
}

function getTemplateVersionDetail(templateId, version, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();

  var target = String(templateId || "").trim();
  if (!target) throw new Error("template_id required");

  var resolvedVersion = resolveTemplateVersionNumber_(target, version);
  var cacheKey = templateVersionDetailCacheKey_(target, resolvedVersion);
  var cached = getCachedJson_(cacheKey);
  if (cached && typeof cached === "object" && Array.isArray(cached.items)) return cached;

  var row = getTemplateVersionRow_(target, resolvedVersion);
  if (!row) throw new Error("Template version not found: " + target + " v" + resolvedVersion);

  var meta = pickTemplateVersionMeta_(row);
  var items = getTemplateVersionItemsFromCache_(target, resolvedVersion);
  if (!Array.isArray(items)) {
    items = normalizeTemplateItems_(parseJsonSafe_(row.items_json, []));
    upsertTemplateVersionItemsCache_(target, resolvedVersion, items);
  }
  var out = {
    template_id: target,
    version: Number(meta.version || resolvedVersion),
    created_at: meta.created_at,
    created_by: meta.created_by,
    source_quote_id: meta.source_quote_id,
    item_count: Number(meta.item_count || items.length),
    summary: meta.summary || {},
    template_meta_snapshot: meta.template_meta_snapshot,
    items: items
  };
  putCachedJson_(cacheKey, out, 21600);
  return out;
}

function saveTemplateVersion(payload, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  return saveTemplateVersionInternal_(payload || {});
}

function saveTemplateVersionInternal_(payload) {
  var data = payload || {};
  return withScriptLock_(15000, function() {
    ensureTemplatesSchemaReady_();

    var nowIso = nowIso_();
    var headersC = getHeaders_(TEMPLATE_CATALOG_SHEET_);
    var colsC = getColMap_(TEMPLATE_CATALOG_SHEET_);
    var shC = getSheet_(TEMPLATE_CATALOG_SHEET_);

    var headersV = getHeaders_(TEMPLATE_VERSIONS_SHEET_);
    var shV = getSheet_(TEMPLATE_VERSIONS_SHEET_);

    var hasCategory = Object.prototype.hasOwnProperty.call(data, "category");
    var hasName = Object.prototype.hasOwnProperty.call(data, "template_name");
    var hasNote = Object.prototype.hasOwnProperty.call(data, "note");
    var targetId = String(data.template_id || "").trim();
    var category = String(data.category || "").trim();
    var templateName = String(data.template_name || "").trim();
    var note = normalizeTemplateTextSummary_(data.note);
    var metaPatch = extractTemplateMetaPatch_(data);
    var sourceQuoteId = String(data.source_quote_id || data.quote_id || "").trim();
    var createdBy = String(data.created_by || "").trim();
    if (!createdBy) createdBy = getSessionUserEmailSafe_() || "admin";

    var items = normalizeTemplateItems_(data.items || []);
    if (!items.length) throw new Error("Template items are empty");

    var catalogRowNo = 0;
    var catalogRow = Array(headersC.length).fill("");
    var existing = null;
    var templateMeta = getDefaultTemplateMeta_();

    if (targetId) {
      existing = readTemplateCatalogRowById_(targetId);
      if (!existing) throw new Error("Template not found: " + targetId);
      catalogRowNo = Number(existing._rowNo || 0);
      if (catalogRowNo < 2) throw new Error("Template row missing: " + targetId);
      catalogRow = shC.getRange(catalogRowNo, 1, 1, headersC.length).getValues()[0];

      if (!hasCategory) category = String(existing.category || "");
      if (!hasName) templateName = String(existing.template_name || "");
      if (!hasNote) note = String(existing.note || "");
      templateMeta = mergeTemplateMeta_(existing, metaPatch);
    } else {
      targetId = uuid_();
      catalogRowNo = shC.getLastRow() + 1;
      ensureRowCapacity_(shC, catalogRowNo);
      setByHeader_(headersC, catalogRow, "template_id", targetId);
      setByHeader_(headersC, catalogRow, "created_at", nowIso);
      setByHeader_(headersC, catalogRow, "is_active", "Y");
      templateMeta = mergeTemplateMeta_(null, metaPatch);
    }

    category = String(category || "").trim() || "General";
    templateName = String(templateName || "").trim();
    if (!templateName) throw new Error("template_name required");
    note = normalizeTemplateTextSummary_(note);

    var rawVersions = listTemplateVersionRowsRaw_(targetId);
    var maxVersion = Number(existing && existing.latest_version || 0);
    for (var i = 0; i < rawVersions.length; i++) {
      var v = Number(rawVersions[i].version || 0);
      if (v > maxVersion) maxVersion = v;
    }
    var nextVersion = maxVersion + 1;

    var summaryObj = buildTemplateSummary_(items);
    var versionRow = Array(headersV.length).fill("");
    setByHeader_(headersV, versionRow, "template_id", targetId);
    setByHeader_(headersV, versionRow, "version", nextVersion);
    setByHeader_(headersV, versionRow, "created_at", nowIso);
    setByHeader_(headersV, versionRow, "created_by", createdBy);
    setByHeader_(headersV, versionRow, "source_quote_id", sourceQuoteId);
    setByHeader_(headersV, versionRow, "items_json", JSON.stringify(items));
    setByHeader_(headersV, versionRow, "summary_json", JSON.stringify(summaryObj));
    setByHeader_(headersV, versionRow, "item_count", items.length);
    setByHeader_(headersV, versionRow, "metadata_snapshot_json", JSON.stringify(templateMeta));

    var versionStart = shV.getLastRow() + 1;
    ensureRowCapacity_(shV, versionStart);
    shV.getRange(versionStart, 1, 1, headersV.length).setValues([versionRow]);
    upsertTemplateVersionItemsCache_(targetId, nextVersion, items);

    setByHeader_(headersC, catalogRow, "template_id", targetId);
    setByHeader_(headersC, catalogRow, "category", category);
    setByHeader_(headersC, catalogRow, "template_name", templateName);
    setByHeader_(headersC, catalogRow, "latest_version", nextVersion);
    setByHeader_(headersC, catalogRow, "latest_item_count", items.length);
    setByHeader_(headersC, catalogRow, "updated_at", nowIso);
    setByHeader_(headersC, catalogRow, "note", note);
    setByHeader_(headersC, catalogRow, "template_type", templateMeta.template_type);
    setByHeader_(headersC, catalogRow, "housing_type", templateMeta.housing_type);
    setByHeader_(headersC, catalogRow, "area_band", templateMeta.area_band);
    setByHeader_(headersC, catalogRow, "style_tags_summary", templateMeta.style_tags_summary);
    setByHeader_(headersC, catalogRow, "tone_tags_summary", templateMeta.tone_tags_summary);
    setByHeader_(headersC, catalogRow, "trade_scope_summary", templateMeta.trade_scope_summary);
    setByHeader_(headersC, catalogRow, "expose_to_prequote", boolToYn_(templateMeta.expose_to_prequote));
    setByHeader_(headersC, catalogRow, "prequote_priority", templateMeta.prequote_priority);
    setByHeader_(headersC, catalogRow, "sort_order", templateMeta.sort_order);
    setByHeader_(headersC, catalogRow, "recommendation_note", templateMeta.recommendation_note);
    setByHeader_(headersC, catalogRow, "target_customer_summary", templateMeta.target_customer_summary);
    setByHeader_(headersC, catalogRow, "budget_band", templateMeta.budget_band);
    setByHeader_(headersC, catalogRow, "recommended_for_summary", templateMeta.recommended_for_summary);
    setByHeader_(headersC, catalogRow, "is_featured_prequote", boolToYn_(templateMeta.is_featured_prequote));
    if (!existing || String(catalogRow[colsC.is_active] || "").trim() === "") {
      setByHeader_(headersC, catalogRow, "is_active", "Y");
    }

    shC.getRange(catalogRowNo, 1, 1, headersC.length).setValues([catalogRow]);
    bumpTemplateListVersion_();

    return {
      ok: true,
      template_id: targetId,
      version: nextVersion,
      latest_version: nextVersion,
      item_count: items.length,
      category: category,
      template_name: templateName,
      template_type: templateMeta.template_type,
      expose_to_prequote: !!templateMeta.expose_to_prequote,
      template_meta: cloneObj_(templateMeta)
    };
  });
}

function setTemplateActive(templateId, isActive, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();

  var target = String(templateId || "").trim();
  if (!target) throw new Error("template_id required");
  var active = !!isActive;

  return withScriptLock_(10000, function() {
    var tpl = readTemplateCatalogRowById_(target);
    if (!tpl) throw new Error("Template not found: " + target);

    var sh = getSheet_(TEMPLATE_CATALOG_SHEET_);
    var headers = getHeaders_(TEMPLATE_CATALOG_SHEET_);
    var rowNo = Number(tpl._rowNo || 0);
    if (rowNo < 2) throw new Error("Template row missing: " + target);
    var row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
    setByHeader_(headers, row, "is_active", templateBoolToCell_(active));
    setByHeader_(headers, row, "updated_at", nowIso_());
    sh.getRange(rowNo, 1, 1, headers.length).setValues([row]);
    bumpTemplateListVersion_();
    return { ok: true, template_id: target, is_active: active };
  });
}

function deleteTemplate(templateId, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();

  var target = String(templateId || "").trim();
  if (!target) throw new Error("template_id required");

  return withScriptLock_(15000, function() {
    var tpl = readTemplateCatalogRowById_(target);
    if (!tpl) throw new Error("Template not found: " + target);

    var shCatalog = getSheet_(TEMPLATE_CATALOG_SHEET_);
    var shVersions = getSheet_(TEMPLATE_VERSIONS_SHEET_);

    var catalogRowNo = Number(tpl._rowNo || 0);
    if (catalogRowNo < 2) throw new Error("Template row missing: " + target);

    var versionRows = listTemplateVersionRowsRaw_(target);
    var versionRowNos = [];
    for (var i = 0; i < versionRows.length; i++) {
      var rn = Number(versionRows[i]._rowNo || 0);
      if (rn >= 2) versionRowNos.push(rn);
    }
    versionRowNos.sort(function(a, b) { return b - a; });
    for (var v = 0; v < versionRowNos.length; v++) {
      shVersions.deleteRow(versionRowNos[v]);
    }

    shCatalog.deleteRow(catalogRowNo);
    var cacheRowsDeleted = deleteTemplateItemsCacheRows_(target);
    bumpTemplateListVersion_();

    return {
      ok: true,
      template_id: target,
      deleted_versions: versionRowNos.length,
      deleted_item_cache_rows: cacheRowsDeleted
    };
  });
}

function toQuoteItemsFromTemplateItems_(items, startSeq) {
  var baseSeq = Number(startSeq || 1);
  var list = normalizeTemplateItems_(items);
  var out = [];
  for (var i = 0; i < list.length; i++) {
    var it = list[i];
    out.push({
      item_id: uuid_(),
      seq: baseSeq + i,
      group_id: String(it.group_id || ""),
      group_label: String(it.group_label || it.process || "기타"),
      group_code: String(it.group_code || ""),
      group_order: Number(it.group_order || 0),
      item_order: Number(it.item_order || (i + 1)),
      name: String(it.name || it.material || ""),
      location: String(it.location || ""),
      detail: String(it.detail || it.spec || ""),
      price: Number(it.price || it.amount || 0),
      price_type: normalizePriceType_(it.price_type || "NORMAL"),
      process: String(it.process || ""),
      material: String(it.material || ""),
      spec: String(it.spec || ""),
      qty: Number(it.qty || 0),
      unit: String(it.unit || ""),
      unit_price: Number(it.unit_price || 0),
      amount: Number(it.amount || (Number(it.qty || 0) * Number(it.unit_price || 0))),
      note: String(it.note || ""),
      material_ref_id: String(it.material_ref_id || ""),
      material_image_id: String(it.material_image_id || ""),
      material_image_name: String(it.material_image_name || "")
    });
  }
  return out;
}

function normalizeImportMode_(mode) {
  var s = String(mode || "").toLowerCase().trim();
  return s === "replace" ? "replace" : "append";
}

function importTemplateIntoQuote(quoteId, templateId, version, mode, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();

  var qid = String(quoteId || "").trim();
  var tid = String(templateId || "").trim();
  if (!qid) throw new Error("quote_id required");
  if (!tid) throw new Error("template_id required");

  var q = findQuoteRow_(qid);
  if (!q) throw new Error("Quote not found: " + qid);

  var resolvedVersion = resolveTemplateVersionNumber_(tid, version);
  var versionRow = getTemplateVersionRow_(tid, resolvedVersion);
  if (!versionRow) throw new Error("Template version not found: " + tid + " v" + resolvedVersion);

  var templateItems = normalizeTemplateItems_(parseJsonSafe_(versionRow.items_json, []));
  if (!templateItems.length) throw new Error("Template has no items");

  var importMode = normalizeImportMode_(mode);
  var merged = [];
  if (importMode === "replace") {
    merged = toQuoteItemsFromTemplateItems_(templateItems, 1);
  } else {
    var existing = normalizeItems_(getItemsFromCache_(qid) || findItems_(qid));
    var appended = toQuoteItemsFromTemplateItems_(templateItems, existing.length + 1);
    merged = existing.concat(appended);
  }

  for (var i = 0; i < merged.length; i++) {
    merged[i].seq = i + 1;
    if (!Number(merged[i].item_order || 0)) merged[i].item_order = i + 1;
  }

  replaceItems_(qid, merged);
  upsertQuoteItemsCache_(qid, merged);

  var vatRate = Number(getSettings_().vat_rate || 0.1);
  var totals = calculateTotalsFromItems_(
    merged,
    q.vat_included,
    vatRate,
    getQuoteTotalOptions_(q)
  );
  updateQuote_(qid, totals);
  bumpQuoteListVersion_();

  return {
    ok: true,
    quote_id: qid,
    template_id: tid,
    version: resolvedVersion,
    mode: importMode,
    imported_count: templateItems.length,
    total_item_count: merged.length
  };
}

function saveQuoteAsTemplateVersion(payload, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();

  var data = payload || {};
  var quoteId = String(data.quote_id || data.quoteId || "").trim();
  var items = normalizeTemplateItems_(data.items || []);

  if (!items.length && quoteId) {
    var q = findQuoteRow_(quoteId);
    if (!q) throw new Error("Quote not found: " + quoteId);
    items = normalizeTemplateItems_(getItemsFromCache_(quoteId) || findItems_(quoteId));
  }
  if (!items.length) throw new Error("No quote items to save as template");

  var templateId = String(data.template_id || "").trim();
  var savePayload = {
    template_id: templateId,
    items: items,
    source_quote_id: quoteId,
    created_by: String(data.created_by || "").trim()
  };

  var isNewTemplate = !templateId || !!data.create_new;
  if (isNewTemplate) savePayload.template_id = "";

  if (isNewTemplate || Object.prototype.hasOwnProperty.call(data, "template_name")) {
    savePayload.template_name = String(data.template_name || "").trim();
  }
  if (isNewTemplate || Object.prototype.hasOwnProperty.call(data, "category")) {
    savePayload.category = String(data.category || "").trim() || "General";
  }
  if (Object.prototype.hasOwnProperty.call(data, "note")) {
    savePayload.note = normalizeTemplateTextSummary_(data.note);
  }
  var metaPatch = extractTemplateMetaPatch_(data);
  var hasMetaPatch = false;
  for (var metaKey in metaPatch) {
    hasMetaPatch = true;
    break;
  }
  if (hasMetaPatch) {
    savePayload.template_meta = metaPatch;
  }

  if (!savePayload.template_id && !String(savePayload.template_name || "").trim()) {
    throw new Error("template_name required for a new template");
  }

  var saved = saveTemplateVersionInternal_(savePayload);
  saved.source_quote_id = quoteId;
  return saved;
}

function generateShare(quoteId, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var qid = String(quoteId || "").trim();
  if (!qid) throw new Error("quoteId required");

  var s = getSettings_();
  var baseUrl = getConfiguredBaseUrl_(s);
  if (!baseUrl) throw new Error("Settings.base_url is required");

  var shareToken = token_();
  var nowMs = Date.now();
  var expireDays = Number(s.share_expire_days || 14);
  var firstFollowUpHours = Number(s.followup_first_hours || 24);

  updateQuote_(qid, {
    share_token: shareToken,
    status: "sent",
    shared_at: new Date(nowMs).toISOString(),
    expire_at: new Date(nowMs + expireDays * 86400000).toISOString(),
    next_action_at: new Date(nowMs + firstFollowUpHours * 3600000).toISOString()
  });

  var urls = buildShareUrlsWithBase_(baseUrl, qid, shareToken);
  bumpQuoteListVersion_();
  return { ok: true, shareToken: shareToken, viewUrl: urls.viewUrl, printUrl: urls.printUrl };
}

function setQuoteApprovalStateByToken_(quoteId, token, targetStateRaw, customerName, customerPhone, options) {
  ensureCoreSchemaReady_();

  var qid = String(quoteId || "").trim();
  var tok = String(token || "").trim();
  if (!qid || !tok) throw new Error("quoteId/token required");

  var opts = options || {};
  var targetState = normalizeApprovalTarget_(targetStateRaw, "");
  if (targetState !== "approved" && targetState !== "not_approved") {
    throw new Error("Invalid approval state");
  }

  var response = null;
  var slackPayload = null;
  var slackDedupKey = "";
  var shouldEnqueueSlack = false;
  var shouldNotifyOwnerApproval = false;
  var ownerEmailDedupKey = "";
  var ownerEmailPayload = null;
  var assigneeNameForSlack = "";

  withScriptLock_(10000, function() {
    var q = findQuoteRow_(qid);
    if (!q) throw new Error("Quote not found");
    if (!q.share_token || String(q.share_token) !== tok) throw new Error("Invalid token");

    var s = getSettings_();
    var oldState = approvalStateFromQuote_(q);
    if (oldState === targetState) {
      var noChangeStatus = (targetState === "approved") ? "approved" : String(q.status || "sent");
      response = {
        ok: true,
        status: noChangeStatus,
        already_approved: (targetState === "approved"),
        already_target_state: true,
        deposit_amount: Number(q.deposit_amount || 0),
        deposit_due_at: String(q.deposit_due_at || ""),
        account_info: String(s.account_info || "")
      };
      return;
    }

    var now = new Date();
    var patch = {
      contact_name: String(customerName || q.contact_name || ""),
      contact_phone: String(customerPhone || q.contact_phone || "")
    };
    var depositAmount = Number(q.deposit_amount || 0);
    var depositDueAt = String(q.deposit_due_at || "");

    if (targetState === "approved") {
      var depositRate = Number(s.deposit_rate || 0.1);
      var dueDays = Number(s.deposit_due_days || 3);
      var total = Number(q.total || 0);
      depositAmount = Math.round(total * depositRate);
      depositDueAt = new Date(Date.now() + dueDays * 86400000).toISOString();
      patch.status = "approved";
      patch.approved_at = nowIso_();
      patch.deposit_amount = depositAmount;
      patch.deposit_due_at = depositDueAt;
      shouldNotifyOwnerApproval = true;
    } else {
      patch.status = String(q.share_token || tok) ? "sent" : "draft";
      patch.approved_at = "";
      patch.deposit_amount = 0;
      patch.deposit_due_at = "";
      depositAmount = 0;
      depositDueAt = "";
    }

    var updated = updateQuote_(qid, patch);
    bumpQuoteListVersion_();

    var eventMillis = parseIsoMs_(updated.updated_at);
    if (isNaN(eventMillis)) eventMillis = now.getTime();
    slackDedupKey = qid + ":" + EVENT_APPROVAL_TOGGLED_ + ":" + oldState + "->" + targetState + ":" + String(eventMillis);
    shouldEnqueueSlack = true;

    assigneeNameForSlack = normalizeAssigneeName_(pickFirstNonEmpty_(updated.contact_name, q.contact_name));
    var assigneeDisplay = assigneeNameForSlack ? ("Assignee: " + assigneeNameForSlack) : "Assignee: (none)";

    slackPayload = {
      quoteId: qid,
      deepLink: buildDeepLinkUrl_(qid, String(q.share_token || tok)),
      customerName: pickFirstNonEmpty_(updated.customer_name, q.customer_name),
      customerPhone: pickFirstNonEmpty_(updated.contact_phone, q.contact_phone),
      customerEmail: pickFirstNonEmpty_(updated.contact_email, updated.customer_email, q.contact_email, q.customer_email),
      siteName: pickFirstNonEmpty_(updated.site_name, q.site_name),
      total: Number(updated.total || q.total || 0),
      assigneeDisplay: assigneeDisplay,
      oldStatus: oldState,
      newStatus: targetState,
      eventTimeIso: formatEventIso_(now)
    };

    response = {
      ok: true,
      status: String(updated.status || ""),
      already_approved: false,
      already_target_state: false,
      deposit_amount: Number(depositAmount || 0),
      deposit_due_at: String(depositDueAt || ""),
      account_info: String(s.account_info || "")
    };

    if (shouldNotifyOwnerApproval) {
      ownerEmailDedupKey = qid + ":" + EVENT_OWNER_APPROVAL_EMAIL_ + ":" + String(eventMillis);
      ownerEmailPayload = { quoteId: qid };
    }
  });

  if (shouldNotifyOwnerApproval && !opts.skipOwnerEmail && ownerEmailPayload) {
    try {
      enqueueOwnerEmailTask_(
        EVENT_OWNER_APPROVAL_EMAIL_,
        qid,
        ownerEmailPayload,
        ownerEmailDedupKey,
        new Date()
      );
    } catch (e) {}
  }
  if (shouldEnqueueSlack && slackPayload) {
    var mentionOrName = resolveSlackMention_(assigneeNameForSlack);
    slackPayload.assigneeDisplay = assigneeNameForSlack
      ? ("Assignee: " + (mentionOrName || assigneeNameForSlack))
      : "Assignee: (none)";
    try { enqueueSlackEvent_(EVENT_APPROVAL_TOGGLED_, slackPayload, slackDedupKey, new Date()); } catch (e2) {}
  }

  return response;
}

function approveQuoteByToken(quoteId, token, customerName, customerPhone) {
  return setQuoteApprovalStateByToken_(
    quoteId,
    token,
    "approved",
    customerName,
    customerPhone,
    { skipOwnerEmail: false }
  );
}

function toggleQuoteApprovalByToken(quoteId, token, targetStatus, customerName, customerPhone) {
  var requested = normalizeApprovalTarget_(targetStatus, "");
  if (!requested) {
    var q = findQuoteRow_(quoteId);
    if (!q) throw new Error("Quote not found");
    requested = approvalStateFromQuote_(q) === "approved" ? "not_approved" : "approved";
  }
  return setQuoteApprovalStateByToken_(
    quoteId,
    token,
    requested,
    customerName,
    customerPhone,
    { skipOwnerEmail: requested !== "approved" }
  );
}

function saveCustomerNoteByToken(quoteId, token, note) {
  ensureCoreSchemaReady_();

  var qid = String(quoteId || "").trim();
  var tok = String(token || "").trim();
  if (!qid || !tok) throw new Error("quoteId/token required");

  var outNote = null;
  var shortNote = "";
  var slackPayload = null;
  var shouldEnqueueSlack = false;
  var slackDedupKey = "";
  var assigneeNameForSlack = "";
  var shouldNotifyOwnerNote = false;
  var ownerEmailDedupKey = "";
  var ownerEmailPayload = null;

  withScriptLock_(10000, function() {
    var q = findQuoteRow_(qid);
    if (!q) throw new Error("Quote not found");
    if (!q.share_token || String(q.share_token) !== tok) throw new Error("Invalid token");

    var n = String(note || "").trim();
    if (!n) throw new Error("Note is empty");

    var nowDate = new Date();
    var nowIso = nowDate.toISOString();
    var noteId = uuid_();
    var shN = getSheet_("CustomerNotes");
    var writeRow = shN.getLastRow() + 1;
    ensureRowCapacity_(shN, writeRow);
    shN.getRange(writeRow, 1, 1, 5).setValues([[noteId, qid, n, nowIso, "customer"]]);

    shortNote = n.length > 60 ? (n.slice(0, 60) + "...") : n;
    updateQuote_(qid, {
      last_customer_note_at: nowIso,
      customer_note_latest: shortNote
    });

    var cachedNotes = getCachedQuoteNotes_(qid);
    if (cachedNotes && Array.isArray(cachedNotes)) {
      cachedNotes.unshift({
        note_id: noteId,
        quote_id: qid,
        note: n,
        created_at: nowIso,
        source: "customer"
      });
      setCachedQuoteNotes_(qid, cachedNotes.slice(0, 500), 600);
    } else {
      invalidateCachedQuoteNotes_(qid);
    }

    slackDedupKey = qid + ":" + EVENT_CHANGE_REQUEST_MEMO_ + ":" + sha256Hex_(normalizeMemoForDedup_(n));
    shouldEnqueueSlack = true;

    assigneeNameForSlack = normalizeAssigneeName_(q.contact_name);
    var assigneeDisplay = assigneeNameForSlack ? ("Assignee: " + assigneeNameForSlack) : "Assignee: (none)";

    slackPayload = {
      quoteId: qid,
      deepLink: buildDeepLinkUrl_(qid, String(q.share_token || tok)),
      customerName: String(q.customer_name || ""),
      customerPhone: String(q.contact_phone || ""),
      customerEmail: pickFirstNonEmpty_(q.contact_email, q.customer_email),
      siteName: String(q.site_name || ""),
      total: Number(q.total || 0),
      assigneeDisplay: assigneeDisplay,
      currentStatus: approvalStateFromQuote_(q),
      memo: n,
      eventTimeIso: formatEventIso_(nowDate)
    };

    outNote = {
      note_id: noteId,
      quote_id: qid,
      note: n,
      created_at: nowIso,
      source: "customer"
    };
    shouldNotifyOwnerNote = true;
    ownerEmailDedupKey = qid + ":" + EVENT_OWNER_NOTE_EMAIL_ + ":" + noteId;
    ownerEmailPayload = {
      quoteId: qid,
      shortNote: shortNote
    };
    bumpQuoteListVersion_();
  });

  if (shouldNotifyOwnerNote && ownerEmailPayload) {
    try {
      enqueueOwnerEmailTask_(
        EVENT_OWNER_NOTE_EMAIL_,
        qid,
        ownerEmailPayload,
        ownerEmailDedupKey,
        new Date()
      );
    } catch (e) {}
  }
  if (shouldEnqueueSlack && slackPayload) {
    var mentionOrName = resolveSlackMention_(assigneeNameForSlack);
    slackPayload.assigneeDisplay = assigneeNameForSlack
      ? ("Assignee: " + (mentionOrName || assigneeNameForSlack))
      : "Assignee: (none)";
    try { enqueueSlackEvent_(EVENT_CHANGE_REQUEST_MEMO_, slackPayload, slackDedupKey, new Date()); } catch (e2) {}
  }

  return { ok: true, note: outNote };
}

function listCustomerNotesByToken(quoteId, token, limit) {
  ensureCoreSchemaReady_();

  var qid = String(quoteId || "").trim();
  var tok = String(token || "").trim();
  if (!qid || !tok) throw new Error("quoteId/token required");

  var q = findQuoteRow_(qid);
  if (!q) throw new Error("Quote not found");
  if (!q.share_token || String(q.share_token) !== tok) throw new Error("Invalid token");

  var max = Math.min(Math.max(Number(limit || 100), 1), 500);
  return listCustomerNotesByQuoteId_(qid, max);
}

function formatAdminDateTime_(value) {
  if (value === null || value === undefined || value === "") return "";
  var d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return String(value || "");
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

function listCustomerNotesAdmin(quoteId, limit, adminCred) {
  assertAdminCredential_(adminCred);
  ensureCoreSchemaReady_();

  var qid = String(quoteId || "").trim();
  if (!qid) throw new Error("quoteId required");

  var max = Math.min(Math.max(Number(limit || 200), 1), 500);
  var list = listCustomerNotesByQuoteId_(qid, max);
  var out = [];
  for (var i = 0; i < list.length; i++) {
    var note = list[i] || {};
    out.push({
      note_id: String(note.note_id || ""),
      quote_id: qid,
      note: String(note.note || ""),
      created_at: formatAdminDateTime_(note.created_at),
      source: String(note.source || "")
    });
  }
  out.sort(function(a, b) {
    return String(b.created_at || "").localeCompare(String(a.created_at || ""));
  });
  return out;
}

function markCustomerNotesRead(quoteId, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  updateQuote_(String(quoteId || "").trim(), { owner_last_seen_note_at: nowIso_() });
  bumpQuoteListVersion_();
  return { ok: true };
}

function listCustomerNotesByQuoteId_(quoteId, limit) {
  var qid = String(quoteId || "").trim();
  if (!qid) return [];

  var max = Math.min(Math.max(Number(limit || 100), 1), 500);
  var cached = getCachedQuoteNotes_(qid);
  if (cached && Array.isArray(cached)) return cached.slice(0, max);

  var sh = getSheet_("CustomerNotes");
  var headers = getHeaders_("CustomerNotes");
  var cols = getColMap_("CustomerNotes");
  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    setCachedQuoteNotes_(qid, [], 180);
    return [];
  }

  var rowNos = findRowNosByColumnValue_(sh, cols.quote_id + 1, qid);
  if (!rowNos.length) {
    setCachedQuoteNotes_(qid, [], 180);
    return [];
  }

  var rowsMap = readRowsMapByRowNos_(sh, rowNos, 1, headers.length);
  var out = [];

  for (var i = 0; i < rowNos.length; i++) {
    var rowNo = rowNos[i];
    var row = rowsMap[rowNo];
    if (!row) continue;
    out.push({
      note_id: String(row[cols.note_id] || ""),
      quote_id: qid,
      note: String(row[cols.note] || ""),
      created_at: String(row[cols.created_at] || ""),
      source: String(row[cols.source] || "")
    });
  }

  out.sort(function(a, b) {
    return String(b.created_at || "").localeCompare(String(a.created_at || ""));
  });
  setCachedQuoteNotes_(qid, out, 600);
  return out.slice(0, max);
}

function trackQuoteViewByToken(quoteId, token, page) {
  ensureCoreSchemaReady_();

  var qid = String(quoteId || "").trim();
  var tok = String(token || "").trim();
  if (!qid || !tok) throw new Error("quoteId/token required");

  var q = findQuoteRow_(qid);
  if (!q) throw new Error("Quote not found");
  if (!q.share_token || String(q.share_token) !== tok) throw new Error("Invalid token");

  var expMs = parseIsoMs_(q.expire_at);
  if (!isNaN(expMs) && expMs < Date.now()) throw new Error("Quote expired");

  logQuoteView_(q, page || "view");
  return { ok: true };
}

function listQuotes(adminPassword, maxCount) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var max = Math.min(Math.max(Number(maxCount || 200), 1), 500);
  var ver = getQuoteListVersion_();
  var cacheKey = "Q_LIST_ALL_" + ver + "_" + String(max);
  var cached = getCachedJson_(cacheKey);
  if (cached && Array.isArray(cached)) return cached;

  var base = listQuotesCore_().slice(0, max);
  var out = [];
  for (var i = 0; i < base.length; i++) {
    var row = cloneObj_(base[i] || {});
    delete row._search_norm;
    out.push(row);
  }
  putCachedJson_(cacheKey, out, 15);
  return out;
}

function quoteDashboardCacheKey_(statusFilter, queryNorm, max, offset) {
  var ver = getQuoteListVersion_();
  return "Q_DASH_" + [
    ver,
    String(statusFilter || ""),
    String(queryNorm || ""),
    String(Number(max || 0)),
    String(Number(offset || 0))
  ].join("_");
}

function quoteCoreListCacheKey_() {
  return "Q_CORE_" + String(getQuoteListVersion_());
}

function listQuotesForDashboard(status, query, limit, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var statusFilter = String(status || "").trim().toLowerCase();
  var queryNorm = normText_(query || "");
  var opts = (limit && typeof limit === "object") ? cloneObj_(limit) : { limit: limit };
  var max = Math.min(Math.max(Number(opts.limit || 50), 1), 500);
  var page = Math.max(Number(opts.page || 1), 1);
  var offset = (opts.offset !== undefined && opts.offset !== null)
    ? Math.max(Math.floor(Number(opts.offset || 0)), 0)
    : Math.max((page - 1) * max, 0);
  var nowMs = Date.now();
  var cacheKey = quoteDashboardCacheKey_(statusFilter, queryNorm, max, offset);
  var out = getOrBuildCachedJsonWithLock_(
    cacheKey,
    20,
    function() {
      var all = listQuotesCore_();
      var rows = [];
      var skipped = 0;
      for (var i = 0; i < all.length; i++) {
        var q = cloneObj_(all[i] || {});
        var computedStatus = String(q.status || "").toLowerCase();
        var expMs = parseIsoMs_(q.expire_at);
        if (computedStatus !== "approved" && !isNaN(expMs) && expMs < nowMs) {
          computedStatus = "expired";
        }

        if (statusFilter && computedStatus !== statusFilter) continue;
        if (queryNorm && !matchesDashboardQuery_(q, queryNorm)) continue;
        if (skipped < offset) {
          skipped++;
          continue;
        }

        q.status = computedStatus;
        delete q._search_norm;
        rows.push(q);
        if (rows.length >= max) break;
      }
      return rows;
    },
    { wait_ms: 1500, retry_ms: 80, fallback_value: [] }
  );
  return Array.isArray(out) ? out : [];
}

function listQuotesCore_() {
  var ver = getQuoteListVersion_();
  if (__QUOTE_LIST_CORE_MEM_ && __QUOTE_LIST_CORE_MEM_VER_ === ver) return __QUOTE_LIST_CORE_MEM_;

  var cacheKey = quoteCoreListCacheKey_();
  var out = getOrBuildCachedJsonWithLock_(
    cacheKey,
    45,
    function() {
      var s = getSettings_();
      var baseUrl = getConfiguredBaseUrl_(s);
      var table = getQuoteTable_();
      var headers = table.headers;
      var values = table.values;
      var list = [];

      for (var i = 0; i < values.length; i++) {
        var obj = rowToObj_(headers, values[i], i + 2);
        if (!obj.quote_id) continue;
        list.push(buildQuoteSummary_(obj, baseUrl));
      }

      list.sort(function(a, b) {
        var up = String(b.updated_at || "").localeCompare(String(a.updated_at || ""));
        if (up !== 0) return up;
        var ap = String(b.approved_at || "").localeCompare(String(a.approved_at || ""));
        if (ap !== 0) return ap;
        var lv = String(b.last_viewed_at || "").localeCompare(String(a.last_viewed_at || ""));
        if (lv !== 0) return lv;
        return String(b.shared_at || "").localeCompare(String(a.shared_at || ""));
      });
      return list;
    },
    { wait_ms: 2200, retry_ms: 120, fallback_value: [] }
  );
  __QUOTE_LIST_CORE_MEM_VER_ = ver;
  __QUOTE_LIST_CORE_MEM_ = Array.isArray(out) ? out : [];
  return __QUOTE_LIST_CORE_MEM_;
}

function buildQuoteSummary_(obj, baseUrl) {
  var token = String(obj.share_token || "").trim();
  var urls = token ? buildShareUrlsWithBase_(baseUrl, String(obj.quote_id || ""), token) : { viewUrl: "", printUrl: "" };

  var lastNoteAt = String(obj.last_customer_note_at || "").trim();
  var seenAt = String(obj.owner_last_seen_note_at || "").trim();
  var hasNewNote = !!lastNoteAt && (!seenAt || lastNoteAt > seenAt);
  var searchNorm = normText_([
    obj.quote_id,
    obj.customer_name,
    obj.site_name,
    obj.contact_name,
    obj.contact_phone,
    obj.memo,
    obj.customer_note_latest
  ].join(" "));

  return {
    quote_id: String(obj.quote_id || ""),
    created_at: String(obj.created_at || ""),
    updated_at: String(obj.updated_at || ""),
    customer_name: String(obj.customer_name || ""),
    site_name: String(obj.site_name || ""),
    contact_name: String(obj.contact_name || ""),
    contact_phone: String(obj.contact_phone || ""),
    memo: String(obj.memo || ""),

    vat_included: String(obj.vat_included || "N"),
    subtotal: Number(obj.subtotal || 0),
    vat: Number(obj.vat || 0),
    total: Number(obj.total || 0),

    share_token: token,
    status: String(obj.status || ""),

    shared_at: String(obj.shared_at || ""),
    expire_at: String(obj.expire_at || ""),
    view_count: Number(obj.view_count || 0),
    last_viewed_at: String(obj.last_viewed_at || ""),
    next_action_at: String(obj.next_action_at || ""),
    last_followup_sent_at: String(obj.last_followup_sent_at || ""),

    approved_at: String(obj.approved_at || ""),
    deposit_amount: Number(obj.deposit_amount || 0),
    deposit_due_at: String(obj.deposit_due_at || ""),

    last_customer_note_at: lastNoteAt,
    customer_note_latest: String(obj.customer_note_latest || ""),
    owner_last_seen_note_at: seenAt,
    has_new_note: hasNewNote,

    view_url: urls.viewUrl,
    print_url: urls.printUrl,
    _search_norm: searchNorm
  };
}

function matchesDashboardQuery_(row, queryNorm) {
  var hay = String(row && row._search_norm || "");
  if (!hay) {
    hay = normText_([
      row.quote_id,
      row.customer_name,
      row.site_name,
      row.contact_name,
      row.contact_phone,
      row.memo,
      row.customer_note_latest
    ].join(" "));
  }
  return hay.indexOf(queryNorm) >= 0;
}

function installAutomationTriggers(adminPassword) {
  assertAdminCredential_(adminPassword);
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var t = triggers[i];
    if (!t.getHandlerFunction) continue;
    var fn = t.getHandlerFunction();
    if (fn === "runAutomation_" || fn === "runSlackQueueWorker_" || fn === "runSlackQueueKick_") {
      ScriptApp.deleteTrigger(t);
    }
  }
  ScriptApp.newTrigger("runAutomation_").timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger("runSlackQueueWorker_").timeBased().everyMinutes(1).create();
  try { CacheService.getScriptCache().put("TRG_SLACKQ_OK", "1", 300); } catch (e) {}
  return { ok: true };
}

function removeAutomationTriggers(adminPassword) {
  assertAdminCredential_(adminPassword);
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var t = triggers[i];
    if (!t.getHandlerFunction) continue;
    var fn = t.getHandlerFunction();
    if (fn === "runAutomation_" || fn === "runSlackQueueWorker_" || fn === "runSlackQueueKick_") {
      ScriptApp.deleteTrigger(t);
    }
  }
  try { CacheService.getScriptCache().remove("TRG_SLACKQ_OK"); } catch (e) {}
  try { CacheService.getScriptCache().remove("TRG_SLACKQ_KICK_PENDING"); } catch (e1) {}
  return { ok: true };
}

function runAutomation_() {
  ensureCoreSchemaReady_();

  var s = getSettings_();
  var ownerTargets = getOwnerEmailTargets_();
  if (!ownerTargets.length) return { ok: true, sent: 0 };

  var cooldownMin = Number(s.followup_cooldown_minutes || 360);
  var repeatHours = Number(s.followup_repeat_hours || 24);
  var maxPerRun = Number(s.max_followups_per_run || 10);
  var baseUrl = getConfiguredBaseUrl_(s);
  if (!baseUrl) return { ok: true, sent: 0 };

  var sh = getSheet_("Quotes");
  var headers = getHeaders_("Quotes");
  var idx = getColMap_("Quotes");
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, sent: 0 };

  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var nowMs = Date.now();
  var sentCount = 0;
  var updates = [];

  for (var i = 0; i < values.length; i++) {
    if (sentCount >= maxPerRun) break;
    var row = values[i];

    var quoteId = String(row[idx.quote_id] || "").trim();
    if (!quoteId) continue;

    var status = String(row[idx.status] || "").toLowerCase();
    if (status !== "sent") continue;
    if (String(row[idx.approved_at] || "").trim()) continue;

    var expMs = parseIsoMs_(row[idx.expire_at]);
    if (!isNaN(expMs) && expMs < nowMs) continue;

    var nextActionMs = parseIsoMs_(row[idx.next_action_at]);
    if (isNaN(nextActionMs) || nextActionMs > nowMs) continue;

    var lastSentMs = parseIsoMs_(row[idx.last_followup_sent_at]);
    if (!isNaN(lastSentMs)) {
      var diffMin = (nowMs - lastSentMs) / 60000;
      if (diffMin < cooldownMin) continue;
    }

    var token = String(row[idx.share_token] || "").trim();
    if (!token) continue;

    var customer = String(row[idx.customer_name] || "");
    var site = String(row[idx.site_name] || "");
    var phone = String(row[idx.contact_phone] || "");
    var total = Number(row[idx.total] || 0);
    var viewCount = Number(row[idx.view_count] || 0);
    var lastViewed = String(row[idx.last_viewed_at] || "").trim();
    var viewUrl = baseUrl + "?page=view&quoteId=" + encodeURIComponent(quoteId) + "&token=" + encodeURIComponent(token);

    var subject = "[QuoteApp] Follow-up: " + (customer || "Customer") + " / " + (site || "Site");
    var body = [
      "Quote follow-up reminder",
      "- Customer/Site: " + customer + " / " + site,
      "- Phone: " + phone,
      "- Total: " + total.toLocaleString("ko-KR") + " KRW",
      "- Views: " + viewCount,
      lastViewed ? ("- Last viewed: " + lastViewed) : "",
      "",
      "Customer link: " + viewUrl
    ].filter(function(v) { return !!v; }).join("\n");

    try {
      sendEmailSafe_(ownerTargets, subject, body);
    } catch (mailErr) {
      continue;
    }

    var lastSentIso = new Date(nowMs).toISOString();
    var nextIso = (repeatHours > 0)
      ? new Date(nowMs + repeatHours * 3600000).toISOString()
      : String(row[idx.next_action_at] || "");

    updates.push({
      rowNo: i + 2,
      last_followup_sent_at: lastSentIso,
      next_action_at: nextIso
    });
    sentCount++;
  }

  if (updates.length) {
    updates.sort(function(a, b) { return a.rowNo - b.rowNo; });

    var startCol = Math.min(idx.last_followup_sent_at, idx.next_action_at) + 1;
    var width = Math.abs(idx.last_followup_sent_at - idx.next_action_at) + 1;
    var outRows = [];
    var rowNos = [];

    for (var u = 0; u < updates.length; u++) {
      var out = Array(width).fill("");
      out[(idx.last_followup_sent_at + 1) - startCol] = updates[u].last_followup_sent_at;
      out[(idx.next_action_at + 1) - startCol] = updates[u].next_action_at;
      outRows.push(out);
      rowNos.push(updates[u].rowNo);
    }

    writeRowsByRowNos_(sh, rowNos, outRows, startCol, width);
    bumpQuoteListVersion_();
  }

  return { ok: true, sent: sentCount };
}

function uploadItemImage(quoteId, itemId, dataUrl, filename, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var qid = String(quoteId || "").trim();
  var iid = String(itemId || "").trim();
  if (!qid || !iid || !dataUrl) throw new Error("quoteId/itemId/dataUrl required");

  var parsed = parseDataUrl_(dataUrl);
  var bytes0 = Utilities.base64Decode(parsed.base64);
  var bytes = tryResizeImageBytes_(bytes0, parsed.mimeType, 900);

  var folder = getOrCreateQuoteUploadFolder_(qid);
  var safeName = sanitizeFilename_(filename || "material.jpg");
  var blob = Utilities.newBlob(bytes, parsed.mimeType, safeName);

  var file = folder.createFile(blob);
  file.setName(qid + "_" + iid + "_" + safeName);

  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  setImageRefCache_(file.getId(), true, 900);

  return { fileId: file.getId(), name: file.getName() };
}

function serveQuoteImage_(e) {
  ensureCoreSchemaReady_();

  var quoteId = String(e && e.parameter && e.parameter.quoteId || "").trim();
  var token = String(e && e.parameter && e.parameter.token || "").trim();
  var fileId = String(e && e.parameter && e.parameter.fileId || "").trim();
  if (!quoteId || !token || !fileId) {
    return ContentService.createTextOutput("missing").setMimeType(ContentService.MimeType.TEXT);
  }

  var q = findQuoteRow_(quoteId);
  if (!q || String(q.share_token || "") !== token) {
    return ContentService.createTextOutput("forbidden").setMimeType(ContentService.MimeType.TEXT);
  }

  var expMs = parseIsoMs_(q.expire_at);
  if (!isNaN(expMs) && expMs < Date.now()) {
    return ContentService.createTextOutput("expired").setMimeType(ContentService.MimeType.TEXT);
  }

  if (!isFileLinkedToQuote_(quoteId, fileId) && !isFileReferencedInSheets_(fileId)) {
    return ContentService.createTextOutput("forbidden").setMimeType(ContentService.MimeType.TEXT);
  }

  var denyKey = "IMG_BLOB_DENY_" + fileId;
  try {
    if (CacheService.getScriptCache().get(denyKey) === "1") {
      return ContentService.createTextOutput("forbidden").setMimeType(ContentService.MimeType.TEXT);
    }
  } catch (e0) {}

  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    try { blob.setContentType(file.getMimeType()); } catch (e1) {}
    try { CacheService.getScriptCache().remove(denyKey); } catch (e2) {}
    return blob;
  } catch (err) {
    try { CacheService.getScriptCache().put(denyKey, "1", 120); } catch (e3) {}
    return ContentService.createTextOutput("forbidden").setMimeType(ContentService.MimeType.TEXT);
  }
}

function servePublicImage_(e) {
  ensureCoreSchemaReady_();

  var fileId = String(e && e.parameter && e.parameter.fileId || "").trim();
  if (!fileId) {
    return ContentService.createTextOutput("missing").setMimeType(ContentService.MimeType.TEXT);
  }
  if (!isLikelyDriveFileId_(fileId)) {
    return ContentService.createTextOutput("403 forbidden").setMimeType(ContentService.MimeType.TEXT);
  }
  if (!isMaterialImageAllowlisted_(fileId)) {
    return ContentService.createTextOutput("403 forbidden").setMimeType(ContentService.MimeType.TEXT);
  }

  var denyKey = "IMG_BLOB_DENY_" + fileId;
  try {
    if (CacheService.getScriptCache().get(denyKey) === "1") {
      return ContentService.createTextOutput("403 forbidden").setMimeType(ContentService.MimeType.TEXT);
    }
  } catch (e0) {}

  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    try { blob.setContentType(file.getMimeType()); } catch (e1) {}
    try { CacheService.getScriptCache().remove(denyKey); } catch (e2) {}
    return blob;
  } catch (err) {
    try { CacheService.getScriptCache().put(denyKey, "1", 120); } catch (e3) {}
    return ContentService.createTextOutput("403 forbidden").setMimeType(ContentService.MimeType.TEXT);
  }
}

function isLikelyDriveFileId_(fileId) {
  var v = String(fileId || "").trim();
  if (!v) return false;
  return /^[A-Za-z0-9_-]{10,}$/.test(v);
}

function materialImageAllowCacheKey_(fileId) {
  return "IMGALLOW_MAT_" + String(fileId || "").trim();
}

function setMaterialImageAllowCache_(fileId, isAllowed, ttlSec) {
  var id = String(fileId || "").trim();
  if (!id) return;
  var ttl = Math.max(Number(ttlSec || 1200), 60);
  try {
    CacheService.getScriptCache().put(materialImageAllowCacheKey_(id), isAllowed ? "1" : "0", ttl);
  } catch (e) {}
}

function clearMaterialImageAllowCache_(fileId) {
  var id = String(fileId || "").trim();
  if (!id) return;
  try { CacheService.getScriptCache().remove(materialImageAllowCacheKey_(id)); } catch (e) {}
}

function isMaterialImageAllowlisted_(fileId) {
  var id = String(fileId || "").trim();
  if (!id) return false;

  var cache = CacheService.getScriptCache();
  var cached = null;
  try { cached = cache.get(materialImageAllowCacheKey_(id)); } catch (e) { cached = null; }
  if (cached === "1") return true;
  if (cached === "0") return false;

  var allowed = sheetColumnHasValue_("Materials", "image_file_id", id);
  setMaterialImageAllowCache_(id, allowed, 1200); // 20 minutes
  return allowed;
}

function isFileReferencedInSheets_(fileId) {
  var target = String(fileId || "").trim();
  if (!target) return false;

  var cacheKey = "IMGREF_" + target;
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached === "1") return true;
  if (cached === "0") return false;

  var checks = [
    { sheet: "Items", col: "material_image_id" },
    { sheet: "Materials", col: "image_file_id" },
    { sheet: "Photos", col: "file_id" }
  ];

  for (var i = 0; i < checks.length; i++) {
    if (sheetColumnHasValue_(checks[i].sheet, checks[i].col, target)) {
      setImageRefCache_(target, true, 300);
      return true;
    }
  }

  setImageRefCache_(target, false, 120);
  return false;
}

function setImageRefCache_(fileId, isReferenced, ttlSec) {
  var target = String(fileId || "").trim();
  if (!target) return;
  var key = "IMGREF_" + target;
  var value = isReferenced ? "1" : "0";
  var ttl = Number(ttlSec || (isReferenced ? 300 : 120));
  try { CacheService.getScriptCache().put(key, value, ttl); } catch (e) {}
}

function markImageRefsCached_(fileIds, ttlSec) {
  var list = Array.isArray(fileIds) ? fileIds : [];
  var ttl = Number(ttlSec || 300);
  var cache = CacheService.getScriptCache();
  for (var i = 0; i < list.length; i++) {
    var id = String(list[i] || "").trim();
    if (!id) continue;
    try { cache.put("IMGREF_" + id, "1", ttl); } catch (e) {}
  }
}

function isFileLinkedToQuote_(quoteId, fileId) {
  var qid = String(quoteId || "").trim();
  var fid = String(fileId || "").trim();
  if (!qid || !fid) return false;

  var cacheKey = "QFILE_" + qid + "_" + fid;
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached === "1") return true;
  if (cached === "0") return false;

  try {
    var itemIdx = getQuoteItemsIndex_(qid);
    if (itemIdx && itemIdx.rowNos && itemIdx.rowNos.length) {
      var shItems = getSheet_("Items");
      var colsItems = getColMap_("Items");
      var minColI = Math.min(colsItems.quote_id, colsItems.material_image_id) + 1;
      var widthI = Math.abs(colsItems.quote_id - colsItems.material_image_id) + 1;
      var rowsI = readRowsMapByRowNos_(shItems, itemIdx.rowNos, minColI, widthI);
      var qPosI = colsItems.quote_id + 1 - minColI;
      var fPosI = colsItems.material_image_id + 1 - minColI;
      for (var ix = 0; ix < itemIdx.rowNos.length; ix++) {
        var rnI = itemIdx.rowNos[ix];
        var rowI = rowsI[rnI];
        if (!rowI) continue;
        if (String(rowI[qPosI] || "").trim() !== qid) continue;
        if (String(rowI[fPosI] || "").trim() !== fid) continue;
        try { cache.put(cacheKey, "1", 300); } catch (eI) {}
        return true;
      }
    }
  } catch (e0) {}

  var targets = [
    { sheet: "Photos", qCol: "quote_id", fCol: "file_id" }
  ];

  for (var t = 0; t < targets.length; t++) {
    var sh = getSheet_(targets[t].sheet);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) continue;

    var cols = getColMap_(targets[t].sheet);
    var qIdx = cols[targets[t].qCol];
    var fIdx = cols[targets[t].fCol];
    if (qIdx === undefined || fIdx === undefined) continue;

    var minCol = Math.min(qIdx, fIdx) + 1;
    var width = Math.abs(qIdx - fIdx) + 1;
    var values = sh.getRange(2, minCol, lastRow - 1, width).getValues();
    var qPos = qIdx + 1 - minCol;
    var fPos = fIdx + 1 - minCol;

    for (var i = 0; i < values.length; i++) {
      if (String(values[i][qPos] || "").trim() !== qid) continue;
      if (String(values[i][fPos] || "").trim() !== fid) continue;
      try { cache.put(cacheKey, "1", 300); } catch (e) {}
      return true;
    }
  }

  try { cache.put(cacheKey, "0", 120); } catch (e) {}
  return false;
}

function sheetColumnSetVersion_(sheetName, colName) {
  var sh = String(sheetName || "").trim();
  var col = String(colName || "").trim();
  if (!sh || !col) return "0";

  if (sh === "Materials" && col === "image_file_id") {
    return "M" + String(getMaterialSearchVersion_());
  }
  if ((sh === "Items" && col === "material_image_id") || (sh === "Photos" && col === "file_id")) {
    return "Q" + String(getQuoteListVersion_());
  }

  try {
    var rowCount = Number(getSheet_(sh).getLastRow() || 0);
    return "R" + String(rowCount);
  } catch (e) {
    return "0";
  }
}

function sheetColumnSetCacheKey_(sheetName, colName, ver) {
  var raw = [
    String(sheetName || "").trim(),
    String(colName || "").trim(),
    String(ver || "").trim()
  ].join("|");
  return "SCSET_" + sha256Hex_(raw).slice(0, 28);
}

function buildSheetColumnValueList_(sheetName, colName) {
  var sh = getSheet_(sheetName);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var cols = getColMap_(sheetName);
  var c = cols[colName];
  if (c === undefined) return [];

  var values = sh.getRange(2, c + 1, lastRow - 1, 1).getValues();
  var out = [];
  var seen = Object.create(null);
  for (var i = 0; i < values.length; i++) {
    var v = String(values[i][0] || "").trim();
    if (!v || seen[v]) continue;
    seen[v] = 1;
    out.push(v);
  }
  return out;
}

function getSheetColumnValueSet_(sheetName, colName) {
  var ver = sheetColumnSetVersion_(sheetName, colName);
  var memKey = [
    String(sheetName || "").trim(),
    String(colName || "").trim(),
    String(ver || "").trim()
  ].join("|");
  if (__SHEET_COLUMN_SET_MEM_[memKey]) return __SHEET_COLUMN_SET_MEM_[memKey];

  var cacheKey = sheetColumnSetCacheKey_(sheetName, colName, ver);
  var list = getOrBuildCachedJsonChunkedWithLock_(
    cacheKey,
    1200,
    function() {
      return buildSheetColumnValueList_(sheetName, colName);
    },
    { wait_ms: 2500, retry_ms: 120, fallback_value: [] }
  );
  var arr = Array.isArray(list) ? list : [];

  var set = Object.create(null);
  for (var i = 0; i < arr.length; i++) {
    var v = String(arr[i] || "").trim();
    if (!v) continue;
    set[v] = 1;
  }
  __SHEET_COLUMN_SET_MEM_[memKey] = set;
  return set;
}

function sheetColumnHasValueDirectScan_(sheetName, colName, target) {
  var t = String(target || "").trim();
  if (!t) return false;
  var sh = getSheet_(sheetName);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  var cols = getColMap_(sheetName);
  var c = cols[colName];
  if (c === undefined) return false;
  var values = sh.getRange(2, c + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === t) return true;
  }
  return false;
}

function sheetColumnHasValue_(sheetName, colName, target) {
  var t = String(target || "").trim();
  if (!t) return false;
  var set = getSheetColumnValueSet_(sheetName, colName);
  if (set[t]) return true;
  var hasAny = false;
  for (var k in set) { hasAny = true; break; }
  if (!hasAny) return sheetColumnHasValueDirectScan_(sheetName, colName, t);
  return false;
}

function ensureMaterialsSheet_() {
  ensureSheetColumns_("Materials", MATERIAL_MASTER_HEADERS_);
}

var MATERIAL_GROUP_MAPPING_ROLE_OPTIONS_ = [
  "LEGACY_REPRESENTATIVE",
  "LEGACY_MAPPING"
];
var MATERIAL_TAG_SOURCE_MATERIALGROUPS_ = "materialgroups";
var MATERIAL_TAG_SOURCE_REF_TYPE_MATERIAL_GROUP_ = "MATERIAL_GROUP";

function materialGroupHeaders_() {
  return [
    "group_key",
    "group_label_ko",
    "mapping_role",
    "space_type",
    "trade_code",
    "material_id",
    "expose_to_prequote",
    "sort_order",
    "recommendation_note",
    "note",
    "is_active",
    "updated_at"
  ];
}

function ensureMaterialGroupsSheet_(overrideId) {
  ensureSheetColumns_(MATERIAL_GROUPS_SHEET_, materialGroupHeaders_(), overrideId);
}

function materialTagHeaders_() {
  return [
    "material_id",
    "tag_type",
    "tag_value",
    "weight",
    "is_active",
    "source",
    "source_ref_type",
    "source_ref_key",
    "note",
    "updated_at"
  ];
}

function ensureMaterialTagsSchema_() {
  ensureSheetColumns_(MATERIAL_TAGS_SHEET_, materialTagHeaders_());
}

function ensureMaterialTagsSheet_(overrideId) {
  ensureSheetColumns_(MATERIAL_TAGS_SHEET_, materialTagHeaders_(), overrideId);
}

function getMaterialTagValuePresets_() {
  return {
    trade: ["BATHROOM", "FLOOR", "WINDOW", "KITCHEN", "DEMOLITION", "WALL", "LIGHTING", "FILM", "PAINT", "FURNITURE"],
    space: ["APARTMENT", "HOUSE", "COMMERCIAL", "OFFICE", "BOTH"],
    mood: ["MODERN", "MINIMAL", "WARM", "COOL", "NATURAL", "VINTAGE", "JAPANDI", "HOTEL_LIKE"],
    tone: ["WHITE", "IVORY", "BEIGE", "GRAY", "GREIGE", "LIGHT_WOOD", "MID_WOOD", "DARK_WOOD", "CHARCOAL", "STONE"],
    texture: ["SMOOTH", "MATTE", "GLOSSY", "WOOD_GRAIN", "STONE", "FABRIC", "CONCRETE"],
    style: ["MODERN", "MINIMAL", "NATURAL", "VINTAGE", "JAPANDI", "SCANDI", "HOTEL_LIKE"],
    feature: ["EASY_CLEAN", "NON_SLIP", "PET_FRIENDLY", "KID_SAFE", "SCRATCH_RESISTANT", "WATER_RESISTANT", "LOW_ODOR", "STORAGE_FRIENDLY"],
    budget: ["ENTRY", "MID", "PREMIUM", "HIGH_END"],
    usage: ["DAILY", "RENTAL", "FAMILY", "PET_HOME", "OFFICE_HEAVY", "SHOWROOM"],
    group_key: [],
    group_label: []
  };
}

function isPrequoteTagType_(tagType) {
  var s = normalizeTagType_(tagType);
  return s === "trade" ||
    s === "space" ||
    s === "mood" ||
    s === "tone" ||
    s === "texture" ||
    s === "style" ||
    s === "feature" ||
    s === "budget" ||
    s === "usage";
}

function normalizeMaterialTagSource_(value) {
  var s = String(value || "").trim().toLowerCase();
  s = s.replace(/[^a-z0-9_]+/g, "_").replace(/^_+|_+$/g, "");
  return s || "catalog_ui";
}

function normalizeMaterialTagSourceRefType_(value) {
  var s = String(value || "").trim().toUpperCase();
  if (s === MATERIAL_TAG_SOURCE_REF_TYPE_MATERIAL_GROUP_) return s;
  return s;
}

function normalizeMaterialTagSourceRefKey_(value, sourceRefType) {
  var s = String(value || "").trim();
  if (!s) return "";
  if (normalizeMaterialTagSourceRefType_(sourceRefType) === MATERIAL_TAG_SOURCE_REF_TYPE_MATERIAL_GROUP_) {
    return normalizeMaterialGroupKey_(s);
  }
  return s;
}

function materialTagUniqueKey_(materialId, tagType, tagValue, source, sourceRefType, sourceRefKey) {
  return [
    String(materialId || "").trim(),
    normalizeTagType_(tagType),
    normalizeTagValue_(tagValue),
    normalizeMaterialTagSource_(source || ""),
    normalizeMaterialTagSourceRefType_(sourceRefType),
    normalizeMaterialTagSourceRefKey_(sourceRefKey, sourceRefType)
  ].join("|");
}

function normalizeMaterialTagPayload_(tagPayload, fallbackMaterialId) {
  var src = tagPayload || {};
  var materialId = String(src.material_id || fallbackMaterialId || "").trim();
  var tagType = normalizeTagType_(src.tag_type || src.type);
  var tagValue = normalizeTagValue_(src.tag_value || src.value);
  if (!materialId) throw new Error("material_id is required");
  if (!tagType) throw new Error("tag_type is required");
  if (!tagValue) throw new Error("tag_value is required");
  return {
    material_id: materialId,
    tag_type: tagType,
    tag_value: tagValue,
    weight: normalizeMaterialTagWeight_(src.weight),
    is_active: normUpperYN_(src.is_active || "Y"),
    source: normalizeMaterialTagSource_(src.source || "catalog_ui"),
    source_ref_type: normalizeMaterialTagSourceRefType_(src.source_ref_type || src.sourceRefType || ""),
    source_ref_key: normalizeMaterialTagSourceRefKey_(src.source_ref_key || src.sourceRefKey || "", src.source_ref_type || src.sourceRefType || ""),
    note: String(src.note || "").trim()
  };
}

function sanitizeMaterialTagsPayload_(materialId, tagsPayload) {
  var targetId = String(materialId || "").trim();
  var input = Array.isArray(tagsPayload) ? tagsPayload : [];
  var out = [];
  var seen = Object.create(null);

  for (var i = 0; i < input.length; i++) {
    var raw = input[i] || {};
    if (ynToBool_(raw._deleted, false)) continue;

    var rawType = String(raw.tag_type || raw.type || "").trim();
    var rawValue = String(raw.tag_value || raw.value || "").trim();
    var rawNote = String(raw.note || "").trim();
    var rawSource = String(raw.source || "").trim();
    var rawWeight = String(raw.weight === undefined || raw.weight === null ? "" : raw.weight).trim();

    if (!rawType && !rawValue && !rawNote && !rawSource && !rawWeight) continue;
    if (!rawType || !rawValue) throw new Error("tag_type and tag_value are required");

    var normalized = normalizeMaterialTagPayload_(raw, targetId);
    var dedupKey = materialTagUniqueKey_(
      targetId,
      normalized.tag_type,
      normalized.tag_value,
      normalized.source,
      normalized.source_ref_type,
      normalized.source_ref_key
    );
    if (seen[dedupKey]) throw new Error("Duplicate tag: " + normalized.tag_type + " / " + normalized.tag_value);
    seen[dedupKey] = 1;
    out.push(normalized);
  }
  return out;
}

function sortMaterialTagsRows_(rows) {
  var list = Array.isArray(rows) ? rows.slice() : [];
  list.sort(function(a, b) {
    var midCmp = String(a.material_id || "").localeCompare(String(b.material_id || ""));
    if (midCmp !== 0) return midCmp;
    var sourceCmp = String(a.source || "").localeCompare(String(b.source || ""));
    if (sourceCmp !== 0) return sourceCmp;
    var refCmp = String(a.source_ref_key || "").localeCompare(String(b.source_ref_key || ""));
    if (refCmp !== 0) return refCmp;
    var typeCmp = String(a.tag_type || "").localeCompare(String(b.tag_type || ""));
    if (typeCmp !== 0) return typeCmp;
    return String(a.tag_value || "").localeCompare(String(b.tag_value || ""));
  });
  return list;
}

function normalizeMaterialTagsRows_(rows) {
  var input = Array.isArray(rows) ? rows : [];
  var now = nowIso_();
  var dedupMap = Object.create(null);

  for (var i = 0; i < input.length; i++) {
    var raw = input[i] || {};
    var rawMaterialId = String(raw.material_id || "").trim();
    var rawTagType = String(raw.tag_type || raw.type || "").trim();
    var rawTagValue = String(raw.tag_value || raw.value || "").trim();
    var rawSource = String(raw.source || "").trim();
    var rawNote = String(raw.note || "").trim();
    var rawSourceRefType = String(raw.source_ref_type || raw.sourceRefType || "").trim();
    var rawSourceRefKey = String(raw.source_ref_key || raw.sourceRefKey || "").trim();
    var rawWeight = String(raw.weight === undefined || raw.weight === null ? "" : raw.weight).trim();
    var rawUpdatedAt = String(raw.updated_at || "").trim();
    if (!rawMaterialId && !rawTagType && !rawTagValue && !rawSource && !rawNote && !rawSourceRefType && !rawSourceRefKey && !rawWeight && !rawUpdatedAt) {
      continue;
    }
    if (!rawMaterialId || !rawTagType || !rawTagValue) continue;

    var normalized = normalizeMaterialTagPayload_(raw, rawMaterialId);
    normalized.updated_at = rawUpdatedAt || now;
    var key = materialTagUniqueKey_(
      normalized.material_id,
      normalized.tag_type,
      normalized.tag_value,
      normalized.source,
      normalized.source_ref_type,
      normalized.source_ref_key
    );
    dedupMap[key] = {
      material_id: normalized.material_id,
      tag_type: normalized.tag_type,
      tag_value: normalized.tag_value,
      weight: normalized.weight,
      is_active: normalized.is_active,
      source: normalized.source,
      source_ref_type: normalized.source_ref_type,
      source_ref_key: normalized.source_ref_key,
      note: normalized.note,
      updated_at: normalized.updated_at
    };
  }

  var keys = Object.keys(dedupMap);
  var out = [];
  for (var k = 0; k < keys.length; k++) out.push(dedupMap[keys[k]]);
  out = sortMaterialTagsRows_(out);
  for (var r = 0; r < out.length; r++) out[r]._rowNo = r + 2;
  return out;
}

function rewriteMaterialTagsSheetInSs_(ss, rows) {
  var spreadsheet = ss || getSpreadsheet_();
  ensureSheetColumnsInSs_(spreadsheet, MATERIAL_TAGS_SHEET_, materialTagHeaders_());
  var sh = getSheetFromSs_(spreadsheet, MATERIAL_TAGS_SHEET_);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var cols = Object.create(null);
  for (var c = 0; c < headers.length; c++) cols[headers[c]] = c;

  var normalizedRows = normalizeMaterialTagsRows_(rows);
  var values = [];
  for (var i = 0; i < normalizedRows.length; i++) {
    var tag = normalizedRows[i];
    var row = Array(headers.length).fill("");
    row[cols.material_id] = tag.material_id;
    row[cols.tag_type] = tag.tag_type;
    row[cols.tag_value] = tag.tag_value;
    row[cols.weight] = normalizeMaterialTagWeight_(tag.weight);
    row[cols.is_active] = normUpperYN_(tag.is_active || "Y");
    if (cols.source !== undefined) row[cols.source] = normalizeMaterialTagSource_(tag.source);
    if (cols.source_ref_type !== undefined) row[cols.source_ref_type] = normalizeMaterialTagSourceRefType_(tag.source_ref_type);
    if (cols.source_ref_key !== undefined) row[cols.source_ref_key] = normalizeMaterialTagSourceRefKey_(tag.source_ref_key, tag.source_ref_type);
    if (cols.note !== undefined) row[cols.note] = String(tag.note || "").trim();
    row[cols.updated_at] = String(tag.updated_at || nowIso_()).trim() || nowIso_();
    values.push(row);
  }

  var lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  if (values.length) {
    ensureRowCapacity_(sh, values.length + 1);
    sh.getRange(2, 1, values.length, headers.length).setValues(values);
  }
  invalidateSheetCache_(MATERIAL_TAGS_SHEET_, spreadsheet.getId());
  invalidateSheetMetaCache_(MATERIAL_TAGS_SHEET_);
  bumpMaterialTagListVersion_();
  return normalizedRows;
}

function rewriteMaterialTagsSheet_(rows) {
  return withScriptLock_(15000, function() {
    ensureMaterialTagsSchema_();
    return rewriteMaterialTagsSheetInSs_(getSpreadsheet_(), rows || []);
  });
}

function buildTagsSummaryFromTagList_(tags) {
  var list = Array.isArray(tags) ? tags : [];
  var out = [];
  var seen = Object.create(null);
  for (var i = 0; i < list.length; i++) {
    var tag = list[i] || {};
    if (normUpperYN_(tag.is_active || "Y") !== "Y") continue;
    if (!isPrequoteTagType_(tag.tag_type)) continue;
    var value = normalizeTagValue_(tag.tag_value);
    if (!value || seen[value]) continue;
    seen[value] = 1;
    out.push(value);
    if (out.length >= 24) break;
  }
  return out.join(", ");
}

function buildMaterialTagStatsMapFromSs_(ss, materialIds) {
  var ids = normalizeStringList_(materialIds || [], function(v) {
    return String(v || "").trim();
  });
  var out = Object.create(null);
  if (!ids.length) return out;

  var tags = listMaterialTagsFromSs_(ss, { material_ids: ids, include_inactive: true });
  for (var i = 0; i < ids.length; i++) {
    out[ids[i]] = {
      total_count: 0,
      active_count: 0,
      prequote_active_count: 0
    };
  }

  for (var t = 0; t < tags.length; t++) {
    var tag = tags[t];
    if (!out[tag.material_id]) {
      out[tag.material_id] = { total_count: 0, active_count: 0, prequote_active_count: 0 };
    }
    out[tag.material_id].total_count += 1;
    if (tag.is_active === "Y") {
      out[tag.material_id].active_count += 1;
      if (isPrequoteTagType_(tag.tag_type)) out[tag.material_id].prequote_active_count += 1;
    }
  }
  return out;
}

function buildMaterialAdminWarnings_(material, tagStats) {
  var m = material || {};
  var stats = tagStats || {};
  var warnings = [];
  if (!!m.is_representative && !m.expose_to_prequote) {
    warnings.push({
      code: "representative_hidden",
      message: "대표 자재인데 가견적 노출이 N이에요."
    });
  }
  if (!!m.is_representative && Number(stats.prequote_active_count || 0) < 1) {
    warnings.push({
      code: "representative_without_prequote_tags",
      message: "대표 자재인데 활성 추천 태그가 없어요."
    });
  }
  return warnings;
}

function decorateMaterialListForAdmin_(ss, materials) {
  var list = Array.isArray(materials) ? materials : [];
  if (!list.length) return list;
  var ids = [];
  for (var i = 0; i < list.length; i++) {
    var id = String(list[i] && list[i].material_id || "").trim();
    if (id) ids.push(id);
  }
  var statsMap = buildMaterialTagStatsMapFromSs_(ss, ids);
  for (var j = 0; j < list.length; j++) {
    var item = list[j] || {};
    var stats = statsMap[String(item.material_id || "").trim()] || { total_count: 0, active_count: 0, prequote_active_count: 0 };
    item.tag_count = Number(stats.total_count || 0);
    item.active_tag_count = Number(stats.active_count || 0);
    item.prequote_tag_count = Number(stats.prequote_active_count || 0);
    item.catalog_warnings = buildMaterialAdminWarnings_(item, stats);
  }
  return list;
}

function normalizeMaterialGroupKey_(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9_]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function normalizeMaterialGroupSpaceType_(value) {
  var raw = String(value || "").trim().toUpperCase();
  if (raw === "RESIDENTIAL" || raw === "COMMERCIAL" || raw === "BOTH") return raw;
  return "BOTH";
}

function normalizeMaterialGroupTradeCode_(value) {
  return normalizeTradeCode_(value);
}

function normalizeMaterialGroupMappingRole_(value) {
  var raw = String(value || "").trim().toUpperCase();
  if (raw === "LEGACY_REPRESENTATIVE") return "LEGACY_REPRESENTATIVE";
  return "LEGACY_MAPPING";
}

function normalizeMaterialGroupSortOrder_(value) {
  var n = Math.floor(normalizeNumber_(value, 0));
  return n > 0 ? n : 0;
}

function normalizeMaterialGroupIsActive_(value) {
  return normUpperYN_(value || "Y");
}

function normalizeMaterialGroupExposeToPrequote_(value) {
  return normUpperYN_(value || "N");
}

function defaultMaterialGroupObj_() {
  return {
    group_key: "",
    original_group_key: "",
    group_label_ko: "",
    mapping_role: "LEGACY_MAPPING",
    space_type: "BOTH",
    trade_code: "",
    material_id: "",
    expose_to_prequote: "N",
    sort_order: 0,
    recommendation_note: "",
    note: "",
    is_active: "Y",
    updated_at: "",
    material_preview: null,
    derived_tag_count: 0,
    expected_derived_tag_count: 0,
    sync_status: "TAGS_NONE"
  };
}

function normalizeMaterialGroupPayload_(groupObj) {
  var src = groupObj || {};
  return {
    group_key: normalizeMaterialGroupKey_(src.group_key),
    original_group_key: normalizeMaterialGroupKey_(src.original_group_key || src.originalGroupKey || src._original_group_key || ""),
    group_label_ko: String(src.group_label_ko || "").trim(),
    mapping_role: normalizeMaterialGroupMappingRole_(src.mapping_role),
    space_type: normalizeMaterialGroupSpaceType_(src.space_type),
    trade_code: normalizeMaterialGroupTradeCode_(src.trade_code),
    material_id: String(src.material_id || "").trim(),
    expose_to_prequote: normalizeMaterialGroupExposeToPrequote_(src.expose_to_prequote),
    sort_order: normalizeMaterialGroupSortOrder_(src.sort_order),
    recommendation_note: String(src.recommendation_note || "").trim(),
    note: String(src.note || "").trim(),
    is_active: normalizeMaterialGroupIsActive_(src.is_active),
    updated_at: String(src.updated_at || "").trim(),
    material_preview: src.material_preview || null,
    derived_tag_count: Number(src.derived_tag_count || 0),
    expected_derived_tag_count: Number(src.expected_derived_tag_count || 0),
    sync_status: String(src.sync_status || "")
  };
}

function materialExistsInSs_(ss, materialId) {
  var id = String(materialId || "").trim();
  if (!id) return false;

  var sh = getSheetFromSs_(ss, "Materials");
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var col = headers.indexOf("material_id");
  if (col < 0) throw new Error("Materials missing column: material_id");

  var values = sh.getRange(2, col + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === id) return true;
  }
  return false;
}

function getMaterialsByIdsFromSs_(ss, materialIds) {
  var ids = Array.isArray(materialIds) ? materialIds : [];
  var wanted = Object.create(null);
  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i] || "").trim();
    if (!id) continue;
    wanted[id] = 1;
  }

  var out = Object.create(null);
  var keys = Object.keys(wanted);
  if (!keys.length) return out;

  var sh = getSheetFromSs_(ss, "Materials");
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return out;

  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var cols = Object.create(null);
  for (var c = 0; c < headers.length; c++) cols[headers[c]] = c;
  if (cols.material_id === undefined) throw new Error("Materials missing column: material_id");

  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  for (var r = 0; r < values.length; r++) {
    var id0 = String(values[r][cols.material_id] || "").trim();
    if (!id0 || !wanted[id0]) continue;
    var rowObj = rowToObj_(headers, values[r], r + 2);
    out[id0] = pickMaterial_(rowObj);
  }
  return out;
}

function readMaterialGroupRowByKeyFromSs_(ss, groupKey) {
  var spreadsheet = ss || getSpreadsheet_();
  ensureSheetColumnsInSs_(spreadsheet, MATERIAL_GROUPS_SHEET_, materialGroupHeaders_());
  var target = normalizeMaterialGroupKey_(groupKey);
  if (!target) return null;

  var sh = getSheetFromSs_(spreadsheet, MATERIAL_GROUPS_SHEET_);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var cols = Object.create(null);
  for (var i = 0; i < headers.length; i++) cols[headers[i]] = i;
  if (cols.group_key === undefined) throw new Error(MATERIAL_GROUPS_SHEET_ + " missing column: group_key");

  var rowNo = findRowNoByColumnValue_(sh, cols.group_key + 1, target);
  if (rowNo < 2) return null;
  var row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
  if (normalizeMaterialGroupKey_(row[cols.group_key]) !== target) return null;
  return toMaterialGroupObj_(row, cols, rowNo);
}

function toMaterialGroupObj_(row, cols, rowNo) {
  var base = normalizeMaterialGroupPayload_({
    group_key: row[cols.group_key],
    original_group_key: row[cols.group_key],
    group_label_ko: row[cols.group_label_ko],
    mapping_role: cols.mapping_role === undefined ? "" : row[cols.mapping_role],
    space_type: row[cols.space_type],
    trade_code: row[cols.trade_code],
    material_id: row[cols.material_id],
    expose_to_prequote: row[cols.expose_to_prequote],
    sort_order: cols.sort_order === undefined ? 0 : row[cols.sort_order],
    recommendation_note: cols.recommendation_note === undefined ? "" : row[cols.recommendation_note],
    note: row[cols.note],
    is_active: row[cols.is_active],
    updated_at: row[cols.updated_at]
  });
  return {
    _rowNo: rowNo,
    group_key: base.group_key,
    original_group_key: base.group_key,
    group_label_ko: base.group_label_ko,
    mapping_role: base.mapping_role,
    space_type: base.space_type,
    trade_code: base.trade_code,
    material_id: base.material_id,
    expose_to_prequote: base.expose_to_prequote,
    sort_order: base.sort_order,
    recommendation_note: base.recommendation_note,
    note: base.note,
    is_active: base.is_active,
    updated_at: base.updated_at
  };
}

function buildDerivedTagsFromMaterialGroup_(groupRow) {
  var group = normalizeMaterialGroupPayload_(groupRow || {});
  var now = nowIso_();
  if (!group.group_key || !group.material_id) return [];

  var specs = [
    { tag_type: "trade", tag_value: group.trade_code, weight: 5 },
    { tag_type: "space", tag_value: group.space_type, weight: 4 },
    { tag_type: "group_key", tag_value: group.group_key, weight: 3 },
    { tag_type: "group_label", tag_value: group.group_label_ko, weight: 2 }
  ];
  var out = [];
  for (var i = 0; i < specs.length; i++) {
    if (!String(specs[i].tag_value || "").trim()) continue;
    out.push(normalizeMaterialTagPayload_({
      material_id: group.material_id,
      tag_type: specs[i].tag_type,
      tag_value: specs[i].tag_value,
      weight: specs[i].weight,
      is_active: "Y",
      source: MATERIAL_TAG_SOURCE_MATERIALGROUPS_,
      source_ref_type: MATERIAL_TAG_SOURCE_REF_TYPE_MATERIAL_GROUP_,
      source_ref_key: group.group_key,
      note: "Derived from MaterialGroups"
    }, group.material_id));
  }
  for (var j = 0; j < out.length; j++) out[j].updated_at = now;
  return out;
}

function removeDerivedTagsForMaterialGroupInSs_(ss, groupKey, materialIdOpt) {
  var spreadsheet = ss || getSpreadsheet_();
  ensureSheetColumnsInSs_(spreadsheet, MATERIAL_TAGS_SHEET_, materialTagHeaders_());
  var targetKey = normalizeMaterialGroupKey_(groupKey);
  var targetMaterialId = String(materialIdOpt || "").trim();
  if (!targetKey) return { ok: true, deleted_count: 0, group_key: "", material_id: targetMaterialId };
  var allTags = listMaterialTagsFromSs_(spreadsheet, { include_inactive: true });
  if (!allTags.length) {
    return { ok: true, deleted_count: 0, group_key: targetKey, material_id: targetMaterialId };
  }

  var out = [];
  var deletedCount = 0;
  for (var i = 0; i < allTags.length; i++) {
    var tag = allTags[i];
    if (
      tag.source === MATERIAL_TAG_SOURCE_MATERIALGROUPS_ &&
      tag.source_ref_type === MATERIAL_TAG_SOURCE_REF_TYPE_MATERIAL_GROUP_ &&
      tag.source_ref_key === targetKey &&
      (!targetMaterialId || tag.material_id === targetMaterialId)
    ) {
      deletedCount += 1;
      continue;
    }
    out.push(tag);
  }
  rewriteMaterialTagsSheetInSs_(spreadsheet, out);
  return {
    ok: true,
    deleted_count: deletedCount,
    group_key: targetKey,
    material_id: targetMaterialId
  };
}

function removeDerivedTagsForMaterialGroup_(groupKey, materialIdOpt) {
  return withScriptLock_(15000, function() {
    ensureMaterialTagsSchema_();
    return removeDerivedTagsForMaterialGroupInSs_(getSpreadsheet_(), groupKey, materialIdOpt);
  });
}

function syncMaterialGroupToTagsInSs_(ss, groupRow, oldGroupRowOpt) {
  var spreadsheet = ss || getSpreadsheet_();
  var current = normalizeMaterialGroupPayload_(groupRow || {});
  var oldGroup = oldGroupRowOpt ? normalizeMaterialGroupPayload_(oldGroupRowOpt) : null;
  ensureSheetColumnsInSs_(spreadsheet, MATERIAL_TAGS_SHEET_, materialTagHeaders_());
  var allTags = listMaterialTagsFromSs_(spreadsheet, { include_inactive: true });
  var cleanupKeys = Object.create(null);
  if (oldGroup && oldGroup.group_key) cleanupKeys[oldGroup.group_key] = 1;
  if (current.group_key) cleanupKeys[current.group_key] = 1;

  var kept = [];
  var removedCount = 0;
  for (var i = 0; i < allTags.length; i++) {
    var existingTag = allTags[i];
    if (
      existingTag.source === MATERIAL_TAG_SOURCE_MATERIALGROUPS_ &&
      existingTag.source_ref_type === MATERIAL_TAG_SOURCE_REF_TYPE_MATERIAL_GROUP_ &&
      cleanupKeys[normalizeMaterialGroupKey_(existingTag.source_ref_key || "")]
    ) {
      removedCount += 1;
      continue;
    }
    kept.push(existingTag);
  }

  var derivedTags = [];
  if (current.is_active === "Y") {
    derivedTags = buildDerivedTagsFromMaterialGroup_(current);
    for (var d = 0; d < derivedTags.length; d++) derivedTags[d].updated_at = nowIso_();
    kept = kept.concat(derivedTags);
  }
  var rewritten = rewriteMaterialTagsSheetInSs_(spreadsheet, kept);
  var actualCount = 0;
  for (var r = 0; r < rewritten.length; r++) {
    var tag = rewritten[r];
    if (
      tag.source === MATERIAL_TAG_SOURCE_MATERIALGROUPS_ &&
      tag.source_ref_type === MATERIAL_TAG_SOURCE_REF_TYPE_MATERIAL_GROUP_ &&
      tag.source_ref_key === current.group_key
    ) {
      actualCount += 1;
    }
  }
  return {
    ok: true,
    mode: current.is_active === "Y" ? "upserted" : "removed",
    group_key: current.group_key,
    material_id: current.material_id,
    removed_count: removedCount,
    upserted_count: actualCount,
    derived_tags: derivedTags
  };
}

function syncMaterialGroupToTags_(groupRow, oldGroupRowOpt) {
  return withScriptLock_(15000, function() {
    ensureMaterialTagsSchema_();
    return syncMaterialGroupToTagsInSs_(getSpreadsheet_(), groupRow || {}, oldGroupRowOpt || null);
  });
}

function materialGroupMatchesFilters_(groupObj, filters, materialPreview) {
  var f = filters || {};
  var g = groupObj || {};
  var m = materialPreview || {};

  var activeFilter = String(f.is_active || f.isActive || "").trim().toUpperCase();
  if (activeFilter === "Y" || activeFilter === "N") {
    if (String(g.is_active || "N").toUpperCase() !== activeFilter) return false;
  }

  var includeInactive = !!(f.include_inactive || f.includeInactive || f.include_all || f.includeAll);
  if (!includeInactive && String(g.is_active || "N") !== "Y") return false;

  var spaceType = normalizeMaterialGroupSpaceType_(f.space_type || f.spaceType || "");
  if ((f.space_type || f.spaceType) && String(g.space_type || "") !== spaceType) return false;

  var tradeCode = normalizeMaterialGroupTradeCode_(f.trade_code || f.tradeCode || "");
  if ((f.trade_code || f.tradeCode) && String(g.trade_code || "") !== tradeCode) return false;

  var exposeToPrequote = String(f.expose_to_prequote || f.exposeToPrequote || "").trim().toUpperCase();
  if ((exposeToPrequote === "Y" || exposeToPrequote === "N") &&
      String(g.expose_to_prequote || "N").toUpperCase() !== exposeToPrequote) {
    return false;
  }

  var q = normText_(f.query || "");
  if (q) {
    var hay = normText_([
      g.group_key,
      g.group_label_ko,
      g.mapping_role,
      g.trade_code,
      g.expose_to_prequote,
      g.recommendation_note,
      g.note,
      g.material_id,
      m.material_id,
      m.name,
      m.brand,
      m.spec
    ].join(" "));
    if (hay.indexOf(q) < 0) return false;
  }
  return true;
}

function buildDerivedTagCountsByGroupKeyFromSs_(ss) {
  var tags = listMaterialTagsFromSs_(ss, {
    source: MATERIAL_TAG_SOURCE_MATERIALGROUPS_,
    source_ref_type: MATERIAL_TAG_SOURCE_REF_TYPE_MATERIAL_GROUP_,
    include_inactive: false
  });
  var counts = Object.create(null);
  for (var i = 0; i < tags.length; i++) {
    var tag = tags[i] || {};
    var key = normalizeMaterialGroupKey_(tag.source_ref_key || "");
    if (!key) continue;
    counts[key] = Number(counts[key] || 0) + 1;
  }
  return counts;
}

function getExpectedDerivedTagCountForMaterialGroup_(groupRow) {
  var group = normalizeMaterialGroupPayload_(groupRow || {});
  if (group.is_active !== "Y") return 0;
  return buildDerivedTagsFromMaterialGroup_(group).length;
}

function computeMaterialGroupSyncStatus_(groupRow, actualCountOpt, expectedCountOpt) {
  var group = normalizeMaterialGroupPayload_(groupRow || {});
  var actualCount = Math.max(Number(actualCountOpt || 0), 0);
  var expectedCount = Math.max(
    expectedCountOpt === undefined || expectedCountOpt === null
      ? getExpectedDerivedTagCountForMaterialGroup_(group)
      : Number(expectedCountOpt || 0),
    0
  );

  if (group.is_active !== "Y") return actualCount > 0 ? "SYNC_MISMATCH" : "INACTIVE";
  if (expectedCount < 1) return actualCount > 0 ? "SYNC_MISMATCH" : "TAGS_NONE";
  if (actualCount === expectedCount) return "SYNCED";
  if (actualCount < 1) return "TAGS_NONE";
  return "SYNC_MISMATCH";
}

function applyMaterialGroupSyncInfo_(groupRow, actualCountOpt, expectedCountOpt, forcedStatusOpt) {
  if (!groupRow) return groupRow;
  var actualCount = Math.max(Number(actualCountOpt || 0), 0);
  var expectedCount = Math.max(
    expectedCountOpt === undefined || expectedCountOpt === null
      ? getExpectedDerivedTagCountForMaterialGroup_(groupRow)
      : Number(expectedCountOpt || 0),
    0
  );
  groupRow.derived_tag_count = actualCount;
  groupRow.expected_derived_tag_count = expectedCount;
  groupRow.sync_status = String(
    forcedStatusOpt || computeMaterialGroupSyncStatus_(groupRow, actualCount, expectedCount)
  );
  return groupRow;
}

function listMaterialGroupsFromSs_(ss, filters) {
  ensureSheetColumnsInSs_(ss, MATERIAL_GROUPS_SHEET_, materialGroupHeaders_());
  ensureSheetColumnsInSs_(ss, MATERIAL_TAGS_SHEET_, materialTagHeaders_());
  var sh = getSheetFromSs_(ss, MATERIAL_GROUPS_SHEET_);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var cols = Object.create(null);
  for (var c = 0; c < headers.length; c++) cols[headers[c]] = c;

  var required = materialGroupHeaders_();
  for (var r0 = 0; r0 < required.length; r0++) {
    if (cols[required[r0]] === undefined) throw new Error(MATERIAL_GROUPS_SHEET_ + " missing column: " + required[r0]);
  }

  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var all = [];
  var materialIds = [];
  var derivedTagCounts = buildDerivedTagCountsByGroupKeyFromSs_(ss);

  for (var i = 0; i < values.length; i++) {
    var obj = toMaterialGroupObj_(values[i], cols, i + 2);
    if (!obj.group_key) continue;
    applyMaterialGroupSyncInfo_(obj, derivedTagCounts[obj.group_key] || 0);
    all.push(obj);
    if (obj.material_id) materialIds.push(obj.material_id);
  }

  var mats = getMaterialsByIdsFromSs_(ss, materialIds);
  var out = [];
  for (var j = 0; j < all.length; j++) {
    var mid0 = String(all[j].material_id || "").trim();
    var preview0 = mid0 ? (mats[mid0] || null) : null;
    if (!materialGroupMatchesFilters_(all[j], filters, preview0)) continue;
    all[j].material_preview = preview0;
    out.push(all[j]);
  }

  out.sort(function(a, b) {
    var sortCmp = Number(a.sort_order || 0) - Number(b.sort_order || 0);
    if (sortCmp !== 0) return sortCmp;
    return String(a.group_key || "").localeCompare(String(b.group_key || ""));
  });
  return out;
}

function upsertMaterialGroupInSs_(ss, groupObj) {
  ensureSheetColumnsInSs_(ss, MATERIAL_GROUPS_SHEET_, materialGroupHeaders_());
  var g = normalizeMaterialGroupPayload_(groupObj || {});
  var groupKey = g.group_key;
  var originalGroupKey = g.original_group_key || groupKey;
  var materialId = g.material_id;
  if (!groupKey) throw new Error("group_key is required");
  if (!materialId) throw new Error("material_id is required");
  if (!materialExistsInSs_(ss, materialId)) throw new Error("Material not found: " + materialId);

  var sh = getSheetFromSs_(ss, MATERIAL_GROUPS_SHEET_);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var cols = Object.create(null);
  for (var c = 0; c < headers.length; c++) cols[headers[c]] = c;
  var required = materialGroupHeaders_();
  for (var r0 = 0; r0 < required.length; r0++) {
    if (cols[required[r0]] === undefined) throw new Error(MATERIAL_GROUPS_SHEET_ + " missing column: " + required[r0]);
  }
  var lastRow = sh.getLastRow();
  var rowNo = 0;
  var duplicateTargetRowNo = 0;

  if (lastRow >= 2) {
    var keyVals = sh.getRange(2, cols.group_key + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < keyVals.length; i++) {
      var normalizedKey = normalizeMaterialGroupKey_(keyVals[i][0]);
      if (!rowNo && normalizedKey === originalGroupKey) {
        rowNo = i + 2;
      }
      if (normalizedKey === groupKey) {
        duplicateTargetRowNo = i + 2;
      }
    }
  }
  if (!rowNo && duplicateTargetRowNo) rowNo = duplicateTargetRowNo;
  if (duplicateTargetRowNo && rowNo && duplicateTargetRowNo !== rowNo) {
    throw new Error("Duplicate group_key: " + groupKey);
  }

  var row;
  if (rowNo > 0) {
    row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
  } else {
    row = Array(headers.length).fill("");
    rowNo = sh.getLastRow() + 1;
  }

  row[cols.group_key] = groupKey;
  row[cols.group_label_ko] = g.group_label_ko;
  row[cols.mapping_role] = g.mapping_role;
  row[cols.space_type] = g.space_type;
  row[cols.trade_code] = g.trade_code;
  row[cols.material_id] = materialId;
  row[cols.expose_to_prequote] = g.expose_to_prequote;
  row[cols.sort_order] = g.sort_order;
  row[cols.recommendation_note] = g.recommendation_note;
  row[cols.note] = g.note;
  row[cols.is_active] = g.is_active;
  row[cols.updated_at] = nowIso_();

  ensureRowCapacity_(sh, rowNo);
  sh.getRange(rowNo, 1, 1, headers.length).setValues([row]);
  invalidateSheetCache_(MATERIAL_GROUPS_SHEET_, ss.getId());
  invalidateSheetMetaCache_(MATERIAL_GROUPS_SHEET_);
  bumpMaterialGroupListVersion_();

  var saved = toMaterialGroupObj_(row, cols, rowNo);
  var mats = getMaterialsByIdsFromSs_(ss, [saved.material_id]);
  saved.material_preview = mats[saved.material_id] || null;
  return saved;
}

function deleteMaterialGroupFromSs_(ss, groupKey) {
  ensureSheetColumnsInSs_(ss, MATERIAL_GROUPS_SHEET_, materialGroupHeaders_());
  var target = normalizeMaterialGroupKey_(groupKey);
  if (!target) throw new Error("group_key is required");

  var sh = getSheetFromSs_(ss, MATERIAL_GROUPS_SHEET_);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var cols = Object.create(null);
  for (var c = 0; c < headers.length; c++) cols[headers[c]] = c;
  if (cols.group_key === undefined) throw new Error(MATERIAL_GROUPS_SHEET_ + " missing column: group_key");
  var lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error("Material group not found: " + target);

  var rows = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var keyVals = sh.getRange(2, cols.group_key + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < keyVals.length; i++) {
    if (normalizeMaterialGroupKey_(keyVals[i][0]) !== target) continue;
    var deletedGroup = toMaterialGroupObj_(rows[i], cols, i + 2);
    sh.deleteRow(i + 2);
    invalidateSheetCache_(MATERIAL_GROUPS_SHEET_, ss.getId());
    invalidateSheetMetaCache_(MATERIAL_GROUPS_SHEET_);
    bumpMaterialGroupListVersion_();
    return { ok: true, group_key: target, deleted_group: deletedGroup };
  }
  throw new Error("Material group not found: " + target);
}

function normalizeMaterialGroupListFilters_(filters) {
  var f = cloneObj_(filters || {});
  var includeInactive = !!(f.include_inactive || f.includeInactive || f.include_all || f.includeAll);
  if (
    f.include_inactive === undefined &&
    f.includeInactive === undefined &&
    f.include_all === undefined &&
    f.includeAll === undefined
  ) {
    includeInactive = true;
  }
  return {
    query: String(f.query || "").trim(),
    space_type: String(f.space_type || f.spaceType || "").trim().toUpperCase(),
    trade_code: String(f.trade_code || f.tradeCode || "").trim().toUpperCase(),
    expose_to_prequote: String(f.expose_to_prequote || f.exposeToPrequote || "").trim().toUpperCase(),
    is_active: String(f.is_active || f.isActive || "").trim().toUpperCase(),
    include_inactive: includeInactive
  };
}

function materialGroupsListCacheKey_(filters) {
  var f = normalizeMaterialGroupListFilters_(filters);
  var st = f.space_type ? normalizeMaterialGroupSpaceType_(f.space_type) : "";
  var tc = f.trade_code ? normalizeMaterialGroupTradeCode_(f.trade_code) : "";
  var active = (f.is_active === "Y" || f.is_active === "N") ? f.is_active : "";
  var raw = [
    "gv=" + String(getMaterialGroupListVersion_()),
    "mv=" + String(getMaterialSearchVersion_()),
    "q=" + normText_(f.query || ""),
    "st=" + st,
    "tc=" + tc,
    "xp=" + String(f.expose_to_prequote || ""),
    "a=" + active,
    "i=" + (f.include_inactive ? "1" : "0")
  ].join("|");
  return "MGRP_LIST_" + sha256Hex_(raw).slice(0, 28);
}

function listMaterialGroupsAdmin(filters, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var opts = normalizeMaterialGroupListFilters_(filters);
  var cacheKey = materialGroupsListCacheKey_(opts);
  var out = getOrBuildCachedJsonWithLock_(
    cacheKey,
    300,
    function() {
      return listMaterialGroupsFromSs_(getSpreadsheet_(), opts);
    },
    { wait_ms: 1600, retry_ms: 80, fallback_value: [] }
  );
  return Array.isArray(out) ? out : [];
}

function upsertMaterialGroupAdmin(groupObj, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  return withScriptLock_(15000, function() {
    var ss = getSpreadsheet_();
    var normalized = normalizeMaterialGroupPayload_(groupObj || {});
    var oldGroup = readMaterialGroupRowByKeyFromSs_(ss, normalized.original_group_key || normalized.group_key);
    var saved = upsertMaterialGroupInSs_(ss, groupObj || {});
    try {
      var sync = syncMaterialGroupToTagsInSs_(ss, saved, oldGroup);
      saved.sync = sync;
      applyMaterialGroupSyncInfo_(saved, saved.is_active === "Y" ? sync.upserted_count : 0);
      return saved;
    } catch (err) {
      saved.sync = {
        ok: false,
        error_message: String((err && err.message) ? err.message : err || "MaterialTags sync failed")
      };
      applyMaterialGroupSyncInfo_(saved, 0, null, "SYNC_ERROR");
      return saved;
    }
  });
}

function deleteMaterialGroupAdmin(groupKey, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  return withScriptLock_(15000, function() {
    var ss = getSpreadsheet_();
    var deleted = deleteMaterialGroupFromSs_(ss, groupKey);
    var deletedGroup = deleted && deleted.deleted_group ? deleted.deleted_group : null;
    if (!deletedGroup || !deletedGroup.group_key) {
      deleted.sync = { ok: true, removed_count: 0, mode: "removed" };
      return deleted;
    }
    try {
      deleted.sync = removeDerivedTagsForMaterialGroupInSs_(ss, deletedGroup.group_key, deletedGroup.material_id);
    } catch (err) {
      deleted.sync = {
        ok: false,
        error_message: String((err && err.message) ? err.message : err || "MaterialTags sync failed")
      };
    }
    return deleted;
  });
}

function getMaterialTagListVersion_() {
  var props = PropertiesService.getScriptProperties();
  var v = String(props.getProperty(MATERIAL_TAG_LIST_VERSION_KEY_) || "").trim();
  if (!v) {
    v = "1";
    props.setProperty(MATERIAL_TAG_LIST_VERSION_KEY_, v);
  }
  return v;
}

function bumpMaterialTagListVersion_() {
  PropertiesService.getScriptProperties().setProperty(MATERIAL_TAG_LIST_VERSION_KEY_, String(Date.now()));
}

function getPrequoteSettingsFromSs_(ss) {
  var spreadsheet = ss || getSpreadsheet_();
  var map = readSettingsMap_(spreadsheet);
  var props = PropertiesService.getScriptProperties();
  return {
    expose_materials_to_prequote: ynToBool_(map.expose_materials_to_prequote, true),
    expose_templates_to_prequote: ynToBool_(map.expose_templates_to_prequote, true),
    default_material_limit: Math.min(Math.max(Number(map.prequote_default_material_limit || 12), 1), 100),
    default_template_limit: Math.min(Math.max(Number(map.prequote_default_template_limit || 12), 1), 100),
    sync_enabled: ynToBool_(map.prequote_sync_enabled, false),
    sync_note: String(map.prequote_sync_note || "").trim(),
    shared_api_key_configured: !!String(props.getProperty("PREQUOTE_SHARED_API_KEY") || "").trim()
  };
}

function getPrequoteSettings_() {
  return getPrequoteSettingsFromSs_(getSpreadsheet_());
}

function getSharedMasterVersionInfoFromSs_(ss) {
  var spreadsheet = ss || getSpreadsheet_();
  var settings = getPrequoteSettingsFromSs_(spreadsheet);
  var materialVersion = getMaterialSearchVersion_();
  var materialTagVersion = getMaterialTagListVersion_();
  var materialGroupVersion = getMaterialGroupListVersion_();
  var templateVersion = getTemplateListVersion_();
  var raw = [
    materialVersion,
    materialTagVersion,
    materialGroupVersion,
    templateVersion,
    settings.expose_materials_to_prequote ? "1" : "0",
    settings.expose_templates_to_prequote ? "1" : "0",
    settings.default_material_limit,
    settings.default_template_limit,
    String(spreadsheet.getId() || "")
  ].join("|");
  return {
    schema_version: CORE_SCHEMA_VERSION_,
    material_version: materialVersion,
    material_tag_version: materialTagVersion,
    material_group_version: materialGroupVersion,
    template_version: templateVersion,
    shared_master_version: sha256Hex_(raw).slice(0, 24),
    generated_at: nowIso_(),
    prequote_settings: settings
  };
}

function getSharedMaterialMasterVersion() {
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  return getSharedMasterVersionInfoFromSs_(getSpreadsheet_());
}

function buildSheetUrlFromSs_(ss, sheetName) {
  var spreadsheet = ss || getSpreadsheet_();
  var sh = spreadsheet.getSheetByName(String(sheetName || "").trim());
  if (!sh) return "";
  return spreadsheet.getUrl() + "#gid=" + String(sh.getSheetId() || "");
}

function buildMaterialImageUrlFromSs_(ss, fileId) {
  var id = String(fileId || "").trim();
  if (!id) return "";
  var spreadsheet = ss || getSpreadsheet_();
  var base = normalizeWebAppBaseUrl_(getSettingFromSs_(spreadsheet, "base_url", ""));
  if (!base) {
    try { base = normalizeWebAppBaseUrl_(getAppUrl_()); } catch (e) { base = ""; }
  }
  if (!base) return "";
  return base + "?page=img&fileId=" + encodeURIComponent(id);
}

function getAdminSpreadsheetUiInfo(adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  var ss = getSpreadsheet_();
  return {
    spreadsheet_id: String(ss.getId() || ""),
    spreadsheet_url: ss.getUrl(),
    materials_url: buildSheetUrlFromSs_(ss, "Materials"),
    material_groups_url: buildSheetUrlFromSs_(ss, MATERIAL_GROUPS_SHEET_),
    material_tags_url: buildSheetUrlFromSs_(ss, MATERIAL_TAGS_SHEET_),
    template_catalog_url: buildSheetUrlFromSs_(ss, TEMPLATE_CATALOG_SHEET_),
    template_versions_url: buildSheetUrlFromSs_(ss, TEMPLATE_VERSIONS_SHEET_)
  };
}

function normalizeMaterialTagWeight_(value) {
  var n = normalizeNumber_(value, 1);
  if (n === 0) return 0;
  return n || 1;
}

function toMaterialTagObj_(row, cols, rowNo) {
  return {
    _rowNo: Number(rowNo || 0),
    material_id: String(row[cols.material_id] || "").trim(),
    tag_type: normalizeTagType_(row[cols.tag_type]),
    tag_value: normalizeTagValue_(row[cols.tag_value]),
    weight: normalizeMaterialTagWeight_(row[cols.weight]),
    is_active: normUpperYN_(row[cols.is_active] || "Y"),
    source: normalizeMaterialTagSource_(cols.source === undefined ? "" : row[cols.source]),
    source_ref_type: normalizeMaterialTagSourceRefType_(cols.source_ref_type === undefined ? "" : row[cols.source_ref_type]),
    source_ref_key: normalizeMaterialTagSourceRefKey_(
      cols.source_ref_key === undefined ? "" : row[cols.source_ref_key],
      cols.source_ref_type === undefined ? "" : row[cols.source_ref_type]
    ),
    note: String(cols.note === undefined ? "" : row[cols.note] || "").trim(),
    updated_at: String(row[cols.updated_at] || "").trim()
  };
}

function listMaterialTagsFromSs_(ss, filters) {
  var spreadsheet = ss || getSpreadsheet_();
  ensureSheetColumnsInSs_(spreadsheet, MATERIAL_TAGS_SHEET_, materialTagHeaders_());
  var sh = getSheetFromSs_(spreadsheet, MATERIAL_TAGS_SHEET_);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var cols = Object.create(null);
  for (var i = 0; i < headers.length; i++) cols[headers[i]] = i;

  var f = filters || {};
  var materialIds = normalizeStringList_(f.material_ids || f.material_id, function(v) {
    return String(v || "").trim();
  });
  var wantedMaterialId = Object.create(null);
  for (var m = 0; m < materialIds.length; m++) wantedMaterialId[materialIds[m]] = 1;
  var tagTypeFilter = normalizeTagType_(f.tag_type || "");
  var sourceFilter = String(f.source || "").trim().toLowerCase();
  var sourceRefTypeFilter = normalizeMaterialTagSourceRefType_(f.source_ref_type || f.sourceRefType || "");
  var sourceRefKeyFilter = normalizeMaterialTagSourceRefKey_(f.source_ref_key || f.sourceRefKey || "", f.source_ref_type || f.sourceRefType || "");
  var includeInactive = !!(f.include_inactive || f.includeInactive);
  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var out = [];

  for (var r = 0; r < values.length; r++) {
    var obj = toMaterialTagObj_(values[r], cols, r + 2);
    if (!obj.material_id || !obj.tag_type || !obj.tag_value) continue;
    if (materialIds.length && !wantedMaterialId[obj.material_id]) continue;
    if (tagTypeFilter && obj.tag_type !== tagTypeFilter) continue;
    if (sourceFilter && obj.source !== sourceFilter) continue;
    if (sourceRefTypeFilter && obj.source_ref_type !== sourceRefTypeFilter) continue;
    if (sourceRefKeyFilter && obj.source_ref_key !== sourceRefKeyFilter) continue;
    if (!includeInactive && obj.is_active !== "Y") continue;
    out.push(obj);
  }

  out.sort(function(a, b) {
    var midCmp = String(a.material_id || "").localeCompare(String(b.material_id || ""));
    if (midCmp !== 0) return midCmp;
    var sourceCmp = String(a.source || "").localeCompare(String(b.source || ""));
    if (sourceCmp !== 0) return sourceCmp;
    var refCmp = String(a.source_ref_key || "").localeCompare(String(b.source_ref_key || ""));
    if (refCmp !== 0) return refCmp;
    var typeCmp = String(a.tag_type || "").localeCompare(String(b.tag_type || ""));
    if (typeCmp !== 0) return typeCmp;
    return String(a.tag_value || "").localeCompare(String(b.tag_value || ""));
  });
  return out;
}

function buildMaterialTagsMapFromSs_(ss, materialIds) {
  var list = listMaterialTagsFromSs_(ss, { material_ids: materialIds || [], include_inactive: false });
  var out = Object.create(null);
  for (var i = 0; i < list.length; i++) {
    var tag = list[i];
    if (!out[tag.material_id]) out[tag.material_id] = [];
    out[tag.material_id].push(tag);
  }
  return out;
}

function replaceMaterialTagsInSs_(ss, materialId, tags) {
  var spreadsheet = ss || getSpreadsheet_();
  var targetId = String(materialId || "").trim();
  if (!targetId) throw new Error("material_id required");
  var allTags = listMaterialTagsFromSs_(spreadsheet, { include_inactive: true });
  var kept = [];
  for (var i = 0; i < allTags.length; i++) {
    var existingTag = allTags[i];
    if (existingTag.material_id === targetId && existingTag.source !== MATERIAL_TAG_SOURCE_MATERIALGROUPS_) continue;
    kept.push(existingTag);
  }

  var input = Array.isArray(tags) ? tags : [];
  var seen = Object.create(null);
  var nowIso = nowIso_();

  for (var t = 0; t < input.length; t++) {
    var tag = normalizeMaterialTagPayload_(input[t] || {}, targetId);
    if (tag.source === MATERIAL_TAG_SOURCE_MATERIALGROUPS_) {
      throw new Error("source=materialgroups tags are managed by MaterialGroups");
    }
    var dedupKey = materialTagUniqueKey_(
      targetId,
      tag.tag_type,
      tag.tag_value,
      tag.source,
      tag.source_ref_type,
      tag.source_ref_key
    );
    if (seen[dedupKey]) continue;
    seen[dedupKey] = 1;
    tag.updated_at = nowIso;
    kept.push(tag);
  }

  var rewritten = rewriteMaterialTagsSheetInSs_(spreadsheet, kept);
  return rewritten.filter(function(tag) {
    return tag.material_id === targetId;
  });
}

function listMaterialTagsAdmin(filters, adminPassword) {
  if (typeof filters === "string") {
    return listMaterialTagsAdmin({ material_id: filters, include_inactive: true }, adminPassword);
  }
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  var opts = cloneObj_(filters || {});
  if ((opts.material_id || opts.material_ids) && opts.include_inactive === undefined && opts.includeInactive === undefined) {
    opts.include_inactive = true;
  }
  return listMaterialTagsFromSs_(getSpreadsheet_(), opts);
}

function replaceMaterialTagsAdmin(materialId, tags, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  return replaceMaterialTagsInSs_(getSpreadsheet_(), materialId, tags);
}

function upsertMaterialTagInSs_(ss, tagPayload) {
  var spreadsheet = ss || getSpreadsheet_();
  var tag = normalizeMaterialTagPayload_(tagPayload || {});
  ensureSheetColumnsInSs_(spreadsheet, MATERIAL_TAGS_SHEET_, materialTagHeaders_());
  if (
    tag.source === MATERIAL_TAG_SOURCE_MATERIALGROUPS_ &&
    (tag.source_ref_type !== MATERIAL_TAG_SOURCE_REF_TYPE_MATERIAL_GROUP_ || !tag.source_ref_key)
  ) {
    throw new Error("Derived materialgroups tags require source_ref_type/source_ref_key");
  }
  if (!materialExistsInSs_(spreadsheet, tag.material_id)) {
    throw new Error("Material not found: " + tag.material_id);
  }
  var wantedKey = materialTagUniqueKey_(
    tag.material_id,
    tag.tag_type,
    tag.tag_value,
    tag.source,
    tag.source_ref_type,
    tag.source_ref_key
  );
  tag.updated_at = nowIso_();

  var allTags = listMaterialTagsFromSs_(spreadsheet, { include_inactive: true });
  var out = [];
  var inserted = false;
  for (var i = 0; i < allTags.length; i++) {
    var existingTag = allTags[i];
    var existingKey = materialTagUniqueKey_(
      existingTag.material_id,
      existingTag.tag_type,
      existingTag.tag_value,
      existingTag.source,
      existingTag.source_ref_type,
      existingTag.source_ref_key
    );
    if (existingKey === wantedKey) {
      if (!inserted) {
        out.push(tag);
        inserted = true;
      }
      continue;
    }
    out.push(existingTag);
  }
  if (!inserted) out.push(tag);

  var rewritten = rewriteMaterialTagsSheetInSs_(spreadsheet, out);
  for (var r = 0; r < rewritten.length; r++) {
    var rowKey = materialTagUniqueKey_(
      rewritten[r].material_id,
      rewritten[r].tag_type,
      rewritten[r].tag_value,
      rewritten[r].source,
      rewritten[r].source_ref_type,
      rewritten[r].source_ref_key
    );
    if (rowKey === wantedKey) return rewritten[r];
  }
  return normalizeMaterialTagPayload_(tag);
}

function deleteMaterialTagInSs_(ss, materialId, tagType, tagValue) {
  var spreadsheet = ss || getSpreadsheet_();
  var targetId = String(materialId || "").trim();
  var targetType = normalizeTagType_(tagType);
  var targetValue = normalizeTagValue_(tagValue);
  if (!targetId) throw new Error("material_id is required");
  if (!targetType) throw new Error("tag_type is required");
  if (!targetValue) throw new Error("tag_value is required");

  var allTags = listMaterialTagsFromSs_(spreadsheet, { include_inactive: true });
  var out = [];
  var deletedCount = 0;
  for (var i = 0; i < allTags.length; i++) {
    var tag = allTags[i];
    if (
      tag.material_id === targetId &&
      tag.tag_type === targetType &&
      tag.tag_value === targetValue &&
      tag.source !== MATERIAL_TAG_SOURCE_MATERIALGROUPS_
    ) {
      deletedCount += 1;
      continue;
    }
    out.push(tag);
  }
  if (deletedCount !== allTags.length || allTags.length) rewriteMaterialTagsSheetInSs_(spreadsheet, out);
  return {
    ok: true,
    deleted: deletedCount,
    material_id: targetId,
    tag_type: targetType,
    tag_value: targetValue
  };
}

function readMaterialRowByIdFromSs_(ss, materialId) {
  var spreadsheet = ss || getSpreadsheet_();
  ensureSheetColumnsInSs_(spreadsheet, "Materials", MATERIAL_MASTER_HEADERS_);
  var targetId = String(materialId || "").trim();
  if (!targetId) return null;

  var sh = getSheetFromSs_(spreadsheet, "Materials");
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var cols = Object.create(null);
  for (var c = 0; c < headers.length; c++) cols[headers[c]] = c;
  if (cols.material_id === undefined) throw new Error("Materials missing column: material_id");

  var rowNo = findRowNoByColumnValue_(sh, cols.material_id + 1, targetId);
  if (rowNo < 2) return null;
  var row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
  if (String(row[cols.material_id] || "").trim() !== targetId) return null;
  return rowToObj_(headers, row, rowNo);
}

function buildMaterialDetailAdminFromSs_(ss, materialId) {
  var spreadsheet = ss || getSpreadsheet_();
  var raw = readMaterialRowByIdFromSs_(spreadsheet, materialId);
  if (!raw) throw new Error("Material not found: " + String(materialId || "").trim());

  var material = pickMaterial_(raw);
  material.image_url = buildMaterialImageUrlFromSs_(spreadsheet, material.image_file_id);
  var tags = listMaterialTagsFromSs_(spreadsheet, { material_id: material.material_id, include_inactive: true });
  var tagStats = buildMaterialTagStatsMapFromSs_(spreadsheet, [material.material_id])[material.material_id] || {
    total_count: 0,
    active_count: 0,
    prequote_active_count: 0
  };
  var legacyGroupMap = buildMaterialGroupMetaMapFromSs_(spreadsheet);
  var legacyGroups = legacyGroupMap.by_material_id[material.material_id] || [];
  return {
    material: material,
    tags: tags,
    tag_count: Number(tagStats.total_count || 0),
    active_tag_count: Number(tagStats.active_count || 0),
    prequote_tag_count: Number(tagStats.prequote_active_count || 0),
    warnings: buildMaterialAdminWarnings_(material, tagStats),
    legacy_groups: legacyGroups,
    generated_tags_summary: buildTagsSummaryFromTagList_(tags),
    fetched_at: nowIso_()
  };
}

function getMaterialDetailAdmin(materialId, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  return buildMaterialDetailAdminFromSs_(getSpreadsheet_(), materialId);
}

function upsertMaterialTagAdmin(tagPayload, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  return upsertMaterialTagInSs_(getSpreadsheet_(), tagPayload || {});
}

function deleteMaterialTagAdmin(materialId, tagType, tagValue, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  return deleteMaterialTagInSs_(getSpreadsheet_(), materialId, tagType, tagValue);
}

function saveMaterialWithTagsAdmin(materialPayload, tagsPayload, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var rawMaterial = cloneObj_(materialPayload || {});
  var imageDataUrl = rawMaterial.image_data_url || rawMaterial.imageDataUrl || null;
  var imageFilename = rawMaterial.image_filename || rawMaterial.imageFilename || null;
  delete rawMaterial.image_data_url;
  delete rawMaterial.imageDataUrl;
  delete rawMaterial.image_filename;
  delete rawMaterial.imageFilename;

  var spreadsheet = getSpreadsheet_();
  return withScriptLock_(10000, function() {
    var draftMaterialId = String(rawMaterial.material_id || "").trim() || "__NEW_MATERIAL__";
    var cleanTags = Array.isArray(tagsPayload)
      ? sanitizeMaterialTagsPayload_(draftMaterialId, tagsPayload)
      : null;

    if ((!rawMaterial.tags_summary || !String(rawMaterial.tags_summary).trim()) && cleanTags) {
      rawMaterial.tags_summary = buildTagsSummaryFromTagList_(cleanTags);
    }

    var saveRes = saveMaterial(rawMaterial, imageDataUrl, imageFilename, adminPassword);
    var materialId = String(saveRes && saveRes.material && saveRes.material.material_id || "").trim();
    if (!materialId) throw new Error("Material save failed");

    if (cleanTags) {
      replaceMaterialTagsInSs_(spreadsheet, materialId, cleanTags.map(function(tag) {
        var cloned = cloneObj_(tag);
        cloned.material_id = materialId;
        return cloned;
      }));
    }

    return buildMaterialDetailAdminFromSs_(spreadsheet, materialId);
  });
}

function listTagValuePresetsAdmin() {
  return {
    tag_types: MATERIAL_TAG_TYPES_.slice(),
    presets: getMaterialTagValuePresets_(),
    generated_at: nowIso_()
  };
}

function ensureSchema(adminPassword) {
  assertAdminCredential_(adminPassword);
  forceEnsureCoreSchema_();
  return {
    ok: true,
    core_schema_version: CORE_SCHEMA_VERSION_,
    materials_headers: getHeaders_("Materials"),
    material_groups_headers: getHeaders_(MATERIAL_GROUPS_SHEET_),
    material_tags_headers: getHeaders_(MATERIAL_TAGS_SHEET_)
  };
}

function appendUniqueTagValue_(bucket, tagType, tagValue) {
  if (!bucket[tagType]) bucket[tagType] = [];
  if (bucket[tagType].indexOf(tagValue) < 0) bucket[tagType].push(tagValue);
}

function hasAnyIntersection_(a, b) {
  var listA = Array.isArray(a) ? a : [];
  var listB = Array.isArray(b) ? b : [];
  if (!listA.length || !listB.length) return false;
  var setB = Object.create(null);
  for (var i = 0; i < listB.length; i++) setB[String(listB[i] || "")] = 1;
  for (var j = 0; j < listA.length; j++) {
    if (setB[String(listA[j] || "")]) return true;
  }
  return false;
}

function buildMaterialGroupMetaMapFromSs_(ss) {
  var spreadsheet = ss || getSpreadsheet_();
  ensureSheetColumnsInSs_(spreadsheet, MATERIAL_GROUPS_SHEET_, materialGroupHeaders_());
  var sh = getSheetFromSs_(spreadsheet, MATERIAL_GROUPS_SHEET_);
  var lastRow = sh.getLastRow();
  var out = {
    by_material_id: Object.create(null),
    by_group_key: Object.create(null)
  };
  if (lastRow < 2) return out;

  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var cols = Object.create(null);
  for (var i = 0; i < headers.length; i++) cols[headers[i]] = i;
  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();

  for (var r = 0; r < values.length; r++) {
    var group = toMaterialGroupObj_(values[r], cols, r + 2);
    if (!group.group_key || !group.material_id || group.is_active !== "Y") continue;
    if (!out.by_material_id[group.material_id]) out.by_material_id[group.material_id] = [];
    out.by_material_id[group.material_id].push({
      group_key: group.group_key,
      group_label_ko: group.group_label_ko,
      trade_code: group.trade_code,
      space_type: group.space_type,
      expose_to_prequote: group.expose_to_prequote,
      note: group.note
    });
    if (!out.by_group_key[group.group_key]) out.by_group_key[group.group_key] = [];
    out.by_group_key[group.group_key].push(group.material_id);
  }
  return out;
}

function buildMaterialMasterSnapshotFromSs_(ss) {
  var spreadsheet = ss || getSpreadsheet_();
  ensureSheetColumnsInSs_(spreadsheet, "Materials", MATERIAL_MASTER_HEADERS_);
  ensureSheetColumnsInSs_(spreadsheet, MATERIAL_GROUPS_SHEET_, materialGroupHeaders_());
  ensureSheetColumnsInSs_(spreadsheet, MATERIAL_TAGS_SHEET_, materialTagHeaders_());

  var versionInfo = getSharedMasterVersionInfoFromSs_(spreadsheet);
  var sh = getSheetFromSs_(spreadsheet, "Materials");
  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return {
      master_version: versionInfo.shared_master_version,
      generated_at: nowIso_(),
      materials: []
    };
  }

  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var cols = Object.create(null);
  for (var c = 0; c < headers.length; c++) cols[headers[c]] = c;
  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var materialIds = [];
  for (var i = 0; i < values.length; i++) {
    var mid = String(values[i][cols.material_id] || "").trim();
    if (mid) materialIds.push(mid);
  }

  var tagsMap = buildMaterialTagsMapFromSs_(spreadsheet, materialIds);
  var groupMetaMap = buildMaterialGroupMetaMapFromSs_(spreadsheet);
  var out = [];

  for (var r = 0; r < values.length; r++) {
    var rowObj = rowToObj_(headers, values[r], r + 2);
    if (!rowObj.material_id) continue;
    var mat = pickMaterial_(rowObj);
    var tagRecords = tagsMap[mat.material_id] || [];
    var tagsByType = Object.create(null);
    var tagWeightMap = Object.create(null);

    for (var t = 0; t < tagRecords.length; t++) {
      var tag = tagRecords[t];
      appendUniqueTagValue_(tagsByType, tag.tag_type, tag.tag_value);
      tagWeightMap[tag.tag_type + "|" + tag.tag_value] = normalizeMaterialTagWeight_(tag.weight);
    }

    var legacyGroups = groupMetaMap.by_material_id[mat.material_id] || [];
    var effectiveTradeCodes = [];
    if (mat.trade_code) effectiveTradeCodes.push(mat.trade_code);
    if (tagsByType.trade) effectiveTradeCodes = effectiveTradeCodes.concat(tagsByType.trade);
    for (var g = 0; g < legacyGroups.length; g++) {
      if (legacyGroups[g].trade_code) effectiveTradeCodes.push(String(legacyGroups[g].trade_code || "").trim());
    }
    effectiveTradeCodes = uniqueStrings_(effectiveTradeCodes);

    var effectiveSpaceValues = [];
    if (mat.space_type) effectiveSpaceValues.push(normalizeTagValue_(mat.space_type));
    if (tagsByType.space) effectiveSpaceValues = effectiveSpaceValues.concat(tagsByType.space);
    for (var s = 0; s < legacyGroups.length; s++) {
      if (legacyGroups[s].space_type) effectiveSpaceValues.push(normalizeTagValue_(legacyGroups[s].space_type));
    }
    effectiveSpaceValues = uniqueStrings_(effectiveSpaceValues);

    var legacyGroupKeys = uniqueStrings_(legacyGroups.map(function(group) {
      return String(group.group_key || "").trim();
    }));
    var summaryTags = splitSummaryTags_(mat.tags_summary);
    var flattenedTagValues = [];
    for (var typeKey in tagsByType) {
      flattenedTagValues = flattenedTagValues.concat(tagsByType[typeKey]);
    }

    out.push({
      material_id: mat.material_id,
      name: mat.name,
      brand: mat.brand,
      spec: mat.spec,
      unit: mat.unit,
      unit_price: mat.unit_price,
      image_file_id: mat.image_file_id,
      image_file_name: mat.image_file_name,
      image_url: buildMaterialImageUrlFromSs_(spreadsheet, mat.image_file_id),
      note: mat.note,
      updated_at: mat.updated_at,
      is_active: !!mat.is_active,
      is_representative: !!mat.is_representative,
      material_type: mat.material_type,
      trade_code: mat.trade_code,
      space_type: mat.space_type,
      sort_order: mat.sort_order,
      expose_to_prequote: !!mat.expose_to_prequote,
      recommendation_score_base: mat.recommendation_score_base,
      price_band: mat.price_band,
      tags_summary: mat.tags_summary,
      recommendation_note: mat.recommendation_note,
      summary_tags: summaryTags,
      tag_records: tagRecords,
      tags_by_type: tagsByType,
      tag_weight_map: tagWeightMap,
      legacy_groups: legacyGroups,
      legacy_group_keys: legacyGroupKeys,
      effective_trade_codes: effectiveTradeCodes,
      effective_space_values: effectiveSpaceValues,
      _search_text: normText_([
        mat.name,
        mat.brand,
        mat.spec,
        mat.note,
        mat.material_type,
        mat.trade_code,
        mat.space_type,
        mat.price_band,
        mat.tags_summary,
        mat.recommendation_note,
        effectiveTradeCodes.join(" "),
        legacyGroupKeys.join(" "),
        flattenedTagValues.join(" ")
      ].join(" "))
    });
  }

  out.sort(function(a, b) {
    var sa = Number(a.sort_order || 0);
    var sb = Number(b.sort_order || 0);
    if (sa !== sb) return sa - sb;
    if (!!a.is_representative !== !!b.is_representative) return a.is_representative ? -1 : 1;
    return String(b.updated_at || "").localeCompare(String(a.updated_at || ""));
  });

  return {
    master_version: versionInfo.shared_master_version,
    generated_at: nowIso_(),
    materials: out
  };
}

function getMaterialMasterSnapshotFromSs_(ss) {
  var spreadsheet = ss || getSpreadsheet_();
  var versionInfo = getSharedMasterVersionInfoFromSs_(spreadsheet);
  var cacheKey = "PREQ_MAT_SNAP_" + sha256Hex_([
    String(spreadsheet.getId() || ""),
    versionInfo.shared_master_version
  ].join("|")).slice(0, 28);
  var snapshot = getOrBuildCachedJsonChunkedWithLock_(
    cacheKey,
    600,
    function() { return buildMaterialMasterSnapshotFromSs_(spreadsheet); },
    { wait_ms: 2200, retry_ms: 120, fallback_value: { master_version: versionInfo.shared_master_version, materials: [] } }
  );
  return snapshot && typeof snapshot === "object" ? snapshot : { master_version: versionInfo.shared_master_version, materials: [] };
}

function normalizePrequoteMaterialFilters_(filters, options, settings) {
  var raw = filters || {};
  var opts = options || {};
  var cfg = settings || getPrequoteSettings_();
  var limit = Math.min(
    Math.max(Number(opts.limit || raw.limit || cfg.default_material_limit || 12), 1),
    100
  );
  return {
    query: String(raw.query || "").trim(),
    group_keys: normalizeStringList_(raw.group_keys || raw.group_key, normalizeMaterialGroupKey_),
    trade_codes: normalizeStringList_(raw.trade_codes || raw.trade_code, normalizeTradeCode_),
    space_type: String(raw.space_type || "").trim() ? normalizeMaterialGroupSpaceType_(raw.space_type) : "",
    mood_tags: normalizeStringList_(raw.mood_tags, normalizeTagValue_),
    tone_tags: normalizeStringList_(raw.tone_tags, normalizeTagValue_),
    texture_tags: normalizeStringList_(raw.texture_tags, normalizeTagValue_),
    style_tags: normalizeStringList_(raw.style_tags, normalizeTagValue_),
    usage_tags: normalizeStringList_(raw.usage_tags, normalizeTagValue_),
    budget_tags: normalizeStringList_(raw.budget_tags, normalizeTagValue_),
    feature_tags: normalizeStringList_(raw.feature_tags, normalizeTagValue_),
    material_types: normalizeStringList_(raw.material_types || raw.material_type, normalizeMaterialType_),
    price_bands: normalizeStringList_(raw.price_bands || raw.price_band, normalizeMaterialPriceBand_),
    representative_only: (raw.representative_only === undefined && raw.representativeOnly === undefined)
      ? true
      : ynToBool_(raw.representative_only !== undefined ? raw.representative_only : raw.representativeOnly, true),
    expose_only: (raw.expose_only === undefined && raw.exposeToPrequote === undefined)
      ? true
      : ynToBool_(raw.expose_only !== undefined ? raw.expose_only : raw.exposeToPrequote, true),
    limit: limit
  };
}

function buildRequestedMaterialTagCriteria_(filters) {
  var f = filters || {};
  return {
    trade: Array.isArray(f.trade_codes) ? f.trade_codes.slice() : [],
    space: f.space_type ? [normalizeTagValue_(f.space_type)] : [],
    mood: Array.isArray(f.mood_tags) ? f.mood_tags.slice() : [],
    tone: Array.isArray(f.tone_tags) ? f.tone_tags.slice() : [],
    texture: Array.isArray(f.texture_tags) ? f.texture_tags.slice() : [],
    style: Array.isArray(f.style_tags) ? f.style_tags.slice() : [],
    usage: Array.isArray(f.usage_tags) ? f.usage_tags.slice() : [],
    budget: Array.isArray(f.budget_tags) ? f.budget_tags.slice() : [],
    feature: Array.isArray(f.feature_tags) ? f.feature_tags.slice() : []
  };
}

function materialPassesPrequoteFilters_(material, filters) {
  var m = material || {};
  var f = filters || {};
  if (!m.material_id || !m.is_active) return false;
  if (f.expose_only && !m.expose_to_prequote) return false;
  if (f.representative_only && !m.is_representative) return false;
  if (f.group_keys && f.group_keys.length && !hasAnyIntersection_(m.legacy_group_keys, f.group_keys)) return false;
  if (f.trade_codes && f.trade_codes.length && !hasAnyIntersection_(m.effective_trade_codes, f.trade_codes)) return false;
  if (f.space_type) {
    if (String(m.space_type || "") !== "BOTH" &&
        String(m.space_type || "") !== f.space_type &&
        !hasAnyIntersection_(m.effective_space_values, [normalizeTagValue_(f.space_type)])) {
      return false;
    }
  }
  if (f.material_types && f.material_types.length && f.material_types.indexOf(String(m.material_type || "")) < 0) return false;
  if (f.price_bands && f.price_bands.length && f.price_bands.indexOf(String(m.price_band || "")) < 0) return false;
  if (f.query) {
    var q = normText_(f.query);
    if (q && String(m._search_text || "").indexOf(q) < 0) return false;
  }
  return true;
}

function scoreMaterialForPrequote_(material, filters) {
  var m = material || {};
  var f = filters || {};
  var requested = buildRequestedMaterialTagCriteria_(f);
  var score = normalizeRecommendationScoreBase_(m.recommendation_score_base);
  var breakdown = [];
  var matchedTagsByType = Object.create(null);
  var pushedReason = Object.create(null);

  if (score) breakdown.push({ type: "base", key: "base", points: score, reason: "recommendation_score_base" });
  if (m.is_representative) {
    score += 2;
    breakdown.push({ type: "flag", key: "representative", points: 2, reason: "representative_material" });
  }
  if (f.group_keys && f.group_keys.length) {
    var matchedGroups = (m.legacy_group_keys || []).filter(function(groupKey) {
      return f.group_keys.indexOf(groupKey) >= 0;
    });
    if (matchedGroups.length) {
      score += matchedGroups.length * 3;
      breakdown.push({ type: "group", key: matchedGroups.join(","), points: matchedGroups.length * 3, reason: "legacy_group_mapping" });
    }
  }
  if (f.material_types && f.material_types.length && f.material_types.indexOf(String(m.material_type || "")) >= 0) {
    score += 3;
    breakdown.push({ type: "field", key: String(m.material_type || ""), points: 3, reason: "material_type_match" });
  }
  if (f.price_bands && f.price_bands.length && f.price_bands.indexOf(String(m.price_band || "")) >= 0) {
    score += 3;
    breakdown.push({ type: "field", key: String(m.price_band || ""), points: 3, reason: "price_band_match" });
  }
  if (f.space_type) {
    if (String(m.space_type || "") === f.space_type) {
      score += 6;
      breakdown.push({ type: "field", key: f.space_type, points: 6, reason: "space_type_exact" });
    } else if (String(m.space_type || "") === "BOTH") {
      score += 3;
      breakdown.push({ type: "field", key: f.space_type, points: 3, reason: "space_type_both" });
    }
  }

  for (var tagType in requested) {
    var wantedValues = requested[tagType] || [];
    var tagMultiplier = Number(MATERIAL_TAG_SCORE_MULTIPLIERS_[tagType] || 2);
    var materialTagValues = (m.tags_by_type && m.tags_by_type[tagType]) ? m.tags_by_type[tagType] : [];
    for (var i = 0; i < wantedValues.length; i++) {
      var wantedValue = String(wantedValues[i] || "");
      if (!wantedValue) continue;
      var reasonKey = tagType + "|" + wantedValue;
      var matched = false;
      var points = 0;
      var reason = "";

      if (tagType === "trade" && (m.effective_trade_codes || []).indexOf(wantedValue) >= 0) {
        matched = true;
        points = Math.max(points, tagMultiplier + 2);
        reason = "trade_match";
      }
      if (tagType === "space" && (m.effective_space_values || []).indexOf(wantedValue) >= 0) {
        matched = true;
        points = Math.max(points, tagMultiplier);
        reason = "space_tag_match";
      }
      if (materialTagValues.indexOf(wantedValue) >= 0) {
        matched = true;
        points = Math.max(points, normalizeMaterialTagWeight_(m.tag_weight_map[reasonKey]) * tagMultiplier);
        reason = "material_tag_match";
      } else if ((m.summary_tags || []).indexOf(wantedValue) >= 0) {
        matched = true;
        points = Math.max(points, Math.max(1, Math.floor(tagMultiplier / 2)));
        reason = "summary_tag_match";
      }

      if (!matched || pushedReason[reasonKey]) continue;
      pushedReason[reasonKey] = 1;
      score += points;
      appendUniqueTagValue_(matchedTagsByType, tagType, wantedValue);
      breakdown.push({ type: tagType, key: wantedValue, points: points, reason: reason });
    }
  }

  breakdown.sort(function(a, b) {
    return Number(b.points || 0) - Number(a.points || 0);
  });
  var primaryReasonTags = [];
  for (var b = 0; b < breakdown.length; b++) {
    var item = breakdown[b];
    if (item.type === "base" || item.type === "flag" || item.type === "field" || item.type === "group") continue;
    var token = item.type + ":" + item.key;
    if (primaryReasonTags.indexOf(token) >= 0) continue;
    primaryReasonTags.push(token);
    if (primaryReasonTags.length >= 3) break;
  }

  return {
    score: score,
    score_breakdown: breakdown,
    matched_tags: matchedTagsByType,
    primary_reason_tags: primaryReasonTags
  };
}

function formatPrequoteMaterialResult_(material, scorePack) {
  var m = material || {};
  var pack = scorePack || { score: 0, score_breakdown: [], matched_tags: {}, primary_reason_tags: [] };
  return {
    material_id: m.material_id,
    brand: m.brand,
    name: m.name,
    spec: m.spec,
    unit: m.unit,
    unit_price: Number(m.unit_price || 0),
    image_file_id: m.image_file_id,
    image_file_name: m.image_file_name,
    image_url: m.image_url || "",
    note: m.note || "",
    material_type: m.material_type || "",
    trade_code: m.trade_code || "",
    space_type: m.space_type || "",
    price_band: m.price_band || "",
    is_representative: !!m.is_representative,
    expose_to_prequote: !!m.expose_to_prequote,
    recommendation_score_base: Number(m.recommendation_score_base || 0),
    sort_order: Number(m.sort_order || 0),
    tags_summary: m.tags_summary || "",
    recommendation_note: m.recommendation_note || "",
    matched_tags: pack.matched_tags || {},
    primary_reason_tags: pack.primary_reason_tags || [],
    score: Number(pack.score || 0),
    score_breakdown: pack.score_breakdown || []
  };
}

function listRepresentativeMaterialsForPrequoteFromSs_(ss, filters, options) {
  var spreadsheet = ss || getSpreadsheet_();
  var settings = getPrequoteSettingsFromSs_(spreadsheet);
  if (!settings.expose_materials_to_prequote) return [];

  var normalized = normalizePrequoteMaterialFilters_(filters, options, settings);
  var snapshot = getMaterialMasterSnapshotFromSs_(spreadsheet);
  var rows = [];
  var list = Array.isArray(snapshot.materials) ? snapshot.materials : [];

  for (var i = 0; i < list.length; i++) {
    var material = list[i];
    if (!materialPassesPrequoteFilters_(material, normalized)) continue;
    rows.push(formatPrequoteMaterialResult_(material, scoreMaterialForPrequote_(material, normalized)));
  }

  rows.sort(function(a, b) {
    if (Number(a.score || 0) !== Number(b.score || 0)) return Number(b.score || 0) - Number(a.score || 0);
    if (!!a.is_representative !== !!b.is_representative) return a.is_representative ? -1 : 1;
    if (Number(a.sort_order || 0) !== Number(b.sort_order || 0)) return Number(a.sort_order || 0) - Number(b.sort_order || 0);
    return String(b.material_id || "").localeCompare(String(a.material_id || ""));
  });
  return rows.slice(0, normalized.limit);
}

function listRepresentativeMaterialsForPrequote(filters, options) {
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  return listRepresentativeMaterialsForPrequoteFromSs_(getSpreadsheet_(), filters, options);
}

function getMaterialDetailForPrequoteFromSs_(ss, materialId) {
  var spreadsheet = ss || getSpreadsheet_();
  var targetId = String(materialId || "").trim();
  if (!targetId) throw new Error("materialId required");
  var snapshot = getMaterialMasterSnapshotFromSs_(spreadsheet);
  var list = Array.isArray(snapshot.materials) ? snapshot.materials : [];
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].material_id || "") !== targetId) continue;
    var detail = formatPrequoteMaterialResult_(list[i], scoreMaterialForPrequote_(list[i], {}));
    detail.tag_records = list[i].tag_records || [];
    detail.tags_by_type = list[i].tags_by_type || {};
    detail.legacy_groups = list[i].legacy_groups || [];
    detail.summary_tags = list[i].summary_tags || [];
    return detail;
  }
  throw new Error("Material not found: " + targetId);
}

function getMaterialDetailForPrequote(materialId) {
  ensureCoreSchemaReady_();
  return getMaterialDetailForPrequoteFromSs_(getSpreadsheet_(), materialId);
}

function getPrequoteMaterialRecommendationFromSs_(ss, input) {
  var spreadsheet = ss || getSpreadsheet_();
  var settings = getPrequoteSettingsFromSs_(spreadsheet);
  var normalized = normalizePrequoteMaterialFilters_(input || {}, {}, settings);
  return {
    master_version: getSharedMasterVersionInfoFromSs_(spreadsheet),
    input: normalized,
    matched_materials: listRepresentativeMaterialsForPrequoteFromSs_(spreadsheet, normalized, { limit: normalized.limit }),
    generated_at: nowIso_()
  };
}

function getPrequoteMaterialRecommendation(input) {
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  return getPrequoteMaterialRecommendationFromSs_(getSpreadsheet_(), input || {});
}

function readTemplateCatalogRowByIdFromSs_(ss, templateId) {
  var spreadsheet = ss || getSpreadsheet_();
  ensureSheetColumnsInSs_(spreadsheet, TEMPLATE_CATALOG_SHEET_, TEMPLATE_CATALOG_HEADERS_);
  var targetId = String(templateId || "").trim();
  if (!targetId) return null;
  var sh = getSheetFromSs_(spreadsheet, TEMPLATE_CATALOG_SHEET_);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var cols = Object.create(null);
  for (var i = 0; i < headers.length; i++) cols[headers[i]] = i;
  if (cols.template_id === undefined) throw new Error("TemplateCatalog missing column: template_id");
  var rowNo = findRowNoByColumnValue_(sh, cols.template_id + 1, targetId);
  if (rowNo < 2) return null;
  var row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
  if (String(row[cols.template_id] || "").trim() !== targetId) return null;
  return pickTemplateCatalog_(rowToObj_(headers, row, rowNo));
}

function listTemplateVersionRowsRawFromSs_(ss, templateId) {
  var spreadsheet = ss || getSpreadsheet_();
  ensureSheetColumnsInSs_(spreadsheet, TEMPLATE_VERSIONS_SHEET_, TEMPLATE_VERSIONS_HEADERS_);
  var targetId = String(templateId || "").trim();
  if (!targetId) return [];
  var sh = getSheetFromSs_(spreadsheet, TEMPLATE_VERSIONS_SHEET_);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var cols = Object.create(null);
  for (var i = 0; i < headers.length; i++) cols[headers[i]] = i;
  if (cols.template_id === undefined) throw new Error("TemplateVersions missing column: template_id");
  var rowNos = findRowNosByColumnValue_(sh, cols.template_id + 1, targetId);
  if (!rowNos.length) return [];
  var rowsMap = readRowsMapByRowNos_(sh, rowNos, 1, headers.length);
  var out = [];
  for (var r = 0; r < rowNos.length; r++) {
    var rowNo = rowNos[r];
    var row = rowsMap[rowNo];
    if (!row) continue;
    if (String(row[cols.template_id] || "").trim() !== targetId) continue;
    out.push(rowToObj_(headers, row, rowNo));
  }
  out.sort(function(a, b) {
    var av = Number(a.version || 0);
    var bv = Number(b.version || 0);
    if (av !== bv) return bv - av;
    return Number(b._rowNo || 0) - Number(a._rowNo || 0);
  });
  return out;
}

function getTemplateVersionRowFromSs_(ss, templateId, version) {
  var rows = listTemplateVersionRowsRawFromSs_(ss, templateId);
  var v = Number(version || 0);
  for (var i = 0; i < rows.length; i++) {
    if (Number(rows[i].version || 0) === v) return rows[i];
  }
  return null;
}

function resolveTemplateVersionNumberFromSs_(ss, templateId, version) {
  var v = Number(version || 0);
  if (v >= 1) return Math.floor(v);
  var tpl = readTemplateCatalogRowByIdFromSs_(ss, templateId);
  if (!tpl) throw new Error("Template not found: " + String(templateId || "").trim());
  var latest = Number(tpl.latest_version || 0);
  if (latest < 1) throw new Error("Template has no versions: " + String(templateId || "").trim());
  return latest;
}

function loadTemplateSummaryMapByVersionFromSs_(ss, versionByTemplateId) {
  var wanted = versionByTemplateId || {};
  var hasWanted = false;
  for (var k in wanted) { hasWanted = true; break; }
  if (!hasWanted) return {};

  var spreadsheet = ss || getSpreadsheet_();
  ensureSheetColumnsInSs_(spreadsheet, TEMPLATE_VERSIONS_SHEET_, TEMPLATE_VERSIONS_HEADERS_);
  var sh = getSheetFromSs_(spreadsheet, TEMPLATE_VERSIONS_SHEET_);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var cols = Object.create(null);
  for (var i = 0; i < headers.length; i++) cols[headers[i]] = i;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var out = Object.create(null);
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var tid = String(row[cols.template_id] || "").trim();
    if (!tid || !Object.prototype.hasOwnProperty.call(wanted, tid)) continue;
    var expectedVer = Number(wanted[tid] || 0);
    var rowVer = Number(row[cols.version] || 0);
    if (expectedVer !== rowVer) continue;
    out[tid] = parseJsonSafe_(row[cols.summary_json], {});
  }
  return out;
}

function buildTemplateCatalogSnapshotFromSs_(ss) {
  var spreadsheet = ss || getSpreadsheet_();
  ensureSheetColumnsInSs_(spreadsheet, TEMPLATE_CATALOG_SHEET_, TEMPLATE_CATALOG_HEADERS_);
  var sh = getSheetFromSs_(spreadsheet, TEMPLATE_CATALOG_SHEET_);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var out = [];
  var versionsByTemplateId = Object.create(null);
  for (var i = 0; i < values.length; i++) {
    var rowObj = pickTemplateCatalog_(rowToObj_(headers, values[i], i + 2));
    if (!rowObj.template_id) continue;
    out.push(rowObj);
    versionsByTemplateId[rowObj.template_id] = Number(rowObj.latest_version || 0);
  }

  var summaryMap = loadTemplateSummaryMapByVersionFromSs_(spreadsheet, versionsByTemplateId);
  for (var j = 0; j < out.length; j++) {
    out[j]._search_text = buildTemplateSearchText_(out[j], summaryMap[out[j].template_id] || {});
  }
  return out;
}

function getTemplateCatalogSnapshotFromSs_(ss) {
  var spreadsheet = ss || getSpreadsheet_();
  var versionInfo = getSharedMasterVersionInfoFromSs_(spreadsheet);
  var cacheKey = "PREQ_TPL_SNAP_" + sha256Hex_([
    String(spreadsheet.getId() || ""),
    versionInfo.shared_master_version
  ].join("|")).slice(0, 28);
  var snapshot = getOrBuildCachedJsonChunkedWithLock_(
    cacheKey,
    600,
    function() { return buildTemplateCatalogSnapshotFromSs_(spreadsheet); },
    { wait_ms: 2200, retry_ms: 120, fallback_value: [] }
  );
  return Array.isArray(snapshot) ? snapshot : [];
}

function normalizePrequoteTemplateFilters_(filters, settings) {
  var raw = filters || {};
  var cfg = settings || getPrequoteSettings_();
  return {
    category: String(raw.category || "").trim(),
    query: String(raw.query || "").trim(),
    housing_type: raw.housing_type !== undefined && raw.housing_type !== null && String(raw.housing_type).trim() !== ""
      ? normalizeHousingType_(raw.housing_type)
      : "",
    area_band: raw.area_band !== undefined && raw.area_band !== null && String(raw.area_band).trim() !== ""
      ? normalizeAreaBand_(raw.area_band)
      : "",
    style_tags: splitSummaryTags_(raw.style_tags || raw.style_tags_summary),
    tone_tags: splitSummaryTags_(raw.tone_tags || raw.tone_tags_summary),
    trade_scope_tags: splitSummaryTags_(raw.trade_scope_tags || raw.trade_scope || raw.trade_scope_summary),
    template_types: normalizeStringList_(raw.template_types || raw.template_type, normalizeTemplateType_),
    include_inactive: ynToBool_(raw.include_inactive, false),
    limit: Math.min(Math.max(Number(raw.limit || cfg.default_template_limit || 12), 1), 100)
  };
}

function templateTypeMatchesFilter_(rowType, allowedTypes) {
  var type = normalizeTemplateType_(rowType);
  var list = Array.isArray(allowedTypes) ? allowedTypes : [];
  if (!list.length) return type === PREQUOTE_TEMPLATE_TYPES_.PREQUOTE_PACKAGE || type === PREQUOTE_TEMPLATE_TYPES_.BOTH;
  if (type === PREQUOTE_TEMPLATE_TYPES_.BOTH) {
    return list.indexOf(PREQUOTE_TEMPLATE_TYPES_.QUOTE) >= 0 ||
      list.indexOf(PREQUOTE_TEMPLATE_TYPES_.PREQUOTE_PACKAGE) >= 0 ||
      list.indexOf(PREQUOTE_TEMPLATE_TYPES_.BOTH) >= 0;
  }
  return list.indexOf(type) >= 0;
}

function templatePassesPrequoteFilters_(row, filters, settings) {
  var tpl = row || {};
  var f = filters || {};
  var cfg = settings || getPrequoteSettings_();
  var styleTags = Array.isArray(f.style_tags) ? f.style_tags : [];
  var toneTags = Array.isArray(f.tone_tags) ? f.tone_tags : [];
  var tradeScopeTags = Array.isArray(f.trade_scope_tags) ? f.trade_scope_tags : [];
  if (!tpl.template_id) return false;
  if (!cfg.expose_templates_to_prequote) return false;
  if (!f.include_inactive && !tpl.is_active) return false;
  if (!tpl.expose_to_prequote) return false;
  if (!templateTypeMatchesFilter_(tpl.template_type, f.template_types)) return false;
  if (f.category && String(tpl.category || "") !== f.category) return false;
  if (f.housing_type && String(tpl.housing_type || "") !== "ALL" && String(tpl.housing_type || "") !== f.housing_type) return false;
  if (f.area_band && String(tpl.area_band || "") !== "ALL" && String(tpl.area_band || "") !== f.area_band) return false;
  if (styleTags.length && !hasAnyIntersection_(splitSummaryTags_(tpl.style_tags_summary), styleTags)) return false;
  if (toneTags.length && !hasAnyIntersection_(splitSummaryTags_(tpl.tone_tags_summary), toneTags)) return false;
  if (tradeScopeTags.length && !hasAnyIntersection_(splitSummaryTags_(tpl.trade_scope_summary), tradeScopeTags)) return false;
  if (f.query) {
    var tokens = splitTemplateQueryTokens_(f.query);
    var searchText = String(tpl._search_text || "");
    for (var i = 0; i < tokens.length; i++) {
      if (searchText.indexOf(tokens[i]) < 0) return false;
    }
  }
  return true;
}

function listTemplateCatalogForPrequoteFromSs_(ss, filters) {
  var spreadsheet = ss || getSpreadsheet_();
  var settings = getPrequoteSettingsFromSs_(spreadsheet);
  var normalized = normalizePrequoteTemplateFilters_(filters || {}, settings);
  var snapshot = getTemplateCatalogSnapshotFromSs_(spreadsheet);
  var out = [];

  for (var i = 0; i < snapshot.length; i++) {
    if (!templatePassesPrequoteFilters_(snapshot[i], normalized, settings)) continue;
    var row = cloneTemplateCatalogRow_(snapshot[i]);
    delete row._search_text;
    out.push(row);
  }

  out.sort(function(a, b) {
    if (!!a.is_featured_prequote !== !!b.is_featured_prequote) return a.is_featured_prequote ? -1 : 1;
    if (Number(a.prequote_priority || 0) !== Number(b.prequote_priority || 0)) return Number(b.prequote_priority || 0) - Number(a.prequote_priority || 0);
    if (Number(a.sort_order || 0) !== Number(b.sort_order || 0)) return Number(a.sort_order || 0) - Number(b.sort_order || 0);
    return String(a.template_name || "").localeCompare(String(b.template_name || ""));
  });
  return out.slice(0, normalized.limit);
}

function listTemplateCatalogForPrequote(filters, options) {
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  var merged = Object.assign({}, filters || {}, options || {});
  return listTemplateCatalogForPrequoteFromSs_(getSpreadsheet_(), merged);
}

function getTemplateVersionItemsForPrequoteFromSs_(ss, templateId, version) {
  var spreadsheet = ss || getSpreadsheet_();
  var tpl = readTemplateCatalogRowByIdFromSs_(spreadsheet, templateId);
  if (!tpl) throw new Error("Template not found: " + String(templateId || "").trim());
  if (!templatePassesPrequoteFilters_(Object.assign({}, tpl, { _search_text: buildTemplateSearchText_(tpl, {}) }), {}, getPrequoteSettingsFromSs_(spreadsheet))) {
    throw new Error("Template is not exposed to prequote: " + String(templateId || "").trim());
  }
  var resolvedVersion = resolveTemplateVersionNumberFromSs_(spreadsheet, templateId, version);
  var row = getTemplateVersionRowFromSs_(spreadsheet, templateId, resolvedVersion);
  if (!row) throw new Error("Template version not found: " + String(templateId || "").trim() + " v" + resolvedVersion);
  var meta = pickTemplateVersionMeta_(row);
  var templateOut = cloneTemplateCatalogRow_(tpl);
  delete templateOut._search_text;
  return {
    template: templateOut,
    template_meta: cloneObj_(templateOut.template_meta || pickTemplateMeta_(templateOut)),
    version: Number(meta.version || resolvedVersion),
    created_at: meta.created_at,
    created_by: meta.created_by,
    source_quote_id: meta.source_quote_id,
    item_count: Number(meta.item_count || 0),
    summary: meta.summary || {},
    template_meta_snapshot: meta.template_meta_snapshot,
    items: normalizeTemplateItems_(parseJsonSafe_(row.items_json, []))
  };
}

function getTemplateVersionItemsForPrequote(templateId, version) {
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  return getTemplateVersionItemsForPrequoteFromSs_(getSpreadsheet_(), templateId, version);
}

function getTemplatePackageDetailForPrequoteFromSs_(ss, templateId, options) {
  var opts = options || {};
  var version = Number(opts.version || 0);
  return getTemplateVersionItemsForPrequoteFromSs_(ss, templateId, version);
}

function getTemplatePackageDetailForPrequote(templateId, options) {
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  return getTemplatePackageDetailForPrequoteFromSs_(getSpreadsheet_(), templateId, options || {});
}

function normText_(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function tokenize_(text) {
  var t = normText_(text);
  if (!t) return [];

  var parts = t.split(" ").filter(function(v) { return !!v; });
  var out = [];
  var seen = Object.create(null);

  for (var i = 0; i < parts.length; i++) {
    var p = String(parts[i] || "").trim();
    if (!p || p.length < 2) continue;

    if (!seen[p]) {
      out.push(p);
      seen[p] = 1;
    }

    var max = Math.min(p.length, 6);
    for (var k = 2; k <= max; k++) {
      var pre = p.slice(0, k);
      if (!seen[pre]) {
        out.push(pre);
        seen[pre] = 1;
      }
    }
  }
  return out;
}

function buildMaterialSearchText_(m) {
  return normText_([
    m && m.name,
    m && m.brand,
    m && m.spec,
    m && m.unit,
    m && m.note,
    m && m.material_type,
    m && m.trade_code,
    m && m.space_type,
    m && m.price_band,
    m && m.tags_summary,
    m && m.recommendation_note
  ].join(" "));
}

function normalizeMaterialType_(value) {
  return normalizeSimpleToken_(value);
}

function normalizeMaterialPriceBand_(value) {
  return normalizeSimpleToken_(value);
}

function normalizeMaterialSortOrder_(value) {
  var n = Math.floor(normalizeNumber_(value, 0));
  return n > 0 ? n : 0;
}

function normalizeRecommendationScoreBase_(value) {
  return normalizeNumber_(value, 0);
}

function normalizeTemplateType_(value) {
  var s = String(value || "").trim().toUpperCase();
  if (s === "QUOTE") s = PREQUOTE_TEMPLATE_TYPES_.QUOTE_ONLY;
  if (s === PREQUOTE_TEMPLATE_TYPES_.PREQUOTE_PACKAGE) return PREQUOTE_TEMPLATE_TYPES_.PREQUOTE_PACKAGE;
  if (s === PREQUOTE_TEMPLATE_TYPES_.BOTH) return PREQUOTE_TEMPLATE_TYPES_.BOTH;
  return PREQUOTE_TEMPLATE_TYPES_.QUOTE_ONLY;
}

function normalizeTemplateSortOrder_(value) {
  var n = Math.floor(normalizeNumber_(value, 0));
  return n > 0 ? n : 0;
}

function parseRefs_(refsStr) {
  var raw = String(refsStr || "").trim();
  if (!raw) return [];

  var parts = raw.split(",");
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var p = String(parts[i] || "").trim();
    if (!p) continue;
    var pair = p.split(":");
    var id = String(pair[0] || "").trim();
    var rowNo = Number(pair[1] || 0);
    if (!id || rowNo < 2) continue;
    out.push({ id: id, rowNo: rowNo });
  }
  return out;
}

function refsToString_(refsArr) {
  var arr = Array.isArray(refsArr) ? refsArr : [];
  var out = [];
  for (var i = 0; i < arr.length; i++) {
    out.push(String(arr[i].id) + ":" + String(arr[i].rowNo));
  }
  return out.join(",");
}

function removeMaterialFromIndex_(materialId) {
  var id = String(materialId || "").trim();
  if (!id) return;

  var sh = getSheet_("MaterialsIndex");
  var headers = getHeaders_("MaterialsIndex");
  var cols = getColMap_("MaterialsIndex");
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var changedRowNos = [];
  var changedRows = [];
  var now = nowIso_();

  for (var i = 0; i < values.length; i++) {
    var refs = parseRefs_(values[i][cols.refs]);
    var filtered = [];
    for (var r = 0; r < refs.length; r++) {
      if (refs[r].id !== id) filtered.push(refs[r]);
    }
    if (filtered.length === refs.length) continue;

    var row = values[i].slice();
    row[cols.refs] = refsToString_(filtered);
    row[cols.updated_at] = now;
    changedRowNos.push(i + 2);
    changedRows.push(row);
  }

  if (changedRowNos.length) {
    writeRowsByRowNos_(sh, changedRowNos, changedRows, 1, headers.length);
  }
}

function upsertMaterialIndex_(materialId, materialRowNo, materialObjOrRow) {
  var id = String(materialId || "").trim();
  var rowNo = Number(materialRowNo || 0);
  if (!id || rowNo < 2) return;

  var m = materialObjOrRow || {};
  var searchText = (m.search_text !== undefined)
    ? String(m.search_text || "")
    : buildMaterialSearchText_(m);
  var tokens = tokenize_(searchText);
  if (!tokens.length) return;

  var sh = getSheet_("MaterialsIndex");
  var headers = getHeaders_("MaterialsIndex");
  var cols = getColMap_("MaterialsIndex");
  var lastRow = sh.getLastRow();
  var now = nowIso_();

  var values = (lastRow >= 2) ? sh.getRange(2, 1, lastRow - 1, headers.length).getValues() : [];
  var tokenToPos = Object.create(null);

  for (var i = 0; i < values.length; i++) {
    var tok = String(values[i][cols.token] || "").trim();
    if (tok) tokenToPos[tok] = i;
  }

  var changedPosSet = Object.create(null);

  for (var v = 0; v < values.length; v++) {
    var refs0 = parseRefs_(values[v][cols.refs]);
    var filtered0 = [];
    for (var z = 0; z < refs0.length; z++) {
      if (refs0[z].id !== id) filtered0.push(refs0[z]);
    }
    if (filtered0.length !== refs0.length) {
      values[v][cols.refs] = refsToString_(filtered0);
      values[v][cols.updated_at] = now;
      changedPosSet[v] = 1;
    }
  }

  var refStr = id + ":" + rowNo;
  var appended = [];

  for (var t = 0; t < tokens.length; t++) {
    var token = tokens[t];
    var pos = tokenToPos[token];
    if (pos === undefined) {
      var row = Array(headers.length).fill("");
      row[cols.token] = token;
      row[cols.refs] = refStr;
      row[cols.updated_at] = now;
      appended.push(row);
      continue;
    }

    var refs = parseRefs_(values[pos][cols.refs]);
    var exists = false;
    for (var k = 0; k < refs.length; k++) {
      if (refs[k].id === id) {
        refs[k].rowNo = rowNo;
        exists = true;
      }
    }
    if (!exists) refs.unshift({ id: id, rowNo: rowNo });

    if (refs.length > 2000) refs = refs.slice(0, 2000);
    values[pos][cols.refs] = refsToString_(refs);
    values[pos][cols.updated_at] = now;
    changedPosSet[pos] = 1;
  }

  var changedRowNos = [];
  var changedRows = [];
  for (var key in changedPosSet) {
    var p = Number(key);
    changedRowNos.push(p + 2);
    changedRows.push(values[p]);
  }

  if (changedRowNos.length) {
    var pairs = [];
    for (var i2 = 0; i2 < changedRowNos.length; i2++) {
      pairs.push({ rowNo: changedRowNos[i2], row: changedRows[i2] });
    }
    pairs.sort(function(a, b) { return a.rowNo - b.rowNo; });
    var sortedNos = [];
    var sortedRows = [];
    for (var p2 = 0; p2 < pairs.length; p2++) {
      sortedNos.push(pairs[p2].rowNo);
      sortedRows.push(pairs[p2].row);
    }
    writeRowsByRowNos_(sh, sortedNos, sortedRows, 1, headers.length);
  }

  if (appended.length) {
    var start = sh.getLastRow() + 1;
    ensureRowCapacity_(sh, start + appended.length - 1);
    sh.getRange(start, 1, appended.length, headers.length).setValues(appended);
  }
}

function materialIndexReadyCacheKey_() {
  return "MATIDX_READY_" + String(getMaterialSearchVersion_());
}

function rebuildMaterialIndexAll_() {
  var shMat = getSheet_("Materials");
  var headersMat = getHeaders_("Materials");
  var colsMat = getColMap_("Materials");
  var lastRow = shMat.getLastRow();

  var tokenMap = Object.create(null);
  if (lastRow >= 2) {
    var values = shMat.getRange(2, 1, lastRow - 1, headersMat.length).getValues();
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var rowNo = i + 2;
      var id = String(row[colsMat.material_id] || "").trim();
      if (!id) continue;
      var searchText = String(row[colsMat.search_text] || "").trim();
      if (!searchText) {
        searchText = buildMaterialSearchText_({
          name: row[colsMat.name],
          brand: row[colsMat.brand],
          spec: row[colsMat.spec],
          unit: row[colsMat.unit],
          note: row[colsMat.note]
        });
      }
      var tokens = tokenize_(searchText);
      for (var t = 0; t < tokens.length; t++) {
        var tok = tokens[t];
        if (!tokenMap[tok]) tokenMap[tok] = [];
        tokenMap[tok].push({ id: id, rowNo: rowNo });
      }
    }
  }

  var shIdx = getSheet_("MaterialsIndex");
  var headersIdx = getHeaders_("MaterialsIndex");
  var colsIdx = getColMap_("MaterialsIndex");

  var idxLastRow = shIdx.getLastRow();
  if (idxLastRow > 1) shIdx.getRange(2, 1, idxLastRow - 1, headersIdx.length).clearContent();

  var tokensOut = Object.keys(tokenMap);
  tokensOut.sort();
  if (!tokensOut.length) return 0;

  var now = nowIso_();
  var rows = [];
  for (var i2 = 0; i2 < tokensOut.length; i2++) {
    var tok2 = tokensOut[i2];
    var refsArr = tokenMap[tok2] || [];
    refsArr.sort(function(a, b) { return Number(b.rowNo || 0) - Number(a.rowNo || 0); });
    if (refsArr.length > 2000) refsArr = refsArr.slice(0, 2000);
    var rowOut = Array(headersIdx.length).fill("");
    rowOut[colsIdx.token] = tok2;
    rowOut[colsIdx.refs] = refsToString_(refsArr);
    rowOut[colsIdx.updated_at] = now;
    rows.push(rowOut);
  }

  ensureRowCapacity_(shIdx, 1 + rows.length);
  var batch = 400;
  for (var s = 0; s < rows.length; s += batch) {
    var chunk = rows.slice(s, s + batch);
    shIdx.getRange(2 + s, 1, chunk.length, headersIdx.length).setValues(chunk);
  }
  return rows.length;
}

function ensureMaterialIndexReady_() {
  var key = materialIndexReadyCacheKey_();
  var cached = getCachedJson_(key);
  if (cached && cached.ok) return;

  withScriptLock_(10000, function() {
    var cached2 = getCachedJson_(key);
    if (cached2 && cached2.ok) return;

    var shIdx = getSheet_("MaterialsIndex");
    if (Number(shIdx.getLastRow() || 0) >= 2) {
      putCachedJson_(key, { ok: true, row_count: Number(shIdx.getLastRow() || 0) - 1 }, 1800);
      return;
    }

    var built = rebuildMaterialIndexAll_();
    putCachedJson_(key, { ok: true, row_count: built }, 1800);
  });
}

function getMaterialIndexMap_() {
  var ver = getMaterialSearchVersion_();
  if (!__MATERIAL_INDEX_MEM_MAP_ || __MATERIAL_INDEX_MEM_VER_ !== ver) {
    __MATERIAL_INDEX_MEM_VER_ = ver;
    __MATERIAL_INDEX_MEM_MAP_ = Object.create(null);
  }
  return __MATERIAL_INDEX_MEM_MAP_;
}

function getMaterialRefsByTokens_(tokens) {
  var list = Array.isArray(tokens) ? tokens : [];
  var out = Object.create(null);
  if (!list.length) return out;

  var ver = getMaterialSearchVersion_();
  var memMap = getMaterialIndexMap_();
  var cache = CacheService.getScriptCache();
  var missing = [];

  for (var i = 0; i < list.length; i++) {
    var token = String(list[i] || "").trim();
    if (!token) continue;

    if (Object.prototype.hasOwnProperty.call(memMap, token)) {
      if (memMap[token]) out[token] = memMap[token];
      continue;
    }

    var key = "MIDXTOK_" + ver + "_" + token;
    var cached = null;
    try { cached = cache.get(key); } catch (e) {}
    if (cached !== null && cached !== undefined) {
      var cachedRef = (cached && cached !== "__MISS__") ? String(cached) : "";
      memMap[token] = cachedRef;
      if (cachedRef) out[token] = cachedRef;
      continue;
    }
    missing.push(token);
  }

  if (!missing.length) return out;

  var sh = getSheet_("MaterialsIndex");
  var cols = getColMap_("MaterialsIndex");
  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    for (var m0 = 0; m0 < missing.length; m0++) {
      memMap[missing[m0]] = "";
      try { cache.put("MIDXTOK_" + ver + "_" + missing[m0], "__MISS__", 300); } catch (e0) {}
    }
    return out;
  }

  var tokenRange = sh.getRange(2, cols.token + 1, lastRow - 1, 1);
  var tokenToRowNo = Object.create(null);
  for (var m = 0; m < missing.length; m++) {
    var tok = missing[m];
    var finder = tokenRange.createTextFinder(tok);
    finder.matchEntireCell(true);
    var hit = finder.findNext();
    if (hit) tokenToRowNo[tok] = Number(hit.getRow() || 0);
  }

  var rowNos = [];
  for (var t = 0; t < missing.length; t++) {
    var rn = Number(tokenToRowNo[missing[t]] || 0);
    if (rn >= 2) rowNos.push(rn);
  }
  rowNos = uniqueSortedRowNos_(rowNos);

  var refsByRowNo = Object.create(null);
  if (rowNos.length) {
    var refsRows = readRowsMapByRowNos_(sh, rowNos, cols.refs + 1, 1);
    for (var r = 0; r < rowNos.length; r++) {
      var rowNo = rowNos[r];
      var row = refsRows[rowNo];
      refsByRowNo[rowNo] = row ? String(row[0] || "") : "";
    }
  }

  for (var m2 = 0; m2 < missing.length; m2++) {
    var tok2 = missing[m2];
    var rn2 = Number(tokenToRowNo[tok2] || 0);
    var refsStr = (rn2 >= 2) ? String(refsByRowNo[rn2] || "") : "";
    memMap[tok2] = refsStr;
    if (refsStr) out[tok2] = refsStr;
    try { cache.put("MIDXTOK_" + ver + "_" + tok2, refsStr || "__MISS__", 600); } catch (e2) {}
  }

  return out;
}

function getMaterialsByRefs_(refs, limit) {
  var sh = getSheet_("Materials");
  var headers = getHeaders_("Materials");
  var cols = getColMap_("Materials");
  var max = Math.max(Number(limit || 50), 1);

  var orderedRefs = Array.isArray(refs) ? refs : [];
  if (!orderedRefs.length) return [];

  var cachedById = Object.create(null);
  var cacheKeys = [];
  for (var i = 0; i < orderedRefs.length; i++) {
    var refId = String(orderedRefs[i] && orderedRefs[i].id || "").trim();
    if (!refId || cachedById[refId]) continue;
    cachedById[refId] = null;
    cacheKeys.push("MATROW_" + refId);
  }

  if (cacheKeys.length) {
    try {
      var hitMap = CacheService.getScriptCache().getAll(cacheKeys);
      for (var ck = 0; ck < cacheKeys.length; ck++) {
        var key = cacheKeys[ck];
        var raw = hitMap && hitMap[key];
        if (!raw) continue;
        var parsed = parseJsonSafe_(raw, null);
        if (!parsed || typeof parsed !== "object") continue;
        var pid = String(parsed.material_id || "").trim();
        if (!pid) continue;
        cachedById[pid] = parsed;
      }
    } catch (e0) {}
  }

  var rowNosRaw = [];
  for (var r0 = 0; r0 < orderedRefs.length; r0++) {
    var ref0 = orderedRefs[r0] || {};
    var id0 = String(ref0.id || "").trim();
    var rn0 = Number(ref0.rowNo || 0);
    if (!id0 || rn0 < 2) continue;
    if (cachedById[id0]) continue;
    rowNosRaw.push(rn0);
  }
  var rowNos = uniqueSortedRowNos_(rowNosRaw);
  var rowsByNo = rowNos.length ? readRowsMapByRowNos_(sh, rowNos, 1, headers.length) : Object.create(null);

  var out = [];
  var seenId = Object.create(null);

  for (var r = 0; r < orderedRefs.length && out.length < max; r++) {
    var ref = orderedRefs[r] || {};
    var refId = String(ref.id || "").trim();
    var rowNo = Number(ref.rowNo || 0);
    if (!refId || rowNo < 2) continue;
    if (seenId[refId]) continue;

    var picked = cachedById[refId];
    if (!picked) {
      var row = rowsByNo[rowNo];
      if (!row) continue;
      if (String(row[cols.material_id] || "").trim() !== refId) continue;
      var obj = rowToObj_(headers, row, rowNo);
      picked = pickMaterial_(obj);
      try { CacheService.getScriptCache().put("MATROW_" + refId, JSON.stringify(picked), 600); } catch (e1) {}
    }
    seenId[refId] = 1;
    out.push(picked);
  }

  return out;
}

function searchMaterials(query, adminPassword, limit) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var spreadsheet = getSpreadsheet_();
  var opts = (limit && typeof limit === "object") ? cloneObj_(limit) : { limit: limit };
  var q = normText_(String(query || "").trim());
  var minQueryLen = Math.max(Number(opts.min_query_len || opts.minQueryLen || 2), 1);
  if (!q || q.length < minQueryLen) return [];

  var max = Math.min(Math.max(Number(opts.limit || 50), 1), 200);
  var page = Math.max(Number(opts.page || 1), 1);
  var offset = (opts.offset !== undefined && opts.offset !== null)
    ? Math.max(Math.floor(Number(opts.offset || 0)), 0)
    : Math.max((page - 1) * max, 0);

  var tokens = pickSearchTokens_(tokenize_(q), 8);
  if (!tokens.length) return [];

  var qCacheKey = materialSearchQueryCacheKey_(q, max, offset);
  var out = getOrBuildCachedJsonWithLock_(
    qCacheKey,
    120,
    function() {
      ensureMaterialIndexReady_();
      var refsByToken = getMaterialRefsByTokens_(tokens);
      if (!refsByToken || typeof refsByToken !== "object") {
        return searchMaterialsFallbackFullScan_(q, offset + max).slice(offset, offset + max);
      }

      var refMap = Object.create(null);
      for (var t = 0; t < tokens.length; t++) {
        var refsStr = refsByToken[tokens[t]];
        if (!refsStr) continue;
        var refs = parseRefs_(refsStr);
        for (var r = 0; r < refs.length; r++) {
          var id = refs[r].id;
          if (!refMap[id] || Number(refs[r].rowNo || 0) > Number(refMap[id].rowNo || 0)) {
            refMap[id] = refs[r];
          }
        }
      }

      var refsOut = [];
      for (var key in refMap) refsOut.push(refMap[key]);
      if (!refsOut.length) {
        return searchMaterialsFallbackFullScan_(q, offset + max).slice(offset, offset + max);
      }

      refsOut.sort(function(a, b) { return Number(b.rowNo || 0) - Number(a.rowNo || 0); });

      var fetchLimit = Math.max((offset + max) * 3, max * 3);
      var pre = getMaterialsByRefs_(refsOut, fetchLimit);
      var filtered = [];
      for (var p = 0; p < pre.length; p++) {
        var m = pre[p];
        var hay = normText_([
          m.name, m.brand, m.spec, m.note,
          m.material_type, m.trade_code, m.space_type, m.price_band, m.tags_summary, m.recommendation_note
        ].join(" "));
        var matched = true;
        for (var t2 = 0; t2 < tokens.length; t2++) {
          if (hay.indexOf(tokens[t2]) < 0) {
            matched = false;
            break;
          }
        }
        if (!matched) continue;
        filtered.push(m);
      }

      filtered.sort(function(a, b) {
        return String(b.updated_at || "").localeCompare(String(a.updated_at || ""));
      });
      return decorateMaterialListForAdmin_(spreadsheet, filtered.slice(offset, offset + max));
    },
    { wait_ms: 1400, retry_ms: 80, fallback_value: [] }
  );
  return Array.isArray(out) ? out.slice(0, max) : [];
}

function searchMaterialsFallbackFullScan_(q, limit) {
  var sh = getSheet_("Materials");
  var headers = getHeaders_("Materials");
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var out = [];
  var qq = normText_(q);
  var tokens = pickSearchTokens_(tokenize_(qq), 8);

  for (var i = 0; i < values.length; i++) {
    var row = rowToObj_(headers, values[i], i + 2);
    if (!row.material_id) continue;

    var hay = normText_(String(row.search_text || [row.name, row.brand, row.spec, row.note].join(" ")));
    var matched = false;
    if (tokens.length) {
      matched = true;
      for (var t = 0; t < tokens.length; t++) {
        if (hay.indexOf(tokens[t]) < 0) {
          matched = false;
          break;
        }
      }
    } else {
      matched = hay.indexOf(qq) >= 0;
    }
    if (matched) out.push(pickMaterial_(row));
  }

  out.sort(function(a, b) {
    return String(b.updated_at || "").localeCompare(String(a.updated_at || ""));
  });
  return decorateMaterialListForAdmin_(getSpreadsheet_(), out.slice(0, Number(limit || 50)));
}

function saveMaterial(material, dataUrl, filename, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var m = material || {};
  var hasOwn = Object.prototype.hasOwnProperty;
  var sh = getSheet_("Materials");
  var headers = getHeaders_("Materials");
  var cols = getColMap_("Materials");
  var now = nowIso_();

  var id = String(m.material_id || "").trim() || uuid_();
  var imageFileId = String(m.image_file_id || "").trim();
  var imageFileName = String(m.image_file_name || "").trim();

  if (dataUrl) {
    var parsed = parseDataUrl_(dataUrl);
    var bytes0 = Utilities.base64Decode(parsed.base64);
    var bytes = tryResizeImageBytes_(bytes0, parsed.mimeType, 900);
    var folder = getOrCreateMaterialsFolder_();
    var safeName = sanitizeFilename_(filename || "material.jpg");
    var blob = Utilities.newBlob(bytes, parsed.mimeType, safeName);
    var file = folder.createFile(blob);
    file.setName("material_" + id + "_" + safeName);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    imageFileId = file.getId();
    imageFileName = file.getName();
  }

  var rowNo = findRowNoByColumnValue_(sh, cols.material_id + 1, id);
  var row;
  var prevImageFileId = "";

  if (rowNo > 0) {
    row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
    prevImageFileId = String(row[cols.image_file_id] || "").trim();
    if (!imageFileId && !dataUrl) imageFileId = prevImageFileId;
    if (!imageFileName && !dataUrl) imageFileName = String(row[cols.image_file_name] || "").trim();
  } else {
    row = Array(headers.length).fill("");
    row[cols.material_id] = id;
    row[cols.created_at] = now;
  }

  row[cols.material_id] = id;
  row[cols.updated_at] = now;
  row[cols.name] = String(m.name || "");
  row[cols.brand] = String(m.brand || "");
  row[cols.spec] = String(m.spec || "");
  row[cols.unit] = String(m.unit || "");
  row[cols.unit_price] = Number(m.unit_price || 0);
  row[cols.note] = String(m.note || "");
  setByHeader_(headers, row, "is_active",
    hasOwn.call(m, "is_active") ? boolToYn_(ynToBool_(m.is_active, true)) : String(getByHeader_(headers, row, "is_active") || "Y"));
  setByHeader_(headers, row, "is_representative",
    hasOwn.call(m, "is_representative") ? boolToYn_(ynToBool_(m.is_representative, false)) : String(getByHeader_(headers, row, "is_representative") || "N"));
  setByHeader_(headers, row, "material_type",
    hasOwn.call(m, "material_type") ? normalizeMaterialType_(m.material_type) : String(getByHeader_(headers, row, "material_type") || "").trim());
  setByHeader_(headers, row, "trade_code",
    hasOwn.call(m, "trade_code") ? normalizeTradeCode_(m.trade_code) : String(getByHeader_(headers, row, "trade_code") || "").trim());
  setByHeader_(headers, row, "space_type",
    hasOwn.call(m, "space_type") ? normalizeMaterialGroupSpaceType_(m.space_type) : normalizeMaterialGroupSpaceType_(getByHeader_(headers, row, "space_type") || "BOTH"));
  setByHeader_(headers, row, "sort_order",
    hasOwn.call(m, "sort_order") ? normalizeMaterialSortOrder_(m.sort_order) : normalizeMaterialSortOrder_(getByHeader_(headers, row, "sort_order")));
  setByHeader_(headers, row, "expose_to_prequote",
    hasOwn.call(m, "expose_to_prequote") ? boolToYn_(ynToBool_(m.expose_to_prequote, false)) : String(getByHeader_(headers, row, "expose_to_prequote") || "N"));
  setByHeader_(headers, row, "recommendation_score_base",
    hasOwn.call(m, "recommendation_score_base") ? normalizeRecommendationScoreBase_(m.recommendation_score_base) : normalizeRecommendationScoreBase_(getByHeader_(headers, row, "recommendation_score_base")));
  setByHeader_(headers, row, "price_band",
    hasOwn.call(m, "price_band") ? normalizeMaterialPriceBand_(m.price_band) : String(getByHeader_(headers, row, "price_band") || "").trim());
  setByHeader_(headers, row, "tags_summary",
    hasOwn.call(m, "tags_summary") ? String(m.tags_summary || "").trim() : String(getByHeader_(headers, row, "tags_summary") || "").trim());
  setByHeader_(headers, row, "recommendation_note",
    hasOwn.call(m, "recommendation_note") ? String(m.recommendation_note || "").trim() : String(getByHeader_(headers, row, "recommendation_note") || "").trim());
  row[cols.image_file_id] = imageFileId;
  row[cols.image_file_name] = imageFileName;
  row[cols.search_text] = buildMaterialSearchText_(rowToObj_(headers, row, rowNo || 0));

  if (rowNo > 0) {
    sh.getRange(rowNo, 1, 1, headers.length).setValues([row]);
  } else {
    rowNo = sh.getLastRow() + 1;
    ensureRowCapacity_(sh, rowNo);
    sh.getRange(rowNo, 1, 1, headers.length).setValues([row]);
  }

  if (prevImageFileId && prevImageFileId !== imageFileId) {
    clearMaterialImageAllowCache_(prevImageFileId);
  }
  if (imageFileId) {
    setImageRefCache_(imageFileId, true, 3600);
    setMaterialImageAllowCache_(imageFileId, true, 1800);
  }
  try {
    upsertMaterialIndex_(id, rowNo, {
      name: m.name,
      brand: m.brand,
      spec: m.spec,
      unit: m.unit,
      note: m.note,
      search_text: row[cols.search_text]
    });
  } catch (e) {}

  bumpMaterialSearchVersion_();
  bumpMaterialTagListVersion_();
  try { CacheService.getScriptCache().remove("MATROW_" + id); } catch (e) {}
  return { material: pickMaterial_(rowToObj_(headers, row, rowNo)) };
}

function deleteMaterial(materialId, deleteImage, adminPassword) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();

  var target = String(materialId || "").trim();
  if (!target) throw new Error("materialId required");

  var sh = getSheet_("Materials");
  var headers = getHeaders_("Materials");
  var cols = getColMap_("Materials");
  var rowNo = findRowNoByColumnValue_(sh, cols.material_id + 1, target);
  if (rowNo < 2) throw new Error("Material not found: " + target);

  var row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
  var imgId = String(row[cols.image_file_id] || "").trim();
  if (imgId) setImageRefCache_(imgId, false, 120);
  if (imgId) clearMaterialImageAllowCache_(imgId);
  if (deleteImage && imgId) {
    try { DriveApp.getFileById(imgId).setTrashed(true); } catch (e) {}
  }
  try { replaceMaterialTagsInSs_(getSpreadsheet_(), target, []); } catch (e0) {}

  sh.deleteRow(rowNo);
  try { removeMaterialFromIndex_(target); } catch (e) {}
  bumpMaterialSearchVersion_();
  bumpMaterialTagListVersion_();
  try { CacheService.getScriptCache().remove("MATROW_" + target); } catch (e) {}
  return { ok: true };
}

function pickMaterial_(row) {
  var src = row || {};
  return {
    material_id: src.material_id,
    name: src.name || "",
    brand: src.brand || "",
    spec: src.spec || "",
    unit: src.unit || "",
    unit_price: Number(src.unit_price || 0),
    image_file_id: src.image_file_id || "",
    image_file_name: src.image_file_name || "",
    note: src.note || "",
    updated_at: src.updated_at || "",
    is_active: ynToBool_(src.is_active, true),
    is_representative: ynToBool_(src.is_representative, false),
    material_type: String(src.material_type || "").trim(),
    trade_code: normalizeTradeCode_(src.trade_code),
    space_type: normalizeMaterialGroupSpaceType_(src.space_type),
    sort_order: normalizeMaterialSortOrder_(src.sort_order),
    expose_to_prequote: ynToBool_(src.expose_to_prequote, false),
    recommendation_score_base: normalizeRecommendationScoreBase_(src.recommendation_score_base),
    price_band: String(src.price_band || "").trim(),
    tags_summary: String(src.tags_summary || "").trim(),
    recommendation_note: String(src.recommendation_note || "").trim()
  };
}

function getOwnerEmailTargets_() {
  var s = getSettings_();
  var raw = [
    s.owner_email,
    s.ownerEmail,
    s.notify_email,
    s.notification_email,
    s.admin_email
  ].map(function(v) { return String(v || "").trim(); }).find(function(v) { return !!v; }) || "";

  if (!raw) return [];
  return raw
    .split(/[,\s;]+/)
    .map(function(v) { return String(v || "").trim(); })
    .filter(function(v) { return isValidEmail_(v); });
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || "").trim());
}

function sendEmailSafe_(to, subject, body) {
  var list = Array.isArray(to) ? to : [to];
  var targets = list
    .map(function(v) { return String(v || "").trim(); })
    .filter(function(v) { return isValidEmail_(v); });
  if (!targets.length) throw new Error("owner_email is missing or invalid");

  var toStr = targets.join(",");
  var sub = String(subject || "");
  var txt = String(body || "");
  var firstErr = null;

  try {
    MailApp.sendEmail(toStr, sub, txt);
    return { ok: true, provider: "MailApp", recipients: targets };
  } catch (e) {
    firstErr = e;
  }

  try {
    GmailApp.sendEmail(toStr, sub, txt);
    return { ok: true, provider: "GmailApp", recipients: targets };
  } catch (e2) {
    var m1 = firstErr && firstErr.message ? firstErr.message : String(firstErr || "");
    var m2 = e2 && e2.message ? e2.message : String(e2 || "");
    throw new Error("Email send failed. MailApp=" + m1 + " / GmailApp=" + m2);
  }
}

function notifyOwnerApproval_(quoteId) {
  var owners = getOwnerEmailTargets_();
  if (!owners.length) return;

  var q = findQuoteRow_(quoteId);
  if (!q) return;

  var s = getSettings_();
  var urls = q.share_token
    ? buildShareUrls_(quoteId, String(q.share_token))
    : { viewUrl: "", printUrl: "" };

  var lines = [
    "Quote approved: " + (q.customer_name || "") + " / " + (q.site_name || ""),
    "Contact: " + (q.contact_name || "") + " / " + (q.contact_phone || ""),
    "Total: " + Number(q.total || 0).toLocaleString("ko-KR") + " KRW",
    q.deposit_amount ? ("Deposit: " + Number(q.deposit_amount || 0).toLocaleString("ko-KR") + " KRW") : "",
    q.deposit_due_at ? ("Due: " + String(q.deposit_due_at || "")) : "",
    s.account_info ? ("Account: " + String(s.account_info || "")) : "",
    "",
    urls.viewUrl ? ("View: " + urls.viewUrl) : "",
    urls.printUrl ? ("Print: " + urls.printUrl) : ""
  ].filter(function(v) { return !!v; }).join("\n");

  sendEmailSafe_(owners, "[QuoteApp] Approved: " + (q.customer_name || "") + "/" + (q.site_name || ""), lines);
}

function notifyOwnerCustomerNote_(quoteId, shortNote) {
  var q = findQuoteRow_(quoteId);
  if (!q) return;
  notifyOwnerCustomerNoteFromQuote_(q, quoteId, shortNote);
}

function notifyOwnerCustomerNoteFromQuote_(quoteObj, quoteId, shortNote) {
  var owners = getOwnerEmailTargets_();
  if (!owners.length) return;

  var q = quoteObj || {};

  var urls = q.share_token ? buildShareUrls_(quoteId, String(q.share_token)) : { viewUrl: "" };
  var body = [
    "Customer note received: " + (q.customer_name || "") + " / " + (q.site_name || ""),
    "Summary: " + (shortNote || ""),
    "",
    urls.viewUrl ? ("View: " + urls.viewUrl) : ""
  ].filter(function(v) { return !!v; }).join("\n");

  sendEmailSafe_(owners, "[QuoteApp] Customer note: " + (q.customer_name || "") + "/" + (q.site_name || ""), body);
}

function logQuoteView_(quoteRowObj, page) {
  if (!quoteRowObj || !quoteRowObj.quote_id || !quoteRowObj._rowNo) return;

  ensureCoreSchemaReady_();

  var shQ = getSheet_("Quotes");
  var headers = getHeaders_("Quotes");
  var cols = getColMap_("Quotes");
  var rowNo = Number(quoteRowObj._rowNo);
  if (!rowNo || rowNo < 2) return;

  var nowIso = nowIso_();
  var followMin = Number(getSettings_().followup_after_view_minutes || 120);

  var lock = LockService.getScriptLock();
  var locked = false;
  try { locked = lock.tryLock(80); } catch (e) { locked = false; }
  if (locked) {
    try {
      var row = shQ.getRange(rowNo, 1, 1, headers.length).getValues()[0];
      row[cols.view_count] = Number(row[cols.view_count] || 0) + 1;
      row[cols.last_viewed_at] = nowIso;

      var st = String(row[cols.status] || "").toLowerCase();
      if (st === "sent") {
        var newNextMs = Date.now() + followMin * 60000;
        var curNextMs = parseIsoMs_(row[cols.next_action_at]);
        if (isNaN(curNextMs) || curNextMs > newNextMs) {
          row[cols.next_action_at] = new Date(newNextMs).toISOString();
        }
      }
      shQ.getRange(rowNo, 1, 1, headers.length).setValues([row]);
    } finally {
      lock.releaseLock();
    }
  }

  getSheet_("ViewLog").appendRow([uuid_(), String(quoteRowObj.quote_id), String(page || "view"), nowIso]);
}

function buildShareUrls_(quoteId, shareToken) {
  var s = getSettings_();
  return buildShareUrlsWithBase_(getConfiguredBaseUrl_(s), quoteId, shareToken);
}

function buildShareUrlsWithBase_(baseUrl, quoteId, shareToken) {
  var base = String(baseUrl || "").trim();
  if (!base) return { viewUrl: "", printUrl: "" };
  return {
    viewUrl: base + "?page=view&quoteId=" + encodeURIComponent(quoteId) + "&token=" + encodeURIComponent(shareToken),
    printUrl: base + "?page=print&quoteId=" + encodeURIComponent(quoteId) + "&token=" + encodeURIComponent(shareToken)
  };
}

function getHeaders_(sheetName) {
  if (__HEADERS_CACHE_[sheetName]) return __HEADERS_CACHE_[sheetName].slice();
  var sh = getSheet_(sheetName);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  __HEADERS_CACHE_[sheetName] = headers;
  delete __COLMAP_CACHE_[sheetName];
  return headers.slice();
}

function getColMap_(sheetName) {
  if (__COLMAP_CACHE_[sheetName]) return __COLMAP_CACHE_[sheetName];
  var headers = getHeaders_(sheetName);
  var map = Object.create(null);
  for (var i = 0; i < headers.length; i++) {
    var key = String(headers[i] || "").trim();
    if (!key) continue;
    if (!Object.prototype.hasOwnProperty.call(map, key)) map[key] = i;
  }
  __COLMAP_CACHE_[sheetName] = map;
  return map;
}

function invalidateSheetMetaCache_(sheetName) {
  delete __HEADERS_CACHE_[sheetName];
  delete __COLMAP_CACHE_[sheetName];
  if (sheetName === "Quotes") invalidateQuoteRowMap_();
  if (sheetName === "QuoteItemsIndex") __QUOTE_ITEMS_INDEX_ROWNO_CACHE_ = null;
  invalidateSheetCache_(sheetName);
}

function colIndex_(headers, name, sheetName) {
  var idx = headers.indexOf(name);
  if (idx < 0) throw new Error(sheetName + " missing column: " + name);
  return idx;
}

function rowToObj_(headers, row, rowNo) {
  var obj = { _rowNo: rowNo };
  for (var i = 0; i < headers.length; i++) obj[headers[i]] = row[i];
  return obj;
}

function setByHeader_(headers, row, name, value) {
  var idx = headers.indexOf(name);
  if (idx >= 0) row[idx] = value;
}

function getByHeader_(headers, row, name) {
  var idx = headers.indexOf(name);
  return idx >= 0 ? row[idx] : "";
}

function getQuoteTable_() {
  var sh = getSheet_("Quotes");
  var headers = getHeaders_("Quotes");
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { headers: headers, values: [] };
  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return { headers: headers, values: values };
}

function quoteRowCacheKey_(quoteId) {
  return "QROW_" + String(quoteId || "").trim();
}

function getCachedQuoteRow_(quoteId) {
  var target = String(quoteId || "").trim();
  if (!target) return null;
  if (__QUOTE_ROW_CACHE_[target]) return cloneObj_(__QUOTE_ROW_CACHE_[target]);
  var cached = getCachedJson_(quoteRowCacheKey_(target));
  if (!cached || typeof cached !== "object") return null;
  __QUOTE_ROW_CACHE_[target] = cloneObj_(cached);
  return cloneObj_(cached);
}

function setCachedQuoteRow_(quoteObj, ttlSec) {
  var q = quoteObj || {};
  var quoteId = String(q.quote_id || "").trim();
  if (!quoteId) return;
  __QUOTE_ROW_CACHE_[quoteId] = cloneObj_(q);
  putCachedJson_(quoteRowCacheKey_(quoteId), q, Number(ttlSec || 300));
}

function invalidateCachedQuoteRow_(quoteId) {
  var target = String(quoteId || "").trim();
  if (!target) return;
  delete __QUOTE_ROW_CACHE_[target];
  removeCacheKey_(quoteRowCacheKey_(target));
}

function getQuoteRowMap_() {
  if (__QUOTE_ROWNO_CACHE_) return __QUOTE_ROWNO_CACHE_;

  var sh = getSheet_("Quotes");
  var cols = getColMap_("Quotes");
  var lastRow = sh.getLastRow();
  var map = Object.create(null);

  if (lastRow >= 2) {
    var qVals = sh.getRange(2, cols.quote_id + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < qVals.length; i++) {
      var id = String(qVals[i][0] || "").trim();
      if (!id) continue;
      map[id] = i + 2;
    }
  }

  __QUOTE_ROWNO_CACHE_ = map;
  return map;
}

function invalidateQuoteRowMap_() {
  __QUOTE_ROWNO_CACHE_ = null;
  __QUOTE_ROW_CACHE_ = Object.create(null);
}

function findQuoteRowNo_(quoteId) {
  var target = String(quoteId || "").trim();
  if (!target) return 0;
  return Number(getQuoteRowMap_()[target] || 0);
}

function findQuoteRow_(quoteId) {
  var target = String(quoteId || "").trim();
  if (!target) return null;

  var cached = getCachedQuoteRow_(target);
  if (cached) return cached;

  var rowNo = findQuoteRowNo_(target);
  if (!rowNo) return null;

  var sh = getSheet_("Quotes");
  var headers = getHeaders_("Quotes");
  var cols = getColMap_("Quotes");
  var row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
  var loadedId = String(row[cols.quote_id] || "").trim();
  if (loadedId !== target) {
    invalidateQuoteRowMap_();
    rowNo = findQuoteRowNo_(target);
    if (!rowNo) return null;
    row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
    loadedId = String(row[cols.quote_id] || "").trim();
    if (loadedId !== target) return null;
  }

  var out = rowToObj_(headers, row, rowNo);
  setCachedQuoteRow_(out, 300);
  return out;
}

function updateQuote_(quoteId, patch) {
  var target = String(quoteId || "").trim();
  if (!target) throw new Error("quoteId required");

  var rowNo = findQuoteRowNo_(target);
  if (!rowNo) throw new Error("Quote not found");

  var sh = getSheet_("Quotes");
  var headers = getHeaders_("Quotes");
  var cols = getColMap_("Quotes");
  var row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];

  var q = patch || {};
  for (var key in q) {
    if (!Object.prototype.hasOwnProperty.call(cols, key)) continue;
    row[cols[key]] = q[key];
  }
  row[cols.quote_id] = target;
  row[cols.updated_at] = nowIso_();

  sh.getRange(rowNo, 1, 1, headers.length).setValues([row]);
  var out = rowToObj_(headers, row, rowNo);
  setCachedQuoteRow_(out, 600);
  return out;
}

function setQuoteToken_(quoteId, token) {
  return updateQuote_(quoteId, { share_token: token });
}

function replaceItems_(quoteId, items) {
  var target = String(quoteId || "").trim();
  if (!target) throw new Error("quoteId required");
  var normalizedItems = normalizeItems_(items);

  var sh = getSheet_("Items");
  var headers = getHeaders_("Items");
  var cols = getColMap_("Items");
  var qCol = cols.quote_id + 1;

  var existingRowNos = findRowNosByColumnValue_(sh, qCol, target);
  var newRows = [];
  var outRowNos = [];
  var outItemIds = [];

  for (var i = 0; i < normalizedItems.length; i++) {
    newRows.push(buildItemSheetRow_(headers, cols, target, normalizedItems[i], i + 1));
  }

  var writeCount = Math.min(existingRowNos.length, newRows.length);
  if (writeCount > 0) {
    writeRowsByRowNos_(
      sh,
      existingRowNos.slice(0, writeCount),
      newRows.slice(0, writeCount),
      1,
      headers.length
    );
    for (var w = 0; w < writeCount; w++) {
      outRowNos.push(existingRowNos[w]);
      outItemIds.push(String(newRows[w][cols.item_id] || ""));
    }
  }

  if (existingRowNos.length > writeCount) {
    clearRowsByRowNos_(
      sh,
      existingRowNos.slice(writeCount),
      1,
      headers.length
    );
  }

  if (newRows.length > writeCount) {
    var appendRows = newRows.slice(writeCount);
    var startRow = sh.getLastRow() + 1;
    ensureRowCapacity_(sh, startRow + appendRows.length - 1);
    sh.getRange(startRow, 1, appendRows.length, headers.length).setValues(appendRows);
    for (var a = 0; a < appendRows.length; a++) {
      outRowNos.push(startRow + a);
      outItemIds.push(String(appendRows[a][cols.item_id] || ""));
    }
  }

  upsertQuoteItemsIndex_(target, outRowNos, outItemIds, nowIso_());
  return { rowNos: outRowNos, itemIds: outItemIds };
}

function normalizeNumberOrZero_(v) {
  var n = Number(v || 0);
  return isNaN(n) ? 0 : n;
}

function normalizePriceType_(v) {
  var s = String(v || "NORMAL").trim().toUpperCase();
  if (s === "OPTION" || s === "INCLUDED" || s === "SERVICE") return s;
  return "NORMAL";
}

function slugifyGroupId_(label) {
  var s = String(label || "").trim().toLowerCase();
  if (!s) return "misc";
  s = s.replace(/[^a-z0-9가-힣]+/g, "-").replace(/^-+|-+$/g, "");
  if (!s) return "misc";
  return s.slice(0, 80);
}

function normalizeItems_(items) {
  var list = Array.isArray(items) ? items : [];
  var out = [];
  var groupOrderMap = Object.create(null);
  var nextGroupOrder = 1;

  for (var i = 0; i < list.length; i++) {
    var it = list[i] || {};
    var seq = normalizeNumberOrZero_(it.seq || (i + 1));
    if (!seq) seq = i + 1;

    var rawName = String((it.name !== undefined && it.name !== null) ? it.name : (it.material || "")).trim();
    var rawDetail = String((it.detail !== undefined && it.detail !== null) ? it.detail : (it.spec || "")).trim();
    var rawLocation = String(it.location || "").trim();
    var rawNote = String(it.note || "").trim();

    var process = String(it.process || "").trim();
    var material = String(it.material || "").trim();
    var spec = String(it.spec || "").trim();

    var groupLabel = String(it.groupLabel || it.group_label || process || "").trim();
    if (!groupLabel) groupLabel = "기타";
    var groupId = String(it.groupId || it.group_id || "").trim();
    if (!groupId) groupId = slugifyGroupId_(groupLabel);
    var groupCode = String((it.groupCode !== undefined && it.groupCode !== null) ? it.groupCode : (it.group_code || "")).trim();

    var groupOrderRaw = Number((it.groupOrder !== undefined && it.groupOrder !== null) ? it.groupOrder : it.group_order);
    var groupOrder = isNaN(groupOrderRaw) ? 0 : groupOrderRaw;
    if (groupOrder <= 0) {
      if (!groupOrderMap[groupId]) {
        groupOrderMap[groupId] = nextGroupOrder++;
      }
      groupOrder = groupOrderMap[groupId];
    } else {
      groupOrderMap[groupId] = groupOrder;
      if (groupOrder >= nextGroupOrder) nextGroupOrder = groupOrder + 1;
    }

    var itemOrderRaw = Number((it.itemOrder !== undefined && it.itemOrder !== null) ? it.itemOrder : it.item_order);
    var itemOrder = isNaN(itemOrderRaw) || itemOrderRaw <= 0 ? seq : itemOrderRaw;

    var qty = normalizeNumberOrZero_(it.qty);
    var unitPrice = normalizeNumberOrZero_(it.unit_price !== undefined ? it.unit_price : it.unitPrice);
    var amountRaw = (it.amount !== undefined && it.amount !== null && String(it.amount) !== "")
      ? Number(it.amount)
      : (qty * unitPrice);
    var amount = isNaN(amountRaw) ? (qty * unitPrice) : amountRaw;

    var priceRaw = (it.price !== undefined && it.price !== null && String(it.price) !== "")
      ? Number(it.price)
      : amount;
    var price = isNaN(priceRaw) ? amount : priceRaw;
    var priceType = normalizePriceType_(it.priceType || it.price_type);

    var normalized = {
      quote_id: String(it.quote_id || it.quoteId || "").trim(),
      item_id: String(it.item_id || "").trim() || uuid_(),
      seq: seq,
      process: process || groupLabel,
      material: material || rawName,
      spec: spec || rawDetail,
      qty: qty || 1,
      unit: String(it.unit || "").trim() || "",
      unit_price: unitPrice || price,
      amount: amount || price,
      note: rawNote,
      material_ref_id: String(it.material_ref_id || it.materialRefId || "").trim(),
      material_image_id: String(it.material_image_id || it.materialImageId || "").trim(),
      material_image_name: String(it.material_image_name || it.materialImageName || "").trim(),
      group_id: groupId,
      group_label: groupLabel,
      group_code: groupCode,
      group_order: groupOrder,
      item_order: itemOrder,
      name: rawName || material,
      location: rawLocation,
      detail: rawDetail || spec,
      price: price,
      price_type: priceType
    };
    out.push(normalized);
  }

  out.sort(function(a, b) {
    var ga = Number(a.group_order || 0);
    var gb = Number(b.group_order || 0);
    if (ga !== gb) return ga - gb;
    var ia = Number(a.item_order || 0);
    var ib = Number(b.item_order || 0);
    if (ia !== ib) return ia - ib;
    return Number(a.seq || 0) - Number(b.seq || 0);
  });
  for (var z = 0; z < out.length; z++) {
    out[z].seq = z + 1;
  }
  return out;
}

function buildItemSheetRow_(headers, cols, quoteId, item, seq) {
  var it = item || {};
  var row = Array(headers.length).fill("");
  var itemId = String(it.item_id || "").trim() || uuid_();

  row[cols.quote_id] = String(quoteId || it.quote_id || "");
  row[cols.item_id] = itemId;
  row[cols.seq] = Number(seq || it.seq || 0);

  row[cols.process] = String(it.process || "");
  row[cols.material] = String(it.material || "");
  row[cols.spec] = String(it.spec || "");
  row[cols.qty] = Number(it.qty || 0);
  row[cols.unit] = String(it.unit || "");
  row[cols.unit_price] = Number(it.unit_price || 0);
  row[cols.amount] = (it.amount !== undefined && it.amount !== null)
    ? Number(it.amount || 0)
    : (Number(it.qty || 0) * Number(it.unit_price || 0));
  row[cols.note] = String(it.note || "");

  row[cols.material_ref_id] = String(it.material_ref_id || "");
  row[cols.material_image_id] = String(it.material_image_id || "");
  row[cols.material_image_name] = String(it.material_image_name || "");

  setByHeader_(headers, row, "group_id", String(it.group_id || ""));
  setByHeader_(headers, row, "group_label", String(it.group_label || ""));
  setByHeader_(headers, row, "group_code", String(it.group_code || ""));
  setByHeader_(headers, row, "group_order", Number(it.group_order || 0));
  setByHeader_(headers, row, "item_order", Number(it.item_order || 0));
  setByHeader_(headers, row, "name", String(it.name || ""));
  setByHeader_(headers, row, "location", String(it.location || ""));
  setByHeader_(headers, row, "detail", String(it.detail || ""));
  setByHeader_(headers, row, "price", Number(it.price || row[cols.amount] || 0));
  setByHeader_(headers, row, "price_type", String(it.price_type || "NORMAL"));
  return row;
}

function quoteItemsScriptCacheKey_(quoteId) {
  return "QITEM_" + String(quoteId || "").trim();
}

function quoteItemsSheetCacheKey_(quoteId) {
  return "QITEM_SHEET_" + String(quoteId || "").trim();
}

function quoteItemsIndexCacheKey_(quoteId) {
  return "QITEM_IDX_" + String(quoteId || "").trim();
}

function parseCsvRowNos_(raw) {
  var parts = String(raw || "").split(",");
  var out = [];
  var seen = Object.create(null);
  for (var i = 0; i < parts.length; i++) {
    var rn = Number(String(parts[i] || "").trim());
    if (rn < 2 || seen[rn]) continue;
    seen[rn] = 1;
    out.push(rn);
  }
  out.sort(function(a, b) { return a - b; });
  return out;
}

function parseCsvItemIds_(raw) {
  var parts = String(raw || "").split(",");
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var s = String(parts[i] || "").trim();
    if (!s) continue;
    out.push(s);
  }
  return out;
}

function toCsvRowNos_(rowNos) {
  var list = uniqueSortedRowNos_(rowNos);
  return list.join(",");
}

function toCsvItemIds_(itemIds) {
  var list = Array.isArray(itemIds) ? itemIds : [];
  var out = [];
  for (var i = 0; i < list.length; i++) {
    var id = String(list[i] || "").trim();
    if (!id) continue;
    out.push(id);
  }
  return out.join(",");
}

function invalidateQuoteItemsIndexCache_(quoteId) {
  var target = String(quoteId || "").trim();
  if (!target) return;
  removeCacheKey_(quoteItemsIndexCacheKey_(target));
}

function getQuoteItemsIndexRowMap_() {
  if (__QUOTE_ITEMS_INDEX_ROWNO_CACHE_) return __QUOTE_ITEMS_INDEX_ROWNO_CACHE_;
  var sh = getSheet_("QuoteItemsIndex");
  var cols = getColMap_("QuoteItemsIndex");
  var lastRow = sh.getLastRow();
  var map = Object.create(null);
  if (lastRow >= 2) {
    var ids = sh.getRange(2, cols.quote_id + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      var qid = String(ids[i][0] || "").trim();
      if (!qid) continue;
      map[qid] = i + 2;
    }
  }
  __QUOTE_ITEMS_INDEX_ROWNO_CACHE_ = map;
  return map;
}

function findQuoteItemsIndexRowNo_(quoteId) {
  var target = String(quoteId || "").trim();
  if (!target) return 0;
  return Number(getQuoteItemsIndexRowMap_()[target] || 0);
}

function invalidateQuoteItemsIndexRowMap_() {
  __QUOTE_ITEMS_INDEX_ROWNO_CACHE_ = null;
}

function getQuoteItemsIndex_(quoteId) {
  var target = String(quoteId || "").trim();
  if (!target) return null;

  var cached = getCachedJson_(quoteItemsIndexCacheKey_(target));
  if (cached && typeof cached === "object" && Array.isArray(cached.rowNos)) {
    return {
      rowNos: uniqueSortedRowNos_(cached.rowNos),
      itemIds: Array.isArray(cached.itemIds) ? cached.itemIds.slice() : []
    };
  }

  var sh = getSheet_("QuoteItemsIndex");
  var headers = getHeaders_("QuoteItemsIndex");
  var cols = getColMap_("QuoteItemsIndex");
  var rowNo = findQuoteItemsIndexRowNo_(target);
  if (rowNo < 2) return null;

  var row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
  if (String(row[cols.quote_id] || "").trim() !== target) {
    invalidateQuoteItemsIndexRowMap_();
    rowNo = findQuoteItemsIndexRowNo_(target);
    if (rowNo < 2) return null;
    row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
  }

  var out = {
    rowNos: parseCsvRowNos_(row[cols.row_ids]),
    itemIds: parseCsvItemIds_(row[cols.item_ids])
  };
  putCachedJson_(quoteItemsIndexCacheKey_(target), out, 1800);
  return out;
}

function upsertQuoteItemsIndex_(quoteId, rowNos, itemIds, updatedAtIso) {
  var target = String(quoteId || "").trim();
  if (!target) return;

  var sh = getSheet_("QuoteItemsIndex");
  var headers = getHeaders_("QuoteItemsIndex");
  var cols = getColMap_("QuoteItemsIndex");
  var rowNo = findQuoteItemsIndexRowNo_(target);
  var row = Array(headers.length).fill("");
  if (rowNo >= 2) {
    row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
    if (String(row[cols.quote_id] || "").trim() !== target) {
      invalidateQuoteItemsIndexRowMap_();
      rowNo = findRowNoByColumnValue_(sh, cols.quote_id + 1, target);
      if (rowNo >= 2) {
        row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
      } else {
        row = Array(headers.length).fill("");
      }
    }
  } else {
    rowNo = sh.getLastRow() + 1;
    ensureRowCapacity_(sh, rowNo);
  }

  row[cols.quote_id] = target;
  row[cols.row_ids] = toCsvRowNos_(rowNos);
  row[cols.item_ids] = toCsvItemIds_(itemIds);
  row[cols.updated_at] = String(updatedAtIso || nowIso_());

  sh.getRange(rowNo, 1, 1, headers.length).setValues([row]);
  invalidateQuoteItemsIndexRowMap_();
  putCachedJson_(quoteItemsIndexCacheKey_(target), {
    rowNos: parseCsvRowNos_(row[cols.row_ids]),
    itemIds: parseCsvItemIds_(row[cols.item_ids])
  }, 1800);
}

function upsertQuoteItemsCache_(quoteId, items) {
  var target = String(quoteId || "").trim();
  if (!target) return;
  var list = normalizeItems_(Array.isArray(items) ? items : []);
  var nowIso = nowIso_();

  try {
    CacheService.getScriptCache().put(
      quoteItemsScriptCacheKey_(target),
      JSON.stringify(list),
      21600
    );
  } catch (e) {}

  try {
    var sh = getSheet_("QuoteItemsCache");
    var headers = getHeaders_("QuoteItemsCache");
    var cols = getColMap_("QuoteItemsCache");
    var rowNo = findRowNoByColumnValue_(sh, cols.quote_id + 1, target);
    var row = Array(headers.length).fill("");

    if (rowNo >= 2) {
      row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
    } else {
      rowNo = sh.getLastRow() + 1;
      ensureRowCapacity_(sh, rowNo);
    }

    row[cols.quote_id] = target;
    row[cols.items_json] = JSON.stringify(list);
    row[cols.updated_at] = nowIso;
    sh.getRange(rowNo, 1, 1, headers.length).setValues([row]);
  } catch (sheetErr) {}

  putCachedJson_(quoteItemsSheetCacheKey_(target), list, 600);
}

function getItemsFromCache_(quoteId) {
  var target = String(quoteId || "").trim();
  if (!target) return null;

  try {
    var cached = CacheService.getScriptCache().get(quoteItemsScriptCacheKey_(target));
    if (!cached) return null;
    return JSON.parse(cached);
  } catch (e) {
    // continue fallback to sheet cache
  }

  var sheetCached = getCachedJson_(quoteItemsSheetCacheKey_(target));
  if (sheetCached && Array.isArray(sheetCached)) {
    try {
      CacheService.getScriptCache().put(
        quoteItemsScriptCacheKey_(target),
        JSON.stringify(sheetCached),
        21600
      );
    } catch (e2) {}
    return sheetCached.slice();
  }

  try {
    var sh = getSheet_("QuoteItemsCache");
    var headers = getHeaders_("QuoteItemsCache");
    var cols = getColMap_("QuoteItemsCache");
    var rowNo = findRowNoByColumnValue_(sh, cols.quote_id + 1, target);
    if (rowNo < 2) return null;

    var row = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
    var parsed = parseJsonSafe_(row[cols.items_json], null);
    if (!Array.isArray(parsed)) return null;

    putCachedJson_(quoteItemsSheetCacheKey_(target), parsed, 600);
    try {
      CacheService.getScriptCache().put(
        quoteItemsScriptCacheKey_(target),
        JSON.stringify(parsed),
        21600
      );
    } catch (e3) {}
    return parsed;
  } catch (err) {
    return null;
  }
}

function recalcTotalsFromCache_(quoteId) {
  var target = String(quoteId || "").trim();
  if (!target) throw new Error("quoteId required");

  var items = getItemsFromCache_(target) || [];
  var q = findQuoteRow_(target);
  if (!q) throw new Error("Quote not found");

  var vatRate = Number(getSettings_().vat_rate || 0.1);
  var totals = calculateTotalsFromItems_(
    items,
    q.vat_included,
    vatRate,
    getQuoteTotalOptions_(q)
  );
  updateQuote_(target, totals);
}

function recalcTotals_(quoteId) {
  var target = String(quoteId || "").trim();
  if (!target) throw new Error("quoteId required");

  var items = getItemsFromCache_(target) || findItems_(target);
  var q = findQuoteRow_(target);
  if (!q) throw new Error("Quote not found");

  var vatRate = Number(getSettings_().vat_rate || 0.1);
  var totals = calculateTotalsFromItems_(
    items,
    q.vat_included,
    vatRate,
    getQuoteTotalOptions_(q)
  );
  updateQuote_(target, totals);
}

function getQuoteTotalOptions_(quoteObjOrOptions) {
  var q = quoteObjOrOptions || {};
  return {
    include_service_in_total: normalizeTotalFlag_(q.include_service_in_total, "N"),
    include_option_in_total: normalizeTotalFlag_(q.include_option_in_total, "N"),
    include_included_in_total: normalizeTotalFlag_(q.include_included_in_total, "N")
  };
}

function shouldIncludeItemInTotal_(priceType, options) {
  var type = normalizePriceType_(priceType);
  var opts = getQuoteTotalOptions_(options || {});
  if (type === "NORMAL") return true;
  if (type === "SERVICE") return opts.include_service_in_total === "Y";
  if (type === "OPTION") return opts.include_option_in_total === "Y";
  if (type === "INCLUDED") return opts.include_included_in_total === "Y";
  return false;
}

function resolveItemPrice_(item) {
  var it = item || {};
  if (it.price !== undefined && it.price !== null && String(it.price) !== "") {
    var p = Number(it.price);
    return isNaN(p) ? 0 : p;
  }
  if (it.amount !== undefined && it.amount !== null && String(it.amount) !== "") {
    var a = Number(it.amount);
    return isNaN(a) ? 0 : a;
  }
  var qty = Number(it.qty || 0);
  var unitPrice = Number(it.unit_price || it.unitPrice || 0);
  if (isNaN(qty)) qty = 0;
  if (isNaN(unitPrice)) unitPrice = 0;
  return qty * unitPrice;
}

function calculateTotalsFromItems_(items, vatIncluded, vatRate, options) {
  var subtotal = 0;
  var list = normalizeItems_(Array.isArray(items) ? items : []);

  for (var i = 0; i < list.length; i++) {
    var it = list[i] || {};
    var price = resolveItemPrice_(it);
    if (!shouldIncludeItemInTotal_(it.price_type || it.priceType, options || {})) continue;
    subtotal += Number(price || 0);
  }

  var includeVat = normUpperYN_(vatIncluded) === "Y";
  var rate = Number(vatRate || 0.1);
  if (isNaN(rate) || rate < 0) rate = 0.1;
  var vat = includeVat ? Math.round(subtotal * rate) : 0;
  var total = includeVat ? (subtotal + vat) : subtotal;
  return { subtotal: subtotal, vat: vat, total: total };
}

function findItems_(quoteId) {
  var target = String(quoteId || "").trim();
  if (!target) return [];

  var sh = getSheet_("Items");
  var headers = getHeaders_("Items");
  var cols = getColMap_("Items");
  var out = [];
  var idx = getQuoteItemsIndex_(target);

  if (idx && idx.rowNos && idx.rowNos.length) {
    var rowsMapByIdx = readRowsMapByRowNos_(sh, idx.rowNos, 1, headers.length);
    var posByItemId = Object.create(null);
    for (var p = 0; p < idx.itemIds.length; p++) {
      posByItemId[String(idx.itemIds[p] || "")] = p;
    }

    var valid = true;
    for (var i = 0; i < idx.rowNos.length; i++) {
      var rowNoByIdx = idx.rowNos[i];
      var rowByIdx = rowsMapByIdx[rowNoByIdx];
      if (!rowByIdx) {
        valid = false;
        continue;
      }
      if (String(rowByIdx[cols.quote_id] || "").trim() !== target) {
        valid = false;
        continue;
      }
      out.push(rowToObj_(headers, rowByIdx, rowNoByIdx));
    }

    if (valid && out.length) {
      out.sort(function(a, b) {
        var pa = Number(posByItemId[String(a.item_id || "")]);
        var pb = Number(posByItemId[String(b.item_id || "")]);
        if (!isNaN(pa) && !isNaN(pb) && pa !== pb) return pa - pb;
        return Number(a.seq || a._rowNo) - Number(b.seq || b._rowNo);
      });
      return normalizeItems_(out);
    }
  }

  var rowNos = findRowNosByColumnValue_(sh, cols.quote_id + 1, target);
  if (!rowNos.length) {
    upsertQuoteItemsIndex_(target, [], [], nowIso_());
    return [];
  }

  var rowsMap = readRowsMapByRowNos_(sh, rowNos, 1, headers.length);
  out = [];

  for (var r = 0; r < rowNos.length; r++) {
    var rowNo = rowNos[r];
    var row = rowsMap[rowNo];
    if (!row) continue;
    if (String(row[cols.quote_id] || "").trim() !== target) continue;
    out.push(rowToObj_(headers, row, rowNo));
  }

  out.sort(function(a, b) {
    return Number(a.seq || a._rowNo) - Number(b.seq || b._rowNo);
  });

  var idxRows = [];
  var idxItems = [];
  for (var z = 0; z < out.length; z++) {
    idxRows.push(Number(out[z]._rowNo || 0));
    idxItems.push(String(out[z].item_id || ""));
  }
  upsertQuoteItemsIndex_(target, idxRows, idxItems, nowIso_());
  return normalizeItems_(out);
}

function findPhotos_(quoteId) {
  var target = String(quoteId || "").trim();
  if (!target) return [];

  var sh = getSheet_("Photos");
  var headers = getHeaders_("Photos");
  var cols = getColMap_("Photos");
  var rowNos = findRowNosByColumnValue_(sh, cols.quote_id + 1, target);
  if (!rowNos.length) return [];

  var rowsMap = readRowsMapByRowNos_(sh, rowNos, 1, headers.length);
  var out = [];

  for (var i = 0; i < rowNos.length; i++) {
    var rowNo = rowNos[i];
    var row = rowsMap[rowNo];
    if (!row) continue;
    if (String(row[cols.quote_id] || "").trim() !== target) continue;
    out.push(rowToObj_(headers, row, rowNo));
  }
  return out;
}

function getOrCreateUploadRootFolder_() {
  if (__UPLOAD_ROOT_FOLDER_CACHE_) return __UPLOAD_ROOT_FOLDER_CACHE_;

  var props = PropertiesService.getScriptProperties();
  var key = "UPLOAD_ROOT_FOLDER_ID";
  var s = getSettings_();
  var configured = String(s.upload_folder_id || "").trim();

  if (configured) {
    try {
      __UPLOAD_ROOT_FOLDER_CACHE_ = DriveApp.getFolderById(configured);
      props.setProperty(key, __UPLOAD_ROOT_FOLDER_CACHE_.getId());
      return __UPLOAD_ROOT_FOLDER_CACHE_;
    } catch (e) {}
  }

  var saved = props.getProperty(key);
  if (saved) {
    try {
      __UPLOAD_ROOT_FOLDER_CACHE_ = DriveApp.getFolderById(saved);
      return __UPLOAD_ROOT_FOLDER_CACHE_;
    } catch (e2) {}
  }

  __UPLOAD_ROOT_FOLDER_CACHE_ = DriveApp.createFolder("quoteapp_uploads");
  props.setProperty(key, __UPLOAD_ROOT_FOLDER_CACHE_.getId());
  return __UPLOAD_ROOT_FOLDER_CACHE_;
}

function getOrCreateQuoteUploadFolder_(quoteId) {
  var qid = String(quoteId || "").trim();
  if (!qid) throw new Error("quoteId required");
  if (__QUOTE_FOLDER_CACHE_[qid]) return __QUOTE_FOLDER_CACHE_[qid];

  var root = getOrCreateUploadRootFolder_();
  var name = "quote_" + qid;
  var it = root.getFoldersByName(name);
  __QUOTE_FOLDER_CACHE_[qid] = it.hasNext() ? it.next() : root.createFolder(name);
  return __QUOTE_FOLDER_CACHE_[qid];
}

function getOrCreateMaterialsFolder_() {
  if (__MATERIALS_FOLDER_CACHE_) return __MATERIALS_FOLDER_CACHE_;
  var root = getOrCreateUploadRootFolder_();
  var name = "materials";
  var it = root.getFoldersByName(name);
  __MATERIALS_FOLDER_CACHE_ = it.hasNext() ? it.next() : root.createFolder(name);
  return __MATERIALS_FOLDER_CACHE_;
}

function parseDataUrl_(dataUrl) {
  var m = String(dataUrl || "").match(/^data:(.+?);base64,(.+)$/);
  if (!m) throw new Error("Invalid dataUrl");
  return { mimeType: m[1], base64: m[2] };
}

function sanitizeFilename_(name) {
  return String(name || "file")
    .replace(/[\\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80);
}

function tryResizeImageBytes_(bytes, mimeType, maxWidth) {
  try {
    var img = ImagesService.openImage(bytes);
    if (img.getWidth() > maxWidth) return img.resize(maxWidth).getBytes();
  } catch (e) {}
  return bytes;
}

function ensureCoreSchemaReady_() {
  if (__SCHEMA_READY_THIS_EXEC_) return;

  var props = PropertiesService.getScriptProperties();
  var ver = props.getProperty(CORE_SCHEMA_FLAG_KEY_);
  if (ver === CORE_SCHEMA_VERSION_) {
    __SCHEMA_READY_THIS_EXEC_ = true;
    return;
  }

  forceEnsureCoreSchema_();
}

function forceEnsureCoreSchema_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var current = props.getProperty(CORE_SCHEMA_FLAG_KEY_);
    if (current !== CORE_SCHEMA_VERSION_) {
      ensureCoreSchema_();
      props.setProperty(CORE_SCHEMA_FLAG_KEY_, CORE_SCHEMA_VERSION_);
    }
    __SCHEMA_READY_THIS_EXEC_ = true;
  } finally {
    lock.releaseLock();
  }
}

function ensureCoreSchema_() {
  ensureSheetColumns_("Settings", ["key", "value"]);
  ensureSheetColumns_("Quotes", [
    "quote_id", "created_at", "updated_at",
    "customer_name", "site_name", "contact_name", "contact_phone", "memo",
    "vat_included",
    "include_service_in_total", "include_option_in_total", "include_included_in_total",
    "subtotal", "vat", "total",
    "share_token", "status",
    "shared_at", "expire_at", "view_count", "last_viewed_at", "next_action_at", "last_followup_sent_at",
    "approved_at", "deposit_amount", "deposit_due_at",
    "last_customer_note_at", "customer_note_latest", "owner_last_seen_note_at"
  ]);
  ensureSheetColumns_("Items", ITEMS_CANONICAL_HEADERS_);
  ensureSheetColumns_("Photos", [
    "quote_id", "photo_id", "file_id", "file_name", "note", "created_at"
  ]);
  ensureSheetColumns_("ViewLog", [
    "log_id", "quote_id", "page", "viewed_at"
  ]);
  ensureSheetColumns_("CustomerNotes", [
    "note_id", "quote_id", "note", "created_at", "source"
  ]);
  ensureMaterialsSheet_();
  ensureMaterialGroupsSheet_();
  ensureMaterialTagsSheet_();
  ensureSheetColumns_("MaterialsIndex", ["token", "refs", "updated_at"]);
  ensureSheetColumns_("SnapshotState", ["snapshot_key", "shared_master_version", "built_at", "row_count", "duration_ms", "status", "note"]);
  ensureSheetColumns_("PerfMetrics", ["metric_id", "measured_at", "metric_name", "duration_ms", "entity_id", "actor_type", "actor_id", "context_json", "success_yn", "note"]);
  ensureSheetColumns_("QuoteItemsCache", ["quote_id", "items_json", "updated_at"]);
  ensureSheetColumns_("QuoteItemsIndex", ["quote_id", "row_ids", "item_ids", "updated_at"]);
  ensureSheetColumns_("SlackQueue", [
    "queue_id",
    "created_at",
    "next_attempt_at",
    "status",
    "event_type",
    "quote_id",
    "dedup_key",
    "payload_json",
    "attempt_count",
    "last_attempt_at",
    "last_http_code",
    "last_error",
    "sent_at"
  ]);

  ensureDefaultSettingsRows_();
  invalidateSettingsCache_();
}

function ensureDefaultSettingsRows_() {
  var sh = getSheet_("Settings");
  var lastRow = sh.getLastRow();

  var defaults = [
    ["admin_password", ""],
    ["vat_rate", "0.1"],
    ["default_vat_included", "N"],
    ["default_include_service_in_total", "N"],
    ["default_include_option_in_total", "N"],
    ["default_include_included_in_total", "N"],
    ["share_expire_days", "14"],
    ["followup_first_hours", "24"],
    ["followup_cooldown_minutes", "360"],
    ["followup_repeat_hours", "24"],
    ["followup_after_view_minutes", "120"],
    ["max_followups_per_run", "10"],
    ["deposit_rate", "0.1"],
    ["deposit_due_days", "3"],
    ["base_url", ""],
    ["upload_folder_id", ""],
    ["owner_email", ""],
    ["account_info", ""],
    ["slack_queue_batch_size", "30"],
    ["slack_retry_max_attempts", "6"],
    ["slack_retry_base_seconds", "60"],
    ["slack_retry_max_seconds", "1800"],
    ["expose_materials_to_prequote", "Y"],
    ["expose_templates_to_prequote", "Y"],
    ["prequote_default_material_limit", "12"],
    ["prequote_default_template_limit", "12"],
    ["prequote_sync_enabled", "N"],
    ["prequote_sync_note", ""],
    ["prequote_allow_snapshot_api", "Y"],
    ["prequote_snapshot_cache_ttl_sec", "600"],
    ["prequote_material_search_mode", "INDEX"],
    ["perf_metrics_enabled", "Y"],
    ["perf_metrics_retention_days", "30"]
  ];

  if (lastRow < 2) {
    ensureRowCapacity_(sh, 1 + defaults.length);
    sh.getRange(2, 1, defaults.length, 2).setValues(defaults);
    return;
  }

  var existing = Object.create(null);
  var keys = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < keys.length; i++) {
    var k = String(keys[i][0] || "").trim();
    if (k) existing[k] = 1;
  }

  var append = [];
  for (var d = 0; d < defaults.length; d++) {
    if (!existing[defaults[d][0]]) append.push(defaults[d]);
  }
  if (!append.length) return;

  var start = sh.getLastRow() + 1;
  ensureRowCapacity_(sh, start + append.length - 1);
  sh.getRange(start, 1, append.length, 2).setValues(append);
}

function ensureSheetColumns_(sheetName, wantedHeaders, overrideId) {
  var ss = getSpreadsheet_(overrideId);
  return ensureSheetColumnsInSs_(ss, sheetName, wantedHeaders);
}

function ensureSheetColumnsInSs_(ss, sheetName, wantedHeaders) {
  var targetSs = ss || getSpreadsheet_();
  var sh = targetSs.getSheetByName(sheetName);
  if (!sh) sh = targetSs.insertSheet(sheetName);

  var headersWanted = Array.isArray(wantedHeaders) ? wantedHeaders.slice() : [];
  if (!headersWanted.length) return sh;

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headersWanted.length).setValues([headersWanted]);
    invalidateSheetCache_(sheetName, targetSs.getId());
    invalidateSheetMetaCache_(sheetName);
    return sh;
  }

  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var missing = [];
  for (var i = 0; i < headersWanted.length; i++) {
    if (headers.indexOf(headersWanted[i]) < 0) missing.push(headersWanted[i]);
  }

  if (missing.length) {
    sh.insertColumnsAfter(lastCol, missing.length);
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  }

  invalidateSheetCache_(sheetName, targetSs.getId());
  invalidateSheetMetaCache_(sheetName);
  return sh;
}

function ensureRowCapacity_(sh, neededLastRow) {
  var maxRows = sh.getMaxRows();
  if (neededLastRow > maxRows) sh.insertRowsAfter(maxRows, neededLastRow - maxRows);
}

function parseIsoMs_(v) {
  var s = String(v || "").trim();
  if (!s) return NaN;
  var ms = new Date(s).getTime();
  return ms;
}

function normUpperYN_(v) {
  var s = String(v || "N").trim().toUpperCase();
  return s === "Y" ? "Y" : "N";
}

function boolToYn_(flag) {
  return flag ? "Y" : "N";
}

function ynToBool_(value, defaultValue) {
  if (value === undefined || value === null || String(value).trim() === "") {
    return !!defaultValue;
  }
  var s = String(value).trim().toUpperCase();
  if (s === "Y" || s === "TRUE" || s === "1" || s === "YES" || s === "ON") return true;
  if (s === "N" || s === "FALSE" || s === "0" || s === "NO" || s === "OFF") return false;
  return !!defaultValue;
}

function normalizeTotalFlag_(value, fallback) {
  if (value === undefined || value === null || String(value).trim() === "") {
    return normUpperYN_(fallback);
  }
  var s = String(value).trim().toUpperCase();
  if (s === "TRUE" || s === "1" || s === "Y" || s === "YES" || s === "ON") return "Y";
  if (s === "FALSE" || s === "0" || s === "N" || s === "NO" || s === "OFF") return "N";
  return normUpperYN_(fallback);
}

function normalizeTagType_(value) {
  var s = String(value || "").trim().toLowerCase();
  if (!s) return "";
  s = s.replace(/[^a-z0-9_]+/g, "_").replace(/^_+|_+$/g, "");
  if (!s) return "";
  return MATERIAL_TAG_TYPES_.indexOf(s) >= 0 ? s : s;
}

function normalizeTagValue_(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9가-힣_]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function normalizeTradeCode_(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9_]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function normalizeSimpleToken_(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9가-힣_]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function normalizeNumber_(value, fallback) {
  var n = Number(value);
  return isNaN(n) ? Number(fallback || 0) : n;
}

function uniqueStrings_(list) {
  var arr = Array.isArray(list) ? list : [];
  var out = [];
  var seen = Object.create(null);
  for (var i = 0; i < arr.length; i++) {
    var s = String(arr[i] || "").trim();
    if (!s || seen[s]) continue;
    seen[s] = 1;
    out.push(s);
  }
  return out;
}

function normalizeStringList_(value, normalizer) {
  var rawList = [];
  if (Array.isArray(value)) {
    rawList = value.slice();
  } else if (value !== undefined && value !== null && String(value).trim() !== "") {
    rawList = String(value).split(/[,\n;|\/]+/);
  }
  var out = [];
  var seen = Object.create(null);
  for (var i = 0; i < rawList.length; i++) {
    var raw = rawList[i];
    var normalized = normalizer ? normalizer(raw) : String(raw || "").trim();
    if (!normalized || seen[normalized]) continue;
    seen[normalized] = 1;
    out.push(normalized);
  }
  return out;
}

function splitSummaryTags_(value) {
  return normalizeStringList_(value, normalizeTagValue_);
}

function cloneObj_(o) {
  var out = {};
  if (!o) return out;
  for (var k in o) out[k] = o[k];
  return out;
}

function findRowNoByColumnValue_(sh, colNo, target) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  var t = String(target || "").trim();
  if (!t) return 0;

  var values = sh.getRange(2, colNo, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === t) return i + 2;
  }
  return 0;
}

function findRowNosByColumnValue_(sh, colNo, target) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var t = String(target || "").trim();
  if (!t) return [];

  var values = sh.getRange(2, colNo, lastRow - 1, 1).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === t) out.push(i + 2);
  }
  return out;
}

function writeRowsByRowNos_(sh, rowNos, rows, startCol, width) {
  if (!rowNos || !rowNos.length) return;
  var col = Number(startCol || 1);
  var w = Number(width || (rows[0] ? rows[0].length : 0));
  if (!w) return;

  var i = 0;
  while (i < rowNos.length) {
    var startIdx = i;
    var startRow = rowNos[i];
    while (i + 1 < rowNos.length && rowNos[i + 1] === rowNos[i] + 1) i++;
    var endIdx = i;
    var blockRows = endIdx - startIdx + 1;
    var blockVals = rows.slice(startIdx, endIdx + 1);
    sh.getRange(startRow, col, blockRows, w).setValues(blockVals);
    i++;
  }
}

function clearRowsByRowNos_(sh, rowNos, startCol, width) {
  if (!rowNos || !rowNos.length) return;
  var col = Number(startCol || 1);
  var w = Number(width || 1);
  if (!w) return;

  var i = 0;
  while (i < rowNos.length) {
    var startRow = rowNos[i];
    while (i + 1 < rowNos.length && rowNos[i + 1] === rowNos[i] + 1) i++;
    var endRow = rowNos[i];
    sh.getRange(startRow, col, endRow - startRow + 1, w).clearContent();
    i++;
  }
}

function readRowsMapByRowNos_(sh, rowNos, startCol, width) {
  var out = Object.create(null);
  if (!rowNos || !rowNos.length) return out;
  var col = Number(startCol || 1);
  var w = Number(width || 1);

  var sorted = uniqueSortedRowNos_(rowNos);
  var i = 0;
  while (i < sorted.length) {
    var startRow = sorted[i];
    while (i + 1 < sorted.length && sorted[i + 1] === sorted[i] + 1) i++;
    var endRow = sorted[i];
    var count = endRow - startRow + 1;
    var values = sh.getRange(startRow, col, count, w).getValues();
    for (var r = 0; r < count; r++) out[startRow + r] = values[r];
    i++;
  }
  return out;
}

function uniqueSortedRowNos_(rowNos) {
  var seen = Object.create(null);
  var out = [];
  var list = Array.isArray(rowNos) ? rowNos : [];
  for (var i = 0; i < list.length; i++) {
    var rn = Number(list[i] || 0);
    if (rn < 2) continue;
    if (seen[rn]) continue;
    seen[rn] = 1;
    out.push(rn);
  }
  out.sort(function(a, b) { return a - b; });
  return out;
}

function getCachedJson_(key) {
  try {
    var raw = CacheService.getScriptCache().get(String(key || ""));
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function putCachedJson_(key, obj, ttlSec) {
  try {
    CacheService.getScriptCache().put(
      String(key || ""),
      JSON.stringify(obj),
      Number(ttlSec || 120)
    );
  } catch (e) {}
}

function removeCacheKey_(key) {
  try { CacheService.getScriptCache().remove(String(key || "")); } catch (e) {}
}

function cacheChunkMetaKey_(baseKey) {
  return String(baseKey || "") + "__META";
}

function cacheChunkPartKey_(baseKey, idx) {
  return String(baseKey || "") + "__P" + String(Number(idx || 0));
}

function getCachedJsonChunked_(baseKey) {
  var key = String(baseKey || "").trim();
  if (!key) return null;
  try {
    var cache = CacheService.getScriptCache();
    var metaRaw = cache.get(cacheChunkMetaKey_(key));
    if (!metaRaw) return null;
    var meta = parseJsonSafe_(metaRaw, null);
    var count = Number(meta && meta.count || 0);
    if (count < 1) return null;
    var chunks = [];
    for (var i = 0; i < count; i++) {
      var part = cache.get(cacheChunkPartKey_(key, i));
      if (part === null || part === undefined) return null;
      chunks.push(String(part));
    }
    return JSON.parse(chunks.join(""));
  } catch (e) {
    return null;
  }
}

function putCachedJsonChunked_(baseKey, obj, ttlSec) {
  var key = String(baseKey || "").trim();
  if (!key) return;
  var ttl = Math.max(Number(ttlSec || 120), 60);
  try {
    var raw = JSON.stringify(obj);
    var cache = CacheService.getScriptCache();
    var chunkSize = 85000;
    var count = Math.max(Math.ceil(raw.length / chunkSize), 1);
    for (var i = 0; i < count; i++) {
      var start = i * chunkSize;
      var end = Math.min(start + chunkSize, raw.length);
      cache.put(cacheChunkPartKey_(key, i), raw.slice(start, end), ttl);
    }
    cache.put(cacheChunkMetaKey_(key), JSON.stringify({ count: count }), ttl);
  } catch (e) {}
}

function removeCachedJsonChunked_(baseKey) {
  var key = String(baseKey || "").trim();
  if (!key) return;
  try {
    var cache = CacheService.getScriptCache();
    var metaRaw = cache.get(cacheChunkMetaKey_(key));
    var meta = parseJsonSafe_(metaRaw, null);
    var count = Number(meta && meta.count || 0);
    if (count > 0) {
      var keys = [];
      for (var i = 0; i < count; i++) keys.push(cacheChunkPartKey_(key, i));
      try { cache.removeAll(keys); } catch (e1) {
        for (var k = 0; k < keys.length; k++) {
          try { cache.remove(keys[k]); } catch (e2) {}
        }
      }
    }
    try { cache.remove(cacheChunkMetaKey_(key)); } catch (e3) {}
  } catch (e4) {}
}

function getOrBuildCachedJsonWithLock_(cacheKey, ttlSec, builderFn, options) {
  var key = String(cacheKey || "").trim();
  if (!key) return (typeof builderFn === "function") ? builderFn() : null;
  var opts = options || {};

  var cached = getCachedJson_(key);
  if (cached !== null) return cached;

  var lock = LockService.getScriptLock();
  var waitMs = Math.max(Number(opts.wait_ms || 1800), 0);
  var gotLock = false;
  try { gotLock = lock.tryLock(waitMs); } catch (e0) { gotLock = false; }
  if (gotLock) {
    try {
      var cached2 = getCachedJson_(key);
      if (cached2 !== null) return cached2;
      var built = (typeof builderFn === "function") ? builderFn() : null;
      if (built !== undefined) putCachedJson_(key, built, ttlSec);
      return built;
    } finally {
      lock.releaseLock();
    }
  }

  var retryMs = Math.max(Number(opts.retry_ms || 80), 0);
  if (retryMs > 0) Utilities.sleep(retryMs);
  var cached3 = getCachedJson_(key);
  if (cached3 !== null) return cached3;
  if (opts && opts.hasOwnProperty("fallback_value")) return opts.fallback_value;
  return (typeof builderFn === "function") ? builderFn() : null;
}

function getOrBuildCachedJsonChunkedWithLock_(cacheKey, ttlSec, builderFn, options) {
  var key = String(cacheKey || "").trim();
  if (!key) return (typeof builderFn === "function") ? builderFn() : null;
  var opts = options || {};

  var cached = getCachedJsonChunked_(key);
  if (cached !== null) return cached;

  var lock = LockService.getScriptLock();
  var waitMs = Math.max(Number(opts.wait_ms || 2200), 0);
  var gotLock = false;
  try { gotLock = lock.tryLock(waitMs); } catch (e0) { gotLock = false; }
  if (gotLock) {
    try {
      var cached2 = getCachedJsonChunked_(key);
      if (cached2 !== null) return cached2;
      var built = (typeof builderFn === "function") ? builderFn() : null;
      if (built !== undefined) putCachedJsonChunked_(key, built, ttlSec);
      return built;
    } finally {
      lock.releaseLock();
    }
  }

  var retryMs = Math.max(Number(opts.retry_ms || 120), 0);
  if (retryMs > 0) Utilities.sleep(retryMs);
  var cached3 = getCachedJsonChunked_(key);
  if (cached3 !== null) return cached3;
  if (opts && opts.hasOwnProperty("fallback_value")) return opts.fallback_value;
  return (typeof builderFn === "function") ? builderFn() : null;
}

function getMaterialSearchVersion_() {
  // Cache key contract:
  // - MATERIAL_SEARCH_VER: invalidates material search query cache + image allowlist column-set cache.
  var props = PropertiesService.getScriptProperties();
  var v = String(props.getProperty("MATERIAL_SEARCH_VER") || "").trim();
  if (!v) {
    v = "1";
    props.setProperty("MATERIAL_SEARCH_VER", v);
  }
  return v;
}

function bumpMaterialSearchVersion_() {
  PropertiesService.getScriptProperties().setProperty("MATERIAL_SEARCH_VER", String(Date.now()));
  __MATERIAL_INDEX_MEM_VER_ = "";
  __MATERIAL_INDEX_MEM_MAP_ = null;
  __SHEET_COLUMN_SET_MEM_ = Object.create(null);
}

function getMaterialGroupListVersion_() {
  // Cache key contract:
  // - MATERIAL_GROUP_LIST_VER: invalidates material-group filter/list caches.
  var props = PropertiesService.getScriptProperties();
  var v = String(props.getProperty(MATERIAL_GROUP_LIST_VERSION_KEY_) || "").trim();
  if (!v) {
    v = "1";
    props.setProperty(MATERIAL_GROUP_LIST_VERSION_KEY_, v);
  }
  return v;
}

function bumpMaterialGroupListVersion_() {
  PropertiesService.getScriptProperties().setProperty(MATERIAL_GROUP_LIST_VERSION_KEY_, String(Date.now()));
}

function getQuoteListVersion_() {
  // Cache key contract:
  // - QUOTE_LIST_VER: invalidates dashboard/list core caches + item/photo reference column-set cache.
  var props = PropertiesService.getScriptProperties();
  var v = String(props.getProperty("QUOTE_LIST_VER") || "").trim();
  if (!v) {
    v = "1";
    props.setProperty("QUOTE_LIST_VER", v);
  }
  return v;
}

function bumpQuoteListVersion_() {
  PropertiesService.getScriptProperties().setProperty("QUOTE_LIST_VER", String(Date.now()));
  __QUOTE_LIST_CORE_MEM_VER_ = "";
  __QUOTE_LIST_CORE_MEM_ = null;
  __SHEET_COLUMN_SET_MEM_ = Object.create(null);
}

function pickSearchTokens_(tokens, maxTokens) {
  var list = Array.isArray(tokens) ? tokens.slice() : [];
  var seen = Object.create(null);
  var uniq = [];
  for (var i = 0; i < list.length; i++) {
    var t = String(list[i] || "").trim();
    if (!t || seen[t]) continue;
    seen[t] = 1;
    uniq.push(t);
  }
  uniq.sort(function(a, b) {
    var dl = b.length - a.length;
    if (dl !== 0) return dl;
    return a.localeCompare(b);
  });
  return uniq.slice(0, Math.max(Number(maxTokens || 8), 1));
}

function quoteNotesCacheKey_(quoteId) {
  return "QNOTES_" + String(quoteId || "").trim();
}

function getCachedQuoteNotes_(quoteId) {
  return getCachedJson_(quoteNotesCacheKey_(quoteId));
}

function setCachedQuoteNotes_(quoteId, list, ttlSec) {
  putCachedJson_(quoteNotesCacheKey_(quoteId), Array.isArray(list) ? list : [], Number(ttlSec || 600));
}

function invalidateCachedQuoteNotes_(quoteId) {
  removeCacheKey_(quoteNotesCacheKey_(quoteId));
}

function materialSearchQueryCacheKey_(queryNorm, limit, offset) {
  var ver = [getMaterialSearchVersion_(), getMaterialTagListVersion_()].join("_");
  var q = normText_(String(queryNorm || ""));
  var max = Math.min(Math.max(Number(limit || 50), 1), 200);
  var off = Math.max(Math.floor(Number(offset || 0)), 0);
  return "MATSEARCH_" + ver + "_" + q + "_" + max + "_" + off;
}

function clearMaterialSearchQueryCachesForPerf_(queryNorm, limit, offset) {
  var q = normText_(String(queryNorm || ""));
  var max = Math.min(Math.max(Number(limit || 50), 1), 200);
  var off = Math.max(Math.floor(Number(offset || 0)), 0);
  removeCacheKey_(materialSearchQueryCacheKey_(q, max, off));

  var tokens = pickSearchTokens_(tokenize_(q), 8);
  var ver = [getMaterialSearchVersion_(), getMaterialTagListVersion_()].join("_");
  if (!tokens.length) return;
  var keys = [];
  for (var i = 0; i < tokens.length; i++) keys.push("MIDXTOK_" + ver + "_" + tokens[i]);
  try { CacheService.getScriptCache().removeAll(keys); } catch (e) {
    for (var k = 0; k < keys.length; k++) removeCacheKey_(keys[k]);
  }
}

function clearQuoteDashboardCachesForPerf_(status, query, limit, offset) {
  var statusFilter = String(status || "").trim().toLowerCase();
  var queryNorm = normText_(query || "");
  var max = Math.min(Math.max(Number(limit || 50), 1), 500);
  var off = Math.max(Math.floor(Number(offset || 0)), 0);
  removeCacheKey_(quoteDashboardCacheKey_(statusFilter, queryNorm, max, off));
  removeCacheKey_(quoteCoreListCacheKey_());
  __QUOTE_LIST_CORE_MEM_VER_ = "";
  __QUOTE_LIST_CORE_MEM_ = null;
}

function clearTemplateListCachesForPerf_(category, query, includeInactive) {
  removeCacheKey_(templateCatalogListCacheKey_(category, query, includeInactive));
  removeCacheKey_(templateCategoryListCacheKey_(includeInactive));
  removeCachedJsonChunked_(templateCatalogSnapshotCacheKey_());
  removeCachedJsonChunked_(templateCatalogCoreSnapshotCacheKey_());
  __TEMPLATE_CATALOG_SNAPSHOT_MEM_VER_ = "";
  __TEMPLATE_CATALOG_SNAPSHOT_MEM_ = null;
  __TEMPLATE_CATALOG_CORE_MEM_VER_ = "";
  __TEMPLATE_CATALOG_CORE_MEM_ = null;
}

function clearMaterialGroupListCachesForPerf_(filters) {
  removeCacheKey_(materialGroupsListCacheKey_(filters || {}));
}

function benchmarkColdWarmMs_(rounds, coldPrepareFn, callFn) {
  var runs = Math.max(Number(rounds || 1), 1);
  var coldRuns = Math.min(runs, 3);

  var cold = benchmarkFnMs_(coldRuns, function(runIdx) {
    if (typeof coldPrepareFn === "function") coldPrepareFn(runIdx);
    callFn(runIdx);
  });

  try { callFn(0); } catch (e1) {}
  var warm = benchmarkFnMs_(runs, function(runIdx) {
    callFn(runIdx);
  });

  return {
    cold_miss: cold,
    warm_hit: warm
  };
}

function runPerfValidation(adminPassword, options) {
  assertAdminCredential_(adminPassword);

  var opts = options || {};
  var rounds = Math.min(Math.max(Number(opts.rounds || 3), 1), 10);
  var materialQuery = normText_(String(opts.material_query || "a"));
  var materialLimit = Math.min(Math.max(Number(opts.material_limit || 50), 1), 200);
  var templateCategory = String(opts.template_category || "").trim();
  var templateQuery = String(opts.template_query || "").trim();
  var templateIncludeInactive = !!opts.template_include_inactive;
  var materialGroupFilters = normalizeMaterialGroupListFilters_({
    query: String(opts.group_query || "").trim(),
    space_type: String(opts.group_space_type || "").trim(),
    trade_code: String(opts.group_trade_code || "").trim(),
    is_active: String(opts.group_is_active || "").trim(),
    include_inactive: (opts.group_include_inactive === undefined) ? true : !!opts.group_include_inactive
  });
  var noteQuoteId = String(opts.quote_id || "").trim();
  var noteToken = String(opts.token || "").trim();

  var report = {
    generated_at: nowIso_(),
    schema_version: CORE_SCHEMA_VERSION_,
    rounds: rounds,
    env: {},
    timings_ms: {
      schema_init: 0,
      server_rpc: {},
      web_get: {}
    },
    integration_scenarios: [],
    notes: []
  };

  var initStart = Date.now();
  forceEnsureCoreSchema_();
  report.timings_ms.schema_init = Date.now() - initStart;

  var s = getSettings_();
  var baseUrl = getConfiguredBaseUrl_(s);
  report.env.base_url = baseUrl || "";
  report.env.spreadsheet_id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID") || "";
  report.env.versions = {
    material_search: getMaterialSearchVersion_(),
    material_group_list: getMaterialGroupListVersion_(),
    template_list: getTemplateListVersion_(),
    quote_list: getQuoteListVersion_()
  };

  report.timings_ms.server_rpc.getPublicAppConfig = benchmarkColdWarmMs_(rounds, function() {
    removeCacheKey_("PUBLIC_APP_CFG_V2");
  }, function() {
    getPublicAppConfig();
  });

  report.timings_ms.server_rpc.listQuotesForDashboard = benchmarkColdWarmMs_(rounds, function() {
    clearQuoteDashboardCachesForPerf_("", "", 50, 0);
  }, function() {
    listQuotesForDashboard("", "", 50, adminPassword);
  });

  report.timings_ms.server_rpc.searchMaterials = benchmarkColdWarmMs_(rounds, function() {
    clearMaterialSearchQueryCachesForPerf_(materialQuery, materialLimit, 0);
  }, function() {
    searchMaterials(materialQuery, adminPassword, { limit: materialLimit, offset: 0 });
  });

  report.timings_ms.server_rpc.listMaterialGroupsAdmin = benchmarkColdWarmMs_(rounds, function() {
    clearMaterialGroupListCachesForPerf_(materialGroupFilters);
  }, function() {
    listMaterialGroupsAdmin(materialGroupFilters, adminPassword);
  });

  report.timings_ms.server_rpc.listTemplateCatalog = benchmarkColdWarmMs_(rounds, function() {
    clearTemplateListCachesForPerf_(templateCategory, templateQuery, templateIncludeInactive);
  }, function() {
    listTemplateCatalog(templateCategory, templateQuery, templateIncludeInactive, adminPassword);
  });

  if (noteQuoteId && noteToken) {
    report.timings_ms.server_rpc.getQuote = benchmarkColdWarmMs_(rounds, function() {
      var q = findQuoteRow_(noteQuoteId);
      if (!q) return;
      var key = buildGetQuoteCacheKey_(noteQuoteId, noteToken, q.updated_at, {
        include_notes: false,
        include_photos: false,
        note_limit: 100
      }, 100);
      removeCacheKey_(key);
      removeCacheKey_(quoteItemsSheetCacheKey_(noteQuoteId));
      try { CacheService.getScriptCache().remove(quoteItemsScriptCacheKey_(noteQuoteId)); } catch (e0) {}
    }, function() {
      getQuote(noteQuoteId, noteToken, { include_notes: false, include_photos: false, note_limit: 100 });
    });

    report.timings_ms.server_rpc.listCustomerNotesByToken = benchmarkColdWarmMs_(rounds, function() {
      removeCacheKey_(quoteNotesCacheKey_(noteQuoteId));
    }, function() {
      listCustomerNotesByToken(noteQuoteId, noteToken, 100);
    });
  } else {
    report.notes.push("quote_id/token not provided. getQuote/listCustomerNotesByToken benchmark skipped.");
  }

  report.integration_scenarios = [
    { name: "quote_create_edit_dashboard", steps: ["createQuote", "saveQuote", "listQuotesForDashboard"] },
    { name: "material_save_search", steps: ["saveMaterial", "searchMaterials", "listMaterialGroupsAdmin"] },
    { name: "template_save_list_detail", steps: ["saveTemplateVersion", "listTemplateCatalog", "getTemplateVersionDetail"] }
  ];

  if (!baseUrl) {
    report.notes.push("Settings.base_url is empty. web_get timings are skipped.");
    return report;
  }

  var pages = ["edit", "catalog", "dashboard"];
  for (var i = 0; i < pages.length; i++) {
    (function(page) {
      report.timings_ms.web_get[page] = benchmarkFnMs_(rounds, function(runIdx) {
        var url = baseUrl + "?page=" + encodeURIComponent(page) + "&_perf_run=" + encodeURIComponent(String(runIdx));
        var res = UrlFetchApp.fetch(url, {
          method: "get",
          followRedirects: true,
          muteHttpExceptions: true
        });
        var code = Number(res.getResponseCode() || 0);
        if (code < 200 || code >= 400) throw new Error("HTTP " + code + " for " + page);
      });
    })(pages[i]);
  }

  report.notes.push("Execution transcript call-count check is manual in Apps Script IDE.");
  return report;
}

function runIntegrationSmokeScenarios(adminPassword, options) {
  assertAdminCredential_(adminPassword);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();

  var opts = options || {};
  var quoteId = String(opts.quote_id || "").trim();
  var token = String(opts.token || "").trim();
  var materialQuery = String(opts.material_query || "LG").trim();
  var templateCategory = String(opts.template_category || "").trim();

  var report = {
    generated_at: nowIso_(),
    ok: true,
    checks: []
  };

  function pushCheck(name, fn) {
    var t0 = Date.now();
    try {
      var meta = fn() || {};
      report.checks.push({ name: name, ok: true, ms: Date.now() - t0, meta: meta });
    } catch (e) {
      report.ok = false;
      report.checks.push({
        name: name,
        ok: false,
        ms: Date.now() - t0,
        error: (e && e.message) ? String(e.message) : String(e)
      });
    }
  }

  pushCheck("dashboard_list", function() {
    var list = listQuotesForDashboard("", "", 20, adminPassword);
    return { count: Array.isArray(list) ? list.length : 0 };
  });
  pushCheck("material_search", function() {
    var list = searchMaterials(materialQuery, adminPassword, { limit: 20 });
    return { count: Array.isArray(list) ? list.length : 0 };
  });
  pushCheck("template_list", function() {
    var list = listTemplateCatalog(templateCategory, "", false, adminPassword);
    return { count: Array.isArray(list) ? list.length : 0 };
  });
  if (quoteId && token) {
    pushCheck("quote_detail", function() {
      var detail = getQuote(quoteId, token, { include_notes: false, include_photos: false, note_limit: 20 });
      return {
        has_quote: !!(detail && detail.quote),
        item_count: Number(detail && detail.items && detail.items.length || 0)
      };
    });
  }

  return report;
}

function benchmarkFnMs_(rounds, fn) {
  var samples = [];
  var errors = [];
  var runs = Math.max(Number(rounds || 1), 1);

  for (var i = 0; i < runs; i++) {
    var t0 = Date.now();
    try {
      fn(i + 1);
      samples.push(Date.now() - t0);
    } catch (e) {
      errors.push((e && e.message) ? String(e.message) : String(e));
    }
  }

  return {
    samples: samples,
    summary: summarizeSamples_(samples),
    errors: errors
  };
}

function summarizeSamples_(samples) {
  var arr = Array.isArray(samples) ? samples.slice() : [];
  if (!arr.length) {
    return { count: 0, min: null, p50: null, avg: null, max: null };
  }

  arr.sort(function(a, b) { return a - b; });
  var sum = 0;
  for (var i = 0; i < arr.length; i++) sum += Number(arr[i] || 0);

  var mid = Math.floor(arr.length / 2);
  var p50 = (arr.length % 2 === 1) ? arr[mid] : Math.round((arr[mid - 1] + arr[mid]) / 2);

  return {
    count: arr.length,
    min: arr[0],
    p50: p50,
    avg: Math.round(sum / arr.length),
    max: arr[arr.length - 1]
  };
}

function runPerfValidationEditorQuick() {
  ensureCoreSchemaReady_();
  var pw = String(getSettings_().admin_password || "").trim();
  if (!pw) throw new Error("Settings.admin_password is empty");

  var report = runPerfValidation(pw, { rounds: 3, material_query: "LG" });
  console.log(JSON.stringify(report));
  return report;
}

function test_bootstrapConfigOnce_() {
  var res = bootstrapConfigOnce_({ force: false });
  var props = PropertiesService.getScriptProperties();
  var hasWebhook = !!String(props.getProperty("SLACK_WEBHOOK_URL") || "").trim();
  var hasMemberMap = !!String(props.getProperty("SLACK_MEMBER_MAP") || "").trim();
  var ss = getSpreadsheet_();
  var hasSlackUsers = !!ss.getSheetByName("SlackUsers");
  var hasDedup = !!ss.getSheetByName("NotificationDedup");
  var hasSlackQueue = !!ss.getSheetByName("SlackQueue");
  if (!hasMemberMap || !hasSlackUsers || !hasDedup || !hasSlackQueue) {
    throw new Error("bootstrapConfigOnce_ verification failed");
  }
  return {
    ok: true,
    bootstrap: res,
    hasWebhook: hasWebhook,
    hasMemberMap: hasMemberMap,
    hasSlackUsers: hasSlackUsers,
    hasNotificationDedup: hasDedup,
    hasSlackQueue: hasSlackQueue
  };
}

function test_memo_sends_once() {
  var quoteId = "TEST_MEMO_ONCE_" + Date.now();
  var memo = "memo test";
  var dedupKey = quoteId + ":" + EVENT_CHANGE_REQUEST_MEMO_ + ":" + sha256Hex_(normalizeMemoForDedup_(memo));
  var now = new Date();

  var first = dedupShouldSendAndStore_(quoteId, EVENT_CHANGE_REQUEST_MEMO_, dedupKey, now);
  if (!first) throw new Error("First memo send should pass dedup");

  var body = buildSlackMessage_(EVENT_CHANGE_REQUEST_MEMO_, {
    quoteId: quoteId,
    customerName: "Test Customer",
    customerPhone: "010-0000-0000",
    assigneeDisplay: "Assignee: Test",
    memo: memo,
    deepLink: "https://example.com",
    eventTimeIso: formatEventIso_(now)
  });
  if (!body || !body.text || !body.blocks || !body.blocks.length) {
    throw new Error("Slack memo message build failed");
  }

  return { ok: true, dedupFirst: first, blockCount: body.blocks.length };
}

function test_memo_double_submit_dedup() {
  var quoteId = "TEST_MEMO_DUP_" + Date.now();
  var memo = "duplicate memo test";
  var dedupKey = quoteId + ":" + EVENT_CHANGE_REQUEST_MEMO_ + ":" + sha256Hex_(normalizeMemoForDedup_(memo));
  var t0 = new Date();

  var first = dedupShouldSendAndStore_(quoteId, EVENT_CHANGE_REQUEST_MEMO_, dedupKey, t0);
  var second = dedupShouldSendAndStore_(quoteId, EVENT_CHANGE_REQUEST_MEMO_, dedupKey, new Date(t0.getTime() + 1000));

  if (!first) throw new Error("First memo submission should send");
  if (second) throw new Error("Second identical memo submission should be deduped");

  return { ok: true, first: first, second: second };
}

function test_approval_toggle_sends_each_transition() {
  var quoteId = "TEST_APPROVAL_TOGGLE_" + Date.now();
  var t1 = Date.now();
  var t2 = t1 + 1;

  var key1 = quoteId + ":" + EVENT_APPROVAL_TOGGLED_ + ":not_approved->approved:" + String(t1);
  var key2 = quoteId + ":" + EVENT_APPROVAL_TOGGLED_ + ":approved->not_approved:" + String(t2);

  var send1 = dedupShouldSendAndStore_(quoteId, EVENT_APPROVAL_TOGGLED_, key1, new Date(t1));
  var send2 = dedupShouldSendAndStore_(quoteId, EVENT_APPROVAL_TOGGLED_, key2, new Date(t2));
  var sendDup = dedupShouldSendAndStore_(quoteId, EVENT_APPROVAL_TOGGLED_, key2, new Date(t2));

  if (!send1 || !send2) throw new Error("Each real transition should send");
  if (sendDup) throw new Error("Identical approval transition should be deduped");

  return { ok: true, firstTransition: send1, secondTransition: send2, duplicateSuppressed: !sendDup };
}

function test_lock_prevents_race() {
  var lock1 = LockService.getScriptLock();
  lock1.waitLock(5000);
  var lock2 = LockService.getScriptLock();
  var secondAcquired = false;

  try {
    secondAcquired = lock2.tryLock(1);
  } finally {
    if (secondAcquired) lock2.releaseLock();
    lock1.releaseLock();
  }

  if (secondAcquired) throw new Error("Race prevention failed: second lock acquired unexpectedly");

  var lock3 = LockService.getScriptLock();
  var thirdAcquired = lock3.tryLock(5000);
  if (!thirdAcquired) throw new Error("Lock did not release correctly");
  lock3.releaseLock();

  return { ok: true, secondAcquired: secondAcquired, thirdAcquired: thirdAcquired };
}

function authorizeExternalRequest_() {
  var res = UrlFetchApp.fetch("https://www.google.com/generate_204", {
    method: "get",
    muteHttpExceptions: true
  });
  return { ok: true, httpCode: Number(res.getResponseCode() || 0) };
}

function test_slackPingMemo_() {
  var hooks = getSlackWebhooks_();
  if (!hooks.memo) throw new Error("SLACK_WEBHOOK_URL.memo is missing");
  return sendSlack_(hooks.memo, {
    text: "[test] memo webhook connectivity",
    blocks: [
      {
        type: "header",
        text: { type: "plain_text", text: "Memo Webhook Test", emoji: true }
      },
      {
        type: "section",
        text: { type: "mrkdwn", text: "Connectivity check from Apps Script" }
      }
    ]
  });
}

function test_slackPingApprove_() {
  var hooks = getSlackWebhooks_();
  if (!hooks.approve) throw new Error("SLACK_WEBHOOK_URL.approve is missing");
  return sendSlack_(hooks.approve, {
    text: "[test] approval webhook connectivity",
    blocks: [
      {
        type: "header",
        text: { type: "plain_text", text: "Approval Webhook Test", emoji: true }
      },
      {
        type: "section",
        text: { type: "mrkdwn", text: "Connectivity check from Apps Script" }
      }
    ]
  });
}

function test_saveCustomerNoteByToken() {
  ensureCoreSchemaReady_();
  var pw = String(getSettings_().admin_password || "").trim();
  if (!pw) throw new Error("Settings.admin_password is empty");

  var created = createQuote(pw);
  var quoteId = String(created.quoteId || "").trim();
  if (!quoteId) throw new Error("Failed to create quote");

  var token = token_();
  updateQuote_(quoteId, {
    share_token: token,
    status: "sent",
    customer_name: "Smoke Customer",
    contact_name: "Smoke Owner",
    contact_phone: "010-1111-2222"
  });

  var memoText = "smoke note " + Date.now();
  var res = saveCustomerNoteByToken(quoteId, token, memoText);
  if (!res || !res.ok || !res.note) throw new Error("saveCustomerNoteByToken failed");

  var list = listCustomerNotesByToken(quoteId, token, 10);
  if (!list || !list.length) throw new Error("Customer note list is empty after save");

  return {
    ok: true,
    quote_id: quoteId,
    note_id: String(res.note.note_id || ""),
    notes_count: list.length
  };
}

function test_approveQuoteByToken_toggle() {
  ensureCoreSchemaReady_();
  var pw = String(getSettings_().admin_password || "").trim();
  if (!pw) throw new Error("Settings.admin_password is empty");

  var created = createQuote(pw);
  var quoteId = String(created.quoteId || "").trim();
  if (!quoteId) throw new Error("Failed to create quote");

  var token = token_();
  updateQuote_(quoteId, {
    share_token: token,
    status: "sent",
    total: 100000,
    customer_name: "Toggle Customer"
  });

  var r1 = approveQuoteByToken(quoteId, token);
  var r2 = toggleQuoteApprovalByToken(quoteId, token, "not_approved");
  var r3 = toggleQuoteApprovalByToken(quoteId, token, "approved");

  if (!r1 || !r1.ok || String(r1.status || "") !== "approved") throw new Error("approve step failed");
  if (!r2 || !r2.ok || String(r2.status || "") === "approved") throw new Error("toggle to not_approved failed");
  if (!r3 || !r3.ok || String(r3.status || "") !== "approved") throw new Error("toggle back to approved failed");

  return {
    ok: true,
    quote_id: quoteId,
    statuses: [r1.status, r2.status, r3.status]
  };
}

function test_materialSearch_perf() {
  ensureCoreSchemaReady_();
  var pw = String(getSettings_().admin_password || "").trim();
  if (!pw) throw new Error("Settings.admin_password is empty");

  var samples = benchmarkFnMs_(5, function() {
    searchMaterials("a", pw, 50);
  });
  return { ok: true, summary: samples.summary, errors: samples.errors };
}

function test_slackQueue_and_worker() {
  ensureCoreSchemaReady_();
  ensureSlackQueueSchemaReady_();

  var payload = {
    quoteId: "TEST_QUEUE_" + Date.now(),
    customerName: "Queue Smoke",
    memo: "queue test"
  };
  var dedupKey = payload.quoteId + ":" + EVENT_CHANGE_REQUEST_MEMO_ + ":" + sha256Hex_(payload.memo);
  var enq = enqueueSlackEvent_(EVENT_CHANGE_REQUEST_MEMO_, payload, dedupKey, new Date());

  var worker = runSlackQueueWorker_();
  return {
    ok: true,
    enqueue: enq,
    worker: worker
  };
}

function test_dedup_and_lock() {
  var memo = test_memo_double_submit_dedup();
  var approval = test_approval_toggle_sends_each_transition();
  var lock = test_lock_prevents_race();
  return { ok: true, memo: memo, approval: approval, lock: lock };
}

function buildTemplateTestItems_(prefix, count) {
  var list = [];
  var n = Math.max(Number(count || 1), 1);
  for (var i = 0; i < n; i++) {
    list.push({
      process: prefix + "_PROC_" + (i + 1),
      material: prefix + "_MAT_" + (i + 1),
      spec: "Spec " + (i + 1),
      qty: i + 1,
      unit: "EA",
      unit_price: 1000 * (i + 1),
      note: "note " + (i + 1)
    });
  }
  return list;
}

function test_createTemplateVersion_() {
  ensureCoreSchemaReady_();
  setupTemplatesSchema_();
  var pw = getAdminPasswordFromSettingsOrThrow_();

  var category = "TEST_TEMPLATE";
  var templateName = "TemplateVersion_" + Date.now();
  var v1 = saveTemplateVersion({
    category: category,
    template_name: templateName,
    note: "v1",
    items: buildTemplateTestItems_("TV1", 2)
  }, pw);
  if (!v1 || !v1.template_id) throw new Error("v1 save failed");
  if (Number(v1.version || 0) !== 1) throw new Error("Expected v1 version=1");

  var v2 = saveTemplateVersion({
    template_id: v1.template_id,
    note: "v2",
    items: buildTemplateTestItems_("TV2", 3)
  }, pw);
  if (Number(v2.version || 0) !== 2) throw new Error("Expected v2 version=2");

  var detail = getTemplateDetail(v1.template_id, pw);
  if (!detail || !detail.template || !detail.versions) throw new Error("Template detail missing");
  if (!detail.versions.length || Number(detail.versions[0].version || 0) !== 2) {
    throw new Error("Version history is invalid");
  }

  return {
    ok: true,
    template_id: v1.template_id,
    versions: detail.versions.map(function(x) { return Number(x.version || 0); })
  };
}

function test_importTemplateIntoQuote_replace_() {
  ensureCoreSchemaReady_();
  setupTemplatesSchema_();
  var pw = getAdminPasswordFromSettingsOrThrow_();

  var created = createQuote(pw);
  var qid = String(created.quoteId || "").trim();
  if (!qid) throw new Error("Quote creation failed");

  saveQuote({
    quoteId: qid,
    quote: { customer_name: "Replace Test" },
    items: [{
      process: "EXISTING_PROCESS",
      material: "EXISTING_MATERIAL",
      spec: "EXISTING_SPEC",
      qty: 1,
      unit: "EA",
      unit_price: 5000,
      note: "existing"
    }]
  }, pw);

  var tv = saveTemplateVersion({
    category: "TEST_IMPORT",
    template_name: "ImportReplace_" + Date.now(),
    items: buildTemplateTestItems_("RPL", 2)
  }, pw);

  var res = importTemplateIntoQuote(qid, tv.template_id, tv.version, "replace", pw);
  if (!res || !res.ok) throw new Error("Replace import failed");

  var items = findItems_(qid);
  if (items.length !== 2) throw new Error("Replace import expected 2 items, got " + items.length);
  if (String(items[0].process || "").indexOf("RPL_PROC_1") < 0) {
    throw new Error("Replace import did not apply template items");
  }

  return {
    ok: true,
    quote_id: qid,
    template_id: tv.template_id,
    imported_count: res.imported_count,
    total_item_count: items.length
  };
}

function test_importTemplateIntoQuote_append_() {
  ensureCoreSchemaReady_();
  setupTemplatesSchema_();
  var pw = getAdminPasswordFromSettingsOrThrow_();

  var created = createQuote(pw);
  var qid = String(created.quoteId || "").trim();
  if (!qid) throw new Error("Quote creation failed");

  saveQuote({
    quoteId: qid,
    quote: { customer_name: "Append Test" },
    items: [{
      process: "BASE_PROCESS",
      material: "BASE_MATERIAL",
      spec: "BASE_SPEC",
      qty: 1,
      unit: "EA",
      unit_price: 7000,
      note: "base"
    }]
  }, pw);

  var tv = saveTemplateVersion({
    category: "TEST_IMPORT",
    template_name: "ImportAppend_" + Date.now(),
    items: buildTemplateTestItems_("APD", 2)
  }, pw);

  var res = importTemplateIntoQuote(qid, tv.template_id, tv.version, "append", pw);
  if (!res || !res.ok) throw new Error("Append import failed");

  var items = findItems_(qid);
  if (items.length !== 3) throw new Error("Append import expected 3 items, got " + items.length);
  if (String(items[0].process || "") !== "BASE_PROCESS") {
    throw new Error("Append import changed existing first row unexpectedly");
  }

  return {
    ok: true,
    quote_id: qid,
    template_id: tv.template_id,
    imported_count: res.imported_count,
    total_item_count: items.length
  };
}

function test_saveQuoteAsTemplateVersion_() {
  ensureCoreSchemaReady_();
  setupTemplatesSchema_();
  var pw = getAdminPasswordFromSettingsOrThrow_();

  var created = createQuote(pw);
  var qid = String(created.quoteId || "").trim();
  if (!qid) throw new Error("Quote creation failed");

  saveQuote({
    quoteId: qid,
    quote: { customer_name: "SaveAsTemplate Test" },
    items: buildTemplateTestItems_("QSV", 2)
  }, pw);

  var first = saveQuoteAsTemplateVersion({
    quote_id: qid,
    category: "TEST_SAVE_QUOTE",
    template_name: "SaveQuote_" + Date.now(),
    note: "from quote"
  }, pw);
  if (!first || !first.template_id || Number(first.version || 0) !== 1) {
    throw new Error("First saveQuoteAsTemplateVersion failed");
  }

  var second = saveQuoteAsTemplateVersion({
    quote_id: qid,
    template_id: first.template_id,
    note: "second version"
  }, pw);
  if (Number(second.version || 0) !== 2) throw new Error("Second template version should be 2");

  var v2 = getTemplateVersionDetail(first.template_id, 2, pw);
  if (!v2 || !Array.isArray(v2.items) || v2.items.length < 1) {
    throw new Error("Saved template v2 has no items");
  }

  return {
    ok: true,
    quote_id: qid,
    template_id: first.template_id,
    versions: [Number(first.version || 0), Number(second.version || 0)],
    v2_item_count: v2.items.length
  };
}

function test_deleteTemplate_() {
  ensureCoreSchemaReady_();
  setupTemplatesSchema_();
  var pw = getAdminPasswordFromSettingsOrThrow_();

  var saved = saveTemplateVersion({
    category: "TEST_DELETE",
    template_name: "DeleteTemplate_" + Date.now(),
    items: buildTemplateTestItems_("DEL", 2)
  }, pw);
  if (!saved || !saved.template_id) throw new Error("Template create failed");

  saveTemplateVersion({
    template_id: saved.template_id,
    items: buildTemplateTestItems_("DEL2", 1)
  }, pw);

  var before = getTemplateDetail(saved.template_id, pw);
  if (!before || !before.versions || before.versions.length < 2) {
    throw new Error("Precondition failed: expected at least 2 versions");
  }

  var deleted = deleteTemplate(saved.template_id, pw);
  if (!deleted || !deleted.ok) throw new Error("deleteTemplate failed");

  var missingOk = false;
  try {
    getTemplateDetail(saved.template_id, pw);
  } catch (e) {
    missingOk = String(e && e.message || "").indexOf("Template not found") >= 0;
  }
  if (!missingOk) throw new Error("Deleted template should not be readable");

  return {
    ok: true,
    template_id: saved.template_id,
    deleted_versions: Number(deleted.deleted_versions || 0)
  };
}

function test_templateListCaching_() {
  ensureCoreSchemaReady_();
  setupTemplatesSchema_();
  var pw = getAdminPasswordFromSettingsOrThrow_();

  var category = "TEST_CACHE";
  var saved = saveTemplateVersion({
    category: category,
    template_name: "CacheTemplate_" + Date.now(),
    items: buildTemplateTestItems_("CACHE", 1)
  }, pw);
  if (!saved || !saved.template_id) throw new Error("Template create for cache test failed");

  var key1 = templateCatalogListCacheKey_(category, "", false);
  removeCacheKey_(key1);
  listTemplateCatalog(category, "", false, pw);
  var cachedRaw = CacheService.getScriptCache().get(key1);
  if (!cachedRaw) throw new Error("Template list cache not populated");

  var versionBefore = getTemplateListVersion_();
  saveTemplateVersion({
    template_id: saved.template_id,
    items: buildTemplateTestItems_("CACHE2", 1)
  }, pw);
  var versionAfter = getTemplateListVersion_();
  if (String(versionBefore) === String(versionAfter)) {
    throw new Error("Template list version did not bump after save");
  }

  var key2 = templateCatalogListCacheKey_(category, "", false);
  if (key1 === key2) throw new Error("Cache key must change after version bump");

  var list = listTemplateCatalog(category, "", false, pw);
  var found = false;
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].template_id || "") === String(saved.template_id)) {
      found = true;
      break;
    }
  }
  if (!found) throw new Error("Template missing from cached list");

  return {
    ok: true,
    key_before: key1,
    key_after: key2,
    version_before: versionBefore,
    version_after: versionAfter
  };
}

function test_itemsCanonicalShape_() {
  ensureCoreSchemaReady_();
  var item = normalizeItems_([{
    quote_id: "Q_CANONICAL",
    group_label: "Floor",
    group_code: "FLOOR",
    name: "Tile",
    detail: "600x600",
    qty: 2,
    unit: "EA",
    unit_price: 1000
  }])[0];
  for (var i = 0; i < ITEMS_CANONICAL_HEADERS_.length; i++) {
    if (!Object.prototype.hasOwnProperty.call(item, ITEMS_CANONICAL_HEADERS_[i])) {
      throw new Error("Missing canonical field: " + ITEMS_CANONICAL_HEADERS_[i]);
    }
  }
  return { ok: true, item: item };
}

function test_prequoteMaterialRecommendation_() {
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  var pw = getAdminPasswordFromSettingsOrThrow_();
  var stamp = String(Date.now());
  var materialRes = saveMaterial({
    name: "PrequoteMat_" + stamp,
    brand: "TestBrand",
    spec: "600x600",
    unit: "EA",
    unit_price: 12345,
    is_active: "Y",
    is_representative: "Y",
    material_type: "TILE",
    trade_code: "FLOOR",
    space_type: "RESIDENTIAL",
    expose_to_prequote: "Y",
    recommendation_score_base: 5,
    price_band: "MID",
    tags_summary: "MODERN, GRAY"
  }, null, null, pw);
  var materialId = String(materialRes && materialRes.material && materialRes.material.material_id || "").trim();
  if (!materialId) throw new Error("Material create failed");

  try {
    replaceMaterialTagsAdmin(materialId, [
      { tag_type: "mood", tag_value: "MODERN", weight: 2, is_active: "Y" },
      { tag_type: "tone", tag_value: "GRAY", weight: 1, is_active: "Y" },
      { tag_type: "feature", tag_value: "EASY_CLEAN", weight: 1, is_active: "Y" }
    ], pw);

    var rec = getPrequoteMaterialRecommendation({
      trade_codes: ["FLOOR"],
      space_type: "RESIDENTIAL",
      mood_tags: ["MODERN"],
      tone_tags: ["GRAY"],
      feature_tags: ["EASY_CLEAN"],
      limit: 5
    });
    var list = rec && rec.matched_materials || [];
    var hit = null;
    for (var i = 0; i < list.length; i++) {
      if (String(list[i].material_id || "") === materialId) {
        hit = list[i];
        break;
      }
    }
    if (!hit) throw new Error("Recommendation did not include saved material");
    if (Number(hit.score || 0) <= 0) throw new Error("Recommendation score should be positive");
    return { ok: true, material_id: materialId, score: hit.score, primary_reason_tags: hit.primary_reason_tags || [] };
  } finally {
    try { replaceMaterialTagsAdmin(materialId, [], pw); } catch (e0) {}
    try { deleteMaterial(materialId, false, pw); } catch (e1) {}
  }
}

function test_materialCatalogAdminSaveWithTags_() {
  ensureCoreSchemaReady_();
  var pw = getAdminPasswordFromSettingsOrThrow_();
  var stamp = String(Date.now());
  var saved = saveMaterialWithTagsAdmin({
    name: "CatalogAdmin_" + stamp,
    brand: "AdminBrand",
    spec: "600x600",
    unit: "EA",
    unit_price: 10000,
    note: "admin save smoke",
    is_active: "Y",
    is_representative: "Y",
    expose_to_prequote: "Y",
    material_type: "TILE",
    trade_code: "FLOOR",
    space_type: "BOTH",
    recommendation_score_base: 3,
    price_band: "MID",
    recommendation_note: "smoke test",
    tags_summary: ""
  }, [
    { tag_type: "trade", tag_value: "FLOOR", weight: 2, is_active: "Y" },
    { tag_type: "mood", tag_value: "MODERN", weight: 1, is_active: "Y", note: "seed" },
    { tag_type: "feature", tag_value: "EASY_CLEAN", weight: 1, is_active: "N", source: "catalog_ui" }
  ], pw);

  var materialId = String(saved && saved.material && saved.material.material_id || "").trim();
  if (!materialId) throw new Error("saveMaterialWithTagsAdmin failed");

  try {
    var detail = getMaterialDetailAdmin(materialId, pw);
    if (!detail || !detail.material || String(detail.material.material_id || "") !== materialId) {
      throw new Error("getMaterialDetailAdmin returned unexpected material");
    }
    if (!Array.isArray(detail.tags) || detail.tags.length !== 3) {
      throw new Error("Expected 3 tags after save");
    }

    var single = upsertMaterialTagAdmin({
      material_id: materialId,
      tag_type: "tone",
      tag_value: "GRAY",
      weight: 2,
      is_active: "Y",
      source: "catalog_ui"
    }, pw);
    if (!single || String(single.tag_value || "") !== "GRAY") throw new Error("upsertMaterialTagAdmin failed");

    var deleted = deleteMaterialTagAdmin(materialId, "feature", "EASY_CLEAN", pw);
    if (!deleted || Number(deleted.deleted || 0) < 1) throw new Error("deleteMaterialTagAdmin failed");

    var list = listMaterialTagsAdmin(materialId, pw);
    if (!Array.isArray(list) || list.length !== 3) {
      throw new Error("Expected 3 tags after delete/upsert flow");
    }

    return {
      ok: true,
      material_id: materialId,
      tag_count: list.length,
      tags_summary: detail.generated_tags_summary || ""
    };
  } finally {
    try { replaceMaterialTagsAdmin(materialId, [], pw); } catch (e0) {}
    try { deleteMaterial(materialId, false, pw); } catch (e1) {}
  }
}

function test_prequoteTemplateCatalog_() {
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  var pw = getAdminPasswordFromSettingsOrThrow_();
  var saved = saveTemplateVersion({
    category: "PREQUOTE_TEST",
    template_name: "PrequoteTemplate_" + Date.now(),
    template_meta: {
      template_type: "PREQUOTE_PACKAGE",
      housing_type: "APARTMENT",
      area_band: "20_29",
      budget_band: "MID",
      style_tags_summary: "MODERN, MINIMAL",
      tone_tags_summary: "BEIGE, LIGHT_WOOD",
      trade_scope_summary: "FLOOR, BATHROOM",
      expose_to_prequote: "Y",
      prequote_priority: 7,
      sort_order: 10,
      target_customer_summary: "신혼부부",
      recommended_for_summary: "타일 상담 초안",
      recommendation_note: "운영자 메모",
      is_featured_prequote: "Y"
    },
    items: buildTemplateTestItems_("PREQ", 2)
  }, pw);
  if (!saved || !saved.template_id) throw new Error("Template create failed");

  try {
    var list = listTemplateCatalogForPrequote({
      housing_type: "APARTMENT",
      area_band: "20_29",
      style_tags: ["MODERN"],
      tone_tags: ["BEIGE"],
      trade_scope: ["FLOOR"],
      limit: 10
    });
    var found = false;
    for (var i = 0; i < list.length; i++) {
      if (String(list[i].template_id || "") === String(saved.template_id || "")) {
        found = true;
        break;
      }
    }
    if (!found) throw new Error("Prequote template list missing saved template");
    var detail = getTemplatePackageDetailForPrequote(saved.template_id);
    if (!detail || !Array.isArray(detail.items) || !detail.items.length) {
      throw new Error("Prequote template detail missing items");
    }
    if (!detail.template || !detail.template.template_meta || !detail.template.template_meta.is_featured_prequote) {
      throw new Error("Prequote template detail missing template meta");
    }
    return { ok: true, template_id: saved.template_id, item_count: detail.items.length };
  } finally {
    try { deleteTemplate(saved.template_id, pw); } catch (e) {}
  }
}

function test_templateMetaRoundTrip_() {
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  var pw = getAdminPasswordFromSettingsOrThrow_();
  var saved = saveTemplateVersion({
    category: "META_TEST",
    template_name: "MetaRoundTrip_" + Date.now(),
    note: "meta smoke",
    template_meta: {
      template_type: "BOTH",
      housing_type: "OFFICETEL",
      area_band: "30_39",
      budget_band: "PREMIUM",
      style_tags_summary: "MODERN|WARM",
      tone_tags_summary: "GRAY|BEIGE",
      trade_scope_summary: "KITCHEN|FLOOR",
      expose_to_prequote: "Y",
      prequote_priority: 9,
      sort_order: 3,
      target_customer_summary: "1인 가구",
      recommended_for_summary: "상담 제안용",
      recommendation_note: "내부 메모",
      is_featured_prequote: "Y"
    },
    items: buildTemplateTestItems_("META", 2)
  }, pw);
  if (!saved || !saved.template_id) throw new Error("Template meta save failed");

  try {
    var detail = getTemplateDetail(saved.template_id, pw);
    var meta = detail && detail.template_meta || {};
    if (String(meta.template_type || "") !== "BOTH") throw new Error("template_type round-trip failed");
    if (String(meta.housing_type || "") !== "OFFICETEL") throw new Error("housing_type round-trip failed");
    if (String(meta.tone_tags_summary || "") !== "GRAY|BEIGE") throw new Error("tone_tags_summary round-trip failed");

    var versionDetail = getTemplateVersionDetail(saved.template_id, saved.version, pw);
    if (!versionDetail || !versionDetail.template_meta_snapshot || String(versionDetail.template_meta_snapshot.budget_band || "") !== "PREMIUM") {
      throw new Error("metadata_snapshot_json round-trip failed");
    }

    var filtered = listTemplateCatalog("META_TEST", "", false, pw, {
      template_type: "BOTH",
      expose_to_prequote: "Y",
      housing_type: "OFFICETEL",
      area_band: "30_39"
    });
    if (!filtered || filtered.length !== 1) throw new Error("admin meta filters failed");

    return { ok: true, template_id: saved.template_id, version: saved.version, filtered_count: filtered.length };
  } finally {
    try { deleteTemplate(saved.template_id, pw); } catch (e) {}
  }
}

// Wrappers without trailing underscore for Apps Script run-menu visibility.
function authorizeExternalRequest() {
  return authorizeExternalRequest_();
}

function testSlackPingMemo() {
  return test_slackPingMemo_();
}

function testSlackPingApprove() {
  return test_slackPingApprove_();
}

// Debug helpers (visible in Apps Script run-menu).
function debugDoGetTemplates() {
  var out = doGet({ parameter: { page: "templates" } });
  var title = "";
  try { title = String(out.getTitle() || ""); } catch (e1) { title = ""; }
  Logger.log("debugDoGetTemplates.title=" + title);
  return { ok: true, title: title };
}

function debugDoGetDashboard() {
  var out = doGet({ parameter: { page: "dashboard" } });
  var title = "";
  try { title = String(out.getTitle() || ""); } catch (e1) { title = ""; }
  Logger.log("debugDoGetDashboard.title=" + title);
  return { ok: true, title: title };
}

function debugServiceUrl() {
  var runtimeUrl = "";
  var cachedUrl = "";
  try { runtimeUrl = String(ScriptApp.getService().getUrl() || ""); } catch (e1) { runtimeUrl = ""; }
  try { cachedUrl = String(getAppUrl_() || ""); } catch (e2) { cachedUrl = ""; }
  Logger.log("debugServiceUrl.runtime=" + runtimeUrl);
  Logger.log("debugServiceUrl.cached=" + cachedUrl);
  return {
    ok: true,
    runtime_url: runtimeUrl,
    cached_app_url: cachedUrl
  };
}

function resetAppUrlCache() {
  try { CacheService.getScriptCache().remove("APP_URL_V2"); } catch (e1) {}
  try { PropertiesService.getScriptProperties().deleteProperty("WEBAPP_URL_CACHE"); } catch (e2) {}
  return {
    ok: true,
    runtime_url: String(ScriptApp.getService().getUrl() || ""),
    app_url_after_reset: String(getAppUrl_() || "")
  };
}
