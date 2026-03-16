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
