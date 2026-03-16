function setupSpreadsheetIdFromManual(spreadsheetId) {
  const props = PropertiesService.getScriptProperties();
  const id = String(spreadsheetId || props.getProperty("BASE_DB_SPREADSHEET_ID") || "").trim();
  if (!id) {
    throw new Error("spreadsheetId required. Pass an ID or set ScriptProperties.BASE_DB_SPREADSHEET_ID.");
  }
  props.setProperty("SPREADSHEET_ID", id);
  __SS_CACHE_BY_ID = Object.create(null);
  __SHEET_CACHE_BY_SS_ID = Object.create(null);
  invalidateSettingsCache_(id);
  return id;
}

function setupSpreadsheetId() {
  const id = SpreadsheetApp.getActiveSpreadsheet().getId();
  PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", id);
  __SS_CACHE_BY_ID = Object.create(null);
  __SHEET_CACHE_BY_SS_ID = Object.create(null);
  invalidateSettingsCache_(id);
  return id;
}

function setupSpreadsheetId_() {
  return setupSpreadsheetId();
}

var __SS_CACHE_BY_ID = Object.create(null);
var __SHEET_CACHE_BY_SS_ID = Object.create(null);
var __SETTINGS_CACHE_KEY_PREFIX = "SETTINGS_V2";
var ADMIN_SESSION_PREFIX_ = "AS_";
var ADMIN_SESSION_CACHE_PREFIX_ = "ADMSESS_";
var ADMIN_SESSION_TTL_SEC_ = 21600;

function getSpreadsheetId_(overrideId) {
  var id = String(overrideId || "").trim();
  if (id) return id;
  id = String(PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID") || "").trim();
  if (!id) throw new Error("SPREADSHEET_ID missing. Run setupSpreadsheetId().");
  return id;
}

function getSpreadsheet_(overrideId) {
  var id = getSpreadsheetId_(overrideId);
  if (__SS_CACHE_BY_ID[id]) return __SS_CACHE_BY_ID[id];
  __SS_CACHE_BY_ID[id] = SpreadsheetApp.openById(id);
  return __SS_CACHE_BY_ID[id];
}

function getSheetFromSs_(ss, name) {
  var key = String(name || "").trim();
  if (!key) throw new Error("Sheet name required");
  if (!ss || typeof ss.getSheetByName !== "function") {
    throw new Error("Spreadsheet instance required");
  }
  var ssId = String(ss.getId() || "").trim();
  if (!ssId) throw new Error("Spreadsheet ID missing");

  if (!__SHEET_CACHE_BY_SS_ID[ssId]) __SHEET_CACHE_BY_SS_ID[ssId] = Object.create(null);
  var cacheForSs = __SHEET_CACHE_BY_SS_ID[ssId];
  if (cacheForSs[key]) return cacheForSs[key];

  var sh = ss.getSheetByName(key);
  if (!sh) throw new Error("Missing sheet: " + key);
  cacheForSs[key] = sh;
  return sh;
}

function getSheet_(name, overrideId) {
  return getSheetFromSs_(getSpreadsheet_(overrideId), name);
}

function invalidateSheetCache_(name, overrideId) {
  var key = String(name || "").trim();
  if (!key) return;

  var explicitId = String(overrideId || "").trim();
  if (explicitId) {
    if (__SHEET_CACHE_BY_SS_ID[explicitId]) delete __SHEET_CACHE_BY_SS_ID[explicitId][key];
    return;
  }

  for (var ssId in __SHEET_CACHE_BY_SS_ID) {
    if (!Object.prototype.hasOwnProperty.call(__SHEET_CACHE_BY_SS_ID, ssId)) continue;
    delete __SHEET_CACHE_BY_SS_ID[ssId][key];
  }
}

function settingsCacheKey_(spreadsheetId) {
  return __SETTINGS_CACHE_KEY_PREFIX + "_" + String(spreadsheetId || "").trim();
}

function getCachedSettings_(spreadsheetId) {
  var id = String(spreadsheetId || "").trim();
  if (!id) return null;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(settingsCacheKey_(id));
  if (!cached) return null;
  try { return JSON.parse(cached); } catch (e) { return null; }
}

function readSettingsMap_(ss) {
  var spreadsheet = ss || getSpreadsheet_();
  var ssId = String(spreadsheet.getId() || "").trim();
  var cachedMap = getCachedSettings_(ssId);
  if (cachedMap) return cachedMap;

  var sh = spreadsheet.getSheetByName("Settings");
  if (!sh) {
    sh = spreadsheet.insertSheet("Settings");
    sh.getRange(1, 1, 1, 2).setValues([["key", "value"]]);
    invalidateSheetCache_("Settings", ssId);
  }

  var lastRow = sh.getLastRow();
  var map = {};
  if (lastRow >= 2) {
    var values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < values.length; i++) {
      var k = String(values[i][0] || "").trim();
      if (!k) continue;
      map[k] = String(values[i][1] || "").trim();
    }
  }

  try {
    CacheService.getScriptCache().put(settingsCacheKey_(ssId), JSON.stringify(map), 120);
  } catch (e) {}
  return map;
}

function getSettings_() {
  return readSettingsMap_(getSpreadsheet_());
}

function getSettingFromSs_(ss, key, defaultValue) {
  var k = String(key || "").trim();
  if (!k) return defaultValue;
  var map = readSettingsMap_(ss);
  if (!Object.prototype.hasOwnProperty.call(map, k)) return defaultValue;
  var value = map[k];
  if (value === undefined || value === null || String(value).trim() === "") return defaultValue;
  return value;
}

function invalidateSettingsCache_(overrideId) {
  var id = "";
  try { id = getSpreadsheetId_(overrideId); } catch (e) { id = String(overrideId || "").trim(); }
  if (!id) return;
  try { CacheService.getScriptCache().remove(settingsCacheKey_(id)); } catch (e2) {}
}

function uuid_() { return Utilities.getUuid(); }
function token_() { return (Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ""); }
function nowIso_() { return new Date().toISOString(); }

function assertAdmin_(password) {
  const s = getSettings_();
  const expected = String(s.admin_password || "").trim();
  if (!expected) throw new Error("Settings.admin_password missing");

  const got = String(password || "").trim();
  if (!got) throw new Error("Admin password required");
  if (got !== expected) throw new Error("Admin password invalid");
}

function createAdminSessionToken_() {
  return ADMIN_SESSION_PREFIX_ + token_().slice(0, 48);
}

function isAdminSessionToken_(cred) {
  return String(cred || "").trim().indexOf(ADMIN_SESSION_PREFIX_) === 0;
}

function getAdminSessionCacheKey_(token) {
  return ADMIN_SESSION_CACHE_PREFIX_ + String(token || "").trim();
}

function assertAdminCredential_(cred) {
  var raw = String(cred || "").trim();
  if (!raw) throw new Error("Admin credential required");

  if (isAdminSessionToken_(raw)) {
    var cacheKey = getAdminSessionCacheKey_(raw);
    var cacheVal = "";
    try {
      cacheVal = String(CacheService.getScriptCache().get(cacheKey) || "");
    } catch (e) {
      cacheVal = "";
    }
    if (!cacheVal) {
      throw new Error("Admin session expired. Please login again.");
    }
    return { type: "session", token: raw };
  }

  assertAdmin_(raw);
  return { type: "password", password: raw };
}

function getAdminPasswordFromSettingsOrThrow_() {
  const s = getSettings_();
  const pw = String(s.admin_password || "").trim();
  if (!pw) {
    throw new Error("Settings.admin_password is empty. Please set it in the Settings sheet.");
  }
  return pw;
}

function assertEditorAdminExecution_() {
  const activeEmail = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  if (!activeEmail) {
    throw new Error("Editor-only run denied. Open Apps Script editor with an admin Google account and run again.");
  }

  const owners = getOwnerEmailTargets_().map(function(v) {
    return String(v || "").trim().toLowerCase();
  }).filter(function(v) { return !!v; });

  if (owners.length && owners.indexOf(activeEmail) < 0) {
    throw new Error("Editor admin check failed. Add your email to Settings.owner_email, then retry.");
  }
}
