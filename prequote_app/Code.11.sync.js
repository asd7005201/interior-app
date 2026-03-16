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
