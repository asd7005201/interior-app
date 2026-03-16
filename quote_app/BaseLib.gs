var BaseLib = BaseLib || {};
var __BASELIB_SS_CACHE_BY_ID = __BASELIB_SS_CACHE_BY_ID || Object.create(null);
var BASELIB_SNAPSHOT_STATE_HEADERS_ = ["snapshot_key", "shared_master_version", "built_at", "row_count", "duration_ms", "status", "note"];
var BASELIB_PERF_METRICS_HEADERS_ = ["metric_id", "measured_at", "metric_name", "duration_ms", "entity_id", "actor_type", "actor_id", "context_json", "success_yn", "note"];

BaseLib.buildImageUrl = function(ctx, fileId) {
  var c = baseLibRequireCtx_(ctx);
  var id = String(fileId || "").trim();
  if (!id) return "";
  var ss = baseLibGetSpreadsheet_(c.baseDbId);
  var base = normalizeWebAppBaseUrl_(c.baseUrl || getSettingFromSs_(ss, "base_url", ""));
  if (!base) base = normalizeWebAppBaseUrl_(getAppUrl_());
  return base ? (base + "?page=img&fileId=" + encodeURIComponent(id)) : "";
};

BaseLib.getSharedMaterialMasterVersion = function(ctx) {
  var c = baseLibRequireCtx_(ctx);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  return getSharedMasterVersionInfoFromSs_(baseLibGetSpreadsheet_(c.baseDbId));
};

BaseLib.getSharedMasterVersion = BaseLib.getSharedMaterialMasterVersion;

BaseLib.searchMaterials = function(ctx, query, options) {
  var c = baseLibRequireCtx_(ctx);
  ensureCoreSchemaReady_();
  var ss = baseLibGetSpreadsheet_(c.baseDbId);
  var startMs = Date.now();
  var success = true;
  var note = "";
  var result = [];
  var opts = options || {};
  var q = normText_(String(query || "").trim());
  var mode = String(getSettingFromSs_(ss, "prequote_material_search_mode", "INDEX") || "INDEX").trim().toUpperCase() || "INDEX";
  var source = String(getSettingFromSs_(ss, "baselib_material_search_source", mode) || mode).trim().toUpperCase() || mode;

  try {
    if (!q || q.length < Math.max(Number(opts.min_query_len || opts.minQueryLen || 2), 1)) return [];
    var limit = Math.min(Math.max(Number(opts.limit || 50), 1), 200);
    var page = Math.max(Number(opts.page || 1), 1);
    var offset = opts.offset !== undefined && opts.offset !== null ? Math.max(Math.floor(Number(opts.offset || 0)), 0) : Math.max((page - 1) * limit, 0);
    if (source === "SNAPSHOT") {
      try {
        ensureTemplatesSchemaReady_();
        var tokensFromSnapshot = pickSearchTokens_(tokenize_(q), 8);
        var snapshot = getMaterialMasterSnapshotFromSs_(ss);
        var items = (snapshot && snapshot.materials ? snapshot.materials : []).filter(function(item) {
          var hay = normText_([item.name, item.brand, item.spec, item.note, item.material_type, item.trade_code, item.space_type, item.price_band, item.tags_summary, item.recommendation_note].join(" "));
          for (var i = 0; i < tokensFromSnapshot.length; i++) {
            if (hay.indexOf(tokensFromSnapshot[i]) < 0) return false;
          }
          return true;
        }).sort(function(a, b) {
          if (!!a.is_representative !== !!b.is_representative) return a.is_representative ? -1 : 1;
          return String(b.updated_at || "").localeCompare(String(a.updated_at || ""));
        });
        result = items.slice(offset, offset + limit);
        mode = "SNAPSHOT";
      } catch (snapshotError) {
        note = "snapshot_fallback:" + String(snapshotError);
        source = "INDEX";
      }
    }
    if (!result.length && mode === "FULL_SCAN") {
      result = searchMaterialsFallbackFullScan_(q, offset + limit).slice(offset, offset + limit);
    } else if (!result.length) {
      try {
        var tokens = pickSearchTokens_(tokenize_(q), 8);
        ensureMaterialIndexReady_();
        var refsByToken = getMaterialRefsByTokens_(tokens);
        var refMap = Object.create(null);
        for (var t = 0; t < tokens.length; t++) {
          var refs = parseRefs_(refsByToken[tokens[t]]);
          for (var r = 0; r < refs.length; r++) {
            if (!refMap[refs[r].id] || Number(refs[r].rowNo || 0) > Number(refMap[refs[r].id].rowNo || 0)) refMap[refs[r].id] = refs[r];
          }
        }
        var refsOut = [];
        for (var key in refMap) if (Object.prototype.hasOwnProperty.call(refMap, key)) refsOut.push(refMap[key]);
        refsOut.sort(function(a, b) { return Number(b.rowNo || 0) - Number(a.rowNo || 0); });
        var pre = getMaterialsByRefs_(refsOut, Math.max((offset + limit) * 3, limit * 3));
        var filtered = [];
        for (var i = 0; i < pre.length; i++) {
          var hay = normText_([pre[i].name, pre[i].brand, pre[i].spec, pre[i].note, pre[i].material_type, pre[i].trade_code, pre[i].space_type, pre[i].price_band, pre[i].tags_summary, pre[i].recommendation_note].join(" "));
          var matched = true;
          for (var j = 0; j < tokens.length; j++) {
            if (hay.indexOf(tokens[j]) < 0) {
              matched = false;
              break;
            }
          }
          if (matched) filtered.push(pre[i]);
        }
        filtered.sort(function(a, b) {
          if (!!a.is_representative !== !!b.is_representative) return a.is_representative ? -1 : 1;
          return String(b.updated_at || "").localeCompare(String(a.updated_at || ""));
        });
        result = filtered.slice(offset, offset + limit);
      } catch (e) {
        note = "index_fallback:" + String(e);
        mode = "INDEX_FALLBACK";
        result = searchMaterialsFallbackFullScan_(q, offset + limit).slice(offset, offset + limit);
      }
    }
    result = result.map(function(item) {
      var out = Object.assign({}, item);
      out.image_url = BaseLib.buildImageUrl(c, out.image_file_id) || out.image_url || "";
      out.title = out.name || out.material_id || "";
      out.subtitle = [out.brand, out.spec].filter(Boolean).join(" / ");
      out.source_type = "QUOTE_DB";
      return out;
    });
    return result;
  } catch (e2) {
    success = false;
    note = String(e2);
    throw e2;
  } finally {
    baseLibLogPerfMetricSafe_(ss, "baselib_material_search", Date.now() - startMs, String(query || "").trim(), { mode: mode, source: source, result_count: result.length }, success, note);
  }
};

BaseLib.getMaterialById = function(ctx, materialId) {
  var c = baseLibRequireCtx_(ctx);
  ensureCoreSchemaReady_();
  var row = getMaterialDetailForPrequoteFromSs_(baseLibGetSpreadsheet_(c.baseDbId), materialId);
  if (!row) return null;
  row.image_url = BaseLib.buildImageUrl(c, row.image_file_id) || row.image_url || "";
  row.title = row.name || row.material_id || "";
  row.subtitle = [row.brand, row.spec].filter(Boolean).join(" / ");
  row.source_type = "QUOTE_DB";
  return row;
};

BaseLib.getRepresentativeMaterialsSnapshotForPrequote = function(ctx, filters, options) {
  var c = baseLibRequireCtx_(ctx);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  var ss = baseLibGetSpreadsheet_(c.baseDbId);
  baseLibAssertSnapshotApiEnabled_(ss);
  var startMs = Date.now();
  var settings = getPrequoteSettingsFromSs_(ss);
  var master = getMaterialMasterSnapshotFromSs_(ss);
  var normalized = baseLibNormalizeMaterialFilters_(filters, options, settings);
  var items = (master.materials || []).filter(function(item) {
    return baseLibMaterialMatches_(item, normalized, settings);
  }).map(function(item) {
    var out = formatPrequoteMaterialResult_(item, scoreMaterialForPrequote_(item, normalized));
    out.image_url = BaseLib.buildImageUrl(c, out.image_file_id) || out.image_url || "";
    out.source_type = "QUOTE_DB";
    return out;
  }).slice(0, normalized.limit);
  var shared = getSharedMasterVersionInfoFromSs_(ss);
  baseLibLogSnapshotState_(ss, "prequote_materials_master", shared.shared_master_version, items.length, Date.now() - startMs, "SUCCESS", "");
  baseLibLogPerfMetricSafe_(ss, "prequote_material_snapshot_build", Date.now() - startMs, shared.shared_master_version, { row_count: items.length }, true, "");
  return { shared_master_version: shared.shared_master_version, built_at: nowIso_(), row_count: items.length, items: items, source: "snapshot_api" };
};

BaseLib.listRepresentativeMaterialsForPrequote = function(ctx, filters, options) {
  return BaseLib.getRepresentativeMaterialsSnapshotForPrequote(ctx, filters, options).items;
};

BaseLib.getMaterialDetailForPrequote = function(ctx, materialId) {
  return BaseLib.getMaterialById(ctx, materialId);
};

BaseLib.getTemplateCatalogSnapshotForPrequote = function(ctx, filters, options) {
  var c = baseLibRequireCtx_(ctx);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  var ss = baseLibGetSpreadsheet_(c.baseDbId);
  baseLibAssertSnapshotApiEnabled_(ss);
  var startMs = Date.now();
  var settings = getPrequoteSettingsFromSs_(ss);
  var normalized = baseLibNormalizeTemplateFilters_(filters, options, settings);
  var master = getTemplateCatalogSnapshotFromSs_(ss);
  var items = (master || []).filter(function(item) {
    return baseLibTemplateMatches_(item, normalized, settings);
  }).map(function(item) {
    var out = cloneTemplateCatalogRow_(item);
    out.source_type = "QUOTE_DB";
    return out;
  }).slice(0, normalized.limit);
  var shared = getSharedMasterVersionInfoFromSs_(ss);
  baseLibLogSnapshotState_(ss, "prequote_templates_master", shared.shared_master_version, items.length, Date.now() - startMs, "SUCCESS", "");
  baseLibLogPerfMetricSafe_(ss, "prequote_template_snapshot_build", Date.now() - startMs, shared.shared_master_version, { row_count: items.length }, true, "");
  return { shared_master_version: shared.shared_master_version, built_at: nowIso_(), row_count: items.length, items: items, source: "snapshot_api" };
};

BaseLib.listTemplateCatalogForPrequote = function(ctx, filters, options) {
  return BaseLib.getTemplateCatalogSnapshotForPrequote(ctx, filters, options).items;
};

BaseLib.getTemplateVersionItemsForPrequote = function(ctx, templateId, version) {
  var c = baseLibRequireCtx_(ctx);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  return getTemplateVersionItemsForPrequoteFromSs_(baseLibGetSpreadsheet_(c.baseDbId), templateId, version);
};

BaseLib.getTemplatePackageDetailForPrequote = function(ctx, templateId, options) {
  var c = baseLibRequireCtx_(ctx);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  return getTemplatePackageDetailForPrequoteFromSs_(baseLibGetSpreadsheet_(c.baseDbId), templateId, options || {});
};

BaseLib.getPrequoteMaterialRecommendation = function(ctx, input) {
  var c = baseLibRequireCtx_(ctx);
  ensureCoreSchemaReady_();
  ensureTemplatesSchemaReady_();
  return getPrequoteMaterialRecommendationFromSs_(baseLibGetSpreadsheet_(c.baseDbId), input || {});
};

BaseLib.listMaterialGroups = function(ctx, filters) {
  var c = baseLibRequireCtx_(ctx);
  ensureCoreSchemaReady_();
  return listMaterialGroupsFromSs_(baseLibGetSpreadsheet_(c.baseDbId), filters || {}).map(function(item) {
    var out = Object.assign({}, item);
    if (out.material_preview) out.material_preview.image_url = BaseLib.buildImageUrl(c, out.material_preview.image_file_id) || out.material_preview.image_url || "";
    return out;
  });
};

BaseLib.upsertMaterialGroup = function(ctx, groupObj) {
  var c = baseLibRequireCtx_(ctx);
  ensureCoreSchemaReady_();
  var saved = upsertMaterialGroupInSs_(baseLibGetSpreadsheet_(c.baseDbId), groupObj || {});
  if (saved && saved.material_preview) saved.material_preview.image_url = BaseLib.buildImageUrl(c, saved.material_preview.image_file_id) || saved.material_preview.image_url || "";
  return saved;
};

BaseLib.deleteMaterialGroup = function(ctx, groupKey) {
  var c = baseLibRequireCtx_(ctx);
  ensureCoreSchemaReady_();
  return deleteMaterialGroupFromSs_(baseLibGetSpreadsheet_(c.baseDbId), groupKey);
};

function baseLibRequireCtx_(ctx) {
  var c = ctx || {};
  var baseDbId = String(c.baseDbId || "").trim();
  if (!baseDbId) throw new Error("ctx.baseDbId is required");
  return { baseDbId: baseDbId, baseUrl: String(c.baseUrl || "").trim() };
}

function baseLibGetSpreadsheet_(spreadsheetId) {
  var id = String(spreadsheetId || "").trim();
  if (!id) throw new Error("spreadsheetId is required");
  if (__BASELIB_SS_CACHE_BY_ID[id]) return __BASELIB_SS_CACHE_BY_ID[id];
  __BASELIB_SS_CACHE_BY_ID[id] = SpreadsheetApp.openById(id);
  return __BASELIB_SS_CACHE_BY_ID[id];
}

function baseLibAssertSnapshotApiEnabled_(ss) {
  if (!ynToBool_(getSettingFromSs_(ss, "prequote_allow_snapshot_api", "Y"), true)) throw new Error("Snapshot API is disabled by Settings.prequote_allow_snapshot_api");
}

function baseLibNormalizeMaterialFilters_(filters, options, settings) {
  var raw = Object.assign({}, filters || {}, options || {});
  return { include_inactive: ynToBool_(raw.include_inactive, false), include_unexposed: ynToBool_(raw.include_unexposed, false), representative_only: ynToBool_(raw.representative_only, false), limit: Math.min(Math.max(Number(raw.limit || settings.default_material_limit || 12), 1), 50000) };
}

function baseLibNormalizeTemplateFilters_(filters, options, settings) {
  var raw = Object.assign({}, filters || {}, options || {});
  return Object.assign({}, raw, { include_inactive: ynToBool_(raw.include_inactive, false), include_unexposed: ynToBool_(raw.include_unexposed, false), limit: Math.min(Math.max(Number(raw.limit || settings.default_template_limit || 12), 1), 50000) });
}

function baseLibMaterialMatches_(item, filters, settings) {
  if (!filters.include_inactive && !item.is_active) return false;
  if (!filters.include_unexposed && (!settings.expose_materials_to_prequote || !item.expose_to_prequote)) return false;
  if (filters.representative_only && !item.is_representative) return false;
  return true;
}

function baseLibTemplateMatches_(item, filters, settings) {
  if (!filters.include_inactive && !item.is_active) return false;
  if (!filters.include_unexposed && (!settings.expose_templates_to_prequote || !item.expose_to_prequote)) return false;
  return true;
}

function baseLibLogSnapshotState_(ss, snapshotKey, sharedVersion, rowCount, durationMs, status, note) {
  ensureSheetColumnsInSs_(ss, "SnapshotState", BASELIB_SNAPSHOT_STATE_HEADERS_);
  var sh = ss.getSheetByName("SnapshotState");
  var headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(function(h) { return String(h || "").trim(); });
  var cols = Object.create(null);
  for (var i = 0; i < headers.length; i++) cols[headers[i]] = i;
  var rowNo = findRowNoByColumnValue_(sh, cols.snapshot_key + 1, snapshotKey);
  var row = Array(headers.length).fill("");
  row[cols.snapshot_key] = snapshotKey;
  row[cols.shared_master_version] = sharedVersion;
  row[cols.built_at] = nowIso_();
  row[cols.row_count] = Number(rowCount || 0);
  row[cols.duration_ms] = Number(durationMs || 0);
  row[cols.status] = status;
  row[cols.note] = note || "";
  if (rowNo >= 2) sh.getRange(rowNo, 1, 1, headers.length).setValues([row]);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, headers.length).setValues([row]);
}

function baseLibLogPerfMetricSafe_(ss, metricName, durationMs, entityId, context, success, note) {
  if (!ynToBool_(getSettingFromSs_(ss, "perf_metrics_enabled", "Y"), true)) return;
  ensureSheetColumnsInSs_(ss, "PerfMetrics", BASELIB_PERF_METRICS_HEADERS_);
  appendRowToSheetInSs_(ss, "PerfMetrics", {
    metric_id: Utilities.getUuid(),
    measured_at: nowIso_(),
    metric_name: String(metricName || "").trim(),
    duration_ms: Math.max(Number(durationMs || 0), 0),
    entity_id: String(entityId || "").trim(),
    actor_type: "SYSTEM",
    actor_id: "BaseLib",
    context_json: JSON.stringify(context || {}),
    success_yn: success === false ? "N" : "Y",
    note: String(note || "").trim()
  });
}

function appendRowToSheetInSs_(ss, sheetName, obj) {
  ensureSheetColumnsInSs_(ss, sheetName, BASELIB_PERF_METRICS_HEADERS_);
  var sh = ss.getSheetByName(sheetName);
  var headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(function(h) { return String(h || "").trim(); });
  var row = headers.map(function(header) { return obj && obj[header] !== undefined ? obj[header] : ""; });
  sh.getRange(sh.getLastRow() + 1, 1, 1, headers.length).setValues([row]);
}
