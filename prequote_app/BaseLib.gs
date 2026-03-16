var BaseLib = BaseLib || {};
var __BASELIB_SS_CACHE_BY_ID = __BASELIB_SS_CACHE_BY_ID || Object.create(null);
var __BASELIB_META_CACHE_BY_SS = __BASELIB_META_CACHE_BY_SS || Object.create(null);
var __BASELIB_TOKEN_CACHE_BY_SS = __BASELIB_TOKEN_CACHE_BY_SS || Object.create(null);

var BASELIB_MATERIAL_GROUP_HEADERS_ = ["group_key", "group_label_ko", "space_type", "trade_code", "material_id", "note", "is_active", "updated_at", "mapping_role", "expose_to_prequote", "sort_order", "recommendation_note"];
var BASELIB_MATERIAL_INDEX_HEADERS_ = ["token", "refs", "updated_at"];
var BASELIB_SNAPSHOT_STATE_HEADERS_ = ["snapshot_key", "shared_master_version", "built_at", "row_count", "duration_ms", "status", "note"];
var BASELIB_PERF_METRICS_HEADERS_ = ["metric_id", "measured_at", "metric_name", "duration_ms", "entity_id", "actor_type", "actor_id", "context_json", "success_yn", "note"];

BaseLib.buildImageUrl = function(ctx, fileId) {
  var c = baseLibRequireCtx_(ctx);
  return baseLibBuildImageUrlFromSs_(baseLibGetSpreadsheet_(c.baseDbId), fileId, c.baseUrl);
};

BaseLib.getSharedMaterialMasterVersion = function(ctx) {
  var c = baseLibRequireCtx_(ctx);
  return baseLibGetSharedMasterVersionInfoFromSs_(baseLibGetSpreadsheet_(c.baseDbId));
};

BaseLib.getSharedMasterVersion = BaseLib.getSharedMaterialMasterVersion;

BaseLib.searchMaterials = function(ctx, query, options) {
  var c = baseLibRequireCtx_(ctx);
  var ss = baseLibGetSpreadsheet_(c.baseDbId);
  var settings = baseLibGetPrequoteSettings_(ss);
  var opts = options || {};
  var q = baseLibNormText_(query);
  var startMs = Date.now();
  var success = true;
  var note = "";
  var result = [];
  var mode = String(settings.prequote_material_search_mode || "INDEX").trim().toUpperCase() || "INDEX";

  try {
    if (!q || q.length < Math.max(Number(opts.min_query_len || opts.minQueryLen || 2), 1)) return [];
    var limit = Math.min(Math.max(Number(opts.limit || 50), 1), 200);
    var page = Math.max(Number(opts.page || 1), 1);
    var offset = opts.offset !== undefined && opts.offset !== null ? Math.max(Math.floor(Number(opts.offset || 0)), 0) : Math.max((page - 1) * limit, 0);
    var versionInfo = baseLibGetSharedMasterVersionInfoFromSs_(ss);
    var cacheKey = "BL_MAT_SEARCH_" + baseLibSha256Hex_([String(ss.getId() || ""), versionInfo.material_version, q, limit, offset].join("|")).slice(0, 28);
    var cached = baseLibGetCachedJson_(cacheKey);
    if (cached && Array.isArray(cached)) {
      mode = "QUERY_CACHE";
      result = cached;
      return result;
    }
    if (mode === "FULL_SCAN") {
      result = baseLibSearchMaterialsFullScan_(c, ss, q, offset, limit);
    } else {
      try {
        result = baseLibSearchMaterialsIndexed_(c, ss, q, offset, limit, versionInfo.material_version);
      } catch (e) {
        note = "index_fallback:" + String(e);
        mode = "INDEX_FALLBACK";
        result = baseLibSearchMaterialsFullScan_(c, ss, q, offset, limit);
      }
    }
    baseLibPutCachedJson_(cacheKey, result, 120);
    return result;
  } catch (e2) {
    success = false;
    note = String(e2);
    throw e2;
  } finally {
    baseLibLogPerfMetricSafe_(ss, "baselib_material_search", Date.now() - startMs, String(query || "").trim(), { mode: mode, result_count: result.length }, success, note, "SYSTEM", "BaseLib");
  }
};

BaseLib.getMaterialById = function(ctx, materialId) {
  var c = baseLibRequireCtx_(ctx);
  var targetId = String(materialId || "").trim();
  if (!targetId) throw new Error("materialId is required");
  var snapshot = baseLibGetMaterialMasterSnapshot_(c);
  for (var i = 0; i < snapshot.items.length; i++) {
    if (String(snapshot.items[i].material_id || "").trim() !== targetId) continue;
    return baseLibFinalizeMaterialForCtx_(c, snapshot.items[i]);
  }
  return null;
};

BaseLib.getRepresentativeMaterialsSnapshotForPrequote = function(ctx, filters, options) {
  var c = baseLibRequireCtx_(ctx);
  var ss = baseLibGetSpreadsheet_(c.baseDbId);
  baseLibAssertSnapshotApiEnabled_(ss);
  return baseLibBuildMaterialSnapshotView_(c, filters, options);
};

BaseLib.listRepresentativeMaterialsForPrequote = function(ctx, filters, options) {
  return baseLibBuildMaterialSnapshotView_(baseLibRequireCtx_(ctx), filters, options).items;
};

BaseLib.getMaterialDetailForPrequote = function(ctx, materialId) {
  var c = baseLibRequireCtx_(ctx);
  var targetId = String(materialId || "").trim();
  if (!targetId) throw new Error("materialId is required");
  var snapshot = baseLibGetMaterialMasterSnapshot_(c);
  for (var i = 0; i < snapshot.items.length; i++) {
    if (String(snapshot.items[i].material_id || "").trim() !== targetId) continue;
    return baseLibFinalizeMaterialForCtx_(c, snapshot.items[i]);
  }
  throw new Error("Material not found: " + targetId);
};

BaseLib.getTemplateCatalogSnapshotForPrequote = function(ctx, filters, options) {
  var c = baseLibRequireCtx_(ctx);
  var ss = baseLibGetSpreadsheet_(c.baseDbId);
  baseLibAssertSnapshotApiEnabled_(ss);
  return baseLibBuildTemplateSnapshotView_(c, filters, options);
};

BaseLib.listTemplateCatalogForPrequote = function(ctx, filters, options) {
  return baseLibBuildTemplateSnapshotView_(baseLibRequireCtx_(ctx), filters, options).items;
};

BaseLib.getTemplateVersionItemsForPrequote = function(ctx, templateId, version) {
  var c = baseLibRequireCtx_(ctx);
  var ss = baseLibGetSpreadsheet_(c.baseDbId);
  var targetId = String(templateId || "").trim();
  if (!targetId) throw new Error("templateId is required");
  var rows = baseLibReadAllRows_(ss, "TemplateVersions");
  var wanted = Number(version || 0);
  var latest = 0;
  var match = null;
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (String(row.template_id || "").trim() !== targetId) continue;
    var rowVersion = Number(row.version || 0);
    if (rowVersion > latest) latest = rowVersion;
    if (wanted >= 1 && rowVersion === wanted) {
      match = row;
      break;
    }
  }
  if (!match && wanted < 1 && latest >= 1) {
    for (var j = 0; j < rows.length; j++) {
      if (String(rows[j].template_id || "").trim() === targetId && Number(rows[j].version || 0) === latest) {
        match = rows[j];
        break;
      }
    }
  }
  if (!match) throw new Error("Template version not found: " + targetId);
  return baseLibSafeJsonParse_(match.items_json, []);
};

BaseLib.getTemplatePackageDetailForPrequote = function(ctx, templateId, options) {
  var c = baseLibRequireCtx_(ctx);
  var targetId = String(templateId || "").trim();
  if (!targetId) throw new Error("templateId is required");
  var snapshot = baseLibGetTemplateMasterSnapshot_(c);
  var template = null;
  for (var i = 0; i < snapshot.items.length; i++) {
    if (String(snapshot.items[i].template_id || "").trim() === targetId) {
      template = baseLibClone_(snapshot.items[i]);
      break;
    }
  }
  if (!template) throw new Error("Template not found: " + targetId);
  var wantedVersion = options && options.version !== undefined ? options.version : template.latest_version;
  return { shared_master_version: snapshot.shared_master_version, template: template, items: BaseLib.getTemplateVersionItemsForPrequote(c, targetId, wantedVersion), summary: baseLibSafeJsonParse_(template.summary_json, {}), metadata_snapshot: baseLibSafeJsonParse_(template.metadata_snapshot_json, {}) };
};

BaseLib.getPrequoteMaterialRecommendation = function(ctx, input) {
  var c = baseLibRequireCtx_(ctx);
  var ss = baseLibGetSpreadsheet_(c.baseDbId);
  var settings = baseLibGetPrequoteSettings_(ss);
  var normalized = baseLibNormalizeMaterialFilters_(input || {}, {}, settings);
  return { master_version: baseLibGetSharedMasterVersionInfoFromSs_(ss), input: normalized, matched_materials: BaseLib.listRepresentativeMaterialsForPrequote(c, normalized, { limit: normalized.limit }), generated_at: baseLibNowIso_() };
};

BaseLib.listMaterialGroups = function(ctx, filters) {
  return baseLibListMaterialGroups_(baseLibRequireCtx_(ctx), filters);
};

BaseLib.upsertMaterialGroup = function(ctx, groupObj) {
  return baseLibUpsertMaterialGroup_(baseLibRequireCtx_(ctx), groupObj);
};

BaseLib.deleteMaterialGroup = function(ctx, groupKey) {
  return baseLibDeleteMaterialGroup_(baseLibRequireCtx_(ctx), groupKey);
};

function baseLibBuildMaterialSnapshotView_(ctx, filters, options) {
  var c = baseLibRequireCtx_(ctx);
  var ss = baseLibGetSpreadsheet_(c.baseDbId);
  var settings = baseLibGetPrequoteSettings_(ss);
  var normalized = baseLibNormalizeMaterialFilters_(filters, options, settings);
  var master = baseLibGetMaterialMasterSnapshot_(c);
  var rows = [];
  for (var i = 0; i < master.items.length; i++) {
    if (!baseLibMaterialMatchesFilters_(master.items[i], normalized, settings)) continue;
    rows.push(baseLibFinalizeMaterialForCtx_(c, master.items[i], normalized));
  }
  rows.sort(function(a, b) {
    if (Number(a.score || 0) !== Number(b.score || 0)) return Number(b.score || 0) - Number(a.score || 0);
    if (!!a.is_representative !== !!b.is_representative) return a.is_representative ? -1 : 1;
    if (Number(a.sort_order || 0) !== Number(b.sort_order || 0)) return Number(a.sort_order || 0) - Number(b.sort_order || 0);
    return String(a.material_id || "").localeCompare(String(b.material_id || ""));
  });
  var limited = rows.slice(0, normalized.limit);
  return { shared_master_version: master.shared_master_version, built_at: master.built_at, row_count: limited.length, items: limited, source: "snapshot_api" };
}

function baseLibBuildTemplateSnapshotView_(ctx, filters, options) {
  var c = baseLibRequireCtx_(ctx);
  var ss = baseLibGetSpreadsheet_(c.baseDbId);
  var settings = baseLibGetPrequoteSettings_(ss);
  var normalized = baseLibNormalizeTemplateFilters_(filters, options, settings);
  var master = baseLibGetTemplateMasterSnapshot_(c);
  var rows = [];
  for (var i = 0; i < master.items.length; i++) {
    if (!baseLibTemplateMatchesFilters_(master.items[i], normalized, settings)) continue;
    rows.push(baseLibFinalizeTemplateForCtx_(c, master.items[i]));
  }
  rows.sort(function(a, b) {
    if (!!a.is_featured_prequote !== !!b.is_featured_prequote) return a.is_featured_prequote ? -1 : 1;
    if (Number(a.prequote_priority || 0) !== Number(b.prequote_priority || 0)) return Number(b.prequote_priority || 0) - Number(a.prequote_priority || 0);
    if (Number(a.sort_order || 0) !== Number(b.sort_order || 0)) return Number(a.sort_order || 0) - Number(b.sort_order || 0);
    return String(a.template_name || "").localeCompare(String(b.template_name || ""));
  });
  var limited = rows.slice(0, normalized.limit);
  return { shared_master_version: master.shared_master_version, built_at: master.built_at, row_count: limited.length, items: limited, source: "snapshot_api" };
}

function baseLibGetMaterialMasterSnapshot_(ctx) {
  var c = baseLibRequireCtx_(ctx);
  var ss = baseLibGetSpreadsheet_(c.baseDbId);
  var settings = baseLibGetPrequoteSettings_(ss);
  var versionInfo = baseLibGetSharedMasterVersionInfoFromSs_(ss);
  var cacheKey = "BL_MAT_SNAP_" + baseLibSha256Hex_([String(ss.getId() || ""), versionInfo.shared_master_version].join("|")).slice(0, 28);
  var snapshot = baseLibGetOrBuildCachedJsonChunkedWithLock_(cacheKey, settings.prequote_snapshot_cache_ttl_sec, function() {
    return baseLibBuildMaterialMasterSnapshot_(ss, versionInfo);
  }, { fallback_value: { shared_master_version: versionInfo.shared_master_version, built_at: baseLibNowIso_(), row_count: 0, items: [] } });
  snapshot = snapshot || {};
  snapshot.items = Array.isArray(snapshot.items) ? snapshot.items : [];
  snapshot.shared_master_version = String(snapshot.shared_master_version || versionInfo.shared_master_version || "").trim();
  snapshot.built_at = String(snapshot.built_at || baseLibNowIso_()).trim();
  snapshot.row_count = Math.max(Number(snapshot.row_count || 0), snapshot.items.length);
  return snapshot;
}

function baseLibGetTemplateMasterSnapshot_(ctx) {
  var c = baseLibRequireCtx_(ctx);
  var ss = baseLibGetSpreadsheet_(c.baseDbId);
  var settings = baseLibGetPrequoteSettings_(ss);
  var versionInfo = baseLibGetSharedMasterVersionInfoFromSs_(ss);
  var cacheKey = "BL_TPL_SNAP_" + baseLibSha256Hex_([String(ss.getId() || ""), versionInfo.shared_master_version].join("|")).slice(0, 28);
  var snapshot = baseLibGetOrBuildCachedJsonChunkedWithLock_(cacheKey, settings.prequote_snapshot_cache_ttl_sec, function() {
    return baseLibBuildTemplateMasterSnapshot_(ss, versionInfo);
  }, { fallback_value: { shared_master_version: versionInfo.shared_master_version, built_at: baseLibNowIso_(), row_count: 0, items: [] } });
  snapshot = snapshot || {};
  snapshot.items = Array.isArray(snapshot.items) ? snapshot.items : [];
  snapshot.shared_master_version = String(snapshot.shared_master_version || versionInfo.shared_master_version || "").trim();
  snapshot.built_at = String(snapshot.built_at || baseLibNowIso_()).trim();
  snapshot.row_count = Math.max(Number(snapshot.row_count || 0), snapshot.items.length);
  return snapshot;
}

function baseLibBuildMaterialMasterSnapshot_(ss, versionInfo) {
  var startMs = Date.now();
  var success = true;
  var note = "";
  try {
    var materials = baseLibReadAllRows_(ss, "Materials");
    var tags = [];
    var groups = [];
    try { tags = baseLibReadAllRows_(ss, "MaterialTags"); } catch (e0) { tags = []; }
    try { groups = baseLibReadAllRows_(ss, "MaterialGroups"); } catch (e1) { groups = []; }
    var tagsByMaterial = Object.create(null);
    for (var i = 0; i < tags.length; i++) {
      var tagRow = tags[i];
      if (!baseLibYnToBool_(tagRow.is_active, true)) continue;
      var tagMaterialId = String(tagRow.material_id || "").trim();
      var tagType = String(tagRow.tag_type || "").trim().toLowerCase();
      var tagValue = String(tagRow.tag_value || "").trim();
      if (!tagMaterialId || !tagType || !tagValue) continue;
      if (!tagsByMaterial[tagMaterialId]) tagsByMaterial[tagMaterialId] = { tag_records: [], tags_by_type: Object.create(null) };
      tagsByMaterial[tagMaterialId].tag_records.push({ tag_type: tagType, tag_value: tagValue, weight: baseLibToNumber_(tagRow.weight, 1), source: String(tagRow.source || "").trim(), note: String(tagRow.note || "").trim(), source_ref_type: String(tagRow.source_ref_type || "").trim(), source_ref_key: String(tagRow.source_ref_key || "").trim() });
      if (!tagsByMaterial[tagMaterialId].tags_by_type[tagType]) tagsByMaterial[tagMaterialId].tags_by_type[tagType] = [];
      baseLibPushUnique_(tagsByMaterial[tagMaterialId].tags_by_type[tagType], tagValue);
    }
    var groupsByMaterial = Object.create(null);
    for (var g = 0; g < groups.length; g++) {
      var groupRow = groups[g];
      if (!baseLibYnToBool_(groupRow.is_active, true)) continue;
      var materialId = String(groupRow.material_id || "").trim();
      var groupKey = String(groupRow.group_key || "").trim();
      if (!materialId || !groupKey) continue;
      if (!groupsByMaterial[materialId]) groupsByMaterial[materialId] = [];
      baseLibPushUnique_(groupsByMaterial[materialId], groupKey);
    }
    var items = [];
    for (var m = 0; m < materials.length; m++) {
      var materialRow = materials[m];
      var item = baseLibNormalizeMaterialRow_(materialRow, tagsByMaterial[String(materialRow.material_id || "").trim()] || null, groupsByMaterial[String(materialRow.material_id || "").trim()] || [], ss);
      if (!item.material_id) continue;
      items.push(item);
    }
    var snapshot = { shared_master_version: String(versionInfo.shared_master_version || "").trim(), built_at: baseLibNowIso_(), row_count: items.length, items: items };
    baseLibLogSnapshotState_(ss, "prequote_materials_master", snapshot.shared_master_version, items.length, Date.now() - startMs, "SUCCESS", "");
    return snapshot;
  } catch (e2) {
    success = false;
    note = String(e2);
    baseLibLogSnapshotState_(ss, "prequote_materials_master", String(versionInfo.shared_master_version || "").trim(), 0, Date.now() - startMs, "FAILED", note);
    throw e2;
  } finally {
    baseLibLogPerfMetricSafe_(ss, "prequote_material_snapshot_build", Date.now() - startMs, String(versionInfo.shared_master_version || "").trim(), {}, success, note, "SYSTEM", "BaseLib");
  }
}

function baseLibBuildTemplateMasterSnapshot_(ss, versionInfo) {
  var startMs = Date.now();
  var success = true;
  var note = "";
  try {
    var catalogRows = baseLibReadAllRows_(ss, "TemplateCatalog");
    var versionRows = [];
    try { versionRows = baseLibReadAllRows_(ss, "TemplateVersions"); } catch (e0) { versionRows = []; }
    var latestByTemplateId = Object.create(null);
    var items = [];
    for (var i = 0; i < catalogRows.length; i++) {
      var row = baseLibNormalizeTemplateRow_(catalogRows[i]);
      if (!row.template_id) continue;
      items.push(row);
      latestByTemplateId[row.template_id] = Number(row.latest_version || 0);
    }
    var latestMeta = Object.create(null);
    for (var j = 0; j < versionRows.length; j++) {
      var versionRow = versionRows[j];
      var tid = String(versionRow.template_id || "").trim();
      var wanted = Number(latestByTemplateId[tid] || 0);
      if (!tid || wanted < 1 || Number(versionRow.version || 0) !== wanted) continue;
      latestMeta[tid] = { summary_json: String(versionRow.summary_json || "").trim(), metadata_snapshot_json: String(versionRow.metadata_snapshot_json || "").trim(), item_count: baseLibToNumber_(versionRow.item_count, 0) };
    }
    for (var k = 0; k < items.length; k++) {
      var meta = latestMeta[items[k].template_id] || {};
      items[k].summary_json = meta.summary_json || "";
      items[k].metadata_snapshot_json = meta.metadata_snapshot_json || "";
      if (!items[k].latest_item_count) items[k].latest_item_count = baseLibToNumber_(meta.item_count, 0);
      items[k]._search_text = baseLibBuildTemplateSearchText_(items[k], baseLibSafeJsonParse_(items[k].summary_json, {}));
    }
    var snapshot = { shared_master_version: String(versionInfo.shared_master_version || "").trim(), built_at: baseLibNowIso_(), row_count: items.length, items: items };
    baseLibLogSnapshotState_(ss, "prequote_templates_master", snapshot.shared_master_version, items.length, Date.now() - startMs, "SUCCESS", "");
    return snapshot;
  } catch (e1) {
    success = false;
    note = String(e1);
    baseLibLogSnapshotState_(ss, "prequote_templates_master", String(versionInfo.shared_master_version || "").trim(), 0, Date.now() - startMs, "FAILED", note);
    throw e1;
  } finally {
    baseLibLogPerfMetricSafe_(ss, "prequote_template_snapshot_build", Date.now() - startMs, String(versionInfo.shared_master_version || "").trim(), {}, success, note, "SYSTEM", "BaseLib");
  }
}

function baseLibSearchMaterialsIndexed_(ctx, ss, q, offset, limit, materialVersion) {
  var tokens = baseLibPickSearchTokens_(baseLibTokenize_(q), 8);
  if (!tokens.length) return [];
  baseLibEnsureMaterialsIndexReady_(ss, materialVersion);
  var refsByToken = baseLibGetMaterialRefsByTokens_(ss, materialVersion, tokens);
  var refMap = Object.create(null);
  for (var t = 0; t < tokens.length; t++) {
    var refs = baseLibParseRefs_(refsByToken[tokens[t]]);
    for (var r = 0; r < refs.length; r++) {
      if (!refMap[refs[r].id] || Number(refs[r].rowNo || 0) > Number(refMap[refs[r].id].rowNo || 0)) refMap[refs[r].id] = refs[r];
    }
  }
  var refsOut = [];
  for (var key in refMap) if (Object.prototype.hasOwnProperty.call(refMap, key)) refsOut.push(refMap[key]);
  refsOut.sort(function(a, b) { return Number(b.rowNo || 0) - Number(a.rowNo || 0); });
  var pre = baseLibGetMaterialsByRefs_(ctx, ss, refsOut, Math.max((offset + limit) * 3, limit * 3));
  var filtered = [];
  for (var i = 0; i < pre.length; i++) {
    var hay = baseLibNormText_(baseLibMaterialSearchTextFromObject_(pre[i]));
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
  return filtered.slice(offset, offset + limit);
}

function baseLibSearchMaterialsFullScan_(ctx, ss, q, offset, limit) {
  var rows = baseLibReadAllRows_(ss, "Materials");
  var tokens = baseLibPickSearchTokens_(baseLibTokenize_(q), 8);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var item = baseLibFinalizeMaterialForCtx_(ctx, baseLibNormalizeMaterialRow_(rows[i], null, [], ss));
    if (!item.material_id) continue;
    var hay = baseLibNormText_(baseLibMaterialSearchTextFromObject_(item));
    var matched = true;
    for (var t = 0; t < tokens.length; t++) {
      if (hay.indexOf(tokens[t]) < 0) {
        matched = false;
        break;
      }
    }
    if (matched) out.push(item);
  }
  out.sort(function(a, b) {
    if (!!a.is_representative !== !!b.is_representative) return a.is_representative ? -1 : 1;
    return String(b.updated_at || "").localeCompare(String(a.updated_at || ""));
  });
  return out.slice(offset, offset + limit);
}

function baseLibGetMaterialsByRefs_(ctx, ss, refs, limit) {
  var sh = baseLibGetSheet_(ss, "Materials");
  var headers = baseLibGetHeaders_(ss, "Materials");
  var max = Math.max(Number(limit || 50), 1);
  var orderedRefs = Array.isArray(refs) ? refs : [];
  if (!orderedRefs.length) return [];
  var rowNos = [];
  var seenRow = Object.create(null);
  for (var i = 0; i < orderedRefs.length && rowNos.length < max * 2; i++) {
    var rowNo = Number(orderedRefs[i].rowNo || 0);
    if (rowNo < 2 || seenRow[rowNo]) continue;
    seenRow[rowNo] = 1;
    rowNos.push(rowNo);
  }
  var rows = baseLibReadRowsByNumbers_(sh, headers, rowNos);
  var rowsByNo = Object.create(null);
  for (var j = 0; j < rows.length; j++) rowsByNo[Number(rows[j]._rowNo || 0)] = rows[j];
  var out = [];
  var seenId = Object.create(null);
  for (var r = 0; r < orderedRefs.length && out.length < max; r++) {
    var ref = orderedRefs[r] || {};
    var refId = String(ref.id || "").trim();
    var refRowNo = Number(ref.rowNo || 0);
    if (!refId || refRowNo < 2 || seenId[refId]) continue;
    var row = rowsByNo[refRowNo];
    if (!row || String(row.material_id || "").trim() !== refId) continue;
    seenId[refId] = 1;
    out.push(baseLibFinalizeMaterialForCtx_(ctx, baseLibNormalizeMaterialRow_(row, null, [], ss)));
  }
  return out;
}

function baseLibEnsureMaterialsIndexReady_(ss, materialVersion) {
  baseLibEnsureSheetHeaders_(ss, "MaterialsIndex", BASELIB_MATERIAL_INDEX_HEADERS_);
  var cacheKey = "BL_MIDX_READY_" + String(ss.getId() || "") + "_" + String(materialVersion || "");
  try {
    if (CacheService.getScriptCache().get(cacheKey) === "1" && baseLibGetSheet_(ss, "MaterialsIndex").getLastRow() >= 2) return;
  } catch (e) {}
  if (baseLibGetSheet_(ss, "MaterialsIndex").getLastRow() < 2) baseLibRebuildMaterialsIndex_(ss);
  try { CacheService.getScriptCache().put(cacheKey, "1", 300); } catch (e2) {}
}

function baseLibRebuildMaterialsIndex_(ss) {
  var rows = baseLibReadAllRows_(ss, "Materials");
  var tokenMap = Object.create(null);
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var id = String(row.material_id || "").trim();
    if (!id) continue;
    var tokens = baseLibTokenize_(baseLibMaterialSearchTextFromObject_(row));
    for (var t = 0; t < tokens.length; t++) {
      if (!tokenMap[tokens[t]]) tokenMap[tokens[t]] = [];
      tokenMap[tokens[t]].push({ id: id, rowNo: Number(row._rowNo || 0) });
    }
  }
  var sh = baseLibGetSheet_(ss, "MaterialsIndex");
  var headers = baseLibGetHeaders_(ss, "MaterialsIndex");
  var lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  var tokensOut = Object.keys(tokenMap).sort();
  if (!tokensOut.length) return 0;
  var rowsOut = [];
  var now = baseLibNowIso_();
  for (var i2 = 0; i2 < tokensOut.length; i2++) {
    var rowOut = baseLibBlankRowObject_(headers);
    rowOut.token = tokensOut[i2];
    rowOut.refs = baseLibRefsToString_(tokenMap[tokensOut[i2]].slice(0, 2000));
    rowOut.updated_at = now;
    rowsOut.push(rowOut);
  }
  baseLibAppendRows_(ss, "MaterialsIndex", rowsOut);
  return rowsOut.length;
}

function baseLibGetMaterialRefsByTokens_(ss, materialVersion, tokens) {
  var ssId = String(ss.getId() || "");
  var memoKey = ssId + "|" + String(materialVersion || "");
  if (!__BASELIB_TOKEN_CACHE_BY_SS[memoKey]) __BASELIB_TOKEN_CACHE_BY_SS[memoKey] = Object.create(null);
  var memo = __BASELIB_TOKEN_CACHE_BY_SS[memoKey];
  var cache = CacheService.getScriptCache();
  var out = Object.create(null);
  var missing = [];
  for (var i = 0; i < tokens.length; i++) {
    var token = String(tokens[i] || "").trim();
    if (!token) continue;
    if (Object.prototype.hasOwnProperty.call(memo, token)) {
      out[token] = memo[token];
      continue;
    }
    try {
      var cached = cache.get("BL_MIDXTOK_" + memoKey + "_" + token);
      if (cached !== null && cached !== undefined) {
        memo[token] = cached === "__MISS__" ? "" : cached;
        out[token] = memo[token];
        continue;
      }
    } catch (e0) {}
    missing.push(token);
  }
  if (!missing.length) return out;
  var sh = baseLibGetSheet_(ss, "MaterialsIndex");
  var cols = baseLibGetColMap_(ss, "MaterialsIndex");
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return out;
  var tokenRange = sh.getRange(2, cols.token + 1, lastRow - 1, 1);
  var tokenToRowNo = Object.create(null);
  for (var m = 0; m < missing.length; m++) {
    var finder = tokenRange.createTextFinder(missing[m]);
    finder.matchEntireCell(true);
    var hit = finder.findNext();
    if (hit) tokenToRowNo[missing[m]] = Number(hit.getRow() || 0);
  }
  var rowNos = [];
  for (var n = 0; n < missing.length; n++) {
    var rowNo = Number(tokenToRowNo[missing[n]] || 0);
    if (rowNo >= 2) rowNos.push(rowNo);
  }
  rowNos = baseLibUniqueSortedNumbers_(rowNos);
  var refsByRowNo = Object.create(null);
  if (rowNos.length) {
    var rows = baseLibReadRowsByNumbers_(sh, baseLibGetHeaders_(ss, "MaterialsIndex"), rowNos);
    for (var r = 0; r < rows.length; r++) refsByRowNo[Number(rows[r]._rowNo || 0)] = String(rows[r].refs || "");
  }
  for (var k = 0; k < missing.length; k++) {
    var token2 = missing[k];
    var refsStr = String(refsByRowNo[Number(tokenToRowNo[token2] || 0)] || "");
    memo[token2] = refsStr;
    out[token2] = refsStr;
    try { cache.put("BL_MIDXTOK_" + memoKey + "_" + token2, refsStr || "__MISS__", 600); } catch (e1) {}
  }
  return out;
}

function baseLibListMaterialGroups_(ctx, filters) {
  var c = baseLibRequireCtx_(ctx);
  var ss = baseLibGetSpreadsheet_(c.baseDbId);
  var rows = baseLibReadAllRows_(ss, "MaterialGroups");
  var opts = filters || {};
  var includeInactive = baseLibYnToBool_(opts.include_inactive, false);
  var groupKey = String(opts.group_key || "").trim();
  var materialId = String(opts.material_id || "").trim();
  var tradeCode = String(opts.trade_code || "").trim().toUpperCase();
  var spaceType = String(opts.space_type || "").trim().toUpperCase();
  var materials = baseLibGetMaterialMasterSnapshot_(c).items;
  var materialMap = Object.create(null);
  for (var i = 0; i < materials.length; i++) materialMap[String(materials[i].material_id || "").trim()] = materials[i];
  return rows.filter(function(row) {
    if (!includeInactive && !baseLibYnToBool_(row.is_active, true)) return false;
    if (groupKey && String(row.group_key || "").trim() !== groupKey) return false;
    if (materialId && String(row.material_id || "").trim() !== materialId) return false;
    if (tradeCode && String(row.trade_code || "").trim().toUpperCase() !== tradeCode) return false;
    if (spaceType && String(row.space_type || "").trim().toUpperCase() !== spaceType) return false;
    return !!String(row.group_key || "").trim();
  }).map(function(row) {
    var item = { group_key: String(row.group_key || "").trim(), group_label_ko: String(row.group_label_ko || "").trim(), space_type: String(row.space_type || "").trim(), trade_code: String(row.trade_code || "").trim(), material_id: String(row.material_id || "").trim(), note: String(row.note || "").trim(), is_active: baseLibYnToBool_(row.is_active, true), updated_at: String(row.updated_at || "").trim(), mapping_role: String(row.mapping_role || "").trim(), expose_to_prequote: baseLibYnToBool_(row.expose_to_prequote, true), sort_order: baseLibToNumber_(row.sort_order, 0), recommendation_note: String(row.recommendation_note || "").trim(), material_preview: null };
    if (item.material_id && materialMap[item.material_id]) item.material_preview = baseLibFinalizeMaterialForCtx_(c, materialMap[item.material_id]);
    return item;
  }).sort(function(a, b) {
    if (a.group_key !== b.group_key) return a.group_key.localeCompare(b.group_key);
    if (a.sort_order !== b.sort_order) return a.sort_order - b.sort_order;
    return a.material_id.localeCompare(b.material_id);
  });
}

function baseLibUpsertMaterialGroup_(ctx, groupObj) {
  var c = baseLibRequireCtx_(ctx);
  var ss = baseLibGetSpreadsheet_(c.baseDbId);
  var payload = groupObj || {};
  var groupKey = String(payload.group_key || "").trim();
  var materialId = String(payload.material_id || "").trim();
  if (!groupKey) throw new Error("group_key is required");
  baseLibEnsureSheetHeaders_(ss, "MaterialGroups", BASELIB_MATERIAL_GROUP_HEADERS_);
  var sh = baseLibGetSheet_(ss, "MaterialGroups");
  var meta = baseLibSheetMeta_(ss, "MaterialGroups");
  var rowNo = 0;
  var rowNos = baseLibFindRowNosByColumnValue_(ss, "MaterialGroups", "group_key", groupKey);
  for (var i = 0; i < rowNos.length; i++) {
    var row = baseLibReadRowsByNumbers_(sh, meta.headers, [rowNos[i]])[0];
    if (!row) continue;
    if (materialId && String(row.material_id || "").trim() !== materialId) continue;
    rowNo = rowNos[i];
    break;
  }
  var fields = { group_key: groupKey, group_label_ko: String(payload.group_label_ko || "").trim(), space_type: String(payload.space_type || "").trim(), trade_code: String(payload.trade_code || "").trim(), material_id: materialId, note: String(payload.note || "").trim(), is_active: baseLibYnToBool_(payload.is_active, true) ? "Y" : "N", updated_at: baseLibNowIso_(), mapping_role: String(payload.mapping_role || "").trim(), expose_to_prequote: baseLibYnToBool_(payload.expose_to_prequote, true) ? "Y" : "N", sort_order: baseLibToNumber_(payload.sort_order, 0), recommendation_note: String(payload.recommendation_note || "").trim() };
  if (rowNo >= 2) baseLibUpdateRowFields_(sh, meta, rowNo, fields);
  else rowNo = baseLibAppendRow_(ss, "MaterialGroups", fields);
  return baseLibListMaterialGroups_(c, { group_key: groupKey, material_id: materialId, include_inactive: true })[0] || null;
}

function baseLibDeleteMaterialGroup_(ctx, groupKey) {
  var c = baseLibRequireCtx_(ctx);
  var ss = baseLibGetSpreadsheet_(c.baseDbId);
  var target = String(groupKey || "").trim();
  if (!target) throw new Error("groupKey is required");
  var rowNos = baseLibFindRowNosByColumnValue_(ss, "MaterialGroups", "group_key", target);
  if (!rowNos.length) return { success: true, updated_count: 0 };
  var sh = baseLibGetSheet_(ss, "MaterialGroups");
  var meta = baseLibSheetMeta_(ss, "MaterialGroups");
  var now = baseLibNowIso_();
  for (var i = 0; i < rowNos.length; i++) baseLibUpdateRowFields_(sh, meta, rowNos[i], { is_active: "N", updated_at: now });
  return { success: true, updated_count: rowNos.length };
}

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

function baseLibGetSheet_(ss, sheetName) {
  var sh = ss.getSheetByName(String(sheetName || "").trim());
  if (!sh) throw new Error("Sheet not found: " + sheetName);
  return sh;
}

function baseLibMetaBucket_(ss) {
  var key = String(ss.getId() || "");
  if (!__BASELIB_META_CACHE_BY_SS[key]) __BASELIB_META_CACHE_BY_SS[key] = { headers: Object.create(null), cols: Object.create(null) };
  return __BASELIB_META_CACHE_BY_SS[key];
}

function baseLibInvalidateSheetMeta_(ss, sheetName) {
  var bucket = baseLibMetaBucket_(ss);
  delete bucket.headers[sheetName];
  delete bucket.cols[sheetName];
}

function baseLibEnsureSheetHeaders_(ss, sheetName, headers) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    baseLibInvalidateSheetMeta_(ss, sheetName);
    return sh;
  }
  var currentHeaders = baseLibGetHeaders_(ss, sheetName);
  var changed = false;
  for (var i = 0; i < headers.length; i++) {
    if (currentHeaders.indexOf(headers[i]) >= 0) continue;
    currentHeaders.push(headers[i]);
    changed = true;
  }
  if (changed) sh.getRange(1, 1, 1, currentHeaders.length).setValues([currentHeaders]);
  baseLibInvalidateSheetMeta_(ss, sheetName);
  return sh;
}

function baseLibGetHeaders_(ss, sheetName) {
  var bucket = baseLibMetaBucket_(ss);
  if (bucket.headers[sheetName]) return bucket.headers[sheetName].slice();
  var sh = baseLibGetSheet_(ss, sheetName);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || "").trim(); });
  bucket.headers[sheetName] = headers;
  delete bucket.cols[sheetName];
  return headers.slice();
}

function baseLibGetColMap_(ss, sheetName) {
  var bucket = baseLibMetaBucket_(ss);
  if (bucket.cols[sheetName]) return bucket.cols[sheetName];
  var headers = baseLibGetHeaders_(ss, sheetName);
  var cols = Object.create(null);
  for (var i = 0; i < headers.length; i++) {
    if (!headers[i]) continue;
    if (!Object.prototype.hasOwnProperty.call(cols, headers[i])) cols[headers[i]] = i;
  }
  bucket.cols[sheetName] = cols;
  return cols;
}

function baseLibSheetMeta_(ss, sheetName) {
  return { headers: baseLibGetHeaders_(ss, sheetName), cols: baseLibGetColMap_(ss, sheetName) };
}

function baseLibBlankRowObject_(headers) {
  var row = {};
  for (var i = 0; i < headers.length; i++) row[headers[i]] = "";
  return row;
}

function baseLibRowToObj_(headers, row, rowNo) {
  var obj = { _rowNo: rowNo };
  for (var i = 0; i < headers.length; i++) obj[headers[i]] = row[i];
  return obj;
}

function baseLibReadAllRows_(ss, sheetName) {
  var sh = baseLibGetSheet_(ss, sheetName);
  var headers = baseLibGetHeaders_(ss, sheetName);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) out.push(baseLibRowToObj_(headers, values[i], i + 2));
  return out;
}

function baseLibReadRowsByNumbers_(sh, headers, rowNos) {
  var wanted = baseLibUniqueSortedNumbers_(rowNos);
  if (!wanted.length) return [];
  var rowsByNo = Object.create(null);
  var start = wanted[0];
  var end = wanted[0];
  var chunks = [];
  for (var i = 1; i < wanted.length; i++) {
    if (wanted[i] === end + 1) end = wanted[i];
    else {
      chunks.push({ start: start, end: end });
      start = wanted[i];
      end = wanted[i];
    }
  }
  chunks.push({ start: start, end: end });
  for (var c = 0; c < chunks.length; c++) {
    var chunk = chunks[c];
    var values = sh.getRange(chunk.start, 1, chunk.end - chunk.start + 1, headers.length).getValues();
    for (var r = 0; r < values.length; r++) rowsByNo[chunk.start + r] = baseLibRowToObj_(headers, values[r], chunk.start + r);
  }
  var out = [];
  for (var j = 0; j < wanted.length; j++) if (rowsByNo[wanted[j]]) out.push(rowsByNo[wanted[j]]);
  return out;
}

function baseLibFindRowNosByColumnValue_(ss, sheetName, colName, value) {
  var needle = String(value || "").trim();
  if (!needle) return [];
  var sh = baseLibGetSheet_(ss, sheetName);
  var cols = baseLibGetColMap_(ss, sheetName);
  var colIdx = cols[colName];
  var lastRow = sh.getLastRow();
  if (colIdx === undefined || lastRow < 2) return [];
  var cells = sh.getRange(2, colIdx + 1, lastRow - 1, 1).createTextFinder(needle).matchEntireCell(true).findAll();
  if (!cells || !cells.length) return [];
  return cells.map(function(cell) { return Number(cell.getRow() || 0); }).filter(function(rowNo) { return rowNo >= 2; });
}

function baseLibAppendRow_(ss, sheetName, fields) {
  return baseLibAppendRows_(ss, sheetName, [fields])[0] || 0;
}

function baseLibAppendRows_(ss, sheetName, rows) {
  var list = Array.isArray(rows) ? rows : [];
  if (!list.length) return [];
  var sh = baseLibGetSheet_(ss, sheetName);
  var headers = baseLibGetHeaders_(ss, sheetName);
  var values = [];
  for (var i = 0; i < list.length; i++) {
    var row = [];
    for (var j = 0; j < headers.length; j++) row.push(list[i] && list[i][headers[j]] !== undefined ? list[i][headers[j]] : "");
    values.push(row);
  }
  var startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, values.length, headers.length).setValues(values);
  var rowNos = [];
  for (var r = 0; r < values.length; r++) rowNos.push(startRow + r);
  return rowNos;
}

function baseLibUpdateRowFields_(sh, meta, rowNo, fields) {
  var row = sh.getRange(rowNo, 1, 1, meta.headers.length).getValues()[0];
  var changed = false;
  for (var key in fields) {
    if (!Object.prototype.hasOwnProperty.call(fields, key)) continue;
    var idx = meta.cols[key];
    if (idx === undefined) continue;
    row[idx] = fields[key];
    changed = true;
  }
  if (changed) sh.getRange(rowNo, 1, 1, meta.headers.length).setValues([row]);
}

function baseLibGetSettingFromSs_(ss, key, defaultValue) {
  var sh = ss.getSheetByName("Settings");
  var target = String(key || "").trim();
  if (!sh || !target) return defaultValue;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return defaultValue;
  var values = sh.getRange(2, 1, lastRow - 1, Math.min(Math.max(sh.getLastColumn(), 2), 2)).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || "").trim() !== target) continue;
    return String(values[i][1] || "").trim() || defaultValue;
  }
  return defaultValue;
}

function baseLibGetPrequoteSettings_(ss) {
  return {
    expose_materials_to_prequote: baseLibYnToBool_(baseLibGetSettingFromSs_(ss, "expose_materials_to_prequote", "Y"), true),
    expose_templates_to_prequote: baseLibYnToBool_(baseLibGetSettingFromSs_(ss, "expose_templates_to_prequote", "Y"), true),
    default_material_limit: Math.min(Math.max(Number(baseLibGetSettingFromSs_(ss, "prequote_default_material_limit", 12) || 12), 1), 100),
    default_template_limit: Math.min(Math.max(Number(baseLibGetSettingFromSs_(ss, "prequote_default_template_limit", 12) || 12), 1), 100),
    prequote_allow_snapshot_api: baseLibYnToBool_(baseLibGetSettingFromSs_(ss, "prequote_allow_snapshot_api", "Y"), true),
    prequote_snapshot_cache_ttl_sec: Math.max(Number(baseLibGetSettingFromSs_(ss, "prequote_snapshot_cache_ttl_sec", 600) || 600), 60),
    prequote_material_search_mode: String(baseLibGetSettingFromSs_(ss, "prequote_material_search_mode", "INDEX") || "INDEX").trim().toUpperCase() || "INDEX",
    perf_metrics_enabled: baseLibYnToBool_(baseLibGetSettingFromSs_(ss, "perf_metrics_enabled", "Y"), true)
  };
}

function baseLibAssertSnapshotApiEnabled_(ss) {
  if (!baseLibGetPrequoteSettings_(ss).prequote_allow_snapshot_api) throw new Error("Snapshot API is disabled by Settings.prequote_allow_snapshot_api");
}

function baseLibGetSharedMasterVersionInfoFromSs_(ss) {
  var settings = baseLibGetPrequoteSettings_(ss);
  var materialStamp = baseLibSheetStamp_(ss, "Materials", ["updated_at", "created_at", "material_id"]);
  var tagStamp = baseLibSheetStamp_(ss, "MaterialTags", ["updated_at", "material_id"]);
  var groupStamp = baseLibSheetStamp_(ss, "MaterialGroups", ["updated_at", "group_key"]);
  var templateStamp = baseLibSheetStamp_(ss, "TemplateCatalog", ["updated_at", "template_id"]);
  var templateVersionStamp = baseLibSheetStamp_(ss, "TemplateVersions", ["created_at", "template_id"]);
  var materialVersion = baseLibSha256Hex_(materialStamp).slice(0, 16);
  var materialTagVersion = baseLibSha256Hex_(tagStamp).slice(0, 16);
  var materialGroupVersion = baseLibSha256Hex_(groupStamp).slice(0, 16);
  var templateVersion = baseLibSha256Hex_([templateStamp, templateVersionStamp].join("|")).slice(0, 16);
  var shared = baseLibSha256Hex_([String(ss.getId() || ""), materialVersion, materialTagVersion, materialGroupVersion, templateVersion, settings.expose_materials_to_prequote ? "1" : "0", settings.expose_templates_to_prequote ? "1" : "0", settings.default_material_limit, settings.default_template_limit].join("|")).slice(0, 24);
  return { schema_version: "baselib-standalone-v1", material_version: materialVersion, material_tag_version: materialTagVersion, material_group_version: materialGroupVersion, template_version: templateVersion, shared_master_version: shared, generated_at: baseLibNowIso_(), prequote_settings: settings };
}

function baseLibSheetStamp_(ss, sheetName, columns) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return sheetName + "|missing";
  var cols = baseLibGetColMap_(ss, sheetName);
  var lastRow = sh.getLastRow();
  var parts = [sheetName, String(lastRow), String(sh.getLastColumn())];
  if (lastRow >= 2) {
    for (var i = 0; i < columns.length; i++) {
      var idx = cols[columns[i]];
      if (idx === undefined) continue;
      var values = sh.getRange(2, idx + 1, lastRow - 1, 1).getDisplayValues();
      var maxValue = "";
      for (var r = 0; r < values.length; r++) {
        var value = String(values[r][0] || "").trim();
        if (value > maxValue) maxValue = value;
      }
      parts.push(columns[i] + "=" + maxValue);
    }
  }
  return parts.join("|");
}

function baseLibBuildImageUrlFromSs_(ss, fileId, baseUrlHint) {
  var id = String(fileId || "").trim();
  if (!id) return "";
  var base = baseLibNormalizeBaseUrl_(baseUrlHint);
  if (!base) base = baseLibNormalizeBaseUrl_(baseLibGetSettingFromSs_(ss, "base_url", ""));
  if (!base && typeof getAppUrl_ === "function") {
    try { base = baseLibNormalizeBaseUrl_(getAppUrl_()); } catch (e) { base = ""; }
  }
  if (!base) return "";
  return base + "?page=img&fileId=" + encodeURIComponent(id);
}

function baseLibLogSnapshotState_(ss, snapshotKey, sharedVersion, rowCount, durationMs, status, note) {
  try {
    baseLibEnsureOperationalSheets_(ss);
    var rowNo = baseLibFindRowNosByColumnValue_(ss, "SnapshotState", "snapshot_key", snapshotKey)[0] || 0;
    var payload = { snapshot_key: String(snapshotKey || "").trim(), shared_master_version: String(sharedVersion || "").trim(), built_at: baseLibNowIso_(), row_count: Math.max(Number(rowCount || 0), 0), duration_ms: Math.max(Number(durationMs || 0), 0), status: String(status || "").trim(), note: String(note || "").trim() };
    if (rowNo >= 2) baseLibUpdateRowFields_(baseLibGetSheet_(ss, "SnapshotState"), baseLibSheetMeta_(ss, "SnapshotState"), rowNo, payload);
    else baseLibAppendRow_(ss, "SnapshotState", payload);
  } catch (e) {}
}

function baseLibLogPerfMetricSafe_(ss, metricName, durationMs, entityId, context, success, note, actorType, actorId) {
  try {
    if (!baseLibGetPrequoteSettings_(ss).perf_metrics_enabled) return;
    baseLibEnsureOperationalSheets_(ss);
    baseLibAppendRow_(ss, "PerfMetrics", { metric_id: baseLibUuid_(), measured_at: baseLibNowIso_(), metric_name: String(metricName || "").trim(), duration_ms: Math.max(Number(durationMs || 0), 0), entity_id: String(entityId || "").trim(), actor_type: String(actorType || "SYSTEM").trim(), actor_id: String(actorId || "").trim(), context_json: JSON.stringify(context || {}), success_yn: success === false ? "N" : "Y", note: String(note || "").trim() });
  } catch (e) {}
}

function baseLibEnsureOperationalSheets_(ss) {
  baseLibEnsureSheetHeaders_(ss, "SnapshotState", BASELIB_SNAPSHOT_STATE_HEADERS_);
  baseLibEnsureSheetHeaders_(ss, "PerfMetrics", BASELIB_PERF_METRICS_HEADERS_);
}

function baseLibNormalizeMaterialRow_(row, tagInfo, groupKeys, ss) {
  var tagsSummary = baseLibParseTagSummaryMap_(row.tags_summary);
  var tagMap = baseLibMergeTagMaps_(tagInfo && tagInfo.tags_by_type ? tagInfo.tags_by_type : {}, tagsSummary);
  var tradeCodes = [];
  baseLibPushUnique_(tradeCodes, String(row.trade_code || "").trim());
  if (tagMap.trade) tradeCodes = baseLibUniqueStrings_(tradeCodes.concat(tagMap.trade));
  return {
    material_id: String(row.material_id || "").trim(),
    name: String(row.name || "").trim(),
    brand: String(row.brand || "").trim(),
    spec: String(row.spec || "").trim(),
    unit: String(row.unit || "").trim(),
    unit_price: baseLibToNumber_(row.unit_price, 0),
    image_file_id: String(row.image_file_id || "").trim(),
    image_file_name: String(row.image_file_name || "").trim(),
    image_url: baseLibBuildImageUrlFromSs_(ss, row.image_file_id, ""),
    note: String(row.note || "").trim(),
    created_at: String(row.created_at || "").trim(),
    updated_at: String(row.updated_at || "").trim(),
    search_text: String(row.search_text || "").trim(),
    is_active: baseLibYnToBool_(row.is_active, true),
    is_representative: baseLibYnToBool_(row.is_representative, false),
    material_type: String(row.material_type || "").trim(),
    trade_code: String(row.trade_code || "").trim(),
    space_type: String(row.space_type || "").trim(),
    expose_to_prequote: baseLibYnToBool_(row.expose_to_prequote, true),
    recommendation_score_base: baseLibToNumber_(row.recommendation_score_base, 0),
    price_band: String(row.price_band || "").trim(),
    tags_summary: String(row.tags_summary || "").trim(),
    sort_order: baseLibToNumber_(row.sort_order, 0),
    recommendation_note: String(row.recommendation_note || "").trim(),
    tags_by_type: tagMap,
    tag_records: tagInfo && Array.isArray(tagInfo.tag_records) ? baseLibClone_(tagInfo.tag_records) : [],
    legacy_group_keys: Array.isArray(groupKeys) ? groupKeys.slice() : [],
    effective_trade_codes: tradeCodes
  };
}

function baseLibFinalizeMaterialForCtx_(ctx, item, filters) {
  var c = baseLibRequireCtx_(ctx);
  var out = baseLibClone_(item || {});
  out.image_url = BaseLib.buildImageUrl(c, out.image_file_id) || out.image_url || "";
  out.title = out.name || out.material_id || "";
  out.subtitle = [out.brand, out.spec].filter(Boolean).join(" / ");
  out.source_type = String(out.source_type || "QUOTE_DB").trim().toUpperCase();
  out.score = baseLibScoreMaterial_(out, filters || {});
  return out;
}

function baseLibNormalizeTemplateRow_(row) {
  return {
    template_id: String(row.template_id || "").trim(),
    category: String(row.category || "").trim(),
    template_name: String(row.template_name || "").trim(),
    latest_version: baseLibToNumber_(row.latest_version, 0),
    created_at: String(row.created_at || "").trim(),
    updated_at: String(row.updated_at || "").trim(),
    is_active: baseLibYnToBool_(row.is_active, true),
    note: String(row.note || "").trim(),
    latest_item_count: baseLibToNumber_(row.latest_item_count, 0),
    template_type: String(row.template_type || "").trim(),
    housing_type: String(row.housing_type || "").trim(),
    area_band: String(row.area_band || "").trim(),
    style_tags_summary: String(row.style_tags_summary || "").trim(),
    trade_scope_summary: String(row.trade_scope_summary || "").trim(),
    expose_to_prequote: baseLibYnToBool_(row.expose_to_prequote, true),
    sort_order: baseLibToNumber_(row.sort_order, 0),
    tone_tags_summary: String(row.tone_tags_summary || "").trim(),
    prequote_priority: baseLibToNumber_(row.prequote_priority, 0),
    recommendation_note: String(row.recommendation_note || "").trim(),
    target_customer_summary: String(row.target_customer_summary || "").trim(),
    budget_band: String(row.budget_band || "").trim(),
    recommended_for_summary: String(row.recommended_for_summary || "").trim(),
    is_featured_prequote: baseLibYnToBool_(row.is_featured_prequote, false),
    summary_json: "",
    metadata_snapshot_json: "",
    _search_text: ""
  };
}

function baseLibFinalizeTemplateForCtx_(ctx, item) {
  var out = baseLibClone_(item || {});
  out.source_type = "QUOTE_DB";
  return out;
}

function baseLibNormalizeMaterialFilters_(filters, options, settings) {
  var raw = baseLibMergeObjects_(filters || {}, options || {});
  var cfg = settings || {};
  return { query: String(raw.query || raw.q || "").trim(), trade_codes: baseLibUniqueStrings_(baseLibNormalizeArray_(raw.trade_codes || raw.trade_code)), material_types: baseLibUniqueStrings_(baseLibNormalizeArray_(raw.material_types || raw.material_type)), space_types: baseLibUniqueStrings_(baseLibNormalizeArray_(raw.space_types || raw.space_type)), price_bands: baseLibUniqueStrings_(baseLibNormalizeArray_(raw.price_bands || raw.price_band)), tags: baseLibUniqueStrings_(baseLibNormalizeArray_(raw.tags || raw.tag_values)), include_inactive: baseLibYnToBool_(raw.include_inactive, false), include_unexposed: baseLibYnToBool_(raw.include_unexposed, false), representative_only: baseLibYnToBool_(raw.representative_only, false), limit: Math.min(Math.max(Number(raw.limit || cfg.default_material_limit || 12), 1), 50000) };
}

function baseLibNormalizeTemplateFilters_(filters, options, settings) {
  var raw = baseLibMergeObjects_(filters || {}, options || {});
  var cfg = settings || {};
  return { query: String(raw.query || raw.q || "").trim(), category: String(raw.category || "").trim(), housing_type: String(raw.housing_type || "").trim(), area_band: String(raw.area_band || "").trim(), budget_band: String(raw.budget_band || "").trim(), template_types: baseLibUniqueStrings_(baseLibNormalizeArray_(raw.template_types || raw.template_type)), style_tags: baseLibSplitSummaryTags_(raw.style_tags || raw.style_tags_summary), tone_tags: baseLibSplitSummaryTags_(raw.tone_tags || raw.tone_tags_summary), trade_scope_tags: baseLibSplitSummaryTags_(raw.trade_scope_tags || raw.trade_scope || raw.trade_scope_summary), include_inactive: baseLibYnToBool_(raw.include_inactive, false), include_unexposed: baseLibYnToBool_(raw.include_unexposed, false), limit: Math.min(Math.max(Number(raw.limit || cfg.default_template_limit || 12), 1), 50000) };
}

function baseLibMaterialMatchesFilters_(item, filters, settings) {
  var row = item || {};
  var f = filters || {};
  var cfg = settings || {};
  if (!row.material_id) return false;
  if (!f.include_inactive && !row.is_active) return false;
  if (!f.include_unexposed && (!cfg.expose_materials_to_prequote || !row.expose_to_prequote)) return false;
  if (f.representative_only && !row.is_representative) return false;
  if (f.trade_codes.length && !baseLibHasAnyIntersection_(row.effective_trade_codes, f.trade_codes)) return false;
  if (f.material_types.length && f.material_types.indexOf(String(row.material_type || "").trim()) < 0) return false;
  if (f.space_types.length && f.space_types.indexOf(String(row.space_type || "").trim()) < 0) return false;
  if (f.price_bands.length && f.price_bands.indexOf(String(row.price_band || "").trim()) < 0) return false;
  if (f.tags.length) {
    var summaryTags = [];
    for (var key in row.tags_by_type || {}) if (Object.prototype.hasOwnProperty.call(row.tags_by_type, key)) summaryTags = summaryTags.concat(row.tags_by_type[key]);
    if (!baseLibHasAnyIntersection_(summaryTags, f.tags)) return false;
  }
  if (f.query) {
    var tokens = baseLibPickSearchTokens_(baseLibTokenize_(f.query), 8);
    var hay = baseLibNormText_(baseLibMaterialSearchTextFromObject_(row));
    for (var i = 0; i < tokens.length; i++) if (hay.indexOf(tokens[i]) < 0) return false;
  }
  return true;
}

function baseLibTemplateMatchesFilters_(item, filters, settings) {
  var row = item || {};
  var f = filters || {};
  var cfg = settings || {};
  if (!row.template_id) return false;
  if (!f.include_inactive && !row.is_active) return false;
  if (!f.include_unexposed && (!cfg.expose_templates_to_prequote || !row.expose_to_prequote)) return false;
  if (f.category && row.category !== f.category) return false;
  if (f.housing_type && row.housing_type && row.housing_type !== "ALL" && row.housing_type !== f.housing_type) return false;
  if (f.area_band && row.area_band && row.area_band !== "ALL" && row.area_band !== f.area_band) return false;
  if (f.budget_band && row.budget_band && row.budget_band !== "ALL" && row.budget_band !== f.budget_band) return false;
  if (f.template_types.length && f.template_types.indexOf(String(row.template_type || "").trim()) < 0) return false;
  if (f.style_tags.length && !baseLibHasAnyIntersection_(baseLibSplitSummaryTags_(row.style_tags_summary), f.style_tags)) return false;
  if (f.tone_tags.length && !baseLibHasAnyIntersection_(baseLibSplitSummaryTags_(row.tone_tags_summary), f.tone_tags)) return false;
  if (f.trade_scope_tags.length && !baseLibHasAnyIntersection_(baseLibSplitSummaryTags_(row.trade_scope_summary), f.trade_scope_tags)) return false;
  if (f.query) {
    var tokens = baseLibPickSearchTokens_(baseLibTokenize_(f.query), 8);
    var hay = String(row._search_text || "");
    for (var i = 0; i < tokens.length; i++) if (hay.indexOf(tokens[i]) < 0) return false;
  }
  return true;
}

function baseLibScoreMaterial_(item, filters) {
  var score = baseLibToNumber_((item || {}).recommendation_score_base, 0);
  if (item && item.is_representative) score += 100;
  if (!filters || !filters.query) return score;
  var tokens = baseLibPickSearchTokens_(baseLibTokenize_(filters.query), 8);
  var hay = baseLibNormText_(baseLibMaterialSearchTextFromObject_(item));
  for (var i = 0; i < tokens.length; i++) if (hay.indexOf(tokens[i]) >= 0) score += Math.max(10 - i, 1);
  return score;
}

function baseLibBuildTemplateSearchText_(template, summary) {
  return baseLibNormText_([template.template_name, template.category, template.template_type, template.housing_type, template.area_band, template.budget_band, template.style_tags_summary, template.tone_tags_summary, template.trade_scope_summary, template.recommendation_note, template.target_customer_summary, template.recommended_for_summary, baseLibExtractTextFromValue_(summary)].join(" "));
}

function baseLibMaterialSearchTextFromObject_(row) {
  var src = row || {};
  return [src.search_text, src.name, src.brand, src.spec, src.note, src.material_type, src.trade_code, src.space_type, src.price_band, src.tags_summary, src.recommendation_note].filter(Boolean).join(" ");
}

function baseLibNormalizeBaseUrl_(value) { return String(value || "").trim().replace(/\/+$/, ""); }
function baseLibNowIso_() { return new Date().toISOString(); }
function baseLibSafeJsonParse_(raw, fallback) { try { return raw ? JSON.parse(raw) : fallback; } catch (e) { return fallback; } }
function baseLibClone_(value) { return JSON.parse(JSON.stringify(value === undefined ? null : value)); }
function baseLibMergeObjects_(base, extra) { var out = {}; var sources = [base || {}, extra || {}]; for (var s = 0; s < sources.length; s++) for (var key in sources[s]) if (Object.prototype.hasOwnProperty.call(sources[s], key)) out[key] = sources[s][key]; return out; }
function baseLibYnToBool_(value, defaultValue) { var raw = String(value === undefined || value === null ? "" : value).trim().toUpperCase(); if (!raw) return !!defaultValue; if (raw === "Y" || raw === "YES" || raw === "TRUE" || raw === "1") return true; if (raw === "N" || raw === "NO" || raw === "FALSE" || raw === "0") return false; return !!defaultValue; }
function baseLibToNumber_(value, defaultValue) { var n = Number(value); return isFinite(n) ? n : (defaultValue == null ? 0 : defaultValue); }
function baseLibNormalizeArray_(value) { if (Array.isArray(value)) return value; var raw = String(value || "").trim(); return raw ? raw.split(/[,\s]+/) : []; }
function baseLibUniqueStrings_(arr) { var list = Array.isArray(arr) ? arr : []; var seen = Object.create(null); var out = []; for (var i = 0; i < list.length; i++) { var value = String(list[i] || "").trim(); if (!value || seen[value]) continue; seen[value] = 1; out.push(value); } return out; }
function baseLibPushUnique_(arr, value) { var raw = String(value || "").trim(); if (!raw || arr.indexOf(raw) >= 0) return arr; arr.push(raw); return arr; }
function baseLibSplitSummaryTags_(raw) { return baseLibUniqueStrings_(String(raw || "").split(/[,\s]+/)); }

function baseLibParseTagSummaryMap_(summary) {
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
    baseLibPushUnique_(out[type], value);
  });
  return out;
}

function baseLibMergeTagMaps_(primary, secondary) {
  var out = Object.create(null);
  var sources = [primary || {}, secondary || {}];
  for (var s = 0; s < sources.length; s++) {
    for (var key in sources[s]) {
      if (!Object.prototype.hasOwnProperty.call(sources[s], key)) continue;
      if (!out[key]) out[key] = [];
      var values = Array.isArray(sources[s][key]) ? sources[s][key] : [sources[s][key]];
      for (var i = 0; i < values.length; i++) baseLibPushUnique_(out[key], values[i]);
    }
  }
  return out;
}

function baseLibHasAnyIntersection_(left, right) { var a = baseLibUniqueStrings_(left || []); var b = baseLibUniqueStrings_(right || []); for (var i = 0; i < a.length; i++) if (b.indexOf(a[i]) >= 0) return true; return false; }
function baseLibNormText_(value) { return String(value || "").toLowerCase().replace(/[^\p{L}\p{N}]+/gu, " ").replace(/\s+/g, " ").trim(); }
function baseLibTokenize_(value) { var text = baseLibNormText_(value); return text ? baseLibUniqueStrings_(text.split(" ")) : []; }
function baseLibPickSearchTokens_(tokens, maxTokens) { var list = baseLibUniqueStrings_(tokens || []); list.sort(function(a, b) { if (b.length !== a.length) return b.length - a.length; return a.localeCompare(b); }); return list.slice(0, Math.max(Number(maxTokens || 8), 1)); }
function baseLibParseRefs_(refsStr) { var raw = String(refsStr || "").trim(); if (!raw) return []; var parts = raw.split(","); var out = []; for (var i = 0; i < parts.length; i++) { var chunk = String(parts[i] || "").trim(); if (!chunk) continue; var pair = chunk.split(":"); var id = String(pair[0] || "").trim(); var rowNo = Number(pair[1] || 0); if (!id || rowNo < 2) continue; out.push({ id: id, rowNo: rowNo }); } return out; }
function baseLibRefsToString_(refs) { var list = Array.isArray(refs) ? refs : []; var out = []; for (var i = 0; i < list.length; i++) out.push(String(list[i].id || "") + ":" + String(list[i].rowNo || 0)); return out.join(","); }
function baseLibUniqueSortedNumbers_(numbers) { var seen = Object.create(null); var out = []; var list = Array.isArray(numbers) ? numbers : []; for (var i = 0; i < list.length; i++) { var value = Number(list[i] || 0); if (value < 2 || seen[value]) continue; seen[value] = 1; out.push(value); } out.sort(function(a, b) { return a - b; }); return out; }
function baseLibExtractTextFromValue_(value) { if (value === null || value === undefined) return ""; if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return String(value); if (Array.isArray(value)) return value.map(baseLibExtractTextFromValue_).join(" "); if (typeof value === "object") { var parts = []; for (var key in value) if (Object.prototype.hasOwnProperty.call(value, key)) parts.push(baseLibExtractTextFromValue_(value[key])); return parts.join(" "); } return ""; }

function baseLibGetCachedJson_(key) { var cacheKey = String(key || "").trim(); if (!cacheKey) return null; try { var raw = CacheService.getScriptCache().get(cacheKey); return raw ? JSON.parse(raw) : null; } catch (e) { return null; } }
function baseLibPutCachedJson_(key, value, ttlSec) { var cacheKey = String(key || "").trim(); if (!cacheKey) return; try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(value), Math.max(Number(ttlSec || 60), 30)); } catch (e) {} }

function baseLibGetCachedJsonChunked_(key) {
  var cacheKey = String(key || "").trim();
  if (!cacheKey) return null;
  try {
    var cache = CacheService.getScriptCache();
    var sizeRaw = cache.get(cacheKey + "_count");
    if (!sizeRaw) return null;
    var count = Number(sizeRaw || 0);
    if (count < 1) return null;
    var parts = [];
    for (var i = 0; i < count; i++) {
      var part = cache.get(cacheKey + "_part_" + i);
      if (part === null || part === undefined) return null;
      parts.push(part);
    }
    return JSON.parse(parts.join(""));
  } catch (e) {
    return null;
  }
}

function baseLibPutCachedJsonChunked_(key, value, ttlSec) {
  var cacheKey = String(key || "").trim();
  if (!cacheKey) return;
  try {
    var raw = JSON.stringify(value);
    var chunkSize = 80000;
    var parts = [];
    for (var i = 0; i < raw.length; i += chunkSize) parts.push(raw.slice(i, i + chunkSize));
    var cache = CacheService.getScriptCache();
    cache.put(cacheKey + "_count", String(parts.length), Math.max(Number(ttlSec || 60), 30));
    for (var p = 0; p < parts.length; p++) cache.put(cacheKey + "_part_" + p, parts[p], Math.max(Number(ttlSec || 60), 30));
  } catch (e) {}
}

function baseLibGetOrBuildCachedJsonChunkedWithLock_(cacheKey, ttlSec, builderFn, options) {
  var cached = baseLibGetCachedJsonChunked_(cacheKey);
  if (cached !== null) return cached;
  var opts = options || {};
  var lock = LockService.getScriptLock();
  var gotLock = false;
  try { gotLock = lock.tryLock(Math.max(Number(opts.wait_ms || 1800), 0)); } catch (e0) { gotLock = false; }
  if (gotLock) {
    try {
      var cachedAgain = baseLibGetCachedJsonChunked_(cacheKey);
      if (cachedAgain !== null) return cachedAgain;
      var built = typeof builderFn === "function" ? builderFn() : null;
      if (built !== undefined) baseLibPutCachedJsonChunked_(cacheKey, built, ttlSec);
      return built;
    } finally {
      lock.releaseLock();
    }
  }
  Utilities.sleep(Math.max(Number(opts.retry_ms || 120), 0));
  var retry = baseLibGetCachedJsonChunked_(cacheKey);
  if (retry !== null) return retry;
  if (Object.prototype.hasOwnProperty.call(opts, "fallback_value")) return opts.fallback_value;
  return typeof builderFn === "function" ? builderFn() : null;
}

function baseLibSha256Hex_(value) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value || ""), Utilities.Charset.UTF_8);
  var out = [];
  for (var i = 0; i < bytes.length; i++) {
    var v = (bytes[i] + 256) % 256;
    out.push((v < 16 ? "0" : "") + v.toString(16));
  }
  return out.join("");
}

function baseLibUuid_() { return Utilities.getUuid(); }
