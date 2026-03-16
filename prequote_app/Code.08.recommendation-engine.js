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
