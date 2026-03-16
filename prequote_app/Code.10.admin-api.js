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
