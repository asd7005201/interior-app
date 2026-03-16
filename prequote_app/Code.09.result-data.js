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
