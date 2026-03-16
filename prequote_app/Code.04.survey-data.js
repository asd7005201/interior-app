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
