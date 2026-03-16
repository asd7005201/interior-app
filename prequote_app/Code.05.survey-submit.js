// ============================================================
//  SURVEY SUBMISSION + ESTIMATE CALCULATION
// ============================================================

/**
 * Submit completed survey and calculate pre-quote estimate.
 * @param {Object} payload - { answers: {code: value}, contact: {name, phone, email, method, note}, files: [] }
 * @returns {Object} { request_id, share_token, estimate, recommendations, flags }
 */
function submitSurvey(payload) {
  ensureSpreadsheetId_();
  ensureSurveySubmitStorageSchema_();
  var startMs = Date.now();
  var error = null;
  var requestId = "";
  try {
    if (!payload || typeof payload !== "object") throw new Error("Invalid payload");
    var answers = payload.answers || {};
    var contact = payload.contact || {};
    var files = Array.isArray(payload.files) ? payload.files : [];
    var settings = getSettings_();

    var name = String(contact.name || "").trim();
    var phone = String(contact.phone || "").trim();
    if (!name) throw new Error("이름을 입력해주세요.");
    if (!phone) throw new Error("연락처를 입력해주세요.");

    requestId = "PQ_" + uuid_().replace(/-/g, "").slice(0, 16).toUpperCase();
    var shareToken = token_().slice(0, 32);
    var now = nowIso_();
    var projectType = String(answers.Q000_PROJECT_TYPE || "").toUpperCase();
    var flowMode = String(answers.Q000_FLOW_MODE || "").toUpperCase();
    var tags = calculateTags_(answers);
    var estimate = calculateEstimate_(answers, projectType);
    var recommendations = [];
    var asyncPostprocess = ynToBool_(settings.survey_submit_async_postprocess, true);
    var postprocessOptions = buildSurveySubmitPostprocessOptions_(settings);
    var duplicatePhoneKey = normalizePhoneKey_(phone);
    var duplicateEmailKey = normalizeEmailKey_(contact.email);
    var requestObj = {
      request_id: requestId,
      created_at: now,
      updated_at: now,
      source_channel: "WEB_SURVEY",
      customer_name: name,
      contact_phone: phone,
      contact_email: String(contact.email || "").trim(),
      preferred_contact_method: String(contact.method || settings.default_contact_method || "KAKAO").toUpperCase(),
      project_type: projectType,
      housing_type: answers.R001_HOUSING_TYPE || answers.C001_BIZ_TYPE || "",
      property_status: "",
      address_text: String(contact.address || answers.Q906_ADDRESS || "").trim(),
      area_py: answers.R002_AREA || answers.C002_AREA || "",
      area_m2: "",
      scope_level: answers.R011_SCOPE_LEVEL || answers.C011_SCOPE_LEVEL || "",
      flow_mode: flowMode,
      preset_bundle_id: answers.Q000_PRESET_RESI || answers.Q000_PRESET_COMM || "",
      tags_json: JSON.stringify(tags),
      estimate_min: estimate.min || 0,
      estimate_max: estimate.max || 0,
      estimate_json: JSON.stringify(estimate),
      risk_flags_json: JSON.stringify(estimate.flags || []),
      recommendation_count: recommendations.length,
      final_estimate_min: estimate.min || 0,
      final_estimate_max: estimate.max || 0,
      final_estimate_source: "AUTO",
      final_recommendation_count: recommendations.length,
      status: "NEW",
      priority: "NORMAL",
      work_state: "NEW",
      survey_version: settings.survey_version || "SURVEY_V6",
      share_token: shareToken,
      customer_note: String(contact.note || "").trim(),
      submitted_at: now,
      schedule_start: answers.R320_START || answers.C320_START || "",
      schedule_end: answers.R321_END || answers.C321_END || "",
      schedule_visit: answers.R322_VISIT || answers.C322_VISIT || "",
      duplicate_phone_key: duplicatePhoneKey,
      duplicate_email_key: duplicateEmailKey,
      duplicate_customer_count: asyncPostprocess ? "" : countDuplicateRequests_(duplicatePhoneKey, duplicateEmailKey, name, requestId),
      review_status_label: REVIEW_STATUS_LABEL_MAP_.AUTO,
      quote_draft_status: "PENDING",
      answers_snapshot_json: JSON.stringify(answers)
    };
    appendRow_("Requests", requestObj);

    var questionCodes = Object.keys(answers);
    if (questionCodes.length > 0) {
      var answerRows = [];
      for (var i = 0; i < questionCodes.length; i++) {
        var code = questionCodes[i];
        var val = answers[code];
        answerRows.push({
          request_id: requestId,
          answer_id: requestId + "_A" + String(i + 1).padStart(3, "0"),
          question_id: code,
          question_code: code,
          answer_type: Array.isArray(val) ? "MULTI" : "SINGLE",
          answer_value_text: Array.isArray(val) ? val.join(",") : String(val || ""),
          answer_value_number: Array.isArray(val) ? "" : (isFinite(Number(val)) ? Number(val) : ""),
          answer_value_json: JSON.stringify(val === undefined ? "" : val),
          normalized_answer: Array.isArray(val) ? val.join(",") : String(val || ""),
          tags_applied_json: "",
          score_delta: "",
          created_at: now
        });
      }
      appendRows_("RequestAnswers", answerRows);
    }

    if (files.length > 0) {
      var fileLimit = Math.max(Number(settings.file_upload_max_count || 10), 1);
      var fileRows = [];
      for (var f = 0; f < Math.min(files.length, fileLimit); f++) {
        fileRows.push(buildRequestFileRow_(requestId, files[f], f, now));
      }
      appendRows_("RequestFiles", fileRows);
    }

    appendRow_("RequestTimeline", {
      event_id: uuid_(),
      request_id: requestId,
      event_at: now,
      actor_type: "CUSTOMER",
      actor_id: name,
      event_type: "SUBMITTED",
      from_status: "",
      to_status: "NEW",
      message: "설문 접수 완료",
      payload_json: "",
      event_label: "설문 접수",
      dedup_key: "request:" + requestId + ":submitted"
    });

    if (asyncPostprocess) {
      try {
        enqueueSurveySubmitPostprocess_(requestObj, estimate, settings);
      } catch (queueError) {
        Logger.log("Survey submit postprocess queue fallback: " + queueError);
        runSurveySubmitPostprocess_(requestId, estimate, postprocessOptions);
      }
    } else {
      runSurveySubmitPostprocess_(requestId, estimate, postprocessOptions);
    }

    cacheResultBootstrap_(requestObj, estimate);

    return {
      request_id: requestId,
      share_token: shareToken,
      estimate: estimate,
      recommendations: recommendations.slice(0, 12),
      flags: estimate.flags || []
    };
  } catch (e) {
    error = e;
    throw e;
  } finally {
    logPerfMetricSafe_(
      "submit_survey",
      Date.now() - startMs,
      requestId,
      {},
      !error,
      error ? String(error) : "",
      "CUSTOMER",
      requestId
    );
  }
}
