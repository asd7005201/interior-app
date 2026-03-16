// ============================================================
//  ESTIMATE CALCULATION ENGINE
// ============================================================

function calculateEstimate_(answers, projectType) {
  var rules = readStaticRowsCached_("EstimateRules").filter(function(r) {
    return ynToBool_(r.is_active);
  });

  // Sort by priority
  rules.sort(function(a, b) { return Number(a.priority || 0) - Number(b.priority || 0); });

  var result = { min: 0, max: 0, base_min: 0, base_max: 0, adjustments: [], flags: [] };

  // Phase 1: BASE_RANGE — find matching base
  var baseRules = rules.filter(function(r) { return r.rule_type === "BASE_RANGE"; });
  var bestBase = null;
  var bestBaseScore = -1;

  for (var i = 0; i < baseRules.length; i++) {
    var br = baseRules[i];
    var score = scoreBaseRule_(br, answers, projectType);
    if (score > bestBaseScore) {
      bestBaseScore = score;
      bestBase = br;
    }
  }

  if (bestBase) {
    var baseVal = safeJsonParse_(bestBase.adjustment_value, {});
    result.base_min = Number(baseVal.min || 0);
    result.base_max = Number(baseVal.max || 0);
    result.min = result.base_min;
    result.max = result.base_max;
    result.adjustments.push({
      rule_id: bestBase.rule_id,
      label: bestBase.display_label || "기본 범위",
      type: "BASE",
      min: result.base_min,
      max: result.base_max
    });
  }

  // Phase 2: RANGE_FACTOR — multiply
  var factorRules = rules.filter(function(r) { return r.rule_type === "RANGE_FACTOR"; });
  for (var f = 0; f < factorRules.length; f++) {
    var fr = factorRules[f];
    if (!matchCondition_(fr.condition_json, answers)) continue;
    var fv = safeJsonParse_(fr.adjustment_value, {});
    var minF = Number(fv.min_factor || 1);
    var maxF = Number(fv.max_factor || 1);
    result.min = Math.round(result.min * minF);
    result.max = Math.round(result.max * maxF);
    result.adjustments.push({
      rule_id: fr.rule_id,
      label: fr.display_label || "",
      type: "FACTOR",
      min_factor: minF,
      max_factor: maxF
    });
  }

  // Phase 3: RANGE_ADD — add flat amounts
  var addRules = rules.filter(function(r) { return r.rule_type === "RANGE_ADD"; });
  for (var a = 0; a < addRules.length; a++) {
    var ar = addRules[a];
    if (!matchCondition_(ar.condition_json, answers)) continue;
    var av = safeJsonParse_(ar.adjustment_value, {});
    var addMin = Number(av.min || 0);
    var addMax = Number(av.max || 0);
    result.min += addMin;
    result.max += addMax;
    result.adjustments.push({
      rule_id: ar.rule_id,
      label: ar.display_label || "",
      type: "ADD",
      min: addMin,
      max: addMax
    });
  }

  // Phase 4: REVIEW_FLAG — collect flags
  var flagRules = rules.filter(function(r) { return r.rule_type === "REVIEW_FLAG"; });
  for (var fl = 0; fl < flagRules.length; fl++) {
    var flr = flagRules[fl];
    if (!matchCondition_(flr.condition_json, answers)) continue;
    var flagVal = safeJsonParse_(flr.adjustment_value, {});
    result.flags.push({
      rule_id: flr.rule_id,
      flag: flagVal.flag || "",
      severity: flagVal.severity || "medium",
      label: flr.display_label || ""
    });
  }

  return result;
}

function scoreBaseRule_(rule, answers, projectType) {
  // Match housing_type
  var rHousing = String(rule.housing_type || "").toUpperCase().trim();
  var aHousing = String(answers.R001_HOUSING_TYPE || answers.C001_BIZ_TYPE || "").toUpperCase().trim();
  if (projectType === "COMMERCIAL") aHousing = String(answers.C001_BIZ_TYPE || "COMMERCIAL").toUpperCase();

  if (rHousing && aHousing && rHousing !== aHousing) return -1;

  // Match area_band
  var rArea = String(rule.area_band || "").toUpperCase().trim();
  var aArea = String(answers.R002_AREA || answers.C002_AREA || "").toUpperCase().trim();
  if (rArea && aArea && rArea !== aArea) return -1;

  // Match condition_json
  var cond = safeJsonParse_(rule.condition_json, {});
  var score = 0;
  for (var key in cond) {
    if (!Object.prototype.hasOwnProperty.call(cond, key)) continue;
    var expected = cond[key];
    var actual = answers[key];
    if (!actual) return -1;
    var actualArr = Array.isArray(actual) ? actual : [String(actual)];
    var expectedArr = Array.isArray(expected) ? expected : [String(expected)];
    var found = false;
    for (var e = 0; e < expectedArr.length; e++) {
      if (actualArr.indexOf(expectedArr[e]) >= 0) { found = true; break; }
    }
    if (!found) return -1;
    score++;
  }

  // Bonus for specificity
  if (rHousing && rHousing === aHousing) score += 2;
  if (rArea && rArea === aArea) score += 2;
  return score;
}

function matchCondition_(conditionJson, answers) {
  var cond = safeJsonParse_(conditionJson, {});
  if (!cond || typeof cond !== "object") return false;
  var keys = Object.keys(cond);
  if (keys.length === 0) return false;

  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var expected = cond[key];
    var actual = answers[key];
    if (!actual && actual !== 0) return false;

    var actualArr = Array.isArray(actual) ? actual : [String(actual)];
    var expectedArr = Array.isArray(expected) ? expected : [String(expected)];

    // ANY match (OR within same key)
    var matched = false;
    for (var e = 0; e < expectedArr.length; e++) {
      if (actualArr.indexOf(expectedArr[e]) >= 0) { matched = true; break; }
    }
    if (!matched) return false;
  }
  return true; // All keys matched (AND across keys)
}
