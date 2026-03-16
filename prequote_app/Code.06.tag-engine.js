// ============================================================
//  TAG CALCULATION ENGINE
// ============================================================

function calculateTags_(answers) {
  var rules = readStaticRowsCached_("SurveyTagRules").filter(function(r) {
    return ynToBool_(r.is_active);
  });

  var tagMap = {}; // { "style:MODERN": { type, value, totalWeight } }

  for (var i = 0; i < rules.length; i++) {
    var rule = rules[i];
    var qCode = String(rule.question_id || "").trim();
    var optCode = String(rule.option_code || "").trim();
    var answerVal = answers[qCode];
    if (!answerVal) continue;

    var match = false;
    if (Array.isArray(answerVal)) {
      match = answerVal.indexOf(optCode) >= 0;
    } else {
      match = String(answerVal) === optCode;
    }

    if (match) {
      var tagKey = String(rule.tag_type || "") + ":" + String(rule.tag_value || "");
      if (!tagMap[tagKey]) {
        tagMap[tagKey] = { type: rule.tag_type, value: rule.tag_value, weight: 0 };
      }
      tagMap[tagKey].weight += Number(rule.weight || 1);
    }
  }

  // Convert to array sorted by weight desc
  var tags = [];
  for (var k in tagMap) {
    if (Object.prototype.hasOwnProperty.call(tagMap, k)) {
      tags.push(tagMap[k]);
    }
  }
  tags.sort(function(a, b) { return b.weight - a.weight; });
  return tags;
}
