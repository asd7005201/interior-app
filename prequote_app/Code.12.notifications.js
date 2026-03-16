// ============================================================
//  OWNER EMAIL TARGETS (for notifications)
// ============================================================

function getOwnerEmailTargets_() {
  var s = getSettings_();
  var raw = String(s.owner_email || "").trim();
  if (!raw) return [];
  return raw.split(",").map(function(v) { return v.trim(); }).filter(Boolean);
}
