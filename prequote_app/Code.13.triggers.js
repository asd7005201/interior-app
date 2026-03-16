// ============================================================
//  TIMED TRIGGERS
// ============================================================

function setupSyncTrigger() {
  ensureSpreadsheetId_();
  assertEditorAdminExecution_();
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === "runScheduledSync_") {
      ScriptApp.deleteTrigger(existing[i]);
    }
  }
  var s = getSettings_();
  var mins = Math.max(Number(s.sync_interval_minutes || 30), 10);
  ScriptApp.newTrigger("runScheduledSync_").timeBased().everyMinutes(mins).create();
  return { interval: mins };
}

function setupOperationalNotificationsTrigger() {
  ensureSpreadsheetId_();
  assertEditorAdminExecution_();
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === "runOperationalNotifications_") {
      ScriptApp.deleteTrigger(existing[i]);
    }
  }
  var settings = getSettings_();
  var mode = String(settings.perf_dashboard_notifications_mode || "TRIGGER").trim().toUpperCase();
  if (mode === "DISABLED" || mode === "OFF" || mode === "NONE") {
    return { enabled: false, mode: mode };
  }
  var mins = Math.max(Number(settings.sync_interval_minutes || 30), 10);
  ScriptApp.newTrigger("runOperationalNotifications_").timeBased().everyMinutes(mins).create();
  return { enabled: true, mode: mode, interval: mins };
}

function runScheduledSync_() {
  ensureSpreadsheetId_();
  var s = getSettings_();
  if (String(s.sync_enabled || "").toUpperCase() !== "Y") return;
  try {
    runWithSyncLock_(function() {
      var versionInfo = null;
      try { versionInfo = getQuoteSharedMasterVersion_(); } catch (e0) { Logger.log("Sync version fetch error: " + e0); }
      try { syncMaterialsCacheInternal_(versionInfo); } catch (e) { Logger.log("Sync materials error: " + e); }
      try { syncTemplatesCacheInternal_(versionInfo); } catch (e2) { Logger.log("Sync templates error: " + e2); }
    });
  } catch (e3) {
    Logger.log("Scheduled sync lock error: " + e3);
  }
}
