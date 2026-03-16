// ============================================================
//  INITIALIZATION
//  에디터에서 최초 1회 실행: initializeApp()
// ============================================================

/**
 * 최초 배포 전 에디터에서 1회 실행해주세요.
 * 컨테이너 바운드(시트에 붙은) 스크립트라면 자동 감지합니다.
 * 독립 스크립트라면 initializeAppManual('스프레드시트ID') 를 사용하세요.
 */
function initializeApp() {
  var id = setupSpreadsheetId();
  Logger.log("SPREADSHEET_ID set: " + id);
  Logger.log("Settings loaded: " + JSON.stringify(getSettings_()));
  return { spreadsheet_id: id, status: "OK" };
}

function initializeAppManual(spreadsheetId) {
  var id = setupSpreadsheetIdFromManual(spreadsheetId);
  Logger.log("SPREADSHEET_ID set (manual): " + id);
  return { spreadsheet_id: id, status: "OK" };
}

/**
 * 모든 진입점에서 호출. Script Properties에 ID가 없으면 자동 감지 시도.
 * getSpreadsheetId_() 자체에 fallback이 있으므로 여기서는 한 번 호출만 해줌.
 */
function ensureSpreadsheetId_() {
  try {
    getSpreadsheetId_();
  } catch (e) {
    // 컨테이너 바운드라면 자동 설정 시도
    try {
      setupSpreadsheetId();
    } catch (e2) {
      throw new Error(
        "스프레드시트 ID를 찾을 수 없습니다.\n" +
        "에디터에서 initializeApp() 을 먼저 실행해주세요.\n" +
        "원본 오류: " + e.message
      );
    }
  }
}
