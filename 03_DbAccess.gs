/***************************************
 * S1 Sales Portal - 03_DbAccess.gs
 * 분리일: 2026-06-19
 * 원칙: 기능 변경 없이 최신 단일 파일을 물리적으로 분리
 ***************************************/

function getMasterSpreadsheet_() {
  return SpreadsheetApp.openById(PORTAL_CONFIG.MASTER_SPREADSHEET_ID);
}

function getWebAppDbSpreadsheet_() {
  const configuredId = String(PORTAL_CONFIG.WEBAPP_DB_SPREADSHEET_ID || '').trim();
  if (configuredId) return SpreadsheetApp.openById(configuredId);

  // 컨테이너 바운드 웹앱이면 현재 연결된 스프레드시트를 웹앱_DB로 사용합니다.
  // standalone 프로젝트이거나 active spreadsheet가 없으면 마스터 파일로 fallback합니다.
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (err) {}

  return getMasterSpreadsheet_();
}

function getPortalDbInfo() {
  const masterSs = getMasterSpreadsheet_();
  const dbSs = getWebAppDbSpreadsheet_();
  return {
    ok: true,
    masterSpreadsheetId: masterSs.getId(),
    masterSpreadsheetName: masterSs.getName(),
    webAppDbSpreadsheetId: dbSs.getId(),
    webAppDbSpreadsheetName: dbSs.getName(),
    separated: masterSs.getId() !== dbSs.getId(),
    webAppSheets: [
      PORTAL_CONFIG.CUSTOMER_INDEX_SHEET_NAME,
      PORTAL_CONFIG.NOTICE_SHEET_NAME,
      PORTAL_CONFIG.CONTACT_HISTORY_SHEET_NAME,
      PORTAL_CONFIG.TODAY_SHEET_NAME,
      PORTAL_CONFIG.CONTACT_PROFILE_SHEET_NAME
    ],
    masterSheets: [
      PORTAL_CONFIG.MASTER_SHEET_NAME,
      PORTAL_CONFIG.SUPPORT_SHEET_NAME
    ],
    supportRequestLocation: 'MASTER_SPREADSHEET_ID'
  };
}

