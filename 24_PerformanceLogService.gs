/***************************************
 * S1 Sales Portal - 24_PerformanceLogService.gs
 * P460: 성능/저장 안정성 로그를 웹앱 DB 스프레드시트에 남기는 경량 진단 계층
 ***************************************/

function getPortalPerfLogSheetNameP460_() {
  return (PORTAL_CONFIG && PORTAL_CONFIG.PERF_LOG_SHEET_NAME) || '성능로그_DB';
}

function getPortalPerfLogHeadersP460_() {
  return (PORTAL_CONFIG && PORTAL_CONFIG.PERF_LOG_HEADERS) || ['기록일시','사용자','세션ID','이벤트','구간','소요ms','rowNo','고객번호','화면','상태','오류','상세JSON'];
}

function ensurePortalPerfLogSheetP460_() {
  const ss = getWebAppDbSpreadsheet_();
  const name = getPortalPerfLogSheetNameP460_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const headers = getPortalPerfLogHeadersP460_();
  const width = Math.max(headers.length, sheet.getLastColumn() || 1);
  const current = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function(v){ return String(v || '').trim(); });
  let needsHeader = false;
  for (let i = 0; i < headers.length; i++) {
    if (current[i] !== headers[i]) { needsHeader = true; break; }
  }
  if (needsHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function appendPortalPerfLogsP460(logs) {
  logs = Array.isArray(logs) ? logs : (logs ? [logs] : []);
  if (!logs.length) return { ok: true, appended: 0, sheetName: getPortalPerfLogSheetNameP460_() };
  const max = 80;
  logs = logs.slice(0, max);
  const user = (Session.getActiveUser && Session.getActiveUser().getEmail()) || '';
  const now = new Date();
  const nowText = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const rows = logs.map(function(log) {
    log = log || {};
    const detail = log.detail || {};
    return [
      nowText,
      String(log.user || user || ''),
      String(log.sessionId || ''),
      String(log.event || log.eventName || ''),
      String(log.phase || log.label || ''),
      Number(log.durationMs || log.ms || 0) || 0,
      String(log.rowNo || (detail && detail.rowNo) || ''),
      String(log.customerNo || (detail && detail.customerNo) || ''),
      String(log.page || (detail && detail.page) || ''),
      String(log.status || ''),
      String(log.error || '').slice(0, 500),
      JSON.stringify(detail || {}).slice(0, 5000)
    ];
  });
  const sheet = ensurePortalPerfLogSheetP460_();
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, getPortalPerfLogHeadersP460_().length).setValues(rows);
  return {
    ok: true,
    appended: rows.length,
    sheetName: sheet.getName(),
    spreadsheetId: sheet.getParent().getId(),
    spreadsheetName: sheet.getParent().getName(),
    url: sheet.getParent().getUrl()
  };
}

function appendPortalServerPerfLogP460_(log) {
  try { return appendPortalPerfLogsP460([log]); } catch (err) { Logger.log('P460 server perf log failed: ' + (err && err.stack || err)); }
  return { ok: false };
}

function getPortalPerfLogSummaryP460(limit) {
  limit = Math.max(1, Math.min(200, Number(limit || 50)));
  const sheet = ensurePortalPerfLogSheetP460_();
  const lastRow = sheet.getLastRow();
  const headers = getPortalPerfLogHeadersP460_();
  if (lastRow < 2) {
    return { ok: true, count: 0, logs: [], sheetName: sheet.getName(), spreadsheetName: sheet.getParent().getName(), url: sheet.getParent().getUrl() };
  }
  const startRow = Math.max(2, lastRow - limit + 1);
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, headers.length).getDisplayValues();
  const logs = values.reverse().map(function(r) {
    const obj = {};
    headers.forEach(function(h, i) { obj[h] = r[i]; });
    return obj;
  });
  return { ok: true, count: lastRow - 1, logs: logs, sheetName: sheet.getName(), spreadsheetName: sheet.getParent().getName(), url: sheet.getParent().getUrl() };
}
