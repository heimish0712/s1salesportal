/***************************************
 * S1 Sales Portal - 12_SetupDiagnostics.gs
 * 분리일: 2026-06-19
 * 원칙: 기능 변경 없이 최신 단일 파일을 물리적으로 분리
 ***************************************/

function setupWebAppDbSheets() {
  const dbSs = getWebAppDbSpreadsheet_();
  const result = {
    ok: true,
    spreadsheetId: dbSs.getId(),
    spreadsheetName: dbSs.getName(),
    sheets: []
  };

  const indexSheet = ensureCustomerSearchIndexSheet_(dbSs);
  result.sheets.push(indexSheet.getName());

  const noticeSheet = ensurePortalNoticeSheet_();
  result.sheets.push(noticeSheet.getName());

  const historySheet = ensureContactHistorySheet_(dbSs);
  result.sheets.push(historySheet.getName());

  const todaySheet = ensurePortalTodaySheet_();
  result.sheets.push(todaySheet.getName());

  // PATCH F: 진행현황 옵션/색상 등 운영 설정_DB를 준비합니다.
  try {
    const settingsResult = setupPortalSettingsSheet();
    result.sheets.push(settingsResult.sheetName || '설정_DB');
  } catch (err) {
    result.sheets.push('설정_DB:ERROR:' + err.message);
  }

  // v34: 영업지원요청은 웹앱_DB가 아니라 마스터시트 파일에 둡니다.
  // setup 시 헤더만 보정하고 결과에는 위치를 명확히 표시합니다.
  const supportSheet = ensurePortalSupportSheet_();
  result.sheets.push('MASTER:' + supportSheet.getName());

  SpreadsheetApp.flush();
  return result;
}

function getPortalMailRuntimeInfo() {
  let activeEmail = '';
  let effectiveEmail = '';
  try { activeEmail = Session.getActiveUser().getEmail(); } catch (err) {}
  try { effectiveEmail = Session.getEffectiveUser().getEmail(); } catch (err) {}

  return {
    ok: true,
    activeUserEmail: activeEmail || '',
    effectiveUserEmail: effectiveEmail || '',
    hasMailAutomationService: (typeof MailAutomationService !== 'undefined'),
    hasSendMailFromDialog: (typeof sendMailFromDialog === 'function'),
    hasGetMailRunProgress: (typeof getMailRunProgress === 'function'),
    hasCancelMailRun: (typeof cancelMailRun === 'function'),
    hasSendMailFromPortalPayload: (typeof sendMailFromPortalPayload === 'function'),
    hasServiceStandardContractKey: isExistingMailFileKeySupported_('serviceStandardContract'),
    masterSpreadsheetId: PORTAL_CONFIG.MASTER_SPREADSHEET_ID,
    webAppDbSpreadsheetId: String(PORTAL_CONFIG.WEBAPP_DB_SPREADSHEET_ID || '').trim() || '(active/fallback)',
    supportRequestLocation: 'MASTER_SPREADSHEET_ID'
  };
}

function debugPortalSupportSourceV56() {
  const masterSs = getMasterSpreadsheet_();
  const sheet = getPortalSupportSheetFromMasterV56_(masterSs);
  const headerMap = getPortalSupportHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), PORTAL_CONFIG.SUPPORT_HEADERS.length);
  const start = Math.max(PORTAL_CONFIG.SUPPORT_DATA_START_ROW, lastRow - 9);
  const values = lastRow >= PORTAL_CONFIG.SUPPORT_DATA_START_ROW
    ? sheet.getRange(start, 1, lastRow - start + 1, lastCol).getDisplayValues()
    : [];
  const sample = values.map(function(row, idx) {
    return buildPortalSupportRowObject_(row, start + idx, headerMap);
  }).filter(function(item) {
    return item && (item.receiptNo || item.requestType || item.customerNo || item.customerName || item.requestText || item.status);
  });
  return {
    ok: true,
    masterSpreadsheetId: masterSs.getId(),
    masterSpreadsheetName: masterSs.getName(),
    supportSheetName: sheet.getName(),
    supportSheetId: sheet.getSheetId(),
    configuredSupportSheetName: PORTAL_CONFIG.SUPPORT_SHEET_NAME,
    headerRow: PORTAL_CONFIG.SUPPORT_HEADER_ROW,
    dataStartRow: PORTAL_CONFIG.SUPPORT_DATA_START_ROW,
    lastRow: lastRow,
    lastCol: lastCol,
    headerMap: headerMap,
    sampleLastRows: sample.slice(-10)
  };
}

function debugPortalMailWorkerConfig_() {
  const props = PropertiesService.getScriptProperties();
  const url = String(props.getProperty('MAIL_WORKER_WEBAPP_URL') || '').trim();
  const secret = String(props.getProperty('MAIL_WORKER_SHARED_SECRET') || '').trim();
  const info = {
    portalScriptId: (() => { try { return ScriptApp.getScriptId(); } catch (e) { return ''; } })(),
    hasWorkerUrl: !!url,
    workerUrlHead: url ? url.slice(0, 70) + (url.length > 70 ? '...' : '') : '',
    hasSecret: !!secret,
    secretLength: secret.length,
    secretHash16: secret ? Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, secret)).replace(/=+$/g, '').slice(0, 16) : ''
  };
  Logger.log(JSON.stringify(info, null, 2));
  return info;
}

function testPortalMailWorkerHealthV42_() {
  const props = PropertiesService.getScriptProperties();
  const workerUrl = String(props.getProperty('MAIL_WORKER_WEBAPP_URL') || '').trim();
  const workerSecret = String(props.getProperty('MAIL_WORKER_SHARED_SECRET') || '').trim();
  if (!workerUrl) throw new Error('MAIL_WORKER_WEBAPP_URL이 없습니다.');
  if (!workerSecret) throw new Error('MAIL_WORKER_SHARED_SECRET이 없습니다.');

  const res = UrlFetchApp.fetch(workerUrl, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify({ secret: workerSecret, action: 'health' }),
    followRedirects: true,
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const text = res.getContentText();
  Logger.log('HTTP ' + code + '\n' + text);
  if (code < 200 || code >= 300) throw new Error('Worker health check HTTP 실패: ' + code + '\n' + text);
  const data = JSON.parse(text);
  if (!data.ok) throw new Error('Worker health check 실패: ' + text);
  return data;
}

function testPortalMailWorkerHealthV44() {
  const props = PropertiesService.getScriptProperties();

  const workerUrl = String(props.getProperty('MAIL_WORKER_WEBAPP_URL') || '').trim();
  const workerSecret = String(props.getProperty('MAIL_WORKER_SHARED_SECRET') || '').trim();

  const info = {
    portalScriptId: (() => {
      try { return ScriptApp.getScriptId(); } catch (e) { return ''; }
    })(),
    hasWorkerUrl: !!workerUrl,
    workerUrlHead: workerUrl ? workerUrl.slice(0, 90) + (workerUrl.length > 90 ? '...' : '') : '',
    hasSecret: !!workerSecret,
    secretLength: workerSecret.length,
    secretHash16: workerSecret
      ? Utilities.base64EncodeWebSafe(
          Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, workerSecret)
        ).replace(/=+$/g, '').slice(0, 16)
      : ''
  };

  if (!workerUrl) {
    Logger.log(JSON.stringify(info, null, 2));
    throw new Error('포털 프로젝트 Script Properties에 MAIL_WORKER_WEBAPP_URL이 없습니다.');
  }

  if (!workerSecret) {
    Logger.log(JSON.stringify(info, null, 2));
    throw new Error('포털 프로젝트 Script Properties에 MAIL_WORKER_SHARED_SECRET이 없습니다.');
  }

  const body = {
    secret: workerSecret,
    action: 'health',
    payload: {},
    client: {
      portalScriptId: info.portalScriptId,
      requestedAt: new Date().toISOString()
    }
  };

  const res = UrlFetchApp.fetch(workerUrl, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify(body),
    followRedirects: true,
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText() || '';

  let parsed = null;
  try {
    parsed = JSON.parse(text);
  } catch (err) {}

  const result = {
    ok: code >= 200 && code < 300 && parsed && parsed.ok === true,
    httpCode: code,
    portalConfig: info,
    responseTextHead: text.slice(0, 2000),
    parsed: parsed
  };

  Logger.log(JSON.stringify(result, null, 2));

  if (!result.ok) {
    throw new Error(
      '메일 Worker health check 실패\n' +
      'HTTP: ' + code + '\n' +
      '응답 앞부분:\n' + text.slice(0, 2000)
    );
  }

  return result;
}

function testPortalMailWorkerSendV44() {
  const testPayload = {
    rowNo: 3,                 // 테스트할 마스터시트 행 번호로 수정
    mode: 'TEST',
    testInput: 'bang@s1samsung.com',
    selectedKeys: ['quote'],  // 견적서만 테스트
    manualTo: [],
    manualCc: ['bang@s1samsung.com'],
    removedCc: [],
    runId: 'PORTAL_WORKER_TEST_' + Utilities.getUuid()
  };

  const result = callExistingMailAutomation_(testPayload);
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

