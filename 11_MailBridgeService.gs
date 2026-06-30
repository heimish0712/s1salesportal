/***************************************
 * S1 Sales Portal - 11_MailBridgeService.gs
 * 분리일: 2026-06-19
 * 원칙: 기능 변경 없이 최신 단일 파일을 물리적으로 분리
 ***************************************/


function preparePortalMailFilesForReview(payload) {
  payload = payload || {};
  const target = assertCustomerTarget_(payload, '파일 확인/수정', { readObject: true });
  const rowNo = target.rowNo;
  const targetCustomerNo = target.customerNo;
  const selectedKeys = Array.isArray(payload.selectedKeys) ? payload.selectedKeys.map(String).filter(Boolean) : [];

  if (!selectedKeys.length) throw new Error('확인/수정할 발송자료를 하나 이상 선택하세요.');
  if (typeof CONFIG !== 'undefined' && CONFIG && Array.isArray(CONFIG.FILE_DEFINITIONS)) {
    const unsupportedKeys = selectedKeys.filter(function(key) {
      return !CONFIG.FILE_DEFINITIONS.some(function(def) { return def && def.key === key; });
    });
    if (unsupportedKeys.length) {
      throw new Error('기존 메일자동화 CONFIG.FILE_DEFINITIONS에 없는 발송자료 key입니다: ' + unsupportedKeys.join(', '));
    }
  }

  const rowObj = target.obj || readMasterRowObject_(target.sheet, rowNo);
  const company = getCompanyValue_(rowObj) || '(회사명 없음)';
  const reviewPayload = {
    rowNo: rowNo,
    customerNo: targetCustomerNo,
    selectedKeys: selectedKeys,
    runId: String(payload.runId || Utilities.getUuid()),
    compareQuoteSheets: payload.compareQuoteSheets || payload.selectedCompareQuoteSheets || null,
    excludedCompareQuoteSheets: payload.excludedCompareQuoteSheets || payload.excludedCompareSheets || null
  };

  const guard = beginPortalIdempotentRequestP26_('MAIL_REVIEW', reviewPayload.runId, { runningTtlMs: 15 * 60 * 1000, doneTtlMs: 60 * 60 * 1000 });
  if (guard && guard.duplicate) {
    if (guard.running) throwPortalDuplicateRunningP26_('파일 확인/수정');
    return { ok: true, duplicate: true, message: '이미 처리된 파일 확인/수정 요청입니다.', company: company, rowNo: rowNo, customerNo: targetCustomerNo, selectedFiles: selectedKeys.map(k => getPortalSendFileLabel_(k)) };
  }
  let result = null;
  try {
    result = callExistingMailAutomationReviewV75_(reviewPayload);
    finishPortalIdempotentRequestP26_(guard, { ok: true, rowNo: rowNo, customerNo: targetCustomerNo });
  } catch (err) {
    failPortalIdempotentRequestP26_(guard);
    throw err;
  }

  try {
    appendPortalActivityLog_({
      actionType: '자료발송',
      screen: '견적/자료 발송',
      rowNo: rowNo,
      customerNo: targetCustomerNo || getMasterFieldValue_(rowObj, 'customerNo') || rowObj['고객번호'] || '',
      company: company,
      summary: '파일 확인/수정: ' + selectedKeys.map(function(k) { return getPortalSendFileLabel_(k); }).join(', '),
      detail: { mode: 'REVIEW', selectedKeys: selectedKeys, runId: reviewPayload.runId, reviewSessionId: result && result.reviewSessionId || '' }
    });
  } catch (err) {}

  return {
    ok: true,
    company: company,
    rowNo: rowNo,
    customerNo: targetCustomerNo,
    selectedFiles: selectedKeys.map(k => getPortalSendFileLabel_(k)),
    reviewSessionId: result && result.reviewSessionId || '',
    requestNo: result && result.requestNo || '',
    folderUrl: result && result.folderUrl || '',
    fileCount: result && result.fileCount || 0,
    files: result && result.files || [],
    result: result || null
  };
}

function sendPortalSingleMail(payload) {
  payload = payload || {};
  const target = assertCustomerTarget_(payload, '자료발송', { readObject: true });
  const rowNo = target.rowNo;
  const targetCustomerNo = target.customerNo;
  const mode = String(payload.mode || '').toUpperCase();
  const selectedKeys = Array.isArray(payload.selectedKeys) ? payload.selectedKeys.map(String).filter(Boolean) : [];

  if (mode !== 'TEST' && mode !== 'CUSTOMER') throw new Error('발송 모드가 올바르지 않습니다.');
  if (!selectedKeys.length) throw new Error('발송자료를 하나 이상 선택하세요.');
  if (typeof CONFIG !== 'undefined' && CONFIG && Array.isArray(CONFIG.FILE_DEFINITIONS)) {
    const unsupportedKeys = selectedKeys.filter(function(key) {
      return !CONFIG.FILE_DEFINITIONS.some(function(def) { return def && def.key === key; });
    });
    if (unsupportedKeys.length) {
      throw new Error('기존 메일자동화 CONFIG.FILE_DEFINITIONS에 없는 발송자료 key입니다: ' + unsupportedKeys.join(', '));
    }
  }

  const sheet = target.sheet;
  const rowObj = target.obj || readMasterRowObject_(sheet, rowNo);
  const company = getCompanyValue_(rowObj) || '(회사명 없음)';
  const email = getMasterFieldValue_(rowObj, 'email') || getValueByHeaderCandidates_(rowObj, ['담당자 이메일 주소', '이메일주소', '이메일']);
  const manualToList = normalizePortalEmailListForServer_(payload.manualTo);

  if (mode === 'CUSTOMER' && !String(email || '').trim() && !manualToList.length) {
    throw new Error('고객에게 발송하려면 담당자 이메일 주소가 필요합니다. 마스터 이메일을 입력하거나 팝업에서 수신자를 추가하세요.');
  }

  const testInput = mode === 'TEST' ? normalizePortalInternalTestEmail_(payload.testInput) : String(payload.testInput || '').trim();

  const sendPayload = {
    rowNo: rowNo,
    mode: mode,
    selectedKeys: selectedKeys,
    testInput: testInput,
    manualTo: payload.manualTo || null,
    manualCc: payload.manualCc || null,
    removedCc: payload.removedCc || [],
    reviewSessionId: String(payload.reviewSessionId || '').trim(),
    runId: String(payload.runId || Utilities.getUuid()),
    customerNo: targetCustomerNo
  };

  const guard = beginPortalIdempotentRequestP26_('MAIL_SEND', sendPayload.runId, { runningTtlMs: 20 * 60 * 1000, doneTtlMs: 2 * 60 * 60 * 1000 });
  if (guard && guard.duplicate) {
    if (guard.running) throwPortalDuplicateRunningP26_('메일 발송');
    return {
      ok: true,
      duplicate: true,
      message: '이미 처리된 메일 발송 요청입니다.',
      company: company,
      mode: mode,
      selectedFiles: selectedKeys.map(k => getPortalSendFileLabel_(k)),
      result: null,
      indexUpdate: null
    };
  }
  let result = null;
  try {
    result = callExistingMailAutomation_(sendPayload);
    finishPortalIdempotentRequestP26_(guard, { ok: true, rowNo: rowNo, customerNo: targetCustomerNo, mode: mode });
  } catch (err) {
    failPortalIdempotentRequestP26_(guard);
    throw err;
  }

  appendContactHistory_({
    customerNo: targetCustomerNo || getMasterFieldValue_(rowObj, 'customerNo') || rowObj['고객번호'] || '',
    company: company,
    rowNo: rowNo,
    type: '자료발송',
    contactRound: '',
    method: mode === 'TEST' ? '테스트발송' : '고객발송',
    status: '',
    contactText: '웹앱에서 자료발송 완료: ' + selectedKeys.map(k => getPortalSendFileLabel_(k)).join(', '),
    note: '',
    nextAction: '',
    nextActionAt: '',
    reflectToMasterMemo: false
  });

  // PATCH R: 메일 Worker가 마스터시트의 마지막발송/발송일시/메모 등을 갱신할 수 있으므로
  // 발송 성공 후 해당 고객의 검색인덱스_DB row를 마스터 기준으로 다시 동기화합니다.
  let indexUpdate = null;
  try {
    indexUpdate = updateCustomerSearchIndexFullAfterMutation_({ rowNo: rowNo, customerNo: targetCustomerNo }, 'MAIL_SEND_RESULT');
  } catch (err) {
    Logger.log('검색인덱스 자료발송 후 갱신 실패: ' + (err && err.stack || err));
    try { indexUpdate = markCustomerSearchIndexDirty_('MAIL_SEND_RESULT', err && err.message ? err.message : String(err)); } catch (e) {}
  }

  try {
    appendPortalActivityLog_({
      actionType: '자료발송',
      screen: '견적/자료 발송',
      rowNo: rowNo,
      customerNo: targetCustomerNo || getMasterFieldValue_(rowObj, 'customerNo') || rowObj['고객번호'] || '',
      company: company,
      summary: (mode === 'TEST' ? '테스트 발송: ' : '고객 발송: ') + selectedKeys.map(function(k) { return getPortalSendFileLabel_(k); }).join(', '),
      detail: { mode: mode, selectedKeys: selectedKeys, runId: sendPayload.runId }
    });
  } catch (err) {}

  return {
    ok: true,
    company: company,
    mode: mode,
    selectedFiles: selectedKeys.map(k => getPortalSendFileLabel_(k)),
    message: (mode === 'TEST' ? '테스트 발송 완료' : '고객 발송 완료'),
    result: result || null,
    indexUpdate: indexUpdate
  };
}

function normalizePortalEmailListForServer_(value) {
  const arr = Array.isArray(value) ? value : String(value || '').split(/[;,\s]+/);
  const out = [];
  arr.forEach(function(v) {
    v = String(v || '').trim();
    if (!v || v === '-') return;
    if (v.indexOf('@') < 0) v += '@s1samsung.com';
    if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v) && out.map(function(x){ return x.toLowerCase(); }).indexOf(v.toLowerCase()) < 0) {
      out.push(v);
    }
  });
  return out;
}

function normalizePortalInternalTestEmail_(value) {
  let email = String(value || '').trim();
  if (!email) throw new Error('나에게 테스트 발송할 이메일 또는 아이디를 입력하세요.');
  if (email.indexOf('@') < 0) email += '@s1samsung.com';
  if (!/^[^@\s]+@s1samsung\.com$/i.test(email)) {
    throw new Error('테스트 발송 주소는 @s1samsung.com 형식이어야 합니다. 외부 메일 테스트는 고객발송 모드에서 수신자를 직접 추가해 주세요.');
  }
  return email;
}

function applyTemporaryPortalRecipientEmail_(sheet, rowNo, manualTo) {
  const toList = normalizePortalEmailListForServer_(manualTo);
  if (!toList.length) return null;
  const headerMap = getHeaderMap_(sheet);
  const headers = masterFieldHeaders_('email', ['담당자 이메일 주소', '이메일주소', '이메일']);
  let col = findFirstExistingHeaderCol_(headerMap, headers);
  if (!col) throw new Error('수신자 임시 반영 실패: 마스터시트에서 담당자 이메일 주소 헤더를 찾지 못했습니다.');
  const cell = sheet.getRange(rowNo, col);
  const oldValue = cell.getValue();
  cell.setValue(toList.join(', '));
  return { rowNo: rowNo, col: col, oldValue: oldValue };
}

function restoreTemporaryPortalRecipientEmail_(sheet, info) {
  if (!info || !info.rowNo || !info.col) return;
  sheet.getRange(info.rowNo, info.col).setValue(info.oldValue || '');
}

function applyTemporarySendCheckboxes_(sheet, rowNo, selectedKeys) {
  const headerMap = getHeaderMap_(sheet);
  const selectedSet = {};
  selectedKeys.forEach(k => selectedSet[k] = true);

  const targets = [];
  PORTAL_SEND_FILE_DEFINITIONS.forEach(def => {
    const col = findFirstHeaderColumn_(headerMap, def.headers);
    if (col) targets.push({ key: def.key, label: def.label, col: col });
  });

  const missingSelected = PORTAL_SEND_FILE_DEFINITIONS
    .filter(def => selectedSet[def.key] && !targets.some(t => t.key === def.key));
  if (missingSelected.length) {
    throw new Error('마스터시트에서 발송자료 체크 헤더를 찾지 못했습니다: ' + missingSelected.map(d => d.label).join(', '));
  }

  const restore = targets.map(t => ({ col: t.col, value: sheet.getRange(rowNo, t.col).getValue() }));
  targets.forEach(t => sheet.getRange(rowNo, t.col).setValue(!!selectedSet[t.key]));
  return { rowNo: rowNo, restore: restore };
}

function restoreTemporarySendCheckboxes_(sheet, restoreInfo) {
  if (!restoreInfo || !restoreInfo.rowNo || !restoreInfo.restore) return;
  restoreInfo.restore.forEach(item => {
    sheet.getRange(restoreInfo.rowNo, item.col).setValue(item.value);
  });
}

function findFirstHeaderColumn_(headerMap, headers) {
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    if (headerMap[h]) return headerMap[h];
  }
  return 0;
}


function callExistingMailAutomationReviewV75_(reviewPayload) {
  // v75: 포털 [견적/자료 발송]의 파일 확인/수정 버튼은 기존 메일자동화 Worker에 위임합니다.
  // Worker 쪽에는 preparePortalMailFilesForReview 액션이 있어야 하며,
  // 구버전 이름을 쓰는 배포본 호환을 위해 몇 가지 액션명을 순차 시도합니다.
  const actions = [
    'preparePortalMailFilesForReview',
    'preparePortalReviewFiles',
    'prepareMailFilesForReview'
  ];
  let lastErr = null;
  for (let i = 0; i < actions.length; i++) {
    try {
      return callMailWorkerActionV44_(actions[i], reviewPayload);
    } catch (err) {
      lastErr = err;
      const msg = String(err && err.message || err || '');
      // 권한/URL/Secret/실행 오류는 다음 액션명으로 해결되지 않으므로 바로 중단합니다.
      if (msg.indexOf('MAIL_WORKER_WEBAPP_URL') >= 0 ||
          msg.indexOf('MAIL_WORKER_SHARED_SECRET') >= 0 ||
          msg.indexOf('HTTP ') >= 0 && msg.indexOf('응답:') >= 0 && msg.indexOf('unknown') < 0 && msg.indexOf('지원하지') < 0 && msg.indexOf('알 수 없는') < 0) {
        throw err;
      }
    }
  }
  throw lastErr || new Error('메일 Worker 파일 확인/수정 액션 호출 실패');
}


/***************************************
 * v83 Mail Worker 설정 안정화
 * - Script Properties UI에서 값이 사라지거나 다른 키로 저장되어도 최대한 복구
 * - 필요 시 PORTAL_MAIL_WORKER_FALLBACK_CONFIG_V83 전역 객체를 별도 파일에 두면 코드 fallback으로 자동 복구
 ***************************************/
function getPortalMailWorkerConfigV83_() {
  const props = PropertiesService.getScriptProperties();
  const urlKeys = [
    'MAIL_WORKER_WEBAPP_URL',
    'PORTAL_MAIL_WORKER_WEBAPP_URL',
    'MAIL_WORKER_WEBAPP_URL_BACKUP'
  ];
  const secretKeys = [
    'MAIL_WORKER_SHARED_SECRET',
    'PORTAL_MAIL_WORKER_SHARED_SECRET',
    'MAIL_WORKER_SHARED_SECRET_BACKUP'
  ];

  let workerUrl = getFirstPortalPropValueV83_(props, urlKeys);
  let workerSecret = getFirstPortalPropValueV83_(props, secretKeys);

  const fallback = getPortalMailWorkerFallbackConfigV83_();
  if (!workerUrl && fallback.webappUrl) workerUrl = fallback.webappUrl;
  if (!workerSecret && fallback.sharedSecret) workerSecret = fallback.sharedSecret;

  // 값이 하나라도 발견되면 표준 키 + backup 키에 즉시 재저장합니다.
  // 즉, 설정 UI에서 표준 키가 비어 보여도 다음 실행 시 다시 살아나게 합니다.
  if (workerUrl) {
    urlKeys.forEach(function(key) {
      try { props.setProperty(key, workerUrl); } catch (err) {}
    });
  }
  if (workerSecret) {
    secretKeys.forEach(function(key) {
      try { props.setProperty(key, workerSecret); } catch (err) {}
    });
  }

  return {
    workerUrl: String(workerUrl || '').trim(),
    workerSecret: String(workerSecret || '').trim(),
    hasUrl: !!String(workerUrl || '').trim(),
    hasSecret: !!String(workerSecret || '').trim(),
    source: fallback.source || 'scriptProperties'
  };
}

function getFirstPortalPropValueV83_(props, keys) {
  for (let i = 0; i < keys.length; i++) {
    const v = String(props.getProperty(keys[i]) || '').trim();
    if (v) return v;
  }
  return '';
}

function getPortalMailWorkerFallbackConfigV83_() {
  try {
    if (typeof PORTAL_MAIL_WORKER_FALLBACK_CONFIG_V83 !== 'undefined' && PORTAL_MAIL_WORKER_FALLBACK_CONFIG_V83) {
      return {
        webappUrl: String(PORTAL_MAIL_WORKER_FALLBACK_CONFIG_V83.WEBAPP_URL || PORTAL_MAIL_WORKER_FALLBACK_CONFIG_V83.webappUrl || '').trim(),
        sharedSecret: String(PORTAL_MAIL_WORKER_FALLBACK_CONFIG_V83.SHARED_SECRET || PORTAL_MAIL_WORKER_FALLBACK_CONFIG_V83.sharedSecret || '').trim(),
        source: 'fallbackConfigObject'
      };
    }
  } catch (err) {}
  return { webappUrl: '', sharedSecret: '', source: 'scriptProperties' };
}

function savePortalMailWorkerConfigV83_(workerUrl, workerSecret) {
  workerUrl = String(workerUrl || '').trim();
  workerSecret = String(workerSecret || '').trim();
  if (!workerUrl) throw new Error('MAIL_WORKER_WEBAPP_URL 값이 비어 있습니다.');
  if (!/^https:\/\/script\.google\.com\/macros\/s\//.test(workerUrl)) {
    throw new Error('MAIL_WORKER_WEBAPP_URL은 Google Apps Script 웹앱 /exec URL이어야 합니다.');
  }
  if (!workerSecret) throw new Error('MAIL_WORKER_SHARED_SECRET 값이 비어 있습니다.');

  const props = PropertiesService.getScriptProperties();
  [
    'MAIL_WORKER_WEBAPP_URL',
    'PORTAL_MAIL_WORKER_WEBAPP_URL',
    'MAIL_WORKER_WEBAPP_URL_BACKUP'
  ].forEach(function(key) { props.setProperty(key, workerUrl); });

  [
    'MAIL_WORKER_SHARED_SECRET',
    'PORTAL_MAIL_WORKER_SHARED_SECRET',
    'MAIL_WORKER_SHARED_SECRET_BACKUP'
  ].forEach(function(key) { props.setProperty(key, workerSecret); });

  return getPortalMailWorkerConfigDebugV83_();
}

function setupPortalMailWorkerConfigV83() {
  const ui = SpreadsheetApp.getUi();
  const current = getPortalMailWorkerConfigV83_();

  const urlRes = ui.prompt(
    '메일 Worker WebApp URL 저장',
    '메일자동화 Worker 웹앱 /exec URL을 붙여넣으세요.\n현재값: ' + maskPortalSecretLikeValueV83_(current.workerUrl),
    ui.ButtonSet.OK_CANCEL
  );
  if (urlRes.getSelectedButton() !== ui.Button.OK) return { ok: false, message: '취소됨' };

  const secretRes = ui.prompt(
    '메일 Worker Shared Secret 저장',
    '포털 프로젝트와 메일 Worker 프로젝트에 공통으로 쓸 secret 값을 붙여넣으세요.\n현재값: ' + maskPortalSecretLikeValueV83_(current.workerSecret),
    ui.ButtonSet.OK_CANCEL
  );
  if (secretRes.getSelectedButton() !== ui.Button.OK) return { ok: false, message: '취소됨' };

  const result = savePortalMailWorkerConfigV83_(urlRes.getResponseText(), secretRes.getResponseText());
  ui.alert('메일 Worker 설정 저장 완료', JSON.stringify(result, null, 2), ui.ButtonSet.OK);
  return result;
}

function debugPortalMailWorkerConfigV83() {
  const result = getPortalMailWorkerConfigDebugV83_();
  Logger.log(JSON.stringify(result, null, 2));
  try {
    SpreadsheetApp.getUi().alert('메일 Worker 설정 확인', JSON.stringify(result, null, 2), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {}
  return result;
}

function getPortalMailWorkerConfigDebugV83_() {
  const cfg = getPortalMailWorkerConfigV83_();
  const props = PropertiesService.getScriptProperties();
  return {
    ok: cfg.hasUrl && cfg.hasSecret,
    source: cfg.source,
    hasMailWorkerUrl: cfg.hasUrl,
    mailWorkerUrlMasked: maskPortalSecretLikeValueV83_(cfg.workerUrl),
    hasMailWorkerSecret: cfg.hasSecret,
    mailWorkerSecretLength: cfg.workerSecret.length,
    mailWorkerSecretHash16: cfg.workerSecret
      ? Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, cfg.workerSecret)).replace(/=+$/g, '').slice(0, 16)
      : '',
    savedKeys: {
      MAIL_WORKER_WEBAPP_URL: !!String(props.getProperty('MAIL_WORKER_WEBAPP_URL') || '').trim(),
      PORTAL_MAIL_WORKER_WEBAPP_URL: !!String(props.getProperty('PORTAL_MAIL_WORKER_WEBAPP_URL') || '').trim(),
      MAIL_WORKER_WEBAPP_URL_BACKUP: !!String(props.getProperty('MAIL_WORKER_WEBAPP_URL_BACKUP') || '').trim(),
      MAIL_WORKER_SHARED_SECRET: !!String(props.getProperty('MAIL_WORKER_SHARED_SECRET') || '').trim(),
      PORTAL_MAIL_WORKER_SHARED_SECRET: !!String(props.getProperty('PORTAL_MAIL_WORKER_SHARED_SECRET') || '').trim(),
      MAIL_WORKER_SHARED_SECRET_BACKUP: !!String(props.getProperty('MAIL_WORKER_SHARED_SECRET_BACKUP') || '').trim()
    }
  };
}

function maskPortalSecretLikeValueV83_(value) {
  value = String(value || '').trim();
  if (!value) return '(없음)';
  if (value.length <= 12) return value.slice(0, 2) + '***' + value.slice(-2);
  return value.slice(0, 10) + '...' + value.slice(-6);
}

function callExistingMailAutomation_(sendPayload) {
  // v42: 메일 발송은 포털 프로젝트 안의 복사본 코드가 아니라,
  // "시트에서 실제로 정상 발송되는" 기존 자동메일 프로젝트의 Web Worker로 위임합니다.
  // 이유: 같은 v65 코드라도 Apps Script 프로젝트/배포/실행환경이 다르면 하이웍스 API 응답이 달라질 수 있었음.
  // 이 방식은 하이웍스 발송부를 포털에서 재구현하지 않고, 정상 발송 확인된 프로젝트에서 그대로 실행합니다.
  const mailWorkerConfig = getPortalMailWorkerConfigV83_();
  const workerUrl = mailWorkerConfig.workerUrl;
  const workerSecret = mailWorkerConfig.workerSecret;

  if (!workerUrl) {
    throw new Error(
      'MAIL_WORKER_WEBAPP_URL 스크립트 속성이 없습니다.\n\n' +
      '시트에서 정상 발송되는 자동메일 프로젝트에 03_MAIL_WORKER_add_to_WORKING_mail_project_v42.gs를 추가/배포한 뒤,\n' +
      '그 웹앱 /exec URL을 포털 프로젝트의 Script Properties > MAIL_WORKER_WEBAPP_URL에 저장하세요.'
    );
  }
  if (!workerSecret) {
    throw new Error(
      'MAIL_WORKER_SHARED_SECRET 스크립트 속성이 없습니다.\n\n' +
      '포털 프로젝트와 메일 Worker 프로젝트 양쪽 Script Properties에 같은 긴 임의 문자열을 저장하세요.'
    );
  }

  const requestBody = {
    secret: workerSecret,
    action: 'sendPortalMail',
    payload: sendPayload,
    client: {
      portalScriptId: (() => { try { return ScriptApp.getScriptId(); } catch (e) { return ''; } })(),
      requestedAt: new Date().toISOString()
    }
  };

  const res = UrlFetchApp.fetch(workerUrl, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify(requestBody),
    followRedirects: true,
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText() || '';
  let data = null;
  try { data = JSON.parse(text); } catch (err) {}

  if (code < 200 || code >= 300) {
    throw new Error('메일 Worker 호출 실패 HTTP ' + code + '\n응답: ' + text.slice(0, 2000));
  }
  if (!data || data.ok !== true) {
    const msg = data && data.message ? data.message : text;
    const detail = data && data.detail ? ('\n상세: ' + String(data.detail).slice(0, 2500)) : '';
    throw new Error('메일 Worker 발송 실패: ' + msg + detail);
  }

  return data.result || data;
}

function getMailRunProgress(runId) {
  runId = String(runId || '').trim();
  if (!runId) {
    return {
      status: 'RUNNING',
      percent: 1,
      message: '진행률 조회 대기 중: runId 없음',
      elapsedSec: 0,
      remainingSec: null,
      progressProxy: 'portal_v44_no_runId'
    };
  }

  try {
    const result = callMailWorkerActionV44_('getProgress', { runId: runId });
    return normalizePortalWorkerProgressV44_(result, runId);
  } catch (err) {
    // 중요:
    // 진행률 조회 실패가 실제 메일 발송 실패는 아닙니다.
    // 여기서 throw 하면 HTML failureHandler가 떠서 "진행률 조회 생략"으로 멈췄습니다.
    // v44에서는 절대 throw하지 않고 RUNNING 상태를 반환하여 계속 재조회하게 합니다.
    return buildPortalProgressRetryFallbackV44_(runId, err);
  }
}

function cancelMailRun(runId) {
  runId = String(runId || '').trim();
  if (!runId) return { ok: false, message: '취소할 runId가 없습니다.' };

  try {
    return callMailWorkerActionV44_('cancelRun', { runId: runId });
  } catch (err) {
    return {
      ok: false,
      message: '메일 Worker 취소 요청 실패: ' + String(err && err.message || err),
      detail: String(err && err.stack || err)
    };
  }
}

function callMailWorkerActionV44_(action, payload) {
  const mailWorkerConfig = getPortalMailWorkerConfigV83_();
  const workerUrl = mailWorkerConfig.workerUrl;
  const workerSecret = mailWorkerConfig.workerSecret;

  if (!workerUrl) throw new Error('MAIL_WORKER_WEBAPP_URL 스크립트 속성이 없습니다. setupPortalMailWorkerConfigV83()로 다시 저장하세요.');
  if (!workerSecret) throw new Error('MAIL_WORKER_SHARED_SECRET 스크립트 속성이 없습니다. setupPortalMailWorkerConfigV83()로 다시 저장하세요.');

  const res = UrlFetchApp.fetch(workerUrl, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify({
      secret: workerSecret,
      action: action,
      payload: payload || {},
      client: {
        portalScriptId: (() => { try { return ScriptApp.getScriptId(); } catch (e) { return ''; } })(),
        requestedAt: new Date().toISOString()
      }
    }),
    followRedirects: true,
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText() || '';
  let data = null;
  try { data = JSON.parse(text); } catch (err) {}

  if (code < 200 || code >= 300) {
    throw new Error('메일 Worker 호출 실패 HTTP ' + code + '\n응답: ' + text.slice(0, 2000));
  }

  if (!data || data.ok !== true) {
    const msg = data && data.message ? data.message : text;
    const detail = data && data.detail ? ('\n상세: ' + String(data.detail).slice(0, 2500)) : '';
    throw new Error('메일 Worker 처리 실패: ' + msg + detail);
  }

  return data.result || data;
}

function normalizePortalWorkerProgressV44_(progress, runId) {
  progress = progress || {};
  const status = String(progress.status || '').trim() || 'RUNNING';
  const percent = Math.max(0, Math.min(100, Number(progress.percent || 0)));
  return {
    status: status,
    percent: percent,
    message: String(progress.message || (status === 'DONE' ? '메일 발송 완료' : '메일 발송 처리 중')),
    startedAt: progress.startedAt || null,
    updatedAt: progress.updatedAt || Date.now(),
    elapsedSec: progress.elapsedSec == null ? 0 : progress.elapsedSec,
    remainingSec: progress.remainingSec == null ? null : progress.remainingSec,
    runId: progress.runId || runId,
    progressProxy: 'worker_v44'
  };
}

function buildPortalProgressRetryFallbackV44_(runId, err) {
  const cache = CacheService.getScriptCache();
  const key = 'PORTAL_PROGRESS_RETRY_START_V44_' + runId;
  let startedAt = Number(cache.get(key) || 0);
  if (!startedAt) {
    startedAt = Date.now();
    cache.put(key, String(startedAt), 21600);
  }
  const elapsedSec = Math.max(0, Math.round((Date.now() - startedAt) / 1000));
  const msg = String(err && err.message || err || '알 수 없는 진행률 조회 오류');
  return {
    status: 'RUNNING',
    percent: Math.min(80, Math.max(8, 8 + Math.floor(elapsedSec / 2))),
    message: '진행률 조회 재시도 중 · 실제 발송은 계속 진행 중',
    elapsedSec: elapsedSec,
    remainingSec: null,
    runId: runId,
    progressProxy: 'portal_v44_retry_fallback',
    progressError: msg.slice(0, 1000)
  };
}

function isExistingMailFileKeySupported_(key) {
  try {
    if (typeof CONFIG !== 'undefined' && CONFIG && Array.isArray(CONFIG.FILE_DEFINITIONS)) {
      return CONFIG.FILE_DEFINITIONS.some(def => def && def.key === key);
    }
  } catch (err) {}
  return key !== 'serviceStandardContract';
}

function getPortalSendFileLabel_(key) {
  const hit = PORTAL_SEND_FILE_DEFINITIONS.find(def => def.key === key);
  return hit ? hit.label : key;
}

function appendContactHistory_(payload) {
  payload = payload || {};

  const dbSs = getWebAppDbSpreadsheet_();
  const author = getPortalCurrentUserName_();

  return appendHistoryRecord_(dbSs, {
    customerNo: payload.customerNo || '',
    company: payload.company || '',
    rowNo: payload.rowNo || '',
    author: author,
    recordType: payload.type || payload.recordType || '자료발송',
    roundNo: payload.contactRound || payload.roundNo || '',
    method: payload.method || '',
    status: payload.status || '',
    content: payload.contactText || payload.content || '',
    note: payload.note || '',
    nextAction: payload.nextAction || '',
    nextActionAt: payload.nextActionAt || '',
    masterApplied: payload.reflectToMasterMemo === true ? 'Y' : 'N'
  });
}

function getPortalMailCompleteSoundDataUrlV48() {
  const fileId = PORTAL_MAIL_COMPLETE_SOUND_FILE_ID_V48;
  const cache = CacheService.getScriptCache();
  const cacheKey = 'PORTAL_MAIL_DONE_SOUND_V48_' + fileId;

  try {
    const cached = cache.get(cacheKey);
    if (cached) {
      return { ok: true, fileId: fileId, dataUrl: cached, cached: true };
    }
  } catch (err) {}

  const blob = DriveApp.getFileById(fileId).getBlob();
  const mimeType = blob.getContentType() || 'audio/mpeg';
  const dataUrl = 'data:' + mimeType + ';base64,' + Utilities.base64Encode(blob.getBytes());

  // CacheService 단일 값 제한이 있어 작은 오디오만 캐시합니다.
  try {
    if (dataUrl.length < 90000) cache.put(cacheKey, dataUrl, 21600);
  } catch (err) {}

  return {
    ok: true,
    fileId: fileId,
    mimeType: mimeType,
    dataUrl: dataUrl,
    cached: false,
    length: dataUrl.length
  };
}



/***************************************
 * P447 발주번호 생성 알림 메일 자동발송 큐
 * - 발주하기 버튼 응답을 지연시키지 않기 위해 포탈은 큐 적재만 수행
 * - 백그라운드 트리거가 메일 Worker action=sendOrderNotificationMail 호출
 ***************************************/
const PORTAL_ORDER_MAIL_QUEUE_SHEET_P447 = '발주메일발송큐';
const PORTAL_ORDER_MAIL_QUEUE_HEADERS_P447 = [
  '요청ID', '상태', '시도횟수', '생성일시', '최종시도일시', '완료일시', '오류',
  '고객번호', '마스터행', '계약번호', '계약행', '고객사명', '영업담당자'
];
const PORTAL_ORDER_MAIL_WORKER_ACTION_P447 = 'sendOrderNotificationMail';

function enqueuePortalOrderNotificationMailP447_(payload) {
  payload = payload || {};
  const customerNo = String(payload.customerNo || '').trim();
  const contractNo = String(payload.contractNo || '').trim();
  const company = String(payload.company || '').trim();
  const salesRep = String(payload.salesRep || '').trim();
  if (!contractNo || !company) return { ok: false, skipped: true, reason: 'contractNo/company empty' };

  const requestId = buildPortalOrderMailRequestIdP447_(customerNo, contractNo, company);
  const sheet = getPortalOrderMailQueueSheetP447_();
  const existing = findPortalOrderMailQueueRowByRequestIdP447_(sheet, requestId);
  if (existing && existing.rowNo) {
    const status = String(existing.values[1] || '').trim();
    if (status === '대기' || status === '발송중' || status === '완료') {
      return { ok: true, duplicate: true, requestId: requestId, status: status };
    }
    // 오류 상태면 재시도 가능하도록 다시 대기로 돌립니다.
    sheet.getRange(existing.rowNo, 2, 1, 6).setValues([['대기', existing.values[2] || 0, existing.values[3] || new Date(), '', '', '']]);
    ensurePortalOrderMailTriggerP447_();
    return { ok: true, retryQueued: true, requestId: requestId };
  }

  sheet.appendRow([
    requestId,
    '대기',
    0,
    new Date(),
    '',
    '',
    '',
    customerNo,
    Number(payload.rowNo) || '',
    contractNo,
    Number(payload.contractRowNo) || '',
    company,
    salesRep
  ]);

  ensurePortalOrderMailTriggerP447_();
  return { ok: true, queued: true, requestId: requestId };
}

function processPortalOrderNotificationMailQueueP447() {
  const sheet = getPortalOrderMailQueueSheetP447_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, processed: 0, pending: 0 };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return { ok: false, skipped: true, reason: 'lock busy' };

  let processed = 0;
  let pending = 0;
  try {
    const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_ORDER_MAIL_QUEUE_HEADERS_P447.length).getValues();
    for (let i = 0; i < values.length; i++) {
      const rowNo = i + 2;
      const row = values[i];
      const status = String(row[1] || '').trim();
      const attempts = Number(row[2]) || 0;
      if (status !== '대기' && status !== '오류') continue;
      if (status === '오류' && attempts >= 3) continue;
      pending++;
      if (processed >= 10) continue;

      const nextAttempts = attempts + 1;
      sheet.getRange(rowNo, 2, 1, 5).setValues([['발송중', nextAttempts, row[3] || new Date(), new Date(), '']]);

      try {
        const result = callMailWorkerActionP447_(PORTAL_ORDER_MAIL_WORKER_ACTION_P447, {
          requestId: String(row[0] || '').trim(),
          customerNo: String(row[7] || '').trim(),
          rowNo: Number(row[8]) || 0,
          contractNo: String(row[9] || '').trim(),
          contractRowNo: Number(row[10]) || 0,
          company: String(row[11] || '').trim(),
          salesRep: String(row[12] || '').trim(),
          to: ['master@s1samsung.com'],
          cc: []
        });
        sheet.getRange(rowNo, 2, 1, 5).setValues([['완료', nextAttempts, row[3] || new Date(), new Date(), new Date()]]);
        sheet.getRange(rowNo, 7).setValue('');
        processed++;
      } catch (err) {
        const msg = String(err && err.message || err || '').slice(0, 1000);
        sheet.getRange(rowNo, 2, 1, 6).setValues([['오류', nextAttempts, row[3] || new Date(), new Date(), '', msg]]);
        processed++;
      }
    }
  } finally {
    lock.releaseLock();
  }

  if (hasPendingPortalOrderMailQueueP447_()) ensurePortalOrderMailTriggerP447_();
  return { ok: true, processed: processed, pending: pending };
}

function callMailWorkerActionP447_(action, payload) {
  const cfg = getPortalMailWorkerConfigV83_();
  if (!cfg.workerUrl) throw new Error('MAIL_WORKER_WEBAPP_URL 스크립트 속성이 없습니다.');
  if (!cfg.workerSecret) throw new Error('MAIL_WORKER_SHARED_SECRET 스크립트 속성이 없습니다.');

  const res = UrlFetchApp.fetch(cfg.workerUrl, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify({
      secret: cfg.workerSecret,
      action: action,
      payload: payload || {},
      client: {
        portalScriptId: (() => { try { return ScriptApp.getScriptId(); } catch (e) { return ''; } })(),
        requestedAt: new Date().toISOString()
      }
    }),
    followRedirects: true,
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText() || '';
  let data = null;
  try { data = JSON.parse(text); } catch (err) {}
  if (code < 200 || code >= 300) throw new Error('메일 Worker 호출 실패 HTTP ' + code + '\n응답: ' + text.slice(0, 2000));
  if (!data || data.ok !== true) {
    const msg = data && data.message ? data.message : text;
    const detail = data && data.detail ? ('\n상세: ' + String(data.detail).slice(0, 2500)) : '';
    throw new Error('발주메일 Worker 실패: ' + msg + detail);
  }
  return data.result || data;
}

function getPortalOrderMailQueueSheetP447_() {
  const ss = getMasterSpreadsheet_();
  let sheet = ss.getSheetByName(PORTAL_ORDER_MAIL_QUEUE_SHEET_P447);
  if (!sheet) {
    sheet = ss.insertSheet(PORTAL_ORDER_MAIL_QUEUE_SHEET_P447);
    try { sheet.hideSheet(); } catch (err) {}
  }
  const current = sheet.getRange(1, 1, 1, PORTAL_ORDER_MAIL_QUEUE_HEADERS_P447.length).getValues()[0];
  const needHeader = PORTAL_ORDER_MAIL_QUEUE_HEADERS_P447.some(function(h, idx) { return String(current[idx] || '') !== h; });
  if (needHeader) {
    sheet.getRange(1, 1, 1, PORTAL_ORDER_MAIL_QUEUE_HEADERS_P447.length).setValues([PORTAL_ORDER_MAIL_QUEUE_HEADERS_P447]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findPortalOrderMailQueueRowByRequestIdP447_(sheet, requestId) {
  requestId = String(requestId || '').trim();
  if (!requestId) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_ORDER_MAIL_QUEUE_HEADERS_P447.length).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === requestId) return { rowNo: i + 2, values: values[i] };
  }
  return null;
}

function buildPortalOrderMailRequestIdP447_(customerNo, contractNo, company) {
  const base = [String(customerNo || '').trim(), String(contractNo || '').trim(), String(company || '').trim()].join('|');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, base);
  return 'ORDERMAIL-' + Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, '').slice(0, 20);
}

function ensurePortalOrderMailTriggerP447_() {
  const handler = 'processPortalOrderNotificationMailQueueP447';
  try {
    const exists = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction && t.getHandlerFunction() === handler; });
    if (exists) return;
    ScriptApp.newTrigger(handler).timeBased().after(60 * 1000).create();
  } catch (err) {
    // 트리거 생성 실패 시 큐에는 남아 있으므로 다음 수동/후속 실행에서 처리 가능
  }
}

function hasPendingPortalOrderMailQueueP447_() {
  const sheet = getPortalOrderMailQueueSheetP447_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const values = sheet.getRange(2, 2, lastRow - 1, 2).getValues();
  return values.some(function(row) {
    const status = String(row[0] || '').trim();
    const attempts = Number(row[1]) || 0;
    return status === '대기' || (status === '오류' && attempts < 3);
  });
}
