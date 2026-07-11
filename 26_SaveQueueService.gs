/***************************************
 * P473 Direct Save Fallback Queue
 * - 정상 저장은 마스터시트에 직접 반영합니다.
 * - 직접 저장이 일시 오류/서비스 busy로 실패한 경우만 저장큐_DB에 남깁니다.
 * - expectedValues 충돌은 덮어쓰지 않고 CONFLICT로 남깁니다.
 ***************************************/

const PORTAL_SAVE_QUEUE_P473 = {
  SHEET_NAME: '저장큐_DB',
  HEADERS: [
    '등록일시',
    '수정일시',
    '작업ID',
    '사용자',
    '세션ID',
    '고객번호',
    'rowNo',
    'methodName',
    'source',
    '상태',
    '우선순위',
    '시도횟수',
    'patchJson',
    'expectedValuesJson',
    'payloadJson',
    'resultJson',
    '마지막오류',
    '적용일시'
  ],
  STATUS: {
    QUEUED: 'QUEUED',
    RUNNING: 'RUNNING',
    DONE: 'DONE',
    RETRY: 'RETRY',
    CONFLICT: 'CONFLICT',
    FAIL: 'FAIL'
  },
  MAX_JOBS_PER_RUN: 5,
  STALE_RUNNING_MINUTES: 3,
  MAX_CONFLICT_REBASE_ATTEMPTS: 3,
  TRIGGER_HANDLER: 'processSaveQueueP473'
};

function getPortalSaveQueueSheetP473_() {
  const ss = getWebAppDbSpreadsheet_();
  const name = PORTAL_SAVE_QUEUE_P473.SHEET_NAME;
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  ensurePortalSaveQueueHeadersP473_(sheet);
  return sheet;
}

function ensurePortalSaveQueueHeadersP473_(sheet) {
  const headers = PORTAL_SAVE_QUEUE_P473.HEADERS;
  const width = Math.max(sheet.getLastColumn(), headers.length);
  let current = [];
  if (sheet.getLastRow() >= 1 && width > 0) {
    current = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function(v) { return String(v || '').trim(); });
  }
  const seen = {};
  current.forEach(function(h) { if (h) seen[h] = true; });
  let changed = current.filter(Boolean).length === 0;
  headers.forEach(function(h) {
    if (!seen[h]) {
      current.push(h);
      seen[h] = true;
      changed = true;
    }
  });
  if (changed) {
    sheet.getRange(1, 1, 1, current.length).setValues([current]);
    try { sheet.setFrozenRows(1); } catch (e) {}
  }
}

function getPortalSaveQueueHeaderIndexP473_(sheet) {
  ensurePortalSaveQueueHeadersP473_(sheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(function(v) { return String(v || '').trim(); });
  const idx = {};
  headers.forEach(function(h, i) { if (h) idx[h] = i + 1; });
  return idx;
}

function getPortalActiveUserEmailP473_() {
  try { return String(Session.getActiveUser().getEmail() || '').trim(); } catch (e) { return ''; }
}

function getPortalSaveOperationIdP473_(methodName, payload) {
  payload = payload || {};
  const existing = String(payload.clientOperationId || payload.operationId || '').trim();
  if (existing) return existing;
  return 'SQP473_' + String(methodName || 'save') + '_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000000);
}

function getPortalSavePatchP473_(methodName, payload) {
  payload = payload || {};
  methodName = String(methodName || '').trim();
  if (payload.values && typeof payload.values === 'object') return payload.values;
  if (Object.prototype.hasOwnProperty.call(payload, 'memo')) return { memo: payload.memo };
  if (methodName === 'savePortalSupportProcessThinP474') {
    return {
      handler: payload.handler || '',
      processContent: payload.processContent || '',
      completedAt: payload.completedAt || '',
      status: payload.status || '',
      autoSendCheck: payload.autoSendCheck || ''
    };
  }
  if (methodName === 'savePortalSupportRequesterEditP527') {
    return {
      requestType: payload.requestType || '',
      requestText: payload.requestText || ''
    };
  }
  if (methodName === 'savePortalSupportRequest') {
    return {
      requestType: payload.requestType || '',
      requester: payload.requester || '',
      customerNo: payload.customerNo || '',
      customerName: payload.customerName || '',
      requestText: payload.requestText || '',
      status: payload.status || '접수'
    };
  }
  if (methodName === 'savePortalNotice' || methodName === 'updatePortalNotice') {
    return {
      id: payload.id || payload.noticeId || '',
      noticeId: payload.noticeId || payload.id || '',
      title: payload.title || '',
      author: payload.author || '',
      noticeDate: payload.noticeDate || '',
      content: payload.content || ''
    };
  }
  return {};
}


function normalizePortalSaveQueueSourceP474_(methodName, payload) {
  payload = payload || {};
  const raw = String(payload.clientSaveSource || payload.source || payload.saveSource || '').trim();
  if (raw) return raw;
  methodName = String(methodName || '').trim();
  if (methodName === 'saveCustomerMemoFast') return 'customer.expandedMemo';
  if (methodName === 'saveCustomerDetailFast' || methodName === 'saveCustomerPatchFastP473') return 'customer.detailPatch';
  if (methodName === 'savePortalSupportProcessThinP474') return 'support.process';
  if (methodName === 'savePortalSupportRequesterEditP527') return 'support.requesterEdit';
  if (methodName === 'savePortalSupportRequest') return 'support.request';
  if (methodName === 'savePortalNotice') return 'notice.save';
  if (methodName === 'updatePortalNotice') return 'notice.update';
  return 'unknown.' + (methodName || 'save');
}

function buildPortalSaveConflictInfoForResponseP474_(methodName, payload, err) {
  if (typeof getPortalDirectSaveConflictInfoP474_ === 'function') {
    const direct = getPortalDirectSaveConflictInfoP474_(err);
    if (direct) return direct;
  }
  return null;
}

function stringifyPortalSaveQueueJsonP473_(value) {
  try { return JSON.stringify(value == null ? null : value); } catch (e) { return String(value || ''); }
}

function parsePortalSaveQueueJsonP473_(text, fallback) {
  try {
    const raw = String(text || '').trim();
    if (!raw) return fallback;
    return JSON.parse(raw);
  } catch (e) { return fallback; }
}

function isPortalServerStaleErrorP473_(err) {
  if (typeof isPortalFieldConflictErrorP474_ === 'function' && isPortalFieldConflictErrorP474_(err)) return true;
  const code = String(err && err.code || '').trim();
  const msg = String(err && err.message || err || '');
  return code === 'PORTAL_STALE_CUSTOMER_VERSION' || code === 'PORTAL_FIELD_CONFLICT_P474' || msg.indexOf('다른 사용자가') >= 0 || msg.indexOf('같은 항목을 먼저 수정') >= 0;
}

function isPortalServerTransientWriteErrorP473_(err) {
  if (!err) return false;
  if (isPortalServerStaleErrorP473_(err)) return false;
  const msg = String(err && err.message || err || '');
  return msg.indexOf('다른 작업 처리 중') >= 0 ||
    msg.indexOf('잠시 후') >= 0 ||
    msg.indexOf('Lock') >= 0 ||
    msg.indexOf('Service invoked too many times') >= 0 ||
    msg.indexOf('Exceeded maximum execution time') >= 0 ||
    msg.indexOf('Exception:') >= 0 && msg.indexOf('Google') >= 0 ||
    msg.indexOf('Internal error') >= 0 ||
    msg.indexOf('마스터시트 저장 확인 실패') >= 0 ||
    msg.indexOf('저장 확인 실패') >= 0 ||
    msg.indexOf('서버') >= 0;
}

function enqueueSaveFallbackP473_(methodName, payload, reason, err, statusOverride) {
  methodName = String(methodName || '').trim() || 'saveCustomerDetailFast';
  payload = Object.assign({}, payload || {});
  payload.clientOperationId = getPortalSaveOperationIdP473_(methodName, payload);
  const operationId = payload.clientOperationId;
  const sheet = getPortalSaveQueueSheetP473_();
  const idx = getPortalSaveQueueHeaderIndexP473_(sheet);
  const lastRow = sheet.getLastRow();
  const statusDone = PORTAL_SAVE_QUEUE_P473.STATUS.DONE;
  const status = statusOverride || (isPortalServerStaleErrorP473_(err) ? PORTAL_SAVE_QUEUE_P473.STATUS.CONFLICT : PORTAL_SAVE_QUEUE_P473.STATUS.QUEUED);
  const user = getPortalActiveUserEmailP473_();
  const patch = getPortalSavePatchP473_(methodName, payload);
  const expected = payload.expectedValues || (Object.prototype.hasOwnProperty.call(payload, 'expectedMemo') ? { memo: payload.expectedMemo } : {});
  const now = new Date();
  const errText = String(err && (err.stack || err.message) || reason || '').slice(0, 3000);

  if (lastRow >= 2 && idx['작업ID']) {
    const opValues = sheet.getRange(2, idx['작업ID'], lastRow - 1, 1).getDisplayValues();
    for (let i = 0; i < opValues.length; i++) {
      if (String(opValues[i][0] || '').trim() === operationId) {
        const rowNo = i + 2;
        const currentStatus = idx['상태'] ? String(sheet.getRange(rowNo, idx['상태']).getDisplayValue() || '').trim() : '';
        if (currentStatus === statusDone && idx['resultJson']) {
          const saved = parsePortalSaveQueueJsonP473_(sheet.getRange(rowNo, idx['resultJson']).getDisplayValue(), null);
          if (saved && typeof saved === 'object') {
            saved.ok = true;
            saved.fromSaveQueueP473 = true;
            saved.duplicateOperation = true;
            return saved;
          }
        }
        if (idx['수정일시']) sheet.getRange(rowNo, idx['수정일시']).setValue(now);
        if (idx['상태'] && currentStatus !== statusDone) sheet.getRange(rowNo, idx['상태']).setValue(status);
        if (idx['마지막오류'] && errText) sheet.getRange(rowNo, idx['마지막오류']).setValue(errText);
        ensureSaveQueueTriggerP473_();
        return buildQueuedSaveResponseP473_(payload, methodName, status, reason || errText, rowNo, err);
      }
    }
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(function(v) { return String(v || '').trim(); });
  const rowObj = {
    '등록일시': now,
    '수정일시': now,
    '작업ID': operationId,
    '사용자': user,
    '세션ID': String(payload.sessionId || payload.clientSessionId || ''),
    '고객번호': String(payload.customerNo || ''),
    'rowNo': Number(payload.rowNo || 0) || '',
    'methodName': methodName,
    'source': normalizePortalSaveQueueSourceP474_(methodName, payload),
    '상태': status,
    '우선순위': Number(payload.priority || 5) || 5,
    '시도횟수': 0,
    'patchJson': stringifyPortalSaveQueueJsonP473_(patch),
    'expectedValuesJson': stringifyPortalSaveQueueJsonP473_(expected),
    'payloadJson': stringifyPortalSaveQueueJsonP473_(payload),
    'resultJson': '',
    '마지막오류': errText,
    '적용일시': ''
  };
  sheet.appendRow(headers.map(function(h) { return Object.prototype.hasOwnProperty.call(rowObj, h) ? rowObj[h] : ''; }));
  ensureSaveQueueTriggerP473_();
  return buildQueuedSaveResponseP473_(payload, methodName, status, reason || errText, sheet.getLastRow(), err);
}

function buildQueuedSaveResponseP473_(payload, methodName, status, message, queueRowNo, err) {
  const patch = getPortalSavePatchP473_(methodName, payload);
  const isConflict = status === PORTAL_SAVE_QUEUE_P473.STATUS.CONFLICT;
  const conflictInfoP474 = isConflict ? buildPortalSaveConflictInfoForResponseP474_(methodName, payload, err) : null;
  return {
    ok: true,
    queuedFallbackP473: !isConflict,
    conflictP473: isConflict,
    conflictP474: isConflict,
    conflictInfoP474: conflictInfoP474,
    applied: false,
    rowNo: Number(payload && payload.rowNo || 0) || 0,
    customerNo: String(payload && payload.customerNo || ''),
    clientOperationId: String(payload && (payload.clientOperationId || payload.operationId) || ''),
    queueRowNo: queueRowNo || '',
    status: status,
    values: patch,
    changedValues: patch,
    changedKeys: Object.keys(patch || {}),
    savedAt: new Date().toISOString(),
    noSynchronousRefresh: true,
    message: isConflict ? '저장 충돌 확인 필요' : '저장 대기 중: 자동으로 다시 반영됩니다.'
  };
}

function ensureSaveQueueTriggerP473_() {
  try {
    const handler = PORTAL_SAVE_QUEUE_P473.TRIGGER_HANDLER;
    const exists = ScriptApp.getProjectTriggers().some(function(t) {
      return t && t.getHandlerFunction && t.getHandlerFunction() === handler;
    });
    if (!exists) ScriptApp.newTrigger(handler).timeBased().everyMinutes(1).create();
  } catch (err) {
    Logger.log('P473 저장큐 트리거 설치 실패: ' + (err && err.stack || err));
  }
}

function processSaveQueueP473(options) {
  options = options || {};
  const sheet = getPortalSaveQueueSheetP473_();
  const idx = getPortalSaveQueueHeaderIndexP473_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, processed: 0, message: '저장큐 비어 있음' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(Number(options.lockWaitMs || 1500))) {
    return { ok: false, processed: 0, message: '저장큐 처리 lock busy' };
  }

  const now = new Date();
  let processed = 0;
  let done = 0;
  let retry = 0;
  let conflict = 0;
  let fail = 0;

  try {
    const width = sheet.getLastColumn();
    const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
    const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function(v) { return String(v || '').trim(); });
    const maxJobs = Math.max(1, Number(options.maxJobs || PORTAL_SAVE_QUEUE_P473.MAX_JOBS_PER_RUN) || 5);

    for (let i = 0; i < values.length && processed < maxJobs; i++) {
      const rowNo = i + 2;
      const row = values[i];
      const obj = {};
      headers.forEach(function(h, c) { if (h) obj[h] = row[c]; });
      let status = String(obj['상태'] || '').trim();
      const runningAt = obj['수정일시'];
      if (status === PORTAL_SAVE_QUEUE_P473.STATUS.RUNNING && runningAt instanceof Date) {
        const ageMin = (now.getTime() - runningAt.getTime()) / 60000;
        if (ageMin >= PORTAL_SAVE_QUEUE_P473.STALE_RUNNING_MINUTES) status = PORTAL_SAVE_QUEUE_P473.STATUS.RETRY;
      }
      const previousAttemptsP513 = Number(obj['시도횟수'] || 0) || 0;
      const conflictRebaseEligibleP513 = status === PORTAL_SAVE_QUEUE_P473.STATUS.CONFLICT && previousAttemptsP513 < (PORTAL_SAVE_QUEUE_P473.MAX_CONFLICT_REBASE_ATTEMPTS || 3);
      if ([PORTAL_SAVE_QUEUE_P473.STATUS.QUEUED, PORTAL_SAVE_QUEUE_P473.STATUS.RETRY].indexOf(status) < 0 && !conflictRebaseEligibleP513) continue;

      processed++;
      if (idx['상태']) sheet.getRange(rowNo, idx['상태']).setValue(PORTAL_SAVE_QUEUE_P473.STATUS.RUNNING);
      if (idx['수정일시']) sheet.getRange(rowNo, idx['수정일시']).setValue(new Date());
      SpreadsheetApp.flush();

      let attempt = Number(obj['시도횟수'] || 0) + 1;
      try {
        const payload = parsePortalSaveQueueJsonP473_(obj['payloadJson'], {});
        const methodName = String(obj['methodName'] || 'saveCustomerDetailFast');
        const result = applyPortalQueuedSaveJobP473_(methodName, payload);
        if (idx['상태']) sheet.getRange(rowNo, idx['상태']).setValue(PORTAL_SAVE_QUEUE_P473.STATUS.DONE);
        if (idx['수정일시']) sheet.getRange(rowNo, idx['수정일시']).setValue(new Date());
        if (idx['시도횟수']) sheet.getRange(rowNo, idx['시도횟수']).setValue(attempt);
        if (idx['resultJson']) sheet.getRange(rowNo, idx['resultJson']).setValue(stringifyPortalSaveQueueJsonP473_(result).slice(0, 45000));
        if (idx['마지막오류']) sheet.getRange(rowNo, idx['마지막오류']).setValue('');
        if (idx['적용일시']) sheet.getRange(rowNo, idx['적용일시']).setValue(new Date());
        done++;
      } catch (err) {
        const staleP513 = isPortalServerStaleErrorP473_(err);
        const transientP513 = !staleP513 && isPortalServerTransientWriteErrorP473_(err);
        const nextStatus = staleP513
          ? PORTAL_SAVE_QUEUE_P473.STATUS.CONFLICT
          : (transientP513 ? PORTAL_SAVE_QUEUE_P473.STATUS.RETRY : PORTAL_SAVE_QUEUE_P473.STATUS.FAIL);
        if (idx['상태']) sheet.getRange(rowNo, idx['상태']).setValue(nextStatus);
        if (idx['수정일시']) sheet.getRange(rowNo, idx['수정일시']).setValue(new Date());
        if (idx['시도횟수']) sheet.getRange(rowNo, idx['시도횟수']).setValue(attempt);
        if (idx['마지막오류']) sheet.getRange(rowNo, idx['마지막오류']).setValue(String(err && (err.stack || err.message) || err).slice(0, 45000));
        if (nextStatus === PORTAL_SAVE_QUEUE_P473.STATUS.CONFLICT) conflict++;
        else if (nextStatus === PORTAL_SAVE_QUEUE_P473.STATUS.RETRY) retry++;
        else fail++;
      }
    }
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
  return { ok: true, processed: processed, done: done, retry: retry, conflict: conflict, fail: fail };
}


function isPortalQueuedCustomerPatchAlreadyAppliedP489_(methodName, payload, patch) {
  methodName = String(methodName || '').trim();
  payload = Object.assign({}, payload || {});
  patch = Object.assign({}, patch || payload.values || {});
  if (methodName === 'saveCustomerMemoFast' && Object.prototype.hasOwnProperty.call(payload, 'memo')) patch = { memo: payload.memo };
  const keys = Object.keys(patch || {});
  if (!keys.length) return false;
  if (['saveCustomerPatchFastP473','saveCustomerDetailFast','saveCustomerMemoFast'].indexOf(methodName) < 0) return false;
  const target = assertCustomerTarget_(payload, '저장큐 충돌 자동정리 확인', { readObject: false });
  const sheet = target.sheet;
  const rowNo = target.rowNo;
  const headerMap = getHeaderMap_(sheet);
  const defMap = typeof getPortalThinSaveDetailDefMapP462_ === 'function' ? getPortalThinSaveDetailDefMapP462_() : {};
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    let col = 0;
    if (key === 'memo') col = findMasterFieldCol_(headerMap, 'memo') || 0;
    else {
      const def = defMap[key];
      if (!def) return false;
      col = findFirstExistingHeaderCol_(headerMap, def.headers || []) || 0;
    }
    if (!col) return false;
    const currentRaw = getPortalCurrentFieldCompareValueP489_(sheet, rowNo, key, col);
    const currentCompare = getPortalMasterConflictCompareTextP489_(key, currentRaw);
    const patchCompare = getPortalMasterConflictCompareTextP489_(key, patch[key]);
    if (currentCompare !== patchCompare) return false;
  }
  return true;
}

function resolveCustomerSupersededSaveConflictsP489_(customerNo, rowNoFilter) {
  const no = String(customerNo || '').trim();
  rowNoFilter = Number(rowNoFilter || 0) || 0;
  if (!no) return { ok: true, resolved: 0 };
  const sheet = getPortalSaveQueueSheetP473_();
  const idx = getPortalSaveQueueHeaderIndexP473_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !idx['고객번호'] || !idx['상태']) return { ok: true, resolved: 0 };
  const width = sheet.getLastColumn();
  const rows = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function(v) { return String(v || '').trim(); });
  let resolved = 0;
  for (let i = 0; i < rows.length; i++) {
    const rowNo = i + 2;
    const row = rows[i];
    const status = String(row[(idx['상태'] || 1) - 1] || '').trim();
    if (status !== PORTAL_SAVE_QUEUE_P473.STATUS.CONFLICT) continue;
    if (String(row[(idx['고객번호'] || 1) - 1] || '').trim() !== no) continue;
    const queuedRowNoP524 = idx['rowNo'] ? (Number(row[idx['rowNo'] - 1] || 0) || 0) : 0;
    if (rowNoFilter && queuedRowNoP524 && queuedRowNoP524 !== rowNoFilter) continue;
    const obj = {};
    headers.forEach(function(h, c) { if (h) obj[h] = row[c]; });
    const methodName = String(obj['methodName'] || 'saveCustomerDetailFast').trim();
    const payload = parsePortalSaveQueueJsonP473_(obj['payloadJson'], {});
    const patch = parsePortalSaveQueueJsonP473_(obj['patchJson'], getPortalSavePatchP473_(methodName, payload));
    try {
      if (!isPortalQueuedCustomerPatchAlreadyAppliedP489_(methodName, payload, patch)) continue;
      const now = new Date();
      if (idx['상태']) sheet.getRange(rowNo, idx['상태']).setValue(PORTAL_SAVE_QUEUE_P473.STATUS.DONE);
      if (idx['수정일시']) sheet.getRange(rowNo, idx['수정일시']).setValue(now);
      if (idx['resultJson']) sheet.getRange(rowNo, idx['resultJson']).setValue(stringifyPortalSaveQueueJsonP473_({ ok: true, resolvedBy: 'MASTER_ALREADY_MATCHES_PATCH_P489', patch: patch }).slice(0, 45000));
      if (idx['마지막오류']) sheet.getRange(rowNo, idx['마지막오류']).setValue('');
      if (idx['적용일시']) sheet.getRange(rowNo, idx['적용일시']).setValue(now);
      resolved++;
    } catch (err) {
      // 자동정리는 보조 기능입니다. 실패하면 기존 CONFLICT 상태를 보존합니다.
    }
  }
  return { ok: true, resolved: resolved };
}

function applyPortalQueuedSaveJobP473_(methodName, payload) {
  methodName = String(methodName || '').trim();
  payload = Object.assign({}, payload || {});
  payload.noSynchronousRefresh = true;
  payload.fastMode = true;
  if (methodName === 'saveCustomerMemoFast') {
    payload.thinSave = true;
    return saveCustomerMemoThinCoreP462_(payload);
  }
  if (methodName === 'saveCustomerPatchFastP473') {
    payload.thinSave = true;
    return applyCustomerPatchDirectP473_(payload);
  }
  if (methodName === 'saveCustomerDetailFast') {
    const values = payload.values || {};
    if (isPortalCustomerThinSavePayloadP462_(payload, values)) return applyCustomerPatchDirectP473_(payload);
    return saveCustomerDetailFastCoreP202_(payload);
  }
  if (methodName === 'savePortalSupportProcessThinP474') {
    return savePortalSupportProcessThinCoreP489_(payload);
  }
  if (methodName === 'savePortalSupportRequesterEditP527') {
    return savePortalSupportRequesterEditCoreP527_(payload);
  }
  if (methodName === 'savePortalSupportRequest') {
    return savePortalSupportRequestCoreP210_(payload);
  }
  if (methodName === 'savePortalNotice') {
    return savePortalNoticeCoreP489_(payload);
  }
  if (methodName === 'updatePortalNotice') {
    return updatePortalNoticeCoreP489_(payload);
  }
  throw new Error('지원하지 않는 저장큐 methodName: ' + methodName);
}


// P521: 자료발송 전 저장큐 차단 내역 상세 조회.
// - 발송 preflight에서 CONFLICT / FAIL / QUEUED / RETRY / RUNNING 상태를 팝업으로 보여주기 위한 읽기 전용 helper입니다.
// - 큐 상태를 변경하지 않고, 저장/발송 로직도 실행하지 않습니다.
function getCustomerSaveQueueBlockingItemsP521(customerNo, limit, options) {
  return getCustomerSaveQueueBlockingItemsP521_(customerNo, limit, options);
}

function trimPortalSaveQueueValueP521_(value, maxLen) {
  const text = String(value == null ? '' : value);
  const limit = Math.max(80, Number(maxLen || 1500) || 1500);
  return text.length > limit ? text.slice(0, limit) + '\n…' : text;
}

function stringifyPortalSaveQueuePreviewP521_(value, maxLen) {
  if (value == null || value === '') return '';
  try {
    if (typeof value === 'object') return trimPortalSaveQueueValueP521_(JSON.stringify(value), maxLen || 1500);
  } catch (e) {}
  return trimPortalSaveQueueValueP521_(value, maxLen || 1500);
}

function getPortalSaveQueueNestedObjectP521_(payload, key) {
  if (!payload || typeof payload !== 'object') return null;
  const obj = payload[key];
  return obj && typeof obj === 'object' ? obj : null;
}

function hasPortalSaveQueueOwnP521_(obj, key) {
  return !!(obj && typeof obj === 'object' && Object.prototype.hasOwnProperty.call(obj, key));
}

function buildPortalSaveQueueFieldSummariesP521_(patch, expected, payload, resultJson) {
  patch = patch && typeof patch === 'object' ? patch : {};
  expected = expected && typeof expected === 'object' ? expected : {};
  payload = payload && typeof payload === 'object' ? payload : {};
  resultJson = resultJson && typeof resultJson === 'object' ? resultJson : {};
  const payloadPatch = getPortalSaveQueueNestedObjectP521_(payload, 'patch') || getPortalSaveQueueNestedObjectP521_(payload, 'values') || {};
  const payloadExpected = getPortalSaveQueueNestedObjectP521_(payload, 'expectedValues') || {};
  const resultCurrent = getPortalSaveQueueNestedObjectP521_(resultJson, 'currentValues') || getPortalSaveQueueNestedObjectP521_(resultJson, 'serverValues') || {};
  const seen = {};
  const keys = [];
  [patch, expected, payloadPatch, payloadExpected, resultCurrent].forEach(function(obj) {
    Object.keys(obj || {}).forEach(function(k) {
      if (!seen[k]) {
        seen[k] = true;
        keys.push(k);
      }
    });
  });
  if (!keys.length && hasPortalSaveQueueOwnP521_(payload, 'memo')) keys.push('memo');
  if (!keys.length) return [];
  return keys.slice(0, 8).map(function(k) {
    let oldValue = hasPortalSaveQueueOwnP521_(expected, k) ? expected[k] : (hasPortalSaveQueueOwnP521_(payloadExpected, k) ? payloadExpected[k] : '');
    let newValue = hasPortalSaveQueueOwnP521_(patch, k) ? patch[k] : (hasPortalSaveQueueOwnP521_(payloadPatch, k) ? payloadPatch[k] : (k === 'memo' && hasPortalSaveQueueOwnP521_(payload, 'memo') ? payload.memo : ''));
    let serverValue = hasPortalSaveQueueOwnP521_(resultCurrent, k) ? resultCurrent[k] : '';
    if (!serverValue && String(resultJson.field || '') === String(k) && hasPortalSaveQueueOwnP521_(resultJson, 'serverValue')) serverValue = resultJson.serverValue;
    if (!serverValue && String(resultJson.field || '') === String(k) && hasPortalSaveQueueOwnP521_(resultJson, 'serverValuePreview')) serverValue = resultJson.serverValuePreview;
    return {
      field: String(k || ''),
      oldValue: trimPortalSaveQueueValueP521_(oldValue, 1800),
      newValue: trimPortalSaveQueueValueP521_(newValue, 1800),
      serverValue: trimPortalSaveQueueValueP521_(serverValue, 1800)
    };
  });
}


// P524: 자료발송 preflight 저장큐 오판 방지 helper.
// - 현재 고객/현재 row에 해당하는 active 상태만 발송 차단 대상으로 봅니다.
// - DONE/적용완료/비차단 상태, 다른 고객, 다른 row는 차단 대상에서 제외합니다.
function normalizePortalSaveQueueOptionsP524_(limit, options) {
  let opt = {};
  let maxLimit = limit;
  if (limit && typeof limit === 'object' && !Array.isArray(limit)) {
    opt = Object.assign({}, limit);
    maxLimit = opt.limit || opt.maxItems || 10;
  } else {
    opt = Object.assign({}, options || {});
  }
  opt.limit = Math.max(1, Math.min(Number(maxLimit || opt.limit || 10) || 10, 20));
  opt.customerNo = String(opt.customerNo || '').trim();
  opt.rowNo = Number(opt.rowNo || 0) || 0;
  opt.operationIds = Array.isArray(opt.operationIds) ? opt.operationIds.map(function(v) { return String(v || '').trim(); }).filter(Boolean) : [];
  opt.operationSet = {};
  opt.operationIds.forEach(function(id) { opt.operationSet[id] = true; });
  return opt;
}

function getPortalSaveQueueBlockingStatusSetP524_() {
  const set = {};
  [
    PORTAL_SAVE_QUEUE_P473.STATUS.QUEUED,
    PORTAL_SAVE_QUEUE_P473.STATUS.RUNNING,
    PORTAL_SAVE_QUEUE_P473.STATUS.RETRY,
    PORTAL_SAVE_QUEUE_P473.STATUS.CONFLICT,
    PORTAL_SAVE_QUEUE_P473.STATUS.FAIL
  ].forEach(function(st) { set[String(st || '').trim()] = true; });
  return set;
}

function parsePortalSaveQueueResultForP524_(row, idx) {
  if (!idx || !idx['resultJson']) return {};
  return parsePortalSaveQueueJsonP473_(row[idx['resultJson'] - 1], {});
}

function hasPortalSaveQueueAppliedMarkerP524_(row, idx, resultJson) {
  const appliedAt = idx && idx['적용일시'] ? String(row[idx['적용일시'] - 1] || '').trim() : '';
  if (appliedAt) return true;
  resultJson = resultJson && typeof resultJson === 'object' ? resultJson : {};
  return !!(
    resultJson.forcedApplyP522 ||
    resultJson.forcedApplyP524 ||
    resultJson.forceApplied ||
    resultJson.applied ||
    resultJson.status === PORTAL_SAVE_QUEUE_P473.STATUS.DONE
  );
}

function getPortalSaveQueueIgnoreReasonP524_(row, idx, filter, status, resultJson) {
  filter = filter || {};
  const blockingStatuses = filter.blockingStatuses || getPortalSaveQueueBlockingStatusSetP524_();
  const customerNo = String(filter.customerNo || '').trim();
  const rowCustomerNo = idx && idx['고객번호'] ? String(row[idx['고객번호'] - 1] || '').trim() : '';
  if (customerNo && rowCustomerNo !== customerNo) return 'differentCustomer';

  status = String(status || '').trim();
  if (!blockingStatuses[status]) return 'nonBlockingStatus';

  const rowNoFilter = Number(filter.rowNo || 0) || 0;
  const queuedRowNo = idx && idx['rowNo'] ? (Number(row[idx['rowNo'] - 1] || 0) || 0) : 0;
  if (rowNoFilter && queuedRowNo && queuedRowNo !== rowNoFilter) return 'rowMismatch';

  const opCol = idx && idx['작업ID'] ? idx['작업ID'] - 1 : -1;
  const operationId = opCol >= 0 ? String(row[opCol] || '').trim() : '';
  if (filter.operationSet && Object.keys(filter.operationSet).length && !filter.operationSet[operationId]) return 'operationNotSelected';

  if (hasPortalSaveQueueAppliedMarkerP524_(row, idx, resultJson)) return 'alreadyApplied';
  return '';
}

function addPortalSaveQueueScanIgnoredP524_(scanStats, reason) {
  reason = String(reason || 'unknown');
  scanStats.ignored++;
  scanStats.ignoredByReason[reason] = (scanStats.ignoredByReason[reason] || 0) + 1;
}

function getCustomerSaveQueueBlockingItemsP521_(customerNo, limit, options) {
  const no = String(customerNo || '').trim();
  const opt = normalizePortalSaveQueueOptionsP524_(limit, options);
  opt.customerNo = no;
  const max = opt.limit;
  const blockingStatuses = getPortalSaveQueueBlockingStatusSetP524_();
  opt.blockingStatuses = blockingStatuses;
  const out = {
    ok: true,
    customerNo: no,
    rowNo: opt.rowNo || '',
    totalBlocking: 0,
    items: [],
    scanStats: {
      scanned: 0,
      matchedCustomer: 0,
      matchedRow: 0,
      ignored: 0,
      ignoredByReason: {},
      blocking: 0
    }
  };
  if (!no) return out;
  const sheet = getPortalSaveQueueSheetP473_();
  const idx = getPortalSaveQueueHeaderIndexP473_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !idx['고객번호'] || !idx['상태']) return out;
  const width = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, lastRow - 1, width).getDisplayValues();
  const customerCol = idx['고객번호'] - 1;
  const statusCol = idx['상태'] - 1;
  const rowNoCol = idx['rowNo'] ? idx['rowNo'] - 1 : -1;

  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    out.scanStats.scanned++;
    const rowCustomerNo = String(row[customerCol] || '').trim();
    const status = String(row[statusCol] || '').trim();
    const resultJson = parsePortalSaveQueueResultForP524_(row, idx);
    const ignoreReason = getPortalSaveQueueIgnoreReasonP524_(row, idx, opt, status, resultJson);
    if (rowCustomerNo === no) {
      out.scanStats.matchedCustomer++;
      const queuedRowNoForStats = rowNoCol >= 0 ? (Number(row[rowNoCol] || 0) || 0) : 0;
      if (!opt.rowNo || !queuedRowNoForStats || queuedRowNoForStats === opt.rowNo) out.scanStats.matchedRow++;
    }
    if (ignoreReason) {
      if (rowCustomerNo === no || ignoreReason !== 'differentCustomer') addPortalSaveQueueScanIgnoredP524_(out.scanStats, ignoreReason);
      continue;
    }

    out.totalBlocking++;
    out.scanStats.blocking++;
    if (out.items.length >= max) continue;

    const patch = idx['patchJson'] ? parsePortalSaveQueueJsonP473_(row[idx['patchJson'] - 1], {}) : {};
    const expected = idx['expectedValuesJson'] ? parsePortalSaveQueueJsonP473_(row[idx['expectedValuesJson'] - 1], {}) : {};
    const payload = idx['payloadJson'] ? parsePortalSaveQueueJsonP473_(row[idx['payloadJson'] - 1], {}) : {};
    const lastError = idx['마지막오류'] ? String(row[idx['마지막오류'] - 1] || '') : '';

    out.items.push({
      sheetRow: i + 2,
      registeredAt: idx['등록일시'] ? row[idx['등록일시'] - 1] : '',
      updatedAt: idx['수정일시'] ? row[idx['수정일시'] - 1] : '',
      appliedAt: idx['적용일시'] ? row[idx['적용일시'] - 1] : '',
      operationId: idx['작업ID'] ? String(row[idx['작업ID'] - 1] || '') : '',
      user: idx['사용자'] ? String(row[idx['사용자'] - 1] || '') : '',
      sessionId: idx['세션ID'] ? String(row[idx['세션ID'] - 1] || '') : '',
      customerNo: no,
      rowNo: idx['rowNo'] ? String(row[idx['rowNo'] - 1] || '') : '',
      methodName: idx['methodName'] ? String(row[idx['methodName'] - 1] || '') : '',
      source: idx['source'] ? String(row[idx['source'] - 1] || '') : '',
      status: status,
      priority: idx['우선순위'] ? String(row[idx['우선순위'] - 1] || '') : '',
      attempts: idx['시도횟수'] ? String(row[idx['시도횟수'] - 1] || '') : '',
      lastError: trimPortalSaveQueueValueP521_(lastError, 1200),
      resultPreview: stringifyPortalSaveQueuePreviewP521_(resultJson, 1200),
      fieldSummaries: buildPortalSaveQueueFieldSummariesP521_(patch, expected, payload, resultJson),
      patchPreview: stringifyPortalSaveQueuePreviewP521_(patch, 1200),
      expectedPreview: stringifyPortalSaveQueuePreviewP521_(expected, 1200),
      payloadPreview: stringifyPortalSaveQueuePreviewP521_(payload, 1200)
    });
  }
  return out;
}

// P522: 자료발송 전 저장큐 차단 팝업에서 사용자가 승인한 경우,
// 저장큐에 남은 신규값을 현재 마스터시트에 강제 반영하고 큐 상태를 DONE으로 정리합니다.
// - expectedValues / expectedMemo / masterVersion 계열은 제거하여 현재 서버값 위에 저장합니다.
// - 사용자가 팝업에서 기존값/신규값을 확인하고 '저장하기'를 누른 경우에만 호출됩니다.
function forceApplyCustomerSaveQueueBlockingItemsP522(customerNo, options) {
  options = options || {};
  const no = String(customerNo || options.customerNo || '').trim();
  const rowNoFilter = Number(options.rowNo || 0) || 0;
  const operationIds = Array.isArray(options.operationIds) ? options.operationIds.map(function(v) { return String(v || '').trim(); }).filter(Boolean) : [];
  const operationSet = {};
  operationIds.forEach(function(id) { operationSet[id] = true; });
  if (!no) return { ok: false, applied: 0, failed: 0, message: '고객번호가 없어 저장큐 내역을 저장할 수 없습니다.' };

  const sheet = getPortalSaveQueueSheetP473_();
  const idx = getPortalSaveQueueHeaderIndexP473_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !idx['고객번호'] || !idx['상태']) return { ok: true, applied: 0, failed: 0, message: '저장큐에 처리할 내역이 없습니다.' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(Number(options.lockWaitMs || 1800) || 1800)) {
    return { ok: false, applied: 0, failed: 0, message: '다른 저장큐 작업이 처리 중입니다. 잠시 후 다시 시도해 주세요.' };
  }

  const blockingStatuses = {};
  [
    PORTAL_SAVE_QUEUE_P473.STATUS.QUEUED,
    PORTAL_SAVE_QUEUE_P473.STATUS.RUNNING,
    PORTAL_SAVE_QUEUE_P473.STATUS.RETRY,
    PORTAL_SAVE_QUEUE_P473.STATUS.CONFLICT,
    PORTAL_SAVE_QUEUE_P473.STATUS.FAIL
  ].forEach(function(st) { blockingStatuses[st] = true; });

  const out = { ok: true, customerNo: no, applied: 0, failed: 0, skipped: 0, results: [], message: '' };

  try {
    const width = sheet.getLastColumn();
    const rows = sheet.getRange(2, 1, lastRow - 1, width).getDisplayValues();
    const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function(v) { return String(v || '').trim(); });
    const customerCol = idx['고객번호'] - 1;
    const statusCol = idx['상태'] - 1;
    const rowNoCol = idx['rowNo'] ? idx['rowNo'] - 1 : -1;
    const opCol = idx['작업ID'] ? idx['작업ID'] - 1 : -1;
    const maxJobs = Math.max(1, Math.min(Number(options.maxJobs || 10) || 10, 20));

    for (let i = rows.length - 1; i >= 0 && (out.applied + out.failed) < maxJobs; i--) {
      const row = rows[i];
      const queueRowNo = i + 2;
      if (String(row[customerCol] || '').trim() !== no) continue;
      const status = String(row[statusCol] || '').trim();
      if (!blockingStatuses[status]) continue;
      const operationId = opCol >= 0 ? String(row[opCol] || '').trim() : '';
      if (operationIds.length && !operationSet[operationId]) continue;
      const queuedRowNo = rowNoCol >= 0 ? Number(row[rowNoCol] || 0) || 0 : 0;
      if (rowNoFilter && queuedRowNo && queuedRowNo !== rowNoFilter) continue;

      const obj = {};
      headers.forEach(function(h, c) { if (h) obj[h] = row[c]; });
      const methodName = String(obj['methodName'] || 'saveCustomerDetailFast').trim();
      const patch = parsePortalSaveQueueJsonP473_(obj['patchJson'], {});
      const rawPayload = parsePortalSaveQueueJsonP473_(obj['payloadJson'], {});
      const forcedPayload = buildPortalForceSavePayloadP522_(methodName, rawPayload, patch, no, queuedRowNo, operationId);
      const now = new Date();

      try {
        if (idx['상태']) sheet.getRange(queueRowNo, idx['상태']).setValue(PORTAL_SAVE_QUEUE_P473.STATUS.RUNNING);
        if (idx['수정일시']) sheet.getRange(queueRowNo, idx['수정일시']).setValue(now);
        SpreadsheetApp.flush();

        const result = applyPortalQueuedSaveJobP473_(methodName, forcedPayload);
        const finalResult = Object.assign({}, result || {}, {
          ok: true,
          forcedApplyP522: true,
          previousStatusP522: status,
          operationIdP522: operationId,
          forcedByP522: getPortalActiveUserEmailP473_(),
          forcedAtP522: new Date().toISOString()
        });
        if (idx['상태']) sheet.getRange(queueRowNo, idx['상태']).setValue(PORTAL_SAVE_QUEUE_P473.STATUS.DONE);
        if (idx['수정일시']) sheet.getRange(queueRowNo, idx['수정일시']).setValue(new Date());
        if (idx['시도횟수']) sheet.getRange(queueRowNo, idx['시도횟수']).setValue((Number(obj['시도횟수'] || 0) || 0) + 1);
        if (idx['resultJson']) sheet.getRange(queueRowNo, idx['resultJson']).setValue(stringifyPortalSaveQueueJsonP473_(finalResult).slice(0, 45000));
        if (idx['마지막오류']) sheet.getRange(queueRowNo, idx['마지막오류']).setValue('');
        if (idx['적용일시']) sheet.getRange(queueRowNo, idx['적용일시']).setValue(new Date());
        out.applied++;
        out.results.push({ ok: true, queueRowNo: queueRowNo, operationId: operationId, methodName: methodName, statusBefore: status });
      } catch (err) {
        const errText = String(err && (err.stack || err.message) || err).slice(0, 45000);
        if (idx['상태']) sheet.getRange(queueRowNo, idx['상태']).setValue(isPortalServerStaleErrorP473_(err) ? PORTAL_SAVE_QUEUE_P473.STATUS.CONFLICT : PORTAL_SAVE_QUEUE_P473.STATUS.FAIL);
        if (idx['수정일시']) sheet.getRange(queueRowNo, idx['수정일시']).setValue(new Date());
        if (idx['시도횟수']) sheet.getRange(queueRowNo, idx['시도횟수']).setValue((Number(obj['시도횟수'] || 0) || 0) + 1);
        if (idx['마지막오류']) sheet.getRange(queueRowNo, idx['마지막오류']).setValue(errText);
        out.failed++;
        out.results.push({ ok: false, queueRowNo: queueRowNo, operationId: operationId, methodName: methodName, statusBefore: status, error: errText.slice(0, 1000) });
      }
    }
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }

  out.ok = out.failed === 0;
  out.message = out.ok
    ? ('저장큐 내역 ' + out.applied + '건을 저장했습니다.')
    : ('저장큐 내역 저장 중 실패 ' + out.failed + '건이 있습니다.');
  if (!out.applied && !out.failed) out.message = '현재 고객에 저장할 차단 내역이 없습니다.';
  const afterBlockingP524 = getCustomerSaveQueueBlockingItemsP521_(no, 10, { rowNo: rowNoFilter });
  out.blockingItems = afterBlockingP524.items;
  out.scanStats = afterBlockingP524.scanStats;
  if (out.blockingItems && out.blockingItems.length) out.ok = false;
  return out;
}

function buildPortalForceSavePayloadP522_(methodName, payload, patch, customerNo, rowNo, operationId) {
  payload = Object.assign({}, payload || {});
  patch = patch && typeof patch === 'object' ? Object.assign({}, patch) : {};
  const method = String(methodName || '').trim();
  const forced = Object.assign({}, payload);
  forced.customerNo = String(forced.customerNo || customerNo || '').trim();
  forced.rowNo = Number(forced.rowNo || rowNo || 0) || rowNo || '';
  forced.clientSaveSource = String(forced.clientSaveSource || forced.source || 'saveQueue.forceApplyP522') + '.forceApplyP522';
  forced.source = forced.clientSaveSource;
  forced.forceApplyP522 = true;
  forced.noSynchronousRefresh = true;
  forced.fastMode = true;
  forced.thinSave = true;
  forced.clientOperationId = String(operationId || forced.clientOperationId || 'FORCE_P522') + '_FORCE_' + new Date().getTime();

  delete forced.expectedValues;
  delete forced.expectedMemo;
  delete forced.baseMasterVersion;
  delete forced.masterVersion;
  delete forced.__masterVersion;
  delete forced.expectedMasterVersion;

  if (Object.keys(patch).length) {
    forced.values = Object.assign({}, patch);
    forced.patch = Object.assign({}, patch);
    if (Object.prototype.hasOwnProperty.call(patch, 'memo')) forced.memo = patch.memo;
  } else if (forced.patch && typeof forced.patch === 'object') {
    forced.values = Object.assign({}, forced.patch || {});
    if (Object.prototype.hasOwnProperty.call(forced.patch, 'memo')) forced.memo = forced.patch.memo;
  } else if (forced.values && typeof forced.values === 'object') {
    forced.values = Object.assign({}, forced.values || {});
    if (Object.prototype.hasOwnProperty.call(forced.values, 'memo')) forced.memo = forced.values.memo;
  }

  if (method === 'saveCustomerMemoFast' && !Object.prototype.hasOwnProperty.call(forced, 'memo') && forced.values && Object.prototype.hasOwnProperty.call(forced.values, 'memo')) {
    forced.memo = forced.values.memo;
  }
  if (method === 'saveCustomerPatchFastP473' || method === 'saveCustomerDetailFast') {
    if (!forced.values || typeof forced.values !== 'object') forced.values = Object.assign({}, patch || {});
  }
  return forced;
}

function flushCustomerPendingOpsP473(customerNo, timeoutMs, rowNo) {
  try { resolveCustomerSupersededSaveConflictsP489_(customerNo, rowNo); } catch (e) {}
  const started = new Date().getTime();
  const limit = Math.max(1000, Number(timeoutMs || 5000) || 5000);
  let last = null;
  while (new Date().getTime() - started < limit) {
    const before = getCustomerSaveQueueStatusCountsP489_(customerNo, rowNo);
    if (before.fail) {
      return { ok: false, pending: before.pending, conflict: before.conflict, fail: before.fail, result: last, message: '저장 실패 작업이 남아 있어 진행할 수 없습니다.', blockingItems: getCustomerSaveQueueBlockingItemsP521_(customerNo, 10, { rowNo: rowNo }).items, scanStats: before.scanStats || null };
    }
    if (!before.pending && !before.conflict) return { ok: true, pending: 0, conflict: 0, fail: 0, result: last, scanStats: before.scanStats || null };
    last = processSaveQueueP473({ maxJobs: 8, lockWaitMs: 800 });
    const after = getCustomerSaveQueueStatusCountsP489_(customerNo, rowNo);
    if (after.fail) {
      return { ok: false, pending: after.pending, conflict: after.conflict, fail: after.fail, result: last, message: '저장 실패 작업이 남아 있어 진행할 수 없습니다.', blockingItems: getCustomerSaveQueueBlockingItemsP521_(customerNo, 10, { rowNo: rowNo }).items, scanStats: after.scanStats || null };
    }
    if (!after.pending && !after.conflict) return { ok: true, pending: 0, conflict: 0, fail: 0, result: last, scanStats: after.scanStats || null };
    if (after.conflict && !after.pending) {
      if (last && Number(last.processed || 0) > 0) {
        Utilities.sleep(300);
        continue;
      }
      return { ok: false, pending: 0, conflict: after.conflict, fail: 0, result: last, message: '자동 재기준 저장으로 해결되지 않은 충돌이 남아 있습니다.', blockingItems: getCustomerSaveQueueBlockingItemsP521_(customerNo, 10, { rowNo: rowNo }).items, scanStats: after.scanStats || null };
    }
    Utilities.sleep(300);
  }
  const finalCounts = getCustomerSaveQueueStatusCountsP489_(customerNo, rowNo);
  return { ok: false, pending: finalCounts.pending, conflict: finalCounts.conflict, fail: finalCounts.fail, result: last, message: '저장 대기 작업이 아직 남아 있습니다.', blockingItems: getCustomerSaveQueueBlockingItemsP521_(customerNo, 10, { rowNo: rowNo }).items, scanStats: finalCounts.scanStats || null };
}

function getCustomerSaveQueueStatusCountsP489_(customerNo, rowNo) {
  const no = String(customerNo || '').trim();
  const rowNoFilter = Number(rowNo || 0) || 0;
  const counts = {
    pending: 0,
    conflict: 0,
    fail: 0,
    done: 0,
    total: 0,
    ignored: 0,
    ignoredByReason: {},
    scanStats: {
      scanned: 0,
      matchedCustomer: 0,
      matchedRow: 0,
      ignored: 0,
      ignoredByReason: {},
      blocking: 0
    }
  };
  if (!no) return counts;
  const sheet = getPortalSaveQueueSheetP473_();
  const idx = getPortalSaveQueueHeaderIndexP473_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !idx['고객번호'] || !idx['상태']) return counts;
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
  const customerCol = idx['고객번호'] - 1;
  const statusCol = idx['상태'] - 1;
  const rowNoCol = idx['rowNo'] ? idx['rowNo'] - 1 : -1;
  const opt = {
    customerNo: no,
    rowNo: rowNoFilter,
    blockingStatuses: getPortalSaveQueueBlockingStatusSetP524_()
  };
  data.forEach(function(row) {
    counts.scanStats.scanned++;
    const rowCustomerNo = String(row[customerCol] || '').trim();
    if (rowCustomerNo !== no) return;
    counts.scanStats.matchedCustomer++;
    const queuedRowNo = rowNoCol >= 0 ? (Number(row[rowNoCol] || 0) || 0) : 0;
    if (!rowNoFilter || !queuedRowNo || queuedRowNo === rowNoFilter) counts.scanStats.matchedRow++;
    const st = String(row[statusCol] || '').trim();
    const resultJson = parsePortalSaveQueueResultForP524_(row, idx);
    const ignoreReason = getPortalSaveQueueIgnoreReasonP524_(row, idx, opt, st, resultJson);
    if (ignoreReason) {
      counts.ignored++;
      counts.ignoredByReason[ignoreReason] = (counts.ignoredByReason[ignoreReason] || 0) + 1;
      addPortalSaveQueueScanIgnoredP524_(counts.scanStats, ignoreReason);
      if (st === PORTAL_SAVE_QUEUE_P473.STATUS.DONE) counts.done++;
      return;
    }
    counts.total++;
    counts.scanStats.blocking++;
    if ([PORTAL_SAVE_QUEUE_P473.STATUS.QUEUED, PORTAL_SAVE_QUEUE_P473.STATUS.RETRY, PORTAL_SAVE_QUEUE_P473.STATUS.RUNNING].indexOf(st) >= 0) counts.pending++;
    else if (st === PORTAL_SAVE_QUEUE_P473.STATUS.CONFLICT) counts.conflict++;
    else if (st === PORTAL_SAVE_QUEUE_P473.STATUS.FAIL) counts.fail++;
  });
  return counts;
}

function getCustomerPendingSaveQueueCountP473_(customerNo, rowNo) {
  const no = String(customerNo || '').trim();
  const rowNoFilter = Number(rowNo || 0) || 0;
  if (!no) return 0;
  const sheet = getPortalSaveQueueSheetP473_();
  const idx = getPortalSaveQueueHeaderIndexP473_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !idx['고객번호'] || !idx['상태']) return 0;
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
  const customerCol = idx['고객번호'] - 1;
  const statusCol = idx['상태'] - 1;
  const rowNoCol = idx['rowNo'] ? idx['rowNo'] - 1 : -1;
  let count = 0;
  data.forEach(function(row) {
    if (String(row[customerCol] || '').trim() !== no) return;
    const queuedRowNo = rowNoCol >= 0 ? (Number(row[rowNoCol] || 0) || 0) : 0;
    if (rowNoFilter && queuedRowNo && queuedRowNo !== rowNoFilter) return;
    const st = String(row[statusCol] || '').trim();
    if ([PORTAL_SAVE_QUEUE_P473.STATUS.QUEUED, PORTAL_SAVE_QUEUE_P473.STATUS.RETRY, PORTAL_SAVE_QUEUE_P473.STATUS.RUNNING].indexOf(st) >= 0) count++;
  });
  return count;
}
