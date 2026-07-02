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
      if ([PORTAL_SAVE_QUEUE_P473.STATUS.QUEUED, PORTAL_SAVE_QUEUE_P473.STATUS.RETRY].indexOf(status) < 0) continue;

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
        const nextStatus = isPortalServerStaleErrorP473_(err)
          ? PORTAL_SAVE_QUEUE_P473.STATUS.CONFLICT
          : (isPortalServerTransientWriteErrorP473_(err) ? PORTAL_SAVE_QUEUE_P473.STATUS.RETRY : PORTAL_SAVE_QUEUE_P473.STATUS.FAIL);
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

function resolveCustomerSupersededSaveConflictsP489_(customerNo) {
  const no = String(customerNo || '').trim();
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

function flushCustomerPendingOpsP473(customerNo, timeoutMs) {
  try { resolveCustomerSupersededSaveConflictsP489_(customerNo); } catch (e) {}
  const started = new Date().getTime();
  const limit = Math.max(1000, Number(timeoutMs || 5000) || 5000);
  let last = null;
  while (new Date().getTime() - started < limit) {
    const before = getCustomerSaveQueueStatusCountsP489_(customerNo);
    if (before.conflict || before.fail) {
      return { ok: false, pending: before.pending, conflict: before.conflict, fail: before.fail, result: last, message: '저장 충돌/실패 작업이 남아 있어 진행할 수 없습니다.' };
    }
    if (!before.pending) return { ok: true, pending: 0, conflict: 0, fail: 0, result: last };
    last = processSaveQueueP473({ maxJobs: 8, lockWaitMs: 800 });
    const after = getCustomerSaveQueueStatusCountsP489_(customerNo);
    if (after.conflict || after.fail) {
      return { ok: false, pending: after.pending, conflict: after.conflict, fail: after.fail, result: last, message: '저장 충돌/실패 작업이 남아 있어 진행할 수 없습니다.' };
    }
    if (!after.pending) return { ok: true, pending: 0, conflict: 0, fail: 0, result: last };
    Utilities.sleep(300);
  }
  const finalCounts = getCustomerSaveQueueStatusCountsP489_(customerNo);
  return { ok: false, pending: finalCounts.pending, conflict: finalCounts.conflict, fail: finalCounts.fail, result: last, message: '저장 대기 작업이 아직 남아 있습니다.' };
}

function getCustomerSaveQueueStatusCountsP489_(customerNo) {
  const no = String(customerNo || '').trim();
  const counts = { pending: 0, conflict: 0, fail: 0, done: 0, total: 0 };
  if (!no) return counts;
  const sheet = getPortalSaveQueueSheetP473_();
  const idx = getPortalSaveQueueHeaderIndexP473_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !idx['고객번호'] || !idx['상태']) return counts;
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
  const customerCol = idx['고객번호'] - 1;
  const statusCol = idx['상태'] - 1;
  data.forEach(function(row) {
    if (String(row[customerCol] || '').trim() !== no) return;
    const st = String(row[statusCol] || '').trim();
    counts.total++;
    if ([PORTAL_SAVE_QUEUE_P473.STATUS.QUEUED, PORTAL_SAVE_QUEUE_P473.STATUS.RETRY, PORTAL_SAVE_QUEUE_P473.STATUS.RUNNING].indexOf(st) >= 0) counts.pending++;
    else if (st === PORTAL_SAVE_QUEUE_P473.STATUS.CONFLICT) counts.conflict++;
    else if (st === PORTAL_SAVE_QUEUE_P473.STATUS.FAIL) counts.fail++;
    else if (st === PORTAL_SAVE_QUEUE_P473.STATUS.DONE) counts.done++;
  });
  return counts;
}

function getCustomerPendingSaveQueueCountP473_(customerNo) {
  const no = String(customerNo || '').trim();
  if (!no) return 0;
  const sheet = getPortalSaveQueueSheetP473_();
  const idx = getPortalSaveQueueHeaderIndexP473_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !idx['고객번호'] || !idx['상태']) return 0;
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
  const customerCol = idx['고객번호'] - 1;
  const statusCol = idx['상태'] - 1;
  let count = 0;
  data.forEach(function(row) {
    if (String(row[customerCol] || '').trim() !== no) return;
    const st = String(row[statusCol] || '').trim();
    if ([PORTAL_SAVE_QUEUE_P473.STATUS.QUEUED, PORTAL_SAVE_QUEUE_P473.STATUS.RETRY, PORTAL_SAVE_QUEUE_P473.STATUS.RUNNING].indexOf(st) >= 0) count++;
  });
  return count;
}
