/***************************************
 * S1 Sales Portal - 22_ConcurrencyTestService.gs
 * STEP10: 관리자 전용 동시성 테스트 패널
 * - 기존 운영 기능을 변경하지 않고 테스트용 서버 API만 추가합니다.
 * - 테스트 데이터는 [CONCURRENCY_TEST:runId] 마커로 식별합니다.
 * - 실제 메일 발송은 이 서비스에서 호출하지 않습니다.
 ***************************************/

const PORTAL_CONCURRENCY_TEST_LOG_SHEET_NAME_P240 = '포털_동시성테스트로그';
const PORTAL_CONCURRENCY_TEST_STATE_SHEET_NAME_P240 = '포털_동시성테스트상태';
const PORTAL_CONCURRENCY_TEST_MAX_COUNT_P240 = 20;
const PORTAL_CONCURRENCY_TEST_LOG_HEADERS_P240 = [
  '기록일시',
  '실행ID',
  '시나리오',
  '순번',
  '결과',
  '소요ms',
  'rowNo',
  'receiptNo',
  'taskId',
  'customerNo',
  '메시지',
  '상세JSON'
];
const PORTAL_CONCURRENCY_TEST_STATE_HEADERS_P240 = [
  '기록일시',
  '실행ID',
  '구분',
  'rowNo',
  'customerNo',
  '원본값',
  '복구여부',
  '복구일시',
  '비고'
];

function getPortalConcurrencyTestInitP240() {
  const perm = assertPortalCanRunConcurrencyTestP240_();
  const recent = getPortalConcurrencyRecentLogsP240_(20);
  return {
    ok: true,
    now: new Date().toISOString(),
    user: getPortalCurrentUserName_(),
    permission: sanitizePortalPermissionForClient_(perm),
    maxCount: PORTAL_CONCURRENCY_TEST_MAX_COUNT_P240,
    scenarios: [
      { key: 'SUPPORT_CREATE', label: '영업지원요청 동시 신규등록', safe: true, cleanup: true },
      { key: 'TODAY_UPSERT', label: '오늘할일 동시 upsert 저장', safe: true, cleanup: true },
      { key: 'CUSTOMER_MEMO', label: '고객 메모 동시 저장', safe: false, cleanup: true }
    ],
    notes: [
      '실제 메일 발송은 호출하지 않습니다.',
      '영업지원/오늘할일 테스트 데이터에는 [CONCURRENCY_TEST:실행ID] 마커가 붙습니다.',
      '고객 메모 테스트는 입력한 rowNo의 실제 메모를 바꾼 뒤 cleanup으로 원복합니다.'
    ],
    recentLogs: recent
  };
}

function runPortalConcurrencyTestActionP240(payload) {
  payload = payload || {};
  assertPortalCanRunConcurrencyTestP240_();
  const scenario = String(payload.scenario || '').trim().toUpperCase();
  const runId = sanitizePortalConcurrencyRunIdP240_(payload.runId || makePortalConcurrencyRunIdP240_());
  const seq = Math.max(1, Math.min(9999, Number(payload.seq) || 1));
  const marker = makePortalConcurrencyMarkerP240_(runId);
  const started = Date.now();
  const jitterMs = Math.max(0, Math.min(1500, Number(payload.jitterMs) || 0));
  if (jitterMs) Utilities.sleep(Math.floor(Math.random() * jitterMs));

  let result = null;
  let ok = false;
  try {
    if (scenario === 'SUPPORT_CREATE') {
      result = runPortalConcurrencySupportCreateP240_(runId, seq, marker, payload);
    } else if (scenario === 'TODAY_UPSERT') {
      result = runPortalConcurrencyTodayUpsertP240_(runId, seq, marker, payload);
    } else if (scenario === 'CUSTOMER_MEMO') {
      result = runPortalConcurrencyCustomerMemoP240_(runId, seq, marker, payload);
    } else {
      throw new Error('지원하지 않는 동시성 테스트 시나리오입니다: ' + scenario);
    }
    ok = true;
    result = result || {};
    result.ok = true;
  } catch (err) {
    ok = false;
    result = {
      ok: false,
      error: String(err && err.message || err),
      errorCode: err && err.code || '',
      stack: String(err && err.stack || '').slice(0, 2000)
    };
  }

  const elapsedMs = Date.now() - started;
  result.runId = runId;
  result.scenario = scenario;
  result.seq = seq;
  result.elapsedMs = elapsedMs;
  result.marker = marker;
  appendPortalConcurrencyTestLogP240_({
    runId: runId,
    scenario: scenario,
    seq: seq,
    ok: ok,
    elapsedMs: elapsedMs,
    rowNo: result.rowNo || '',
    receiptNo: result.receiptNo || '',
    taskId: result.taskId || '',
    customerNo: result.customerNo || '',
    message: result.message || result.error || '',
    detail: result
  });
  return result;
}

function summarizePortalConcurrencyTestRunP240(payload) {
  payload = payload || {};
  assertPortalCanRunConcurrencyTestP240_();
  const scenario = String(payload.scenario || '').trim().toUpperCase();
  const runId = sanitizePortalConcurrencyRunIdP240_(payload.runId || '');
  const results = Array.isArray(payload.results) ? payload.results : [];
  const success = results.filter(function(r) { return !!(r && r.ok); }).length;
  const failed = results.length - success;
  const elapsedValues = results.map(function(r) { return Number(r && r.elapsedMs) || 0; }).filter(function(v) { return v >= 0; });
  const avgMs = elapsedValues.length ? Math.round(elapsedValues.reduce(function(a, b) { return a + b; }, 0) / elapsedValues.length) : 0;
  const maxMs = elapsedValues.length ? Math.max.apply(null, elapsedValues) : 0;
  const duplicateRowNos = findPortalConcurrencyDuplicatesP240_(results, 'rowNo');
  const duplicateReceipts = findPortalConcurrencyDuplicatesP240_(results, 'receiptNo');
  const duplicateTaskIds = findPortalConcurrencyDuplicatesP240_(results, 'taskId');
  const verification = verifyPortalConcurrencyScenarioP240_(scenario, runId, results);
  const pass = failed === 0 && duplicateRowNos.length === 0 && duplicateReceipts.length === 0 && duplicateTaskIds.length === 0 && verification.ok !== false;

  const summary = {
    ok: true,
    pass: pass,
    runId: runId,
    scenario: scenario,
    total: results.length,
    success: success,
    failed: failed,
    avgMs: avgMs,
    maxMs: maxMs,
    duplicateRowNos: duplicateRowNos,
    duplicateReceipts: duplicateReceipts,
    duplicateTaskIds: duplicateTaskIds,
    verification: verification,
    completedAt: new Date().toISOString()
  };
  appendPortalConcurrencyTestLogP240_({
    runId: runId,
    scenario: scenario,
    seq: 'SUMMARY',
    ok: pass,
    elapsedMs: maxMs,
    message: pass ? 'PASS' : 'CHECK_REQUIRED',
    detail: summary
  });
  return summary;
}

function cleanupPortalConcurrencyTestRunP240(payload) {
  payload = payload || {};
  assertPortalCanRunConcurrencyTestP240_();
  const runId = sanitizePortalConcurrencyRunIdP240_(payload.runId || '');
  if (!runId) throw new Error('정리할 실행ID가 없습니다.');
  const marker = makePortalConcurrencyMarkerP240_(runId);
  const support = cleanupPortalConcurrencySupportRowsP240_(marker);
  const today = cleanupPortalConcurrencyTodayRowsP240_(marker);
  const customer = payload.restoreCustomer === false ? { restored: 0, skipped: true } : restorePortalConcurrencyCustomerMemosP240_(runId);
  const result = { ok: true, runId: runId, marker: marker, support: support, today: today, customer: customer };
  appendPortalConcurrencyTestLogP240_({
    runId: runId,
    scenario: 'CLEANUP',
    seq: '-',
    ok: true,
    message: '테스트 데이터 정리 완료',
    detail: result
  });
  return result;
}

function runPortalConcurrencySupportCreateP240_(runId, seq, marker, payload) {
  const clientRequestId = 'CONC-' + runId + '-SUPPORT-' + seq;
  const res = savePortalSupportRequest({
    requestType: '기타요청',
    requester: '동시성테스트',
    customerNo: '',
    customerName: 'TEST_동시성검증_' + runId,
    requestText: marker + ' 영업지원요청 동시등록 seq=' + seq,
    status: '접수',
    processContent: '',
    clientRequestId: clientRequestId,
    masterMemo: '동시성 테스트 데이터입니다. cleanup으로 삭제하세요.'
  });
  return {
    rowNo: res && res.rowNo || '',
    receiptNo: res && res.receiptNo || '',
    duplicate: !!(res && res.duplicate),
    message: res && res.message || '영업지원요청 저장 완료'
  };
}

function runPortalConcurrencyTodayUpsertP240_(runId, seq, marker, payload) {
  const dateText = normalizePortalTodoDate_(payload.date || new Date());
  const taskId = 'CONC_TODAY_' + runId + '_' + seq;
  const res = savePortalTodosForDate({
    date: dateText,
    baseDateVersion: String(payload.baseDateVersion || 'STALE_BASE_' + runId),
    baseTaskIds: [],
    tasks: [{
      id: taskId,
      content: marker + ' 오늘할일 동시저장 seq=' + seq,
      done: false,
      tags: ['동시성테스트'],
      category: '할일',
      sourceType: 'MANUAL',
      priority: '보통',
      detail: '동시성 테스트 데이터입니다. cleanup으로 삭제 처리하세요.'
    }]
  });
  return {
    taskId: taskId,
    selectedDate: dateText,
    rowNo: '',
    saveMeta: res && res.saveMeta || null,
    mergedDueToStaleBase: !!(res && res.mergedDueToStaleBase),
    message: '오늘할일 저장 완료'
  };
}

function runPortalConcurrencyCustomerMemoP240_(runId, seq, marker, payload) {
  const rowNo = Number(payload.customerRowNo || payload.rowNo) || 0;
  const customerNo = String(payload.customerNo || '').trim();
  if (!rowNo || rowNo < PORTAL_CONFIG.DATA_START_ROW) {
    throw new Error('고객 메모 테스트는 테스트용 고객 rowNo를 입력해야 실행됩니다. 실제 고객 메모가 바뀔 수 있으므로 TEST 고객 행만 사용하세요.');
  }
  const original = capturePortalConcurrencyCustomerMemoOriginalP240_(runId, rowNo, customerNo);
  const memoText = marker + ' 고객메모 동시저장 seq=' + seq + '\n실행시각=' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const res = saveCustomerMemoFast({
    rowNo: rowNo,
    customerNo: customerNo || original.customerNo || '',
    memo: memoText
  });
  return {
    rowNo: res && res.rowNo || rowNo,
    customerNo: res && res.customerNo || customerNo || original.customerNo || '',
    masterVersion: res && res.masterVersion || '',
    message: res && res.message || '고객 메모 저장 완료'
  };
}

function capturePortalConcurrencyCustomerMemoOriginalP240_(runId, rowNo, customerNo) {
  const stateSheet = ensurePortalConcurrencyTestStateSheetP240_();
  const lock = LockService.getScriptLock();
  lock.waitLock(8000);
  try {
    const lastRow = stateSheet.getLastRow();
    if (lastRow >= 2) {
      const values = stateSheet.getRange(2, 1, lastRow - 1, PORTAL_CONCURRENCY_TEST_STATE_HEADERS_P240.length).getValues();
      for (let i = 0; i < values.length; i++) {
        if (String(values[i][1] || '') === runId && String(values[i][2] || '') === 'CUSTOMER_MEMO_ORIGINAL' && Number(values[i][3]) === rowNo) {
          return { rowNo: rowNo, customerNo: String(values[i][4] || customerNo || ''), originalMemo: String(values[i][5] || '') };
        }
      }
    }
    const master = getMasterSpreadsheet_();
    const sheet = master.getSheetByName(PORTAL_CONFIG.MASTER_SHEET_NAME);
    if (!sheet) throw new Error('마스터시트를 찾지 못했습니다.');
    const headerMap = getHeaderMap_(sheet);
    const memoCol = findMasterFieldCol_(headerMap, 'memo') || findFirstExistingHeaderCol_(headerMap, PORTAL_CONFIG.MEMO_HEADER_CANDIDATES || ['메모']);
    const customerNoCol = findFirstExistingHeaderCol_(headerMap, ['고객번호']);
    if (!memoCol) throw new Error('메모 컬럼을 찾지 못했습니다.');
    const originalMemo = String(sheet.getRange(rowNo, memoCol).getDisplayValue() || '');
    const resolvedCustomerNo = customerNo || (customerNoCol ? String(sheet.getRange(rowNo, customerNoCol).getDisplayValue() || '').trim() : '');
    stateSheet.appendRow([new Date(), runId, 'CUSTOMER_MEMO_ORIGINAL', rowNo, resolvedCustomerNo, originalMemo, '', '', 'cleanup 시 복구 대상']);
    return { rowNo: rowNo, customerNo: resolvedCustomerNo, originalMemo: originalMemo };
  } finally {
    try { lock.releaseLock(); } catch (err) {}
  }
}

function restorePortalConcurrencyCustomerMemosP240_(runId) {
  const stateSheet = ensurePortalConcurrencyTestStateSheetP240_();
  const lastRow = stateSheet.getLastRow();
  if (lastRow < 2) return { restored: 0, skipped: 0 };
  const width = PORTAL_CONCURRENCY_TEST_STATE_HEADERS_P240.length;
  const values = stateSheet.getRange(2, 1, lastRow - 1, width).getValues();
  let restored = 0;
  let skipped = 0;
  values.forEach(function(row, i) {
    const rowNo = Number(row[3]) || 0;
    const already = String(row[6] || '').toUpperCase() === 'Y';
    if (String(row[1] || '') !== runId || String(row[2] || '') !== 'CUSTOMER_MEMO_ORIGINAL' || !rowNo || already) {
      skipped++;
      return;
    }
    const customerNo = String(row[4] || '').trim();
    const originalMemo = String(row[5] || '');
    try {
      saveCustomerMemoFast({ rowNo: rowNo, customerNo: customerNo, memo: originalMemo });
      const sheetRow = i + 2;
      stateSheet.getRange(sheetRow, 7, 1, 2).setValues([['Y', new Date()]]);
      restored++;
    } catch (err) {
      stateSheet.getRange(i + 2, 9).setValue('복구 실패: ' + String(err && err.message || err));
      skipped++;
    }
  });
  return { restored: restored, skipped: skipped };
}

function cleanupPortalConcurrencySupportRowsP240_(marker) {
  const masterSs = getMasterSpreadsheet_();
  const sheet = getPortalSupportSheetFastV64_(masterSs);
  const headerMap = getPortalSupportHeaderMap_(sheet);
  const requestCol = headerMap['요청업무'];
  if (!requestCol) return { deleted: 0, reason: '요청업무 컬럼 없음' };
  const lastRow = sheet.getLastRow();
  if (lastRow < PORTAL_CONFIG.SUPPORT_DATA_START_ROW) return { deleted: 0 };
  const rowCount = lastRow - PORTAL_CONFIG.SUPPORT_DATA_START_ROW + 1;
  const values = sheet.getRange(PORTAL_CONFIG.SUPPORT_DATA_START_ROW, requestCol, rowCount, 1).getDisplayValues();
  const rows = [];
  values.forEach(function(r, i) {
    if (String(r[0] || '').indexOf(marker) >= 0) rows.push(PORTAL_CONFIG.SUPPORT_DATA_START_ROW + i);
  });
  deleteRowsBottomUpP240_(sheet, rows);
  try { bumpPortalSupportCacheBustV64_(); } catch (err) {}
  return { deleted: rows.length };
}

function cleanupPortalConcurrencyTodayRowsP240_(marker) {
  const sheet = ensurePortalTodaySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { markedDeleted: 0 };
  const headers = PORTAL_CONFIG.TODAY_HEADERS;
  const contentCol = headers.indexOf('내용') + 1;
  const deleteCol = headers.indexOf('삭제여부') + 1;
  if (!contentCol || !deleteCol) return { markedDeleted: 0, reason: '오늘할일 헤더 없음' };
  const rowCount = lastRow - 1;
  const values = sheet.getRange(2, contentCol, rowCount, 1).getDisplayValues();
  const rows = [];
  values.forEach(function(r, i) {
    if (String(r[0] || '').indexOf(marker) >= 0) rows.push(i + 2);
  });
  rows.forEach(function(rowNo) {
    sheet.getRange(rowNo, deleteCol).setValue('Y');
  });
  return { markedDeleted: rows.length };
}

function verifyPortalConcurrencyScenarioP240_(scenario, runId, results) {
  const marker = makePortalConcurrencyMarkerP240_(runId);
  if (scenario === 'SUPPORT_CREATE') {
    const masterSs = getMasterSpreadsheet_();
    const sheet = getPortalSupportSheetFastV64_(masterSs);
    const headerMap = getPortalSupportHeaderMap_(sheet);
    const requestCol = headerMap['요청업무'];
    if (!requestCol || sheet.getLastRow() < PORTAL_CONFIG.SUPPORT_DATA_START_ROW) return { ok: false, activeRows: 0 };
    const values = sheet.getRange(PORTAL_CONFIG.SUPPORT_DATA_START_ROW, requestCol, sheet.getLastRow() - PORTAL_CONFIG.SUPPORT_DATA_START_ROW + 1, 1).getDisplayValues();
    const count = values.filter(function(r) { return String(r[0] || '').indexOf(marker) >= 0; }).length;
    return { ok: count >= results.filter(function(r) { return r && r.ok; }).length, activeRows: count };
  }
  if (scenario === 'TODAY_UPSERT') {
    const sheet = ensurePortalTodaySheet_();
    if (sheet.getLastRow() < 2) return { ok: false, activeRows: 0 };
    const headers = PORTAL_CONFIG.TODAY_HEADERS;
    const contentCol = headers.indexOf('내용') + 1;
    const deleteCol = headers.indexOf('삭제여부') + 1;
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(headers.length, sheet.getLastColumn())).getDisplayValues();
    const count = rows.filter(function(r) {
      return String(r[contentCol - 1] || '').indexOf(marker) >= 0 && String(r[deleteCol - 1] || '').toUpperCase() !== 'Y';
    }).length;
    return { ok: count >= results.filter(function(r) { return r && r.ok; }).length, activeRows: count };
  }
  if (scenario === 'CUSTOMER_MEMO') {
    return { ok: true, note: '고객 메모는 최종값 1개가 남는 last-write-wins 구조입니다. cleanup으로 원복하세요.' };
  }
  return { ok: true };
}

function appendPortalConcurrencyTestLogP240_(entry) {
  entry = entry || {};
  try {
    const sheet = ensurePortalConcurrencyTestLogSheetP240_();
    const detail = JSON.stringify(entry.detail || {});
    const row = [
      new Date(),
      entry.runId || '',
      entry.scenario || '',
      entry.seq == null ? '' : entry.seq,
      entry.ok ? 'OK' : 'FAIL',
      Number(entry.elapsedMs) || 0,
      entry.rowNo || '',
      entry.receiptNo || '',
      entry.taskId || '',
      entry.customerNo || '',
      String(entry.message || '').slice(0, 500),
      detail.length > 45000 ? detail.slice(0, 45000) + '...[truncated]' : detail
    ];
    const lock = LockService.getScriptLock();
    try { lock.waitLock(4000); } catch (lockErr) {}
    try { sheet.appendRow(row); } finally { try { lock.releaseLock(); } catch (err) {} }
  } catch (err) {
    Logger.log('동시성 테스트 로그 기록 실패: ' + (err && err.stack || err));
  }
}

function getPortalConcurrencyRecentLogsP240_(limit) {
  limit = Math.max(1, Math.min(100, Number(limit) || 20));
  const sheet = ensurePortalConcurrencyTestLogSheetP240_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const start = Math.max(2, lastRow - limit + 1);
  const values = sheet.getRange(start, 1, lastRow - start + 1, PORTAL_CONCURRENCY_TEST_LOG_HEADERS_P240.length).getDisplayValues();
  return values.reverse().map(function(r) {
    return {
      at: r[0],
      runId: r[1],
      scenario: r[2],
      seq: r[3],
      result: r[4],
      elapsedMs: r[5],
      message: r[10]
    };
  });
}

function ensurePortalConcurrencyTestLogSheetP240_() {
  const ss = getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(PORTAL_CONCURRENCY_TEST_LOG_SHEET_NAME_P240);
  if (!sheet) sheet = ss.insertSheet(PORTAL_CONCURRENCY_TEST_LOG_SHEET_NAME_P240);
  ensureSheetHeaders_(sheet, PORTAL_CONCURRENCY_TEST_LOG_HEADERS_P240);
  try { sheet.setFrozenRows(1); sheet.getRange(1, 1, 1, PORTAL_CONCURRENCY_TEST_LOG_HEADERS_P240.length).setFontWeight('bold').setBackground('#f2f4f7'); } catch (err) {}
  return sheet;
}

function ensurePortalConcurrencyTestStateSheetP240_() {
  const ss = getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(PORTAL_CONCURRENCY_TEST_STATE_SHEET_NAME_P240);
  if (!sheet) sheet = ss.insertSheet(PORTAL_CONCURRENCY_TEST_STATE_SHEET_NAME_P240);
  ensureSheetHeaders_(sheet, PORTAL_CONCURRENCY_TEST_STATE_HEADERS_P240);
  try { sheet.setFrozenRows(1); sheet.getRange(1, 1, 1, PORTAL_CONCURRENCY_TEST_STATE_HEADERS_P240.length).setFontWeight('bold').setBackground('#f2f4f7'); } catch (err) {}
  return sheet;
}

function assertPortalCanRunConcurrencyTestP240_() {
  const perm = getPortalCurrentPermission_();
  if (!perm || perm.active === false || !perm.canUseAdminHome) {
    throw new Error('동시성 테스트는 관리자/서무 권한에서만 실행할 수 있습니다.');
  }
  return perm;
}

function sanitizePortalConcurrencyRunIdP240_(value) {
  return String(value || '').trim().replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40);
}

function makePortalConcurrencyRunIdP240_() {
  return 'CT' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss') + '_' + Math.floor(Math.random() * 10000);
}

function makePortalConcurrencyMarkerP240_(runId) {
  runId = sanitizePortalConcurrencyRunIdP240_(runId || makePortalConcurrencyRunIdP240_());
  return '[CONCURRENCY_TEST:' + runId + ']';
}

function findPortalConcurrencyDuplicatesP240_(results, key) {
  const seen = {};
  const dup = {};
  (results || []).forEach(function(r) {
    const v = String(r && r[key] || '').trim();
    if (!v) return;
    if (seen[v]) dup[v] = true;
    seen[v] = true;
  });
  return Object.keys(dup);
}

function deleteRowsBottomUpP240_(sheet, rows) {
  rows = (rows || []).map(function(v) { return Number(v) || 0; }).filter(Boolean).sort(function(a, b) { return b - a; });
  rows.forEach(function(rowNo) {
    try { sheet.deleteRow(rowNo); } catch (err) {}
  });
}
