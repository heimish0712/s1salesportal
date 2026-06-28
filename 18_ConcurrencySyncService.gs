
/***************************************
 * S1 Sales Portal - 18_ConcurrencySyncService.gs
 * P2-1/P2-2: 동시수정 안정화 + 마스터시트 직접 수정 동기화
 ***************************************/

const PORTAL_MASTER_META_HEADERS_P201 = ['수정일시', '수정버전', '최종수정자'];
const PORTAL_MASTER_SYNC_HANDLER_P201 = 'onMasterSheetEditSyncP201';

function withPortalScriptLockP201_(label, callback, options) {
  options = options || {};
  const attempts = Math.max(1, Number(options.attempts) || 4);
  const waitMs = Math.max(100, Number(options.waitMs) || 700);
  const sleepBaseMs = Math.max(100, Number(options.sleepBaseMs) || 180);
  const lock = LockService.getScriptLock();
  let locked = false;
  let lastErr = null;
  for (let i = 0; i < attempts; i++) {
    try {
      locked = lock.tryLock(waitMs);
      if (locked) break;
    } catch (err) {
      lastErr = err;
    }
    Utilities.sleep(sleepBaseMs * (i + 1));
  }
  if (!locked) {
    const err = new Error('다른 작업 처리 중입니다. 잠시 후 자동으로 다시 시도해 주세요.');
    err.code = 'PORTAL_LOCK_BUSY';
    err.label = label || '';
    err.detail = lastErr && lastErr.message ? lastErr.message : '';
    throw err;
  }
  try {
    return callback();
  } finally {
    try { lock.releaseLock(); } catch (err) {}
  }
}

function getPortalNowTextP201_(date) {
  return Utilities.formatDate(date || new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function getPortalVersionTextP201_(date) {
  return Utilities.formatDate(date || new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmssSSS');
}

function markPortalMasterDataChangedP201_(detail) {
  const now = new Date();
  const version = getPortalVersionTextP201_(now);
  PropertiesService.getScriptProperties().setProperties({
    PORTAL_MASTER_DATA_VERSION: version,
    PORTAL_MASTER_DATA_CHANGED_AT: getPortalNowTextP201_(now),
    PORTAL_MASTER_DATA_CHANGED_DETAIL: String(detail || '').slice(0, 500)
  }, true);
  return version;
}

function getPortalClientDataVersionP201() {
  const props = PropertiesService.getScriptProperties();
  let indexMeta = {};
  try { indexMeta = getCustomerSearchIndexMeta(); } catch (err) { indexMeta = { error: err && err.message ? err.message : String(err) }; }
  return {
    ok: true,
    now: getPortalNowTextP201_(new Date()),
    masterVersion: props.getProperty('PORTAL_MASTER_DATA_VERSION') || '',
    masterChangedAt: props.getProperty('PORTAL_MASTER_DATA_CHANGED_AT') || '',
    masterChangedDetail: props.getProperty('PORTAL_MASTER_DATA_CHANGED_DETAIL') || '',
    customerIndexVersion: props.getProperty('CUSTOMER_SEARCH_INDEX_VERSION') || (indexMeta && indexMeta.version) || '',
    customerIndexBuiltAt: props.getProperty('CUSTOMER_SEARCH_INDEX_BUILT_AT') || (indexMeta && indexMeta.builtAt) || '',
    customerIndexDirty: props.getProperty('CUSTOMER_SEARCH_INDEX_DIRTY') || 'N',
    customerIndexDirtyReason: props.getProperty('CUSTOMER_SEARCH_INDEX_DIRTY_REASON') || '',
    indexMeta: indexMeta || {}
  };
}

function installMasterSheetEditSyncTriggerP201() {
  const ss = getMasterSpreadsheet_();
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t && t.getHandlerFunction && t.getHandlerFunction() === PORTAL_MASTER_SYNC_HANDLER_P201) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger(PORTAL_MASTER_SYNC_HANDLER_P201)
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  return {
    ok: true,
    message: '마스터시트 수정 감지 트리거를 설치했습니다.',
    spreadsheetId: ss.getId(),
    handler: PORTAL_MASTER_SYNC_HANDLER_P201
  };
}

function onMasterSheetEditSyncP201(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  if (!sheet || sheet.getName() !== PORTAL_CONFIG.MASTER_SHEET_NAME) return;

  const rowStart = range.getRow();
  const rowEnd = rowStart + range.getNumRows() - 1;
  if (rowEnd < PORTAL_CONFIG.DATA_START_ROW) return;

  const colStart = range.getColumn();
  const colEnd = colStart + range.getNumColumns() - 1;
  const headerMap = getHeaderMap_(sheet);
  const editedMetaOnly = PORTAL_MASTER_META_HEADERS_P201.every(function(h) { return false; });
  const metaCols = PORTAL_MASTER_META_HEADERS_P201.map(function(h) { return findFirstExistingHeaderCol_(headerMap, [h]); }).filter(Boolean);
  const editOnlyMetaCols = metaCols.length && colStart >= Math.min.apply(null, metaCols) && colEnd <= Math.max.apply(null, metaCols) &&
    Array.from({length: colEnd - colStart + 1}, function(_, i) { return colStart + i; }).every(function(c) { return metaCols.indexOf(c) >= 0; });
  if (editOnlyMetaCols) return;

  withPortalScriptLockP201_('master-onedit-sync', function() {
    const freshMap = getHeaderMap_(sheet);
    const updatedAtCol = ensureMasterColumn_(sheet, freshMap, '수정일시');
    const versionCol = ensureMasterColumn_(sheet, getHeaderMap_(sheet), '수정버전');
    const editorCol = ensureMasterColumn_(sheet, getHeaderMap_(sheet), '최종수정자');
    const now = new Date();
    const nowText = getPortalNowTextP201_(now);
    const versionBase = getPortalVersionTextP201_(now);
    let editor = '';
    try { editor = String(Session.getActiveUser().getEmail() || '').trim(); } catch (err) {}
    if (!editor) editor = 'sheet-edit';

    const candidateRows = [];
    for (let r = Math.max(PORTAL_CONFIG.DATA_START_ROW, rowStart); r <= rowEnd; r++) candidateRows.push(r);

    // P2-4: 빈 행에는 수정일시/수정버전/최종수정자를 찍지 않습니다.
    // 대량 행 추가, 빈 영역 붙여넣기, 필터/드롭다운 조작 시 빈 행 전체에 메타가 찍히는 문제 방지.
    const targetRows = filterMeaningfulMasterRowsForMetaP204_(sheet, candidateRows, getHeaderMap_(sheet));
    if (!targetRows.length) return false;

    const metaByRow = writePortalMasterMetaBatchP240_(sheet, targetRows, {
      updatedAtCol: updatedAtCol,
      versionCol: versionCol,
      editorCol: editorCol,
      nowText: nowText,
      versionBase: versionBase,
      editor: editor
    });

    const changedColumnNames = getPortalEditedMasterColumnNamesP209_(sheet, getHeaderMap_(sheet), colStart, colEnd);
    const customerNoCol = findFirstExistingHeaderCol_(getHeaderMap_(sheet), ['고객번호']);
    const customerNoByRow = readPortalMasterColumnValuesByRowsP240_(sheet, targetRows, customerNoCol);
    const queueItems = targetRows.map(function(r) {
      return {
        rowNo: r,
        customerNo: customerNoByRow[r] || '',
        changedColumns: changedColumnNames.join(', '),
        editor: editor,
        status: 'PENDING',
        masterVersion: metaByRow[r] && metaByRow[r].version ? metaByRow[r].version : ''
      };
    });
    SpreadsheetApp.flush();

    let queueIds = [];
    if (typeof appendPortalChangeQueueRowsP209_ === 'function') {
      try { queueIds = appendPortalChangeQueueRowsP209_(queueItems); }
      catch (queueErr) { markCustomerSearchIndexDirty_('CHANGE_QUEUE_APPEND_FAIL', queueErr && queueErr.message ? queueErr.message : String(queueErr)); }
    }

    const maxImmediate = 25;
    const processedQueueIds = [];
    targetRows.slice(0, maxImmediate).forEach(function(r, idx) {
      try {
        const res = updateCustomerSearchIndexRow_(r);
        if (queueIds[idx] && res && res.ok && typeof markPortalChangeQueueRowsDoneByIdsP209_ === 'function') processedQueueIds.push(queueIds[idx]);
      } catch (err) {
        markCustomerSearchIndexDirty_('MASTER_ONEDIT_ROW_REFRESH_FAIL', 'row ' + r + ': ' + (err && err.message ? err.message : String(err)));
      }
    });
    if (processedQueueIds.length && typeof markPortalChangeQueueRowsDoneByIdsP209_ === 'function') {
      try { markPortalChangeQueueRowsDoneByIdsP209_(processedQueueIds, 'DONE', ''); } catch (err) {}
    }
    if (targetRows.length > maxImmediate) {
      markCustomerSearchIndexDirty_('MASTER_ONEDIT_MANY_ROWS', targetRows.length + ' rows edited; 변경큐에 남겨 재처리 필요');
    }
    markPortalMasterDataChangedP201_('마스터시트 직접수정 rows=' + targetRows.join(','));
    return true;
  }, { attempts: 3, waitMs: 500, sleepBaseMs: 150 });
}



/**
 * STEP9/P240: 마스터시트 onEdit 메타 컬럼을 행별 setValue 반복 대신 batch setValues로 기록합니다.
 * - 수정일시/최종수정자는 같은 값이지만, 수정버전은 row별 고유값을 유지합니다.
 * - targetRows가 필터링으로 끊겨 있을 수 있으므로 연속 행 그룹 단위로 나누어 씁니다.
 * - 메타 컬럼이 서로 떨어져 있어도 중간 업무 컬럼을 덮어쓰지 않도록 컬럼별 batch write를 사용합니다.
 */
function writePortalMasterMetaBatchP240_(sheet, targetRows, options) {
  options = options || {};
  targetRows = normalizePortalRowNumbersP240_(targetRows);
  if (!sheet || !targetRows.length) return {};

  const updatedAtCol = Number(options.updatedAtCol) || 0;
  const versionCol = Number(options.versionCol) || 0;
  const editorCol = Number(options.editorCol) || 0;
  const nowText = String(options.nowText || '');
  const versionBase = String(options.versionBase || getPortalVersionTextP201_(new Date()));
  const editor = String(options.editor || 'sheet-edit');

  const metaByRow = {};
  targetRows.forEach(function(rowNo, idx) {
    metaByRow[rowNo] = {
      updatedAt: nowText,
      version: versionBase + '-' + rowNo + '-' + idx,
      editor: editor
    };
  });

  const groups = groupContiguousRowsP240_(targetRows);
  groups.forEach(function(group) {
    const rows = [];
    for (let r = group.start; r <= group.end; r++) rows.push(r);
    if (updatedAtCol) {
      sheet.getRange(group.start, updatedAtCol, rows.length, 1)
        .setValues(rows.map(function(r) { return [metaByRow[r].updatedAt]; }));
    }
    if (versionCol) {
      sheet.getRange(group.start, versionCol, rows.length, 1)
        .setValues(rows.map(function(r) { return [metaByRow[r].version]; }));
    }
    if (editorCol) {
      sheet.getRange(group.start, editorCol, rows.length, 1)
        .setValues(rows.map(function(r) { return [metaByRow[r].editor]; }));
    }
  });
  return metaByRow;
}

function normalizePortalRowNumbersP240_(rowNumbers) {
  const seen = {};
  return (rowNumbers || [])
    .map(function(v) { return Number(v) || 0; })
    .filter(function(v) { return v >= PORTAL_CONFIG.DATA_START_ROW; })
    .filter(function(v) {
      if (seen[v]) return false;
      seen[v] = true;
      return true;
    })
    .sort(function(a, b) { return a - b; });
}

function groupContiguousRowsP240_(rowNumbers) {
  rowNumbers = normalizePortalRowNumbersP240_(rowNumbers);
  if (!rowNumbers.length) return [];
  const groups = [];
  let start = rowNumbers[0];
  let prev = rowNumbers[0];
  for (let i = 1; i < rowNumbers.length; i++) {
    const row = rowNumbers[i];
    if (row === prev + 1) {
      prev = row;
      continue;
    }
    groups.push({ start: start, end: prev, count: prev - start + 1 });
    start = row;
    prev = row;
  }
  groups.push({ start: start, end: prev, count: prev - start + 1 });
  return groups;
}

function readPortalMasterColumnValuesByRowsP240_(sheet, targetRows, colNo) {
  targetRows = normalizePortalRowNumbersP240_(targetRows);
  colNo = Number(colNo) || 0;
  const out = {};
  if (!sheet || !targetRows.length || !colNo) return out;
  groupContiguousRowsP240_(targetRows).forEach(function(group) {
    const values = sheet.getRange(group.start, colNo, group.count, 1).getDisplayValues();
    for (let i = 0; i < values.length; i++) {
      out[group.start + i] = String(values[i][0] || '').trim();
    }
  });
  return out;
}

function filterMeaningfulMasterRowsForMetaP204_(sheet, rowNumbers, headerMap) {
  rowNumbers = (rowNumbers || []).map(function(v) { return Number(v) || 0; })
    .filter(function(v) { return v >= PORTAL_CONFIG.DATA_START_ROW; });
  if (!rowNumbers.length) return [];

  headerMap = headerMap || getHeaderMap_(sheet);
  const lastCol = sheet.getLastColumn();
  const minRow = Math.min.apply(null, rowNumbers);
  const maxRow = Math.max.apply(null, rowNumbers);
  const displayRows = sheet.getRange(minRow, 1, maxRow - minRow + 1, lastCol).getDisplayValues();

  const importantHeaders = [
    '고객번호', '회사명', '건물명', '현재 영업 진행 상황', '영업담당자', '견적담당',
    '고객사 담당자', '담당자', '대표전화', '전화번호', '직통번호', '이메일',
    '주소', '상세주소', '고객사 상세 주소', '메모', '최종 견적가', '최종견적가'
  ];
  const metaHeaders = PORTAL_MASTER_META_HEADERS_P201.concat(['원본수정시각', '마스터원본버전']);
  const importantCols = [];
  importantHeaders.forEach(function(h) {
    const col = findFirstExistingHeaderCol_(headerMap, [h]);
    if (col && importantCols.indexOf(col) < 0) importantCols.push(col);
  });

  const metaCols = metaHeaders.map(function(h) { return findFirstExistingHeaderCol_(headerMap, [h]); }).filter(Boolean);

  return rowNumbers.filter(function(rowNo) {
    const values = displayRows[rowNo - minRow] || [];
    if (!values.length) return false;

    // 고객번호/회사명 등 핵심 헤더 중 하나라도 값이 있으면 실데이터 행으로 봅니다.
    for (let i = 0; i < importantCols.length; i++) {
      const v = String(values[importantCols[i] - 1] || '').trim();
      if (v) return true;
    }

    // 혹시 헤더명이 달라져도, 메타 컬럼을 제외한 실제 행 데이터가 하나라도 있으면 허용합니다.
    // 단순 드롭다운/수식 공란 행은 displayValue가 비어 있으므로 여기서 걸러집니다.
    for (let c = 1; c <= values.length; c++) {
      if (metaCols.indexOf(c) >= 0) continue;
      const v = String(values[c - 1] || '').trim();
      if (v) return true;
    }
    return false;
  });
}


function getPortalEditedMasterColumnNamesP209_(sheet, headerMap, colStart, colEnd) {
  headerMap = headerMap || getHeaderMap_(sheet);
  const reverse = {};
  Object.keys(headerMap || {}).forEach(function(h) {
    const c = Number(headerMap[h]) || 0;
    if (c && !reverse[c]) reverse[c] = h;
  });
  const names = [];
  for (let c = Number(colStart) || 1; c <= (Number(colEnd) || Number(colStart) || 1); c++) {
    const name = reverse[c] || ('COL' + c);
    if (PORTAL_MASTER_META_HEADERS_P201.indexOf(name) >= 0) continue;
    names.push(name);
  }
  return names;
}

/**
 * P2-4: 이미 잘못 찍힌 빈 행의 관리용 메타 값을 정리하는 수동 실행 함수.
 * Apps Script에서 필요 시 1회 실행하세요.
 */
function clearBlankMasterMetaRowsP204() {
  const sheet = getMasterSheet_();
  const headerMap = getHeaderMap_(sheet);
  const updatedAtCol = findFirstExistingHeaderCol_(headerMap, ['수정일시']);
  const versionCol = findFirstExistingHeaderCol_(headerMap, ['수정버전']);
  const editorCol = findFirstExistingHeaderCol_(headerMap, ['최종수정자']);
  const metaCols = [updatedAtCol, versionCol, editorCol].filter(Boolean);
  if (!metaCols.length) return { ok: true, cleared: 0, message: '관리용 헤더가 없습니다.' };

  const lastRow = sheet.getLastRow();
  if (lastRow < PORTAL_CONFIG.DATA_START_ROW) return { ok: true, cleared: 0 };
  const rows = [];
  for (let r = PORTAL_CONFIG.DATA_START_ROW; r <= lastRow; r++) rows.push(r);
  const meaningful = {};
  filterMeaningfulMasterRowsForMetaP204_(sheet, rows, headerMap).forEach(function(r) { meaningful[r] = true; });

  let cleared = 0;
  rows.forEach(function(r) {
    if (meaningful[r]) return;
    let hasMeta = false;
    metaCols.forEach(function(c) {
      if (String(sheet.getRange(r, c).getDisplayValue() || '').trim()) hasMeta = true;
    });
    if (!hasMeta) return;
    metaCols.forEach(function(c) { sheet.getRange(r, c).clearContent(); });
    cleared++;
  });
  SpreadsheetApp.flush();
  return { ok: true, cleared: cleared };
}

function syncEditedMasterRowsOnceP201(rowNumbers) {
  rowNumbers = Array.isArray(rowNumbers) ? rowNumbers : [rowNumbers];
  const rows = rowNumbers.map(function(v) { return Number(v) || 0; }).filter(function(v) { return v >= PORTAL_CONFIG.DATA_START_ROW; });
  if (!rows.length) throw new Error('갱신할 마스터 행번호가 없습니다.');
  const result = [];
  rows.forEach(function(r) {
    try { result.push(updateCustomerSearchIndexRow_(r)); }
    catch (err) { result.push({ ok: false, rowNo: r, error: err && err.message ? err.message : String(err) }); }
  });
  markPortalMasterDataChangedP201_('수동 행 갱신 rows=' + rows.join(','));
  return { ok: true, rows: result };
}
