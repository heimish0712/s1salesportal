/***************************************
 * S1 Sales Portal - 21_ChangeQueueService.gs
 * P2-9: 마스터 변경큐 + 검색인덱스 재처리 안정화
 ***************************************/

const PORTAL_CHANGE_QUEUE_SHEET_NAME_P209 = '변경큐_DB';
const PORTAL_CHANGE_QUEUE_HEADERS_P209 = [
  '변경ID', '일시', '시트명', '행번호', '고객번호', '변경컬럼', '변경자',
  '상태', '처리일시', '오류메시지', '시도횟수'
];

function ensurePortalChangeQueueSheetP209_() {
  const ss = getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(PORTAL_CHANGE_QUEUE_SHEET_NAME_P209);
  if (!sheet) sheet = ss.insertSheet(PORTAL_CHANGE_QUEUE_SHEET_NAME_P209);
  ensurePortalChangeQueueHeadersP209_(sheet);
  return sheet;
}

function ensurePortalChangeQueueHeadersP209_(sheet) {
  const width = Math.max(sheet.getLastColumn(), PORTAL_CHANGE_QUEUE_HEADERS_P209.length, 1);
  const current = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const existing = {};
  current.forEach(function(h, i) { if (h) existing[h] = i + 1; });
  let changed = false;
  if (sheet.getLastRow() < 1 || current.every(function(h) { return !h; })) {
    sheet.getRange(1, 1, 1, PORTAL_CHANGE_QUEUE_HEADERS_P209.length).setValues([PORTAL_CHANGE_QUEUE_HEADERS_P209]);
    changed = true;
  } else {
    PORTAL_CHANGE_QUEUE_HEADERS_P209.forEach(function(h) {
      if (existing[h]) return;
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(h);
      existing[h] = col;
      changed = true;
    });
  }
  try {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold').setBackground('#f2f4f7');
  } catch (err) {}
  return changed;
}

function getPortalChangeQueueHeaderMapP209_(sheet) {
  sheet = sheet || ensurePortalChangeQueueSheetP209_();
  ensurePortalChangeQueueHeadersP209_(sheet);
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), PORTAL_CHANGE_QUEUE_HEADERS_P209.length)).getDisplayValues()[0];
  const map = {};
  headers.forEach(function(h, i) {
    h = String(h || '').trim();
    if (h && !map[h]) map[h] = i + 1;
  });
  return map;
}

function getPortalChangeQueueNowP209_(date) {
  return Utilities.formatDate(date || new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function makePortalChangeIdP209_(rowNo, suffix) {
  const now = new Date();
  return 'CHG-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmssSSS') + '-' + String(rowNo || '0') + '-' + String(suffix || Math.floor(Math.random() * 10000));
}

function appendPortalChangeQueueRowsP209_(items) {
  items = Array.isArray(items) ? items : [items];
  items = items.filter(Boolean);
  if (!items.length) return [];
  const sheet = ensurePortalChangeQueueSheetP209_();
  const headers = PORTAL_CHANGE_QUEUE_HEADERS_P209.slice();
  const nowText = getPortalChangeQueueNowP209_();
  const rows = items.map(function(item, idx) {
    const rowNo = Number(item.rowNo || item.masterRow || 0) || 0;
    const map = {
      '변경ID': item.changeId || makePortalChangeIdP209_(rowNo, idx),
      '일시': item.at || nowText,
      '시트명': item.sheetName || (PORTAL_CONFIG && PORTAL_CONFIG.MASTER_SHEET_NAME) || '마스터시트(신규)',
      '행번호': rowNo || '',
      '고객번호': item.customerNo || '',
      '변경컬럼': item.changedColumns || item.changedColumn || '',
      '변경자': item.editor || '',
      '상태': item.status || 'PENDING',
      '처리일시': item.processedAt || '',
      '오류메시지': item.error || '',
      '시도횟수': Number(item.attempts || 0) || 0
    };
    return headers.map(function(h) { return map[h] == null ? '' : map[h]; });
  });
  const start = Math.max(2, sheet.getLastRow() + 1);
  sheet.getRange(start, 1, rows.length, headers.length).setValues(rows);
  return rows.map(function(r) { return r[0]; });
}

function markPortalChangeQueueRowP209_(sheet, rowNo, status, error) {
  if (!rowNo || rowNo < 2) return;
  sheet = sheet || ensurePortalChangeQueueSheetP209_();
  const map = getPortalChangeQueueHeaderMapP209_(sheet);
  const nowText = getPortalChangeQueueNowP209_();
  const statusCol = map['상태'];
  const processedCol = map['처리일시'];
  const errorCol = map['오류메시지'];
  const attemptsCol = map['시도횟수'];
  if (statusCol) sheet.getRange(rowNo, statusCol).setValue(status || 'DONE');
  if (processedCol) sheet.getRange(rowNo, processedCol).setValue(nowText);
  if (errorCol) sheet.getRange(rowNo, errorCol).setValue(String(error || '').slice(0, 1000));
  if (attemptsCol) {
    const current = Number(sheet.getRange(rowNo, attemptsCol).getValue() || 0) || 0;
    sheet.getRange(rowNo, attemptsCol).setValue(current + 1);
  }
}

function markPortalChangeQueueRowsDoneByIdsP209_(ids, status, error) {
  ids = Array.isArray(ids) ? ids : [ids];
  const idSet = {};
  ids.forEach(function(id) { id = String(id || '').trim(); if (id) idSet[id] = true; });
  if (!Object.keys(idSet).length) return { ok: true, updated: 0 };
  const sheet = ensurePortalChangeQueueSheetP209_();
  const map = getPortalChangeQueueHeaderMapP209_(sheet);
  const idCol = map['변경ID'];
  if (!idCol || sheet.getLastRow() < 2) return { ok: true, updated: 0 };
  const vals = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getDisplayValues();
  let updated = 0;
  vals.forEach(function(r, i) {
    if (!idSet[String(r[0] || '').trim()]) return;
    markPortalChangeQueueRowP209_(sheet, i + 2, status || 'DONE', error || '');
    updated++;
  });
  return { ok: true, updated: updated };
}

function getPortalChangeQueueStatsP209() {
  const sheet = ensurePortalChangeQueueSheetP209_();
  const map = getPortalChangeQueueHeaderMapP209_(sheet);
  const lastRow = sheet.getLastRow();
  const stats = { total: Math.max(0, lastRow - 1), pending: 0, error: 0, done: 0, skipped: 0, processing: 0, samples: [] };
  if (lastRow < 2) return { ok: true, stats: stats };
  const width = Math.max(sheet.getLastColumn(), PORTAL_CHANGE_QUEUE_HEADERS_P209.length);
  const values = sheet.getRange(2, 1, lastRow - 1, width).getDisplayValues();
  values.forEach(function(row, idx) {
    const status = String(row[(map['상태'] || 8) - 1] || '').trim().toUpperCase() || 'PENDING';
    if (status === 'DONE') stats.done++;
    else if (status === 'SKIPPED') stats.skipped++;
    else if (status === 'ERROR') stats.error++;
    else if (status === 'PROCESSING') stats.processing++;
    else stats.pending++;
    if ((status === 'PENDING' || status === 'ERROR') && stats.samples.length < 10) {
      stats.samples.push({
        rowNo: idx + 2,
        changeId: row[(map['변경ID'] || 1) - 1] || '',
        masterRow: row[(map['행번호'] || 4) - 1] || '',
        customerNo: row[(map['고객번호'] || 5) - 1] || '',
        status: status,
        error: row[(map['오류메시지'] || 10) - 1] || ''
      });
    }
  });
  return { ok: true, stats: stats };
}

function processPortalChangeQueueP209(options) {
  options = options || {};
  const userInfo = typeof getPortalSystemHealthUserP208_ === 'function' ? getPortalSystemHealthUserP208_() : { allowed: true };
  if (typeof assertPortalSystemHealthAllowedP208_ === 'function') assertPortalSystemHealthAllowedP208_(userInfo);

  const limit = Math.max(1, Math.min(200, Number(options.limit) || 50));
  const includeErrors = options.includeErrors !== false;
  const sheet = ensurePortalChangeQueueSheetP209_();
  const map = getPortalChangeQueueHeaderMapP209_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, processed: 0, success: 0, skipped: 0, failed: 0, results: [] };

  const width = Math.max(sheet.getLastColumn(), PORTAL_CHANGE_QUEUE_HEADERS_P209.length);
  const values = sheet.getRange(2, 1, lastRow - 1, width).getDisplayValues();
  const targets = [];
  values.forEach(function(row, idx) {
    if (targets.length >= limit) return;
    const status = String(row[(map['상태'] || 8) - 1] || '').trim().toUpperCase() || 'PENDING';
    if (status !== 'PENDING' && !(includeErrors && status === 'ERROR')) return;
    const masterRow = Number(row[(map['행번호'] || 4) - 1] || 0) || 0;
    if (!masterRow) return;
    targets.push({ queueRow: idx + 2, rowNo: masterRow, changeId: row[(map['변경ID'] || 1) - 1] || '' });
  });
  if (!targets.length) return { ok: true, processed: 0, success: 0, skipped: 0, failed: 0, results: [] };

  return withPortalScriptLockP201_('change-queue-reprocess', function() {
    const result = { ok: true, processed: 0, success: 0, skipped: 0, failed: 0, results: [] };
    targets.forEach(function(t) {
      result.processed++;
      try {
        const masterSheet = getMasterSheet_();
        const headerMap = getHeaderMap_(masterSheet);
        const meaningful = filterMeaningfulMasterRowsForMetaP204_(masterSheet, [t.rowNo], headerMap);
        if (!meaningful.length) {
          try { updateCustomerSearchIndexRow_(t.rowNo); } catch (cleanupErr) {}
          markPortalChangeQueueRowP209_(sheet, t.queueRow, 'SKIPPED', '실데이터 없는 행이라 건너뜀');
          result.skipped++;
          result.results.push({ rowNo: t.rowNo, status: 'SKIPPED' });
          return;
        }
        const res = updateCustomerSearchIndexRow_(t.rowNo);
        if (res && res.ok) {
          markPortalChangeQueueRowP209_(sheet, t.queueRow, 'DONE', '');
          result.success++;
          result.results.push({ rowNo: t.rowNo, status: 'DONE', indexRow: res.indexRow || '' });
        } else {
          const msg = res && res.reason ? res.reason : '검색인덱스 갱신 실패';
          markPortalChangeQueueRowP209_(sheet, t.queueRow, 'ERROR', msg);
          result.failed++;
          result.results.push({ rowNo: t.rowNo, status: 'ERROR', error: msg });
        }
      } catch (err) {
        const msg = err && err.message ? err.message : String(err);
        markPortalChangeQueueRowP209_(sheet, t.queueRow, 'ERROR', msg);
        result.failed++;
        result.results.push({ rowNo: t.rowNo, status: 'ERROR', error: msg });
      }
    });
    markPortalMasterDataChangedP201_('변경큐 재처리 success=' + result.success + ', failed=' + result.failed + ', skipped=' + result.skipped);
    return result;
  }, { attempts: 3, waitMs: 700, sleepBaseMs: 180 });
}
