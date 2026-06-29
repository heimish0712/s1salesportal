/***************************************
 * S1 Sales Portal - 06_ContactHistoryService.gs
 * 분리일: 2026-06-19
 * 원칙: 기능 변경 없이 최신 단일 파일을 물리적으로 분리
 ***************************************/

function addContactHistory(payload) {
  payload = payload || {};
  const target = assertCustomerTarget_(payload, '컨택이력 저장', { readObject: true });
  assertPortalCanAccessCustomerTarget_(target, '컨택이력 저장');
  const rowNo = target.rowNo;

  const roundNo = Number(payload.roundNo);
  const method = String(payload.method || '').trim();
  const contactAt = String(payload.contactAt || '').trim() || new Date();
  const content = String(payload.content || payload.tag || payload.tags || '').trim();
  const nextAction = String(payload.nextAction || '').trim();
  let nextActionAt = String(payload.nextActionAt || '').trim();
  const nextActionAuthor = String(payload.nextActionAuthor || payload.nextAssignee || '').trim();
  const tag = String(payload.tag || payload.tags || '').trim();
  const nextActionTags = (typeof normalizePortalTodayTags_ === 'function') ? normalizePortalTodayTags_(payload.nextActionTags || payload.todoTags || '') : String(payload.nextActionTags || payload.todoTags || '').split(',').map(function(v){ return String(v || '').trim(); }).filter(Boolean);
  const nextActionTagText = nextActionTags.join(', ');

  if (!roundNo) throw new Error('몇 차 컨택인지 선택하세요.');
  if (!method) throw new Error('연락수단을 선택하세요.');
  if (!content) throw new Error('컨택내용 또는 태그를 입력하세요.');
  // P432: 다음 액션을 입력했는데 일시를 비워두면 오늘 할 일에 즉시 보이도록 현재 시각을 사용합니다.
  if (nextAction && !nextActionAt) nextActionAt = new Date();

  const masterSheet = target.sheet;
  const masterObj = target.obj || readMasterRowObject_(masterSheet, rowNo);
  const customerNo = getMasterFieldValue_(masterObj, 'customerNo') || masterObj['고객번호'] || '';
  const company = getCompanyValue_(masterObj);
  if (!company && !customerNo) throw new Error('해당 행에서 고객정보를 찾지 못했습니다.');

  const author = String(payload.author || '').trim() || getPortalCurrentUserName_();
  const contactAtText = formatMemoDateTime_(contactAt);
  // P432: 마스터 메모에는 컨택 이력만 남깁니다. 다음 액션은 오늘 할 일 쪽에서만 노출합니다.
  const masterLogText = '[컨택이력][' + roundNo + '차][' + method + '] ' + contactAtText + ' ' + content + '(' + author + ')';

  const updatedMemo = appendToMasterMemo_(masterSheet, rowNo, masterLogText);

  const historyId = appendHistoryRecord_(getWebAppDbSpreadsheet_(), {
    customerNo,
    company,
    rowNo,
    author,
    recordType: '컨택이력',
    roundNo: roundNo + '차',
    method,
    tag,
    status: '',
    content,
    note: '',
    nextAction,
    nextActionAt,
    nextActionTags: nextActionTagText,
    nextActionAuthor: nextActionAuthor || author,
    masterApplied: 'Y',
    clientRequestId: payload.clientRequestId || ''
  });

  let indexUpdate = null;
  try {
    indexUpdate = updateCustomerSearchIndexAfterMutation_({ rowNo: rowNo, customerNo: customerNo }, { memo: updatedMemo }, 'CONTACT_HISTORY_MEMO');
  } catch (err) {
    Logger.log('검색인덱스 컨택이력 갱신 실패: ' + (err && err.stack || err));
    try { indexUpdate = markCustomerSearchIndexDirty_('CONTACT_HISTORY_MEMO', err && err.message ? err.message : String(err)); } catch (e) {}
  }
  try { CacheService.getScriptCache().remove('PORTAL_DASHBOARD_V27'); } catch (err) {}
  try { CacheService.getScriptCache().remove('PORTAL_DASHBOARD_V46_FAST_HOME'); } catch (err) {}
  if (nextAction) {
    try {
      const nextDateKey = (typeof normalizePortalTodoDate_ === 'function') ? normalizePortalTodoDate_(nextActionAt || new Date()) : '';
      if (nextDateKey) CacheService.getScriptCache().remove('PORTAL_TODAY_NEXT_RAW_P370_' + nextDateKey);
    } catch (err) {}
  }

  try {
    appendPortalActivityLog_({
      actionType: '컨택이력추가',
      screen: '고객 상세/컨택이력',
      rowNo: rowNo,
      customerNo: customerNo,
      company: company,
      summary: '[' + roundNo + '차][' + method + '] ' + content,
      detail: { tag: tag, nextAction: nextAction, nextActionAt: nextActionAt, nextActionTags: nextActionTagText, nextActionAuthor: nextActionAuthor || author, historyId: historyId }
    });
  } catch (err) {}

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: '컨택이력이 마스터 메모와 컨택이력_DB에 저장되었습니다.',
    historyId,
    rowNo,
    customerNo,
    company,
    updatedMemo,
    nextActionSaved: !!nextAction,
    nextActionAt: nextActionAt ? formatMemoDateTime_(nextActionAt) : '',
    nextActionAuthor: nextActionAuthor || author,
    indexUpdate: indexUpdate
  };
}

function addTempMemo(payload) {
  payload = payload || {};
  const target = assertCustomerTarget_(payload, '임시메모 저장', { readObject: true });
  assertPortalCanAccessCustomerTarget_(target, '임시메모 저장');
  const rowNo = target.rowNo;

  const memoAt = String(payload.memoAt || '').trim();
  const memo = String(payload.memo || '').trim();
  if (!memoAt) throw new Error('임시메모 날짜/시간을 입력하세요.');
  if (!memo) throw new Error('임시메모 내용을 입력하세요.');

  const masterSheet = target.sheet;
  const masterObj = target.obj || readMasterRowObject_(masterSheet, rowNo);
  const customerNo = getMasterFieldValue_(masterObj, 'customerNo') || masterObj['고객번호'] || '';
  const company = getCompanyValue_(masterObj);
  const author = getPortalCurrentUserName_();

  const historyId = appendHistoryRecord_(getWebAppDbSpreadsheet_(), {
    customerNo,
    company,
    rowNo,
    author,
    recordType: '임시메모',
    roundNo: '',
    method: '',
    status: '',
    content: memo,
    note: '임시메모 일시: ' + formatMemoDateTime_(memoAt),
    nextAction: '',
    nextActionAt: '',
    masterApplied: 'N'
  });

  return {
    ok: true,
    message: '임시메모가 저장되었습니다. 마스터시트 메모에는 반영하지 않았습니다.',
    historyId,
    rowNo
  };
}

function updateContractStatus(payload) {
  payload = payload || {};
  const target = assertCustomerTarget_(payload, '계약진행상태 변경', { readObject: true });
  assertPortalCanAccessCustomerTarget_(target, '계약진행상태 변경');
  const rowNo = target.rowNo;

  const status = String(payload.status || '').trim();
  const changedAt = String(payload.changedAt || '').trim();
  const note = String(payload.note || '').trim();

  if (!status) throw new Error('계약진행상태를 선택하세요.');
  if (!changedAt) throw new Error('상태 변경 날짜/시간을 입력하세요.');

  const masterSheet = target.sheet;
  let headerMap = getHeaderMap_(masterSheet);
  const statusCol = findMasterFieldCol_(headerMap, 'status') ||
    ensureMasterFieldColumn_(masterSheet, headerMap, 'status');

  const beforeObj = target.obj || readMasterRowObject_(masterSheet, rowNo);
  const beforeStatus = getStatusValueFromObj_(beforeObj);
  const customerNo = getMasterFieldValue_(beforeObj, 'customerNo') || beforeObj['고객번호'] || '';
  const company = getCompanyValue_(beforeObj);
  const author = getPortalCurrentUserName_();

  masterSheet.getRange(rowNo, statusCol).setValue(status);

  const changedAtText = formatMemoDateTime_(changedAt);
  const notePart = note ? (' ' + note) : '';
  const masterLogText = '[계약진행상태 변경][' + status + '] ' + changedAtText + notePart + '(' + author + ')';
  const updatedMemo = appendToMasterMemo_(masterSheet, rowNo, masterLogText);

  const historyId = appendHistoryRecord_(getWebAppDbSpreadsheet_(), {
    customerNo,
    company,
    rowNo,
    author,
    recordType: '계약진행상태 변경',
    roundNo: '',
    method: '',
    status,
    content: note,
    note: '이전상태: ' + (beforeStatus || '-'),
    nextAction: '',
    nextActionAt: '',
    masterApplied: 'Y'
  });

  let indexUpdate = null;
  try {
    indexUpdate = updateCustomerSearchIndexAfterMutation_({ rowNo: rowNo, customerNo: customerNo }, { status: status, memo: updatedMemo }, 'STATUS_CHANGE');
  } catch (err) {
    Logger.log('검색인덱스 상태변경 갱신 실패: ' + (err && err.stack || err));
    try { indexUpdate = markCustomerSearchIndexDirty_('STATUS_CHANGE', err && err.message ? err.message : String(err)); } catch (e) {}
  }

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: '계약진행상태와 마스터 메모 로그가 저장되었습니다.',
    rowNo,
    customerNo,
    status,
    historyId,
    updatedMemo,
    indexUpdate: indexUpdate
  };
}

function addContactMemo(payload) {
  return addContactHistory(payload);
}

function updateCustomerStatus(payload) {
  return updateContractStatus(payload);
}

function appendHistoryRecord_(ss, record) {
  const sheet = ensureContactHistorySheet_(ss);
  const now = new Date();
  const historyId = 'CH-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + shortUuid_();
  const headerMap = getSimpleHeaderMapFromRow_(sheet, 1);

  const rowObj = {
    '이력ID': historyId,
    '고객번호': record.customerNo || '',
    '회사명': record.company || '',
    '마스터행': record.rowNo || '',
    '작성일시': now,
    '작성자': record.author || getPortalCurrentUserName_(),
    '기록구분': record.recordType || '',
    '차수': record.roundNo || '',
    '연락수단': record.method || '',
    '태그': record.tag || '',
    '계약진행상태': record.status || '',
    '컨택내용': record.content || '',
    '특이사항': record.note || '',
    '다음액션': record.nextAction || '',
    '다음액션일시': record.nextActionAt ? formatMemoDateTime_(record.nextActionAt) : '',
    '다음액션태그': record.nextActionTags || '',
    '다음액션담당자': record.nextActionAuthor || record.author || '',
    '마스터메모반영': record.masterApplied || 'N',
    '클라이언트요청ID': record.clientRequestId || ''
  };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(h => String(h || '').trim());
  const row = headers.map(h => rowObj[h] != null ? rowObj[h] : '');
  sheet.appendRow(row);
  return historyId;
}

function appendToMasterMemo_(masterSheet, rowNo, logText) {
  const headerMap = getHeaderMap_(masterSheet);
  const memoCol = findMasterFieldCol_(headerMap, 'memo') ||
    ensureMasterFieldColumn_(masterSheet, headerMap, 'memo');

  const range = masterSheet.getRange(rowNo, memoCol);
  const current = String(range.getDisplayValue() || '').trim();
  const next = current ? (current + '\n' + logText) : logText;
  range.setValue(next);
  return next;
}

function getContactHistoryByCustomer_(customerNo, rowNo) {
  const ss = getWebAppDbSpreadsheet_();
  const sheet = ss.getSheetByName(PORTAL_CONFIG.CONTACT_HISTORY_SHEET_NAME);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h || '').trim());
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i; });

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  const keyCustomerNo = String(customerNo || '').trim();
  const keyRowNo = String(rowNo || '').trim();

  const rows = values
    .map(row => {
      const oldMemo = cellByHeader_(row, map, '메모');
      const oldType = cellByHeader_(row, map, '컨택방식');
      const oldResult = cellByHeader_(row, map, '통화결과');
      const oldNextDate = cellByHeader_(row, map, '다음연락일');
      return {
        historyId: cellByHeader_(row, map, '이력ID'),
        customerNo: cellByHeader_(row, map, '고객번호'),
        company: cellByHeader_(row, map, '회사명'),
        rowNo: cellByHeader_(row, map, '마스터행'),
        createdAt: cellByHeader_(row, map, '작성일시'),
        author: cellByHeader_(row, map, '작성자'),
        recordType: cellByHeader_(row, map, '기록구분') || oldType || '이력',
        roundNo: cellByHeader_(row, map, '차수'),
        method: cellByHeader_(row, map, '연락수단') || oldType,
        status: cellByHeader_(row, map, '계약진행상태') || oldResult,
        content: cellByHeader_(row, map, '컨택내용') || oldMemo,
        note: cellByHeader_(row, map, '특이사항'),
        nextAction: cellByHeader_(row, map, '다음액션'),
        nextActionAt: cellByHeader_(row, map, '다음액션일시') || oldNextDate,
        nextActionTags: cellByHeader_(row, map, '다음액션태그'),
        nextActionAuthor: cellByHeader_(row, map, '다음액션담당자'),
        masterApplied: cellByHeader_(row, map, '마스터메모반영')
      };
    })
    .filter(item => {
      if (keyCustomerNo && String(item.customerNo || '').trim() === keyCustomerNo) return true;
      if (keyRowNo && String(item.rowNo || '').trim() === keyRowNo) return true;
      return false;
    })
    .reverse()
    .slice(0, PORTAL_CONFIG.CONTACT_HISTORY_MAX_PER_CUSTOMER || 30);

  return rows;
}

function ensureContactHistorySheet_(ss) {
  ss = ss || getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(PORTAL_CONFIG.CONTACT_HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PORTAL_CONFIG.CONTACT_HISTORY_SHEET_NAME);
    sheet.getRange(1, 1, 1, PORTAL_CONFIG.CONTACT_HISTORY_HEADERS.length)
      .setValues([PORTAL_CONFIG.CONTACT_HISTORY_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, PORTAL_CONFIG.CONTACT_HISTORY_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#f2f4f7');
    sheet.autoResizeColumns(1, PORTAL_CONFIG.CONTACT_HISTORY_HEADERS.length);
    return sheet;
  }

  ensureSheetHeaders_(sheet, PORTAL_CONFIG.CONTACT_HISTORY_HEADERS);
  return sheet;
}

