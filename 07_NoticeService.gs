/***************************************
 * S1 Sales Portal - 07_NoticeService.gs
 * PATCH S: 권한_DB 기반 공지 작성/삭제 권한 적용
 ***************************************/

function getPortalNotices(limit) {
  limit = Math.max(0, Number(limit) || 0);
  const active = getPortalNoticeActiveListCachedP16_();
  return limit ? active.slice(0, limit) : active;
}

function getPortalNoticeActiveListCachedP16_() {
  const cache = CacheService.getScriptCache();
  const key = 'S1_PORTAL_NOTICE_ACTIVE_LIST_P16';

  try {
    const cached = cache.get(key);
    if (cached) {
      const parsed = JSON.parse(cached);
      if (Array.isArray(parsed)) return parsed;
    }
  } catch (err) {}

  const rows = readPortalNoticeRows_();
  const active = rows
    .filter(item => String(item.deleted || '').toUpperCase() !== 'Y')
    .sort((a, b) => {
      const at = a.createdAtValue ? a.createdAtValue.getTime() : 0;
      const bt = b.createdAtValue ? b.createdAtValue.getTime() : 0;
      return bt - at;
    })
    .map(item => {
      const copy = Object.assign({}, item);
      delete copy.createdAtValue;
      return copy;
    });

  try { cache.put(key, JSON.stringify(active), 60); } catch (err) {}
  return active;
}

function clearPortalNoticeCacheP16_() {
  try { CacheService.getScriptCache().remove('S1_PORTAL_NOTICE_ACTIVE_LIST_P16'); } catch (err) {}
}

function getPortalNoticeDetail(noticeId) {
  const id = String(noticeId || '').trim();
  if (!id) throw new Error('공지 ID가 없습니다.');

  const sheet = ensurePortalNoticeSheet_();
  const rows = readPortalNoticeRows_(sheet);
  const hit = rows.find(item => String(item.id || '') === id && String(item.deleted || '').toUpperCase() !== 'Y');
  if (!hit) throw new Error('공지사항을 찾지 못했습니다.');

  const result = Object.assign({}, hit);
  delete result.createdAtValue;
  return result;
}


function findPortalNoticeRowByIdP489_(sheet, noticeId) {
  const id = String(noticeId || '').trim();
  if (!sheet || !id) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_CONFIG.NOTICE_HEADERS.length).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === id) {
      return { rowNo: i + 2, values: values[i], item: { id: String(values[i][0] || '').trim(), deleted: String(values[i][7] || '').trim() } };
    }
  }
  return null;
}


function normalizePortalNoticeCompareTextP489_(key, value) {
  key = String(key || '').trim();
  if (key === 'noticeDate') {
    const d = parsePortalDateOnly_(value);
    if (d) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value == null ? '' : value).replace(/\r?\n/g, '\n').trim();
}

function savePortalNotice(payload) {
  payload = Object.assign({}, payload || {});
  payload.clientSaveSource = String(payload.clientSaveSource || 'notice.save');
  payload.noticeId = String(payload.noticeId || payload.id || '').trim() || Utilities.getUuid().slice(0, 8);
  try {
    return savePortalNoticeCoreP489_(payload);
  } catch (err) {
    if (typeof enqueueSaveFallbackP473_ === 'function' &&
        typeof isPortalServerTransientWriteErrorP473_ === 'function' &&
        (isPortalServerTransientWriteErrorP473_(err) || (typeof isPortalServerStaleErrorP473_ === 'function' && isPortalServerStaleErrorP473_(err)))) {
      return enqueueSaveFallbackP473_('savePortalNotice', payload, 'NOTICE_DIRECT_SAVE_FAILED', err);
    }
    throw err;
  }
}

function savePortalNoticeCoreP489_(payload) {
  assertPortalCanWriteNotice_();
  payload = payload || {};
  const title = String(payload.title || '').trim();
  const content = String(payload.content || '').trim();
  if (!title) throw new Error('공지 제목을 입력하세요.');
  if (!content) throw new Error('공지 내용을 입력하세요.');

  const sheet = ensurePortalNoticeSheet_();
  const now = new Date();
  const createdAt = parsePortalDateOnly_(payload.noticeDate) || now;
  const author = String(payload.author || '').trim() || getPortalCurrentUserName_();
  const id = String(payload.noticeId || payload.id || '').trim() || Utilities.getUuid().slice(0, 8);

  const existing = findPortalNoticeRowByIdP489_(sheet, id);
  if (existing && existing.item && String(existing.item.deleted || '').toUpperCase() !== 'Y') {
    return { ok: true, duplicate: true, notice: getPortalNoticeDetail(id), notices: getPortalNotices(0), message: '이미 처리된 공지사항 저장입니다.' };
  }

  sheet.appendRow([
    id,
    now,
    Utilities.formatDate(createdAt, Session.getScriptTimeZone(), 'yyyy. MM. dd'),
    author,
    title,
    content,
    '',
    '',
    '',
    ''
  ]);

  SpreadsheetApp.flush();
  clearPortalNoticeCacheP16_();
  try {
    appendPortalActivityLog_({
      actionType: '공지사항',
      screen: '공지사항',
      summary: '공지 작성: ' + title,
      detail: { noticeId: id }
    });
  } catch (err) {}
  return { ok: true, notice: getPortalNoticeDetail(id), notices: getPortalNotices(0) };
}

function updatePortalNotice(payload) {
  payload = Object.assign({}, payload || {});
  payload.clientSaveSource = String(payload.clientSaveSource || 'notice.update');
  try {
    return updatePortalNoticeCoreP489_(payload);
  } catch (err) {
    if (typeof enqueueSaveFallbackP473_ === 'function' &&
        typeof isPortalServerTransientWriteErrorP473_ === 'function' &&
        (isPortalServerTransientWriteErrorP473_(err) || (typeof isPortalServerStaleErrorP473_ === 'function' && isPortalServerStaleErrorP473_(err)))) {
      return enqueueSaveFallbackP473_('updatePortalNotice', payload, 'NOTICE_UPDATE_DIRECT_SAVE_FAILED', err);
    }
    throw err;
  }
}

function updatePortalNoticeCoreP489_(payload) {
  assertPortalCanWriteNotice_();
  payload = payload || {};

  const id = String(payload.id || payload.noticeId || '').trim();
  const title = String(payload.title || '').trim();
  const content = String(payload.content || '').trim();

  if (!id) throw new Error('수정할 공지 ID가 없습니다.');
  if (!title) throw new Error('공지 제목을 입력하세요.');
  if (!content) throw new Error('공지 내용을 입력하세요.');

  const sheet = ensurePortalNoticeSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('공지사항이 없습니다.');

  const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_CONFIG.NOTICE_HEADERS.length).getValues();
  const now = new Date();
  const noticeDate = parsePortalDateOnly_(payload.noticeDate);
  const author = String(payload.author || '').trim() || getPortalCurrentUserName_();
  const editor = getPortalCurrentUserName_();
  const tz = Session.getScriptTimeZone();

  for (let i = 0; i < values.length; i++) {
    const rowId = String(values[i][0] || '').trim();
    const deleted = String(values[i][7] || '').trim().toUpperCase();
    if (rowId !== id || deleted === 'Y') continue;

    const rowNo = i + 2;
    const rowValues = values[i].slice();

    if (payload.expectedValues && typeof payload.expectedValues === 'object') {
      const currentExpectedMap = {
        noticeDate: String(rowValues[2] || '').trim(),
        author: String(rowValues[3] || '').trim(),
        title: String(rowValues[4] || '').trim(),
        content: String(rowValues[5] || '').replace(/\r?\n/g, '\n')
      };
      Object.keys(currentExpectedMap).forEach(function(k) {
        if (!Object.prototype.hasOwnProperty.call(payload.expectedValues, k)) return;
        const expectedText = normalizePortalNoticeCompareTextP489_(k, payload.expectedValues[k]);
        const currentText = normalizePortalNoticeCompareTextP489_(k, currentExpectedMap[k]);
        const nextText = normalizePortalNoticeCompareTextP489_(k, payload[k]);
        if (currentText !== expectedText && currentText !== nextText) {
          const err = new Error('다른 사용자가 같은 공지사항을 먼저 수정했습니다. 최신 내용을 확인한 뒤 다시 저장해 주세요.');
          err.code = 'PORTAL_FIELD_CONFLICT_P474';
          err.conflictInfoP474 = { key: k, label: k, currentValue: currentText, expectedValue: expectedText, attemptedValue: nextText, rowNo: rowNo, source: 'notice.update' };
          throw err;
        }
      });
    }

    rowValues[2] = noticeDate ? Utilities.formatDate(noticeDate, tz, 'yyyy. MM. dd') : String(payload.noticeDate || rowValues[2] || '').trim();
    rowValues[3] = author;
    rowValues[4] = title;
    rowValues[5] = content;
    rowValues[8] = now;
    rowValues[9] = editor;

    sheet.getRange(rowNo, 1, 1, PORTAL_CONFIG.NOTICE_HEADERS.length).setValues([rowValues]);

    SpreadsheetApp.flush();
    clearPortalNoticeCacheP16_();
    try {
      appendPortalActivityLog_({
        actionType: '공지사항',
        screen: '공지사항',
        summary: '공지 수정: ' + title,
        detail: { noticeId: id, editor: editor }
      });
    } catch (err) {}

    return { ok: true, notice: getPortalNoticeDetail(id), notices: getPortalNotices(0) };
  }

  throw new Error('수정할 공지사항을 찾지 못했습니다.');
}

function deletePortalNotices(ids) {
  assertPortalCanDeleteNotice_();
  ids = Array.isArray(ids) ? ids.map(String).filter(Boolean) : [];
  if (!ids.length) throw new Error('삭제할 공지를 선택하세요.');

  const sheet = ensurePortalNoticeSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, deleted: 0, notices: [] };

  const idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const idSet = {};
  ids.forEach(id => idSet[id] = true);

  let deleted = 0;
  idValues.forEach((row, i) => {
    const id = String(row[0] || '').trim();
    if (idSet[id]) {
      sheet.getRange(i + 2, 8).setValue('Y');
      deleted++;
    }
  });

  SpreadsheetApp.flush();
  clearPortalNoticeCacheP16_();
  try {
    appendPortalActivityLog_({
      actionType: '공지사항',
      screen: '공지사항',
      summary: '공지 삭제: ' + deleted + '건',
      detail: { noticeIds: ids, deleted: deleted }
    });
  } catch (err) {}
  return { ok: true, deleted: deleted, notices: getPortalNotices(0) };
}

function confirmPortalNotice(noticeId) {
  const id = String(noticeId || '').trim();
  if (!id) throw new Error('공지 ID가 없습니다.');

  const sheet = ensurePortalNoticeSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('공지사항이 없습니다.');

  const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_CONFIG.NOTICE_HEADERS.length).getValues();
  const user = getPortalCurrentUserName_();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy. MM. dd HH:mm');

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === id) {
      const current = String(values[i][6] || '').trim();
      const nextItem = user + ' (' + now + ')';
      const exists = current.indexOf(user + ' (') >= 0 || current.indexOf(user) >= 0;
      const next = exists ? current : (current ? current + '\n' + nextItem : nextItem);
      sheet.getRange(i + 2, 7).setValue(next);
      SpreadsheetApp.flush();
      clearPortalNoticeCacheP16_();
      return { ok: true, notice: getPortalNoticeDetail(id) };
    }
  }

  throw new Error('공지사항을 찾지 못했습니다.');
}

function ensurePortalNoticeSheet_() {
  const ss = getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(PORTAL_CONFIG.NOTICE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PORTAL_CONFIG.NOTICE_SHEET_NAME);
    sheet.getRange(1, 1, 1, PORTAL_CONFIG.NOTICE_HEADERS.length).setValues([PORTAL_CONFIG.NOTICE_HEADERS]);
    sheet.setFrozenRows(1);
  }

  const headers = sheet.getRange(1, 1, 1, Math.max(PORTAL_CONFIG.NOTICE_HEADERS.length, sheet.getLastColumn())).getDisplayValues()[0];
  const needInit = PORTAL_CONFIG.NOTICE_HEADERS.some((h, i) => String(headers[i] || '').trim() !== h);
  if (needInit && sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, PORTAL_CONFIG.NOTICE_HEADERS.length).setValues([PORTAL_CONFIG.NOTICE_HEADERS]);
  } else if (needInit && sheet.getLastRow() >= 1) {
    PORTAL_CONFIG.NOTICE_HEADERS.forEach((h, i) => {
      if (!String(headers[i] || '').trim()) sheet.getRange(1, i + 1).setValue(h);
    });
  }
  return sheet;
}

function readPortalNoticeRows_(sheet) {
  sheet = sheet || ensurePortalNoticeSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    seedDefaultPortalNotices_(sheet);
  }

  const lr = sheet.getLastRow();
  if (lr < 2) return [];

  const values = sheet.getRange(2, 1, lr - 1, PORTAL_CONFIG.NOTICE_HEADERS.length).getValues();
  return values.map(row => {
    const createdAt = row[1] instanceof Date ? row[1] : parsePortalDateOnly_(row[2]);
    return {
      id: String(row[0] || '').trim(),
      createdAt: formatDateTimeText_(row[1]),
      createdAtValue: createdAt,
      noticeDate: String(row[2] || '').trim() || formatDateText_(row[1]),
      author: String(row[3] || '').trim(),
      title: String(row[4] || '').trim(),
      content: String(row[5] || '').trim(),
      confirmers: String(row[6] || '').trim(),
      deleted: String(row[7] || '').trim(),
      updatedAt: formatDateTimeText_(row[8]),
      updatedAtValue: row[8] instanceof Date ? row[8] : null,
      updatedBy: String(row[9] || '').trim(),
      updatedDate: row[8] ? formatDateTimeText_(row[8]) : ''
    };
  }).filter(item => item.id || item.title || item.content);
}

function seedDefaultPortalNotices_(sheet) {
  sheet = sheet || ensurePortalNoticeSheet_();
  if (sheet.getLastRow() >= 2) return;

  const now = new Date();
  const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy. MM. dd');
  const rows = [
    [Utilities.getUuid().slice(0, 8), now, today, '문형진', '영업전산 개편 공지', '안녕하세요 여러분 영업전산이 새로운 내용을 가지게 되었습니다.\n서무님들이 개편한 마스터시트와 연동된 웹앱 기능을 만들어보았습니다.\n테스트 부탁드립니다.', '', '', '', ''],
    [Utilities.getUuid().slice(0, 8), now, '2026. 06. 16', '문형진', '디엠정보기술 사업종료', '최근 저희와 함께 일하던 디엠정보기술이 사업을 종료했습니다. 수행사 배정 시 참고 부탁드립니다.', '', '', '', ''],
    [Utilities.getUuid().slice(0, 8), now, '2026. 06. 16', '문형진', '영진약품 변경 관련 내용', '기존 디엠에서 케이제이로 수행사를 변경하는 내용입니다. 관련 고객 응대 시 확인 부탁드립니다.', '', '', '', '']
  ];
  sheet.getRange(2, 1, rows.length, PORTAL_CONFIG.NOTICE_HEADERS.length).setValues(rows);
  SpreadsheetApp.flush();
  clearPortalNoticeCacheP16_();
}

function parsePortalDateOnly_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const text = String(value || '').trim();
  if (!text) return null;

  const m = text.match(/(\d{4})[.\-\/\s]+(\d{1,2})[.\-\/\s]+(\d{1,2})/);
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return isNaN(d.getTime()) ? null : d;
}

