/***************************************
 * S1 Sales Portal - 08_TodayService.gs
 * 분리일: 2026-06-19
 * 원칙: 기능 변경 없이 최신 단일 파일을 물리적으로 분리
 ***************************************/

function getPortalTodayData(dateText) {
  const selectedDate = normalizePortalTodoDate_(dateText || new Date());
  return {
    ok: true,
    selectedDate: selectedDate,
    todos: getPortalTodosForDate_(selectedDate),
    nextActions: getContactNextActionsForDate_(selectedDate),
    actionOptions: PORTAL_NEXT_ACTION_OPTIONS
  };
}

function savePortalTodosForDate(payload) {
  payload = payload || {};
  const selectedDate = normalizePortalTodoDate_(payload.date || new Date());
  const todos = Array.isArray(payload.todos) ? payload.todos : [];
  const sheet = ensurePortalTodaySheet_();
  const now = new Date();
  const user = getCurrentUserLabel_();

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_CONFIG.TODAY_HEADERS.length).getValues();
    values.forEach(function(row, i) {
      const rowDate = normalizePortalTodoDate_(row[1]);
      if (rowDate === selectedDate && String(row[7] || '').toUpperCase() !== 'Y') {
        sheet.getRange(i + 2, 8).setValue('Y');
      }
    });
  }

  const rows = todos
    .map(function(item, idx) {
      item = item || {};
      const content = String(item.content || '').trim();
      if (!content) return null;
      return [
        Utilities.getUuid().slice(0, 8),
        selectedDate,
        idx + 1,
        content,
        item.done ? 'Y' : '',
        String(item.author || '').trim() || user,
        now,
        ''
      ];
    })
    .filter(Boolean);

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, PORTAL_CONFIG.TODAY_HEADERS.length).setValues(rows);
  }

  try {
    appendPortalActivityLog_({
      actionType: '오늘할일',
      screen: '오늘 할 일',
      summary: '오늘 할 일 저장: ' + selectedDate + ' / ' + rows.length + '건',
      detail: { date: selectedDate, count: rows.length }
    });
  } catch (err) {}

  SpreadsheetApp.flush();
  return getPortalTodayData(selectedDate);
}

function ensurePortalTodaySheet_() {
  const ss = getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(PORTAL_CONFIG.TODAY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PORTAL_CONFIG.TODAY_SHEET_NAME);
    sheet.getRange(1, 1, 1, PORTAL_CONFIG.TODAY_HEADERS.length).setValues([PORTAL_CONFIG.TODAY_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, PORTAL_CONFIG.TODAY_HEADERS.length).setFontWeight('bold').setBackground('#f2f4f7');
    return sheet;
  }
  ensureSheetHeaders_(sheet, PORTAL_CONFIG.TODAY_HEADERS);
  return sheet;
}

function getPortalTodosForDate_(selectedDate) {
  const sheet = ensurePortalTodaySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_CONFIG.TODAY_HEADERS.length).getValues();
  return values
    .map(function(row) {
      return {
        id: String(row[0] || '').trim(),
        date: normalizePortalTodoDate_(row[1]),
        order: Number(row[2]) || 0,
        content: String(row[3] || '').trim(),
        done: String(row[4] || '').toUpperCase() === 'Y',
        author: String(row[5] || '').trim(),
        updatedAt: formatDateTimeText_(row[6]),
        deleted: String(row[7] || '').trim()
      };
    })
    .filter(function(item) {
      return item.date === selectedDate && item.deleted.toUpperCase() !== 'Y' && item.content;
    })
    .sort(function(a, b) { return (a.order || 0) - (b.order || 0); });
}

function getContactNextActionsForDate_(selectedDate) {
  const ss = getWebAppDbSpreadsheet_();
  const sheet = ss.getSheetByName(PORTAL_CONFIG.CONTACT_HISTORY_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const map = {};
  headers.forEach(function(h, i) { if (h) map[h] = i; });
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();

  const perm = getPortalCurrentPermission_ ? getPortalCurrentPermission_() : null;
  const canViewAll = !!(perm && (perm.isAdmin || perm.canViewAllCustomers || perm.canViewAllActivityLogs));
  const myNames = [];
  if (perm) {
    myNames.push(perm.name, perm.salesRepName);
    (perm.salesRepAliases || []).forEach(function(v) { myNames.push(v); });
  }
  const myNameKeys = myNames
    .map(function(v) { return normalizePortalNameForPermission_ ? normalizePortalNameForPermission_(v) : String(v || '').replace(/\s+/g, ''); })
    .filter(Boolean);

  return values.map(function(row) {
      const nextAt = cellByHeader_(row, map, '다음액션일시') || cellByHeader_(row, map, '다음연락일');
      const nextDate = normalizePortalTodoDate_(nextAt);
      return {
        historyId: cellByHeader_(row, map, '이력ID'),
        customerNo: cellByHeader_(row, map, '고객번호'),
        company: cellByHeader_(row, map, '회사명'),
        rowNo: cellByHeader_(row, map, '마스터행'),
        nextAction: cellByHeader_(row, map, '다음액션'),
        nextActionAt: nextAt,
        nextDate: nextDate,
        content: cellByHeader_(row, map, '컨택내용'),
        author: cellByHeader_(row, map, '작성자')
      };
    })
    .filter(function(item) {
      if (!(item.nextDate === selectedDate && (item.nextAction || item.nextActionAt))) return false;
      if (canViewAll) return true;
      const authorKey = normalizePortalNameForPermission_ ? normalizePortalNameForPermission_(item.author) : String(item.author || '').replace(/\s+/g, '');
      return !myNameKeys.length || myNameKeys.indexOf(authorKey) >= 0;
    })
    .sort(function(a, b) { return String(a.nextActionAt || '').localeCompare(String(b.nextActionAt || '')); });
}

function normalizePortalTodoDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const text = String(value || '').trim();
  if (!text) return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) return iso[1] + '-' + String(iso[2]).padStart(2, '0') + '-' + String(iso[3]).padStart(2, '0');

  const dot = text.match(/(\d{4})[.\/\s-]+(\d{1,2})[.\/\s-]+(\d{1,2})/);
  if (dot) return dot[1] + '-' + String(dot[2]).padStart(2, '0') + '-' + String(dot[3]).padStart(2, '0');

  const short = text.match(/(\d{2})[.\/\s-]+(\d{1,2})[.\/\s-]+(\d{1,2})/);
  if (short) return '20' + short[1] + '-' + String(short[2]).padStart(2, '0') + '-' + String(short[3]).padStart(2, '0');

  const d = new Date(text);
  return isNaN(d.getTime()) ? Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

