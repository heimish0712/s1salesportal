/***************************************
 * S1 Sales Portal - 08_TodayService.gs
 * 오늘 할 일 통합형 개편
 * - 수동 할 일 + 컨택 예정/다음액션을 하나의 '오늘 할 일'로 통합
 * - 완료 체크 시 오늘 할 일 카운트에서 제외
 ***************************************/

function getPortalTodayData(dateText) {
  const selectedDate = normalizePortalTodoDate_(dateText || new Date());
  const storedTodos = getPortalTodosForDate_(selectedDate);
  const nextActions = getContactNextActionsForDate_(selectedDate);
  const tasks = buildPortalTodayUnifiedTasks_(storedTodos, nextActions, selectedDate);
  return {
    ok: true,
    selectedDate: selectedDate,
    tasks: tasks,
    todos: tasks,
    nextActions: nextActions,
    actionOptions: PORTAL_NEXT_ACTION_OPTIONS
  };
}

function savePortalTodosForDate(payload) {
  payload = payload || {};
  const selectedDate = normalizePortalTodoDate_(payload.date || new Date());
  const todos = Array.isArray(payload.tasks) ? payload.tasks : (Array.isArray(payload.todos) ? payload.todos : []);
  const sheet = ensurePortalTodaySheet_();
  const now = new Date();
  const user = getCurrentUserLabel_();
  const headerLen = PORTAL_CONFIG.TODAY_HEADERS.length;

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, headerLen).getValues();
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
      const sourceType = String(item.sourceType || '').trim() || 'MANUAL';
      const sourceId = String(item.sourceId || '').trim();
      const category = String(item.category || '').trim() || (sourceType === 'CONTACT_NEXT' ? '컨택예정' : '수동');
      return [
        String(item.id || '').trim() || Utilities.getUuid().slice(0, 8),
        selectedDate,
        idx + 1,
        content,
        item.done ? 'Y' : '',
        String(item.author || '').trim() || user,
        now,
        '',
        category,
        sourceType,
        sourceId,
        String(item.customerNo || '').trim(),
        String(item.company || '').trim(),
        String(item.rowNo || '').trim(),
        String(item.timeText || item.nextActionAt || '').trim(),
        String(item.priority || '').trim()
      ];
    })
    .filter(Boolean);

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headerLen).setValues(rows);
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

function buildPortalTodayUnifiedTasks_(storedTodos, nextActions, selectedDate) {
  storedTodos = Array.isArray(storedTodos) ? storedTodos : [];
  nextActions = Array.isArray(nextActions) ? nextActions : [];
  const result = [];
  const sourceTodoMap = {};

  storedTodos.forEach(function(todo) {
    const sourceType = String(todo.sourceType || '').trim();
    const sourceId = String(todo.sourceId || '').trim();
    if (sourceType && sourceType !== 'MANUAL' && sourceId) {
      sourceTodoMap[sourceType + '|' + sourceId] = todo;
    } else {
      result.push(normalizePortalTodayTask_(todo, selectedDate));
    }
  });

  nextActions.forEach(function(action) {
    const sourceId = String(action.sourceId || action.historyId || '').trim() || [action.customerNo, action.rowNo, action.nextActionAt, action.nextAction].join('|');
    const sourceKey = 'CONTACT_NEXT|' + sourceId;
    const persisted = sourceTodoMap[sourceKey];
    const defaultContent = buildPortalTodayContactTaskText_(action);
    if (persisted) {
      const merged = Object.assign({}, persisted, {
        category: persisted.category || '컨택예정',
        sourceType: 'CONTACT_NEXT',
        sourceId: sourceId,
        customerNo: action.customerNo || persisted.customerNo || '',
        company: action.company || persisted.company || '',
        rowNo: action.rowNo || persisted.rowNo || '',
        timeText: action.nextActionAt || persisted.timeText || '',
        nextAction: action.nextAction || '',
        rawContent: action.content || ''
      });
      if (!String(merged.content || '').trim()) merged.content = defaultContent;
      result.push(normalizePortalTodayTask_(merged, selectedDate));
    } else {
      result.push(normalizePortalTodayTask_({
        id: 'contact_' + Utilities.base64EncodeWebSafe(sourceId).slice(0, 16),
        date: selectedDate,
        order: 9000 + result.length,
        content: defaultContent,
        done: false,
        author: action.author || '',
        category: '컨택예정',
        sourceType: 'CONTACT_NEXT',
        sourceId: sourceId,
        customerNo: action.customerNo || '',
        company: action.company || '',
        rowNo: action.rowNo || '',
        timeText: action.nextActionAt || '',
        priority: '오늘',
        nextAction: action.nextAction || '',
        rawContent: action.content || ''
      }, selectedDate));
    }
  });

  result.sort(function(a, b) {
    const ad = a.done ? 1 : 0;
    const bd = b.done ? 1 : 0;
    if (ad !== bd) return ad - bd;
    const at = String(a.timeText || '99:99');
    const bt = String(b.timeText || '99:99');
    const ac = a.category === '컨택예정' ? 0 : 1;
    const bc = b.category === '컨택예정' ? 0 : 1;
    if (at !== bt) return at.localeCompare(bt);
    if (ac !== bc) return ac - bc;
    return (Number(a.order) || 0) - (Number(b.order) || 0);
  });

  return result;
}

function normalizePortalTodayTask_(task, selectedDate) {
  task = task || {};
  const sourceType = String(task.sourceType || '').trim() || 'MANUAL';
  const sourceId = String(task.sourceId || '').trim();
  return {
    id: String(task.id || '').trim() || Utilities.getUuid().slice(0, 8),
    date: normalizePortalTodoDate_(task.date || selectedDate || new Date()),
    order: Number(task.order) || 0,
    content: String(task.content || '').trim(),
    done: !!task.done,
    author: String(task.author || '').trim(),
    updatedAt: String(task.updatedAt || '').trim(),
    deleted: String(task.deleted || '').trim(),
    category: String(task.category || '').trim() || (sourceType === 'CONTACT_NEXT' ? '컨택예정' : '수동'),
    sourceType: sourceType,
    sourceId: sourceId,
    customerNo: String(task.customerNo || '').trim(),
    company: String(task.company || '').trim(),
    rowNo: String(task.rowNo || '').trim(),
    timeText: String(task.timeText || task.nextActionAt || '').trim(),
    priority: String(task.priority || '').trim(),
    nextAction: String(task.nextAction || '').trim(),
    rawContent: String(task.rawContent || '').trim()
  };
}

function buildPortalTodayContactTaskText_(action) {
  action = action || {};
  const parts = [];
  if (action.company) parts.push('[' + action.company + ']');
  if (action.nextAction) parts.push(action.nextAction);
  if (action.nextActionAt) parts.push(String(action.nextActionAt).replace(/^\d{4}-\d{1,2}-\d{1,2}\s*/, ''));
  if (action.content) parts.push('- ' + action.content);
  return parts.join(' ').trim() || '컨택 예정 확인';
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

  const lastCol = Math.max(sheet.getLastColumn(), PORTAL_CONFIG.TODAY_HEADERS.length);
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const map = {};
  headers.forEach(function(h, i) { if (h) map[h] = i; });
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return values
    .map(function(row) {
      const item = {
        id: String(row[map['할일ID']] || '').trim(),
        date: normalizePortalTodoDate_(row[map['일자']]),
        order: Number(row[map['순번']]) || 0,
        content: String(row[map['내용']] || '').trim(),
        done: String(row[map['완료여부']] || '').toUpperCase() === 'Y',
        author: String(row[map['작성자']] || '').trim(),
        updatedAt: formatDateTimeText_(row[map['수정일시']]),
        deleted: String(row[map['삭제여부']] || '').trim(),
        category: String(row[map['분류']] || '').trim(),
        sourceType: String(row[map['출처유형']] || '').trim() || 'MANUAL',
        sourceId: String(row[map['출처ID']] || '').trim(),
        customerNo: String(row[map['고객번호']] || '').trim(),
        company: String(row[map['회사명']] || '').trim(),
        rowNo: String(row[map['마스터행']] || '').trim(),
        timeText: String(row[map['시간']] || '').trim(),
        priority: String(row[map['우선순위']] || '').trim()
      };
      if (!item.category) item.category = item.sourceType === 'CONTACT_NEXT' ? '컨택예정' : '수동';
      return item;
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
  const canViewAll = !!(perm && (perm.canUseAdminHome || perm.canViewAllActivityLogs || perm.canReadAllSupport));
  const myNames = [];
  if (perm) {
    myNames.push(perm.name, perm.salesRepName);
    (perm.salesRepAliases || []).forEach(function(v) { myNames.push(v); });
  }
  const myNameKeys = myNames
    .map(function(v) { return normalizePortalNameForPermission_ ? normalizePortalNameForPermission_(v) : String(v || '').replace(/\s+/g, ''); })
    .filter(Boolean);

  return values.map(function(row, idx) {
      const nextAt = cellByHeader_(row, map, '다음액션일시') || cellByHeader_(row, map, '다음연락일');
      const nextDate = normalizePortalTodoDate_(nextAt);
      const historyId = cellByHeader_(row, map, '이력ID') || ('row' + (idx + 2));
      const customerNo = cellByHeader_(row, map, '고객번호');
      const rowNo = cellByHeader_(row, map, '마스터행');
      const nextAction = cellByHeader_(row, map, '다음액션');
      return {
        historyId: historyId,
        sourceId: historyId || [customerNo, rowNo, nextAt, nextAction].join('|'),
        customerNo: customerNo,
        company: cellByHeader_(row, map, '회사명'),
        rowNo: rowNo,
        nextAction: nextAction,
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
