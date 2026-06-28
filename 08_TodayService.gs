/***************************************
 * S1 Sales Portal - 08_TodayService.gs
 * 오늘 할 일 태그형 개편
 * - 수동/컨택 예정 액션을 하나의 할 일 목록으로 통합
 * - 사용자 정의 해시태그 다중 적용
 * - 작성시간/처리내용/완료예정시간/완료시간 관리
 ***************************************/

function getPortalTodayData(dateText, options) {
  options = options || {};
  const selectedDate = normalizePortalTodoDate_(dateText || new Date());
  const access = getPortalTodayAccessContextP350_(options);
  try { carryOverPortalTodayOpenTasksP350_(selectedDate, access); } catch (err) { console.warn('오늘할일 이월 처리 실패', err); }
  const storedTodos = getPortalTodosForDate_(selectedDate, access);
  let nextActions = getContactNextActionsForDate_(selectedDate);
  nextActions = filterPortalTodayItemsByAccessP350_(nextActions, access);
  const tasks = buildPortalTodayUnifiedTasks_(storedTodos, nextActions, selectedDate);
  const dateMeta = getPortalTodayDateMetaP130_(selectedDate, storedTodos);
  return {
    ok: true,
    selectedDate: selectedDate,
    dateVersion: dateMeta.version,
    activeTaskIds: tasks.map(function(t) { return String((t && t.id) || '').trim(); }).filter(Boolean),
    storedTaskIds: dateMeta.activeTaskIds,
    tasks: tasks,
    todos: tasks,
    nextActions: nextActions,
    tagOptions: getPortalTodayTagOptions_(),
    hiddenTags: getPortalTodayHiddenTags_(),
    actionOptions: PORTAL_NEXT_ACTION_OPTIONS,
    canViewAllTodos: !!access.canViewAll,
    authorFilter: access.authorFilter || 'ALL',
    currentUserLabel: access.currentUserLabel || '',
    authorOptions: access.canViewAll ? getPortalTodayAuthorOptionsP350_(selectedDate) : []
  };
}

function getPortalTodayAccessContextP350_(options) {
  options = options || {};
  const perm = getPortalCurrentPermission_ ? getPortalCurrentPermission_() : null;
  const canViewAll = !!(perm && perm.active !== false && (perm.canUseAdminHome || perm.canReadAllSupport));
  const currentUserLabel = String((perm && (perm.displayName || perm.name || perm.salesRepName)) || getCurrentUserLabel_() || '').trim();
  const myNames = [];
  if (perm) {
    myNames.push(perm.name, perm.displayName, perm.salesRepName, perm.email);
    (perm.salesRepAliases || []).forEach(function(v) { myNames.push(v); });
  }
  myNames.push(currentUserLabel);
  const myKeys = myNames.map(portalTodayAuthorKeyP350_).filter(Boolean);
  let authorFilter = String(options.assigneeFilter || options.authorFilter || options.author || 'ALL').trim();
  if (!authorFilter) authorFilter = 'ALL';
  const authorFilterKey = (canViewAll && authorFilter !== 'ALL') ? portalTodayAuthorKeyP350_(authorFilter) : '';
  return {
    perm: perm,
    canViewAll: canViewAll,
    currentUserLabel: currentUserLabel,
    myKeys: myKeys,
    authorFilter: canViewAll ? authorFilter : currentUserLabel,
    authorFilterKey: authorFilterKey
  };
}

function portalTodayAuthorKeyP350_(value) {
  if (typeof normalizePortalNameForPermission_ === 'function') return normalizePortalNameForPermission_(value);
  return String(value || '').replace(/\s+/g, '').toLowerCase();
}

function portalTodayAuthorMatchesAccessP350_(author, access) {
  access = access || getPortalTodayAccessContextP350_({});
  const key = portalTodayAuthorKeyP350_(author);
  if (access.canViewAll) {
    if (!access.authorFilterKey) return true;
    return key && key === access.authorFilterKey;
  }
  if (!access.myKeys || !access.myKeys.length) return true;
  return !key || access.myKeys.indexOf(key) >= 0;
}

function filterPortalTodayItemsByAccessP350_(items, access) {
  return (Array.isArray(items) ? items : []).filter(function(item) {
    return portalTodayAuthorMatchesAccessP350_((item && item.author) || '', access);
  });
}

function getPortalTodayAuthorOptionsP350_(selectedDate) {
  const seen = {};
  const out = [];
  function add(name) {
    name = String(name || '').trim();
    if (!name) return;
    const key = portalTodayAuthorKeyP350_(name);
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push(name);
  }
  try {
    const sheet = ensurePortalTodaySheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const headerLen = PORTAL_CONFIG.TODAY_HEADERS.length;
      const values = sheet.getRange(2, 1, lastRow - 1, headerLen).getValues();
      values.forEach(function(row) {
        const item = portalTodayRowToItemP130_(row);
        if (item.date === selectedDate && String(item.deleted || '').toUpperCase() !== 'Y' && item.content) add(item.author);
      });
    }
  } catch (err) {}
  try { getContactNextActionsForDate_(selectedDate).forEach(function(item) { add(item.author); }); } catch (err) {}
  return out.sort(function(a, b) { return String(a).localeCompare(String(b), 'ko'); });
}

function carryOverPortalTodayOpenTasksP350_(selectedDate, access) {
  selectedDate = normalizePortalTodoDate_(selectedDate || new Date());
  const todayStr = normalizePortalTodoDate_(new Date());
  if (String(selectedDate) < String(todayStr)) return { moved: 0 };
  const sheet = ensurePortalTodaySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { moved: 0 };
  const headerLen = PORTAL_CONFIG.TODAY_HEADERS.length;
  const values = sheet.getRange(2, 1, lastRow - 1, headerLen).getValues();
  const selectedActiveIds = {};
  values.forEach(function(row) {
    const item = portalTodayRowToItemP130_(row);
    if (item.date === selectedDate && String(item.deleted || '').toUpperCase() !== 'Y' && item.content) selectedActiveIds[item.id] = true;
  });
  const moveRows = [];
  const deleteRows = [];
  values.forEach(function(row, idx) {
    const item = portalTodayRowToItemP130_(row);
    if (!item.id || !item.content) return;
    if (String(item.deleted || '').toUpperCase() === 'Y') return;
    if (item.done) return;
    if (!portalTodayAuthorMatchesAccessP350_(item.author, access)) return;
    if (!(item.date && String(item.date) < String(todayStr))) return;
    const rowNo = idx + 2;
    if (selectedActiveIds[item.id]) deleteRows.push(rowNo);
    else moveRows.push(rowNo);
  });
  if (!moveRows.length && !deleteRows.length) return { moved: 0, deleted: 0 };
  const now = new Date();
  if (moveRows.length) {
    sheet.getRangeList(moveRows.map(function(r) { return 'B' + r; })).setValue(selectedDate);
    sheet.getRangeList(moveRows.map(function(r) { return 'G' + r; })).setValue(now);
  }
  if (deleteRows.length) markPortalTodayRowsDeletedP130_(sheet, deleteRows);
  SpreadsheetApp.flush();
  return { moved: moveRows.length, deleted: deleteRows.length };
}

function savePortalTodosForDate(payload) {
  payload = payload || {};
  const writeMeta = runPortalTodayWriteLockedP130_('today-save', function() {
    return savePortalTodosForDateCoreP130_(payload || {});
  });

  // 활동로그는 핵심 저장 Lock 밖에서 처리합니다. 로그 실패가 오늘할일 저장 성공을 막으면 안 됩니다.
  try {
    appendPortalActivityLog_({
      actionType: '오늘할일',
      screen: '오늘 할 일',
      summary: '오늘 할 일 저장: ' + writeMeta.selectedDate + ' / 저장 ' + writeMeta.upsertedCount + '건 / 삭제 ' + writeMeta.deletedCount + '건',
      detail: writeMeta
    });
  } catch (err) {}

  const data = getPortalTodayData(writeMeta.selectedDate);
  data.saved = true;
  data.saveMeta = writeMeta;
  data.mergedDueToStaleBase = !!writeMeta.mergedDueToStaleBase;
  return data;
}

function runPortalTodayWriteLockedP130_(label, callback) {
  if (typeof withPortalScriptLockP201_ === 'function') {
    return withPortalScriptLockP201_(label || 'today-save', callback, { attempts: 5, waitMs: 900, sleepBaseMs: 220 });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(12000);
  try {
    return callback();
  } finally {
    try { lock.releaseLock(); } catch (err) {}
  }
}

function savePortalTodosForDateCoreP130_(payload) {
  payload = payload || {};
  const selectedDate = normalizePortalTodoDate_(payload.date || new Date());
  const access = getPortalTodayAccessContextP350_(payload || {});
  const todos = Array.isArray(payload.tasks) ? payload.tasks : (Array.isArray(payload.todos) ? payload.todos : []);
  const sheet = ensurePortalTodaySheet_();
  const now = new Date();
  const user = getCurrentUserLabel_();
  const headerLen = PORTAL_CONFIG.TODAY_HEADERS.length;
  const existingRows = readPortalTodayRowsForDateP130_(sheet, selectedDate, headerLen, access);
  const currentMeta = getPortalTodayDateMetaFromRowsP130_(selectedDate, existingRows);
  const baseDateVersion = String(payload.baseDateVersion || '').trim();
  const mergedDueToStaleBase = !!(baseDateVersion && currentMeta.version && baseDateVersion !== currentMeta.version);
  const hasExplicitBaseIds = Array.isArray(payload.baseTaskIds);
  const baseTaskIds = (hasExplicitBaseIds ? payload.baseTaskIds : currentMeta.activeTaskIds)
    .map(function(v) { return String(v || '').trim(); })
    .filter(Boolean);
  const baseIdSet = baseTaskIds.reduce(function(acc, id) { acc[id] = true; return acc; }, {});

  const existingById = {};
  const duplicateDeleteRowNos = [];
  existingRows.forEach(function(rowInfo) {
    if (!rowInfo.active || !rowInfo.id) return;
    if (!existingById[rowInfo.id]) existingById[rowInfo.id] = rowInfo;
    else duplicateDeleteRowNos.push(rowInfo.rowNo);
  });

  const incomingIds = {};
  const appendRows = [];
  let upsertedCount = 0;
  todos.forEach(function(rawItem, idx) {
    const item = Object.assign({}, rawItem || {});
    if (!access.canViewAll) item.author = access.currentUserLabel || item.author || getCurrentUserLabel_();
    else if (!String(item.author || '').trim() && access.authorFilter && access.authorFilter !== 'ALL') item.author = access.authorFilter;
    const content = String(item.content || '').trim();
    if (!content) return;
    const id = String(item.id || '').trim() || Utilities.getUuid().slice(0, 8);
    incomingIds[id] = true;
    const prevRowInfo = existingById[id] || null;
    const prev = prevRowInfo ? prevRowInfo.item : {};
    const rowValues = buildPortalTodayRowValuesP130_(item, {
      id: id,
      selectedDate: selectedDate,
      order: idx + 1,
      prev: prev,
      now: now,
      user: user
    });
    if (prevRowInfo && prevRowInfo.rowNo) {
      sheet.getRange(prevRowInfo.rowNo, 1, 1, headerLen).setValues([rowValues]);
    } else {
      appendRows.push(rowValues);
    }
    upsertedCount++;
  });

  const deleteRowNos = duplicateDeleteRowNos.slice();
  existingRows.forEach(function(rowInfo) {
    if (!rowInfo.active || !rowInfo.id) return;
    if (baseIdSet[rowInfo.id] && !incomingIds[rowInfo.id]) deleteRowNos.push(rowInfo.rowNo);
  });
  markPortalTodayRowsDeletedP130_(sheet, deleteRowNos);

  if (appendRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, headerLen).setValues(appendRows);
  }

  SpreadsheetApp.flush();
  return {
    ok: true,
    selectedDate: selectedDate,
    baseDateVersion: baseDateVersion,
    previousDateVersion: currentMeta.version,
    mergedDueToStaleBase: mergedDueToStaleBase,
    explicitBaseIds: hasExplicitBaseIds,
    baseTaskCount: baseTaskIds.length,
    upsertedCount: upsertedCount,
    appendedCount: appendRows.length,
    updatedCount: upsertedCount - appendRows.length,
    deletedCount: deleteRowNos.length
  };
}

function buildPortalTodayRowValuesP130_(item, context) {
  item = item || {};
  context = context || {};
  const now = context.now || new Date();
  const prev = context.prev || {};
  const sourceType = String(item.sourceType || '').trim() || 'MANUAL';
  const done = !!item.done;
  const createdAt = parsePortalTodayDateTime_(item.createdAt) || parsePortalTodayDateTime_(prev.createdAt) || now;
  const completedAt = done
    ? (parsePortalTodayDateTime_(item.completedAt) || parsePortalTodayDateTime_(prev.completedAt) || now)
    : null;
  const dueAt = done ? '' : String(item.dueAt || item.timeText || '').trim();
  return [
    String(context.id || item.id || '').trim() || Utilities.getUuid().slice(0, 8),
    context.selectedDate,
    Number(context.order) || 0,
    String(item.content || '').trim(),
    done ? 'Y' : '',
    String(item.author || '').trim() || String(prev.author || '').trim() || context.user || getCurrentUserLabel_(),
    now,
    '',
    String(item.category || '').trim() || (sourceType === 'CONTACT_NEXT' ? '다음액션' : '할일'),
    sourceType,
    String(item.sourceId || '').trim(),
    String(item.customerNo || '').trim(),
    String(item.company || '').trim(),
    String(item.rowNo || '').trim(),
    String(item.timeText || item.nextActionAt || '').trim(),
    String(item.priority || '').trim(),
    createdAt,
    normalizePortalTodayTags_(item.tags).join(', '),
    String(item.detail || '').trim(),
    dueAt,
    completedAt || ''
  ];
}

function readPortalTodayRowsForDateP130_(sheet, selectedDate, headerLen, access) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, headerLen).getValues();
  return values.map(function(row, i) {
      const item = portalTodayRowToItemP130_(row);
      const deleted = String(row[7] || '').toUpperCase() === 'Y';
      return {
        rowNo: i + 2,
        values: row,
        id: item.id,
        date: item.date,
        deleted: deleted,
        active: item.date === selectedDate && !deleted && !!item.content && portalTodayAuthorMatchesAccessP350_(item.author, access),
        item: item
      };
    })
    .filter(function(rowInfo) { return rowInfo.date === selectedDate; });
}

function portalTodayRowToItemP130_(row) {
  row = row || [];
  return {
    id: String(row[0] || '').trim(),
    date: normalizePortalTodoDate_(row[1]),
    order: Number(row[2]) || 0,
    content: String(row[3] || '').trim(),
    done: String(row[4] || '').toUpperCase() === 'Y',
    author: String(row[5] || '').trim(),
    updatedAt: formatDateTimeText_(row[6]),
    deleted: String(row[7] || '').trim(),
    category: String(row[8] || '').trim(),
    sourceType: String(row[9] || '').trim() || 'MANUAL',
    sourceId: String(row[10] || '').trim(),
    customerNo: String(row[11] || '').trim(),
    company: String(row[12] || '').trim(),
    rowNo: String(row[13] || '').trim(),
    timeText: String(row[14] || '').trim(),
    priority: String(row[15] || '').trim(),
    createdAt: formatPortalTodayInputDateTime_(row[16] || row[6] || ''),
    tags: normalizePortalTodayTags_(row[17]),
    detail: String(row[18] || '').trim(),
    dueAt: formatPortalTodayInputDateTime_(row[19] || ''),
    completedAt: formatPortalTodayInputDateTime_(row[20] || '')
  };
}

function markPortalTodayRowsDeletedP130_(sheet, rowNos) {
  const unique = {};
  const a1s = [];
  (Array.isArray(rowNos) ? rowNos : []).forEach(function(rowNo) {
    rowNo = Number(rowNo) || 0;
    if (rowNo < 2 || unique[rowNo]) return;
    unique[rowNo] = true;
    a1s.push('H' + rowNo);
  });
  if (a1s.length) sheet.getRangeList(a1s).setValue('Y');
}

function getPortalTodayDateMetaP130_(selectedDate, storedTodos) {
  const rows = (Array.isArray(storedTodos) ? storedTodos : getPortalTodosForDate_(selectedDate))
    .map(function(item) { return { active: true, id: String((item && item.id) || '').trim(), item: item || {} }; });
  return getPortalTodayDateMetaFromRowsP130_(selectedDate, rows);
}

function getPortalTodayDateMetaFromRowsP130_(selectedDate, rows) {
  const active = (Array.isArray(rows) ? rows : [])
    .filter(function(rowInfo) { return rowInfo && rowInfo.active && rowInfo.id; })
    .map(function(rowInfo) {
      const item = rowInfo.item || {};
      return {
        id: String(rowInfo.id || item.id || '').trim(),
        updatedAt: String(item.updatedAt || '').trim(),
        done: item.done ? 'Y' : '',
        content: String(item.content || '').trim()
      };
    })
    .filter(function(item) { return item.id; })
    .sort(function(a, b) { return a.id.localeCompare(b.id); });
  const source = selectedDate + '|' + active.map(function(item) {
    return [item.id, item.updatedAt, item.done, item.content].join(':');
  }).join('|');
  return {
    version: digestPortalTodayVersionP130_(source),
    activeTaskIds: active.map(function(item) { return item.id; })
  };
}

function digestPortalTodayVersionP130_(source) {
  source = String(source || '');
  try {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, source, Utilities.Charset.UTF_8);
    return bytes.map(function(b) {
      b = b < 0 ? b + 256 : b;
      return ('0' + b.toString(16)).slice(-2);
    }).join('');
  } catch (err) {
    return String(source.length) + '_' + Utilities.base64EncodeWebSafe(source).slice(0, 24);
  }
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
    }
    result.push(normalizePortalTodayTask_(todo, selectedDate));
  });

  nextActions.forEach(function(action) {
    const sourceId = String(action.sourceId || action.historyId || '').trim() || [action.customerNo, action.rowNo, action.nextActionAt, action.nextAction].join('|');
    const sourceKey = 'CONTACT_NEXT|' + sourceId;
    if (sourceTodoMap[sourceKey]) return;
    result.push(normalizePortalTodayTask_({
      id: 'contact_' + Utilities.base64EncodeWebSafe(sourceId).slice(0, 16),
      date: selectedDate,
      order: 9000 + result.length,
      content: buildPortalTodayContactTaskText_(action),
      done: false,
      author: action.author || '',
      category: '컨택예정',
      sourceType: 'CONTACT_NEXT',
      sourceId: sourceId,
      customerNo: action.customerNo || '',
      company: action.company || '',
      rowNo: action.rowNo || '',
      timeText: action.nextActionAt || '',
      dueAt: action.nextActionAt || '',
      priority: '오늘',
      tags: normalizePortalNextActionTags_(action.nextActionTags || action.nextAction || ''),
      nextAction: action.nextAction || '',
      rawContent: action.content || ''
    }, selectedDate));
  });

  result.sort(function(a, b) {
    const ad = a.done ? 1 : 0;
    const bd = b.done ? 1 : 0;
    if (ad !== bd) return ad - bd;
    const at = String(a.dueAt || a.timeText || '99:99');
    const bt = String(b.dueAt || b.timeText || '99:99');
    if (at !== bt) return at.localeCompare(bt);
    return (Number(a.order) || 0) - (Number(b.order) || 0);
  });

  return result;
}

function normalizePortalTodayTask_(task, selectedDate) {
  task = task || {};
  const sourceType = String(task.sourceType || '').trim() || 'MANUAL';
  const createdAt = task.createdAt || task.updatedAt || new Date();
  const tags = normalizePortalTodayTags_(task.tags);
  return {
    id: String(task.id || '').trim() || Utilities.getUuid().slice(0, 8),
    date: normalizePortalTodoDate_(task.date || selectedDate || new Date()),
    order: Number(task.order) || 0,
    content: String(task.content || '').trim(),
    done: !!task.done,
    author: String(task.author || '').trim(),
    updatedAt: formatDateTimeText_(task.updatedAt || ''),
    deleted: String(task.deleted || '').trim(),
    category: String(task.category || '').trim() || (sourceType === 'CONTACT_NEXT' ? '다음액션' : '할일'),
    sourceType: sourceType,
    sourceId: String(task.sourceId || '').trim(),
    customerNo: String(task.customerNo || '').trim(),
    company: String(task.company || '').trim(),
    rowNo: String(task.rowNo || '').trim(),
    timeText: String(task.timeText || task.nextActionAt || '').trim(),
    priority: String(task.priority || '').trim(),
    tags: tags,
    detail: String(task.detail || '').trim(),
    dueAt: formatPortalTodayInputDateTime_(task.dueAt || ''),
    completedAt: formatPortalTodayInputDateTime_(task.completedAt || ''),
    createdAt: formatPortalTodayInputDateTime_(createdAt),
    nextAction: String(task.nextAction || '').trim(),
    rawContent: String(task.rawContent || '').trim()
  };
}

function normalizePortalNextActionTags_(value) {
  const tags = normalizePortalTodayTags_(value);
  const text = String(value || '').trim();
  function add(tag) { if (tag && tags.indexOf(tag) < 0) tags.push(tag); }
  add('컨택');
  if (/전화|통화|콜|연락/.test(text)) add('전화');
  if (/메일|자료|발송|양식/.test(text)) add('메일');
  if (/견적/.test(text)) add('견적');
  if (/계약|서류/.test(text)) add('계약');
  if (/재확인|확인|보류/.test(text)) add('재확인');
  return tags;
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

function getPortalTodosForDate_(selectedDate, access) {
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
        priority: String(row[map['우선순위']] || '').trim(),
        createdAt: formatPortalTodayInputDateTime_(row[map['작성일시']] || row[map['수정일시']] || ''),
        tags: normalizePortalTodayTags_(row[map['태그']]),
        detail: String(row[map['처리내용']] || '').trim(),
        dueAt: formatPortalTodayInputDateTime_(row[map['완료예정시간']] || ''),
        completedAt: formatPortalTodayInputDateTime_(row[map['완료시간']] || '')
      };
      if (!item.category) item.category = item.sourceType === 'CONTACT_NEXT' ? '다음액션' : '할일';
      return item;
    })
    .filter(function(item) {
      return item.date === selectedDate && item.deleted.toUpperCase() !== 'Y' && item.content && portalTodayAuthorMatchesAccessP350_(item.author, access);
    })
    .sort(function(a, b) { return (a.order || 0) - (b.order || 0); });
}

const PORTAL_TODAY_HIDDEN_TAGS_PREF_KEY_P123 = 'todayHiddenTags';

function getPortalTodayHiddenTags_() {
  try {
    const perm = getPortalCurrentPermission_ ? getPortalCurrentPermission_() : null;
    const email = String((perm && perm.email) || getPortalActiveUserEmail_() || '').trim().toLowerCase();
    if (!email || typeof ensurePortalUserPrefSheetP121_ !== 'function' || typeof readPortalUserPrefMapP121_ !== 'function') return [];
    const map = readPortalUserPrefMapP121_(ensurePortalUserPrefSheetP121_(), email);
    return normalizePortalTodayTags_(map[PORTAL_TODAY_HIDDEN_TAGS_PREF_KEY_P123] || '');
  } catch (err) {
    return [];
  }
}

function savePortalTodayHiddenTags(tags) {
  const hidden = normalizePortalTodayTags_(tags || '');
  const perm = getPortalCurrentPermission_ ? getPortalCurrentPermission_() : null;
  const email = String((perm && perm.email) || getPortalActiveUserEmail_() || '').trim().toLowerCase();
  if (!email) throw new Error('현재 접속자 이메일을 확인하지 못해 태그 설정을 저장할 수 없습니다.');
  const displayName = (perm && (perm.displayName || perm.name)) || email || '웹앱사용자';
  if (typeof ensurePortalUserPrefSheetP121_ === 'function' && typeof upsertPortalUserPrefP121_ === 'function') {
    upsertPortalUserPrefP121_(ensurePortalUserPrefSheetP121_(), email, displayName, PORTAL_TODAY_HIDDEN_TAGS_PREF_KEY_P123, hidden.join(', '));
  }
  return { ok: true, hiddenTags: hidden, tagOptions: getPortalTodayTagOptions_() };
}

function getPortalTodayTagOptions_() {
  const defaults = Array.isArray(PORTAL_CONFIG.TODAY_DEFAULT_TAGS) ? PORTAL_CONFIG.TODAY_DEFAULT_TAGS : [];
  const hidden = getPortalTodayHiddenTags_().reduce(function(acc, tag) { acc[tag] = true; return acc; }, {});
  const seen = {};
  const tags = [];
  function add(v) {
    v = String(v || '').replace(/^#+/, '').trim();
    if (!v || seen[v] || hidden[v]) return;
    seen[v] = true;
    tags.push(v);
  }
  defaults.forEach(add);
  try {
    const sheet = ensurePortalTodaySheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const lastCol = Math.max(sheet.getLastColumn(), PORTAL_CONFIG.TODAY_HEADERS.length);
      const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
      const idx = headers.indexOf('태그');
      if (idx >= 0) {
        const vals = sheet.getRange(2, idx + 1, lastRow - 1, 1).getDisplayValues();
        vals.forEach(function(r) { normalizePortalTodayTags_(r[0]).forEach(add); });
      }
    }
  } catch (err) {}
  return tags.slice(0, 30);
}

function normalizePortalTodayTags_(value) {
  let arr = [];
  if (Array.isArray(value)) arr = value;
  else arr = String(value || '').split(/[#,，,;；\s]+/);
  const seen = {};
  return arr.map(function(v) { return String(v || '').replace(/^#+/, '').trim(); })
    .filter(function(v) { if (!v || seen[v]) return false; seen[v] = true; return true; });
}

function parsePortalTodayDateTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const text = String(value || '').trim();
  if (!text) return null;
  const local = text.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
  if (local) return new Date(Number(local[1]), Number(local[2]) - 1, Number(local[3]), Number(local[4]), Number(local[5]));
  const dot = text.match(/(\d{2,4})[.\/-](\d{1,2})[.\/-](\d{1,2})\s+(\d{1,2}):(\d{2})/);
  if (dot) {
    const y = String(dot[1]).length === 2 ? 2000 + Number(dot[1]) : Number(dot[1]);
    return new Date(y, Number(dot[2]) - 1, Number(dot[3]), Number(dot[4]), Number(dot[5]));
  }
  const d = new Date(text);
  return isNaN(d.getTime()) ? null : d;
}

function formatPortalTodayInputDateTime_(value) {
  const d = parsePortalTodayDateTime_(value);
  if (!d) return '';
  const pad = function(n) { return String(n).padStart(2, '0'); };
  return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + 'T' + pad(d.getHours()) + ':' + pad(d.getMinutes());
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
        author: cellByHeader_(row, map, '작성자'),
        nextActionTags: cellByHeader_(row, map, '다음액션태그')
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
