/***************************************
 * S1 Sales Portal - 08_TodayService.gs
 * 오늘 할 일 운영 안정화 버전
 * - 원천: 오늘할일_DB
 * - 추가/수정/삭제는 클라이언트에서 모은 현재 목록을 [저장] 시 한 번에 반영
 * - 미완료 과거 할 일은 오늘 화면 진입 시 오늘 날짜로 자동 이월
 * - 서무/admin은 전체/담당자별 조회, 일반 사용자는 본인 할 일만 조회
 ***************************************/

function getPortalTodayData(dateText, options) {
  options = options || {};
  const selectedDate = normalizePortalTodoDate_(dateText || new Date());
  const access = getPortalTodayAccessContextP360_(options);

  if (!options.skipCarryover && selectedDate === normalizePortalTodoDate_(new Date())) {
    try { carryOverPortalTodayOpenTasksCachedP370_(selectedDate, access); } catch (err) { console.warn('오늘할일 이월 처리 실패', err); }
  }

  const sheet = ensurePortalTodaySheet_();
  const headerLen = PORTAL_CONFIG.TODAY_HEADERS.length;
  const allDateRows = readPortalTodaySheetRowsForDateAllP370_(sheet, selectedDate, headerLen, { includeDeleted: true });
  const accessRows = allDateRows.filter(function(rowInfo) {
    return portalTodayAuthorMatchesAccessP360_((rowInfo.item && rowInfo.item.author) || '', access);
  });

  const storedTodos = accessRows
    .filter(function(rowInfo) { return !rowInfo.deleted && rowInfo.item && rowInfo.item.content; })
    .map(function(rowInfo) { return rowInfo.item; })
    .sort(function(a, b) { return (Number(a.order) || 0) - (Number(b.order) || 0); });

  const deletedSourceMap = {};
  accessRows.forEach(function(rowInfo) {
    const item = rowInfo.item || {};
    if (String(item.deleted || '').toUpperCase() !== 'Y') return;
    if (item.id) deletedSourceMap['ID|' + item.id] = true;
    if (item.sourceType && item.sourceId) deletedSourceMap[String(item.sourceType) + '|' + String(item.sourceId)] = true;
  });

  const rawNextActions = options.skipNextActions ? [] : getContactNextActionsRawForDateP370_(selectedDate);
  const nextActions = filterPortalTodayItemsByAccessP360_(rawNextActions, access);
  const tasks = buildPortalTodayUnifiedTasksP360_(storedTodos, nextActions, selectedDate, deletedSourceMap);
  const dateMeta = getPortalTodayDateMetaP360_(selectedDate, tasks);

  const authorOptions = access.canViewAll ? buildPortalTodayAuthorOptionsFromRowsP370_(allDateRows, rawNextActions) : [];

  return {
    ok: true,
    selectedDate: selectedDate,
    dateVersion: dateMeta.version,
    activeTaskIds: dateMeta.activeTaskIds,
    tasks: tasks,
    todos: tasks,
    nextActions: nextActions,
    tagOptions: getPortalTodayTagOptionsFastP370_(tasks),
    hiddenTags: getPortalTodayHiddenTags_(),
    actionOptions: typeof PORTAL_NEXT_ACTION_OPTIONS !== 'undefined' ? PORTAL_NEXT_ACTION_OPTIONS : [],
    canViewAllTodos: !!access.canViewAll,
    authorFilter: access.authorFilter || 'ALL',
    currentUserLabel: access.currentUserLabel || '',
    authorOptions: authorOptions,
    loadedAt: formatDateTimeText_(new Date())
  };
}


function readPortalTodayMatchingDateRowsFastP380_(sheet, selectedDate, headerLen) {
  selectedDate = normalizePortalTodoDate_(selectedDate || new Date());
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // 속도 개선: 오늘할일_DB 전체 A:U를 매번 읽지 않고, B열 날짜만 먼저 읽은 뒤
  // 해당 날짜의 행만 묶어서 가져옵니다. 데이터가 날짜별로 쌓일수록 차이가 큽니다.
  const dateValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const rowNos = [];
  dateValues.forEach(function(row, idx) {
    if (normalizePortalTodoDate_(row[0]) === selectedDate) rowNos.push(idx + 2);
  });
  if (!rowNos.length) return [];

  const groups = [];
  rowNos.forEach(function(rowNo) {
    const last = groups.length ? groups[groups.length - 1] : null;
    if (last && rowNo === last.end + 1) last.end = rowNo;
    else groups.push({ start: rowNo, end: rowNo });
  });

  const result = [];
  groups.forEach(function(g) {
    const values = sheet.getRange(g.start, 1, g.end - g.start + 1, headerLen).getValues();
    values.forEach(function(row, idx) {
      result.push({ rowNo: g.start + idx, values: row });
    });
  });
  return result;
}

function getPortalTodayAssignableAuthorOptionsP380_() {
  const cache = CacheService.getScriptCache();
  const bust = (typeof getPortalPermissionCacheBustP230_ === 'function') ? getPortalPermissionCacheBustP230_() : 'v1';
  const key = 'PORTAL_TODAY_AUTHORS_P380_' + bust;
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached) || []; } catch (err) {}
  }

  const seen = {};
  const out = [];
  function add(name) {
    name = String(name || '').trim();
    if (!name) return;
    const keyName = portalTodayAuthorKeyP360_(name);
    if (!keyName || seen[keyName]) return;
    seen[keyName] = true;
    out.push(name);
  }

  try {
    const sheet = ensurePortalPermissionSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 15)).getValues();
      values.forEach(function(row) {
        const active = String(row[14] || 'Y').trim().toUpperCase() !== 'N';
        if (!active) return;
        add(row[6] || row[1] || row[0]);
      });
    }
  } catch (err) {}

  out.sort(function(a, b) { return String(a).localeCompare(String(b), 'ko'); });
  try { cache.put(key, JSON.stringify(out).slice(0, 90000), 300); } catch (err) {}
  return out;
}

function readPortalTodaySheetRowsForDateAllP370_(sheet, selectedDate, headerLen, options) {
  options = options || {};
  const matchingRows = readPortalTodayMatchingDateRowsFastP380_(sheet, selectedDate, headerLen);
  return matchingRows.map(function(rowInfo) {
    const item = portalTodayRowToItemP360_(rowInfo.values);
    const deleted = String(item.deleted || '').toUpperCase() === 'Y';
    return {
      rowNo: rowInfo.rowNo,
      values: rowInfo.values,
      id: item.id,
      date: item.date,
      deleted: deleted,
      active: item.date === selectedDate && !deleted && !!item.content,
      access: true,
      item: item
    };
  }).filter(function(rowInfo) {
    if (rowInfo.date !== selectedDate) return false;
    if (options.includeDeleted) return true;
    return !rowInfo.deleted && !!rowInfo.item.content;
  });
}

function buildPortalTodayAuthorOptionsFromRowsP370_(allDateRows, rawNextActions) {
  const seen = {};
  const out = [];
  function add(name) {
    name = String(name || '').trim();
    if (!name) return;
    const key = portalTodayAuthorKeyP360_(name);
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push(name);
  }

  // 서무/admin 담당자 필터는 그날 할 일이 있는 사람만 나오면 안 됩니다.
  // 권한_DB의 활성 사용자 전체를 기본 옵션으로 깔고, 실제 데이터 작성자도 추가합니다.
  getPortalTodayAssignableAuthorOptionsP380_().forEach(add);
  (Array.isArray(allDateRows) ? allDateRows : []).forEach(function(rowInfo) {
    const item = (rowInfo && rowInfo.item) || {};
    if (String(item.deleted || '').toUpperCase() === 'Y') return;
    if (item.content) add(item.author);
  });
  (Array.isArray(rawNextActions) ? rawNextActions : []).forEach(function(item) { add(item && item.author); });
  return out.sort(function(a, b) { return String(a).localeCompare(String(b), 'ko'); });
}

function getPortalTodayTagOptionsFastP370_(tasks) {
  const defaults = Array.isArray(PORTAL_CONFIG.TODAY_DEFAULT_TAGS) ? PORTAL_CONFIG.TODAY_DEFAULT_TAGS : ['전화','메일','견적','컨택','계약','서류','영업지원','재확인','긴급','내부처리'];
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
  (Array.isArray(tasks) ? tasks : []).forEach(function(task) {
    normalizePortalTodayTags_((task && task.tags) || '').forEach(add);
  });
  return tags.slice(0, 30);
}

function carryOverPortalTodayOpenTasksCachedP370_(selectedDate, access) {
  const cache = CacheService.getScriptCache();
  const key = 'PORTAL_TODAY_CARRY_P370_' + selectedDate + '_' + (access && access.canViewAll ? String(access.authorFilter || 'ALL') : 'ME');
  if (cache.get(key)) return { moved: 0, cached: true };
  const result = carryOverPortalTodayOpenTasksP360_(selectedDate, access);
  cache.put(key, '1', 90);
  return result;
}

function getContactNextActionsRawForDateP370_(selectedDate) {
  const cache = CacheService.getScriptCache();
  const key = 'PORTAL_TODAY_NEXT_RAW_P370_' + selectedDate;
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached) || []; } catch (err) {}
  }
  const ss = getWebAppDbSpreadsheet_();
  const sheet = ss.getSheetByName(PORTAL_CONFIG.CONTACT_HISTORY_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const map = {};
  headers.forEach(function(h, i) { if (h) map[h] = i; });
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();

  const result = values.map(function(row, idx) {
      const nextAt = cellByHeader_(row, map, '다음액션일시') || cellByHeader_(row, map, '다음연락일');
      const nextDate = normalizePortalTodoDate_(nextAt);
      const historyId = cellByHeader_(row, map, '이력ID') || ('row' + (idx + 2));
      const customerNo = cellByHeader_(row, map, '고객번호');
      const rowNo = cellByHeader_(row, map, '마스터행');
      const nextAction = cellByHeader_(row, map, '다음액션');
      const nextActionAuthor = cellByHeader_(row, map, '다음액션담당자') || cellByHeader_(row, map, '작성자');
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
        author: nextActionAuthor,
        nextActionAuthor: nextActionAuthor,
        nextActionTags: cellByHeader_(row, map, '다음액션태그')
      };
    })
    .filter(function(item) { return item.nextDate === selectedDate && (item.nextAction || item.nextActionAt); })
    .sort(function(a, b) { return String(a.nextActionAt || '').localeCompare(String(b.nextActionAt || '')); });
  try { cache.put(key, JSON.stringify(result).slice(0, 95000), 120); } catch (err) {}
  return result;
}

function getPortalTodayAccessContextP360_(options) {
  options = options || {};
  const perm = getPortalCurrentPermission_ ? getPortalCurrentPermission_() : null;
  const canViewAll = !!(perm && perm.active !== false && perm.canUseAdminHome);
  const currentUserLabel = String((perm && (perm.displayName || perm.name || perm.salesRepName)) || getCurrentUserLabel_() || '').trim();
  const names = [];
  if (perm) {
    names.push(perm.name, perm.displayName, perm.salesRepName, perm.email);
    (perm.salesRepAliases || []).forEach(function(v) { names.push(v); });
  }
  names.push(currentUserLabel);
  const myKeys = names.map(portalTodayAuthorKeyP360_).filter(Boolean);
  let authorFilter = String(options.assigneeFilter || options.authorFilter || options.author || 'ALL').trim();
  if (!authorFilter) authorFilter = 'ALL';
  const authorFilterKey = (canViewAll && authorFilter !== 'ALL') ? portalTodayAuthorKeyP360_(authorFilter) : '';
  return {
    perm: perm,
    canViewAll: canViewAll,
    currentUserLabel: currentUserLabel,
    myKeys: myKeys,
    authorFilter: canViewAll ? authorFilter : currentUserLabel,
    authorFilterKey: authorFilterKey
  };
}

function portalTodayAuthorKeyP360_(value) {
  if (typeof normalizePortalNameForPermission_ === 'function') return normalizePortalNameForPermission_(value);
  return String(value || '').replace(/\s+/g, '').toLowerCase();
}

function portalTodayAuthorMatchesAccessP360_(author, access) {
  access = access || getPortalTodayAccessContextP360_({});
  const key = portalTodayAuthorKeyP360_(author);
  if (access.canViewAll) {
    if (!access.authorFilterKey) return true;
    return !!key && key === access.authorFilterKey;
  }
  if (!access.myKeys || !access.myKeys.length) return true;
  return !key || access.myKeys.indexOf(key) >= 0;
}

function filterPortalTodayItemsByAccessP360_(items, access) {
  return (Array.isArray(items) ? items : []).filter(function(item) {
    return portalTodayAuthorMatchesAccessP360_((item && item.author) || '', access);
  });
}

function getPortalTodayAuthorOptionsP360_(selectedDate) {
  try {
    const sheet = ensurePortalTodaySheet_();
    const allRows = readPortalTodaySheetRowsForDateAllP370_(sheet, selectedDate, PORTAL_CONFIG.TODAY_HEADERS.length, { includeDeleted: true });
    const rawNext = getContactNextActionsRawForDateP370_(selectedDate);
    return buildPortalTodayAuthorOptionsFromRowsP370_(allRows, rawNext);
  } catch (err) {
    return [];
  }
}

function carryOverPortalTodayOpenTasksP360_(selectedDate, access) {
  selectedDate = normalizePortalTodoDate_(selectedDate || new Date());
  const todayStr = normalizePortalTodoDate_(new Date());
  if (selectedDate !== todayStr) return { moved: 0, deleted: 0 };

  const sheet = ensurePortalTodaySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { moved: 0, deleted: 0 };

  // 이월 판정에는 A:H까지만 필요합니다. 전체 컬럼을 읽지 않아도 됩니다.
  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const rowsToMove = [];

  values.forEach(function(row, idx) {
    const item = portalTodayRowToItemP360_(row);
    if (!item.id || !item.content) return;
    if (String(item.deleted || '').toUpperCase() === 'Y') return;
    if (item.done) return;
    if (!item.date || item.date >= todayStr) return;
    if (!portalTodayAuthorMatchesAccessP360_(item.author, access)) return;
    rowsToMove.push(idx + 2);
  });

  if (!rowsToMove.length) return { moved: 0, deleted: 0 };
  const now = new Date();
  sheet.getRangeList(rowsToMove.map(function(r) { return 'B' + r; })).setValue(todayStr);
  sheet.getRangeList(rowsToMove.map(function(r) { return 'G' + r; })).setValue(now);
  if (!payload.noFlush) SpreadsheetApp.flush();
  return { moved: rowsToMove.length, deleted: 0 };
}

function savePortalTodosForDate(payload) {
  payload = payload || {};
  const writeMeta = runPortalTodayWriteLockedP360_('today-save', function() {
    return savePortalTodosForDateCoreP360_(payload || {});
  });

  if (!payload.skipActivityLog && !payload.fastMode) {
    try {
      appendPortalActivityLog_({
        actionType: '오늘할일',
        screen: '오늘 할 일',
        summary: '오늘 할 일 저장: ' + writeMeta.selectedDate + ' / 저장 ' + writeMeta.upsertedCount + '건 / 삭제 ' + writeMeta.deletedCount + '건',
        detail: writeMeta
      });
    } catch (err) {}
  }

  // STEP36: 저장 직후 다시 컨택이력/원장 전체를 조회하면 체감 저장 시간이 길어집니다.
  // 방금 저장한 payload를 기준으로 화면에 필요한 응답을 즉시 구성하고, 다음 메뉴 진입/백그라운드 갱신에서 서버 현재값을 맞춥니다.
  const data = buildPortalTodaySavedResponseFastP360_(payload, writeMeta);
  data.saved = true;
  data.saveMeta = writeMeta;
  return data;
}



function buildPortalTodaySavedResponseFastP360_(payload, writeMeta) {
  payload = payload || {};
  writeMeta = writeMeta || {};
  const selectedDate = normalizePortalTodoDate_(writeMeta.selectedDate || payload.date || new Date());
  const access = getPortalTodayAccessContextP360_(payload || {});
  const raw = Array.isArray(payload.tasks) ? payload.tasks : (Array.isArray(payload.todos) ? payload.todos : []);
  const deletedIds = {};
  (Array.isArray(payload.deletedTasks) ? payload.deletedTasks : []).forEach(function(t) {
    const id = String((t && t.id) || '').trim();
    if (id) deletedIds[id] = true;
  });
  const seen = {};
  const tasks = [];
  raw.forEach(function(item, idx) {
    let t = normalizePortalTodayTaskP360_(item || {}, selectedDate);
    if (!t.content || deletedIds[t.id] || seen[t.id]) return;
    if (!access.canViewAll) t.author = access.currentUserLabel || t.author || getCurrentUserLabel_();
    else if (!String(t.author || '').trim()) t.author = (access.authorFilter && access.authorFilter !== 'ALL') ? access.authorFilter : (access.currentUserLabel || getCurrentUserLabel_());
    t.order = idx + 1;
    t.updatedAt = formatDateTimeText_(new Date());
    seen[t.id] = true;
    tasks.push(t);
  });
  return {
    ok: true,
    selectedDate: selectedDate,
    dateVersion: 'saved-' + new Date().getTime(),
    activeTaskIds: tasks.map(function(t) { return t.id; }),
    tasks: tasks,
    todos: tasks,
    nextActions: [],
    tagOptions: getPortalTodayTagOptionsFastP370_(tasks),
    hiddenTags: getPortalTodayHiddenTags_(),
    actionOptions: typeof PORTAL_NEXT_ACTION_OPTIONS !== 'undefined' ? PORTAL_NEXT_ACTION_OPTIONS : [],
    canViewAllTodos: !!access.canViewAll,
    authorFilter: access.authorFilter || 'ALL',
    currentUserLabel: access.currentUserLabel || '',
    authorOptions: access.canViewAll ? getPortalTodayAssignableAuthorOptionsP380_() : [],
    loadedAt: formatDateTimeText_(new Date()),
    fastSavedResponse: true
  };
}

function runPortalTodayWriteLockedP360_(label, callback) {
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


function portalTodayRowValuesEquivalentP463_(a, b) {
  a = Array.isArray(a) ? a : [];
  b = Array.isArray(b) ? b : [];
  const len = Math.max(a.length, b.length);
  for (let i = 0; i < len; i++) {
    // updatedAt(G열, index 6)은 저장 시각이라 내용 비교에서 제외합니다.
    if (i === 6) continue;
    const av = a[i] instanceof Date ? formatDateTimeText_(a[i]) : String(a[i] == null ? '' : a[i]).trim();
    const bv = b[i] instanceof Date ? formatDateTimeText_(b[i]) : String(b[i] == null ? '' : b[i]).trim();
    if (av !== bv) return false;
  }
  return true;
}

function savePortalTodosForDateCoreP360_(payload) {
  payload = payload || {};
  const selectedDate = normalizePortalTodoDate_(payload.date || new Date());
  const access = getPortalTodayAccessContextP360_(payload || {});
  const incomingRaw = Array.isArray(payload.tasks) ? payload.tasks : (Array.isArray(payload.todos) ? payload.todos : []);
  const deletedRaw = Array.isArray(payload.deletedTasks) ? payload.deletedTasks : [];
  const baseTasks = Array.isArray(payload.baseTasks) ? payload.baseTasks : [];
  const baseTaskIds = (Array.isArray(payload.baseTaskIds) ? payload.baseTaskIds : baseTasks.map(function(t) { return t && t.id; }))
    .map(function(v) { return String(v || '').trim(); })
    .filter(Boolean);
  const baseIdSet = baseTaskIds.reduce(function(acc, id) { acc[id] = true; return acc; }, {});

  const sheet = ensurePortalTodaySheet_();
  const now = new Date();
  const user = getCurrentUserLabel_();
  const headerLen = PORTAL_CONFIG.TODAY_HEADERS.length;
  const allRowsForDate = readPortalTodaySheetRowsForDateP360_(sheet, selectedDate, headerLen, access, { includeDeleted: true });

  const existingActiveById = {};
  const existingAnyById = {};
  const duplicateDeleteRows = [];
  allRowsForDate.forEach(function(rowInfo) {
    if (!rowInfo.id) return;
    if (!existingAnyById[rowInfo.id]) existingAnyById[rowInfo.id] = rowInfo;
    if (!rowInfo.deleted && rowInfo.item.content) {
      if (!existingActiveById[rowInfo.id]) existingActiveById[rowInfo.id] = rowInfo;
      else duplicateDeleteRows.push(rowInfo.rowNo);
    }
  });

  const incomingIds = {};
  const appendRows = [];
  let upsertedCount = 0;

  incomingRaw.forEach(function(rawItem, idx) {
    const item = Object.assign({}, rawItem || {});
    item.content = String(item.content || '').trim();
    if (!item.content) return;
    if (!access.canViewAll) {
      item.author = access.currentUserLabel || item.author || user;
    } else if (!String(item.author || '').trim()) {
      item.author = (access.authorFilter && access.authorFilter !== 'ALL') ? access.authorFilter : (access.currentUserLabel || user);
    }
    const id = String(item.id || '').trim() || Utilities.getUuid().slice(0, 12);
    incomingIds[id] = true;
    const prevRowInfo = existingAnyById[id] || null;
    const prev = prevRowInfo ? prevRowInfo.item : {};
    const rowValues = buildPortalTodayRowValuesP360_(item, {
      id: id,
      selectedDate: selectedDate,
      order: idx + 1,
      prev: prev,
      now: now,
      user: user,
      deleted: ''
    });
    if (prevRowInfo && prevRowInfo.rowNo) {
      // P463: 기존값과 실질적으로 달라진 행만 씁니다. 이전에는 매 저장마다 모든 할 일을 다시 setValues 해서 느렸습니다.
      if (!portalTodayRowValuesEquivalentP463_(rowValues, prevRowInfo.values || [])) {
        sheet.getRange(prevRowInfo.rowNo, 1, 1, headerLen).setValues([rowValues]);
        upsertedCount++;
      }
    } else {
      appendRows.push(rowValues);
      upsertedCount++;
    }
  });

  const deleteRowNos = duplicateDeleteRows.slice();
  allRowsForDate.forEach(function(rowInfo) {
    if (rowInfo.deleted || !rowInfo.active || !rowInfo.id) return;
    if (baseIdSet[rowInfo.id] && !incomingIds[rowInfo.id]) deleteRowNos.push(rowInfo.rowNo);
  });

  const tombstoneRows = [];
  const explicitDeletedById = {};
  deletedRaw.forEach(function(raw) {
    const t = normalizePortalTodayTaskP360_(raw || {}, selectedDate);
    if (t.id) explicitDeletedById[t.id] = t;
  });
  Object.keys(explicitDeletedById).forEach(function(id) {
    const t = explicitDeletedById[id];
    const existing = existingAnyById[id];
    if (existing && existing.rowNo && !existing.deleted) {
      deleteRowNos.push(existing.rowNo);
      return;
    }
    if (!existing && t && String(t.sourceType || '') === 'CONTACT_NEXT' && !incomingIds[id]) {
      tombstoneRows.push(buildPortalTodayRowValuesP360_(t, {
        id: id,
        selectedDate: selectedDate,
        order: 9999,
        prev: {},
        now: now,
        user: user,
        deleted: 'Y'
      }));
    }
  });

  const baseTaskById = {};
  baseTasks.forEach(function(raw) {
    const t = normalizePortalTodayTaskP360_(raw || {}, selectedDate);
    if (t.id) baseTaskById[t.id] = t;
  });
  baseTaskIds.forEach(function(id) {
    if (incomingIds[id]) return;
    if (existingAnyById[id]) return;
    const t = baseTaskById[id];
    if (!t || String(t.sourceType || '') !== 'CONTACT_NEXT') return;
    tombstoneRows.push(buildPortalTodayRowValuesP360_(t, {
      id: id,
      selectedDate: selectedDate,
      order: 9999,
      prev: {},
      now: now,
      user: user,
      deleted: 'Y'
    }));
  });

  markPortalTodayRowsDeletedP360_(sheet, deleteRowNos);

  const rowsToAppend = appendRows.concat(tombstoneRows);
  if (rowsToAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, headerLen).setValues(rowsToAppend);
  }

  if (!payload.noFlush) SpreadsheetApp.flush();
  return {
    ok: true,
    selectedDate: selectedDate,
    authorFilter: access.authorFilter || 'ALL',
    baseTaskCount: baseTaskIds.length,
    upsertedCount: upsertedCount,
    appendedCount: appendRows.length,
    tombstoneCount: tombstoneRows.length,
    updatedCount: upsertedCount - appendRows.length,
    deletedCount: deleteRowNos.length + tombstoneRows.length
  };
}

function buildPortalTodayRowValuesP360_(item, context) {
  item = item || {};
  context = context || {};
  const now = context.now || new Date();
  const prev = context.prev || {};
  const sourceType = String(item.sourceType || prev.sourceType || '').trim() || 'MANUAL';
  const done = !!item.done;
  const createdAt = parsePortalTodayDateTime_(item.createdAt) || parsePortalTodayDateTime_(prev.createdAt) || now;
  const completedAt = done
    ? (parsePortalTodayDateTime_(item.completedAt) || parsePortalTodayDateTime_(prev.completedAt) || now)
    : null;
  const dueAt = done ? '' : String(item.dueAt || item.timeText || '').trim();
  return [
    String(context.id || item.id || '').trim() || Utilities.getUuid().slice(0, 12),
    context.selectedDate,
    Number(context.order) || 0,
    String(item.content || prev.content || '').trim(),
    done ? 'Y' : '',
    String(item.author || prev.author || '').trim() || context.user || getCurrentUserLabel_(),
    now,
    String(context.deleted || '').toUpperCase() === 'Y' ? 'Y' : '',
    String(item.category || prev.category || '').trim() || (sourceType === 'CONTACT_NEXT' ? '다음액션' : '할일'),
    sourceType,
    String(item.sourceId || prev.sourceId || '').trim(),
    String(item.customerNo || prev.customerNo || '').trim(),
    String(item.company || prev.company || '').trim(),
    String(item.rowNo || prev.rowNo || '').trim(),
    String(item.timeText || item.nextActionAt || prev.timeText || '').trim(),
    String(item.priority || prev.priority || '').trim(),
    createdAt,
    normalizePortalTodayTags_(item.tags || prev.tags || '').join(', '),
    String(item.detail || '').trim(),
    dueAt,
    completedAt || ''
  ];
}

function readPortalTodaySheetRowsForDateP360_(sheet, selectedDate, headerLen, access, options) {
  options = options || {};
  const matchingRows = readPortalTodayMatchingDateRowsFastP380_(sheet, selectedDate, headerLen);
  return matchingRows.map(function(rowInfo) {
    const item = portalTodayRowToItemP360_(rowInfo.values);
    const deleted = String(item.deleted || '').toUpperCase() === 'Y';
    const matchesDate = item.date === selectedDate;
    const matchesAccess = portalTodayAuthorMatchesAccessP360_(item.author, access);
    return {
      rowNo: rowInfo.rowNo,
      values: rowInfo.values,
      id: item.id,
      date: item.date,
      deleted: deleted,
      active: matchesDate && !deleted && !!item.content && matchesAccess,
      access: matchesAccess,
      item: item
    };
  }).filter(function(rowInfo) {
    if (rowInfo.date !== selectedDate) return false;
    if (!rowInfo.access) return false;
    if (options.includeDeleted) return true;
    return !rowInfo.deleted && !!rowInfo.item.content;
  });
}

function portalTodayRowToItemP360_(row) {
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

function portalTodayRowToItemP130_(row) { return portalTodayRowToItemP360_(row); }

function markPortalTodayRowsDeletedP360_(sheet, rowNos) {
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

function markPortalTodayRowsDeletedP130_(sheet, rowNos) { return markPortalTodayRowsDeletedP360_(sheet, rowNos); }

function getPortalTodayDateMetaP360_(selectedDate, tasks) {
  const active = (Array.isArray(tasks) ? tasks : [])
    .map(function(item) {
      item = item || {};
      return {
        id: String(item.id || '').trim(),
        updatedAt: String(item.updatedAt || '').trim(),
        done: item.done ? 'Y' : '',
        content: String(item.content || '').trim(),
        author: String(item.author || '').trim()
      };
    })
    .filter(function(item) { return item.id && item.content; })
    .sort(function(a, b) { return a.id.localeCompare(b.id); });
  const source = selectedDate + '|' + active.map(function(item) {
    return [item.id, item.updatedAt, item.done, item.content, item.author].join(':');
  }).join('|');
  return {
    version: digestPortalTodayVersionP360_(source),
    activeTaskIds: active.map(function(item) { return item.id; })
  };
}

function getPortalTodayDateMetaP130_(selectedDate, storedTodos) { return getPortalTodayDateMetaP360_(selectedDate, storedTodos); }
function getPortalTodayDateMetaFromRowsP130_(selectedDate, rows) {
  const tasks = (Array.isArray(rows) ? rows : []).filter(function(r) { return r && r.active; }).map(function(r) { return r.item || {}; });
  return getPortalTodayDateMetaP360_(selectedDate, tasks);
}

function digestPortalTodayVersionP360_(source) {
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
function digestPortalTodayVersionP130_(source) { return digestPortalTodayVersionP360_(source); }

function getPortalTodayDeletedSourceMapP360_(selectedDate, access) {
  const map = {};
  try {
    const sheet = ensurePortalTodaySheet_();
    const rows = readPortalTodaySheetRowsForDateP360_(sheet, selectedDate, PORTAL_CONFIG.TODAY_HEADERS.length, access, { includeDeleted: true });
    rows.forEach(function(rowInfo) {
      const item = rowInfo.item || {};
      if (String(item.deleted || '').toUpperCase() !== 'Y') return;
      if (item.id) map['ID|' + item.id] = true;
      if (item.sourceType && item.sourceId) map[String(item.sourceType) + '|' + String(item.sourceId)] = true;
    });
  } catch (err) {}
  return map;
}

function buildPortalTodayUnifiedTasksP360_(storedTodos, nextActions, selectedDate, deletedSourceMap) {
  storedTodos = Array.isArray(storedTodos) ? storedTodos : [];
  nextActions = Array.isArray(nextActions) ? nextActions : [];
  deletedSourceMap = deletedSourceMap || {};
  const result = [];
  const sourceTodoMap = {};

  storedTodos.forEach(function(todo) {
    const normalized = normalizePortalTodayTaskP360_(todo, selectedDate);
    const sourceType = String(normalized.sourceType || '').trim();
    const sourceId = String(normalized.sourceId || '').trim();
    if (sourceType && sourceType !== 'MANUAL' && sourceId) sourceTodoMap[sourceType + '|' + sourceId] = normalized;
    result.push(normalized);
  });

  nextActions.forEach(function(action) {
    const sourceId = String(action.sourceId || action.historyId || '').trim() || [action.customerNo, action.rowNo, action.nextActionAt, action.nextAction].join('|');
    const sourceKey = 'CONTACT_NEXT|' + sourceId;
    const taskId = 'contact_' + Utilities.base64EncodeWebSafe(sourceId).slice(0, 18);
    if (sourceTodoMap[sourceKey]) return;
    if (deletedSourceMap[sourceKey] || deletedSourceMap['ID|' + taskId]) return;
    result.push(normalizePortalTodayTaskP360_({
      id: taskId,
      date: selectedDate,
      done: false,
      category: '다음액션',
      sourceType: 'CONTACT_NEXT',
      sourceId: sourceId,
      customerNo: action.customerNo || '',
      company: action.company || '',
      rowNo: action.rowNo || '',
      timeText: action.nextActionAt || '',
      dueAt: action.nextActionAt || '',
      content: buildPortalTodayContactTaskTextP360_(action),
      author: action.author || '',
      tags: normalizePortalNextActionTagsP360_(action.nextAction || action.nextActionTags || '')
    }, selectedDate));
  });

  return result.sort(function(a, b) {
    if (!!a.done !== !!b.done) return a.done ? 1 : -1;
    const ad = String(a.dueAt || a.timeText || '');
    const bd = String(b.dueAt || b.timeText || '');
    if (ad || bd) return ad.localeCompare(bd);
    return (Number(a.order) || 0) - (Number(b.order) || 0);
  });
}
function buildPortalTodayUnifiedTasks_(storedTodos, nextActions, selectedDate) { return buildPortalTodayUnifiedTasksP360_(storedTodos, nextActions, selectedDate, {}); }

function normalizePortalTodayTaskP360_(task, selectedDate) {
  task = task || {};
  const sourceType = String(task.sourceType || '').trim() || 'MANUAL';
  return {
    id: String(task.id || '').trim() || Utilities.getUuid().slice(0, 12),
    date: normalizePortalTodoDate_(task.date || selectedDate || new Date()),
    order: Number(task.order) || 0,
    done: !!task.done,
    content: String(task.content || '').trim(),
    detail: String(task.detail || '').trim(),
    author: String(task.author || '').trim(),
    category: String(task.category || '').trim() || (sourceType === 'CONTACT_NEXT' ? '다음액션' : '할일'),
    sourceType: sourceType,
    sourceId: String(task.sourceId || '').trim(),
    customerNo: String(task.customerNo || '').trim(),
    company: String(task.company || '').trim(),
    rowNo: String(task.rowNo || '').trim(),
    timeText: String(task.timeText || task.nextActionAt || '').trim(),
    priority: String(task.priority || '').trim(),
    updatedAt: String(task.updatedAt || '').trim(),
    createdAt: formatPortalTodayInputDateTime_(task.createdAt || task.updatedAt || new Date()),
    dueAt: formatPortalTodayInputDateTime_(task.dueAt || task.timeText || ''),
    completedAt: formatPortalTodayInputDateTime_(task.completedAt || ''),
    tags: normalizePortalTodayTags_(task.tags),
    nextAction: String(task.nextAction || '').trim(),
    rawContent: String(task.rawContent || '').trim()
  };
}
function normalizePortalTodayTask_(task, selectedDate) { return normalizePortalTodayTaskP360_(task, selectedDate); }

function normalizePortalNextActionTagsP360_(value) {
  const tags = normalizePortalTodayTags_(value);
  function add(tag) { if (tag && tags.indexOf(tag) < 0) tags.push(tag); }
  const text = String(value || '');
  add('컨택');
  if (/전화|통화|콜|연락/.test(text)) add('전화');
  if (/메일|자료|발송|양식/.test(text)) add('메일');
  if (/견적/.test(text)) add('견적');
  if (/계약|서류/.test(text)) add('계약');
  if (/재확인|확인|보류/.test(text)) add('재확인');
  return tags;
}
function normalizePortalNextActionTags_(value) { return normalizePortalNextActionTagsP360_(value); }

function buildPortalTodayContactTaskTextP360_(action) {
  action = action || {};
  const parts = [];
  if (action.company) parts.push('[' + action.company + ']');
  if (action.nextAction) parts.push(action.nextAction);
  if (action.nextActionAt) parts.push(String(action.nextActionAt).replace(/^\d{4}-\d{1,2}-\d{1,2}\s*/, ''));
  if (action.content) parts.push('- ' + action.content);
  return parts.join(' ').trim() || '컨택 예정 확인';
}
function buildPortalTodayContactTaskText_(action) { return buildPortalTodayContactTaskTextP360_(action); }

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
  access = access || getPortalTodayAccessContextP360_({});
  const sheet = ensurePortalTodaySheet_();
  const rows = readPortalTodaySheetRowsForDateP360_(sheet, selectedDate, PORTAL_CONFIG.TODAY_HEADERS.length, access, { includeDeleted: false });
  return rows
    .map(function(rowInfo) { return rowInfo.item; })
    .filter(function(item) { return item && item.content; })
    .sort(function(a, b) { return (Number(a.order) || 0) - (Number(b.order) || 0); });
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
  const defaults = Array.isArray(PORTAL_CONFIG.TODAY_DEFAULT_TAGS) ? PORTAL_CONFIG.TODAY_DEFAULT_TAGS : ['전화','메일','견적','컨택','계약','서류','영업지원','재확인','긴급','내부처리'];
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
      const vals = sheet.getRange(2, 18, lastRow - 1, 1).getDisplayValues();
      vals.forEach(function(r) { normalizePortalTodayTags_(r[0]).forEach(add); });
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

function getContactNextActionsForDate_(selectedDate, access) {
  access = access || getPortalTodayAccessContextP360_({});
  return filterPortalTodayItemsByAccessP360_(getContactNextActionsRawForDateP370_(selectedDate), access)
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

// Backward-compatible aliases retained for older internal calls.
function buildPortalTodayRowValuesP130_(item, context) { return buildPortalTodayRowValuesP360_(item, context); }
function readPortalTodayRowsForDateP130_(sheet, selectedDate, headerLen, access) { return readPortalTodaySheetRowsForDateP360_(sheet, selectedDate, headerLen, access, { includeDeleted: false }); }
