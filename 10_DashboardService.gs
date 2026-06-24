/***************************************
 * S1 Sales Portal - 10_DashboardService.gs
 * v77: HOME 역할별 대시보드 집계
 * - ADMIN/서무: 전체 고객 / 처리할 요청 / 계약 완료 / 재확인 필요
 * - SALES: 나의 고객 / 오늘 할 일 / 나의 계약 완료 / 재확인 필요
 * - HOME 최초 렌더링은 캐시 우선, 실제 마스터 스캔은 refreshPortalDashboardCache()에서 수행
 ***************************************/

function getPortalDashboard(options) {
  options = options || {};
  const perm = getPortalCurrentPermission_();
  const keys = getPortalDashboardCacheKeysV77_(perm);
  const cache = CacheService.getScriptCache();
  const props = PropertiesService.getScriptProperties();

  let cached = '';
  try { cached = cache.get(keys.cacheKey); } catch (err) {}
  if (cached) {
    try {
      const data = JSON.parse(cached);
      data.fromCache = true;
      data.needsRefresh = isPortalDashboardStaleV77_(data.cachedAt);
      data.fastHome = true;
      return normalizePortalDashboardResultV77_(data, perm);
    } catch (err) {}
  }

  let propCached = '';
  try { propCached = props.getProperty(keys.propKey); } catch (err) {}
  if (propCached) {
    try {
      const data = JSON.parse(propCached);
      data.fromCache = true;
      data.fromScriptProperties = true;
      data.needsRefresh = isPortalDashboardStaleV77_(data.cachedAt);
      data.fastHome = true;
      try { cache.put(keys.cacheKey, JSON.stringify(data), PORTAL_CONFIG.DASHBOARD_CACHE_SECONDS || 300); } catch (err) {}
      return normalizePortalDashboardResultV77_(data, perm);
    } catch (err) {}
  }

  // 최초 접속 시에도 HOME을 막지 않기 위해 역할별 라벨만 즉시 내려줍니다.
  return normalizePortalDashboardResultV77_({
    total: '-',
    sent: '-',
    completed: '-',
    needCheck: '-',
    todayTodo: '-',
    pendingSupport: '-',
    contractCompleted: '-',
    cachedAt: '',
    fromCache: false,
    needsRefresh: true,
    initializing: true,
    fastHome: true
  }, perm);
}

function getPortalHomeData() {
  return {
    dashboard: getPortalDashboard({ fast: true }),
    notices: getPortalNotices(3),
    fastHome: true
  };
}

function refreshPortalDashboardCache() {
  const perm = getPortalCurrentPermission_();
  const keys = getPortalDashboardCacheKeysV77_(perm);
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(500)) {
    const current = getPortalDashboard({ fast: true });
    current.refreshSkipped = true;
    current.message = '대시보드 갱신이 이미 진행 중입니다.';
    return current;
  }

  try {
    const result = buildPortalDashboardStatsV77_(perm);
    savePortalDashboardCacheV77_(result, perm);
    return result;
  } catch (err) {
    const fallback = getPortalDashboard({ fast: true });
    fallback.refreshFailed = true;
    fallback.message = String(err && err.message || err);
    fallback.detail = String(err && err.stack || err);
    return fallback;
  } finally {
    try { lock.releaseLock(); } catch (err) {}
  }
}


// PATCH P1-14: HOME은 고객 전체열람권한과 별개로 권한_DB의 역할/기본범위를 기준으로 카드 구성을 결정합니다.
// SALES 계정은 전체고객열람=Y여도 HOME에서는 본인 기준(나의 고객/오늘 할 일/나의 계약 완료)을 봅니다.
function isPortalDashboardAllScopeV78_(perm) {
  // PATCH P1-15: HOME 관리자형/영업담당자형 판단은 전체고객열람이 아니라
  // 권한_DB에서 계산한 canUseAdminHome만 사용합니다.
  perm = perm || getPortalCurrentPermission_();
  if (typeof canPortalUseAdminHome_ === 'function') return canPortalUseAdminHome_(perm);
  return !!(perm && perm.active !== false && perm.canUseAdminHome);
}

function isPortalDashboardOwnCustomerRowAllowedV78_(row, perm) {
  perm = perm || getPortalCurrentPermission_();
  if (!perm || !perm.active) return false;
  if (isPortalDashboardAllScopeV78_(perm)) return true;
  const salesRep = normalizePortalNameForPermission_(getPortalCustomerSalesRepFromRow_(row));
  const aliases = [];
  if (Array.isArray(perm.salesRepAliases)) aliases.push.apply(aliases, perm.salesRepAliases);
  if (perm.salesRepName) aliases.push(perm.salesRepName);
  if (perm.name) aliases.push(perm.name);
  if (!salesRep || !aliases.length) return false;
  return aliases.some(function(alias) {
    const a = normalizePortalNameForPermission_(alias);
    return a && (salesRep === a || salesRep.indexOf(a) >= 0 || a.indexOf(salesRep) >= 0);
  });
}

function buildPortalDashboardStatsV77_(perm) {
  perm = perm || getPortalCurrentPermission_();
  const isAllScope = isPortalDashboardAllScopeV78_(perm);

  const allMasterRows = getMasterObjects_()
    .filter(function(row) { return String(getCompanyValue_(row) || '').trim() !== ''; });

  const scopedRows = isAllScope
    ? allMasterRows
    : allMasterRows.filter(function(row) { return isPortalDashboardOwnCustomerRowAllowedV78_(row, perm); });

  const scopedCustomerNoMap = buildPortalDashboardCustomerNoMapV77_(scopedRows);
  const needCheck = countPortalDashboardNeedCheckRowsV77_(scopedRows);
  const contractCompleted = countPortalDashboardContractCompletedV77_(perm, scopedCustomerNoMap);
  const todayTodo = countPortalDashboardTodayWorkV77_(perm, scopedCustomerNoMap);
  const pendingSupport = countPortalDashboardPendingSupportV77_(perm, scopedCustomerNoMap);

  const resultBase = {
    allScope: isAllScope,
    roleLevel: perm && perm.level || '',
    permissionName: perm && perm.name || '',
    salesRepName: perm && perm.salesRepName || '',
    totalCustomers: scopedRows.length,
    todayTodo: todayTodo,
    pendingSupport: pendingSupport,
    contractCompleted: contractCompleted,
    needCheck: needCheck,
    cachedAt: new Date().toISOString(),
    fromCache: false,
    refreshed: true,
    needsRefresh: false,
    fastHome: true
  };

  return normalizePortalDashboardResultV77_(resultBase, perm);
}

function normalizePortalDashboardResultV77_(data, perm) {
  data = data || {};
  perm = perm || getPortalCurrentPermission_();
  const isAllScope = isPortalDashboardAllScopeV78_(perm);
  const keys = getPortalDashboardCacheKeysV77_(perm);

  const totalCustomers = normalizePortalDashboardMetricValueV77_(data.totalCustomers != null ? data.totalCustomers : data.total);
  const todayTodo = normalizePortalDashboardMetricValueV77_(data.todayTodo);
  const pendingSupport = normalizePortalDashboardMetricValueV77_(data.pendingSupport != null ? data.pendingSupport : data.sent);
  const contractCompleted = normalizePortalDashboardMetricValueV77_(data.contractCompleted != null ? data.contractCompleted : data.completed);
  const needCheck = normalizePortalDashboardMetricValueV77_(data.needCheck);

  const secondLabel = isAllScope ? '처리할 요청' : '오늘 할 일';
  const secondValue = isAllScope ? pendingSupport : todayTodo;
  const totalLabel = isAllScope ? '전체 고객' : '나의 고객';
  const contractLabel = isAllScope ? '계약 완료' : '나의 계약 완료';

  data.total = totalCustomers;
  data.sent = secondValue;              // 기존 프론트 id(statSent) 호환용
  data.completed = contractCompleted;   // 기존 프론트 id(statCompleted) 호환용
  data.needCheck = needCheck;
  data.todayTodo = todayTodo;
  data.pendingSupport = pendingSupport;
  data.contractCompleted = contractCompleted;
  data.allScope = isAllScope;
  data.cacheScopeKey = keys.scopeKey;
  data.clientCacheKey = 'portalHomeCache:v77:' + keys.clientScopeKey;
  data.permission = sanitizePortalPermissionForClient_(perm);
  data.cards = [
    { key: 'totalCustomers', label: totalLabel, value: totalCustomers, scope: isAllScope ? 'ALL' : 'OWN' },
    { key: isAllScope ? 'pendingSupport' : 'todayTodo', label: secondLabel, value: secondValue, scope: isAllScope ? 'ALL' : 'OWN' },
    { key: 'contractCompleted', label: contractLabel, value: contractCompleted, scope: isAllScope ? 'ALL' : 'OWN' },
    { key: 'needCheck', label: '재확인 필요', value: needCheck, scope: isAllScope ? 'ALL' : 'OWN' }
  ];

  ['total', 'sent', 'completed', 'needCheck'].forEach(function(k) {
    if (data[k] === undefined || data[k] === null || data[k] === '') data[k] = '-';
  });
  return data;
}

function normalizePortalDashboardMetricValueV77_(value) {
  if (value === undefined || value === null || value === '') return '-';
  if (typeof value === 'number' && isNaN(value)) return '-';
  return value;
}

function getPortalDashboardCacheKeysV77_(perm) {
  perm = perm || getPortalCurrentPermission_();
  const isAllScope = isPortalDashboardAllScopeV78_(perm);
  const rawScope = isAllScope
    ? 'ALL_ADMIN'
    : ('OWN_' + (perm && (perm.salesRepName || perm.name || perm.email) || 'GUEST'));
  const clientScopeKey = normalizePortalDashboardClientScopeV77_(rawScope);
  const scopeKey = normalizePortalDashboardCacheScopeV77_(rawScope);
  return {
    scopeKey: scopeKey,
    clientScopeKey: clientScopeKey,
    cacheKey: 'PORTAL_DASHBOARD_V77_ROLE_' + scopeKey,
    propKey: 'PORTAL_DASHBOARD_V77_ROLE_' + scopeKey + '_JSON'
  };
}


function normalizePortalDashboardClientScopeV77_(value) {
  const text = String(value || '').trim() || 'UNKNOWN';
  return text.replace(/[^A-Za-z0-9가-힣_]+/g, '_').slice(0, 80);
}

function normalizePortalDashboardCacheScopeV77_(value) {
  const text = String(value || '').trim() || 'UNKNOWN';
  const digest = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text)
  ).replace(/=+$/g, '').slice(0, 14);
  return text.replace(/[^A-Za-z0-9가-힣_]+/g, '_').slice(0, 40) + '_' + digest;
}

function savePortalDashboardCacheV77_(result, perm) {
  const keys = getPortalDashboardCacheKeysV77_(perm);
  const json = JSON.stringify(result || {});
  try { CacheService.getScriptCache().put(keys.cacheKey, json, PORTAL_CONFIG.DASHBOARD_CACHE_SECONDS || 300); } catch (err) {}
  try { PropertiesService.getScriptProperties().setProperty(keys.propKey, json); } catch (err) {}
}

function isPortalDashboardStaleV77_(cachedAt) {
  if (!cachedAt) return true;
  const t = new Date(cachedAt).getTime();
  if (!t || isNaN(t)) return true;
  const maxAgeMs = Math.max(60, Number(PORTAL_CONFIG.DASHBOARD_CACHE_SECONDS || 300)) * 1000;
  return Date.now() - t > maxAgeMs;
}

function buildPortalDashboardCustomerNoMapV77_(rows) {
  const map = {};
  (rows || []).forEach(function(row) {
    const no = String(getMasterFieldValue_(row, 'customerNo') || row.customerNo || row['고객번호'] || '').trim();
    if (no) map[no] = true;
  });
  return map;
}

function countPortalDashboardNeedCheckRowsV77_(rows) {
  const needKeywords = ['재컨택', '재확인', '확인필요', '확인 필요', '검토', '상태지정필요', '상태 지정 필요', '부재', '재연락'];
  const excludeKeywords = ['계약완료', '계약 완료', '영업종료', '거절', '오류', '제외', '계약중도취소'];
  return (rows || []).filter(function(row) {
    const status = String(getStatusValueFromObj_(row) || '').trim();
    const memo = String(getMemoValueFromObj_(row) || '').trim();
    const target = status + ' ' + memo;
    if (containsAny_(status, excludeKeywords)) return false;
    return containsAny_(target, needKeywords);
  }).length;
}

function countPortalDashboardContractCompletedV77_(perm, scopedCustomerNoMap) {
  const rows = getPortalDashboardContractRowsV77_();
  if (!rows.length) return 0;

  const isAllScope = isPortalDashboardAllScopeV78_(perm);
  if (isAllScope) return rows.length;

  return rows.filter(function(item) {
    const customerNo = String(item.customerNo || '').trim();
    if (customerNo && scopedCustomerNoMap && scopedCustomerNoMap[customerNo]) return true;
    return isPortalDashboardNameMatchV77_(item.contractRep, perm);
  }).length;
}

function getPortalDashboardContractRowsV77_() {
  try {
    const ss = getMasterSpreadsheet_();
    const sheetName = (typeof PORTAL_CONTRACT_COMPLETE_SHEET_NAME_V69 !== 'undefined')
      ? PORTAL_CONTRACT_COMPLETE_SHEET_NAME_V69
      : '수주확정/계약완료';
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return [];

    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
    const map = buildPortalDashboardSimpleHeaderMapV77_(headers);
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();

    return values.map(function(row) {
      return {
        contractNo: getPortalDashboardCellV77_(row, map, ['계약번호']),
        customerNo: getPortalDashboardCellV77_(row, map, ['고객번호']),
        company: getPortalDashboardCellV77_(row, map, ['고객사명', '회사명', '고객명']),
        contractRep: getPortalDashboardCellV77_(row, map, ['계약담당자', '계약 담당자', '영업담당자', '견적담당'])
      };
    }).filter(function(item) {
      // 계약번호/고객번호만 미리 적어둔 준비행은 계약 완료 건으로 보지 않습니다.
      return String(item.company || '').trim() !== '';
    });
  } catch (err) {
    Logger.log('계약완료 대시보드 집계 실패: ' + (err && err.stack || err));
    return [];
  }
}

function countPortalDashboardPendingSupportV77_(perm, scopedCustomerNoMap) {
  try {
    if (typeof getPortalSupportRows_ !== 'function') return 0;
    const rows = getPortalSupportRows_();
    const isAllScope = isPortalDashboardAllScopeV78_(perm);
    return rows.filter(function(item) {
      if (!item || !(item.requestText || item.customerNo || item.customerName)) return false;
      const status = String(item.status || '접수').replace(/\s+/g, '').trim();
      if (status === '완료' || status === '반려') return false;
      if (isAllScope) return true;
      const no = String(item.customerNo || '').trim();
      if (no && scopedCustomerNoMap && scopedCustomerNoMap[no]) return true;
      return isPortalDashboardNameMatchV77_(item.requester, perm);
    }).length;
  } catch (err) {
    Logger.log('영업지원 pending 대시보드 집계 실패: ' + (err && err.stack || err));
    return 0;
  }
}

function countPortalDashboardTodayWorkV77_(perm, scopedCustomerNoMap) {
  try {
    const today = (typeof normalizePortalTodoDate_ === 'function')
      ? normalizePortalTodoDate_(new Date())
      : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const isAllScope = isPortalDashboardAllScopeV78_(perm);

    let todoCount = 0;
    if (typeof getPortalTodosForDate_ === 'function') {
      const todos = getPortalTodosForDate_(today) || [];
      todoCount = todos.filter(function(item) {
        if (!item || item.done) return false;
        if (isAllScope) return true;
        return isPortalDashboardNameMatchV77_(item.author, perm);
      }).length;
    }

    let nextActionCount = 0;
    if (typeof getContactNextActionsForDate_ === 'function') {
      const actions = getContactNextActionsForDate_(today) || [];
      nextActionCount = actions.filter(function(item) {
        if (!item) return false;
        if (isAllScope) return true;
        const no = String(item.customerNo || '').trim();
        if (no && scopedCustomerNoMap && scopedCustomerNoMap[no]) return true;
        return isPortalDashboardNameMatchV77_(item.author, perm);
      }).length;
    }

    return todoCount + nextActionCount;
  } catch (err) {
    Logger.log('오늘 할 일 대시보드 집계 실패: ' + (err && err.stack || err));
    return 0;
  }
}

function buildPortalDashboardSimpleHeaderMapV77_(headers) {
  const map = { exact: {}, normalized: {} };
  (headers || []).forEach(function(header, idx) {
    const h = String(header || '').trim();
    if (!h) return;
    map.exact[h] = idx;
    map.normalized[normalizePortalDashboardHeaderV77_(h)] = idx;
  });
  return map;
}

function getPortalDashboardCellV77_(row, map, candidates) {
  candidates = Array.isArray(candidates) ? candidates : [candidates];
  for (let i = 0; i < candidates.length; i++) {
    const h = String(candidates[i] || '').trim();
    if (!h) continue;
    if (Object.prototype.hasOwnProperty.call(map.exact, h)) {
      const v = String(row[map.exact[h]] || '').trim();
      if (v) return v;
    }
    const key = normalizePortalDashboardHeaderV77_(h);
    if (Object.prototype.hasOwnProperty.call(map.normalized, key)) {
      const v2 = String(row[map.normalized[key]] || '').trim();
      if (v2) return v2;
    }
  }
  return '';
}

function normalizePortalDashboardHeaderV77_(value) {
  return String(value || '').replace(/[\s\n\r\t\/·_()（）\-]+/g, '').trim().toLowerCase();
}

function isPortalDashboardNameMatchV77_(value, perm) {
  const text = normalizePortalNameForPermission_(value || '');
  if (!text || !perm) return false;
  const aliases = [];
  if (Array.isArray(perm.salesRepAliases)) aliases.push.apply(aliases, perm.salesRepAliases);
  if (perm.salesRepName) aliases.push(perm.salesRepName);
  if (perm.name) aliases.push(perm.name);
  return aliases.some(function(alias) {
    const a = normalizePortalNameForPermission_(alias);
    return a && (text === a || text.indexOf(a) >= 0 || a.indexOf(text) >= 0);
  });
}
