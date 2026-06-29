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


/***************************************
 * P443: HOME 수주현황 위젯
 * - 원천: 수주확정/계약완료 시트(listContractCompleteRowsV69)
 * - ADMIN/서무: 전체 수주 현황
 * - SALES: 본인 계약담당자 수주 현황
 ***************************************/
const PORTAL_HOME_SALES_STATUS_CACHE_PREFIX_P443 = 'PORTAL_HOME_SALES_STATUS_P443_';
const PORTAL_HOME_SALES_STATUS_CACHE_SECONDS_P443 = 180;

function getPortalHomeSalesStatus(options) {
  options = options || {};
  const perm = getPortalCurrentPermission_();
  const isAllScope = isPortalDashboardAllScopeV78_(perm);
  const cacheKey = makePortalHomeSalesStatusCacheKeyP443_(perm);
  const cache = CacheService.getScriptCache();

  if (!options.force) {
    try {
      const cached = cache.get(cacheKey);
      if (cached) {
        const parsed = JSON.parse(cached);
        parsed.fromCache = true;
        return parsed;
      }
    } catch (err) {}
  }

  let list;
  try {
    list = (typeof listContractCompleteRowsV69 === 'function')
      ? listContractCompleteRowsV69({ force: true })
      : { rows: getPortalDashboardContractRowsV77_() };
  } catch (err) {
    return {
      ok: false,
      message: '수주현황 데이터를 불러오지 못했습니다: ' + String(err && err.message || err),
      scope: isAllScope ? 'ALL' : 'OWN',
      allScope: isAllScope,
      permissionName: perm && (perm.salesRepName || perm.name || '') || '',
      generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    };
  }

  const rawRows = Array.isArray(list && list.rows) ? list.rows : [];
  const scopedRows = rawRows.filter(function(item) {
    if (!item || !String(item.company || item.customerNo || item.contractNo || '').trim()) return false;
    if (isAllScope) return true;
    return isPortalDashboardNameMatchV77_(item.contractRep, perm);
  }).map(function(item) {
    return normalizePortalHomeSalesRowP443_(item);
  }).filter(function(item) { return !!item.company; });

  scopedRows.sort(comparePortalHomeSalesRowsP443_);

  const today = stripPortalHomeSalesTimeP443_(new Date());
  const weekStart = getPortalHomeSalesWeekStartP443_(today);
  const weekEnd = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate() + 7);
  const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);
  const last30Start = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 29);

  const thisWeekRows = filterPortalHomeSalesByDateRangeP443_(scopedRows, weekStart, weekEnd);
  const thisMonthRows = filterPortalHomeSalesByDateRangeP443_(scopedRows, monthStart, nextMonth);
  const last30Rows = filterPortalHomeSalesByDateRangeP443_(scopedRows, last30Start, new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1));

  const repSummary = buildPortalHomeSalesRepSummaryP443_(scopedRows, monthStart, nextMonth);
  const periodBuckets = buildPortalHomeSalesPeriodBucketsP443_(scopedRows, today);

  const result = {
    ok: true,
    scope: isAllScope ? 'ALL' : 'OWN',
    allScope: isAllScope,
    permissionName: perm && (perm.salesRepName || perm.name || '') || '',
    title: isAllScope ? '전체 수주현황' : '나의 수주현황',
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    sourceTotal: rawRows.length,
    total: summarizePortalHomeSalesRowsP443_(scopedRows),
    thisWeek: summarizePortalHomeSalesRowsP443_(thisWeekRows),
    thisMonth: summarizePortalHomeSalesRowsP443_(thisMonthRows),
    last30: summarizePortalHomeSalesRowsP443_(last30Rows),
    repSummary: repSummary.slice(0, 12),
    periodBuckets: periodBuckets,
    thisWeekList: thisWeekRows.slice(0, 8),
    recentList: scopedRows.slice(0, 10),
    fromCache: false
  };

  try {
    const json = JSON.stringify(result);
    if (json.length < 90000) cache.put(cacheKey, json, PORTAL_HOME_SALES_STATUS_CACHE_SECONDS_P443);
  } catch (err) {}

  return result;
}

function makePortalHomeSalesStatusCacheKeyP443_(perm) {
  const dashKeys = getPortalDashboardCacheKeysV77_(perm || getPortalCurrentPermission_());
  return (PORTAL_HOME_SALES_STATUS_CACHE_PREFIX_P443 + dashKeys.scopeKey).slice(0, 240);
}

function normalizePortalHomeSalesRowP443_(item) {
  item = item || {};
  const contractDateRaw = item.contractDate || item.contractDateText || item['계약일자(발주번호 부여일)'] || '';
  const dateObj = parsePortalHomeSalesDateP443_(contractDateRaw);
  const amount = parsePortalHomeSalesAmountP443_(item.contractPrice || item.contractAmount || item.price || item['계약가'] || '');
  const contractNo = cleanPortalHomeNumberTextP443_(item.contractNo);
  const customerNo = cleanPortalHomeNumberTextP443_(item.customerNo);
  return {
    contractNo: contractNo,
    customerNo: customerNo,
    company: String(item.company || '').trim(),
    contractRep: normalizePortalHomeSalesPersonNameP443_(item.contractRep || ''),
    contractDate: dateObj ? Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy.MM.dd.') : String(contractDateRaw || '').trim(),
    contractDateIso: dateObj ? Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
    contractDateTime: dateObj ? dateObj.getTime() : 0,
    amount: amount,
    amountText: formatPortalHomeSalesCurrencyP443_(amount),
    vendor: String(item.vendor || '').trim(),
    region: String(item.region || item.regionCity || '').trim(),
    grade: String(item.grade || '').trim(),
    contractPeriod: String(item.contractPeriod || '').trim()
  };
}

function comparePortalHomeSalesRowsP443_(a, b) {
  const ad = Number(a && a.contractDateTime || 0);
  const bd = Number(b && b.contractDateTime || 0);
  if (ad !== bd) return bd - ad;
  const an = Number(String(a && a.contractNo || '').replace(/[^0-9.\-]/g, ''));
  const bn = Number(String(b && b.contractNo || '').replace(/[^0-9.\-]/g, ''));
  if (!isNaN(an) && !isNaN(bn) && an !== bn) return bn - an;
  return String(b && b.company || '').localeCompare(String(a && a.company || ''), 'ko', { numeric: true, sensitivity: 'base' });
}

function parsePortalHomeSalesDateP443_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return stripPortalHomeSalesTimeP443_(value);
  let text = String(value == null ? '' : value).trim();
  if (!text) return null;

  const serialText = text.replace(/,/g, '');
  if (/^\d{5}(?:\.0+)?$/.test(serialText)) {
    const serial = Number(serialText);
    if (!isNaN(serial) && serial > 30000) {
      return stripPortalHomeSalesTimeP443_(new Date(Math.round((serial - 25569) * 86400 * 1000)));
    }
  }

  text = text.replace(/년|월/g, '.').replace(/일/g, '.').replace(/\s+/g, ' ');
  let m = text.match(/(20\d{2}|\d{2})\s*[.\-/]\s*(\d{1,2})\s*[.\-/]\s*(\d{1,2})/);
  if (m) {
    let y = Number(m[1]);
    if (y < 100) y += 2000;
    const d = new Date(y, Number(m[2]) - 1, Number(m[3]));
    return isNaN(d.getTime()) ? null : stripPortalHomeSalesTimeP443_(d);
  }
  const fallback = new Date(value);
  return isNaN(fallback.getTime()) ? null : stripPortalHomeSalesTimeP443_(fallback);
}

function stripPortalHomeSalesTimeP443_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function getPortalHomeSalesWeekStartP443_(date) {
  const d = stripPortalHomeSalesTimeP443_(date || new Date());
  const day = d.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + diff);
}

function filterPortalHomeSalesByDateRangeP443_(rows, start, endExclusive) {
  const startMs = start.getTime();
  const endMs = endExclusive.getTime();
  return (rows || []).filter(function(row) {
    const t = Number(row && row.contractDateTime || 0);
    return t && t >= startMs && t < endMs;
  });
}

function summarizePortalHomeSalesRowsP443_(rows) {
  rows = Array.isArray(rows) ? rows : [];
  const amount = rows.reduce(function(sum, row) { return sum + Number(row.amount || 0); }, 0);
  return {
    count: rows.length,
    amount: amount,
    amountText: formatPortalHomeSalesCurrencyP443_(amount)
  };
}

function buildPortalHomeSalesRepSummaryP443_(rows, monthStart, nextMonth) {
  const map = {};
  (rows || []).forEach(function(row) {
    const rep = normalizePortalHomeSalesPersonNameP443_(row.contractRep || '') || '미지정';
    if (!map[rep]) map[rep] = { rep: rep, count: 0, amount: 0, monthCount: 0, monthAmount: 0, recentDateTime: 0 };
    map[rep].count += 1;
    map[rep].amount += Number(row.amount || 0);
    map[rep].recentDateTime = Math.max(map[rep].recentDateTime, Number(row.contractDateTime || 0));
    if (row.contractDateTime >= monthStart.getTime() && row.contractDateTime < nextMonth.getTime()) {
      map[rep].monthCount += 1;
      map[rep].monthAmount += Number(row.amount || 0);
    }
  });
  return Object.keys(map).map(function(k) {
    const item = map[k];
    item.amountText = formatPortalHomeSalesCurrencyP443_(item.amount);
    item.monthAmountText = formatPortalHomeSalesCurrencyP443_(item.monthAmount);
    return item;
  }).sort(function(a, b) {
    if (b.count !== a.count) return b.count - a.count;
    if (b.amount !== a.amount) return b.amount - a.amount;
    return b.recentDateTime - a.recentDateTime;
  });
}

function buildPortalHomeSalesPeriodBucketsP443_(rows, today) {
  const thisMonthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  const buckets = [];
  for (let i = 0; i < 4; i++) {
    const start = new Date(thisMonthStart.getFullYear(), thisMonthStart.getMonth() - i, 1);
    const end = new Date(start.getFullYear(), start.getMonth() + 1, 1);
    const periodRows = filterPortalHomeSalesByDateRangeP443_(rows, start, end);
    const summary = summarizePortalHomeSalesRowsP443_(periodRows);
    buckets.push({
      label: Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy.MM'),
      count: summary.count,
      amount: summary.amount,
      amountText: summary.amountText
    });
  }
  return buckets;
}

function parsePortalHomeSalesAmountP443_(value) {
  if (typeof value === 'number') return isNaN(value) ? 0 : value;
  const text = String(value == null ? '' : value).replace(/[^0-9.\-]/g, '');
  const n = Number(text);
  return isNaN(n) ? 0 : n;
}

function formatPortalHomeSalesCurrencyP443_(amount) {
  const n = Number(amount || 0);
  if (!n) return '₩0';
  return '₩' + Math.round(n).toLocaleString('ko-KR');
}

function normalizePortalHomeSalesPersonNameP443_(value) {
  let text = String(value || '').trim();
  if (!text) return '';
  text = text.replace(/\([^)]*\)/g, '').replace(/\[[^\]]*\]/g, '').trim();
  text = text.replace(/\s+/g, ' ');
  const tokens = text.split(' ').filter(Boolean);
  // 팀/직급이 섞여 들어와도 실제 이름만 남기는 약식 보정입니다.
  const ranks = ['대표', '책임', '차장', '과장', '대리', '주임', '팀장', '실장', '부장', '사원'];
  const cleaned = tokens.filter(function(t) { return ranks.indexOf(t) < 0 && !/팀$|부$|실$/.test(t); });
  return cleaned.length ? cleaned[cleaned.length - 1] : tokens[0];
}

function cleanPortalHomeNumberTextP443_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!text) return '';
  if (/^\d+(?:\.0+)?$/.test(text)) return String(Math.round(Number(text)));
  return text;
}
