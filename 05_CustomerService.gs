/***************************************
 * S1 Sales Portal - 05_CustomerService.gs
 * 분리일: 2026-06-19
 * PATCH Q: 마스터시트 헤더 매핑 중앙화 적용
 * v70: 회사명 공란 행은 고객 리스트/검색인덱스 표시 제외
 ***************************************/

// PATCH K-2: 검색인덱스_DB를 상세 Lite DB로 확장합니다.
// 기존 PORTAL_CONFIG.CUSTOMER_INDEX_HEADERS는 유지하되, 서버 내부에서는 아래 확장 헤더를 사용합니다.
// 주의: 이 DB는 source of truth가 아니라 검색/빠른 표시용 snapshot입니다.
function getCustomerSearchIndexHeadersK2_() {
  const base = Array.isArray(PORTAL_CONFIG.CUSTOMER_INDEX_HEADERS) ? PORTAL_CONFIG.CUSTOMER_INDEX_HEADERS.slice() : [];
  const extra = [
    '발주번호',
    '마스터시트 최초등록일',
    '지역구분',
    '고객등급',
    '연면적',
    '관리등급',
    '건물 유형',
    '계약단위',
    '계약시작일',
    '계약종료일',
    '제보자',
    '관리자 선임 여부',
    '유지점검',
    '성능점검',
    '부가세',
    '할인율',
    '용역신청서특약사항',
    '마지막발송',
    '발송일시',
    '상세Lite여부',
    '마스터원본버전',
    '최종수정자'
  ];
  const seen = {};
  const result = [];
  base.concat(extra).forEach(function(h) {
    h = String(h || '').trim();
    if (!h || seen[h]) return;
    seen[h] = true;
    result.push(h);
  });
  return result;
}

function getIndexCellByAnyHeaderK2_(row, map, headers) {
  headers = Array.isArray(headers) ? headers : [headers];
  for (let i = 0; i < headers.length; i++) {
    const v = cellByIndexHeader_(row, map, headers[i]);
    if (String(v || '').trim() !== '') return v;
  }
  return '';
}

function normalizeMasterHeaderKeyK2_(header) {
  // PATCH Q: 04_MasterRepository.gs의 공통 normalizer를 우선 사용합니다.
  if (typeof normalizeMasterHeaderKey_ === 'function') return normalizeMasterHeaderKey_(header);
  return String(header || '').replace(/[\s\n\r\t]+/g, '').trim().toLowerCase();
}

function getCustomerMasterHeaderValueK2_(obj, fieldKeyOrCandidates) {
  obj = obj || {};

  // PATCH Q: field key(customerNo, grade, businessNo 등)가 들어오면 01_Schema.gs의 중앙 schema를 사용합니다.
  if (typeof fieldKeyOrCandidates === 'string' && typeof getMasterFieldDef_ === 'function' && getMasterFieldDef_(fieldKeyOrCandidates)) {
    return getMasterFieldValue_(obj, fieldKeyOrCandidates);
  }

  const candidates = Array.isArray(fieldKeyOrCandidates) ? fieldKeyOrCandidates : [fieldKeyOrCandidates];
  return getValueByHeaderCandidates_(obj, candidates || []);
}

function getCustomerIndexObjectValueK2_(obj, fieldKeyOrCandidates) {
  return getCustomerMasterHeaderValueK2_(obj || {}, fieldKeyOrCandidates || []);
}

function isCustomerSearchIndexK2Ready_() {
  const sheet = ensureCustomerSearchIndexSheet_();
  const lastRow = sheet.getLastRow();
  const width = Math.max(sheet.getLastColumn(), getCustomerSearchIndexHeadersK2_().length);
  const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const required = ['연면적', '계약단위', '관리자 선임 여부', '유지점검', '성능점검', '부가세', '상세Lite여부'];
  const missing = required.some(function(h) { return headers.indexOf(h) < 0; });
  if (missing) return false;
  if (lastRow < 2) return false;
  const liteCol = headers.indexOf('상세Lite여부') + 1;
  if (!liteCol) return false;
  const sampleCount = Math.min(20, lastRow - 1);
  const vals = sheet.getRange(2, liteCol, sampleCount, 1).getDisplayValues();
  return vals.some(function(r) { return String(r[0] || '').trim().toUpperCase() === 'Y'; });
}


function searchCustomers(keyword) {
  return searchCustomersPaged(keyword, 0, 100).rows;
}

function searchCustomersPaged(keyword, page, pageSize, sortMode, scopeMode) {
  keyword = normalizeSearchKeyword_(keyword);
  page = Math.max(0, Number(page) || 0);
  pageSize = Math.max(1, Math.min(100, Number(pageSize) || PORTAL_CONFIG.SEARCH_PAGE_SIZE || 20));

  const permForScope = getPortalCurrentPermission_();
  const effectiveScopeMode = normalizeCustomerSearchScopeModeP04_(scopeMode, permForScope);
  const indexData = getCustomerSearchIndexData(permForScope);
  const rows = filterCustomerIndexRowsByPermissionAndScopeP04_(indexData.rows || [], effectiveScopeMode, permForScope, true);
  const filtered = sortCustomerIndexRowsV67_(filterCustomerIndexRows_(rows, keyword), sortMode || 'customerNoDigitsDesc');
  const total = filtered.length;
  const start = page * pageSize;
  const pageRows = filtered.slice(start, start + pageSize);

  return {
    rows: pageRows,
    total: total,
    page: page,
    pageSize: pageSize,
    start: start,
    end: Math.min(start + pageRows.length, total),
    hasPrev: page > 0,
    hasNext: start + pageSize < total,
    keyword: keyword,
    sortMode: sortMode || 'customerNoDigitsDesc',
    scopeMode: effectiveScopeMode,
    source: '검색인덱스_DB',
    version: indexData.version || '',
    builtAt: indexData.builtAt || ''
  };
}


function normalizeCustomerSearchScopeModeP04_(scopeMode, perm) {
  const raw = String(scopeMode || '').trim().toUpperCase();
  if (raw === 'OWN' || raw === 'ALL') return raw;

  // P2-5: 클라이언트가 scope를 누락하거나 오래된 코드에서 호출해도
  // SALES 계정의 기본 조회 범위는 전체가 아니라 반드시 OWN입니다.
  perm = perm || getPortalCurrentPermission_();
  const defaultScope = String(perm && perm.defaultScope || '').trim().toUpperCase();
  if (defaultScope === 'OWN' || defaultScope === 'ALL') return defaultScope;
  const level = String(perm && perm.level || '').trim().toUpperCase();
  return level === 'SALES' ? 'OWN' : 'ALL';
}

function isCustomerIndexRowOwnedByPermissionP04_(row, perm) {
  perm = perm || getPortalCurrentPermission_();
  const salesRep = normalizePortalNameForPermission_(getPortalCustomerSalesRepFromRow_(row));
  const aliases = perm.salesRepAliases || splitPortalPermissionAliases_(perm.salesRepName || perm.name || '');
  if (!salesRep || !aliases.length) return false;
  return aliases.some(function(alias) {
    const a = normalizePortalNameForPermission_(alias);
    return a && (salesRep === a || salesRep.indexOf(a) >= 0 || a.indexOf(salesRep) >= 0);
  });
}

function filterCustomerIndexRowsByPermissionAndScopeP04_(rows, scopeMode, perm, alreadyPermissionFiltered) {
  rows = Array.isArray(rows) ? rows : [];
  perm = perm || getPortalCurrentPermission_();
  if (!alreadyPermissionFiltered) rows = filterPortalCustomerRowsByPermission_(rows, perm);
  const canAll = !!(perm && (perm.canViewAllCustomers || perm.isAdmin));
  const scope = canAll ? normalizeCustomerSearchScopeModeP04_(scopeMode, perm) : 'OWN';
  if (scope !== 'OWN') return rows;
  return rows.filter(function(row) { return isCustomerIndexRowOwnedByPermissionP04_(row, perm); });
}

const CUSTOMER_SEARCH_INDEX_SCHEMA_VERSION_P360 = 'P360_MEMO5000_PREWARM';

function getCustomerSearchIndexData(permForFilter) {
  const sheet = ensureCustomerSearchIndexSheet_();
  const props = PropertiesService.getScriptProperties();
  let rows = getCustomerSearchIndexRows_();

  // STEP22/P310: 검색인덱스는 빠른 조회용 보조 캐시일 뿐이고, 원천은 항상 영업관리대장 마스터시트입니다.
  // 마스터 직접수정/onEdit/변경큐 실패로 dirty 표시가 있으면, 다음 조회에서 먼저 마스터시트 기준 재생성을 시도합니다.
  // 즉, dirty 상태의 오래된 검색인덱스를 웹앱 목록의 원천처럼 계속 노출하지 않습니다.
  let rebuildInfo = null;
  const needsK2Rebuild = !isCustomerSearchIndexK2Ready_();
  const indexDirtyBefore = props.getProperty('CUSTOMER_SEARCH_INDEX_DIRTY') === 'Y';
  const schemaChangedP360 = props.getProperty('CUSTOMER_SEARCH_INDEX_SCHEMA_VERSION') !== CUSTOMER_SEARCH_INDEX_SCHEMA_VERSION_P360;

  // v31 FIX 유지: 인덱스가 비어 있거나 구조가 맞지 않거나 dirty이면 첫 요청자가 짧게 rebuild합니다.
  // STEP36: 메모 보관 길이/프리워밍 정책 변경 시 기존 350자 인덱스가 계속 쓰이지 않도록 스키마 버전도 봅니다.
  if (needsK2Rebuild || !rows.length || indexDirtyBefore || schemaChangedP360) {
    const reason = schemaChangedP360 ? 'P360_SCHEMA_MEMO_PREWARM' : (indexDirtyBefore ? 'DIRTY_MASTER_PRIORITY' : (needsK2Rebuild ? 'K2_DETAIL_LITE' : 'EMPTY_INDEX'));
    rebuildInfo = rebuildCustomerSearchIndex({ auto: true, maxWaitMs: 700, reason: reason, skipFormat: true });
    if (rebuildInfo && rebuildInfo.ok) {
      rows = getCustomerSearchIndexRows_();
    }
  }

  // PATCH S: 권한_DB 기준 고객 목록 제한.
  // 영업팀은 기본적으로 본인 영업담당자명과 매칭되는 고객만 내려줍니다.
  rows = filterPortalCustomerRowsByPermission_(rows, permForFilter);

  const dirtyAfter = props.getProperty('CUSTOMER_SEARCH_INDEX_DIRTY') === 'Y';
  return {
    ok: true,
    rebuilding: !!(rebuildInfo && rebuildInfo.locked),
    needsK2Rebuild: !!(needsK2Rebuild && !(rebuildInfo && rebuildInfo.ok)),
    dirty: dirtyAfter,
    dirtyReason: props.getProperty('CUSTOMER_SEARCH_INDEX_DIRTY_REASON') || '',
    dirtyAt: props.getProperty('CUSTOMER_SEARCH_INDEX_DIRTY_AT') || '',
    masterVersion: props.getProperty('PORTAL_MASTER_DATA_VERSION') || '',
    masterChangedAt: props.getProperty('PORTAL_MASTER_DATA_CHANGED_AT') || '',
    message: rebuildInfo && rebuildInfo.locked ? rebuildInfo.message : '',
    version: props.getProperty('CUSTOMER_SEARCH_INDEX_VERSION') || getCustomerSearchIndexVersion_(),
    builtAt: props.getProperty('CUSTOMER_SEARCH_INDEX_BUILT_AT') || '',
    total: rows.length,
    rows: rows
  };
}


/**
 * STEP22/P310: 서무/admin 전용 마스터시트 기준 검색인덱스 강제갱신.
 * - 웹앱의 빠른 검색인덱스/브라우저 캐시가 영업관리대장 최신값보다 우선 보이는 상황을 복구합니다.
 * - 권한은 서버에서 다시 검증합니다.
 */
function forceCustomerSearchIndexFromMasterP310(options) {
  options = options || {};
  const perm = getPortalCurrentPermission_();
  if (!perm || perm.active === false || !perm.canUseAdminHome) {
    throw new Error('마스터시트 기준 강제갱신은 서무/admin 계정에서만 사용할 수 있습니다.');
  }
  const startedAt = new Date();
  const result = rebuildCustomerSearchIndex({
    auto: false,
    maxWaitMs: Math.max(3000, Math.min(15000, Number(options.maxWaitMs) || 10000)),
    reason: 'ADMIN_FORCE_MASTER_PRIORITY'
  });
  if (!result || !result.ok) {
    return Object.assign({ ok: false }, result || {}, {
      message: result && result.message ? result.message : '검색인덱스 강제갱신에 실패했습니다.'
    });
  }

  try { markPortalMasterDataChangedP201_('admin-force-search-index-from-master'); } catch (err) {}
  try { clearCustomerSearchIndexDirty_(); } catch (err) {}
  try { if (typeof clearContractCompleteCacheV69_ === 'function') clearContractCompleteCacheV69_(); } catch (err) {}
  try { if (typeof clearFocusCustomerCacheP290 === 'function') clearFocusCustomerCacheP290(); } catch (err) {}

  const fresh = getCustomerSearchIndexData(perm);
  const elapsedMs = new Date().getTime() - startedAt.getTime();
  return {
    ok: true,
    message: '영업관리대장 마스터시트 기준으로 검색인덱스_DB를 강제갱신했습니다.',
    rebuilt: result.rebuilt || 0,
    version: result.version || (fresh && fresh.version) || '',
    builtAt: result.builtAt || (fresh && fresh.builtAt) || '',
    elapsedMs: elapsedMs,
    indexData: fresh
  };
}

function getCustomerSearchIndexMeta() {
  const sheet = ensureCustomerSearchIndexSheet_();
  const props = PropertiesService.getScriptProperties();
  return {
    ok: true,
    sheetName: PORTAL_CONFIG.CUSTOMER_INDEX_SHEET_NAME,
    version: props.getProperty('CUSTOMER_SEARCH_INDEX_VERSION') || getCustomerSearchIndexVersion_(),
    builtAt: props.getProperty('CUSTOMER_SEARCH_INDEX_BUILT_AT') || '',
    dirty: props.getProperty('CUSTOMER_SEARCH_INDEX_DIRTY') === 'Y',
    dirtyReason: props.getProperty('CUSTOMER_SEARCH_INDEX_DIRTY_REASON') || '',
    dirtyAt: props.getProperty('CUSTOMER_SEARCH_INDEX_DIRTY_AT') || '',
    total: Math.max(0, sheet.getLastRow() - 1),
    headers: getCustomerSearchIndexHeadersK2_()
  };
}


// =========================
// PATCH R: 검색인덱스_DB 정합성/dirty 관리 공통 함수
// =========================
function touchCustomerSearchIndexVersion_(now) {
  now = now || new Date();
  const ts = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const version = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  PropertiesService.getScriptProperties().setProperties({
    CUSTOMER_SEARCH_INDEX_VERSION: version,
    CUSTOMER_SEARCH_INDEX_BUILT_AT: ts,
    CUSTOMER_SEARCH_INDEX_DIRTY: 'N',
    CUSTOMER_SEARCH_INDEX_DIRTY_REASON: '',
    CUSTOMER_SEARCH_INDEX_DIRTY_AT: ''
  }, true);
  return { version: version, builtAt: ts };
}

function markCustomerSearchIndexDirty_(reason, detail) {
  const now = new Date();
  const ts = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const msg = [reason || 'UNKNOWN', detail || ''].filter(Boolean).join(' | ');
  PropertiesService.getScriptProperties().setProperties({
    CUSTOMER_SEARCH_INDEX_DIRTY: 'Y',
    CUSTOMER_SEARCH_INDEX_DIRTY_REASON: msg.slice(0, 500),
    CUSTOMER_SEARCH_INDEX_DIRTY_AT: ts
  }, true);
  return { ok: false, dirty: true, reason: reason || 'UNKNOWN', dirtyAt: ts, detail: String(detail || '') };
}

function clearCustomerSearchIndexDirty_() {
  PropertiesService.getScriptProperties().setProperties({
    CUSTOMER_SEARCH_INDEX_DIRTY: 'N',
    CUSTOMER_SEARCH_INDEX_DIRTY_REASON: '',
    CUSTOMER_SEARCH_INDEX_DIRTY_AT: ''
  }, true);
}

function updateCustomerSearchIndexAfterMutation_(targetOrPayload, values, reason) {
  targetOrPayload = targetOrPayload || {};
  values = values || {};
  const rowNo = Number(targetOrPayload.rowNo || targetOrPayload.masterRow || 0) || 0;
  const customerNo = String(targetOrPayload.customerNo || '').trim();
  if (!rowNo) return markCustomerSearchIndexDirty_(reason || 'MUTATION', 'rowNo 없음');
  try {
    const result = updateCustomerSearchIndexRowFastByPatch_(rowNo, customerNo, values);
    if (!result || result.ok === false) {
      return markCustomerSearchIndexDirty_(reason || 'MUTATION', result && result.reason ? result.reason : 'fast patch 실패');
    }
    return result;
  } catch (err) {
    return markCustomerSearchIndexDirty_(reason || 'MUTATION', err && err.message ? err.message : String(err));
  }
}

function updateCustomerSearchIndexFullAfterMutation_(targetOrPayload, reason) {
  targetOrPayload = targetOrPayload || {};
  const rowNo = Number(targetOrPayload.rowNo || targetOrPayload.masterRow || 0) || 0;
  if (!rowNo) return markCustomerSearchIndexDirty_(reason || 'FULL_REFRESH', 'rowNo 없음');
  try {
    const result = updateCustomerSearchIndexRow_(rowNo);
    if (!result || result.ok === false) {
      return markCustomerSearchIndexDirty_(reason || 'FULL_REFRESH', result && result.reason ? result.reason : 'full row refresh 실패');
    }
    return result;
  } catch (err) {
    return markCustomerSearchIndexDirty_(reason || 'FULL_REFRESH', err && err.message ? err.message : String(err));
  }
}

function buildCustomerSearchIndexObjectFromRow_(headers, row) {
  const obj = {};
  (headers || []).forEach(function(h, i) {
    h = String(h || '').trim();
    if (h) obj[h] = row[i] == null ? '' : row[i];
  });
  return obj;
}

function normalizeCustomerIndexCompareValue_(value) {
  return String(value == null ? '' : value).replace(/\r\n/g, '\n').trim();
}

function getCustomerSearchIndexConsistencyReport(options) {
  options = options || {};
  const limit = Math.max(1, Math.min(3000, Number(options.limit) || 500));
  const indexSheet = ensureCustomerSearchIndexSheet_(getWebAppDbSpreadsheet_());
  const indexLastRow = indexSheet.getLastRow();
  const width = Math.max(indexSheet.getLastColumn(), getCustomerSearchIndexHeadersK2_().length);
  const indexHeaders = indexSheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const indexMap = {};
  indexHeaders.forEach(function(h, i) { if (h) indexMap[h] = i; });

  const indexByRowNo = {};
  if (indexLastRow >= 2) {
    const indexValues = indexSheet.getRange(2, 1, indexLastRow - 1, width).getDisplayValues();
    indexValues.forEach(function(row, i) {
      const rowNo = Number(row[indexMap['rowNo']] || 0);
      if (rowNo) indexByRowNo[rowNo] = { indexRow: i + 2, obj: buildCustomerSearchIndexObjectFromRow_(indexHeaders, row) };
    });
  }

  const masterObjects = getMasterObjects_().filter(function(obj) {
    return isCustomerMasterObjectVisibleInListV70_(obj);
  }).slice(0, limit);
  const fields = [
    ['customerNo', '고객번호'],
    ['orderNo', '발주번호'],
    ['company', '회사명'],
    ['salesRep', '영업담당자'],
    ['status', '진행현황'],
    ['customerRank', '고객등급'],
    ['contact', '담당자'],
    ['phone', '전화번호'],
    ['directPhone', '직통번호'],
    ['email', '담당자 이메일'],
    ['vendor', '수행사'],
    ['finalQuote', '최종 견적가'],
    ['area', '연면적'],
    ['grade', '관리등급'],
    ['contractUnit', '계약단위'],
    ['contractStartDate', '계약시작일'],
    ['contractEndDate', '계약종료일'],
    ['s1Referrer', '제보자'],
    ['appointment', '관리자 선임 여부'],
    ['maintenance', '유지점검'],
    ['performance', '성능점검'],
    ['vat', '부가세'],
    ['discountRate', '할인율'],
    ['specialTerms', '용역신청서특약사항'],
    ['lastSent', '마지막발송'],
    ['sentAt', '발송일시']
  ];

  const mismatches = [];
  masterObjects.forEach(function(obj) {
    const rowNo = Number(obj.__rowNo || 0);
    if (!rowNo) return;
    const indexHit = indexByRowNo[rowNo];
    if (!indexHit) {
      mismatches.push({ rowNo: rowNo, customerNo: getCustomerListValue_(obj, 'customerNo'), company: getCustomerListValue_(obj, 'company'), type: 'MISSING_INDEX_ROW', diffs: [] });
      return;
    }
    const diffs = [];
    fields.forEach(function(pair) {
      const key = pair[0];
      const indexHeader = pair[1];
      let masterValue = key === 'memo'
        ? shortenTextForIndex_(getCustomerListValue_(obj, key), PORTAL_CONFIG.CUSTOMER_INDEX_MEMO_MAX_LENGTH || 350)
        : getCustomerListValue_(obj, key);
      if (key === 'specialTerms') masterValue = getCustomerIndexObjectValueK2_(obj, 'specialTerms');
      const indexValue = indexHit.obj[indexHeader] || '';
      if (normalizeCustomerIndexCompareValue_(masterValue) !== normalizeCustomerIndexCompareValue_(indexValue)) {
        diffs.push({ key: key, header: indexHeader, master: masterValue || '', index: indexValue || '' });
      }
    });
    if (diffs.length) mismatches.push({ rowNo: rowNo, customerNo: getCustomerListValue_(obj, 'customerNo'), company: getCustomerListValue_(obj, 'company'), type: 'VALUE_MISMATCH', diffs: diffs.slice(0, 10) });
  });

  return {
    ok: true,
    checked: masterObjects.length,
    indexRows: Math.max(0, indexLastRow - 1),
    mismatchCount: mismatches.length,
    mismatches: mismatches.slice(0, Math.max(1, Math.min(100, Number(options.maxResults) || 50))),
    meta: getCustomerSearchIndexMeta()
  };
}

function repairCustomerSearchIndexConsistency(options) {
  options = options || {};
  const limit = Math.max(1, Math.min(3000, Number(options.limit) || 500));
  const report = getCustomerSearchIndexConsistencyReport({ limit: limit, maxResults: limit });
  const targets = (report.mismatches || []).map(function(item) { return Number(item.rowNo) || 0; }).filter(Boolean);
  let repaired = 0;
  const errors = [];
  targets.forEach(function(rowNo) {
    try {
      const res = updateCustomerSearchIndexRow_(rowNo);
      if (res && res.ok) repaired++;
    } catch (err) {
      errors.push({ rowNo: rowNo, message: err && err.message ? err.message : String(err) });
    }
  });
  return { ok: !errors.length, checked: report.checked, targetCount: targets.length, repaired: repaired, errors: errors, before: report };
}

function rebuildCustomerSearchIndex(options) {
  options = options || {};
  const maxWaitMs = Math.max(500, Math.min(10000, Number(options.maxWaitMs) || 5000));
  const lock = LockService.getScriptLock();

  // v31 FIX: waitLock(30000)은 웹앱 동시 접속 때 사용자가 30초 기다리다가
  // '잠금 시간초과'를 맞기 쉬워서, 짧게 시도 후 안전하게 반환하도록 변경합니다.
  if (!lock.tryLock(maxWaitMs)) {
    return {
      ok: false,
      locked: true,
      message: '검색인덱스_DB 생성/갱신이 이미 진행 중입니다. 잠시 후 다시 시도하세요.'
    };
  }

  try {
    const masterSs = getMasterSpreadsheet_();
    const masterSheet = masterSs.getSheetByName(PORTAL_CONFIG.MASTER_SHEET_NAME);
    if (!masterSheet) throw new Error('마스터시트(신규)를 찾지 못했습니다.');
    const indexSheet = ensureCustomerSearchIndexSheet_(getWebAppDbSpreadsheet_());
    const masterRows = getMasterObjects_().filter(function(obj) {
      // v70: 회사명이 없는 준비 행은 검색인덱스_DB 생성 대상에서 제외합니다.
      return isCustomerMasterObjectVisibleInListV70_(obj);
    });
    const now = new Date();
    const rows = masterRows.map(function(obj) { return buildCustomerSearchIndexRow_(obj, now); });

    indexSheet.clearContents();
    indexSheet.getRange(1, 1, 1, getCustomerSearchIndexHeadersK2_().length).setValues([getCustomerSearchIndexHeadersK2_()]);
    if (rows.length) {
      formatCustomerSearchIndexContractCellsP340_(indexSheet, 2, rows.length);
      indexSheet.getRange(2, 1, rows.length, getCustomerSearchIndexHeadersK2_().length).setValues(rows);
    }
    indexSheet.setFrozenRows(1);
    indexSheet.getRange(1, 1, 1, getCustomerSearchIndexHeadersK2_().length).setFontWeight('bold').setBackground('#f2f4f7');
    // STEP23/P320: autoResizeColumns는 대량 인덱스 재생성에서 체감 지연이 커서 생략합니다.
    // 검색 성능/데이터 정합성과 무관한 서식 작업입니다.
    if (!options.skipFormat) { try { indexSheet.autoResizeColumns(1, getCustomerSearchIndexHeadersK2_().length); } catch (err) {} }

    const version = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
    const builtAt = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    PropertiesService.getScriptProperties().setProperties({
      CUSTOMER_SEARCH_INDEX_VERSION: version,
      CUSTOMER_SEARCH_INDEX_BUILT_AT: builtAt,
      CUSTOMER_SEARCH_INDEX_DIRTY: 'N',
      CUSTOMER_SEARCH_INDEX_DIRTY_REASON: '',
      CUSTOMER_SEARCH_INDEX_DIRTY_AT: '',
      CUSTOMER_SEARCH_INDEX_SCHEMA_VERSION: CUSTOMER_SEARCH_INDEX_SCHEMA_VERSION_P360
    }, true);
    SpreadsheetApp.flush();
    return { ok: true, rebuilt: rows.length, version: version, builtAt: builtAt, sheetName: PORTAL_CONFIG.CUSTOMER_INDEX_SHEET_NAME };
  } finally {
    try { lock.releaseLock(); } catch (err) {}
  }
}

function getCustomerSearchIndexRows_() {
  const sheet = ensureCustomerSearchIndexSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), getCustomerSearchIndexHeadersK2_().length)).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const map = {};
  headers.forEach(function(h, i) { if (h) map[h] = i; });
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getDisplayValues();
  return values.map(function(row) {
    const fullAddress = cellByIndexHeader_(row, map, '주소');
    const item = {
      rowNo: Number(cellByIndexHeader_(row, map, 'rowNo')) || 0,
      customerNo: cellByIndexHeader_(row, map, '고객번호'),
      orderNo: cellByIndexHeader_(row, map, '발주번호'),
      company: cellByIndexHeader_(row, map, '회사명'),
      salesRep: cellByIndexHeader_(row, map, '영업담당자'),
      status: cellByIndexHeader_(row, map, '진행현황'),
      customerRank: cellByIndexHeader_(row, map, '고객등급'),
      contact: cellByIndexHeader_(row, map, '담당자'),
      phone: cellByIndexHeader_(row, map, '전화번호'),
      directPhone: cellByIndexHeader_(row, map, '직통번호'),
      email: cellByIndexHeader_(row, map, '담당자 이메일'),
      vendor: cellByIndexHeader_(row, map, '수행사'),
      finalQuote: cellByIndexHeader_(row, map, '최종 견적가'),
      memo: cellByIndexHeader_(row, map, '메모요약'),
      address: shortenAddressForList_(fullAddress),
      fullAddress: fullAddress,
      searchText: String(cellByIndexHeader_(row, map, '검색문자열') || '').toLowerCase(),
      indexUpdatedAt: cellByIndexHeader_(row, map, '인덱스갱신시각'),

      // PATCH K-2 detail-lite fields. 검색/빠른 표시용 snapshot이며 저장 기준은 아닙니다.
      firstRegisteredAt: cellByIndexHeader_(row, map, '마스터시트 최초등록일'),
      region: cellByIndexHeader_(row, map, '지역구분'),
      area: cellByIndexHeader_(row, map, '연면적'),
      grade: cellByIndexHeader_(row, map, '관리등급'),
      buildingType: cellByIndexHeader_(row, map, '건물 유형'),
      contractUnit: cellByIndexHeader_(row, map, '계약단위'),
      contractStartDate: cellByIndexHeader_(row, map, '계약시작일'),
      contractEndDate: cellByIndexHeader_(row, map, '계약종료일'),
      s1Referrer: cellByIndexHeader_(row, map, '제보자'),
      appointment: cellByIndexHeader_(row, map, '관리자 선임 여부'),
      maintenance: cellByIndexHeader_(row, map, '유지점검'),
      performance: cellByIndexHeader_(row, map, '성능점검'),
      vat: cellByIndexHeader_(row, map, '부가세'),
      discountRate: cellByIndexHeader_(row, map, '할인율'),
      specialTerms: cellByIndexHeader_(row, map, '용역신청서특약사항'),
      lastSent: cellByIndexHeader_(row, map, '마지막발송'),
      sentAt: cellByIndexHeader_(row, map, '발송일시'),
      indexLite: cellByIndexHeader_(row, map, '상세Lite여부'),
      masterVersion: cellByIndexHeader_(row, map, '마스터원본버전'),
      masterUpdatedAt: cellByIndexHeader_(row, map, '원본수정시각'),
      masterEditor: cellByIndexHeader_(row, map, '최종수정자'),
      __source: 'index',
      __summary: true,
      __lite: true,
      __masterBacked: false
    };
    return item;
  }).filter(function(item) {
    // v70: 마스터시트에 고객번호만 미리 적어둔 준비 행은 고객 목록에서 제외합니다.
    // 고객상세검색의 유효 고객 기준은 `회사명`이 비어 있지 않은 행입니다.
    return item.rowNo && isCustomerListCompanyPresentV70_(item.company);
  });
}

function isCustomerListCompanyPresentV70_(value) {
  return String(value == null ? '' : value).trim() !== '';
}

function isCustomerMasterObjectVisibleInListV70_(obj) {
  obj = obj || {};
  return isCustomerListCompanyPresentV70_(getCustomerListValue_(obj, 'company'));
}


function filterCustomerIndexRows_(rows, keyword) {
  keyword = normalizeSearchKeyword_(keyword);
  rows = rows || [];
  if (!keyword) return rows;

  // v67: 검색어가 숫자만이면 고객번호 exact match를 최우선으로 사용합니다.
  // 고객번호가 존재하면 전화번호/금액 등에 우연히 포함된 결과는 뒤섞지 않습니다.
  if (/^[0-9]+$/.test(keyword)) {
    const exact = rows.filter(function(r) {
      return normalizeCustomerNoForKey_(r && r.customerNo) === keyword;
    });
    if (exact.length) return exact;
  }

  return rows.filter(function(r) {
    return doesCustomerRowMatchKeyword_(r, keyword);
  });
}


function sortCustomerIndexRowsV67_(rows, sortMode) {
  rows = Array.isArray(rows) ? rows.slice() : [];
  sortMode = sortMode || 'customerNoDigitsDesc';
  return rows.map(function(row, idx) { return { row: row, idx: idx }; }).sort(function(a, b) {
    const cmp = compareCustomerIndexRowsV67_(a.row, b.row, sortMode);
    return cmp || (a.idx - b.idx);
  }).map(function(item) { return item.row; });
}

function compareCustomerIndexRowsV67_(a, b, sortMode) {
  const headerSort = parseCustomerHeaderSortModeP15_(sortMode);
  if (headerSort) return compareCustomerIndexFieldGenericP15_(a && a[headerSort.key], b && b[headerSort.key], headerSort.dir);
  if (sortMode === 'customerNoAsc') return compareNumberWithMissingLastV67_(parseCustomerNoNumberV67_(a && a.customerNo), parseCustomerNoNumberV67_(b && b.customerNo), true);
  if (sortMode === 'customerNoDesc') return compareNumberWithMissingLastV67_(parseCustomerNoNumberV67_(a && a.customerNo), parseCustomerNoNumberV67_(b && b.customerNo), false);
  if (sortMode === 'finalQuoteAsc') return compareNumberWithMissingLastV67_(parseMoneyNumberV67_(a && a.finalQuote), parseMoneyNumberV67_(b && b.finalQuote), true);
  if (sortMode === 'finalQuoteDesc') return compareNumberWithMissingLastV67_(parseMoneyNumberV67_(a && a.finalQuote), parseMoneyNumberV67_(b && b.finalQuote), false);
  if (sortMode === 'contactLatest') return compareNumberWithMissingLastV67_(parseLatestContactTimeV67_(a && a.memo), parseLatestContactTimeV67_(b && b.memo), false);
  if (sortMode === 'contactOldest') return compareNumberWithMissingLastV67_(parseLatestContactTimeV67_(a && a.memo), parseLatestContactTimeV67_(b && b.memo), true);
  return compareCustomerNoDigitsDescV67_(a && a.customerNo, b && b.customerNo);
}


/**
 * PATCH P1-15B: 고객 상세 검색 표 헤더 클릭 정렬용 sortMode 파서입니다.
 * 클라이언트에서 전달하는 sortMode 형식: header:<검색인덱스키>:asc|desc
 * 기존 고객번호/견적가/최근컨택 정렬 모드는 그대로 compareCustomerIndexRowsV67_에서 처리합니다.
 */
function parseCustomerHeaderSortModeP15_(sortMode) {
  const m = String(sortMode == null ? '' : sortMode).match(/^header:([A-Za-z0-9_\-]+):(asc|desc)$/);
  if (!m) return null;
  const key = m[1];
  const dir = m[2];
  const allowed = {
    favorite: true,
    customerNo: true,
    salesRep: true,
    company: true,
    status: true,
    contact: true,
    phone: true,
    directPhone: true,
    vendor: true,
    finalQuote: true,
    memo: true,
    customerRank: true,
    grade: true,
    region: true,
    area: true,
    buildingType: true,
    contractUnit: true,
    email: true
  };
  if (!allowed[key]) return null;
  return { key: key, dir: dir };
}

function compareCustomerIndexFieldGenericP15_(a, b, dir) {
  const asc = String(dir || 'asc') !== 'desc';
  const av = normalizeCustomerSortValueP15_(a);
  const bv = normalizeCustomerSortValueP15_(b);
  if (av.missing && bv.missing) return 0;
  if (av.missing) return 1;
  if (bv.missing) return -1;

  let cmp;
  if (av.type === 'number' && bv.type === 'number') {
    cmp = av.value - bv.value;
  } else if (av.type === 'date' && bv.type === 'date') {
    cmp = av.value - bv.value;
  } else {
    cmp = String(av.text).localeCompare(String(bv.text), 'ko', { numeric: true, sensitivity: 'base' });
  }
  return asc ? cmp : -cmp;
}

function normalizeCustomerSortValueP15_(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw || raw === '-') return { missing: true, text: '' };

  // 금액, 면적, 고객번호처럼 숫자 의미가 강한 값은 숫자로 정렬합니다.
  const numericRaw = raw.replace(/[₩￦원,\s]/g, '').replace(/㎡/g, '');
  if (/^-?\d+(?:\.\d+)?$/.test(numericRaw)) {
    return { missing: false, type: 'number', value: Number(numericRaw), text: raw };
  }

  // 26.06.23 08:27 / 2026-06-23 08:27 등 날짜형 메모 일부도 가능하면 날짜로 정렬합니다.
  const dt = parseCustomerSortDateP15_(raw);
  if (!isNaN(dt)) return { missing: false, type: 'date', value: dt, text: raw };

  return { missing: false, type: 'text', text: raw };
}

function parseCustomerSortDateP15_(text) {
  const raw = String(text == null ? '' : text).trim();
  let m = raw.match(/((?:20)?\d{2})[.\/-](\d{1,2})[.\/-](\d{1,2})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (m) {
    let y = Number(m[1]);
    if (y < 100) y += 2000;
    const t = new Date(y, Number(m[2]) - 1, Number(m[3]), Number(m[4] || 0), Number(m[5] || 0), 0, 0).getTime();
    return t;
  }
  m = raw.match(/(\d{1,2})[.\/-](\d{1,2})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (m) {
    const nowYear = new Date().getFullYear();
    const t = new Date(nowYear, Number(m[1]) - 1, Number(m[2]), Number(m[3] || 0), Number(m[4] || 0), 0, 0).getTime();
    return t;
  }
  return NaN;
}

function compareCustomerNoDigitsDescV67_(a, b) {
  const pa = parseCustomerNoPartsV67_(a);
  const pb = parseCustomerNoPartsV67_(b);
  if (pa.missing && pb.missing) return 0;
  if (pa.missing) return 1;
  if (pb.missing) return -1;
  if (pb.digits !== pa.digits) return pb.digits - pa.digits;
  return pb.number - pa.number;
}

function parseCustomerNoPartsV67_(value) {
  const raw = String(value == null ? '' : value).replace(/[^0-9]/g, '');
  if (!raw) return { missing: true, digits: 0, number: 0 };
  return { missing: false, digits: raw.length, number: Number(raw) || 0 };
}

function parseCustomerNoNumberV67_(value) {
  const raw = String(value == null ? '' : value).replace(/[^0-9]/g, '');
  return raw ? Number(raw) : NaN;
}

function parseMoneyNumberV67_(value) {
  const raw = String(value == null ? '' : value).replace(/[^0-9.\-]/g, '');
  if (!raw || raw === '-' || raw === '.') return NaN;
  return Number(raw);
}

function compareNumberWithMissingLastV67_(a, b, asc) {
  const ma = isNaN(a);
  const mb = isNaN(b);
  if (ma && mb) return 0;
  if (ma) return 1;
  if (mb) return -1;
  return asc ? (a - b) : (b - a);
}

function parseLatestContactTimeV67_(memo) {
  const text = String(memo == null ? '' : memo);
  if (!text.trim()) return NaN;
  const nowYear = new Date().getFullYear();
  const candidates = [];
  const patterns = [
    /(?:^|[^0-9])((?:20)?\d{2})[.\/-](\d{1,2})[.\/-](\d{1,2})(?:\s+(\d{1,2}):(\d{2}))?/g,
    /(?:^|[^0-9])(\d{1,2})[.\/-](\d{1,2})(?:\s+(\d{1,2}):(\d{2}))?/g
  ];
  patterns.forEach(function(re, idx) {
    let m;
    while ((m = re.exec(text)) !== null) {
      let year, month, day, hour, minute;
      if (idx === 0) {
        year = Number(m[1]);
        if (year < 100) year += 2000;
        month = Number(m[2]);
        day = Number(m[3]);
        hour = Number(m[4] || 0);
        minute = Number(m[5] || 0);
      } else {
        year = nowYear;
        month = Number(m[1]);
        day = Number(m[2]);
        hour = Number(m[3] || 0);
        minute = Number(m[4] || 0);
      }
      if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        const t = new Date(year, month - 1, day, hour || 0, minute || 0, 0, 0).getTime();
        if (!isNaN(t)) candidates.push(t);
      }
    }
  });
  if (!candidates.length) return NaN;
  return Math.max.apply(null, candidates);
}

function doesCustomerRowMatchKeyword_(r, keyword) {
  keyword = normalizeSearchKeyword_(keyword);
  if (!keyword) return true;

  const fields = [
    r && r.customerNo,
    r && r.company,
    r && (r.fullAddress || r.address),
    r && r.contact,
    r && r.phone,
    r && r.directPhone,
    r && r.email,
    r && r.salesRep,
    r && r.vendor,
    r && r.status,
    r && r.finalQuote,
    r && r.region,
    r && r.area,
    r && r.grade,
    r && r.buildingType,
    r && r.contractUnit,
    r && r.appointment,
    r && r.maintenance,
    r && r.performance,
    r && r.vat
  ];

  const text = normalizeCustomerSearchText_(fields.join(' '));
  const compactText = compactCustomerSearchText_(text);
  const tokens = keyword.split(/\s+/).filter(Boolean);

  return tokens.every(function(token) {
    const t = normalizeCustomerSearchText_(token);
    const compact = compactCustomerSearchText_(t);
    return text.indexOf(t) >= 0 || (compact && compactText.indexOf(compact) >= 0);
  });
}

function normalizeCustomerSearchText_(value) {
  return String(value || '').toLowerCase().trim();
}

function compactCustomerSearchText_(value) {
  return normalizeCustomerSearchText_(value).replace(/[\s\-_.()\[\]{}\/\\]+/g, '');
}

function ensureCustomerSearchIndexSheet_(ss) {
  ss = ss || getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(PORTAL_CONFIG.CUSTOMER_INDEX_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PORTAL_CONFIG.CUSTOMER_INDEX_SHEET_NAME);
    sheet.getRange(1, 1, 1, getCustomerSearchIndexHeadersK2_().length).setValues([getCustomerSearchIndexHeadersK2_()]);
    sheet.setFrozenRows(1);
  }
  const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), getCustomerSearchIndexHeadersK2_().length)).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const existing = {};
  current.forEach(function(h, i) { if (h) existing[h] = i + 1; });
  let col = Math.max(sheet.getLastColumn(), 1);
  getCustomerSearchIndexHeadersK2_().forEach(function(h) {
    if (!existing[h]) { col += 1; sheet.getRange(1, col).setValue(h); existing[h] = col; }
  });
  sheet.getRange(1, 1, 1, getCustomerSearchIndexHeadersK2_().length).setFontWeight('bold').setBackground('#f2f4f7');
  return sheet;
}

function formatCustomerSearchIndexContractCellsP340_(sheet, startRow, numRows) {
  if (!sheet || !startRow || !numRows) return;
  try {
    const width = Math.max(sheet.getLastColumn(), getCustomerSearchIndexHeadersK2_().length);
    const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
    ['계약단위', '유지점검', '성능점검'].forEach(function(header) {
      const col = headers.indexOf(header) + 1;
      if (col > 0) sheet.getRange(startRow, col, numRows, 1).setNumberFormat('@');
    });
  } catch (err) {}
}

function buildCustomerSearchIndexRow_(obj, now) {
  now = now || new Date();
  obj = obj || {};
  const customerNo = getCustomerListValue_(obj, 'customerNo');
  const orderNo = getCustomerListValue_(obj, 'orderNo');
  const company = getCustomerListValue_(obj, 'company');
  const salesRep = getCustomerListValue_(obj, 'salesRep');
  const status = getCustomerListValue_(obj, 'status');
  const customerRank = getCustomerListValue_(obj, 'customerRank');
  const contact = getCustomerListValue_(obj, 'contact');
  const phone = getCustomerListValue_(obj, 'phone');
  const directPhone = getCustomerListValue_(obj, 'directPhone');
  const email = getCustomerListValue_(obj, 'email');
  const vendor = getCustomerListValue_(obj, 'vendor');
  const finalQuote = getCustomerListValue_(obj, 'finalQuote');
  const memoRaw = getCustomerListValue_(obj, 'memo');
  const memo = shortenTextForIndex_(memoRaw, PORTAL_CONFIG.CUSTOMER_INDEX_MEMO_MAX_LENGTH || 350);
  const address = getCustomerListValue_(obj, 'address');

  const firstRegisteredAt = getCustomerIndexObjectValueK2_(obj, 'firstRegisteredAt');
  const region = getCustomerIndexObjectValueK2_(obj, 'region');
  const area = getCustomerIndexObjectValueK2_(obj, 'area');
  const grade = getCustomerIndexObjectValueK2_(obj, 'grade');
  const buildingType = getCustomerIndexObjectValueK2_(obj, 'buildingType');
  const contractUnit = normalizePortalContractFieldForDbP280_('contractUnit', getCustomerIndexObjectValueK2_(obj, 'contractUnit'));
  const contractStartDate = formatPortalContractDateForDisplayP420_(getCustomerIndexObjectValueK2_(obj, 'contractStartDate'));
  const contractEndDate = formatPortalContractDateForDisplayP420_(getCustomerIndexObjectValueK2_(obj, 'contractEndDate'));
  const appointment = getCustomerIndexObjectValueK2_(obj, 'appointment');
  const maintenance = normalizePortalContractFieldForDbP280_('maintenance', getCustomerIndexObjectValueK2_(obj, 'maintenance'));
  const performance = normalizePortalContractFieldForDbP280_('performance', getCustomerIndexObjectValueK2_(obj, 'performance'));
  const vat = getCustomerIndexObjectValueK2_(obj, 'vat');
  const discountRate = getCustomerIndexObjectValueK2_(obj, 'discountRate');
  const specialTerms = getCustomerIndexObjectValueK2_(obj, 'specialTerms');
  const s1Referrer = getCustomerIndexObjectValueK2_(obj, 's1Referrer');
  const lastSent = getCustomerIndexObjectValueK2_(obj, 'lastSent');
  const sentAt = getCustomerIndexObjectValueK2_(obj, 'sentAt');
  const masterUpdatedAt = getCustomerIndexObjectValueK2_(obj, ['수정일시', '최종수정일시', '수정 시각']);
  const masterVersion = getCustomerIndexObjectValueK2_(obj, ['수정버전', '마스터수정버전']);
  const masterEditor = getCustomerIndexObjectValueK2_(obj, ['최종수정자', '수정자']);

  // v35 FIX: 검색문자열에는 메모/작성자 로그를 넣지 않습니다.
  // PATCH K-2: 상세 lite 필드 중 검색에 실질적으로 필요한 값은 searchText에 포함합니다.
  const searchText = shortenTextForIndex_([
    customerNo, orderNo, company, salesRep, status, customerRank, contact, phone, directPhone, email, vendor, finalQuote, address,
    region, area, grade, buildingType, contractUnit, contractStartDate, contractEndDate, s1Referrer, appointment, maintenance, performance, vat
  ].join(' ').toLowerCase(), PORTAL_CONFIG.CUSTOMER_INDEX_SEARCH_TEXT_MAX_LENGTH || 1500);
  const ts = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const rowMap = {
    'rowNo': obj.__rowNo || '',
    '고객번호': customerNo,
    '발주번호': orderNo,
    '회사명': company,
    '영업담당자': salesRep,
    '진행현황': status,
    '고객등급': customerRank,
    '담당자': contact,
    '전화번호': phone,
    '직통번호': directPhone,
    '담당자 이메일': email,
    '수행사': vendor,
    '최종 견적가': finalQuote,
    '메모요약': memo,
    '주소': address,
    '검색문자열': searchText,
    '원본수정시각': masterUpdatedAt || '',
    '인덱스갱신시각': ts,
    '마스터시트 최초등록일': firstRegisteredAt,
    '지역구분': region,
    '연면적': area,
    '관리등급': grade,
    '건물 유형': buildingType,
    '계약단위': contractUnit,
    '계약시작일': contractStartDate,
    '계약종료일': contractEndDate,
    '제보자': s1Referrer,
    '관리자 선임 여부': appointment,
    '유지점검': maintenance,
    '성능점검': performance,
    '부가세': vat,
    '할인율': discountRate,
    '용역신청서특약사항': specialTerms,
    '마지막발송': lastSent,
    '발송일시': sentAt,
    '상세Lite여부': 'Y',
    '마스터원본버전': masterVersion || masterUpdatedAt || ts,
    '최종수정자': masterEditor || ''
  };
  return getCustomerSearchIndexHeadersK2_().map(function(h) { return rowMap[h] == null ? '' : rowMap[h]; });
}


function deleteCustomerSearchIndexRowByRowNoV70_(rowNo) {
  rowNo = Number(rowNo) || 0;
  if (!rowNo) return { ok: true, deleted: false, reason: 'invalid rowNo' };
  const indexSheet = ensureCustomerSearchIndexSheet_(getWebAppDbSpreadsheet_());
  const lastRow = indexSheet.getLastRow();
  if (lastRow < 2) return { ok: true, deleted: false, reason: 'empty index' };
  const headers = indexSheet.getRange(1, 1, 1, Math.max(indexSheet.getLastColumn(), getCustomerSearchIndexHeadersK2_().length)).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const rowNoColIdx = headers.indexOf('rowNo');
  if (rowNoColIdx < 0) return { ok: true, deleted: false, reason: 'rowNo header missing' };
  const values = indexSheet.getRange(2, rowNoColIdx + 1, lastRow - 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (Number(values[i][0]) === rowNo) {
      indexSheet.deleteRow(i + 2);
      return { ok: true, deleted: true, indexRow: i + 2 };
    }
  }
  return { ok: true, deleted: false, reason: 'not found' };
}

function updateCustomerSearchIndexRow_(rowNo) {
  rowNo = Number(rowNo) || 0;
  if (!rowNo || rowNo < PORTAL_CONFIG.DATA_START_ROW) return { ok: false, reason: 'invalid rowNo' };
  const masterSs = getMasterSpreadsheet_();
  const masterSheet = masterSs.getSheetByName(PORTAL_CONFIG.MASTER_SHEET_NAME);
  if (!masterSheet) throw new Error('마스터시트(신규)를 찾지 못했습니다.');
  const obj = readMasterRowObject_(masterSheet, rowNo);
  obj.__rowNo = rowNo;
  if (!isCustomerMasterObjectVisibleInListV70_(obj)) {
    const deleted = deleteCustomerSearchIndexRowByRowNoV70_(rowNo);
    const touched = touchCustomerSearchIndexVersion_(new Date());
    return {
      ok: true,
      skipped: true,
      reason: '회사명 공란 행은 고객 목록/검색인덱스 표시 대상에서 제외',
      rowNo: rowNo,
      deleted: deleted.deleted,
      version: touched.version,
      builtAt: touched.builtAt
    };
  }
  const indexSheet = ensureCustomerSearchIndexSheet_(getWebAppDbSpreadsheet_());
  const rowValues = buildCustomerSearchIndexRow_(obj, new Date());
  const lastRow = indexSheet.getLastRow();
  let targetRow = 0;
  if (lastRow >= 2) {
    const rowNoValues = indexSheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    for (let i = 0; i < rowNoValues.length; i++) if (Number(rowNoValues[i][0]) === rowNo) { targetRow = i + 2; break; }
  }
  if (!targetRow) targetRow = Math.max(2, indexSheet.getLastRow() + 1);
  formatCustomerSearchIndexContractCellsP340_(indexSheet, targetRow, 1);
  indexSheet.getRange(targetRow, 1, 1, getCustomerSearchIndexHeadersK2_().length).setValues([rowValues]);
  const touched = touchCustomerSearchIndexVersion_(new Date());
  return { ok: true, rowNo: rowNo, indexRow: targetRow, version: touched.version, builtAt: touched.builtAt };
}


function updateCustomerSearchIndexRowFastByPatch_(rowNo, customerNo, values) {
  // PATCH D: 상세 저장 후 검색인덱스_DB를 마스터 재조회 없이 patch 값만으로 빠르게 갱신합니다.
  rowNo = Number(rowNo) || 0;
  values = values || {};
  if (!rowNo || rowNo < PORTAL_CONFIG.DATA_START_ROW) return { ok: false, reason: 'invalid rowNo' };

  const indexSheet = ensureCustomerSearchIndexSheet_(getWebAppDbSpreadsheet_());
  const lastRow = indexSheet.getLastRow();
  if (lastRow < 2) return updateCustomerSearchIndexRow_(rowNo);

  const width = Math.max(indexSheet.getLastColumn(), getCustomerSearchIndexHeadersK2_().length);
  const headers = indexSheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const hmap = {};
  headers.forEach(function(h, i) { if (h) hmap[h] = i + 1; });
  const rowNoCol = hmap['rowNo'];
  const customerNoCol = hmap['고객번호'];
  if (!rowNoCol) return updateCustomerSearchIndexRow_(rowNo);

  let targetRow = 0;
  const cno = String(customerNo || '').trim();
  if (cno && customerNoCol) {
    const cnoValues = indexSheet.getRange(2, customerNoCol, lastRow - 1, 1).getDisplayValues();
    for (let i = 0; i < cnoValues.length; i++) {
      if (String(cnoValues[i][0] || '').trim() === cno) { targetRow = i + 2; break; }
    }
  }
  if (!targetRow) {
    const rowNoValues = indexSheet.getRange(2, rowNoCol, lastRow - 1, 1).getDisplayValues();
    for (let i = 0; i < rowNoValues.length; i++) {
      if (Number(rowNoValues[i][0]) === rowNo) { targetRow = i + 2; break; }
    }
  }
  if (!targetRow) return updateCustomerSearchIndexRow_(rowNo);

  const row = indexSheet.getRange(targetRow, 1, 1, width).getDisplayValues()[0];
  const setByHeader = function(header, value) {
    const col = hmap[header];
    if (!col) return;
    row[col - 1] = String(value == null ? '' : value);
  };

  const keyToHeader = {
    customerNo: '고객번호',
    orderNo: '발주번호',
    company: '회사명',
    salesRep: '영업담당자',
    status: '진행현황',
    customerRank: '고객등급',
    contact: '담당자',
    phone: '전화번호',
    directPhone: '직통번호',
    email: '담당자 이메일',
    vendor: '수행사',
    finalQuote: '최종 견적가',
    memo: '메모요약',
    address: '주소',
    firstRegisteredAt: '마스터시트 최초등록일',
    region: '지역구분',
    area: '연면적',
    grade: '관리등급',
    buildingType: '건물 유형',
    contractUnit: '계약단위',
    appointment: '관리자 선임 여부',
    maintenance: '유지점검',
    performance: '성능점검',
    vat: '부가세',
    discountRate: '할인율',
    specialTerms: '용역신청서특약사항',
    s1Referrer: '제보자',
    lastSent: '마지막발송',
    sentAt: '발송일시'
  };

  Object.keys(keyToHeader).forEach(function(key) {
    if (!Object.prototype.hasOwnProperty.call(values, key)) return;
    let v = key === 'memo'
      ? shortenTextForIndex_(values[key], PORTAL_CONFIG.CUSTOMER_INDEX_MEMO_MAX_LENGTH || 350)
      : values[key];
    if (key === 'contractUnit' || key === 'maintenance' || key === 'performance') {
      v = normalizePortalContractFieldForDbP280_(key, v);
    }
    setByHeader(keyToHeader[key], v);
  });
  setByHeader('rowNo', rowNo);
  if (cno) setByHeader('고객번호', cno);

  const getH = function(header) {
    const col = hmap[header];
    return col ? String(row[col - 1] || '') : '';
  };
  const searchText = shortenTextForIndex_([
    getH('고객번호'), getH('발주번호'), getH('회사명'), getH('영업담당자'), getH('진행현황'), getH('고객등급'), getH('담당자'),
    getH('전화번호'), getH('직통번호'), getH('담당자 이메일'), getH('수행사'), getH('최종 견적가'), getH('주소'),
    getH('지역구분'), getH('연면적'), getH('관리등급'), getH('건물 유형'), getH('계약단위'),
    getH('계약시작일'), getH('계약종료일'), getH('제보자'), getH('관리자 선임 여부'), getH('유지점검'), getH('성능점검'), getH('부가세')
  ].join(' ').toLowerCase(), PORTAL_CONFIG.CUSTOMER_INDEX_SEARCH_TEXT_MAX_LENGTH || 1500);
  setByHeader('검색문자열', searchText);

  const now = new Date();
  const ts = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  setByHeader('인덱스갱신시각', ts);
  setByHeader('상세Lite여부', 'Y');
  setByHeader('원본수정시각', values.__metaUpdatedAt || getH('원본수정시각'));
  setByHeader('마스터원본버전', values.__metaMasterVersion || ts);
  setByHeader('최종수정자', values.__metaEditor || getH('최종수정자'));

  formatCustomerSearchIndexContractCellsP340_(indexSheet, targetRow, 1);
  indexSheet.getRange(targetRow, 1, 1, width).setValues([row]);
  const touched = touchCustomerSearchIndexVersion_(now);
  return { ok: true, rowNo: rowNo, indexRow: targetRow, version: touched.version, builtAt: touched.builtAt, fastPatch: true };
}

function getCustomerDetailsBatch(rowNos) {
  // STEP8 / P240: 상세 프리패치 권한필터 성능 개선
  // 기존 구현은 rowNo마다 고객 접근권한 검사를 별도 호출해
  // 권한 확인 단계에서 마스터 행을 다시 읽었습니다. 여기서는 시트/헤더/권한을 1회만 준비하고,
  // 실제 상세 데이터 읽기와 같은 rowValues에서 권한 필터까지 처리합니다.
  rowNos = Array.isArray(rowNos) ? rowNos.map(Number).filter(Boolean) : [];
  const unique = [];
  const seen = {};
  rowNos.forEach(function(rowNo) {
    if (!rowNo || rowNo < PORTAL_CONFIG.DATA_START_ROW || seen[rowNo]) return;
    seen[rowNo] = true;
    unique.push(rowNo);
  });

  const limit = Number(PORTAL_CONFIG.CUSTOMER_DETAIL_PREFETCH_BATCH_SIZE || 50);
  const targets = unique.slice(0, Math.max(1, Math.min(100, limit)));
  if (!targets.length) return [];

  const ss = getMasterSpreadsheet_();
  const sheet = ss.getSheetByName(PORTAL_CONFIG.MASTER_SHEET_NAME);
  if (!sheet) throw new Error('마스터시트(신규)를 찾지 못했습니다.');

  const headerMap = getHeaderMap_(sheet);
  const lastCol = sheet.getLastColumn();
  const perm = getPortalCurrentPermission_();
  const minRow = Math.min.apply(null, targets);
  const maxRow = Math.max.apply(null, targets);
  const span = maxRow - minRow + 1;
  const targetSet = {};
  targets.forEach(function(r) { targetSet[r] = true; });

  const resultByRow = {};
  const orderLookupP260 = (typeof getPortalCustomerOrderLookupP260_ === 'function')
    ? getPortalCustomerOrderLookupP260_({ force: false })
    : null;

  const addRowDetail = function(rowNo, rowValues) {
    try {
      const obj = buildMasterRowObjectFromValues_(headerMap, rowValues);
      obj.__rowNo = rowNo;
      if (!isPortalCustomerRowAllowedForPermission_(obj, perm)) return;
      resultByRow[rowNo] = buildCustomerDetailFromObj_(obj, rowNo, { includeLogs: false, includeRaw: false, lite: true, includeOrderInfo: true, orderLookup: orderLookupP260 });
    } catch (err) {
      resultByRow[rowNo] = { rowNo: rowNo, error: String(err && err.message || err) };
    }
  };

  // 보통 검색결과 20~50건은 행 번호가 어느 정도 인접하므로 범위 1회 읽기가 가장 빠릅니다.
  // 행 간격이 너무 넓으면 불필요한 대량 읽기를 피하기 위해 필요한 행만 개별 조회합니다.
  // 단, 권한 확인은 이미 읽은 rowValues와 동일한 객체로 처리해 중복 시트 접근을 제거합니다.
  if (span <= 350) {
    const values = sheet.getRange(minRow, 1, span, lastCol).getDisplayValues();
    values.forEach(function(rowValues, idx) {
      const rowNo = minRow + idx;
      if (!targetSet[rowNo]) return;
      addRowDetail(rowNo, rowValues);
    });
  } else {
    targets.forEach(function(rowNo) {
      try {
        const values = sheet.getRange(rowNo, 1, 1, lastCol).getDisplayValues()[0];
        addRowDetail(rowNo, values);
      } catch (err) {
        resultByRow[rowNo] = { rowNo: rowNo, error: String(err && err.message || err) };
      }
    });
  }

  // 기존 동작과 동일하게 권한이 없는 고객은 결과에서 제외합니다.
  // 요청 순서는 유지하되, 서버 오류가 난 행은 오류 객체로 반환합니다.
  return targets.map(function(rowNo) {
    return resultByRow[rowNo] || null;
  }).filter(Boolean);
}

function getCustomerDetailsBatchFast(rowNos) {
  return getCustomerDetailsBatch(rowNos);
}

function shortenTextForIndex_(value, maxLen) {
  value = String(value || '').trim();
  maxLen = Number(maxLen) || 500;
  return value.length <= maxLen ? value : value.slice(0, maxLen) + '...';
}

function getCustomerSearchIndexVersion_() {
  const sheet = ensureCustomerSearchIndexSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';
  const lastCol = Math.max(sheet.getLastColumn(), getCustomerSearchIndexHeadersK2_().length);
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const updatedCol = headers.indexOf('인덱스갱신시각') + 1;
  const masterUpdatedCol = headers.indexOf('원본수정시각') + 1;
  const masterVersionCol = headers.indexOf('마스터원본버전') + 1;
  const masterEditorCol = headers.indexOf('최종수정자') + 1;
  if (!updatedCol) return '';
  const values = sheet.getRange(2, updatedCol, lastRow - 1, 1).getDisplayValues();
  let maxText = '';
  values.forEach(function(r) { if (String(r[0] || '') > maxText) maxText = String(r[0] || ''); });
  return maxText ? maxText.replace(/[^0-9]/g, '') : '';
}

function normalizeSearchKeyword_(keyword) {
  return String(keyword || '').trim().toLowerCase();
}

function buildMasterRowObjectFromValues_(headerMap, rowValues) {
  const obj = {};
  Object.keys(headerMap || {}).forEach(function(header) {
    obj[header] = rowValues[headerMap[header] - 1] || '';
  });
  return obj;
}

function getPortalCustomerMasterMetaP202_(sheet, rowNo, headerMap) {
  headerMap = headerMap || getHeaderMap_(sheet);
  const colUpdatedAt = findFirstExistingHeaderCol_(headerMap, ['수정일시', '최종수정일시', '수정 시각']);
  const colVersion = findFirstExistingHeaderCol_(headerMap, ['수정버전', '마스터수정버전']);
  const colEditor = findFirstExistingHeaderCol_(headerMap, ['최종수정자', '수정자']);
  return {
    updatedAt: colUpdatedAt ? String(sheet.getRange(rowNo, colUpdatedAt).getDisplayValue() || '').trim() : '',
    version: colVersion ? String(sheet.getRange(rowNo, colVersion).getDisplayValue() || '').trim() : '',
    editor: colEditor ? String(sheet.getRange(rowNo, colEditor).getDisplayValue() || '').trim() : ''
  };
}

function getPortalBaseMasterVersionFromPayloadP202_(payload) {
  payload = payload || {};
  return String(payload.baseMasterVersion || payload.masterVersion || payload.__masterVersion || payload.expectedMasterVersion || '').trim();
}

function makePortalStaleCustomerErrorP202_(currentMeta, baseVersion) {
  const msg = '다른 사용자가 이 고객 정보를 먼저 수정했습니다. 최신 내용을 다시 불러온 뒤 저장해 주세요.';
  const err = new Error(msg);
  err.code = 'PORTAL_STALE_CUSTOMER_VERSION';
  err.currentMasterVersion = currentMeta && currentMeta.version || '';
  err.currentMasterUpdatedAt = currentMeta && currentMeta.updatedAt || '';
  err.baseMasterVersion = baseVersion || '';
  return err;
}


function getPortalDetailDefByKeyP458_(key) {
  key = String(key || '');
  const defs = (typeof getPortalCustomerAllDetailDefsP436_ === 'function') ? getPortalCustomerAllDetailDefsP436_() : [];
  for (let i = 0; i < defs.length; i++) {
    if (defs[i] && defs[i].key === key) return defs[i];
  }
  return null;
}

function isPortalExpectedValuesStillFreshP458_(sheet, rowNo, payload) {
  payload = payload || {};
  const expected = payload.expectedValues || {};
  const values = payload.values || (Object.prototype.hasOwnProperty.call(payload, 'memo') ? { memo: payload.memo } : {});
  const keys = Object.keys(values).filter(function(k) { return Object.prototype.hasOwnProperty.call(expected, k); });
  if (!keys.length) return false;
  const headerMap = getHeaderMap_(sheet);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const currentDisplayRow = sheet.getRange(rowNo, 1, 1, lastCol).getDisplayValues()[0];
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const def = getPortalDetailDefByKeyP458_(key);
    if (!def) return false;
    const col = findFirstExistingHeaderCol_(headerMap, def.headers || []);
    if (!col) return false;
    const expectedText = getPortalMasterCompareTextP280_(key, expected[key]);
    const currentText = String(currentDisplayRow[col - 1] || '').trim();
    if (currentText !== expectedText) return false;
  }
  return true;
}

function getPortalOperationCacheKeyP458_(payload, suffix) {
  payload = payload || {};
  const opId = String(payload.clientOperationId || '').trim();
  if (!opId) return '';
  const customerNo = String(payload.customerNo || '').trim();
  const rowNo = String(payload.rowNo || '').trim();
  return 'P458_OP_' + (suffix || 'RESULT') + '_' + opId + '_' + rowNo + '_' + customerNo;
}

function getPortalCachedOperationResultP458_(payload) {
  const key = getPortalOperationCacheKeyP458_(payload, 'RESULT');
  if (!key) return null;
  try {
    const raw = CacheService.getScriptCache().get(key);
    if (!raw) return null;
    const res = JSON.parse(raw);
    if (res && typeof res === 'object') {
      res.ok = true;
      res.duplicateOperation = true;
      return res;
    }
  } catch (err) {}
  return null;
}

function putPortalCachedOperationResultP458_(payload, result) {
  const key = getPortalOperationCacheKeyP458_(payload, 'RESULT');
  if (!key || !result) return;
  try {
    const compact = {
      ok: true,
      rowNo: result.rowNo || payload.rowNo || '',
      customerNo: result.customerNo || payload.customerNo || '',
      changedFields: result.changedFields || [],
      changedKeys: result.changedKeys || [],
      values: result.values || {},
      changedValues: result.changedValues || {},
      masterVersion: result.masterVersion || '',
      masterUpdatedAt: result.masterUpdatedAt || '',
      masterEditor: result.masterEditor || '',
      savedAt: result.savedAt || new Date().toISOString(),
      verified: result.verified === true,
      fastPatch: result.fastPatch !== false,
      clientOperationId: result.clientOperationId || payload.clientOperationId || '',
      noSynchronousRefresh: result.noSynchronousRefresh === true,
      message: result.message || '저장 완료'
    };
    CacheService.getScriptCache().put(key, JSON.stringify(compact), 21600);
  } catch (err) {}
}

function assertPortalCustomerVersionFreshP202_(sheet, rowNo, payload) {
  const baseVersion = getPortalBaseMasterVersionFromPayloadP202_(payload);
  const currentMeta = getPortalCustomerMasterMetaP202_(sheet, rowNo);
  // 기존 데이터/구버전 클라이언트/수정버전 미기록 행은 우선 허용합니다.
  // P458: rowVersion이 달라도 사용자가 바꾸려는 필드의 expectedValue가 그대로면 저장을 살립니다.
  // 예: 메모 저장 후 상태값 저장처럼 서로 다른 필드의 연속 수정은 stale로 막지 않습니다.
  if (baseVersion && currentMeta.version && baseVersion !== currentMeta.version) {
    if (isPortalExpectedValuesStillFreshP458_(sheet, rowNo, payload)) {
      return currentMeta;
    }
    throw makePortalStaleCustomerErrorP202_(currentMeta, baseVersion);
  }
  return currentMeta;
}

function bumpPortalCustomerMasterMetaP202_(sheet, rowNo, reason) {
  // STEP40/P400: 메타 3칸은 저장 응답 체감속도에 직접 영향을 줍니다.
  // 기존처럼 setValue 3회를 따로 호출하지 않고, 인접 컬럼이면 setValues 1회로 처리합니다.
  const headerMap = getHeaderMap_(sheet);
  const updatedAtCol = ensureMasterColumn_(sheet, headerMap, '수정일시');
  const versionCol = ensureMasterColumn_(sheet, getHeaderMap_(sheet), '수정버전');
  const editorCol = ensureMasterColumn_(sheet, getHeaderMap_(sheet), '최종수정자');
  const now = new Date();
  const nowText = (typeof getPortalNowTextP201_ === 'function')
    ? getPortalNowTextP201_(now)
    : Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const version = ((typeof getPortalVersionTextP201_ === 'function')
    ? getPortalVersionTextP201_(now)
    : Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmssSSS')) + '-' + rowNo;
  let editor = '';
  try { editor = String(Session.getActiveUser().getEmail() || '').trim(); } catch (err) {}
  if (!editor) editor = 'webapp';

  writePortalCustomerMetaCellsP400_(sheet, rowNo, [
    { col: updatedAtCol, value: nowText },
    { col: versionCol, value: version },
    { col: editorCol, value: editor }
  ]);

  try {
    if (typeof markPortalMasterDataChangedP201_ === 'function') {
      markPortalMasterDataChangedP201_(String(reason || 'webapp-customer-save') + ' row=' + rowNo);
    }
  } catch (err) {}
  return { updatedAt: nowText, version: version, editor: editor };
}

function writePortalCustomerMetaCellsP400_(sheet, rowNo, cells) {
  cells = (cells || []).filter(function(c) { return c && Number(c.col) > 0; })
    .sort(function(a, b) { return Number(a.col) - Number(b.col); });
  if (!sheet || !rowNo || !cells.length) return;

  let blockStart = Number(cells[0].col);
  let blockValues = [cells[0].value];
  let prevCol = Number(cells[0].col);
  const flushBlock = function() {
    sheet.getRange(rowNo, blockStart, 1, blockValues.length).setValues([blockValues]);
  };

  for (let i = 1; i < cells.length; i++) {
    const col = Number(cells[i].col);
    if (col === prevCol + 1) {
      blockValues.push(cells[i].value);
      prevCol = col;
    } else {
      flushBlock();
      blockStart = col;
      blockValues = [cells[i].value];
      prevCol = col;
    }
  }
  flushBlock();
}

function queueCustomerSearchIndexRefreshAfterSaveP400_(rowNo, customerNo, changedKeys, meta, reason) {
  // STEP40/P400: 고객 저장 응답 전에 검색인덱스 행 갱신/로그 append까지 기다리지 않습니다.
  // 현재 화면은 클라이언트 optimistic patch로 즉시 최신화하고,
  // 검색인덱스는 dirty 표시만 해 백그라운드 최신화 대상으로 넘깁니다.
  try {
    if (typeof markCustomerSearchIndexDirty_ === 'function') {
      return markCustomerSearchIndexDirty_(reason || 'WEBAPP_CUSTOMER_SAVE', 'row=' + rowNo + ', cno=' + (customerNo || '') + ', keys=' + (changedKeys || []).join(','));
    }
  } catch (err) {
    Logger.log('검색인덱스 dirty 표시 실패: ' + (err && err.stack || err));
  }
  return { ok: true, deferred: true, rowNo: rowNo, customerNo: customerNo, meta: meta || null };
}

function shouldUseDeferredCustomerSavePostProcessP400_(changedKeys) {
  changedKeys = changedKeys || [];
  // P455: 일반 연락처/상태 저장은 기존처럼 dirty 표시 후 빠르게 반환합니다.
  // 단, 연면적·할인율·최종 견적가·계약조건 계열은 시트 수식/표시값을 바로 확정해야 하므로
  // 저장 응답 전에 마스터 단건 재조회 + 검색인덱스 행 갱신 경로를 탑니다.
  if (shouldRefreshIndexFromMasterAfterContractSaveP112_(changedKeys)) return false;
  return true;
}

function addPortalCustomerMetaToDetailP202_(detail, obj, rowNo) {
  detail = detail || {};
  obj = obj || {};
  const updatedAt = getCustomerIndexObjectValueK2_(obj, ['수정일시', '최종수정일시', '수정 시각']);
  const version = getCustomerIndexObjectValueK2_(obj, ['수정버전', '마스터수정버전']);
  const editor = getCustomerIndexObjectValueK2_(obj, ['최종수정자', '수정자']);
  detail.masterUpdatedAt = updatedAt || '';
  detail.masterVersion = version || '';
  detail.masterEditor = editor || '';
  detail.__masterUpdatedAt = updatedAt || '';
  detail.__masterVersion = version || '';
  detail.__masterEditor = editor || '';
  return detail;
}

function getCustomerDetail(rowNo) {
  // v56 PATCH A: 기존 rowNo 호출은 호환하되, 내부에서는 고객 target resolver를 거칩니다.
  const target = assertCustomerTarget_({ rowNo: rowNo }, '고객 상세조회', { readObject: true });
  assertPortalCanAccessCustomerTarget_(target, '고객 상세조회');
  return buildCustomerDetailFromObj_(target.obj, target.rowNo, { includeLogs: true, includeRaw: true, lite: false });
}

function getCustomerDetailByCustomerNo(customerNo, fallbackRowNo) {
  // v56 PATCH A: 신규 표준 상세조회 API. customerNo를 기준키로 우선 사용하고 rowNo는 검증용 보조값으로만 사용합니다.
  const target = assertCustomerTarget_({ customerNo: customerNo, rowNo: fallbackRowNo }, '고객 상세조회', { readObject: true });
  assertPortalCanAccessCustomerTarget_(target, '고객 상세조회');
  return buildCustomerDetailFromObj_(target.obj, target.rowNo, { includeLogs: true, includeRaw: true, lite: false });
}


/**
 * P453: 마스터시트 `발주번호` 기준 발주여부 정보 생성
 */
function buildPortalCustomerOrderInfoFromMasterOrderNoP453_(customerNo, company, orderNo, rowNo) {
  const no = String(orderNo == null ? '' : orderNo).trim();
  return {
    exists: !!no,
    contractNo: no,
    rowNo: 0,
    customerNo: String(customerNo || '').trim(),
    company: String(company || '').trim(),
    masterRowNo: Number(rowNo) || 0,
    source: 'masterOrderNo'
  };
}

function buildCustomerDetailFromObj_(obj, rowNo, options) {
  options = options || {};
  rowNo = Number(rowNo) || Number(obj && obj.__rowNo) || 0;
  obj = obj || {};
  obj.__rowNo = rowNo;

  const customerNo = getCustomerMasterHeaderValueK2_(obj, 'customerNo') || obj['고객번호'] || '';
  const orderNo = getCustomerMasterHeaderValueK2_(obj, 'orderNo') || obj['발주번호'] || '';
  const company = getCompanyValue_(obj);
  const status = getStatusValueFromObj_(obj);
  const customerRank = getCustomerMasterHeaderValueK2_(obj, 'customerRank');
  const address = getCustomerMasterHeaderValueK2_(obj, 'address');

  // PATCH M-FIX2: 펼침/상세 화면에서 바로 쓰는 계약조건 필드는 detail 객체에 직접 싣습니다.
  // 특히 관리등급은 마스터시트 O열의 '관리등급' 헤더 값을 기준으로 읽되, 열주소가 아니라 헤더명으로만 매핑합니다.
  const firstRegisteredAt = getCustomerMasterHeaderValueK2_(obj, 'firstRegisteredAt');
  const region = getCustomerMasterHeaderValueK2_(obj, 'region');
  const area = getCustomerMasterHeaderValueK2_(obj, 'area');
  const grade = getCustomerMasterHeaderValueK2_(obj, 'grade');
  const buildingType = getCustomerMasterHeaderValueK2_(obj, 'buildingType');
  const finalQuote = getCustomerMasterHeaderValueK2_(obj, 'finalQuote');
  const contractUnit = normalizePortalContractFieldForDbP280_('contractUnit', getCustomerMasterHeaderValueK2_(obj, 'contractUnit'));
  const contractStartDate = formatPortalContractDateForDisplayP420_(getCustomerMasterHeaderValueK2_(obj, 'contractStartDate'));
  const contractEndDate = formatPortalContractDateForDisplayP420_(getCustomerMasterHeaderValueK2_(obj, 'contractEndDate'));
  const appointment = getCustomerMasterHeaderValueK2_(obj, 'appointment');
  const maintenance = normalizePortalContractFieldForDbP280_('maintenance', getCustomerMasterHeaderValueK2_(obj, 'maintenance'));
  const performance = normalizePortalContractFieldForDbP280_('performance', getCustomerMasterHeaderValueK2_(obj, 'performance'));
  const vat = getCustomerMasterHeaderValueK2_(obj, 'vat');
  const discountRate = getCustomerMasterHeaderValueK2_(obj, 'discountRate');
  const specialTerms = getCustomerMasterHeaderValueK2_(obj, 'specialTerms');
  const s1Referrer = getCustomerMasterHeaderValueK2_(obj, 's1Referrer');
  // P475: 장기미접촉 TM 이관 정보는 마스터 메모와 분리된 읽기 전용 표시 영역에서 사용합니다.
  const longNoContactTransferred = getCustomerMasterHeaderValueK2_(obj, 'longNoContactTransferred');
  const tmProgressStatus = getCustomerMasterHeaderValueK2_(obj, 'tmProgressStatus');
  const tmContactContent = getCustomerMasterHeaderValueK2_(obj, 'tmContactContent');

  // P453: 고객상세의 발주여부는 마스터시트 해당 행의 `발주번호`를 기준으로 봅니다.
  // 수주확정/계약완료 시트에 부분 행이 있더라도 마스터 발주번호가 비어 있으면 X(발주하기)로 표시해야 합니다.
  const orderInfoP250 = buildPortalCustomerOrderInfoFromMasterOrderNoP453_(customerNo, company, orderNo, rowNo);

  const detail = {
    rowNo: rowNo,
    company: company,
    customerNo: customerNo,
    orderNo: orderNo,
    address: address,
    firstRegisteredAt: firstRegisteredAt,
    region: region,
    area: area,
    grade: grade,
    buildingType: buildingType,
    finalQuote: finalQuote,
    contractUnit: contractUnit,
    contractStartDate: contractStartDate,
    contractEndDate: contractEndDate,
    appointment: appointment,
    maintenance: maintenance,
    performance: performance,
    vat: vat,
    discountRate: discountRate,
    specialTerms: specialTerms,
    s1Referrer: s1Referrer,
    longNoContactTransferred: longNoContactTransferred,
    tmProgressStatus: tmProgressStatus,
    tmContactContent: tmContactContent,
    contact: getCustomerMasterHeaderValueK2_(obj, 'contact'),
    phone: getCustomerMasterHeaderValueK2_(obj, 'phone'),
    directPhone: getCustomerMasterHeaderValueK2_(obj, 'directPhone'),
    email: getCustomerMasterHeaderValueK2_(obj, 'email'),
    salesRep: getCustomerMasterHeaderValueK2_(obj, 'salesRep'),
    vendor: getCustomerMasterHeaderValueK2_(obj, 'vendor'),
    status: status,
    customerRank: customerRank,
    statusOptions: buildStatusOptions_(status),
    customerRankOptions: buildCustomerRankOptions_(customerRank),
    memo: getMemoValueFromObj_(obj),
    businessRegistrationReceived: getCustomerMasterHeaderValueK2_(obj, 'businessRegistrationReceived'),
    serviceApplicationReceived: getCustomerMasterHeaderValueK2_(obj, 'serviceApplicationReceived'),
    appointmentReportReceived: getCustomerMasterHeaderValueK2_(obj, 'appointmentReportReceived'),
    businessNo: getCustomerMasterHeaderValueK2_(obj, 'businessNo'),
    businessLegalName: getCustomerMasterHeaderValueK2_(obj, 'businessLegalName'),
    representativeName: getCustomerMasterHeaderValueK2_(obj, 'representativeName'),
    businessAddress: getCustomerMasterHeaderValueK2_(obj, 'businessAddress'),
    lastSent: getCustomerMasterHeaderValueK2_(obj, 'lastSent') || obj['마지막발송'] || '',
    sentAt: getCustomerMasterHeaderValueK2_(obj, 'sentAt') || obj['발송일시'] || '',
    contactRounds: PORTAL_CONFIG.CONTACT_ROUNDS,
    contactMethods: PORTAL_CONFIG.CONTACT_METHODS,
    detailFields: buildDetailFieldValues_(obj),
    quoteCalcDefaults: buildQuoteCalcDefaults_(obj),
    contactLogs: options.includeLogs === false ? [] : getContactHistoryByCustomer_(customerNo, rowNo),
    orderInfo: orderInfoP250 || { exists: false, contractNo: '', rowNo: 0, customerNo: customerNo, company: company },
    raw: options.includeRaw === false ? {} : obj,
    __lite: !!options.lite,
    __source: options.lite ? 'master-lite' : 'master',
    __masterBacked: true,
    __loadedAt: new Date().toISOString()
  };
  addPortalCustomerMetaToDetailP202_(detail, obj, rowNo);
  return detail;
}

function buildDetailFieldValues_(obj) {
  const statusValue = getStatusValueFromObj_(obj);
  const result = { basic: [], contract: [] };
  Object.keys(PORTAL_DETAIL_FIELDS).forEach(section => {
    result[section] = PORTAL_DETAIL_FIELDS[section].map(def => {
      let value = def.key === 'status' ? statusValue : getCustomerMasterHeaderValueK2_(obj, def.headers || []);
      if (def.key === 'contractStartDate' || def.key === 'contractEndDate' || def.key === 'firstRegisteredAt') value = formatPortalContractDateForDisplayP420_(value);
      const field = Object.assign({}, def, { value: value || '' });
      if (def.optionsSource === 'statusOptions') field.options = buildStatusOptions_(statusValue);
      if (def.optionsSource === 'customerRankOptions') field.options = buildCustomerRankOptions_(value);
      return field;
    });
  });
  return result;
}

function buildQuoteCalcDefaults_(obj) {
  const vals = {};
  (PORTAL_DETAIL_FIELDS.contract || []).forEach(def => vals[def.key] = getCustomerMasterHeaderValueK2_(obj, def.headers || []));
  return vals;
}

function getPortalCustomerAllDetailDefsP436_() {
  const result = [];
  Object.keys(PORTAL_DETAIL_FIELDS || {}).forEach(function(section) {
    const list = PORTAL_DETAIL_FIELDS[section] || [];
    list.forEach(function(def) { if (def && def.key) result.push(def); });
  });
  return result;
}


function runPortalCustomerWriteLockedP202_(label, callback) {
  const startedP460 = new Date().getTime();
  // P461: 고객 저장은 실패를 빨리 감지하고 클라이언트 큐/재시도에 맡깁니다.
  // 서버에서 4~5초씩 lock을 기다리면 한 번의 busy가 전체 UX를 30초 이상 끌고 갑니다.
  const lockOptionsP460 = { attempts: 1, waitMs: 250, sleepBaseMs: 80 };
  const finishP460 = function(res, error) {
    const elapsed = new Date().getTime() - startedP460;
    if (res && typeof res === 'object') {
      res.timingP460 = {
        label: String(label || ''),
        totalMs: elapsed,
        lockWaitMs: Number(lockOptionsP460.__lockWaitMsP459 || lockOptionsP460.__lockWaitMsP460 || 0) || 0,
        lockAttempts: Number(lockOptionsP460.__lockAttemptsP459 || lockOptionsP460.__lockAttemptsP460 || 0) || 0
      };
    }
    // P461: 저장 함수 안에서 성능로그_DB에 직접 append하지 않습니다.
    // P460에서는 모든 느린 저장이 서버 로그 시트 append까지 기다려 체감 저장 시간이 더 늘었습니다.
    // 서버는 timingP460만 응답에 실어 보내고, 시트 기록은 클라이언트 배치 flush가 담당합니다.
    return res;
  };
  try {
    let res;
    if (typeof withPortalScriptLockP201_ === 'function') {
      res = withPortalScriptLockP201_(label, callback, lockOptionsP460);
    } else {
      res = callback();
    }
    return finishP460(res, null);
  } catch (err) {
    try {
      const elapsed = new Date().getTime() - startedP460;
      err.timingP460 = {
        label: String(label || ''),
        totalMs: elapsed,
        lockWaitMs: Number(lockOptionsP460.__lockWaitMsP459 || lockOptionsP460.__lockWaitMsP460 || 0) || 0,
        lockAttempts: Number(lockOptionsP460.__lockAttemptsP459 || lockOptionsP460.__lockAttemptsP460 || 0) || 0
      };
      // P461: 오류도 우선 응답으로 돌려보냅니다. 오류 로그는 클라이언트 perf buffer가 시트에 배치 저장합니다.
    } catch (logErr) {}
    throw err;
  }
}


// P462: 초경량 필드 저장 경로 -------------------------------------------------
// 목적: 상태/주소/사업자등록증/특약/계약조건 등 사용자가 이미 화면에서 계산한 patch를
// 마스터시트에 바로 반영합니다. 기존 saveCustomerDetailFastCoreP202_는 계약조건 보정,
// 전체 행 display read, 검증 flush, 후처리 분기 때문에 단일 필드도 3~10초가 걸렸습니다.
function isPortalCustomerThinSavePayloadP462_(payload, values) {
  payload = payload || {};
  values = values || payload.values || {};
  const keys = Object.keys(values || {});
  if (!keys.length) return false;
  if (payload.thinSave === true) return true;
  // 저장 출처가 명확하고 한두 칸 patch인 경우는 초경량 저장으로 처리합니다.
  if (payload.fastMode === true && payload.noSynchronousRefresh === true && keys.length <= 4) return true;
  return false;
}

function getPortalThinSaveDetailDefMapP462_() {
  const defs = getPortalCustomerAllDetailDefsP436_();
  const map = {};
  defs.forEach(function(def) { if (def && def.key) map[def.key] = def; });
  return map;
}

function saveCustomerDetailThinCoreP462_(payload) {
  payload = payload || {};
  const started = new Date().getTime();
  const target = assertCustomerTarget_(payload, '고객 상세정보 초경량 저장', { readObject: false });
  assertPortalCanAccessCustomerTarget_(target, '고객 상세정보 초경량 저장');
  const rowNo = target.rowNo;
  const customerNo = target.customerNo;
  const sheet = target.sheet;
  let values = normalizePortalCustomerLocationValues_(Object.assign({}, payload.values || {}));
  values = preparePortalContractValuesForSaveP112_(values || {}, { sheet: sheet, rowNo: rowNo, requireFull: false });
  values = applyPortalAutoGradeByAreaForSaveP451_(values);

  const keys = Object.keys(values || {});
  if (!keys.length) {
    return { ok: true, rowNo: rowNo, customerNo: customerNo, changedFields: [], changedKeys: [], values: {}, changedValues: {}, savedAt: new Date().toISOString(), message: '저장할 변경값이 없습니다.', thinSaveP462: true };
  }

  const defMap = getPortalThinSaveDetailDefMapP462_();
  let headerMap = getHeaderMap_(sheet);
  const expected = payload.expectedValues || {};
  const targets = [];
  const changed = [];
  const changedKeys = [];
  const changedValues = {};
  const appliedValues = {};

  keys.forEach(function(key) {
    const def = defMap[key];
    if (!canWritePortalDetailFieldP451_(def, values)) return;
    const col = findFirstExistingHeaderCol_(headerMap, def.headers || []) || ensureMasterColumn_(sheet, headerMap, (def.headers && def.headers[0]) || def.label);
    headerMap = getHeaderMap_(sheet);
    const nextValue = getPortalMasterCompareTextP280_(key, values[key]);
    const writeValue = getPortalMasterWriteValueP280_(key, values[key]);
    const currentText = String(sheet.getRange(rowNo, col).getDisplayValue() || '').trim();
    if (Object.prototype.hasOwnProperty.call(expected, key)) {
      const expectedText = getPortalMasterCompareTextP280_(key, expected[key]);
      if (currentText !== expectedText && currentText !== nextValue) {
        throw makePortalFieldConflictErrorP474_({
          rowNo: rowNo,
          customerNo: customerNo,
          key: key,
          label: def.label || key,
          currentValue: currentText,
          expectedValue: expectedText,
          attemptedValue: nextValue,
          payload: payload,
          source: 'customer.detailThin'
        });
      }
    }
    appliedValues[key] = nextValue;
    if (currentText !== nextValue) {
      targets.push({ key: key, label: def.label || key, col: col, writeValue: writeValue, value: nextValue, oldValue: currentText });
      changed.push(def.label || key);
      changedKeys.push(key);
      changedValues[key] = nextValue;
    }
  });

  if (targets.length) {
    targets.sort(function(a, b) { return a.col - b.col; });
    // 필드 format 적용은 필요한 경우만 수행합니다. 모든 저장마다 과한 서식 보정은 속도를 잡아먹습니다.
    targets.forEach(function(t) {
      if (['finalQuote','contractStartDate','contractEndDate','area','discountRate'].indexOf(String(t.key || '')) >= 0) {
        try { applyPortalMasterCellFormatP433_(sheet, rowNo, t.col, t.key); } catch (fmtErr) {}
      }
    });
    let blockStart = targets[0].col;
    let prevCol = targets[0].col;
    let blockValues = [targets[0].writeValue];
    const flushBlock = function() { sheet.getRange(rowNo, blockStart, 1, blockValues.length).setValues([blockValues]); };
    for (let i = 1; i < targets.length; i++) {
      const t = targets[i];
      if (t.col === prevCol + 1) {
        blockValues.push(t.writeValue);
        prevCol = t.col;
      } else {
        flushBlock();
        blockStart = t.col;
        prevCol = t.col;
        blockValues = [t.writeValue];
      }
    }
    flushBlock();
  }

  let masterMeta = null;
  let indexUpdate = null;
  let fieldChangeLogP474 = null;
  if (targets.length) {
    if (payload.skipMasterMeta !== true) {
      masterMeta = bumpPortalCustomerMasterMetaP202_(sheet, rowNo, 'webapp-customer-thin-save');
    }
    fieldChangeLogP474 = recordPortalFieldChangesP474_(targets.map(function(t) {
      return { key: t.key, label: t.label, oldValue: t.oldValue, newValue: t.value };
    }), Object.assign({}, payload, { rowNo: rowNo, customerNo: customerNo, source: normalizePortalSaveSourceP474_(payload, 'customer.detailThin') }));
    if (payload.skipIndexQueue !== true) indexUpdate = queueCustomerSearchIndexRefreshAfterSaveP400_(rowNo, customerNo, changedKeys, masterMeta, 'WEBAPP_CUSTOMER_THIN_SAVE');
  }

  return {
    ok: true,
    rowNo: rowNo,
    customerNo: customerNo,
    changedFields: changed,
    changedKeys: changedKeys,
    values: appliedValues,
    changedValues: changedValues,
    indexUpdate: indexUpdate,
    fieldChangeLogP474: fieldChangeLogP474,
    masterVersion: masterMeta && masterMeta.version || '',
    masterUpdatedAt: masterMeta && masterMeta.updatedAt || '',
    masterEditor: masterMeta && masterMeta.editor || '',
    savedAt: new Date().toISOString(),
    verified: false,
    fastPatch: true,
    thinSaveP462: true,
    clientOperationId: payload.clientOperationId || '',
    noSynchronousRefresh: true,
    timingInnerP462: { totalMs: new Date().getTime() - started, keyCount: keys.length, changedCount: changedKeys.length },
    message: changed.length ? ('저장 완료: ' + changed.join(', ')) : '변경된 값이 없습니다.'
  };
}

function saveCustomerMemoThinCoreP462_(rowNoOrPayload, memo) {
  const started = new Date().getTime();
  const payload = (rowNoOrPayload && typeof rowNoOrPayload === 'object') ? rowNoOrPayload : { rowNo: rowNoOrPayload, memo: memo };
  const target = assertCustomerTarget_(payload, '고객 메모 초경량 저장', { readObject: false });
  assertPortalCanAccessCustomerTarget_(target, '고객 메모 초경량 저장');
  const rowNo = target.rowNo;
  const customerNo = target.customerNo;
  const sheet = target.sheet;
  const headerMap = getHeaderMap_(sheet);
  const memoCol = findMasterFieldCol_(headerMap, 'memo') || ensureMasterFieldColumn_(sheet, headerMap, 'memo');
  const nextMemo = String(payload.memo == null ? '' : payload.memo);
  const currentMemo = String(sheet.getRange(rowNo, memoCol).getDisplayValue() || '');
  if (payload.expectedValues && Object.prototype.hasOwnProperty.call(payload.expectedValues, 'memo')) {
    const expectedMemo = String(payload.expectedValues.memo == null ? '' : payload.expectedValues.memo);
    if (currentMemo !== expectedMemo && currentMemo !== nextMemo) {
      throw makePortalFieldConflictErrorP474_({
        rowNo: rowNo,
        customerNo: customerNo,
        key: 'memo',
        label: '마스터시트 메모',
        currentValue: currentMemo,
        expectedValue: expectedMemo,
        attemptedValue: nextMemo,
        payload: payload,
        source: 'customer.memoThin'
      });
    }
  }
  if (currentMemo !== nextMemo) sheet.getRange(rowNo, memoCol).setValue(nextMemo);
  let masterMeta = null;
  let indexUpdate = null;
  let fieldChangeLogP474 = null;
  if (currentMemo !== nextMemo) {
    if (payload.skipMasterMeta !== true) masterMeta = bumpPortalCustomerMasterMetaP202_(sheet, rowNo, 'webapp-customer-memo-thin-save');
    fieldChangeLogP474 = recordPortalFieldChangesP474_([{ key: 'memo', label: '마스터시트 메모', oldValue: currentMemo, newValue: nextMemo }], Object.assign({}, payload, { rowNo: rowNo, customerNo: customerNo, source: normalizePortalSaveSourceP474_(payload, 'customer.memoThin') }));
    if (payload.skipIndexQueue !== true) indexUpdate = queueCustomerSearchIndexRefreshAfterSaveP400_(rowNo, customerNo, ['memo'], masterMeta, 'WEBAPP_CUSTOMER_MEMO_THIN_SAVE');
  }
  return {
    ok: true,
    rowNo: rowNo,
    customerNo: customerNo,
    memo: nextMemo,
    changedFields: currentMemo !== nextMemo ? ['메모'] : [],
    changedKeys: currentMemo !== nextMemo ? ['memo'] : [],
    values: { memo: nextMemo },
    changedValues: currentMemo !== nextMemo ? { memo: nextMemo } : {},
    indexUpdate: indexUpdate,
    fieldChangeLogP474: fieldChangeLogP474,
    masterVersion: masterMeta && masterMeta.version || '',
    masterUpdatedAt: masterMeta && masterMeta.updatedAt || '',
    masterEditor: masterMeta && masterMeta.editor || '',
    savedAt: new Date().toISOString(),
    thinSaveP462: true,
    timingInnerP462: { totalMs: new Date().getTime() - started, changed: currentMemo !== nextMemo },
    message: '메모 저장 완료'
  };
}


// P474: field-level direct save conflict + field change log ------------------
// 원칙: 같은 고객이라도 다른 필드는 서로 막지 않습니다. expectedValues가 제공된 필드만 현재값과 비교합니다.
const PORTAL_FIELD_CHANGE_LOG_P474 = {
  SHEET_NAME: '필드변경로그_DB',
  HEADERS: ['변경일시','작업ID','사용자','세션ID','고객번호','rowNo','source','필드명','헤더명','이전값요약','변경값요약','상태']
};

function getPortalActiveUserEmailP474_() {
  try { return String(Session.getActiveUser().getEmail() || '').trim(); } catch (e) { return ''; }
}

function getPortalFieldChangeLogSheetP474_() {
  const ss = getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(PORTAL_FIELD_CHANGE_LOG_P474.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(PORTAL_FIELD_CHANGE_LOG_P474.SHEET_NAME);
  const headers = PORTAL_FIELD_CHANGE_LOG_P474.HEADERS;
  const width = Math.max(sheet.getLastColumn(), headers.length);
  let current = [];
  if (sheet.getLastRow() >= 1) current = sheet.getRange(1,1,1,width).getDisplayValues()[0].map(function(v){ return String(v||'').trim(); });
  const seen = {}; current.forEach(function(h){ if(h) seen[h]=true; });
  let changed = current.filter(Boolean).length === 0;
  headers.forEach(function(h){ if(!seen[h]){ current.push(h); seen[h]=true; changed=true; } });
  if (changed) { sheet.getRange(1,1,1,current.length).setValues([current]); try { sheet.setFrozenRows(1); } catch(e){} }
  return sheet;
}

function truncatePortalFieldLogValueP474_(value) {
  const text = String(value == null ? '' : value);
  return text.length > 900 ? text.slice(0,900) + '…' : text;
}

function normalizePortalSaveSourceP474_(payload, fallback) {
  payload = payload || {};
  const src = String(payload.clientSaveSource || payload.source || payload.saveSource || fallback || '').trim();
  if (src) return src;
  return 'unknown.customerSave';
}

function recordPortalFieldChangesP474_(items, ctx) {
  items = Array.isArray(items) ? items : [];
  if (!items.length) return { ok: true, count: 0 };
  ctx = ctx || {};
  try {
    const sheet = getPortalFieldChangeLogSheetP474_();
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(function(v){ return String(v||'').trim(); });
    const now = new Date();
    const user = getPortalActiveUserEmailP474_();
    const rows = items.map(function(item){
      const obj = {
        '변경일시': now,
        '작업ID': String(ctx.operationId || ctx.clientOperationId || ''),
        '사용자': user,
        '세션ID': String(ctx.sessionId || ctx.clientSessionId || ''),
        '고객번호': String(ctx.customerNo || ''),
        'rowNo': Number(ctx.rowNo || 0) || '',
        'source': normalizePortalSaveSourceP474_(ctx, ''),
        '필드명': String(item.key || ''),
        '헤더명': String(item.label || item.header || item.key || ''),
        '이전값요약': truncatePortalFieldLogValueP474_(item.oldValue),
        '변경값요약': truncatePortalFieldLogValueP474_(item.newValue),
        '상태': 'DONE'
      };
      return headers.map(function(h){ return Object.prototype.hasOwnProperty.call(obj,h) ? obj[h] : ''; });
    });
    sheet.getRange(sheet.getLastRow()+1,1,rows.length,headers.length).setValues(rows);
    return { ok: true, count: rows.length };
  } catch (err) {
    Logger.log('P474 필드변경로그 기록 실패: ' + (err && err.stack || err));
    return { ok: false, count: 0, error: String(err && err.message || err) };
  }
}

function findLatestPortalFieldChangeP474_(customerNo, rowNo, fieldKey) {
  try {
    const sheet = getPortalFieldChangeLogSheetP474_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(function(v){ return String(v||'').trim(); });
    const idx = {}; headers.forEach(function(h,i){ if(h) idx[h]=i; });
    const rowCount = Math.min(500, lastRow-1);
    const values = sheet.getRange(lastRow-rowCount+1,1,rowCount,headers.length).getValues();
    const cno = String(customerNo || '').trim();
    const rno = Number(rowNo || 0) || 0;
    const fkey = String(fieldKey || '').trim();
    for (let i=values.length-1; i>=0; i--) {
      const row = values[i];
      if (idx['고객번호'] != null && cno && String(row[idx['고객번호']] || '').trim() !== cno) continue;
      if (idx['rowNo'] != null && rno && Number(row[idx['rowNo']] || 0) !== rno) continue;
      if (idx['필드명'] != null && fkey && String(row[idx['필드명']] || '').trim() !== fkey) continue;
      return {
        changedAt: row[idx['변경일시']] instanceof Date ? row[idx['변경일시']].toISOString() : String(row[idx['변경일시']] || ''),
        user: idx['사용자'] != null ? String(row[idx['사용자']] || '') : '',
        source: idx['source'] != null ? String(row[idx['source']] || '') : '',
        field: idx['필드명'] != null ? String(row[idx['필드명']] || '') : fkey,
        header: idx['헤더명'] != null ? String(row[idx['헤더명']] || '') : '',
        oldValueSummary: idx['이전값요약'] != null ? String(row[idx['이전값요약']] || '') : '',
        newValueSummary: idx['변경값요약'] != null ? String(row[idx['변경값요약']] || '') : ''
      };
    }
  } catch (err) {
    Logger.log('P474 최신 필드 변경 조회 실패: ' + (err && err.stack || err));
  }
  return null;
}

function makePortalFieldConflictErrorP474_(ctx) {
  ctx = ctx || {};
  const latest = findLatestPortalFieldChangeP474_(ctx.customerNo, ctx.rowNo, ctx.key) || {};
  const info = {
    type: 'FIELD_CONFLICT',
    rowNo: Number(ctx.rowNo || 0) || 0,
    customerNo: String(ctx.customerNo || ''),
    field: String(ctx.key || ''),
    label: String(ctx.label || ctx.key || ''),
    source: normalizePortalSaveSourceP474_(ctx.payload || {}, ctx.source || ''),
    expectedValue: String(ctx.expectedValue == null ? '' : ctx.expectedValue),
    currentValue: String(ctx.currentValue == null ? '' : ctx.currentValue),
    attemptedValue: String(ctx.attemptedValue == null ? '' : ctx.attemptedValue),
    latestChange: latest
  };
  const err = new Error('다른 사용자가 같은 항목을 먼저 수정했습니다. 최신 내용을 확인한 뒤 다시 저장해 주세요.');
  err.code = 'PORTAL_FIELD_CONFLICT_P474';
  err.conflictInfoP474 = info;
  return err;
}

function isPortalFieldConflictErrorP474_(err) {
  return String(err && err.code || '') === 'PORTAL_FIELD_CONFLICT_P474' || !!(err && err.conflictInfoP474);
}

function getPortalDirectSaveConflictInfoP474_(err) {
  return (err && err.conflictInfoP474) || null;
}

function getPortalCurrentFieldValueP474_(sheet, rowNo, key, col) {
  if (!sheet || !rowNo || !col) return '';
  if (String(key || '') === 'memo') return String(sheet.getRange(rowNo, col).getValue() || '');
  return String(sheet.getRange(rowNo, col).getDisplayValue() || '').trim();
}

function getCustomerFieldChangesSinceP474(payload) {
  payload = payload || {};
  const sinceMs = Number(payload.sinceMs || 0) || 0;
  const visible = Array.isArray(payload.visibleCustomers) ? payload.visibleCustomers : [];
  const wanted = {};
  visible.forEach(function(item){
    const rowNo = Number(item && item.rowNo || 0) || 0;
    const customerNo = String(item && item.customerNo || '').trim();
    if (rowNo || customerNo) wanted[rowNo + ':' + customerNo] = true;
  });
  const sheet = getPortalFieldChangeLogSheetP474_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, changes: [], nowMs: Date.now() };
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(function(v){ return String(v||'').trim(); });
  const idx = {}; headers.forEach(function(h,i){ if(h) idx[h]=i; });
  const rowCount = Math.min(600, lastRow-1);
  const values = sheet.getRange(lastRow-rowCount+1,1,rowCount,headers.length).getValues();
  const changes = [];
  values.forEach(function(row){
    const d = row[idx['변경일시']];
    const ms = d instanceof Date ? d.getTime() : Date.parse(String(d || ''));
    if (sinceMs && (!ms || ms <= sinceMs)) return;
    const rowNo = Number(row[idx['rowNo']] || 0) || 0;
    const customerNo = String(row[idx['고객번호']] || '').trim();
    if (Object.keys(wanted).length && !wanted[rowNo + ':' + customerNo]) return;
    changes.push({
      rowNo: rowNo,
      customerNo: customerNo,
      field: String(row[idx['필드명']] || ''),
      label: String(row[idx['헤더명']] || ''),
      valueSummary: String(row[idx['변경값요약']] || ''),
      changedBy: String(row[idx['사용자']] || ''),
      changedAt: d instanceof Date ? d.toISOString() : String(d || ''),
      changedAtMs: ms || 0,
      source: String(row[idx['source']] || '')
    });
  });
  return { ok: true, changes: changes, nowMs: Date.now() };
}


function saveCustomerDetail(payload) {
  return runPortalCustomerWriteLockedP202_('customer-detail-save', function() {
    return saveCustomerDetailCoreP202_(payload);
  });
}

function saveCustomerDetailFast(payload) {
  payload = payload || {};
  const cachedP458 = getPortalCachedOperationResultP458_(payload);
  if (cachedP458) return cachedP458;

  const valuesP462 = payload && payload.values || {};

  // P473: 정상적인 단일/경량 저장은 전역 ScriptLock 없이 마스터시트에 직접 저장합니다.
  // 기존 구조는 3~7초짜리 저장이 전역 lock을 잡고, 다른 사용자는 250ms 후 "다른 작업 처리 중"으로 튕겼습니다.
  // 경량 저장은 rowNo/customerNo + expectedValues로 필드 단위 충돌을 검사하므로 전역 lock이 필요 없습니다.
  if (isPortalCustomerThinSavePayloadP462_(payload, valuesP462)) {
    try {
      const res = saveCustomerDetailThinCoreP462_(payload);
      res.directSaveP473 = true;
      putPortalCachedOperationResultP458_(payload, res);
      return res;
    } catch (err) {
      // 충돌/일시 오류는 사용자에게 busy 실패로 던지지 않고 저장큐_DB에 남깁니다.
      if (typeof enqueueSaveFallbackP473_ === 'function' && (isPortalServerTransientWriteErrorP473_(err) || isPortalServerStaleErrorP473_(err))) {
        return enqueueSaveFallbackP473_('saveCustomerDetailFast', payload, 'DIRECT_SAVE_FAILED', err);
      }
      throw err;
    }
  }

  // 무거운 계약조건/복합 저장은 기존 lock 경로를 유지하되, busy는 큐 fallback으로 전환합니다.
  try {
    return runPortalCustomerWriteLockedP202_('customer-detail-fast-save', function() {
      const cachedInsideLockP458 = getPortalCachedOperationResultP458_(payload);
      if (cachedInsideLockP458) return cachedInsideLockP458;
      const res = saveCustomerDetailFastCoreP202_(payload);
      putPortalCachedOperationResultP458_(payload, res);
      return res;
    });
  } catch (err) {
    if (typeof enqueueSaveFallbackP473_ === 'function' && (isPortalServerTransientWriteErrorP473_(err) || isPortalServerStaleErrorP473_(err))) {
      return enqueueSaveFallbackP473_('saveCustomerDetailFast', payload, 'LOCKED_SAVE_FAILED', err);
    }
    throw err;
  }
}

function saveCustomerMemoFast(rowNoOrPayload, memo) {
  const payloadP462 = (rowNoOrPayload && typeof rowNoOrPayload === 'object') ? rowNoOrPayload : { rowNo: rowNoOrPayload, memo: memo };
  payloadP462.thinSave = true;
  const cachedP458 = getPortalCachedOperationResultP458_(payloadP462);
  if (cachedP458) return cachedP458;

  // P473: 메모는 가장 민감한 입력 필드이므로 전역 lock 없이 expectedMemo/expectedValues 기준으로 직접 저장합니다.
  try {
    const res = saveCustomerMemoThinCoreP462_(payloadP462);
    res.directSaveP473 = true;
    putPortalCachedOperationResultP458_(payloadP462, res);
    return res;
  } catch (err) {
    if (typeof enqueueSaveFallbackP473_ === 'function' && (isPortalServerTransientWriteErrorP473_(err) || isPortalServerStaleErrorP473_(err))) {
      return enqueueSaveFallbackP473_('saveCustomerMemoFast', payloadP462, 'MEMO_DIRECT_SAVE_FAILED', err);
    }
    throw err;
  }
}


// P451: 고객상세에서 연면적이 수정되면 관리등급도 같은 저장 트랜잭션에서 자동 재계산합니다.
// 관리등급 필드는 화면에서 직접 수정 불가이지만, 연면적 변경에 따른 파생값은 마스터시트에 같이 반영해야 합니다.
function applyPortalAutoGradeByAreaForSaveP451_(values) {
  values = Object.assign({}, values || {});
  if (!Object.prototype.hasOwnProperty.call(values, 'area')) return values;

  const area = parsePortalDecimalNumberP433_(values.area);
  if (area === '' || !(Number(area) > 0)) {
    values.grade = '';
    return values;
  }

  values.grade = calculateGradeByArea_(area);
  return values;
}

function canWritePortalDetailFieldP451_(def, values) {
  if (!def || !def.key) return false;
  if (def.editable !== false) return true;

  // 관리등급은 사용자가 직접 수정하는 필드가 아니라 연면적 변경의 파생값으로만 저장 허용.
  if (def.key === 'grade' && Object.prototype.hasOwnProperty.call(values || {}, 'area')) return true;
  return false;
}

function saveCustomerDetailCoreP202_(payload) {
  payload = payload || {};
  const target = assertCustomerTarget_(payload, '고객 상세정보 저장', { readObject: false });
  assertPortalCanAccessCustomerTarget_(target, '고객 상세정보 저장');
  const rowNo = target.rowNo;
  const customerNo = target.customerNo;
  let values = payload.values || {};
  values = normalizePortalCustomerLocationValues_(values);
  const sheet = target.sheet;
  assertPortalCustomerVersionFreshP202_(sheet, rowNo, payload);
  values = preparePortalContractValuesForSaveP112_(values, { sheet: sheet, rowNo: rowNo, requireFull: false });
  values = applyPortalAutoGradeByAreaForSaveP451_(values);

  const allDefs = getPortalCustomerAllDetailDefsP436_();
  let headerMap = getHeaderMap_(sheet);
  const changed = [];

  allDefs.forEach(def => {
    if (!canWritePortalDetailFieldP451_(def, values)) return;
    if (!Object.prototype.hasOwnProperty.call(values, def.key)) return;
    const col = findFirstExistingHeaderCol_(headerMap, def.headers || []) || ensureMasterColumn_(sheet, headerMap, (def.headers && def.headers[0]) || def.label);
    headerMap = getHeaderMap_(sheet);
    const nextValue = getPortalMasterCompareTextP280_(def.key, values[def.key]);
    const writeValue = getPortalMasterWriteValueP280_(def.key, values[def.key]);
    const range = sheet.getRange(rowNo, col);
    const prevValue = String(range.getDisplayValue() || '').trim();
    if (prevValue !== nextValue) {
      applyPortalMasterCellFormatP433_(sheet, rowNo, col, def.key);
      range.setValue(writeValue);
      changed.push(def.label);
    }
  });

  let indexUpdate = null;
  let masterMetaP202 = null;
  if (changed.length) {
    masterMetaP202 = bumpPortalCustomerMasterMetaP202_(sheet, rowNo, 'webapp-customer-detail-save');
    indexUpdate = queueCustomerSearchIndexRefreshAfterSaveP400_(rowNo, customerNo, changed, masterMetaP202, 'WEBAPP_CUSTOMER_DETAIL_SAVE');
    try { if (typeof syncContractCompleteFromCustomerMasterP420_ === 'function') syncContractCompleteFromCustomerMasterP420_({ sheet: sheet, rowNo: rowNo, customerNo: customerNo, changedKeys: changed, source: 'customerDetailSave' }); } catch (syncErrP420) { Logger.log('수주확정/계약완료 동기화 실패: ' + (syncErrP420 && syncErrP420.stack || syncErrP420)); }
    try { CacheService.getScriptCache().remove('PORTAL_DASHBOARD_V46_FAST_HOME'); } catch (err) {}
  }

  if (changed.length) {
    try {
      appendPortalActivityLog_({
        actionType: '고객정보수정',
        screen: '고객 상세',
        rowNo: rowNo,
        customerNo: customerNo,
        summary: '상세정보 수정: ' + changed.join(', '),
        detail: { changedFields: changed, values: values }
      });
    } catch (err) {}
  }

  return {
    ok: true,
    rowNo: rowNo,
    customerNo: customerNo,
    changedFields: changed,
    indexUpdate: indexUpdate,
    masterVersion: masterMetaP202 && masterMetaP202.version || '',
    masterUpdatedAt: masterMetaP202 && masterMetaP202.updatedAt || '',
    masterEditor: masterMetaP202 && masterMetaP202.editor || '',
    message: changed.length ? ('상세정보 저장 완료: ' + changed.join(', ')) : '변경된 값이 없습니다.',
    detail: getCustomerDetail(rowNo)
  };
}


function verifyPortalCustomerFastSaveAppliedP430_(sheet, rowNo, changedTargets) {
  if (!sheet || !rowNo || !changedTargets || !changedTargets.length) return true;
  const failed = [];
  changedTargets.forEach(function(t) {
    try {
      const actual = String(sheet.getRange(rowNo, t.col).getDisplayValue() || '').trim();
      const expected = String(t.value == null ? '' : t.value).trim();
      if (actual !== expected) {
        failed.push(t.label + ' expected=[' + expected + '] actual=[' + actual + ']');
      }
    } catch (err) {
      failed.push(t.label + ' 확인 실패: ' + (err && err.message || err));
    }
  });
  if (failed.length) {
    throw new Error('마스터시트 저장 확인 실패: ' + failed.join(' / '));
  }
  return true;
}

function saveCustomerDetailFastCoreP202_(payload) {
  // v45: 저장 후 상세 전체 재조회 금지. 변경된 필드만 쓰고 결과만 반환합니다.
  // v56 PATCH A: customerNo를 기준키로 우선 검증하고, rowNo는 보조 위치값으로만 사용합니다.
  // PATCH D: 셀별 get/set 반복을 줄이고, 검색인덱스는 마스터 재조회 없이 patch 갱신합니다.
  // PATCH P1-12: 계약조건이 바뀐 경우에는 시트 수식 반영을 위해 단건 마스터 재조회로 인덱스/상세를 확정합니다.
  payload = payload || {};
  const target = assertCustomerTarget_(payload, '고객 상세정보 저장', { readObject: false });
  assertPortalCanAccessCustomerTarget_(target, '고객 상세정보 저장');
  const rowNo = target.rowNo;
  const customerNo = target.customerNo;
  let values = payload.values || {};
  values = normalizePortalCustomerLocationValues_(values);
  const sheet = target.sheet;
  assertPortalCustomerVersionFreshP202_(sheet, rowNo, payload);
  values = preparePortalContractValuesForSaveP112_(values, { sheet: sheet, rowNo: rowNo, requireFull: false });
  values = applyPortalAutoGradeByAreaForSaveP451_(values);

  const allDefs = getPortalCustomerAllDetailDefsP436_();
  let headerMap = getHeaderMap_(sheet);
  const targets = [];
  const appliedValues = {};

  allDefs.forEach(function(def) {
    if (!canWritePortalDetailFieldP451_(def, values)) return;
    if (!Object.prototype.hasOwnProperty.call(values, def.key)) return;
    const col = findFirstExistingHeaderCol_(headerMap, def.headers || []) || ensureMasterColumn_(sheet, headerMap, (def.headers && def.headers[0]) || def.label);
    headerMap = getHeaderMap_(sheet);
    const nextValue = getPortalMasterCompareTextP280_(def.key, values[def.key]);
    const writeValue = getPortalMasterWriteValueP280_(def.key, values[def.key]);
    targets.push({ key: def.key, label: def.label, col: col, value: nextValue, writeValue: writeValue });
    appliedValues[def.key] = nextValue;
  });

  if (!targets.length) {
    return {
      ok: true,
      rowNo: rowNo,
      customerNo: customerNo,
      changedFields: [],
      changedKeys: [],
      values: {},
      appliedValues: {},
      savedAt: new Date().toISOString(),
      message: '저장할 수 있는 변경값이 없습니다.'
    };
  }

  const lastCol = Math.max(sheet.getLastColumn(), Math.max.apply(null, targets.map(function(t) { return t.col; })));
  const currentDisplayRow = sheet.getRange(rowNo, 1, 1, lastCol).getDisplayValues()[0];
  const changedTargets = [];
  targets.forEach(function(t) {
    const prevValue = String(currentDisplayRow[t.col - 1] || '').trim();
    if (prevValue !== t.value) changedTargets.push(t);
  });

  const changed = changedTargets.map(function(t) { return t.label; });
  const changedKeys = changedTargets.map(function(t) { return t.key; });
  const changedValues = {};
  changedTargets.forEach(function(t) { changedValues[t.key] = t.value; });

  if (changedTargets.length) {
    changedTargets.forEach(function(t) { applyPortalMasterCellFormatP433_(sheet, rowNo, t.col, t.key); });
    changedTargets.sort(function(a, b) { return a.col - b.col; });
    let blockStart = changedTargets[0].col;
    let blockValues = [changedTargets[0].writeValue];
    let prevCol = changedTargets[0].col;

    const flushBlock = function() {
      sheet.getRange(rowNo, blockStart, 1, blockValues.length).setValues([blockValues]);
    };

    for (let i = 1; i < changedTargets.length; i++) {
      const t = changedTargets[i];
      if (t.col === prevCol + 1) {
        blockValues.push(t.writeValue);
        prevCol = t.col;
      } else {
        flushBlock();
        blockStart = t.col;
        blockValues = [t.writeValue];
        prevCol = t.col;
      }
    }
    flushBlock();

    // P455: 고객상세/계약조건 저장은 빠른 화면 반영보다 실제 마스터 반영 확정이 우선입니다.
    // 특히 연면적·할인율·최종 견적가·계약조건 계열은 저장 직후 수식/표시값 재조회에 쓰이므로
    // 여기서 즉시 커밋하고 저장값을 1차 검증합니다.
    const verifyDirectWriteP455 = changedTargets.some(function(t) {
      return ['area','grade','buildingType','finalQuote','contractUnit','contractStartDate','contractEndDate','appointment','maintenance','performance','vat','discountRate','specialTerms'].indexOf(String(t.key || '')) >= 0;
    });
    if (verifyDirectWriteP455) {
      SpreadsheetApp.flush();
      verifyPortalCustomerFastSaveAppliedP430_(sheet, rowNo, changedTargets);
    }
  }

  let indexUpdate = null;
  let masterMetaP202 = null;
  // P457: 상세 모달 저장은 화면 보호/체감속도를 위해 동기 상세 재조회·인덱스 직접갱신을 생략할 수 있습니다.
  // 마스터 셀 저장과 수정버전 bump는 즉시 수행하고, 검색인덱스/상세 보강은 dirty queue/다음 조회에 맡깁니다.
  const noSynchronousRefreshP457 = payload.noSynchronousRefresh === true;
  const shouldRefreshMaster = !noSynchronousRefreshP457 && shouldRefreshIndexFromMasterAfterContractSaveP112_(changedKeys);
  const deferPostProcessP400 = noSynchronousRefreshP457 || shouldUseDeferredCustomerSavePostProcessP400_(changedKeys);
  let refreshedDetail = null;
  const ultraFastP433 = payload.fastMode === true;
  if (changedTargets.length) {
    masterMetaP202 = bumpPortalCustomerMasterMetaP202_(sheet, rowNo, 'webapp-customer-fast-save');
    if (deferPostProcessP400) {
      indexUpdate = queueCustomerSearchIndexRefreshAfterSaveP400_(rowNo, customerNo, changedKeys, masterMetaP202, 'WEBAPP_CUSTOMER_FAST_SAVE');
    } else {
      // 계약조건처럼 파생 수식 확인이 필요한 저장만 기존 동기 재조회 경로를 유지합니다.
      SpreadsheetApp.flush();
      try { indexUpdate = updateCustomerSearchIndexRow_(rowNo); } catch (err) { Logger.log('검색인덱스 갱신 실패: ' + (err && err.stack || err)); }
      try { refreshedDetail = getCustomerDetail(rowNo); } catch (err) { Logger.log('계약조건 저장 후 상세 재조회 실패: ' + (err && err.stack || err)); }
    }
    if (!ultraFastP433) {
      try { if (typeof syncContractCompleteFromCustomerMasterP420_ === 'function') syncContractCompleteFromCustomerMasterP420_({ sheet: sheet, rowNo: rowNo, customerNo: customerNo, changedKeys: changedKeys, source: 'customerFastSave' }); } catch (syncErrP420) { Logger.log('수주확정/계약완료 동기화 실패: ' + (syncErrP420 && syncErrP420.stack || syncErrP420)); }
    }
    try { CacheService.getScriptCache().remove('PORTAL_DASHBOARD_V46_FAST_HOME'); } catch (err) {}
  }

  let auditLoggedP430 = false;
  if (changedTargets.length && payload.requireAuditLog !== false && !ultraFastP433) {
    try {
      auditLoggedP430 = appendPortalActivityLog_({
        actionType: '고객정보수정',
        screen: payload.clientSaveSource === 'inlineEdit' ? '고객 목록 즉시수정' : '고객 상세',
        rowNo: rowNo,
        customerNo: customerNo,
        summary: '상세정보 빠른 저장: ' + changed.join(', '),
        detail: {
          changedFields: changed,
          changedKeys: changedKeys,
          values: changedValues,
          clientSaveSource: payload.clientSaveSource || '',
          clientRequestId: payload.clientRequestId || '',
          clientSavedAt: payload.clientSavedAt || ''
        }
      }) === true;
    } catch (auditErrP430) {
      Logger.log('빠른 상세저장 작업로그 기록 실패: ' + (auditErrP430 && auditErrP430.stack || auditErrP430));
    }
  }

  return {
    ok: true,
    rowNo: rowNo,
    customerNo: customerNo,
    changedFields: changed,
    changedKeys: changedKeys,
    values: appliedValues,
    changedValues: changedValues,
    indexUpdate: indexUpdate,
    masterVersion: masterMetaP202 && masterMetaP202.version || '',
    masterUpdatedAt: masterMetaP202 && masterMetaP202.updatedAt || '',
    masterEditor: masterMetaP202 && masterMetaP202.editor || '',
    savedAt: new Date().toISOString(),
    verified: true,
    auditLogged: auditLoggedP430,
    fastPatch: !shouldRefreshMaster,
    clientOperationId: payload.clientOperationId || '',
    noSynchronousRefresh: noSynchronousRefreshP457,
    detail: refreshedDetail,
    refreshedFromMaster: !!refreshedDetail,
    message: changed.length ? ('상세정보 저장 완료: ' + changed.join(', ')) : '변경된 값이 없습니다.'
  };
}

function saveCustomerMemoFastCoreP202_(rowNoOrPayload, memo) {
  // v45: 메모 한 칸만 저장. 상세 재조회/전체 인덱스 재생성 금지.
  // v56 PATCH A: 기존 saveCustomerMemoFast(rowNo, memo)도 호환하되, 신규 호출은 {customerNo,rowNo,memo}를 권장합니다.
  const payload = (rowNoOrPayload && typeof rowNoOrPayload === 'object')
    ? rowNoOrPayload
    : { rowNo: rowNoOrPayload, memo: memo };
  const target = assertCustomerTarget_(payload, '고객 메모 저장', { readObject: false });
  assertPortalCanAccessCustomerTarget_(target, '고객 메모 저장');
  const targetRowNo = target.rowNo;
  const customerNo = target.customerNo;
  const sheet = target.sheet;
  assertPortalCustomerVersionFreshP202_(sheet, targetRowNo, payload);

  const headerMap = getHeaderMap_(sheet);
  const memoCol = findMasterFieldCol_(headerMap, 'memo') || ensureMasterFieldColumn_(sheet, headerMap, 'memo');
  const nextMemo = String(payload.memo == null ? '' : payload.memo);
  sheet.getRange(targetRowNo, memoCol).setValue(nextMemo);
  const masterMetaP202 = bumpPortalCustomerMasterMetaP202_(sheet, targetRowNo, 'webapp-customer-memo-save');

  let indexUpdate = queueCustomerSearchIndexRefreshAfterSaveP400_(targetRowNo, customerNo, ['memo'], masterMetaP202, 'WEBAPP_CUSTOMER_MEMO_SAVE');
  try { CacheService.getScriptCache().remove('PORTAL_DASHBOARD_V46_FAST_HOME'); } catch (err) {}

  return {
    ok: true,
    rowNo: targetRowNo,
    customerNo: customerNo,
    memo: nextMemo,
    indexUpdate: indexUpdate,
    masterVersion: masterMetaP202 && masterMetaP202.version || '',
    masterUpdatedAt: masterMetaP202 && masterMetaP202.updatedAt || '',
    masterEditor: masterMetaP202 && masterMetaP202.editor || '',
    savedAt: new Date().toISOString(),
    message: '메모 저장 완료'
  };
}

function updateCustomerSearchIndexMemoFast_(rowNo, memoValue, meta) {
  rowNo = Number(rowNo) || 0;
  if (!rowNo || rowNo < PORTAL_CONFIG.DATA_START_ROW) return { ok: false, reason: 'invalid rowNo' };

  const indexSheet = ensureCustomerSearchIndexSheet_(getWebAppDbSpreadsheet_());
  const lastRow = indexSheet.getLastRow();
  if (lastRow < 2) return { ok: false, reason: 'empty index' };

  const headers = indexSheet.getRange(1, 1, 1, Math.max(indexSheet.getLastColumn(), getCustomerSearchIndexHeadersK2_().length)).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const rowNoCol = headers.indexOf('rowNo') + 1;
  const memoCol = headers.indexOf('메모요약') + 1;
  const updatedCol = headers.indexOf('인덱스갱신시각') + 1;
  const masterUpdatedCol = headers.indexOf('원본수정시각') + 1;
  const masterVersionCol = headers.indexOf('마스터원본버전') + 1;
  const masterEditorCol = headers.indexOf('최종수정자') + 1;
  if (!rowNoCol || !memoCol) return { ok: false, reason: 'missing index headers' };
  meta = meta || {};

  const rowNoValues = indexSheet.getRange(2, rowNoCol, lastRow - 1, 1).getDisplayValues();
  let targetRow = 0;
  for (let i = 0; i < rowNoValues.length; i++) {
    if (Number(rowNoValues[i][0]) === rowNo) { targetRow = i + 2; break; }
  }
  if (!targetRow) return { ok: false, reason: 'index row not found' };

  const now = new Date();
  const ts = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  indexSheet.getRange(targetRow, memoCol).setValue(shortenTextForIndex_(memoValue, PORTAL_CONFIG.CUSTOMER_INDEX_MEMO_MAX_LENGTH || 350));
  if (updatedCol) indexSheet.getRange(targetRow, updatedCol).setValue(ts);
  if (masterUpdatedCol && meta.updatedAt) indexSheet.getRange(targetRow, masterUpdatedCol).setValue(meta.updatedAt);
  if (masterVersionCol && meta.version) indexSheet.getRange(targetRow, masterVersionCol).setValue(meta.version);
  if (masterEditorCol && meta.editor) indexSheet.getRange(targetRow, masterEditorCol).setValue(meta.editor);

  const touched = touchCustomerSearchIndexVersion_(now);
  return { ok: true, rowNo: rowNo, indexRow: targetRow, version: touched.version, builtAt: touched.builtAt };
}




// STEP42/P420: 계약시작일/계약종료일 저장·표시 보정
function parsePortalContractDateP420_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  const text = String(value || '').trim().replace(/\s*([.\/-])\s*/g, '$1');
  if (!text) return null;
  let m = text.match(/^(\d{2})[.\/-](\d{1,2})[.\/-](\d{1,2})\.?$/);
  if (m) return new Date(2000 + Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  m = text.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})\.?$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  m = text.match(/^(\d{2})년\s*(\d{1,2})월\s*(\d{1,2})일$/);
  if (m) return new Date(2000 + Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  m = text.match(/^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d = new Date(text);
  return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function formatPortalContractDateForDisplayP420_(value) {
  const d = parsePortalContractDateP420_(value);
  if (!d) return String(value == null ? '' : value).trim();
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy.MM.dd.');
}

function getPortalDefaultContractStartDateP420_() {
  return new Date(2026, 6, 1); // 2026-07-01
}

function calculatePortalContractEndDateP420_(startValue, contractUnitValue) {
  const start = parsePortalContractDateP420_(startValue) || getPortalDefaultContractStartDateP420_();
  const months = normalizePortalContractUnitMonthsP280_(contractUnitValue) || 12;
  return new Date(start.getFullYear(), start.getMonth() + months, start.getDate() - 1);
}

function normalizePortalContractDatePayloadFieldsP420_(values, options) {
  options = options || {};
  values = Object.assign({}, values || {});
  const hasStart = Object.prototype.hasOwnProperty.call(values, 'contractStartDate');
  const hasEnd = Object.prototype.hasOwnProperty.call(values, 'contractEndDate');
  const hasUnit = Object.prototype.hasOwnProperty.call(values, 'contractUnit');

  if (options.requireFull && !hasStart && !String(values.contractStartDate || '').trim()) {
    values.contractStartDate = formatPortalContractDateForDisplayP420_(getPortalDefaultContractStartDateP420_());
  }

  if ((options.requireFull || hasStart || hasUnit) && !hasEnd) {
    const startText = values.contractStartDate || (options.current && options.current.contractStartDate) || formatPortalContractDateForDisplayP420_(getPortalDefaultContractStartDateP420_());
    const unitText = values.contractUnit || (options.current && options.current.contractUnit) || '12';
    values.contractEndDate = formatPortalContractDateForDisplayP420_(calculatePortalContractEndDateP420_(startText, unitText));
  }

  ['contractStartDate', 'contractEndDate'].forEach(function(key) {
    if (Object.prototype.hasOwnProperty.call(values, key)) {
      values[key] = formatPortalContractDateForDisplayP420_(values[key]);
    }
  });
  return values;
}

function isPortalContractDateKeyP420_(key) {
  key = String(key || '').trim();
  return key === 'contractStartDate' || key === 'contractEndDate';
}

function applyPortalContractDateCellFormatP420_(sheet, rowNo, col, key) {
  if (!sheet || !rowNo || !col || !isPortalContractDateKeyP420_(key)) return;
  try { sheet.getRange(rowNo, col).setNumberFormat('yyyy.MM.dd.'); } catch (err) {}
}

// PATCH P280: 계약조건 숫자형 저장 호환
// - 마스터시트 U/W/X(계약단위/유지점검/성능점검)는 숫자 셀을 기준으로 사용합니다.
// - 화면에는 "12개월", "2회"처럼 보여도 저장 payload와 시트 write 값은 12, 2 같은 숫자로 정규화합니다.
function isPortalZeroDateLikeP340_(value) {
  // Google Sheets에서 숫자 0이 날짜 서식 컬럼에 들어가면
  // 1899. 12. 30 또는 1899-12-30처럼 표시될 수 있습니다.
  // 유지점검/성능점검 0회가 날짜로 보이지 않게 여기서 0으로 복구합니다.
  if (value instanceof Date) {
    const y = value.getFullYear();
    const m = value.getMonth() + 1;
    const d = value.getDate();
    return y === 1899 && m === 12 && (d === 30 || d === 31);
  }
  const text = String(value == null ? '' : value).trim();
  if (!text) return false;
  const compact = text.replace(/\s+/g, '');
  if (/^1899[.\-\/]12[.\\/](30|31)$/.test(compact)) return true;
  if (/^1899년12월(30|31)일$/.test(compact)) return true;
  return false;
}

function normalizePortalContractUnitMonthsP280_(value) {
  if (isPortalZeroDateLikeP340_(value)) return '';
  const text = String(value == null ? '' : value).trim();
  if (!text || text === '-' || text === '–' || text === '—') return '';
  if (/^1년$/.test(text)) return 12;
  if (/^반년$/.test(text)) return 6;
  const m = text.match(/\d+/);
  if (!m) return '';
  const n = Number(m[0]);
  if (!isFinite(n) || n < 1 || n > 12) return '';
  return n;
}

function normalizePortalInspectionCountP280_(value) {
  if (isPortalZeroDateLikeP340_(value)) return 0;
  const text = String(value == null ? '' : value).trim();
  if (!text || text === '-' || text === '–' || text === '—') return '';
  const m = text.match(/\d+/);
  if (!m) return '';
  const n = Number(m[0]);
  if (!isFinite(n) || n < 0 || n > 12) return '';
  return n;
}

function isPortalContractNumericKeyP340_(key) {
  key = String(key || '').trim();
  return key === 'contractUnit' || key === 'maintenance' || key === 'performance';
}

function isPortalMasterDateKeyP433_(key) {
  key = String(key || '').trim();
  return key === 'contractStartDate' || key === 'contractEndDate' || key === 'firstRegisteredAt';
}

function isPortalMasterNumberFormatKeyP433_(key) {
  key = String(key || '').trim();
  return key === 'contractUnit' || key === 'maintenance' || key === 'performance' || key === 'area' || key === 'finalQuote' || key === 'discountRate';
}

function getPortalMasterNumberFormatP433_(key) {
  key = String(key || '').trim();
  if (key === 'contractUnit') return '0"개월"';
  if (key === 'maintenance' || key === 'performance') return '0"회"';
  if (key === 'area') return '#,##0.##';
  if (key === 'finalQuote') return '₩#,##0';
  if (key === 'discountRate') return '0.##';
  return '';
}

function applyPortalMasterCellFormatP433_(sheet, rowNo, col, key) {
  if (!sheet || !rowNo || !col) return;
  const range = sheet.getRange(rowNo, col);
  try {
    if (isPortalMasterDateKeyP433_(key)) {
      range.setNumberFormat('yyyy.MM.dd.');
      return;
    }
    const fmt = getPortalMasterNumberFormatP433_(key);
    if (fmt) range.setNumberFormat(fmt);
  } catch (err) {}
}

function applyPortalContractCellNumberFormatP340_(sheet, rowNo, col, key) {
  applyPortalMasterCellFormatP433_(sheet, rowNo, col, key);
}

function normalizePortalContractFieldForDbP280_(key, value) {
  key = String(key || '').trim();
  if (key === 'contractUnit') {
    const months = normalizePortalContractUnitMonthsP280_(value);
    return months === '' ? '' : String(months);
  }
  if (key === 'maintenance' || key === 'performance') {
    const count = normalizePortalInspectionCountP280_(value);
    return count === '' ? '' : String(count);
  }
  return value == null ? '' : String(value).trim();
}

function normalizePortalContractPayloadFieldsP280_(values) {
  values = Object.assign({}, values || {});
  ['contractUnit', 'maintenance', 'performance'].forEach(function(key) {
    if (Object.prototype.hasOwnProperty.call(values, key)) {
      values[key] = normalizePortalContractFieldForDbP280_(key, values[key]);
    }
  });
  return values;
}

function parsePortalDecimalNumberP433_(value) {
  if (value == null || value === '') return '';
  if (value instanceof Date) return '';
  const text = String(value).trim();
  if (!text) return '';
  const normalized = text.replace(/,/g, '').replace(/₩/g, '').replace(/원/g, '').replace(/%/g, '').trim();
  const n = Number(normalized.replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? '' : n;
}

function roundPortalNumberP433_(num, decimals) {
  num = Number(num);
  if (!isFinite(num)) return '';
  decimals = Number(decimals) || 0;
  const factor = Math.pow(10, decimals);
  return Math.round(num * factor) / factor;
}

function formatPortalNumberTextP433_(num, decimals, comma) {
  num = Number(num);
  if (!isFinite(num)) return '';
  decimals = Number(decimals) || 0;
  const rounded = roundPortalNumberP433_(num, decimals);
  let text = String(rounded);
  if (text.indexOf('e') >= 0 || text.indexOf('E') >= 0) text = rounded.toFixed(decimals);
  if (text.indexOf('.') >= 0) text = text.replace(/0+$/g, '').replace(/\.$/, '');
  if (comma) {
    const parts = text.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    text = parts.join('.');
  }
  return text;
}

function getPortalMasterWriteValueP280_(key, value) {
  key = String(key || '').trim();
  if (key === 'contractUnit') {
    const months = normalizePortalContractUnitMonthsP280_(value);
    return months === '' ? '' : months;
  }
  if (key === 'maintenance' || key === 'performance') {
    const count = normalizePortalInspectionCountP280_(value);
    return count === '' ? '' : count;
  }
  if (key === 'area') {
    const n = parsePortalDecimalNumberP433_(value);
    return n === '' ? '' : roundPortalNumberP433_(n, 2);
  }
  if (key === 'finalQuote') {
    const n = parsePortalDecimalNumberP433_(value);
    return n === '' ? '' : Math.round(n);
  }
  if (key === 'discountRate') {
    const n = parsePortalDecimalNumberP433_(value);
    return n === '' ? '' : roundPortalNumberP433_(n, 2);
  }
  if (isPortalMasterDateKeyP433_(key)) {
    const d = parsePortalContractDateP420_(value);
    return d || '';
  }
  return String(value == null ? '' : value).trim();
}

function getPortalMasterCompareTextP280_(key, value) {
  key = String(key || '').trim();
  if (isPortalMasterDateKeyP433_(key)) return formatPortalContractDateForDisplayP420_(value);
  const writeValue = getPortalMasterWriteValueP280_(key, value);
  if (writeValue == null || writeValue === '') return '';
  if (key === 'contractUnit') return String(writeValue) + '개월';
  if (key === 'maintenance' || key === 'performance') return String(writeValue) + '회';
  if (key === 'area') return formatPortalNumberTextP433_(writeValue, 2, true);
  if (key === 'finalQuote') return '₩' + formatPortalNumberTextP433_(writeValue, 0, true);
  if (key === 'discountRate') return formatPortalNumberTextP433_(writeValue, 2, false);
  return String(writeValue).trim();
}

// PATCH P1-12: 계약조건 저장 검증 + 할인율 공란 0 보정
// - 변수 입력 부족이 시트 수식에 남지 않도록 계약 계산 필수값을 저장 전 차단합니다.
// - 할인율은 공란이면 0으로 강제 저장합니다.
const PORTAL_CONTRACT_REQUIRED_KEYS_P112 = ['area','contractUnit','appointment','maintenance','performance','vat'];
const PORTAL_CONTRACT_REQUIRED_LABELS_P112 = {
  area: '연면적',
  contractUnit: '계약단위',
  contractStartDate: '계약시작일',
  contractEndDate: '계약종료일',
  appointment: '관리자 선임 여부',
  maintenance: '유지점검 횟수',
  performance: '성능점검 횟수',
  vat: '부가세 여부'
};
const PORTAL_CONTRACT_SAVE_KEYS_P112 = PORTAL_CONTRACT_REQUIRED_KEYS_P112.concat(['discountRate','grade','finalQuote','targetFinalPrice']);

function isPortalContractValueBlankP112_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!text) return true;
  if (text === '-' || text === '–' || text === '—') return true;
  if (/^(선택|미선택|입력필요|변수입력부족|변수 입력 부족)$/i.test(text)) return true;
  return false;
}

function isPortalContractSaveTouchedP112_(values) {
  values = values || {};
  return PORTAL_CONTRACT_SAVE_KEYS_P112.some(function(key) {
    return Object.prototype.hasOwnProperty.call(values, key);
  });
}

function preparePortalContractValuesForSaveP112_(values, options) {
  options = options || {};
  values = normalizePortalContractPayloadFieldsP280_(Object.assign({}, values || {}));
  const requireFull = !!options.requireFull;
  const touchedContractP461 = isPortalContractSaveTouchedP112_(values);
  const touchedContractDateP461 = requireFull || Object.prototype.hasOwnProperty.call(values, 'contractStartDate') || Object.prototype.hasOwnProperty.call(values, 'contractEndDate') || Object.prototype.hasOwnProperty.call(values, 'contractUnit');

  // P461: 상태/메모/주소 같은 일반 필드 저장에서 계약조건 보정을 위해 마스터 row 전체를 읽지 않습니다.
  // P460 로그상 status 1칸 저장도 6~12초 걸렸는데, 불필요한 readMasterRowObject_가 주요 병목이었습니다.
  let currentForDatesP420 = null;
  let currentObjP461 = null;
  if (touchedContractDateP461 && options && options.sheet && options.rowNo) {
    try {
      currentObjP461 = readMasterRowObject_(options.sheet, Number(options.rowNo));
      currentForDatesP420 = {
        contractStartDate: getCustomerMasterHeaderValueK2_(currentObjP461, 'contractStartDate'),
        contractEndDate: getCustomerMasterHeaderValueK2_(currentObjP461, 'contractEndDate'),
        contractUnit: getCustomerMasterHeaderValueK2_(currentObjP461, 'contractUnit')
      };
    } catch (err) {}
  }
  values = normalizePortalContractDatePayloadFieldsP420_(values, { requireFull: requireFull, current: currentForDatesP420 });
  const shouldValidate = requireFull || touchedContractP461;
  if (!shouldValidate) return values;

  const effective = {};
  if (options.sheet && options.rowNo) {
    try {
      const obj = currentObjP461 || readMasterRowObject_(options.sheet, Number(options.rowNo));
      PORTAL_CONTRACT_REQUIRED_KEYS_P112.concat(['discountRate']).forEach(function(key) {
        effective[key] = getCustomerMasterHeaderValueK2_(obj, key);
      });
    } catch (err) {}
  }

  Object.keys(values || {}).forEach(function(key) { effective[key] = values[key]; });

  if (isPortalContractValueBlankP112_(effective.discountRate)) {
    effective.discountRate = '0';
    values.discountRate = '0';
  }

  const missing = [];
  PORTAL_CONTRACT_REQUIRED_KEYS_P112.forEach(function(key) {
    if (isPortalContractValueBlankP112_(effective[key])) missing.push(PORTAL_CONTRACT_REQUIRED_LABELS_P112[key] || key);
  });

  if (missing.length) {
    throw new Error('계약조건 저장 불가: ' + missing.join(', ') + '을(를) 입력해야 합니다. 할인율은 공란이면 0으로 저장됩니다.');
  }
  return values;
}

function shouldRefreshIndexFromMasterAfterContractSaveP112_(changedKeys) {
  changedKeys = changedKeys || [];
  return changedKeys.some(function(key) {
    return PORTAL_CONTRACT_SAVE_KEYS_P112.indexOf(String(key || '')) >= 0;
  });
}



function normalizePortalQuoteBasisMonthKeyP463_(value) {
  const n = normalizePortalContractUnitMonthsP280_(value);
  return n === '' ? '' : String(n);
}
function makePortalQuoteBasisObjectP463_(row, map, grade) {
  const norm = normalizeGrade_(grade);
  let areaAddUnit = parseMoney_(cellByHeaderIndex_(row, map, ['단가_연면적가산', '연면적가산', '연면적 가산', '면적가산단가', '단가_면적가산']));
  // P465: 기존 계약기준/캐시에 가산단가가 비어 있어도 특급 기본 가산단가 200,000원을 보장합니다.
  if (!(areaAddUnit > 0) && norm.indexOf('특급') >= 0) areaAddUnit = 200000;
  return {
    grade: String(grade || '').trim(),
    appointmentUnit: parseMoney_(cellByHeaderIndex_(row, map, ['단가_선임', '선임단가', '관리자선임단가'])),
    maintenanceUnit: parseMoney_(cellByHeaderIndex_(row, map, ['단가_유지', '유지단가', '유지점검단가'])),
    performanceUnit: parseMoney_(cellByHeaderIndex_(row, map, ['단가_성능', '성능단가', '성능점검단가'])),
    areaAddUnit: areaAddUnit
  };
}

function calculatePortalSpecialAreaSurchargeP464_(area, basis) {
  area = Number(area) || 0;
  basis = basis || {};
  const basisGrade = normalizeGrade_(basis.grade || basis.managementGrade || '');
  let addUnit = Number(basis.areaAddUnit) || Number(basis.specialAreaAddUnit) || 0;
  if (!(addUnit > 0) && basisGrade.indexOf('특급') >= 0) addUnit = 200000;
  const flooredArea = Math.floor(Math.max(area, 0) / 10000) * 10000;
  if (!(area > 90000) || !(addUnit > 0)) {
    return { areaAddUnit: addUnit, flooredArea: flooredArea, surchargeArea: 0, surchargeUnits: 0, surchargeAmount: 0 };
  }
  // P465: 특급은 90,000㎡까지 기본단가 적용, 초과분은 10,000㎡ 미만 절삭 후 10,000㎡당 단가_연면적가산을 더합니다.
  // 예: 197,200㎡ → 190,000㎡로 보고, (190,000-90,000)/10,000=10단위 × 200,000 = 2,000,000원 가산.
  const surchargeArea = Math.max(0, flooredArea - 90000);
  const surchargeUnits = Math.floor(surchargeArea / 10000);
  const surchargeAmount = surchargeUnits * addUnit;
  return { areaAddUnit: addUnit, flooredArea: flooredArea, surchargeArea: surchargeArea, surchargeUnits: surchargeUnits, surchargeAmount: surchargeAmount };
}

function getPortalQuoteBasisMapP462() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'PORTAL_QUOTE_BASIS_MAP_P465';
  try {
    const raw = cache.get(cacheKey);
    if (raw) return JSON.parse(raw);
  } catch (err) {}
  const ss = getMasterSpreadsheet_();
  const sheet = ss.getSheetByName('계약기준');
  if (!sheet) throw new Error('계약기준 시트를 찾지 못했습니다.');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const result = { ok: true, loadedAt: new Date().toISOString(), basis: {}, byGradeMonth: {} };
  if (lastRow < 2) return result;
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h){ return String(h || '').trim(); });
  const map = {};
  headers.forEach(function(h, i) { if (h) map[h] = i; });
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  values.forEach(function(row) {
    const grade = cellByHeaderIndex_(row, map, ['등급', '관리등급']);
    const norm = normalizeGrade_(grade);
    if (!norm) return;
    const basis = makePortalQuoteBasisObjectP463_(row, map, grade);
    const monthValue = cellByHeaderIndex_(row, map, ['계약단위', '계약개월', '개월', '계약기간']);
    const monthKey = normalizePortalQuoteBasisMonthKeyP463_(monthValue);
    if (monthKey) result.byGradeMonth[norm + '|' + monthKey] = basis;
    if (!result.basis[norm]) result.basis[norm] = basis;
  });
  try { cache.put(cacheKey, JSON.stringify(result), 21600); } catch (err) {}
  return result;
}

function calculateQuoteDiscount(payload) {
  payload = payload || {};

  let grade = String(payload.grade || '').trim();
  const area = parseMoney_(payload.area);
  if (!grade && area > 0) {
    grade = calculateGradeByArea_(area);
  }

  const discountTextRaw = String(payload.discountRate || '').trim();
  const discountText = discountTextRaw === '' ? '0' : discountTextRaw;
  const contractUnitText = String(payload.contractUnit || '').trim();
  const appointmentText = String(payload.appointment || '').trim();
  const maintenanceText = String(payload.maintenance || '').trim();
  const performanceText = String(payload.performance || '').trim();
  const vatText = String(payload.vat || '').trim();
  const normalizedContractMonthsP280 = normalizePortalContractUnitMonthsP280_(contractUnitText);
  const normalizedMaintenanceCountP280 = normalizePortalInspectionCountP280_(maintenanceText);
  const normalizedPerformanceCountP280 = normalizePortalInspectionCountP280_(performanceText);

  const anyInput = [grade, discountTextRaw, contractUnitText, appointmentText, maintenanceText, performanceText, vatText, String(payload.targetFinalPrice || '').trim(), String(payload.area || '').trim()]
    .some(v => String(v || '').trim() !== '');
  if (!anyInput) {
    return { ok: true, blank: true, finalPrice: '', finalPriceText: '', message: '' };
  }

  const missing = [];
  if (!(area > 0)) missing.push('연면적');
  if (!grade) missing.push('관리등급');
  if (isPortalContractValueBlankP112_(contractUnitText) || normalizedContractMonthsP280 === '') missing.push('계약단위');
  if (isPortalContractValueBlankP112_(appointmentText)) missing.push('관리자 선임 여부');
  if (isPortalContractValueBlankP112_(maintenanceText) || normalizedMaintenanceCountP280 === '') missing.push('유지점검');
  if (isPortalContractValueBlankP112_(performanceText) || normalizedPerformanceCountP280 === '') missing.push('성능점검');
  if (isPortalContractValueBlankP112_(vatText)) missing.push('부가세');

  if (missing.length) {
    throw new Error('변수 입력 부족: ' + missing.join(', ') + '을(를) 입력해야 합니다. 할인율은 공란이면 0%로 계산합니다.');
  }

  const basis = getContractBasisForGrade_(grade, contractUnitText || String(normalizedContractMonthsP280 || 12));
  if (!basis) throw new Error('계약기준 시트에서 관리등급 [' + (grade || '-') + '] 기준단가를 찾지 못했습니다.');

  const months = normalizedContractMonthsP280 || 12;
  const hasAppointment = appointmentText === '선임' || (appointmentText.indexOf('선임') >= 0 && appointmentText.indexOf('비선임') < 0 && appointmentText.indexOf('해당없음') < 0);
  const maintenanceCount = normalizedMaintenanceCountP280 === '' ? 0 : normalizedMaintenanceCountP280;
  const performanceCount = normalizedPerformanceCountP280 === '' ? 0 : normalizedPerformanceCountP280;

  const appointmentAmount = hasAppointment ? basis.appointmentUnit * months : 0;
  const maintenanceAmount = basis.maintenanceUnit * maintenanceCount;
  const performanceAmount = basis.performanceUnit * performanceCount;
  const specialAreaSurcharge = normalizeGrade_(grade).indexOf('특급') >= 0 ? calculatePortalSpecialAreaSurchargeP464_(area, basis) : { areaAddUnit: Number(basis.areaAddUnit) || 0, flooredArea: Math.floor(Math.max(area, 0) / 10000) * 10000, surchargeArea: 0, surchargeUnits: 0, surchargeAmount: 0 };
  const areaSurchargeAmount = Number(specialAreaSurcharge.surchargeAmount) || 0;
  const subtotal = appointmentAmount + maintenanceAmount + performanceAmount + areaSurchargeAmount;

  const vatMultiplier = vatText.indexOf('포함') >= 0 ? 1.1 : 1;
  const originPrice = subtotal * vatMultiplier;

  const targetFinal = parseMoney_(payload.targetFinalPrice);
  let discountRate = parseFloat(String(discountText).replace(/[^0-9.\-]/g, ''));
  let finalPrice = 0;

  if (targetFinal > 0 && originPrice > 0) {
    discountRate = (1 - targetFinal / originPrice) * 100;
    finalPrice = roundDownToUnit_(targetFinal, 10000);
  } else if (!isNaN(discountRate)) {
    finalPrice = roundDownToUnit_(originPrice * (1 - discountRate / 100), 10000);
  } else {
    discountRate = 0;
    finalPrice = roundDownToUnit_(originPrice, 10000);
  }

  const roundedDiscountRate = Math.round(discountRate * 100) / 100;

  return {
    ok: true,
    grade: grade,
    months: months,
    contractUnit: String(months),
    hasAppointment: hasAppointment,
    maintenanceCount: maintenanceCount,
    performanceCount: performanceCount,
    appointmentUnit: basis.appointmentUnit,
    maintenanceUnit: basis.maintenanceUnit,
    performanceUnit: basis.performanceUnit,
    areaAddUnit: Number(basis.areaAddUnit) || 0,
    appointmentAmount: appointmentAmount,
    maintenanceAmount: maintenanceAmount,
    performanceAmount: performanceAmount,
    areaSurchargeAmount: areaSurchargeAmount,
    specialAreaSurcharge: specialAreaSurcharge,
    subtotal: subtotal,
    vatMultiplier: vatMultiplier,
    originPrice: originPrice,
    discountRate: roundedDiscountRate,
    finalPrice: finalPrice,
    finalPriceText: formatWon_(finalPrice),
    message: '계산 완료'
  };
}

function getContractBasisForGrade_(grade, contractUnit) {
  grade = String(grade || '').trim();
  if (!grade) return null;
  const ss = getMasterSpreadsheet_();
  const sheet = ss.getSheetByName('계약기준');
  if (!sheet) throw new Error('계약기준 시트를 찾지 못했습니다.');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return null;

  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h || '').trim());
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i; });
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

  const gradeNorm = normalizeGrade_(grade);
  const requestedMonth = normalizePortalQuoteBasisMonthKeyP463_(contractUnit);
  let fallback = null;
  for (const row of values) {
    const rowGrade = cellByHeaderIndex_(row, map, ['등급', '관리등급']);
    if (normalizeGrade_(rowGrade) !== gradeNorm) continue;
    const basis = makePortalQuoteBasisObjectP463_(row, map, rowGrade);
    const rowMonth = normalizePortalQuoteBasisMonthKeyP463_(cellByHeaderIndex_(row, map, ['계약단위', '계약개월', '개월', '계약기간']));
    if (requestedMonth && rowMonth && rowMonth === requestedMonth) return basis;
    if (!fallback) fallback = basis;
  }
  return fallback;
}

function getNewCustomerTemplate() {
  const emptyObj = {};
  const todayText = formatPortalContractDateForDisplayP420_(new Date());
  const detail = {
    rowNo: 0,
    isNew: true,
    company: '신규 고객 등록',
    customerNo: '',
    address: '',
    contact: '',
    phone: '',
    directPhone: '',
    email: '',
    salesRep: '',
    vendor: '',
    status: '견적제출완료',
    customerRank: '',
    statusOptions: buildStatusOptions_('견적제출완료'),
    customerRankOptions: buildCustomerRankOptions_(''),
    memo: '',
    lastSent: '',
    sentAt: '',
    contactRounds: PORTAL_CONFIG.CONTACT_ROUNDS,
    contactMethods: PORTAL_CONFIG.CONTACT_METHODS,
    detailFields: buildDetailFieldValues_(emptyObj),
    quoteCalcDefaults: {},
    contactLogs: [],
    raw: {}
  };

  // 신규 등록 폼 기본값 보정
  const setField = function(section, key, value) {
    const arr = detail.detailFields && detail.detailFields[section] || [];
    arr.forEach(function(f) { if (f.key === key) f.value = value; });
  };
  setField('basic', 'firstRegisteredAt', todayText);
  setField('basic', 'status', '견적제출완료');
  setField('contract', 'contractUnit', '12');
  setField('contract', 'contractStartDate', formatPortalContractDateForDisplayP420_(getPortalDefaultContractStartDateP420_()));
  setField('contract', 'contractEndDate', formatPortalContractDateForDisplayP420_(calculatePortalContractEndDateP420_(getPortalDefaultContractStartDateP420_(), '12')));
  setField('contract', 'vat', '별도');
  setField('contract', 'appointment', '선임');
  setField('contract', 'maintenance', '2');
  setField('contract', 'performance', '1');

  return detail;
}

function applyCustomerRegistrationCalculation(payload) {
  payload = payload || {};
  const values = payload.values || {};
  const area = parseMoney_(values.area);
  const grade = calculateGradeByArea_(area);
  if (!grade) throw new Error('연면적을 입력해야 관리등급을 계산할 수 있습니다.');

  const calcPayload = Object.assign({}, values, { grade: grade });
  const quote = calculateQuoteDiscount(calcPayload);

  return {
    ok: true,
    grade: grade,
    finalQuote: quote.finalPriceText || quote.message || '',
    finalQuoteRaw: quote.finalPrice || '',
    quote: quote,
    message: '관리등급과 최종견적가를 계산했습니다.'
  };
}

function saveRegistrationCustomer(payload) {
  payload = payload || {};
  const rowNoInput = Number(payload.rowNo || 0);
  let values = payload.values || {};
  values = normalizePortalCustomerLocationValues_(values);
  values = preparePortalContractValuesForSaveP112_(values, { requireFull: true });
  let calc = payload.calculation || {};

  // P0-2: 고객 등록/수정은 저장 버튼 1회로 계산값 적용과 저장을 함께 처리합니다.
  // 기존 applyCustomerRegistrationCalculation API는 호환용으로 유지하되, 저장 시 계산값이 없거나
  // autoCalculate가 지정된 경우 서버에서 관리등급/최종견적가를 다시 계산해 저장합니다.
  const shouldAutoCalculate = payload.autoCalculate !== false && (!calc || !calc.grade || !calc.finalQuote);
  if (shouldAutoCalculate) {
    calc = applyCustomerRegistrationCalculation({ values: values });
  }

  const ss = getMasterSpreadsheet_();
  const sheet = ss.getSheetByName(PORTAL_CONFIG.MASTER_SHEET_NAME);
  if (!sheet) throw new Error('마스터시트(신규)를 찾지 못했습니다.');

  let headerMap = getHeaderMap_(sheet);
  const allDefs = getPortalCustomerAllDetailDefsP436_();

  const isNew = !rowNoInput;
  const rowNo = isNew ? findFirstAvailableMasterCustomerRow_(sheet, headerMap) : rowNoInput;
  if (!isNew && rowNo < PORTAL_CONFIG.DATA_START_ROW) throw new Error('수정할 고객 행 번호가 올바르지 않습니다.');
  if (isNew && !rowNo) throw new Error('마스터시트에서 신규 고객을 넣을 빈 행을 찾지 못했습니다. 현재 영업 진행 상황~메모가 비어 있는 행을 먼저 확보해 주세요.');

  const currentObj = readMasterRowObject_(sheet, rowNo);
  const existingCustomerNo = getMasterFieldValue_(currentObj, 'customerNo');
  const existingRegisteredAt = getMasterFieldValue_(currentObj, 'firstRegisteredAt');

  const nextCustomerNo = isNew ? (existingCustomerNo || generateNextCustomerNo_(sheet)) : existingCustomerNo;
  const registeredAt = isNew
    ? (parsePortalContractDateP420_(existingRegisteredAt) || new Date())
    : existingRegisteredAt;

  const merged = Object.assign({}, values);
  if (isNew) {
    merged.customerNo = nextCustomerNo;
    merged.firstRegisteredAt = registeredAt;
    if (!String(merged.status || '').trim()) merged.status = '견적제출완료';
  }
  if (calc.grade) merged.grade = calc.grade;
  if (calc.finalQuote) merged.finalQuote = calc.finalQuote;

  allDefs.forEach(function(def) {
    if (!def || !def.key) return;

    // 기존 고객 수정 시 고객번호/최초등록일은 보존합니다.
    if (!isNew && (def.key === 'customerNo' || def.key === 'firstRegisteredAt')) return;

    // 신규/수정 등록 화면에서는 관리등급/최종견적가도 적용 계산값을 저장합니다.
    if (!Object.prototype.hasOwnProperty.call(merged, def.key)) return;

    const headerName = (def.headers && def.headers[0]) || def.label;
    let col = findFirstExistingHeaderCol_(headerMap, def.headers || []);
    if (!col) {
      col = ensureMasterColumn_(sheet, headerMap, headerName);
      headerMap = getHeaderMap_(sheet);
    }
    const writeRangeP340 = sheet.getRange(rowNo, col);
    applyPortalMasterCellFormatP433_(sheet, rowNo, col, def.key);
    writeRangeP340.setValue(getPortalMasterWriteValueP280_(def.key, merged[def.key]));
  });

  SpreadsheetApp.flush();

  let indexUpdate = null;
  try { indexUpdate = updateCustomerSearchIndexRow_(rowNo); } catch (err) { Logger.log('검색인덱스 신규/수정 고객 갱신 실패: ' + (err && err.stack || err)); }
  if (!isNew) {
    try { if (typeof syncContractCompleteFromCustomerMasterP420_ === 'function') syncContractCompleteFromCustomerMasterP420_({ sheet: sheet, rowNo: rowNo, customerNo: nextCustomerNo, changedKeys: Object.keys(merged || {}), source: 'registrationEditSave' }); } catch (syncErrP420) { Logger.log('수주확정/계약완료 동기화 실패: ' + (syncErrP420 && syncErrP420.stack || syncErrP420)); }
  }
  try { CacheService.getScriptCache().remove('PORTAL_DASHBOARD_V27'); } catch (err) {}
  try { CacheService.getScriptCache().remove('PORTAL_DASHBOARD_V46_FAST_HOME'); } catch (err) {}

  try {
    appendPortalActivityLog_({
      actionType: '고객정보수정',
      screen: '신규 고객 등록',
      rowNo: rowNo,
      customerNo: nextCustomerNo,
      company: values.company || '',
      summary: isNew ? '신규 고객 등록: ' + (values.company || nextCustomerNo || rowNo) : '등록 고객 수정: ' + (values.company || nextCustomerNo || rowNo),
      detail: { isNew: isNew, values: values, calculation: calc }
    });
  } catch (err) {}

  return {
    ok: true,
    indexUpdate: indexUpdate,
    isNew: isNew,
    rowNo: rowNo,
    customerNo: nextCustomerNo,
    calculation: {
      grade: calc.grade || '',
      finalQuote: calc.finalQuote || '',
      finalQuoteRaw: calc.finalQuoteRaw || '',
      quote: calc.quote || null
    },
    message: isNew ? ('신규 고객을 빈 슬롯 ' + rowNo + '행에 등록했습니다.') : '고객 정보를 수정 저장했습니다.',
    detail: getCustomerDetail(rowNo)
  };
}

function findFirstAvailableMasterCustomerRow_(sheet, headerMap) {
  const lastRow = sheet.getLastRow();
  if (lastRow < PORTAL_CONFIG.DATA_START_ROW) return 0;

  const statusCol = findMasterFieldCol_(headerMap, 'status');
  const firstRegCol = findMasterFieldCol_(headerMap, 'firstRegisteredAt');
  const regionCol = findMasterFieldCol_(headerMap, 'region');
  const vendorCol = findMasterFieldCol_(headerMap, 'vendor');
  const salesRepCol = findMasterFieldCol_(headerMap, 'salesRep');
  const companyCol = findMasterFieldCol_(headerMap, 'company');
  const memoCol = findMasterFieldCol_(headerMap, 'memo');

  const cols = [statusCol, firstRegCol, regionCol, vendorCol, salesRepCol, companyCol, memoCol]
    .filter(function(col, idx, arr) { return col && arr.indexOf(col) === idx; });

  // 헤더가 비정상이라 기준 컬럼을 못 찾으면 신규 저장을 중단합니다.
  if (!cols.length) return 0;

  const start = PORTAL_CONFIG.DATA_START_ROW;
  const values = sheet.getRange(start, 1, lastRow - start + 1, sheet.getLastColumn()).getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const hasAnyDataInCustomerArea = cols.some(function(col) {
      return String(row[col - 1] || '').trim() !== '';
    });
    if (!hasAnyDataInCustomerArea) {
      return start + i;
    }
  }

  return 0;
}

function calculateGradeByArea_(area) {
  area = Number(area) || 0;
  if (!area) return '';
  if (area < 10000) return '초급';
  if (area < 15000) return '초급';
  if (area < 30000) return '중급';
  if (area < 60000) return '고급';
  return '특급';
}

function generateNextCustomerNo_(sheet) {
  const headerMap = getHeaderMap_(sheet);
  const col = findMasterFieldCol_(headerMap, 'customerNo');
  if (!col) return '10000';
  const lastRow = sheet.getLastRow();
  if (lastRow < PORTAL_CONFIG.DATA_START_ROW) return '10000';
  const values = sheet.getRange(PORTAL_CONFIG.DATA_START_ROW, col, lastRow - PORTAL_CONFIG.DATA_START_ROW + 1, 1).getDisplayValues();
  let maxNo = 0;
  values.forEach(function(row) {
    const n = Number(String(row[0] || '').replace(/[^0-9]/g, ''));
    if (!isNaN(n) && n > maxNo) maxNo = n;
  });
  return String(maxNo ? maxNo + 1 : 10000);
}

