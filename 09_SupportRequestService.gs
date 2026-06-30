/***************************************
 * S1 Sales Portal - 09_SupportRequestService.gs
 * PATCH S: 권한_DB 기반 영업지원 완료처리 권한 적용
 ***************************************/



// PATCH S-FIX2: 공통 영업지원 옵션 상수는 00_Config.gs에서만 선언합니다.
// 이 파일에서는 PORTAL_SUPPORT_REQUEST_TYPES / PORTAL_SUPPORT_REQUESTER_OPTIONS /
// PORTAL_SUPPORT_STATUS_OPTIONS / PORTAL_SUPPORT_MEMO_HEADER를 절대 재선언하지 않습니다.

function getPortalSupportCurrentHandlerNameP260_(perm) {
  perm = perm || getPortalCurrentPermission_();
  const email = String((perm && perm.email) || getPortalSessionEmail_() || '').trim().toLowerCase();
  const map = PORTAL_CONFIG.USER_DISPLAY_NAME_MAP || {};
  if (email && map[email]) return String(map[email] || '').trim();

  const name = String(perm && perm.name || '').trim();
  if (name && name.indexOf('@') < 0) return name;

  try {
    const fallback = String(getCurrentUserLabel_() || '').trim();
    if (fallback) return fallback;
  } catch (err) {}
  return name || email || '웹앱사용자';
}

function getPortalSupportDelegatedRequesterNamesP250_() {
  return (PORTAL_SUPPORT_REQUESTER_OPTIONS || ['문형진', '방수원', '박새봄', '김경아', '최보람', '이옥희']).slice();
}

function canPortalChooseDelegatedSupportRequesterP250_(perm) {
  perm = perm || getPortalCurrentPermission_();
  return !!(perm && perm.active !== false && (perm.canUseAdminHome || perm.canCompleteSupport || perm.isAdmin));
}

function buildPortalSupportRequesterOptions_(currentUser, perm) {
  perm = perm || getPortalCurrentPermission_();
  currentUser = String(currentUser || getPortalCurrentUserName_() || '').trim();
  const out = [];
  function add(v) {
    v = String(v || '').trim();
    if (v && out.indexOf(v) < 0) out.push(v);
  }

  if (canPortalChooseDelegatedSupportRequesterP250_(perm)) {
    getPortalSupportDelegatedRequesterNamesP250_().forEach(add);
    add(currentUser);
  } else {
    add(currentUser);
  }
  return out;
}

function assertPortalSupportRequesterAllowedP250_(requester, perm) {
  perm = perm || getPortalCurrentPermission_();
  requester = String(requester || '').trim();
  const currentUser = getPortalCurrentUserName_();
  const options = buildPortalSupportRequesterOptions_(currentUser, perm);
  const normalizedRequester = normalizePortalNameForPermission_(requester);
  const ok = options.some(function(v) {
    return normalizePortalNameForPermission_(v) === normalizedRequester;
  });
  if (ok) return true;
  if (canPortalChooseDelegatedSupportRequesterP250_(perm)) {
    throw new Error('요청자는 지정된 요청자 목록에서만 선택할 수 있습니다.');
  }
  throw new Error('요청자는 본인만 선택할 수 있습니다.');
}


// v64: 영업지원요청 목록 조회 속도 개선용 서버 캐시/읽기 전용 시트 접근
// - 목록 조회 때마다 헤더 보정/서식 쓰기를 하지 않습니다.
// - 최근 목록은 ScriptCache에 row 데이터만 저장하고, 사용자명/권한은 매 호출마다 현재 계정 기준으로 붙입니다.
const PORTAL_SUPPORT_RECENT_CACHE_PREFIX_V64 = 'supportRecent:v64:';
const PORTAL_SUPPORT_CACHE_BUST_PROP_V64 = 'PORTAL_SUPPORT_CACHE_BUST_V64';
const PORTAL_SUPPORT_RECENT_CACHE_TTL_SEC_V64 = 90;

function getPortalSupportCacheBustV64_() {
  try {
    return PropertiesService.getScriptProperties().getProperty(PORTAL_SUPPORT_CACHE_BUST_PROP_V64) || '0';
  } catch (err) {
    return '0';
  }
}

function bumpPortalSupportCacheBustV64_() {
  try {
    PropertiesService.getScriptProperties().setProperty(PORTAL_SUPPORT_CACHE_BUST_PROP_V64, String(Date.now()));
  } catch (err) {}
}

function getPortalSupportSheetFastV64_(masterSs) {
  masterSs = masterSs || getMasterSpreadsheet_();
  let sheet = masterSs.getSheetByName(PORTAL_CONFIG.SUPPORT_SHEET_NAME);
  if (!sheet) {
    const target = String(PORTAL_CONFIG.SUPPORT_SHEET_NAME || '').replace(/\s+/g, '');
    sheet = masterSs.getSheets().find(function(sh) {
      return String(sh.getName() || '').replace(/\s+/g, '') === target;
    }) || null;
  }
  // 시트가 없을 때만 생성/헤더 보정 루틴을 탑니다. 평상시 조회는 쓰기 작업 없이 끝냅니다.
  return sheet || getPortalSupportSheetFromMasterV56_(masterSs);
}

function makePortalSupportRecentCacheKeyV64_(options) {
  options = options || {};
  return PORTAL_SUPPORT_RECENT_CACHE_PREFIX_V64 + [
    getPortalSupportCacheBustV64_(),
    Number(options.limit) || 100,
    Number(options.maxScanRows || options.readRows) || 2500,
    Number(options.chunkSize) || 250
  ].join(':');
}

function readPortalSupportRecentCacheV64_(key) {
  try {
    const raw = CacheService.getScriptCache().get(key);
    if (!raw) return null;
    const obj = JSON.parse(raw);
    if (!obj || !Array.isArray(obj.rows)) return null;
    return obj;
  } catch (err) {
    return null;
  }
}

function writePortalSupportRecentCacheV64_(key, rows, total, meta) {
  try {
    const slimRows = (rows || []).slice(0, 120).map(function(r) {
      return {
        rowNo: r.rowNo,
        receiptNo: r.receiptNo,
        requestType: r.requestType,
        requester: r.requester,
        handler: r.handler,
        customerNo: r.customerNo,
        customerName: r.customerName,
        requestText: r.requestText,
        processContent: r.processContent,
        status: r.status,
        completedAt: r.completedAt,
        masterMemo: r.masterMemo
      };
    });
    const payload = JSON.stringify({ rows: slimRows, total: Number(total) || 0, meta: meta || {}, cachedAt: new Date().toISOString() });
    // Apps Script CacheService 단일 값 제한을 넘으면 조용히 캐시를 포기합니다.
    if (payload.length < 95000) CacheService.getScriptCache().put(key, payload, PORTAL_SUPPORT_RECENT_CACHE_TTL_SEC_V64);
  } catch (err) {}
}

function getPortalSupportData(options) {
  options = options || {};
  const keyword = String(options.keyword || '').trim().toLowerCase();
  const statusFilter = String(options.status || '').trim();
  const limit = Math.max(1, Math.min(500, Number(options.limit) || 200));

  const perm = getPortalCurrentPermission_();
  const rows = getPortalSupportRowsFastV64_();
  const visibleRows = filterPortalSupportRowsByPermissionP115_(rows, perm);
  const filtered = filterPortalSupportRowsV55_(visibleRows, keyword, statusFilter)
    .sort(sortPortalSupportRowsDescV55_)
    .slice(0, limit);

  return {
    rows: filtered,
    total: visibleRows.length,
    filtered: filtered.length,
    requestTypes: PORTAL_SUPPORT_REQUEST_TYPES,
    statusOptions: PORTAL_SUPPORT_STATUS_OPTIONS,
    currentUser: getPortalCurrentUserName_(),
    requesterOptions: buildPortalSupportRequesterOptions_(getPortalCurrentUserName_(), perm),
    permission: sanitizePortalPermissionForClient_(perm),
    meta: { mode: 'full', recent: false, complete: true, limit: limit, keyword: keyword, status: statusFilter, loadedAt: new Date().toISOString() }
  };
}


// v55: 영업지원 요청 빠른 초기 로딩용 API
// - 전체 시트를 매번 읽지 않고, 최근 물리 행부터 제한된 범위만 먼저 읽습니다.
// - 프론트에서는 이 결과를 먼저 보여주고, 필요할 때만 전체 목록을 백그라운드에서 보강합니다.
function getPortalSupportRecentData(options) {
  options = options || {};
  const keyword = String(options.keyword || '').trim().toLowerCase();
  const statusFilter = String(options.status || '').trim();
  const limit = Math.max(10, Math.min(200, Number(options.limit) || 80));

  // v56: 하단 빈 행/서식/과거 이관 흔적 때문에 마지막 물리 행 근처만 읽으면 0건으로 보일 수 있어
  // MASTER_SPREADSHEET_ID 파일의 영업지원요청 시트를 뒤에서부터 chunk 단위로 스캔합니다.
  const maxScanRows = Math.max(limit, Math.min(5000, Number(options.maxScanRows || options.readRows) || 1800));
  const chunkSize = Math.max(50, Math.min(500, Number(options.chunkSize) || 220));

  const cacheable = !keyword && !statusFilter && options.noCache !== true;
  const cacheKey = cacheable ? makePortalSupportRecentCacheKeyV64_({ limit: limit, maxScanRows: maxScanRows, chunkSize: chunkSize }) : '';
  if (cacheable) {
    const cached = readPortalSupportRecentCacheV64_(cacheKey);
    if (cached) {
      const meta = cached.meta || {};
      meta.cacheHit = true;
      meta.loadedAt = meta.loadedAt || cached.cachedAt || new Date().toISOString();
      return buildPortalSupportDataResponseV55_(cached.rows || [], cached.total || (cached.rows || []).length, meta);
    }
  }

  const masterSs = getMasterSpreadsheet_();
  const sheet = getPortalSupportSheetFastV64_(masterSs);
  const headerMap = getPortalSupportHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  const dataStart = PORTAL_CONFIG.SUPPORT_DATA_START_ROW;
  const lastCol = Math.max(sheet.getLastColumn(), PORTAL_CONFIG.SUPPORT_HEADERS.length);
  const totalPhysical = Math.max(0, lastRow - dataStart + 1);

  if (lastRow < dataStart || totalPhysical <= 0) {
    return buildPortalSupportDataResponseV55_([], 0, {
      mode: 'recent',
      recent: true,
      complete: true,
      scannedRows: 0,
      physicalReadRows: 0,
      source: 'MASTER_SPREADSHEET_ID',
      spreadsheetId: masterSs.getId(),
      spreadsheetName: masterSs.getName(),
      sheetName: sheet.getName(),
      limit: limit,
      loadedAt: new Date().toISOString()
    });
  }

  let endRow = lastRow;
  let scannedPhysical = 0;
  let meaningfulScanned = 0;
  const matched = [];
  const seenRows = {};

  while (endRow >= dataStart && scannedPhysical < maxScanRows && matched.length < limit) {
    const remainingScan = maxScanRows - scannedPhysical;
    const size = Math.min(chunkSize, remainingScan, endRow - dataStart + 1);
    const startRow = endRow - size + 1;
    const values = sheet.getRange(startRow, 1, size, lastCol).getDisplayValues();

    // 최신 물리 행부터 먼저 처리
    for (let i = values.length - 1; i >= 0; i--) {
      const rowNo = startRow + i;
      if (seenRows[rowNo]) continue;
      seenRows[rowNo] = true;
      const item = buildPortalSupportRowObject_(values[i], rowNo, headerMap);
      const meaningful = item && (item.requestType || item.requester || item.customerNo || item.customerName || item.requestText || item.processContent || item.status);
      if (!meaningful) continue;
      meaningfulScanned++;
      if (filterPortalSupportRowsV55_([item], keyword, statusFilter).length) matched.push(item);
      if (matched.length >= limit) break;
    }

    scannedPhysical += size;
    endRow = startRow - 1;
  }

  const sorted = matched.sort(sortPortalSupportRowsDescV55_).slice(0, limit);
  const complete = scannedPhysical >= totalPhysical || endRow < dataStart;

  const meta = {
    mode: 'recent',
    recent: true,
    complete: complete,
    scannedRows: meaningfulScanned,
    physicalReadRows: scannedPhysical,
    maxScanRows: maxScanRows,
    limit: limit,
    keyword: keyword,
    status: statusFilter,
    source: 'MASTER_SPREADSHEET_ID',
    spreadsheetId: masterSs.getId(),
    spreadsheetName: masterSs.getName(),
    sheetName: sheet.getName(),
    lastRow: lastRow,
    lastCol: lastCol,
    loadedAt: new Date().toISOString()
  };
  if (cacheable) writePortalSupportRecentCacheV64_(cacheKey, sorted, totalPhysical, meta);
  return buildPortalSupportDataResponseV55_(sorted, totalPhysical, meta);
}

function getPortalSupportSheetFromMasterV56_(masterSs) {
  masterSs = masterSs || getMasterSpreadsheet_();
  let sheet = masterSs.getSheetByName(PORTAL_CONFIG.SUPPORT_SHEET_NAME);
  if (!sheet) {
    // 혹시 공백이 들어간 시트명으로 만들어져 있어도 기존 데이터를 우선 찾습니다.
    const target = String(PORTAL_CONFIG.SUPPORT_SHEET_NAME || '').replace(/\s+/g, '');
    sheet = masterSs.getSheets().find(function(sh) {
      return String(sh.getName() || '').replace(/\s+/g, '') === target;
    }) || null;
  }
  if (!sheet) {
    sheet = masterSs.insertSheet(PORTAL_CONFIG.SUPPORT_SHEET_NAME);
    sheet.getRange(1, 1).setValue('영업지원 요청 작성 방법: 고객번호 입력 후 불러오기 또는 직접 입력하여 요청사항을 작성합니다.');
    sheet.getRange(2, 1).setValue('서무 처리 후 처리내용, 완료 시각, 진행 상태를 입력합니다.');
    sheet.getRange(PORTAL_CONFIG.SUPPORT_HEADER_ROW, 1, 1, PORTAL_CONFIG.SUPPORT_HEADERS.length).setValues([PORTAL_CONFIG.SUPPORT_HEADERS]);
    sheet.setFrozenRows(PORTAL_CONFIG.SUPPORT_HEADER_ROW);
  }
  ensurePortalSupportHeaderLayout_(sheet);
  return sheet;
}

function debugPortalSupportSourceV56() {
  const masterSs = getMasterSpreadsheet_();
  const sheet = getPortalSupportSheetFromMasterV56_(masterSs);
  const headerMap = getPortalSupportHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), PORTAL_CONFIG.SUPPORT_HEADERS.length);
  const start = Math.max(PORTAL_CONFIG.SUPPORT_DATA_START_ROW, lastRow - 9);
  const values = lastRow >= PORTAL_CONFIG.SUPPORT_DATA_START_ROW
    ? sheet.getRange(start, 1, lastRow - start + 1, lastCol).getDisplayValues()
    : [];
  const sample = values.map(function(row, idx) {
    return buildPortalSupportRowObject_(row, start + idx, headerMap);
  }).filter(function(item) {
    return item && (item.receiptNo || item.requestType || item.customerNo || item.customerName || item.requestText || item.status);
  });
  return {
    ok: true,
    masterSpreadsheetId: masterSs.getId(),
    masterSpreadsheetName: masterSs.getName(),
    supportSheetName: sheet.getName(),
    supportSheetId: sheet.getSheetId(),
    configuredSupportSheetName: PORTAL_CONFIG.SUPPORT_SHEET_NAME,
    headerRow: PORTAL_CONFIG.SUPPORT_HEADER_ROW,
    dataStartRow: PORTAL_CONFIG.SUPPORT_DATA_START_ROW,
    lastRow: lastRow,
    lastCol: lastCol,
    headerMap: headerMap,
    sampleLastRows: sample.slice(-10)
  };
}

function filterPortalSupportRowsV55_(rows, keyword, statusFilter) {
  keyword = String(keyword || '').trim().toLowerCase();
  statusFilter = String(statusFilter || '').trim();
  return (rows || []).filter(function(item) {
    if (statusFilter && item.status !== statusFilter) return false;
    if (!keyword) return true;
    const text = [
      item.receiptNo,
      item.requestType,
      item.requester,
      item.handler,
      item.customerNo,
      item.customerName,
      item.requestText,
      item.processContent,
      item.status,
      item.masterMemo
    ].join(' ').toLowerCase();
    return text.indexOf(keyword) >= 0;
  });
}

function filterPortalSupportRowsByPermissionP115_(rows, perm) {
  rows = Array.isArray(rows) ? rows : [];
  perm = perm || getPortalCurrentPermission_();
  if (!perm || perm.active === false) return [];
  if (canPortalReadAllSupport_(perm)) return rows;
  return rows.filter(function(item) { return isPortalSupportRowMineP115_(item, perm); });
}

function isPortalSupportRowMineP115_(item, perm) {
  item = item || {};
  perm = perm || getPortalCurrentPermission_();
  const names = [];
  if (perm.name) names.push(perm.name);
  if (perm.salesRepName) names.push(perm.salesRepName);
  (perm.salesRepAliases || []).forEach(function(x) { if (x) names.push(x); });
  const fields = [item.requester, item.handler, item.customerName].map(normalizePortalNameForPermission_);
  return names.some(function(name) {
    const n = normalizePortalNameForPermission_(name || '');
    if (!n) return false;
    return fields.some(function(f) { return f && (f === n || f.indexOf(n) >= 0 || n.indexOf(f) >= 0); });
  });
}

function sortPortalSupportRowsDescV55_(a, b) {
  const an = Number(String(a && a.receiptNo || '').replace(/[^0-9]/g, '')) || 0;
  const bn = Number(String(b && b.receiptNo || '').replace(/[^0-9]/g, '')) || 0;
  if (an !== bn) return bn - an;
  return (Number(b && b.rowNo) || 0) - (Number(a && a.rowNo) || 0);
}

function buildPortalSupportDataResponseV55_(rows, total, meta) {
  meta = meta || {};
  const perm = getPortalCurrentPermission_();
  const visibleRows = filterPortalSupportRowsByPermissionP115_(rows || [], perm);
  return {
    rows: visibleRows || [],
    total: Number(total) || visibleRows.length || 0,
    filtered: (visibleRows || []).length,
    requestTypes: PORTAL_SUPPORT_REQUEST_TYPES,
    statusOptions: PORTAL_SUPPORT_STATUS_OPTIONS,
    currentUser: getPortalCurrentUserName_(),
    requesterOptions: buildPortalSupportRequesterOptions_(getPortalCurrentUserName_(), perm),
    permission: sanitizePortalPermissionForClient_(perm),
    meta: meta
  };
}

// v78: 상세/처리 팝업 빠른 로딩용 상세 캐시 API
// - 목록 행에서 즉시 팝업을 채우고, 서버 상세는 ScriptCache를 먼저 봅니다.
// - 저장 시 bumpPortalSupportCacheBustV64_()가 호출되어 목록/상세 캐시가 같이 무효화됩니다.
const PORTAL_SUPPORT_DETAIL_CACHE_PREFIX_V78 = 'supportDetail:v448:';
const PORTAL_SUPPORT_DETAIL_CACHE_TTL_SEC_V78 = 300;

function makePortalSupportDetailCacheKeyV78_(rowNo) {
  return PORTAL_SUPPORT_DETAIL_CACHE_PREFIX_V78 + getPortalSupportCacheBustV64_() + ':' + String(Number(rowNo) || 0);
}

function getPortalSupportRequestDetail(rowNo, options) {
  return getPortalSupportRequestDetailFastV78(rowNo, options || {});
}


function getPortalSupportMasterCustomerInfoForRequestP448_(customerNo, customerName) {
  customerNo = String(customerNo || '').trim();
  customerName = String(customerName || '').trim();
  const normalizedNo = customerNo.replace(/\s+/g, '').toLowerCase();
  const normalizedName = customerName.replace(/\s+/g, '').toLowerCase();

  if (!normalizedNo && !normalizedName) return null;

  try {
    const indexRows = getCustomerSearchIndexRows_();
    let hitIndex = null;
    if (normalizedNo) {
      hitIndex = indexRows.find(function(r) {
        const no = String(r.customerNo || '').replace(/\s+/g, '').toLowerCase();
        return no && no === normalizedNo;
      }) || null;
    }
    if (!hitIndex && normalizedName) {
      hitIndex = indexRows.find(function(r) {
        const company = String(r.company || '').replace(/\s+/g, '').toLowerCase();
        return company && company.indexOf(normalizedName) >= 0;
      }) || null;
    }
    if (hitIndex && hitIndex.rowNo) {
      const masterSs = getMasterSpreadsheet_();
      const masterSheet = masterSs.getSheetByName(PORTAL_CONFIG.MASTER_SHEET_NAME);
      if (masterSheet) {
        const obj = readMasterRowObject_(masterSheet, Number(hitIndex.rowNo));
        obj.__rowNo = Number(hitIndex.rowNo);
        return buildPortalSupportCustomerInfo_(obj);
      }
    }
  } catch (err) {
    Logger.log('영업지원 상세: 검색인덱스 기반 마스터 고객정보 조회 실패 - ' + (err && err.message || err));
  }

  try {
    const rows = getMasterObjects_();
    let hit = null;
    if (normalizedNo) {
      hit = rows.find(function(r) {
        const no = String(r['고객번호'] || '').replace(/\s+/g, '').toLowerCase();
        return no && no === normalizedNo;
      }) || null;
    }
    if (!hit && normalizedName) {
      hit = rows.find(function(r) {
        const company = String(getCompanyValue_(r) || '').replace(/\s+/g, '').toLowerCase();
        return company && company.indexOf(normalizedName) >= 0;
      }) || null;
    }
    return hit ? buildPortalSupportCustomerInfo_(hit) : null;
  } catch (err) {
    Logger.log('영업지원 상세: 마스터 직접 고객정보 조회 실패 - ' + (err && err.message || err));
    return null;
  }
}

function enrichPortalSupportItemWithMasterCustomerInfoP448_(item) {
  item = item || {};
  const masterInfo = getPortalSupportMasterCustomerInfoForRequestP448_(item.customerNo, item.customerName);
  if (!masterInfo) {
    item.customerInfoSource = 'supportSheet';
    return item;
  }

  item.customerNo = masterInfo.customerNo || item.customerNo || '';
  item.customerName = masterInfo.customerName || item.customerName || '';
  item.discountRate = masterInfo.discountRate || '';
  item.quoteAmount = masterInfo.quoteAmount || '';
  item.area = masterInfo.area || '';
  item.region = masterInfo.region || '';
  item.contactName = masterInfo.contactName || '';
  item.contactEmail = masterInfo.contactEmail || '';
  item.masterMemo = masterInfo.masterMemo || '';
  item.contractSummary = masterInfo.contractSummary || item.contractSummary || '';
  item.customerInfoSource = 'master';
  return item;
}

function getPortalSupportRequestDetailFastV78(rowNo, options) {
  options = options || {};
  rowNo = Number(rowNo);
  if (!rowNo || rowNo < PORTAL_CONFIG.SUPPORT_DATA_START_ROW) throw new Error('영업지원 요청 행 번호가 올바르지 않습니다.');

  const cacheKey = makePortalSupportDetailCacheKeyV78_(rowNo);
  if (options.force !== true) {
    try {
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) {
        const parsed = JSON.parse(cached);
        if (parsed && parsed.item) {
          parsed.item = enrichPortalSupportItemWithMasterCustomerInfoP448_(parsed.item);
          const perm = getPortalCurrentPermission_();
          parsed.fromCache = true;
          parsed.loadedAt = parsed.loadedAt || new Date().toISOString();
          parsed.currentUser = getPortalCurrentUserName_();
          parsed.requesterOptions = buildPortalSupportRequesterOptions_(parsed.currentUser, perm);
          parsed.permission = sanitizePortalPermissionForClient_(perm);
          return parsed;
        }
      }
    } catch (err) {}
  }

  const sheet = getPortalSupportSheetFastV64_(getMasterSpreadsheet_());
  const headerMap = getPortalSupportHeaderMap_(sheet);
  const values = sheet.getRange(rowNo, 1, 1, Math.max(sheet.getLastColumn(), PORTAL_CONFIG.SUPPORT_HEADERS.length)).getDisplayValues()[0];
  const item = enrichPortalSupportItemWithMasterCustomerInfoP448_(buildPortalSupportRowObject_(values, rowNo, headerMap));
  if (!item || (!item.requestText && !item.customerNo && !item.customerName)) throw new Error('영업지원 요청을 찾지 못했습니다.');

  const perm = getPortalCurrentPermission_();
  const result = {
    item: item,
    requestTypes: PORTAL_SUPPORT_REQUEST_TYPES,
    statusOptions: PORTAL_SUPPORT_STATUS_OPTIONS,
    currentUser: getPortalCurrentUserName_(),
    requesterOptions: buildPortalSupportRequesterOptions_(getPortalCurrentUserName_(), perm),
    permission: sanitizePortalPermissionForClient_(perm),
    fromCache: false,
    loadedAt: new Date().toISOString()
  };

  try {
    const json = JSON.stringify(result);
    if (json.length < 90000) CacheService.getScriptCache().put(cacheKey, json, PORTAL_SUPPORT_DETAIL_CACHE_TTL_SEC_V78);
  } catch (err) {}

  return result;
}

function getPortalSupportCustomerLookup(query) {
  query = String(query || '').trim();
  if (!query) throw new Error('고객번호 또는 고객명을 입력하세요.');

  const normalizedQuery = query.replace(/\s+/g, '').toLowerCase();
  let hitIndex = null;

  try {
    const indexRows = getCustomerSearchIndexRows_();
    hitIndex = indexRows.find(function(r) {
      const no = String(r.customerNo || '').replace(/\s+/g, '').toLowerCase();
      return no && no === normalizedQuery;
    }) || indexRows.find(function(r) {
      const company = String(r.company || '').replace(/\s+/g, '').toLowerCase();
      return company && company.indexOf(normalizedQuery) >= 0;
    });
  } catch (err) {
    Logger.log('영업지원 고객 조회: 검색인덱스 조회 실패, 마스터 직접 조회로 fallback - ' + (err && err.message || err));
  }

  if (hitIndex && hitIndex.rowNo) {
    const masterSs = getMasterSpreadsheet_();
    const masterSheet = masterSs.getSheetByName(PORTAL_CONFIG.MASTER_SHEET_NAME);
    if (!masterSheet) throw new Error('마스터시트(신규)를 찾지 못했습니다.');
    const obj = readMasterRowObject_(masterSheet, Number(hitIndex.rowNo));
    obj.__rowNo = Number(hitIndex.rowNo);
    return buildPortalSupportCustomerInfo_(obj);
  }

  // 검색인덱스_DB가 비어 있거나 아직 생성되지 않은 경우에만 전체 마스터 조회로 fallback합니다.
  const rows = getMasterObjects_();
  let hit = rows.find(function(r) {
    const no = String(r['고객번호'] || '').replace(/\s+/g, '').toLowerCase();
    return no && no === normalizedQuery;
  });
  if (!hit) {
    hit = rows.find(function(r) {
      const company = String(getCompanyValue_(r) || '').replace(/\s+/g, '').toLowerCase();
      return company && company.indexOf(normalizedQuery) >= 0;
    });
  }

  if (!hit) throw new Error('마스터시트에서 해당 고객을 찾지 못했습니다. 고객번호/회사명을 확인하세요.');
  return buildPortalSupportCustomerInfo_(hit);
}

function searchPortalSupportCustomers(keyword) {
  keyword = String(keyword || '').trim();
  if (!keyword) return [];
  const normalizedKeyword = keyword.replace(/\s+/g, '').toLowerCase();

  try {
    return getCustomerSearchIndexRows_()
      .filter(function(r) {
        const no = String(r.customerNo || '').replace(/\s+/g, '').toLowerCase();
        const company = String(r.company || '').replace(/\s+/g, '').toLowerCase();
        const contact = String(r.contact || '').replace(/\s+/g, '').toLowerCase();
        const email = String(r.email || '').replace(/\s+/g, '').toLowerCase();
        return (no && no.indexOf(normalizedKeyword) >= 0) ||
          (company && company.indexOf(normalizedKeyword) >= 0) ||
          (contact && contact.indexOf(normalizedKeyword) >= 0) ||
          (email && email.indexOf(normalizedKeyword) >= 0);
      })
      .slice(0, 20)
      .map(buildPortalSupportCustomerInfoFromIndex_);
  } catch (err) {
    Logger.log('영업지원 고객 검색: 검색인덱스 조회 실패, 마스터 직접 조회로 fallback - ' + (err && err.message || err));
  }

  return getMasterObjects_()
    .filter(function(r) {
      const no = String(r['고객번호'] || '').replace(/\s+/g, '').toLowerCase();
      const company = String(getCompanyValue_(r) || '').replace(/\s+/g, '').toLowerCase();
      const contact = String(getValueByHeaderCandidates_(r, ['고객사 담당자', '담당자', '담당자명']) || '').replace(/\s+/g, '').toLowerCase();
      return (no && no.indexOf(normalizedKeyword) >= 0) ||
        (company && company.indexOf(normalizedKeyword) >= 0) ||
        (contact && contact.indexOf(normalizedKeyword) >= 0);
    })
    .slice(0, 20)
    .map(buildPortalSupportCustomerInfo_);
}

function savePortalSupportRequest(payload) {
  assertPortalCanWriteSupport_();
  const result = runPortalSupportWriteLockedP210_('support-request-save', function() {
    return savePortalSupportRequestCoreP210_(payload || {});
  });

  // 활동로그는 핵심 저장 Lock 밖에서 처리합니다. 로그 실패가 접수/저장 성공을 막으면 안 됩니다.
  try {
    appendPortalActivityLog_({
      actionType: '영업지원요청',
      screen: '영업지원 요청',
      customerNo: result && result.audit ? result.audit.customerNo : '',
      company: result && result.audit ? result.audit.customerName : '',
      summary: result && result.audit ? result.audit.summary : '영업지원 요청 저장',
      detail: result && result.audit ? result.audit.detail : {}
    });
  } catch (err) {}

  if (result) delete result.audit;
  return result;
}

function runPortalSupportWriteLockedP210_(label, callback) {
  if (typeof withPortalScriptLockP201_ === 'function') {
    return withPortalScriptLockP201_(label || 'support-request-save', callback, { attempts: 5, waitMs: 900, sleepBaseMs: 220 });
  }
  return callback();
}

function savePortalSupportRequestCoreP210_(payload) {
  payload = payload || {};
  const permForSupport = getPortalCurrentPermission_();
  const sheet = ensurePortalSupportSheet_();
  const headerMap = getPortalSupportHeaderMap_(sheet);
  ensurePortalSupportHeadersForWriteP210_(sheet, headerMap);

  const clientRequestId = String(payload.clientRequestId || payload.requestId || '').trim();
  if (clientRequestId) {
    const dupSupport = findPortalSupportByClientRequestIdP26_(sheet, headerMap, clientRequestId);
    if (dupSupport) {
      return {
        ok: true,
        duplicate: true,
        rowNo: dupSupport.rowNo,
        receiptNo: dupSupport.receiptNo || '',
        message: '이미 처리된 영업지원 요청입니다.',
        item: dupSupport,
        audit: {
          customerNo: dupSupport.customerNo || '',
          customerName: dupSupport.customerName || '',
          summary: '중복 영업지원 요청 무시: ' + (dupSupport.requestText || ''),
          detail: { receiptNo: dupSupport.receiptNo || '', rowNo: dupSupport.rowNo, duplicate: true }
        }
      };
    }
  }

  let rowNo = Number(payload.rowNo) || 0;
  const isNew = !rowNo;
  if (isNew) {
    rowNo = findFirstBlankPortalSupportRow_(sheet, headerMap);
  } else if (rowNo < PORTAL_CONFIG.SUPPORT_DATA_START_ROW) {
    throw new Error('영업지원 요청 행 번호가 올바르지 않습니다.');
  }
  ensurePortalSupportRowExistsP210_(sheet, rowNo);

  const writeLastCol = getPortalSupportWriteLastColP210_(sheet, headerMap);
  let existingRowValues = [];
  if (!isNew) {
    existingRowValues = sheet.getRange(rowNo, 1, 1, writeLastCol).getDisplayValues()[0] || [];
  } else {
    existingRowValues = new Array(writeLastCol).fill('');
  }

  const existingItemForPermission = !isNew
    ? buildPortalSupportRowObject_(existingRowValues, rowNo, headerMap)
    : null;

  const requestType = String(payload.requestType || '').trim();
  const requester = String(payload.requester || '').trim() || getPortalCurrentUserName_();
  const customerNo = String(payload.customerNo || '').trim();
  const customerName = String(payload.customerName || '').trim();
  const requestText = String(payload.requestText || '').trim();
  let status = String(payload.status || '').trim() || '접수';

  if (!permForSupport.canCompleteSupport) {
    if (isNew) {
      status = '접수';
    } else {
      status = String(existingItemForPermission && existingItemForPermission.status || status || '접수').trim() || '접수';
    }
    if (String(payload.status || '').trim() === '완료') {
      throw new Error('영업지원 완료 처리 권한이 없습니다.');
    }
  }

  if (!requestType) throw new Error('업무유형을 선택하세요.');
  if (!requester) throw new Error('요청자를 입력하세요.');
  assertPortalSupportRequesterAllowedP250_(requester, permForSupport);
  if (!requestText) throw new Error('요청업무를 입력하세요.');

  const masterCustomerInfoP448 = getPortalSupportMasterCustomerInfoForRequestP448_(customerNo, customerName);
  const resolvedCustomerNoP448 = masterCustomerInfoP448 && masterCustomerInfoP448.customerNo ? masterCustomerInfoP448.customerNo : customerNo;
  const resolvedCustomerNameP448 = masterCustomerInfoP448 && masterCustomerInfoP448.customerName ? masterCustomerInfoP448.customerName : customerName;

  const receiptNo = String(payload.receiptNo || '').trim() || getExistingOrNextPortalSupportReceiptNo_(sheet, headerMap, rowNo);
  let completedAt = normalizePortalSupportCompletedAt_(payload.completedAt, status);
  let handlerValue = String(payload.handler || '').trim();
  let processContentValue = String(payload.processContent || '').trim();
  let autoSendCheckValue = String(payload.autoSendCheck || '').trim();

  const statusForHandlerP260 = String(status || '').trim();
  const supportProcessingTouchedP260 =
    statusForHandlerP260 === '처리중' ||
    statusForHandlerP260 === '완료' ||
    !!String(payload.completedAt || '').trim() ||
    !!String(payload.processContent || '').trim() ||
    !!String(payload.autoSendCheck || '').trim();
  if (permForSupport.canCompleteSupport && supportProcessingTouchedP260 && !handlerValue) {
    handlerValue = getPortalSupportCurrentHandlerNameP260_(permForSupport);
  }

  if (!permForSupport.canCompleteSupport) {
    if (isNew) {
      completedAt = '';
      handlerValue = '';
      processContentValue = '';
      autoSendCheckValue = '';
    } else {
      completedAt = String(existingItemForPermission && existingItemForPermission.completedAt || '').trim();
      handlerValue = String(existingItemForPermission && existingItemForPermission.handler || '').trim();
      processContentValue = String(existingItemForPermission && existingItemForPermission.processContent || '').trim();
      autoSendCheckValue = String(existingItemForPermission && existingItemForPermission.autoSendCheck || '').trim();
    }
  }

  const rowObj = {
    '접수번호': receiptNo,
    '업무유형': requestType,
    '요청자': requester,
    '처리자': handlerValue,
    '고객번호': resolvedCustomerNoP448,
    '고객명': resolvedCustomerNameP448,
    '요청업무': requestText,
    '할인율': masterCustomerInfoP448 ? String(masterCustomerInfoP448.discountRate || '').trim() : String(payload.discountRate || '').trim(),
    '견적금액': masterCustomerInfoP448 ? String(masterCustomerInfoP448.quoteAmount || '').trim() : String(payload.quoteAmount || '').trim(),
    '연면적': masterCustomerInfoP448 ? String(masterCustomerInfoP448.area || '').trim() : String(payload.area || '').trim(),
    '지역': masterCustomerInfoP448 ? String(masterCustomerInfoP448.region || '').trim() : String(payload.region || '').trim(),
    '담당자 이름': masterCustomerInfoP448 ? String(masterCustomerInfoP448.contactName || '').trim() : String(payload.contactName || '').trim(),
    '담당자 이메일': masterCustomerInfoP448 ? String(masterCustomerInfoP448.contactEmail || '').trim() : String(payload.contactEmail || '').trim(),
    '처리내용': processContentValue,
    '완료 시각': completedAt,
    '진행 상태': status,
    [PORTAL_SUPPORT_MEMO_HEADER]: masterCustomerInfoP448 ? String(masterCustomerInfoP448.masterMemo || '').trim() : String(payload.masterMemo || '').trim(),
    '자동발송 확인': autoSendCheckValue,
    '클라이언트요청ID': clientRequestId
  };

  const writtenRowValues = writePortalSupportRowAtomicP210_(sheet, rowNo, headerMap, rowObj, existingRowValues, writeLastCol);

  let masterApply = { applied: false };
  if (isNew && customerNo) {
    try {
      masterApply = appendPortalSupportRequestToMasterMemo_(payload, receiptNo);
    } catch (err) {
      Logger.log('영업지원 요청 마스터 메모 반영 실패: ' + (err && err.stack || err));
      masterApply = { applied: false, error: String(err && err.message || err) };
    }
  }

  SpreadsheetApp.flush();
  bumpPortalSupportCacheBustV64_();

  const item = enrichPortalSupportItemWithMasterCustomerInfoP448_(buildPortalSupportRowObject_(writtenRowValues, rowNo, headerMap));
  return {
    ok: true,
    rowNo: rowNo,
    receiptNo: receiptNo,
    masterApplied: !!(masterApply && masterApply.applied),
    masterApply: masterApply,
    message: isNew ? '영업지원 요청이 접수되었습니다.' : '영업지원 요청이 저장되었습니다.',
    item: item,
    audit: {
      customerNo: resolvedCustomerNoP448,
      customerName: resolvedCustomerNameP448,
      summary: (isNew ? '요청 접수: ' : '요청 저장: ') + requestText,
      detail: { receiptNo: receiptNo, requestType: requestType, status: status, rowNo: rowNo }
    }
  };
}

function ensurePortalSupportHeadersForWriteP210_(sheet, headerMap) {
  (PORTAL_CONFIG.SUPPORT_HEADERS || []).forEach(function(header) {
    if (!headerMap[header]) ensurePortalSupportColumn_(sheet, headerMap, header);
  });
}

function ensurePortalSupportRowExistsP210_(sheet, rowNo) {
  rowNo = Number(rowNo) || 0;
  if (!rowNo) return;
  const maxRows = sheet.getMaxRows();
  if (rowNo > maxRows) {
    sheet.insertRowsAfter(maxRows, rowNo - maxRows);
  }
}

function getPortalSupportWriteLastColP210_(sheet, headerMap) {
  const maxHeaderCol = (PORTAL_CONFIG.SUPPORT_HEADERS || []).reduce(function(maxCol, header) {
    return Math.max(maxCol, Number(headerMap[header]) || 0);
  }, 0);
  return Math.max(maxHeaderCol, sheet.getLastColumn(), PORTAL_CONFIG.SUPPORT_HEADERS.length, 1);
}

function writePortalSupportRowAtomicP210_(sheet, rowNo, headerMap, rowObj, existingRowValues, writeLastCol) {
  writeLastCol = Number(writeLastCol) || getPortalSupportWriteLastColP210_(sheet, headerMap);
  const nextRowValues = (existingRowValues || []).slice(0, writeLastCol);
  while (nextRowValues.length < writeLastCol) nextRowValues.push('');

  const writeCols = [];
  (PORTAL_CONFIG.SUPPORT_HEADERS || []).forEach(function(header) {
    const col = Number(headerMap[header]) || 0;
    if (!col) return;
    nextRowValues[col - 1] = rowObj[header] == null ? '' : rowObj[header];
    writeCols.push(col);
  });

  const groups = groupConsecutivePortalSupportColsP210_(writeCols);
  groups.forEach(function(group) {
    const width = group.end - group.start + 1;
    const slice = nextRowValues.slice(group.start - 1, group.end);
    sheet.getRange(rowNo, group.start, 1, width).setValues([slice]);
  });

  return nextRowValues;
}

function groupConsecutivePortalSupportColsP210_(cols) {
  cols = (cols || []).map(function(c) { return Number(c) || 0; }).filter(Boolean).sort(function(a, b) { return a - b; });
  const unique = [];
  cols.forEach(function(c) {
    if (unique.indexOf(c) < 0) unique.push(c);
  });
  const groups = [];
  unique.forEach(function(c) {
    const last = groups[groups.length - 1];
    if (!last || c > last.end + 1) {
      groups.push({ start: c, end: c });
    } else {
      last.end = c;
    }
  });
  return groups;
}

function findPortalSupportByClientRequestIdP26_(sheet, headerMap, clientRequestId) {
  clientRequestId = String(clientRequestId || '').trim();
  if (!clientRequestId) return null;
  const col = headerMap['클라이언트요청ID'] || ensurePortalSupportColumn_(sheet, headerMap, '클라이언트요청ID');
  const lastRow = sheet.getLastRow();
  if (lastRow < PORTAL_CONFIG.SUPPORT_DATA_START_ROW) return null;
  const values = sheet.getRange(PORTAL_CONFIG.SUPPORT_DATA_START_ROW, 1, lastRow - PORTAL_CONFIG.SUPPORT_DATA_START_ROW + 1, Math.max(sheet.getLastColumn(), PORTAL_CONFIG.SUPPORT_HEADERS.length)).getDisplayValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    if (String(row[col - 1] || '').trim() !== clientRequestId) continue;
    return buildPortalSupportRowObject_(row, PORTAL_CONFIG.SUPPORT_DATA_START_ROW + i, headerMap);
  }
  return null;
}

function ensurePortalSupportSheet_() {
  // v56: 영업지원요청은 포털 컨테이너/웹앱_DB가 아니라 반드시 MASTER_SPREADSHEET_ID 파일에서 읽고 씁니다.
  return getPortalSupportSheetFromMasterV56_(getMasterSpreadsheet_());
}

function ensurePortalSupportHeaderLayout_(sheet) {
  const currentLastCol = Math.max(sheet.getLastColumn(), PORTAL_CONFIG.SUPPORT_HEADERS.length);
  const headerValues = sheet.getRange(PORTAL_CONFIG.SUPPORT_HEADER_ROW, 1, 1, currentLastCol).getDisplayValues()[0];
  const existing = {};
  headerValues.forEach(function(h, i) {
    h = String(h || '').trim();
    if (h) existing[h] = i + 1;
  });

  if (!Object.keys(existing).length) {
    sheet.getRange(PORTAL_CONFIG.SUPPORT_HEADER_ROW, 1, 1, PORTAL_CONFIG.SUPPORT_HEADERS.length).setValues([PORTAL_CONFIG.SUPPORT_HEADERS]);
  } else {
    let nextCol = sheet.getLastColumn();
    PORTAL_CONFIG.SUPPORT_HEADERS.forEach(function(h) {
      if (!existing[h]) {
        nextCol += 1;
        sheet.getRange(PORTAL_CONFIG.SUPPORT_HEADER_ROW, nextCol).setValue(h);
        existing[h] = nextCol;
      }
    });
  }

  sheet.getRange(PORTAL_CONFIG.SUPPORT_HEADER_ROW, 1, 1, Math.max(sheet.getLastColumn(), PORTAL_CONFIG.SUPPORT_HEADERS.length))
    .setFontWeight('bold')
    .setBackground('#f2f4f7');
}

function tryAutoMigratePortalSupportFromMasterIfEmpty_(targetSheet) {
  if (!targetSheet) return false;

  const webAppDbId = String(PORTAL_CONFIG.WEBAPP_DB_SPREADSHEET_ID || '').trim();
  const masterId = String(PORTAL_CONFIG.MASTER_SPREADSHEET_ID || '').trim();
  if (!webAppDbId || !masterId || webAppDbId === masterId) return false;

  if (hasPortalSupportMeaningfulRows_(targetSheet)) return false;

  let sourceSheet = null;
  try {
    const masterSs = getMasterSpreadsheet_();
    sourceSheet = masterSs.getSheetByName(PORTAL_CONFIG.SUPPORT_SHEET_NAME);
  } catch (err) {
    Logger.log('영업지원요청 자동 이관: 마스터 파일 접근 실패 - ' + (err && err.message || err));
    return false;
  }

  if (!sourceSheet) return false;
  if (sourceSheet.getSheetId && targetSheet.getSheetId && sourceSheet.getSheetId() === targetSheet.getSheetId()) return false;
  if (!hasPortalSupportMeaningfulRows_(sourceSheet)) return false;

  const lastRow = sourceSheet.getLastRow();
  const lastCol = Math.max(sourceSheet.getLastColumn(), PORTAL_CONFIG.SUPPORT_HEADERS.length);
  const values = sourceSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  targetSheet.clearContents();
  targetSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  targetSheet.setFrozenRows(PORTAL_CONFIG.SUPPORT_HEADER_ROW);

  Logger.log('영업지원요청 자동 이관 완료: ' + (lastRow - PORTAL_CONFIG.SUPPORT_DATA_START_ROW + 1) + '행');
  return true;
}

function hasPortalSupportMeaningfulRows_(sheet) {
  if (!sheet) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < PORTAL_CONFIG.SUPPORT_DATA_START_ROW) return false;

  const lastCol = Math.max(sheet.getLastColumn(), PORTAL_CONFIG.SUPPORT_HEADERS.length);
  const headerMap = getPortalSupportHeaderMap_(sheet);
  const checkHeaders = ['업무유형', '요청자', '처리자', '고객번호', '고객명', '요청업무', '처리내용', '완료 시각', '진행 상태'];
  const values = sheet.getRange(PORTAL_CONFIG.SUPPORT_DATA_START_ROW, 1, lastRow - PORTAL_CONFIG.SUPPORT_DATA_START_ROW + 1, lastCol).getDisplayValues();

  return values.some(function(row) {
    return checkHeaders.some(function(h) {
      const col = headerMap[h];
      return col && String(row[col - 1] || '').trim() !== '';
    });
  });
}

function migratePortalSupportRequestsToWebAppDb() {
  // v34: 영업지원요청은 다시 마스터시트 파일을 기준으로 사용합니다.
  // 이전 v33 이관 함수명을 눌러도 데이터가 이동되지 않도록 안전하게 no-op 처리합니다.
  const sheet = ensurePortalSupportSheet_();
  return {
    ok: true,
    message: 'v34 기준 영업지원요청은 마스터시트 파일의 영업지원요청 시트를 사용합니다. 이관 작업은 필요하지 않습니다.',
    supportRequestLocation: 'MASTER_SPREADSHEET_ID',
    spreadsheetId: getMasterSpreadsheet_().getId(),
    sheetName: sheet.getName()
  };
}

function getPortalSupportHeaderMap_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), PORTAL_CONFIG.SUPPORT_HEADERS.length);
  const headers = sheet.getRange(PORTAL_CONFIG.SUPPORT_HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];
  const map = {};
  headers.forEach(function(h, i) {
    h = String(h || '').trim();
    if (h && !map[h]) map[h] = i + 1;
  });
  return map;
}

function ensurePortalSupportColumn_(sheet, headerMap, headerName) {
  if (headerMap[headerName]) return headerMap[headerName];
  const col = sheet.getLastColumn() + 1;
  sheet.getRange(PORTAL_CONFIG.SUPPORT_HEADER_ROW, col).setValue(headerName);
  headerMap[headerName] = col;
  return col;
}


function getPortalSupportRowsFastV64_() {
  const sheet = getPortalSupportSheetFastV64_(getMasterSpreadsheet_());
  const headerMap = getPortalSupportHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < PORTAL_CONFIG.SUPPORT_DATA_START_ROW) return [];

  const values = sheet.getRange(
    PORTAL_CONFIG.SUPPORT_DATA_START_ROW,
    1,
    lastRow - PORTAL_CONFIG.SUPPORT_DATA_START_ROW + 1,
    Math.max(sheet.getLastColumn(), PORTAL_CONFIG.SUPPORT_HEADERS.length)
  ).getDisplayValues();

  return values
    .map(function(row, idx) {
      return buildPortalSupportRowObject_(row, PORTAL_CONFIG.SUPPORT_DATA_START_ROW + idx, headerMap);
    })
    .filter(function(item) {
      return item && (item.requestType || item.requester || item.customerNo || item.customerName || item.requestText || item.processContent || item.status);
    });
}

function getPortalSupportRows_() {
  const sheet = ensurePortalSupportSheet_();
  const headerMap = getPortalSupportHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < PORTAL_CONFIG.SUPPORT_DATA_START_ROW) return [];

  const values = sheet.getRange(
    PORTAL_CONFIG.SUPPORT_DATA_START_ROW,
    1,
    lastRow - PORTAL_CONFIG.SUPPORT_DATA_START_ROW + 1,
    Math.max(sheet.getLastColumn(), PORTAL_CONFIG.SUPPORT_HEADERS.length)
  ).getDisplayValues();

  return values
    .map(function(row, idx) {
      return buildPortalSupportRowObject_(row, PORTAL_CONFIG.SUPPORT_DATA_START_ROW + idx, headerMap);
    })
    .filter(function(item) {
      return item && (item.requestType || item.requester || item.customerNo || item.customerName || item.requestText || item.processContent || item.status);
    });
}

function buildPortalSupportRowObject_(row, rowNo, headerMap) {
  function v(header) {
    const col = headerMap[header];
    return col ? String(row[col - 1] || '').trim() : '';
  }
  const item = {
    rowNo: rowNo,
    receiptNo: v('접수번호'),
    requestType: v('업무유형'),
    requester: v('요청자'),
    handler: v('처리자'),
    customerNo: v('고객번호'),
    customerName: v('고객명'),
    requestText: v('요청업무'),
    discountRate: v('할인율'),
    quoteAmount: v('견적금액'),
    area: v('연면적'),
    region: v('지역'),
    contactName: v('담당자 이름'),
    contactEmail: v('담당자 이메일'),
    processContent: v('처리내용'),
    completedAt: v('완료 시각'),
    status: v('진행 상태'),
    masterMemo: v(PORTAL_SUPPORT_MEMO_HEADER),
    autoSendCheck: v('자동발송 확인')
  };
  if (!item.status && (item.requestText || item.customerNo || item.customerName)) item.status = '접수';
  return item;
}

function findFirstBlankPortalSupportRow_(sheet, headerMap) {
  const lastRow = Math.max(sheet.getLastRow(), PORTAL_CONFIG.SUPPORT_DATA_START_ROW);
  const lastCol = Math.max(sheet.getLastColumn(), PORTAL_CONFIG.SUPPORT_HEADERS.length);
  const numRows = Math.max(1, lastRow - PORTAL_CONFIG.SUPPORT_DATA_START_ROW + 1);
  const values = sheet.getRange(PORTAL_CONFIG.SUPPORT_DATA_START_ROW, 1, numRows, lastCol).getDisplayValues();

  const checkHeaders = ['업무유형', '요청자', '처리자', '고객번호', '고객명', '요청업무', '처리내용', '완료 시각', '진행 상태'];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const hasData = checkHeaders.some(function(h) {
      const col = headerMap[h];
      return col && String(row[col - 1] || '').trim() !== '';
    });
    if (!hasData) return PORTAL_CONFIG.SUPPORT_DATA_START_ROW + i;
  }
  return lastRow + 1;
}

function getExistingOrNextPortalSupportReceiptNo_(sheet, headerMap, rowNo) {
  const receiptCol = headerMap['접수번호'];
  if (receiptCol && rowNo) {
    const existing = String(sheet.getRange(rowNo, receiptCol).getDisplayValue() || '').trim();
    if (existing) return existing;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < PORTAL_CONFIG.SUPPORT_DATA_START_ROW || !receiptCol) return '1';
  const values = sheet.getRange(PORTAL_CONFIG.SUPPORT_DATA_START_ROW, receiptCol, lastRow - PORTAL_CONFIG.SUPPORT_DATA_START_ROW + 1, 1).getDisplayValues();
  let maxNo = 0;
  values.forEach(function(row) {
    const n = Number(String(row[0] || '').replace(/[^0-9]/g, '')) || 0;
    if (n > maxNo) maxNo = n;
  });
  return String(maxNo + 1);
}

function normalizePortalSupportCompletedAt_(value, status) {
  const text = String(value || '').trim();
  if (text) return text;
  if (String(status || '').trim() === '완료') {
    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yy-MM-dd HH:mm');
  }
  return '';
}

function buildPortalSupportCustomerInfo_(r) {
  const contractUnit = getValueByHeaderCandidates_(r, ['계약단위']);
  const appointment = getValueByHeaderCandidates_(r, ['관리자\n선임 여부', '관리자 선임 여부', '관리자선임여부', '선임 여부']);
  const maintenance = getValueByHeaderCandidates_(r, ['유지점검']);
  const performance = getValueByHeaderCandidates_(r, ['성능점검']);
  const vat = getValueByHeaderCandidates_(r, ['부가세']);
  const discountRate = getValueByHeaderCandidates_(r, ['할인율(%)', '할인률(%)', '할인율', '할인률']);
  const quoteAmount = getValueByHeaderCandidates_(r, ['최종 견적가', '최종견적가', '최종단가', '최종 견적금액']);
  const status = getStatusValueFromObj_(r);
  const contractSummary = buildPortalSupportContractSummary_({
    contractUnit: contractUnit,
    appointment: appointment,
    maintenance: maintenance,
    performance: performance,
    vat: vat,
    discountRate: discountRate,
    quoteAmount: quoteAmount,
    status: status
  });

  return {
    rowNo: r.__rowNo || '',
    customerNo: r['고객번호'] || '',
    customerName: getCompanyValue_(r),
    discountRate: discountRate,
    quoteAmount: quoteAmount,
    area: getValueByHeaderCandidates_(r, ['연면적']),
    region: getValueByHeaderCandidates_(r, ['지역구분', '지역']),
    contactName: getValueByHeaderCandidates_(r, ['고객사 담당자', '담당자', '담당자명']),
    contactEmail: getValueByHeaderCandidates_(r, ['담당자 이메일 주소', '이메일주소', '이메일']),
    contractUnit: contractUnit,
    appointment: appointment,
    maintenance: maintenance,
    performance: performance,
    vat: vat,
    status: status,
    contractSummary: contractSummary,
    masterMemo: getMemoValueFromObj_(r)
  };
}

function buildPortalSupportCustomerInfoFromIndex_(r) {
  r = r || {};
  return {
    rowNo: r.rowNo || '',
    customerNo: r.customerNo || '',
    customerName: r.company || '',
    discountRate: '',
    quoteAmount: r.finalQuote || '',
    area: '',
    region: '',
    contactName: r.contact || '',
    contactEmail: r.email || '',
    contractUnit: '',
    appointment: '',
    maintenance: '',
    performance: '',
    vat: '',
    contractSummary: '',
    masterMemo: r.memo || '',
    status: r.status || ''
  };
}

function buildPortalSupportContractSummary_(info) {
  info = info || {};
  const parts = [];
  if (info.contractUnit) parts.push(String(info.contractUnit));
  const appointmentLabel = normalizePortalSupportAppointmentLabel_(info.appointment);
  if (appointmentLabel) parts.push('선임 ' + appointmentLabel);
  if (info.maintenance) parts.push('유지 ' + String(info.maintenance));
  if (info.performance) parts.push('성능 ' + String(info.performance));
  if (info.discountRate) parts.push('할인율 ' + String(info.discountRate).replace(/%$/, '') + '%');
  if (info.vat) parts.push('부가세 ' + String(info.vat));
  if (info.quoteAmount) parts.push('최종견적가 ' + String(info.quoteAmount));
  if (info.status) parts.push('진행상태 ' + String(info.status));
  return parts.join(' / ');
}

function normalizePortalSupportAppointmentLabel_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const lower = text.toLowerCase();
  if (text.indexOf('미') >= 0 || text.indexOf('없') >= 0 || lower.indexOf('x') >= 0 || lower.indexOf('n') === 0) return 'X';
  if (text.indexOf('해당없음') >= 0 || text.indexOf('해당 없음') >= 0) return 'X';
  if (text.indexOf('선임') >= 0 || lower.indexOf('o') >= 0 || lower.indexOf('y') === 0 || text.indexOf('유') >= 0) return 'O';
  return text;
}

function appendPortalSupportRequestToMasterMemo_(payload, receiptNo) {
  payload = payload || {};
  const customerNo = String(payload.customerNo || '').trim();
  const masterRowNo = Number(payload.masterRowNo) || 0;
  if (!customerNo && !masterRowNo) return { applied: false, reason: 'NO_CUSTOMER_KEY' };

  const masterSs = getMasterSpreadsheet_();
  const masterSheet = masterSs.getSheetByName(PORTAL_CONFIG.MASTER_SHEET_NAME);
  if (!masterSheet) throw new Error('마스터시트(신규)를 찾지 못했습니다.');

  let rowNo = masterRowNo;
  if (!rowNo || rowNo < PORTAL_CONFIG.DATA_START_ROW) {
    rowNo = findMasterRowNoByCustomerNo_(masterSheet, customerNo);
  }
  if (!rowNo || rowNo < PORTAL_CONFIG.DATA_START_ROW) return { applied: false, reason: 'MASTER_ROW_NOT_FOUND' };

  const nowText = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yy.MM.dd HH:mm');
  const requestType = String(payload.requestType || '').trim();
  const requester = String(payload.requester || '').trim();
  const requestText = String(payload.requestText || '').trim();
  const customerName = String(payload.customerName || '').trim();
  const logText = '[' + nowText + '] [영업지원요청 접수' + (receiptNo ? ' #' + receiptNo : '') + '] ' +
    [requestType, customerName, requester ? '요청자: ' + requester : '', requestText].filter(Boolean).join(' / ');

  const updatedMemo = appendToMasterMemo_(masterSheet, rowNo, logText);
  return { applied: true, rowNo: rowNo, updatedMemo: updatedMemo };
}

function findMasterRowNoByCustomerNo_(masterSheet, customerNo) {
  customerNo = String(customerNo || '').trim();
  if (!customerNo) return 0;
  const headerMap = getHeaderMap_(masterSheet);
  const col = findFirstExistingHeaderCol_(headerMap, ['고객번호']);
  if (!col) return 0;
  const lastRow = masterSheet.getLastRow();
  if (lastRow < PORTAL_CONFIG.DATA_START_ROW) return 0;
  const values = masterSheet.getRange(PORTAL_CONFIG.DATA_START_ROW, col, lastRow - PORTAL_CONFIG.DATA_START_ROW + 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === customerNo) return PORTAL_CONFIG.DATA_START_ROW + i;
  }
  return 0;
}
