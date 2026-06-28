/***************************************
 * S1 Sales Portal - 20_SystemHealthService.gs
 * P2-8: 동시사용/동기화 상태 점검 화면
 ***************************************/

const PORTAL_SYSTEM_HEALTH_REQUIRED_MASTER_HEADERS_P208 = [
  '고객번호', '현재 영업 진행 상황', '마스터시트 최종등록일', '지역구분', '수행사', '영업담당자', '회사명', '메모',
  '수정일시', '수정버전', '최종수정자'
];

const PORTAL_SYSTEM_HEALTH_REQUIRED_DB_SHEETS_P208 = [
  '검색인덱스_DB', '고객즐겨찾기_DB', '컨택이력_DB', '공지사항_DB', '작업로그_DB', '변경큐_DB', '오늘할일_DB', '권한_DB'
];

function getPortalSystemHealthP208() {
  const started = new Date();
  const userInfo = getPortalSystemHealthUserP208_();
  assertPortalSystemHealthAllowedP208_(userInfo);

  const masterSs = SpreadsheetApp.openById(PORTAL_CONFIG.MASTER_SPREADSHEET_ID);
  const webSs = getPortalSystemHealthWebDbSpreadsheetP208_();
  const masterSheet = masterSs.getSheetByName(PORTAL_CONFIG.MASTER_SHEET_NAME);

  const checks = [];
  const summary = {
    ok: 0,
    warn: 0,
    danger: 0,
    info: 0
  };

  function addCheck(item) {
    item = item || {};
    item.status = item.status || 'info';
    checks.push(item);
    if (summary[item.status] !== undefined) summary[item.status]++;
    else summary.info++;
  }

  addCheck({
    group: '기본 연결',
    name: '마스터 스프레드시트',
    status: masterSs ? 'ok' : 'danger',
    message: masterSs ? ('연결됨: ' + masterSs.getName()) : '마스터 스프레드시트를 열 수 없습니다.',
    detail: PORTAL_CONFIG.MASTER_SPREADSHEET_ID
  });

  addCheck({
    group: '기본 연결',
    name: '웹앱 DB 스프레드시트',
    status: webSs ? 'ok' : 'danger',
    message: webSs ? ('연결됨: ' + webSs.getName()) : '웹앱 DB 스프레드시트를 열 수 없습니다.',
    detail: webSs ? webSs.getId() : ''
  });

  if (!masterSheet) {
    addCheck({ group: '마스터시트', name: PORTAL_CONFIG.MASTER_SHEET_NAME, status: 'danger', message: '마스터시트를 찾지 못했습니다.' });
  } else {
    addCheck({
      group: '마스터시트',
      name: '데이터 행 수',
      status: 'ok',
      message: String(Math.max(0, masterSheet.getLastRow() - PORTAL_CONFIG.DATA_START_ROW + 1)) + '행',
      detail: 'lastRow=' + masterSheet.getLastRow() + ', lastCol=' + masterSheet.getLastColumn()
    });

    const headerMap = getPortalSystemHeaderMapP208_(masterSheet, PORTAL_CONFIG.HEADER_ROW);
    const missingMasterHeaders = PORTAL_SYSTEM_HEALTH_REQUIRED_MASTER_HEADERS_P208.filter(function(h) { return !headerMap[h]; });
    addCheck({
      group: '마스터시트',
      name: '필수 헤더',
      status: missingMasterHeaders.length ? 'danger' : 'ok',
      message: missingMasterHeaders.length ? ('누락: ' + missingMasterHeaders.join(', ')) : '필수 헤더 확인 완료',
      detail: missingMasterHeaders.join(', ')
    });

    const blankMeta = countBlankMasterMetaRowsP208_(masterSheet, headerMap);
    addCheck({
      group: '마스터시트',
      name: '빈 행 수정메타',
      status: blankMeta.count ? 'warn' : 'ok',
      message: blankMeta.count ? ('빈 행에 수정메타 ' + blankMeta.count + '건 남아 있음') : '빈 행 수정메타 없음',
      detail: blankMeta.sampleRows.length ? ('샘플 행: ' + blankMeta.sampleRows.join(', ')) : ''
    });
  }

  const triggerCheck = checkPortalMasterSyncTriggerP208_(masterSs.getId());
  addCheck({
    group: '트리거/동기화',
    name: '마스터 onEdit 트리거',
    status: triggerCheck.installed ? 'ok' : 'danger',
    message: triggerCheck.installed ? '설치됨' : '미설치: installMasterSheetEditSyncTriggerP201() 실행 필요',
    detail: triggerCheck.detail
  });

  const permissionTriggerCheck = checkPortalPermissionCacheTriggerP350_();
  addCheck({
    group: '트리거/동기화',
    name: '권한_DB 캐시 무효화 트리거',
    status: permissionTriggerCheck.installed ? 'ok' : 'warn',
    message: permissionTriggerCheck.installed ? '설치됨' : '미설치: 권한_DB 수정 후 권한 반영이 늦을 수 있음',
    detail: permissionTriggerCheck.detail || '설정 필요 시 시스템 점검에서 설치 버튼 실행'
  });

  const permissionCacheCheck = checkPortalPermissionCacheP350_();
  addCheck({
    group: '권한/캐시',
    name: '현재 계정 권한 캐시',
    status: permissionCacheCheck.error ? 'warn' : 'ok',
    message: permissionCacheCheck.error ? permissionCacheCheck.error : ('cacheHit=' + permissionCacheCheck.cacheHit + ', ttl=' + permissionCacheCheck.ttlSeconds + '초'),
    detail: permissionCacheCheck.detail
  });

  const version = getPortalSystemVersionInfoP208_();
  addCheck({
    group: '트리거/동기화',
    name: '마스터 변경 버전',
    status: version.masterVersion ? 'ok' : 'warn',
    message: version.masterVersion ? ('최근 변경: ' + (version.masterChangedAt || version.masterVersion)) : '아직 기록된 변경 버전 없음',
    detail: version.masterVersion
  });
  addCheck({
    group: '검색인덱스',
    name: '검색인덱스 상태',
    status: version.customerIndexDirty === 'Y' ? 'warn' : 'ok',
    message: version.customerIndexDirty === 'Y' ? ('재생성 필요: ' + (version.customerIndexDirtyReason || '사유 미기록')) : '정상',
    detail: 'version=' + (version.customerIndexVersion || '') + ', builtAt=' + (version.customerIndexBuiltAt || '')
  });

  PORTAL_SYSTEM_HEALTH_REQUIRED_DB_SHEETS_P208.forEach(function(name) {
    const sheet = webSs.getSheetByName(name);
    addCheck({
      group: '웹앱 DB',
      name: name,
      status: sheet ? 'ok' : (name === '작업로그_DB' ? 'warn' : 'danger'),
      message: sheet ? (Math.max(0, sheet.getLastRow() - 1) + '행') : '시트 없음',
      detail: sheet ? ('lastCol=' + sheet.getLastColumn()) : ''
    });
  });

  const favoriteCheck = checkPortalFavoriteDuplicatesP208_(webSs);
  addCheck({
    group: '즐겨찾기',
    name: '중복/삭제상태',
    status: favoriteCheck.duplicateActiveCount ? 'warn' : 'ok',
    message: favoriteCheck.duplicateActiveCount ? ('활성 중복 ' + favoriteCheck.duplicateActiveCount + '건') : '활성 중복 없음',
    detail: favoriteCheck.detail
  });

  const requestGuard = checkPortalRequestGuardPropsP208_();
  addCheck({
    group: '중복요청 가드',
    name: '처리중 요청',
    status: requestGuard.staleRunningCount ? 'warn' : 'ok',
    message: requestGuard.runningCount ? ('처리중 ' + requestGuard.runningCount + '건 / 오래된 처리중 ' + requestGuard.staleRunningCount + '건') : '처리중 요청 없음',
    detail: requestGuard.detail
  });

  const changeQueue = checkPortalChangeQueueP209_();
  addCheck({
    group: '마스터 변경큐',
    name: '검색인덱스 재처리 대기',
    status: changeQueue.error ? 'danger' : ((changeQueue.pending + changeQueue.errorCount) ? 'warn' : 'ok'),
    message: changeQueue.error ? changeQueue.error : ((changeQueue.pending + changeQueue.errorCount) ? ('대기 ' + changeQueue.pending + '건 / 오류 ' + changeQueue.errorCount + '건') : '미처리 변경건 없음'),
    detail: changeQueue.detail
  });

  const todayCheck = checkPortalTodayRowsP208_(webSs);
  addCheck({
    group: '오늘 할 일',
    name: '필수값 점검',
    status: todayCheck.badRows.length ? 'warn' : 'ok',
    message: todayCheck.badRows.length ? ('점검 필요 ' + todayCheck.badRows.length + '건') : '필수값 이상 없음',
    detail: todayCheck.badRows.slice(0, 10).join(', ')
  });

  const todayStats = getPortalTodaySheetStatsP350_(webSs);
  addCheck({
    group: '오늘 할 일',
    name: '데이터 누적/삭제 현황',
    status: todayStats.error ? 'warn' : (todayStats.deletedRows > 500 ? 'warn' : 'ok'),
    message: todayStats.error ? todayStats.error : ('전체 ' + todayStats.totalRows + '건 / 활성 ' + todayStats.activeRows + '건 / 삭제 ' + todayStats.deletedRows + '건 / 오늘 ' + todayStats.todayRows + '건'),
    detail: todayStats.detail
  });

  const supportStats = getPortalSupportSheetStatsP350_(masterSs);
  addCheck({
    group: '영업지원요청',
    name: '시트/캐시 상태',
    status: supportStats.error ? 'warn' : 'ok',
    message: supportStats.error ? supportStats.error : ('전체 ' + supportStats.totalRows + '건 / 미완료 ' + supportStats.openRows + '건 / cacheBust=' + supportStats.cacheBust),
    detail: supportStats.detail
  });

  const elapsedMs = new Date().getTime() - started.getTime();
  return {
    ok: true,
    checkedAt: getPortalSystemNowP208_(new Date()),
    elapsedMs: elapsedMs,
    user: userInfo,
    summary: summary,
    checks: checks
  };
}

function repairPortalBlankMasterMetaP208() {
  const userInfo = getPortalSystemHealthUserP208_();
  assertPortalSystemHealthAllowedP208_(userInfo);
  if (typeof clearBlankMasterMetaRowsP204 === 'function') {
    return clearBlankMasterMetaRowsP204();
  }
  throw new Error('빈 행 수정메타 정리 함수가 없습니다. 18_ConcurrencySyncService.gs를 확인해 주세요.');
}

function installPortalMasterSyncTriggerFromHealthP208() {
  const userInfo = getPortalSystemHealthUserP208_();
  assertPortalSystemHealthAllowedP208_(userInfo);
  if (typeof installMasterSheetEditSyncTriggerP201 === 'function') {
    return installMasterSheetEditSyncTriggerP201();
  }
  throw new Error('마스터시트 수정 감지 트리거 설치 함수가 없습니다.');
}

function processPortalChangeQueueFromHealthP209() {
  const userInfo = getPortalSystemHealthUserP208_();
  assertPortalSystemHealthAllowedP208_(userInfo);
  if (typeof processPortalChangeQueueP209 === 'function') {
    return processPortalChangeQueueP209({ limit: 80, includeErrors: true });
  }
  throw new Error('변경큐 재처리 함수가 없습니다. 21_ChangeQueueService.gs를 확인해 주세요.');
}

function installPortalPermissionCacheBusterFromHealthP350() {
  const userInfo = getPortalSystemHealthUserP208_();
  assertPortalSystemHealthAllowedP208_(userInfo);
  if (typeof installPortalPermissionCacheBusterTriggerP230 === 'function') {
    return installPortalPermissionCacheBusterTriggerP230();
  }
  throw new Error('권한 캐시 무효화 트리거 설치 함수가 없습니다. 15_PermissionService.gs를 확인해 주세요.');
}

function clearPortalPermissionCacheFromHealthP350() {
  const userInfo = getPortalSystemHealthUserP208_();
  assertPortalSystemHealthAllowedP208_(userInfo);
  if (typeof clearPortalPermissionCache === 'function') {
    return clearPortalPermissionCache();
  }
  throw new Error('권한 캐시 무효화 함수가 없습니다. 15_PermissionService.gs를 확인해 주세요.');
}

function getPortalSystemHealthUserP208_() {
  let email = '';
  try { email = String(Session.getActiveUser().getEmail() || '').trim(); } catch (err) {}
  let permission = null;
  try {
    const p = getPortalCurrentPermission();
    permission = p && p.permission ? p.permission : p;
  } catch (err) {}
  const level = permission && permission.level ? String(permission.level) : '';
  const displayName = permission && permission.displayName ? String(permission.displayName) : '';
  const allowed = !!(permission && permission.active !== false && (permission.canUseAdminHome || permission.canCompareActivityLogs || level === 'ADMIN'));
  return { email: email, level: level || 'UNKNOWN', displayName: displayName || email || 'unknown', allowed: allowed };
}

function assertPortalSystemHealthAllowedP208_(userInfo) {
  if (userInfo && userInfo.allowed) return;
  throw new Error('시스템 점검 화면은 관리자/서무 권한에서만 사용할 수 있습니다.');
}

function getPortalSystemHealthWebDbSpreadsheetP208_() {
  if (PORTAL_CONFIG.WEBAPP_DB_SPREADSHEET_ID) {
    return SpreadsheetApp.openById(PORTAL_CONFIG.WEBAPP_DB_SPREADSHEET_ID);
  }
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (err) {}
  return SpreadsheetApp.openById(PORTAL_CONFIG.MASTER_SPREADSHEET_ID);
}

function getPortalSystemHeaderMapP208_(sheet, headerRow) {
  const lastCol = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(headerRow || 1, 1, 1, lastCol).getDisplayValues()[0];
  const map = {};
  headers.forEach(function(h, i) {
    h = String(h || '').trim();
    if (h && !map[h]) map[h] = i + 1;
  });
  return map;
}

function findPortalSystemHeaderColP208_(headerMap, candidates) {
  candidates = Array.isArray(candidates) ? candidates : [candidates];
  for (let i = 0; i < candidates.length; i++) {
    const key = String(candidates[i] || '').trim();
    if (key && headerMap[key]) return headerMap[key];
  }
  return 0;
}

function getPortalSystemNowP208_(date) {
  return Utilities.formatDate(date || new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function countBlankMasterMetaRowsP208_(sheet, headerMap) {
  const metaCols = ['수정일시', '수정버전', '최종수정자'].map(function(h) { return findPortalSystemHeaderColP208_(headerMap, h); }).filter(Boolean);
  if (!metaCols.length) return { count: 0, sampleRows: [] };
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < PORTAL_CONFIG.DATA_START_ROW) return { count: 0, sampleRows: [] };
  const values = sheet.getRange(PORTAL_CONFIG.DATA_START_ROW, 1, lastRow - PORTAL_CONFIG.DATA_START_ROW + 1, lastCol).getDisplayValues();
  let count = 0;
  const sampleRows = [];
  const importantHeaders = ['고객번호', '회사명', '건물명', '현재 영업 진행 상황', '영업담당자', '견적담당', '고객사 담당자', '대표전화', '직통번호', '이메일', '주소', '메모', '최종 견적가'];
  const importantCols = importantHeaders.map(function(h) { return findPortalSystemHeaderColP208_(headerMap, h); }).filter(Boolean);

  values.forEach(function(row, idx) {
    const rowNo = PORTAL_CONFIG.DATA_START_ROW + idx;
    const hasImportant = importantCols.some(function(c) { return String(row[c - 1] || '').trim(); });
    if (hasImportant) return;
    const hasMeta = metaCols.some(function(c) { return String(row[c - 1] || '').trim(); });
    if (!hasMeta) return;
    count++;
    if (sampleRows.length < 20) sampleRows.push(rowNo);
  });
  return { count: count, sampleRows: sampleRows };
}

function checkPortalMasterSyncTriggerP208_(spreadsheetId) {
  let installed = false;
  const detail = [];
  try {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      const handler = t.getHandlerFunction ? t.getHandlerFunction() : '';
      const sourceId = t.getTriggerSourceId ? t.getTriggerSourceId() : '';
      if (handler === 'onMasterSheetEditSyncP201') {
        installed = true;
        detail.push(handler + ' / ' + sourceId);
      }
    });
  } catch (err) {
    detail.push('트리거 조회 실패: ' + (err && err.message ? err.message : String(err)));
  }
  return { installed: installed, detail: detail.join('\n') };
}

function getPortalSystemVersionInfoP208_() {
  const props = PropertiesService.getScriptProperties();
  return {
    masterVersion: props.getProperty('PORTAL_MASTER_DATA_VERSION') || '',
    masterChangedAt: props.getProperty('PORTAL_MASTER_DATA_CHANGED_AT') || '',
    customerIndexVersion: props.getProperty('CUSTOMER_SEARCH_INDEX_VERSION') || '',
    customerIndexBuiltAt: props.getProperty('CUSTOMER_SEARCH_INDEX_BUILT_AT') || '',
    customerIndexDirty: props.getProperty('CUSTOMER_SEARCH_INDEX_DIRTY') || 'N',
    customerIndexDirtyReason: props.getProperty('CUSTOMER_SEARCH_INDEX_DIRTY_REASON') || ''
  };
}

function checkPortalPermissionCacheTriggerP350_() {
  const handler = (typeof PORTAL_PERMISSION_CACHE_BUSTER_HANDLER_P230 !== 'undefined')
    ? PORTAL_PERMISSION_CACHE_BUSTER_HANDLER_P230
    : 'onPortalPermissionSheetEditP230';
  let installed = false;
  const detail = [];
  try {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      const h = t && t.getHandlerFunction ? t.getHandlerFunction() : '';
      const sourceId = t && t.getTriggerSourceId ? t.getTriggerSourceId() : '';
      if (h === handler) {
        installed = true;
        detail.push(h + ' / ' + sourceId);
      }
    });
  } catch (err) {
    detail.push('트리거 조회 실패: ' + (err && err.message ? err.message : String(err)));
  }
  return { installed: installed, detail: detail.join('\n') };
}

function checkPortalPermissionCacheP350_() {
  try {
    if (typeof getPortalPermissionCacheStatus !== 'function') {
      return { error: '권한 캐시 상태 함수 없음', detail: '', cacheHit: false, ttlSeconds: 0 };
    }
    const res = getPortalPermissionCacheStatus();
    return {
      error: '',
      cacheHit: res && res.cacheHit ? 'Y' : 'N',
      ttlSeconds: Number(res && res.ttlSeconds || 0) || 0,
      detail: 'email=' + ((res && res.emailKey) || '') + '\ncacheBust=' + ((res && res.cacheBust) || '') + '\nguestTtl=' + ((res && res.guestTtlSeconds) || 0)
    };
  } catch (err) {
    return { error: err && err.message ? err.message : String(err), detail: '', cacheHit: 'N', ttlSeconds: 0 };
  }
}

function getPortalTodaySheetStatsP350_(webSs) {
  try {
    const sheet = webSs.getSheetByName(PORTAL_CONFIG.TODAY_SHEET_NAME || '오늘할일_DB');
    if (!sheet || sheet.getLastRow() < 2) return { totalRows: 0, activeRows: 0, deletedRows: 0, todayRows: 0, detail: '오늘할일_DB 데이터 없음' };
    const headerMap = getPortalSystemHeaderMapP208_(sheet, 1);
    const dateCol = findPortalSystemHeaderColP208_(headerMap, '일자');
    const doneCol = findPortalSystemHeaderColP208_(headerMap, '완료여부');
    const ownerCol = findPortalSystemHeaderColP208_(headerMap, ['담당자', '작성자']);
    const deletedCol = findPortalSystemHeaderColP208_(headerMap, '삭제여부');
    const lastRow = sheet.getLastRow();
    const width = Math.max(sheet.getLastColumn(), deletedCol || 1, doneCol || 1, dateCol || 1, ownerCol || 1);
    const values = sheet.getRange(2, 1, lastRow - 1, width).getDisplayValues();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    let active = 0, deleted = 0, todayRows = 0, done = 0;
    const owners = {};
    values.forEach(function(r) {
      const isDeleted = deletedCol ? String(r[deletedCol - 1] || '').trim().toUpperCase() === 'Y' : false;
      if (isDeleted) { deleted++; return; }
      active++;
      if (doneCol && String(r[doneCol - 1] || '').trim().toUpperCase() === 'Y') done++;
      const dateText = dateCol ? normalizePortalSystemDateKeyP350_(r[dateCol - 1]) : '';
      if (dateText === today) todayRows++;
      const owner = ownerCol ? String(r[ownerCol - 1] || '').trim() : '';
      if (owner) owners[owner] = (owners[owner] || 0) + 1;
    });
    const topOwners = Object.keys(owners).sort(function(a,b){ return owners[b]-owners[a]; }).slice(0, 10).map(function(k){ return k + ' ' + owners[k] + '건'; });
    return {
      totalRows: values.length,
      activeRows: active,
      deletedRows: deleted,
      todayRows: todayRows,
      detail: '완료 ' + done + '건 / 미완료 ' + Math.max(0, active - done) + '건' + (topOwners.length ? '\n담당자별 상위: ' + topOwners.join(', ') : '')
    };
  } catch (err) {
    return { error: err && err.message ? err.message : String(err), totalRows: 0, activeRows: 0, deletedRows: 0, todayRows: 0, detail: '' };
  }
}

function getPortalSupportSheetStatsP350_(masterSs) {
  try {
    const sheet = masterSs.getSheetByName(PORTAL_CONFIG.SUPPORT_SHEET_NAME || '영업지원요청');
    if (!sheet) return { error: '영업지원요청 시트 없음', totalRows: 0, openRows: 0, cacheBust: '', detail: '' };
    const dataStart = Number(PORTAL_CONFIG.SUPPORT_DATA_START_ROW || 4);
    const lastRow = sheet.getLastRow();
    if (lastRow < dataStart) return { totalRows: 0, openRows: 0, cacheBust: getPortalSupportCacheBustP350_(), detail: '데이터 없음' };
    const headerMap = getPortalSystemHeaderMapP208_(sheet, Number(PORTAL_CONFIG.SUPPORT_HEADER_ROW || 3));
    const statusCol = findPortalSystemHeaderColP208_(headerMap, ['처리상태', '상태']);
    const requesterCol = findPortalSystemHeaderColP208_(headerMap, ['요청자', '요청담당자']);
    const width = Math.max(sheet.getLastColumn(), statusCol || 1, requesterCol || 1);
    const values = sheet.getRange(dataStart, 1, lastRow - dataStart + 1, width).getDisplayValues();
    let open = 0;
    const requesters = {};
    values.forEach(function(r) {
      const st = statusCol ? String(r[statusCol - 1] || '').trim() : '';
      if (!st || !/완료|반려/.test(st)) open++;
      const req = requesterCol ? String(r[requesterCol - 1] || '').trim() : '';
      if (req) requesters[req] = (requesters[req] || 0) + 1;
    });
    const top = Object.keys(requesters).sort(function(a,b){ return requesters[b]-requesters[a]; }).slice(0, 10).map(function(k){ return k + ' ' + requesters[k] + '건'; });
    return { totalRows: values.length, openRows: open, cacheBust: getPortalSupportCacheBustP350_(), detail: top.length ? ('요청자별 상위: ' + top.join(', ')) : '' };
  } catch (err) {
    return { error: err && err.message ? err.message : String(err), totalRows: 0, openRows: 0, cacheBust: '', detail: '' };
  }
}

function getPortalSupportCacheBustP350_() {
  try {
    if (typeof getPortalSupportCacheBustV64_ === 'function') return getPortalSupportCacheBustV64_();
  } catch (err) {}
  try { return PropertiesService.getScriptProperties().getProperty('PORTAL_SUPPORT_CACHE_BUST_V64') || ''; } catch (err2) {}
  return '';
}

function normalizePortalSystemDateKeyP350_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const text = String(value == null ? '' : value).trim();
  const m = text.match(/(20\d{2}|\d{2})[.\-\/]\s*(\d{1,2})[.\-\/]\s*(\d{1,2})/);
  if (!m) return text;
  let y = Number(m[1]);
  if (y < 100) y += 2000;
  return y + '-' + ('0' + Number(m[2])).slice(-2) + '-' + ('0' + Number(m[3])).slice(-2);
}

function checkPortalFavoriteDuplicatesP208_(webSs) {
  const sheet = webSs.getSheetByName('고객즐겨찾기_DB');
  if (!sheet || sheet.getLastRow() < 2) return { duplicateActiveCount: 0, detail: '즐겨찾기 행 없음' };
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(9, sheet.getLastColumn())).getDisplayValues();
  const active = {};
  const duplicates = [];
  values.forEach(function(r, i) {
    const deleted = String(r[7] || '').trim().toUpperCase() === 'Y';
    if (deleted) return;
    const key = [String(r[1] || '').trim(), String(r[3] || '').trim(), String(r[4] || '').trim()].join('|');
    if (!key.replace(/\|/g, '')) return;
    if (active[key]) duplicates.push(i + 2);
    else active[key] = i + 2;
  });
  return { duplicateActiveCount: duplicates.length, detail: duplicates.length ? ('중복행: ' + duplicates.slice(0, 20).join(', ')) : '정상' };
}

function checkPortalRequestGuardPropsP208_() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const now = Date.now();
  let running = 0;
  let stale = 0;
  const sample = [];
  Object.keys(props).forEach(function(k) {
    if (k.indexOf('PORTAL_REQ_GUARD__') !== 0) return;
    try {
      const state = JSON.parse(props[k]);
      if (state.status === 'RUNNING') {
        running++;
        const ageMin = Math.round((now - Number(state.ts || 0)) / 60000);
        if (ageMin >= 15) stale++;
        if (sample.length < 10) sample.push(state.type + '/' + state.requestId + '/' + ageMin + '분');
      }
    } catch (err) {}
  });
  return { runningCount: running, staleRunningCount: stale, detail: sample.join('\n') };
}


function checkPortalChangeQueueP209_() {
  if (typeof getPortalChangeQueueStatsP209 !== 'function') {
    return { pending: 0, errorCount: 0, detail: '변경큐 서비스 파일 없음', error: '변경큐 서비스 파일이 없습니다.' };
  }
  try {
    const res = getPortalChangeQueueStatsP209();
    const stats = (res && res.stats) || {};
    const samples = Array.isArray(stats.samples) ? stats.samples : [];
    return {
      pending: Number(stats.pending || 0) || 0,
      errorCount: Number(stats.error || 0) || 0,
      detail: samples.length ? samples.map(function(s) {
        return '마스터행 ' + (s.masterRow || '-') + ' / 고객번호 ' + (s.customerNo || '-') + ' / ' + (s.status || '') + (s.error ? ' / ' + s.error : '');
      }).join('\n') : ('전체 ' + (stats.total || 0) + '건, 완료 ' + (stats.done || 0) + '건, 건너뜀 ' + (stats.skipped || 0) + '건')
    };
  } catch (err) {
    return { pending: 0, errorCount: 0, detail: '', error: err && err.message ? err.message : String(err) };
  }
}

function checkPortalTodayRowsP208_(webSs) {
  const sheet = webSs.getSheetByName(PORTAL_CONFIG.TODAY_SHEET_NAME || '오늘할일_DB');
  if (!sheet || sheet.getLastRow() < 2) return { badRows: [] };
  const headerMap = getPortalSystemHeaderMapP208_(sheet, 1);
  const idCol = findPortalSystemHeaderColP208_(headerMap, '할일ID');
  const dateCol = findPortalSystemHeaderColP208_(headerMap, '일자');
  const contentCol = findPortalSystemHeaderColP208_(headerMap, ['내용', '할 일']);
  const deletedCol = findPortalSystemHeaderColP208_(headerMap, '삭제여부');
  const width = sheet.getLastColumn();
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getDisplayValues();
  const bad = [];
  values.forEach(function(r, i) {
    const rowNo = i + 2;
    const deleted = deletedCol ? String(r[deletedCol - 1] || '').trim().toUpperCase() === 'Y' : false;
    if (deleted) return;
    const id = idCol ? String(r[idCol - 1] || '').trim() : '';
    const date = dateCol ? String(r[dateCol - 1] || '').trim() : '';
    const content = contentCol ? String(r[contentCol - 1] || '').trim() : '';
    if (!id || !date || !content) bad.push(rowNo);
  });
  return { badRows: bad };
}
