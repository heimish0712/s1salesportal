/***************************************
 * S1 Sales Portal - 20_SystemHealthService.gs
 * P2-8: 동시사용/동기화 상태 점검 화면
 ***************************************/

const PORTAL_SYSTEM_HEALTH_REQUIRED_MASTER_HEADERS_P208 = [
  '고객번호', '현재 영업 진행 상황', '마스터시트 최종등록일', '지역구분', '수행사', '영업담당자', '회사명', '메모',
  '수정일시', '수정버전', '최종수정자'
];

const PORTAL_SYSTEM_HEALTH_REQUIRED_DB_SHEETS_P208 = [
  '검색인덱스_DB',
  '고객즐겨찾기_DB',
  '컨택이력_DB',
  '공지사항_DB',
  '작업로그_DB',
  '변경큐_DB',
  '오늘할일_DB',
  '권한_DB',
  '저장큐_DB',
  '성능로그_DB',
  '담당자프로필_DB',
  '고객분류_폴더_DB',
  '고객분류_고객_DB'
];

// P517: 신규 운영진단 대상 시트 중, 아직 생성 전일 수 있는 시트는 누락 시 주의로 표시합니다.
const PORTAL_SYSTEM_HEALTH_OPTIONAL_DB_SHEETS_P517 = [
  '작업로그_DB',
  '저장큐_DB',
  '성능로그_DB',
  '담당자프로필_DB',
  '고객분류_폴더_DB',
  '고객분류_고객_DB'
];

const PORTAL_SYSTEM_HEALTH_SAVE_QUEUE_FALLBACK_HEADERS_P518 = [
  '등록일시',
  '수정일시',
  '작업ID',
  '사용자',
  '세션ID',
  '고객번호',
  'rowNo',
  'methodName',
  'source',
  '상태',
  '우선순위',
  '시도횟수',
  'patchJson',
  'expectedValuesJson',
  'payloadJson',
  'resultJson',
  '마지막오류',
  '적용일시'
];

const PORTAL_SYSTEM_HEALTH_SAVE_QUEUE_CORE_HEADERS_P518 = [
  '작업ID',
  '고객번호',
  'rowNo',
  'source',
  '상태',
  'patchJson',
  'expectedValuesJson',
  'resultJson',
  '마지막오류'
];

const PORTAL_SYSTEM_HEALTH_SAVE_FLOW_FUNCTIONS_P518 = [
  { name: 'saveCustomerPatchFastP473', label: '직접 저장 진입점', required: true },
  { name: 'processSaveQueueP473', label: '저장큐 처리기', required: true },
  { name: 'flushCustomerPendingOpsP473', label: '자료발송 전 pending flush', required: true }
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

  const releaseInfoP517 = getPortalReleaseInfoP517_();
  addCheck({
    group: '릴리즈',
    name: '기준본',
    status: releaseInfoP517.exists ? (releaseInfoP517.businessLogicChanged ? 'warn' : 'ok') : 'warn',
    message: releaseInfoP517.message,
    detail: releaseInfoP517.detail
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
      status: sheet ? 'ok' : getPortalSystemHealthMissingDbStatusP517_(name),
      message: sheet ? (Math.max(0, sheet.getLastRow() - 1) + '행') : '시트 없음',
      detail: sheet ? ('lastCol=' + sheet.getLastColumn()) : '신규/선택 기능 시트는 사용 전 생성되지 않았을 수 있습니다.'
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

  const saveQueueCheck = checkPortalSaveQueueP517_(webSs);
  addCheck({
    group: '저장 안정성',
    name: '저장큐_DB 상태',
    status: saveQueueCheck.status,
    message: saveQueueCheck.message,
    detail: saveQueueCheck.detail
  });

  const saveFlowFunctionCheck = checkPortalSaveFlowFunctionsP518_();
  addCheck({
    group: '저장 안정성',
    name: '저장 핵심 함수',
    status: saveFlowFunctionCheck.status,
    message: saveFlowFunctionCheck.message,
    detail: saveFlowFunctionCheck.detail
  });

  const saveQueueSchemaCheck = checkPortalSaveQueueSchemaP518_(webSs);
  addCheck({
    group: '저장 안정성',
    name: '저장큐_DB 스키마',
    status: saveQueueSchemaCheck.status,
    message: saveQueueSchemaCheck.message,
    detail: saveQueueSchemaCheck.detail
  });

  const recentSavePerfCheck = checkPortalRecentSavePerformanceP518_(webSs);
  addCheck({
    group: '저장 안정성',
    name: '최근 저장 오류/지연',
    status: recentSavePerfCheck.status,
    message: recentSavePerfCheck.message,
    detail: recentSavePerfCheck.detail
  });

  const perfLogCheck = checkPortalPerformanceLogP517_(webSs);
  addCheck({
    group: '성능로그',
    name: '최근 오류/지연',
    status: perfLogCheck.status,
    message: perfLogCheck.message,
    detail: perfLogCheck.detail
  });

  const folderDbCheck = checkPortalCustomerFolderDbP517_(webSs);
  addCheck({
    group: '나의 고객 폴더',
    name: '고객분류 DB',
    status: folderDbCheck.status,
    message: folderDbCheck.message,
    detail: folderDbCheck.detail
  });

  const contactProfileCheck = checkPortalContactProfileDbP517_(webSs);
  addCheck({
    group: '담당자 프로필',
    name: '담당자프로필_DB',
    status: contactProfileCheck.status,
    message: contactProfileCheck.message,
    detail: contactProfileCheck.detail
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


/***************************************
 * P517: 운영 안정화 스프린트용 시스템 점검 보강
 * - 저장/자료발송/검색 로직에는 개입하지 않고, 최근 상태를 읽기만 합니다.
 ***************************************/
function getPortalReleaseInfoP517_() {
  try {
    if (typeof PORTAL_RELEASE_INFO === 'undefined' || !PORTAL_RELEASE_INFO) {
      return { exists: false, businessLogicChanged: false, message: '릴리즈 기준본 정보 없음', detail: '00_Config.gs의 PORTAL_RELEASE_INFO 확인 필요' };
    }
    const info = PORTAL_RELEASE_INFO;
    const version = String(info.version || '').trim() || '-';
    const releaseName = String(info.releaseName || '').trim() || '-';
    const baselineZip = String(info.baselineZip || '').trim() || '-';
    const baselineTimestamp = String(info.baselineTimestamp || '').trim() || '-';
    const businessLogicChanged = info.businessLogicChanged === true;
    const milestones = Array.isArray(info.includedMilestones) ? info.includedMilestones.join(', ') : '';
    const notes = Array.isArray(info.notes) ? info.notes.join('\n') : '';
    return {
      exists: true,
      businessLogicChanged: businessLogicChanged,
      message: version + ' / ' + baselineZip + ' / 업무로직변경=' + (businessLogicChanged ? 'Y' : 'N'),
      detail: 'releaseName=' + releaseName + '\nbaselineTimestamp=' + baselineTimestamp + (milestones ? '\nincluded=' + milestones : '') + (notes ? '\nnotes=' + notes : '')
    };
  } catch (err) {
    return { exists: false, businessLogicChanged: false, message: '릴리즈 기준본 정보 확인 실패', detail: getPortalSystemHealthErrorTextP517_(err) };
  }
}

function getPortalSystemHealthMissingDbStatusP517_(name) {
  name = String(name || '').trim();
  return PORTAL_SYSTEM_HEALTH_OPTIONAL_DB_SHEETS_P517.indexOf(name) >= 0 ? 'warn' : 'danger';
}

function getPortalSaveQueueSheetNameP517_() {
  try {
    if (typeof PORTAL_SAVE_QUEUE_P473 !== 'undefined' && PORTAL_SAVE_QUEUE_P473 && PORTAL_SAVE_QUEUE_P473.SHEET_NAME) return PORTAL_SAVE_QUEUE_P473.SHEET_NAME;
  } catch (err) {}
  return '저장큐_DB';
}

function getPortalPerfLogSheetNameP517_() {
  try {
    if (typeof getPortalPerfLogSheetNameP460_ === 'function') return getPortalPerfLogSheetNameP460_();
  } catch (err) {}
  try {
    if (PORTAL_CONFIG && PORTAL_CONFIG.PERF_LOG_SHEET_NAME) return PORTAL_CONFIG.PERF_LOG_SHEET_NAME;
  } catch (err2) {}
  return '성능로그_DB';
}

function getPortalContactProfileSheetNameP517_() {
  try {
    if (PORTAL_CONFIG && PORTAL_CONFIG.CONTACT_PROFILE_SHEET_NAME) return PORTAL_CONFIG.CONTACT_PROFILE_SHEET_NAME;
  } catch (err) {}
  return '담당자프로필_DB';
}

function getPortalCustomerFolderSheetNamesP517_() {
  let folderSheetName = '고객분류_폴더_DB';
  let itemSheetName = '고객분류_고객_DB';
  try {
    if (typeof PORTAL_MY_CUSTOMER_FOLDER_SHEET_P497 !== 'undefined' && PORTAL_MY_CUSTOMER_FOLDER_SHEET_P497) folderSheetName = PORTAL_MY_CUSTOMER_FOLDER_SHEET_P497;
  } catch (err) {}
  try {
    if (typeof PORTAL_MY_CUSTOMER_FOLDER_ITEM_SHEET_P497 !== 'undefined' && PORTAL_MY_CUSTOMER_FOLDER_ITEM_SHEET_P497) itemSheetName = PORTAL_MY_CUSTOMER_FOLDER_ITEM_SHEET_P497;
  } catch (err2) {}
  return { folderSheetName: folderSheetName, itemSheetName: itemSheetName };
}

function getPortalSystemRecentRowsP517_(sheet, limit, minWidth) {
  const lastRow = sheet ? sheet.getLastRow() : 0;
  const width = Math.max(sheet ? sheet.getLastColumn() : 1, Number(minWidth || 1) || 1);
  const headers = sheet ? sheet.getRange(1, 1, 1, width).getDisplayValues()[0] : [];
  if (!sheet || lastRow < 2) return { headers: headers, rows: [], startRow: 2, lastRow: lastRow, totalRows: 0, width: width };
  const totalRows = lastRow - 1;
  const count = Math.min(Number(limit || 100) || 100, totalRows);
  const startRow = Math.max(2, lastRow - count + 1);
  const rows = sheet.getRange(startRow, 1, count, width).getDisplayValues();
  return { headers: headers, rows: rows, startRow: startRow, lastRow: lastRow, totalRows: totalRows, width: width };
}

function getPortalSystemHeaderMapFromHeadersP517_(headers) {
  const map = {};
  (Array.isArray(headers) ? headers : []).forEach(function(h, i) {
    h = String(h || '').trim();
    if (h && !map[h]) map[h] = i + 1;
  });
  return map;
}

function getPortalSystemColFromHeadersP517_(headers, candidates) {
  const map = getPortalSystemHeaderMapFromHeadersP517_(headers);
  return findPortalSystemHeaderColP208_(map, candidates);
}

function getPortalSystemRequiredHeaderStatusP517_(sheet, requiredHeaders) {
  requiredHeaders = Array.isArray(requiredHeaders) ? requiredHeaders : [];
  if (!sheet) return { ok: false, missing: requiredHeaders.slice(0), detail: '시트 없음' };
  if (!requiredHeaders.length) return { ok: true, missing: [], detail: '필수 헤더 정의 없음' };
  const width = Math.max(sheet.getLastColumn() || 1, requiredHeaders.length);
  const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const existing = {};
  headers.forEach(function(h) { if (h) existing[h] = true; });
  const missing = requiredHeaders.filter(function(h) { return !existing[h]; });
  return { ok: missing.length === 0, missing: missing, detail: missing.length ? ('누락: ' + missing.join(', ')) : '필수 헤더 확인 완료' };
}

function checkPortalSaveQueueP517_(webSs) {
  try {
    const sheetName = getPortalSaveQueueSheetNameP517_();
    const sheet = webSs.getSheetByName(sheetName);
    if (!sheet) {
      return { status: 'warn', message: sheetName + ' 없음: fallback 저장이 아직 발생하지 않았을 수 있음', detail: '정상 저장은 마스터시트에 직접 반영되므로, 큐 시트는 예외 발생 전까지 없을 수 있습니다.' };
    }
    const recent = getPortalSystemRecentRowsP517_(sheet, 500, 18);
    const statusCol = getPortalSystemColFromHeadersP517_(recent.headers, '상태');
    const jobCol = getPortalSystemColFromHeadersP517_(recent.headers, '작업ID');
    const customerCol = getPortalSystemColFromHeadersP517_(recent.headers, '고객번호');
    const rowNoCol = getPortalSystemColFromHeadersP517_(recent.headers, 'rowNo');
    const updatedCol = getPortalSystemColFromHeadersP517_(recent.headers, ['수정일시', 'updatedAt']);
    const createdCol = getPortalSystemColFromHeadersP517_(recent.headers, ['등록일시', 'createdAt']);
    const errorCol = getPortalSystemColFromHeadersP517_(recent.headers, '마지막오류');
    if (!statusCol) return { status: 'danger', message: sheetName + ' 상태 헤더 없음', detail: '저장큐_DB 헤더를 확인해 주세요.' };
    const counts = { QUEUED: 0, RUNNING: 0, DONE: 0, RETRY: 0, CONFLICT: 0, FAIL: 0, BLANK: 0, OTHER: 0, STALE_RUNNING: 0 };
    const samples = [];
    const staleRunningMinutes = getPortalSaveQueueStaleRunningMinutesP518_();
    const nowMs = new Date().getTime();
    recent.rows.forEach(function(row, i) {
      const status = String(row[statusCol - 1] || '').trim().toUpperCase();
      const key = Object.prototype.hasOwnProperty.call(counts, status) ? status : (status ? 'OTHER' : 'BLANK');
      counts[key]++;
      let ageMin = 0;
      if (status === 'RUNNING') {
        const baseValue = updatedCol ? row[updatedCol - 1] : (createdCol ? row[createdCol - 1] : '');
        const ts = getPortalSystemDateMsP518_(baseValue);
        if (ts) {
          ageMin = Math.max(0, Math.floor((nowMs - ts) / 60000));
          if (ageMin >= staleRunningMinutes) counts.STALE_RUNNING++;
        }
      }
      if ((status === 'QUEUED' || status === 'RUNNING' || status === 'RETRY' || status === 'CONFLICT' || status === 'FAIL' || !status) && samples.length < 10) {
        const rowNo = recent.startRow + i;
        const jobId = jobCol ? String(row[jobCol - 1] || '').trim() : '';
        const customerNo = customerCol ? String(row[customerCol - 1] || '').trim() : '';
        const masterRow = rowNoCol ? String(row[rowNoCol - 1] || '').trim() : '';
        const err = errorCol ? String(row[errorCol - 1] || '').trim() : '';
        samples.push('큐행 ' + rowNo + ' / ' + (status || '상태공란') + (ageMin ? ' / ' + ageMin + '분 경과' : '') + (customerNo ? ' / 고객번호 ' + customerNo : '') + (masterRow ? ' / 마스터행 ' + masterRow : '') + (jobId ? ' / ' + jobId : '') + (err ? ' / ' + err : ''));
      }
    });
    const risky = counts.FAIL + counts.CONFLICT;
    const waiting = counts.QUEUED + counts.RUNNING + counts.RETRY;
    const blank = counts.BLANK;
    const stale = counts.STALE_RUNNING;
    const status = risky ? 'danger' : ((stale || waiting || blank) ? 'warn' : 'ok');
    const parts = [];
    ['QUEUED','RUNNING','STALE_RUNNING','RETRY','CONFLICT','FAIL','DONE','BLANK','OTHER'].forEach(function(k) { if (counts[k]) parts.push(k + ' ' + counts[k]); });
    const message = recent.totalRows ? ('전체 ' + recent.totalRows + '건 / 최근 ' + recent.rows.length + '행 기준 · ' + (parts.length ? parts.join(' / ') : '대기 저장 없음')) : '대기 저장 없음';
    return { status: status, message: message, detail: samples.length ? samples.join('\n') : '최근 대기/충돌/실패 저장 없음' };
  } catch (err) {
    return { status: 'warn', message: '저장큐 상태 확인 실패', detail: getPortalSystemHealthErrorTextP517_(err) };
  }
}

function checkPortalPerformanceLogP517_(webSs) {
  try {
    const sheetName = getPortalPerfLogSheetNameP517_();
    const sheet = webSs.getSheetByName(sheetName);
    if (!sheet) return { status: 'warn', message: sheetName + ' 없음: 아직 시트 로그가 저장되지 않았을 수 있음', detail: '브라우저 로그 저장 버튼 또는 자동 로그 적재 후 생성됩니다.' };
    const recent = getPortalSystemRecentRowsP517_(sheet, 200, 12);
    const statusCol = getPortalSystemColFromHeadersP517_(recent.headers, '상태');
    const errorCol = getPortalSystemColFromHeadersP517_(recent.headers, ['오류', '에러']);
    const durationCol = getPortalSystemColFromHeadersP517_(recent.headers, ['소요ms', 'durationMs']);
    const eventCol = getPortalSystemColFromHeadersP517_(recent.headers, ['이벤트', 'event']);
    const phaseCol = getPortalSystemColFromHeadersP517_(recent.headers, ['구간', 'phase']);
    const screenCol = getPortalSystemColFromHeadersP517_(recent.headers, ['화면', 'screen']);
    let errorCount = 0;
    let slowCount = 0;
    let blankSourceCount = 0;
    const samples = [];
    recent.rows.forEach(function(row, i) {
      const rowNo = recent.startRow + i;
      const statusText = statusCol ? String(row[statusCol - 1] || '').trim() : '';
      const errText = errorCol ? String(row[errorCol - 1] || '').trim() : '';
      const duration = durationCol ? Number(String(row[durationCol - 1] || '').replace(/[^0-9.\-]/g, '')) || 0 : 0;
      const eventText = eventCol ? String(row[eventCol - 1] || '').trim() : '';
      const phaseText = phaseCol ? String(row[phaseCol - 1] || '').trim() : '';
      const screenText = screenCol ? String(row[screenCol - 1] || '').trim() : '';
      const isError = !!errText || /fail|error|err|오류|실패/i.test(statusText);
      const isSlow = duration >= 5000;
      const isBlankSource = !eventText && !phaseText && !screenText;
      if (isError) errorCount++;
      if (isSlow) slowCount++;
      if (isBlankSource) blankSourceCount++;
      if ((isError || isSlow || isBlankSource) && samples.length < 10) {
        samples.push('로그행 ' + rowNo + ' / ' + (screenText || '-') + ' / ' + (eventText || phaseText || '-') + (duration ? ' / ' + duration + 'ms' : '') + (statusText ? ' / ' + statusText : '') + (errText ? ' / ' + errText : ''));
      }
    });
    const status = errorCount ? 'warn' : ((slowCount || blankSourceCount) ? 'warn' : 'ok');
    const message = recent.totalRows ? ('최근 ' + recent.rows.length + '행 기준 · 오류 ' + errorCount + '건 / 5초 이상 ' + slowCount + '건 / source 공란 ' + blankSourceCount + '건') : '성능로그 데이터 없음';
    return { status: status, message: message, detail: samples.length ? samples.join('\n') : '최근 오류/지연 특이사항 없음' };
  } catch (err) {
    return { status: 'warn', message: '성능로그 상태 확인 실패', detail: getPortalSystemHealthErrorTextP517_(err) };
  }
}


/***************************************
 * P518: 저장/동기화 플로우 진단
 * - 저장 실행/큐 처리/발송 차단 로직에는 개입하지 않고, 함수/스키마/최근 로그만 읽습니다.
 ***************************************/
function getPortalSaveQueueRequiredHeadersP518_() {
  try {
    if (typeof PORTAL_SAVE_QUEUE_P473 !== 'undefined' && PORTAL_SAVE_QUEUE_P473 && Array.isArray(PORTAL_SAVE_QUEUE_P473.HEADERS)) {
      return PORTAL_SAVE_QUEUE_P473.HEADERS.slice(0);
    }
  } catch (err) {}
  return PORTAL_SYSTEM_HEALTH_SAVE_QUEUE_FALLBACK_HEADERS_P518.slice(0);
}

function getPortalSaveQueueStaleRunningMinutesP518_() {
  try {
    if (typeof PORTAL_SAVE_QUEUE_P473 !== 'undefined' && PORTAL_SAVE_QUEUE_P473 && PORTAL_SAVE_QUEUE_P473.STALE_RUNNING_MINUTES) {
      return Math.max(1, Number(PORTAL_SAVE_QUEUE_P473.STALE_RUNNING_MINUTES) || 3);
    }
  } catch (err) {}
  return 3;
}

function isPortalSystemFunctionAvailableP518_(name) {
  name = String(name || '').trim();
  if (name === 'saveCustomerPatchFastP473') return typeof saveCustomerPatchFastP473 === 'function';
  if (name === 'processSaveQueueP473') return typeof processSaveQueueP473 === 'function';
  if (name === 'flushCustomerPendingOpsP473') return typeof flushCustomerPendingOpsP473 === 'function';
  return false;
}

function checkPortalSaveFlowFunctionsP518_() {
  try {
    const missing = [];
    const available = [];
    const detail = [];
    PORTAL_SYSTEM_HEALTH_SAVE_FLOW_FUNCTIONS_P518.forEach(function(item) {
      const ok = isPortalSystemFunctionAvailableP518_(item.name);
      if (ok) available.push(item.name);
      else if (item.required) missing.push(item.name);
      detail.push((ok ? 'OK ' : 'MISSING ') + item.name + ' - ' + item.label);
    });
    return {
      status: missing.length ? 'danger' : 'ok',
      message: missing.length ? ('필수 저장 함수 누락 ' + missing.length + '개: ' + missing.join(', ')) : ('필수 저장 함수 ' + available.length + '/' + PORTAL_SYSTEM_HEALTH_SAVE_FLOW_FUNCTIONS_P518.length + '개 확인'),
      detail: detail.join('\n')
    };
  } catch (err) {
    return { status: 'warn', message: '저장 핵심 함수 확인 실패', detail: getPortalSystemHealthErrorTextP517_(err) };
  }
}

function checkPortalSaveQueueSchemaP518_(webSs) {
  try {
    const sheetName = getPortalSaveQueueSheetNameP517_();
    const sheet = webSs.getSheetByName(sheetName);
    if (!sheet) {
      return { status: 'warn', message: sheetName + ' 없음: fallback 저장 발생 전이면 정상일 수 있음', detail: '시트가 생성되면 필수 헤더를 확인합니다.' };
    }
    const requiredHeaders = getPortalSaveQueueRequiredHeadersP518_();
    const headerStatus = getPortalSystemRequiredHeaderStatusP517_(sheet, requiredHeaders);
    const width = Math.max(sheet.getLastColumn() || 1, requiredHeaders.length || 1);
    const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
    const headerSet = {};
    headers.forEach(function(h) { if (h) headerSet[h] = true; });
    const missingCore = PORTAL_SYSTEM_HEALTH_SAVE_QUEUE_CORE_HEADERS_P518.filter(function(h) { return !headerSet[h]; });
    const missingAll = requiredHeaders.filter(function(h) { return !headerSet[h]; });
    let status = 'ok';
    if (missingCore.length) status = 'danger';
    else if (missingAll.length) status = 'warn';
    const lastRow = sheet.getLastRow();
    return {
      status: status,
      message: missingAll.length ? ('필수 헤더 누락 ' + missingAll.length + '개' + (missingCore.length ? ' / 핵심 누락 ' + missingCore.length + '개' : '')) : ('필수 헤더 ' + requiredHeaders.length + '개 확인'),
      detail: (missingAll.length ? ('전체 누락: ' + missingAll.join(', ')) : '필수 헤더 확인 완료') + (missingCore.length ? '\n핵심 누락: ' + missingCore.join(', ') : '') + '\n행수=' + Math.max(0, lastRow - 1) + ', lastCol=' + sheet.getLastColumn()
    };
  } catch (err) {
    return { status: 'warn', message: '저장큐 스키마 확인 실패', detail: getPortalSystemHealthErrorTextP517_(err) };
  }
}

function checkPortalRecentSavePerformanceP518_(webSs) {
  try {
    const sheetName = getPortalPerfLogSheetNameP517_();
    const sheet = webSs.getSheetByName(sheetName);
    if (!sheet) return { status: 'warn', message: sheetName + ' 없음: 최근 저장 로그 확인 불가', detail: '성능로그_DB가 생성되면 저장 관련 오류/지연을 확인합니다.' };
    const recent = getPortalSystemRecentRowsP517_(sheet, 200, 12);
    const statusCol = getPortalSystemColFromHeadersP517_(recent.headers, '상태');
    const errorCol = getPortalSystemColFromHeadersP517_(recent.headers, ['오류', '에러']);
    const durationCol = getPortalSystemColFromHeadersP517_(recent.headers, ['소요ms', 'durationMs']);
    const eventCol = getPortalSystemColFromHeadersP517_(recent.headers, ['이벤트', 'event']);
    const phaseCol = getPortalSystemColFromHeadersP517_(recent.headers, ['구간', 'phase']);
    const screenCol = getPortalSystemColFromHeadersP517_(recent.headers, ['화면', 'screen']);
    const customerCol = getPortalSystemColFromHeadersP517_(recent.headers, ['고객번호', 'customerNo']);
    const rowNoCol = getPortalSystemColFromHeadersP517_(recent.headers, ['rowNo', '마스터행']);
    const detailCol = getPortalSystemColFromHeadersP517_(recent.headers, ['상세JSON', 'detailJson']);
    let saveLogCount = 0;
    let errorCount = 0;
    let slowCount = 0;
    let busyCount = 0;
    let blankSourceCount = 0;
    const samples = [];
    recent.rows.forEach(function(row, i) {
      const rowNo = recent.startRow + i;
      const statusText = statusCol ? String(row[statusCol - 1] || '').trim() : '';
      const errText = errorCol ? String(row[errorCol - 1] || '').trim() : '';
      const duration = durationCol ? Number(String(row[durationCol - 1] || '').replace(/[^0-9.\-]/g, '')) || 0 : 0;
      const eventText = eventCol ? String(row[eventCol - 1] || '').trim() : '';
      const phaseText = phaseCol ? String(row[phaseCol - 1] || '').trim() : '';
      const screenText = screenCol ? String(row[screenCol - 1] || '').trim() : '';
      const customerNo = customerCol ? String(row[customerCol - 1] || '').trim() : '';
      const masterRow = rowNoCol ? String(row[rowNoCol - 1] || '').trim() : '';
      const detailText = detailCol ? String(row[detailCol - 1] || '').trim() : '';
      const combined = [statusText, errText, eventText, phaseText, screenText, detailText].join(' ');
      if (!isPortalSaveRelatedLogP518_(combined)) return;
      saveLogCount++;
      const isError = !!errText || /fail|error|err|오류|실패|충돌|conflict/i.test(statusText + ' ' + errText);
      const isSlow = duration >= 5000;
      const isBusy = /다른 작업 처리 중|server busy|busy|LockService|lock|잠금|timeout|Exceeded maximum execution time|Service invoked too many times|Internal error|서버/i.test(combined);
      const isBlankSource = !eventText && !phaseText && !screenText;
      if (isError) errorCount++;
      if (isSlow) slowCount++;
      if (isBusy) busyCount++;
      if (isBlankSource) blankSourceCount++;
      if ((isError || isSlow || isBusy || isBlankSource) && samples.length < 10) {
        samples.push('로그행 ' + rowNo + ' / ' + (screenText || '-') + ' / ' + (eventText || phaseText || '-') + (customerNo ? ' / 고객번호 ' + customerNo : '') + (masterRow ? ' / 마스터행 ' + masterRow : '') + (duration ? ' / ' + duration + 'ms' : '') + (statusText ? ' / ' + statusText : '') + (errText ? ' / ' + errText : ''));
      }
    });
    const status = (errorCount || busyCount) ? 'warn' : ((slowCount || blankSourceCount) ? 'warn' : 'ok');
    const message = saveLogCount ? ('최근 저장관련 ' + saveLogCount + '건 / 오류 ' + errorCount + '건 / busy·lock ' + busyCount + '건 / 5초이상 ' + slowCount + '건 / source공란 ' + blankSourceCount + '건') : '최근 200행 내 저장 관련 로그 없음';
    return { status: status, message: message, detail: samples.length ? samples.join('\n') : '최근 저장 관련 오류/지연 샘플 없음' };
  } catch (err) {
    return { status: 'warn', message: '최근 저장 오류/지연 확인 실패', detail: getPortalSystemHealthErrorTextP517_(err) };
  }
}

function isPortalSaveRelatedLogP518_(text) {
  text = String(text || '');
  return /save|저장|memo|메모|detail|patch|queue|flush|pending|customer\.detail|customer\.expandedMemo|customer\.detailPatch|support\.process|support\.request|자료발송|발송|send|LockService|busy|마스터시트/i.test(text);
}

function getPortalSystemDateMsP518_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value.getTime();
  const text = String(value == null ? '' : value).trim();
  if (!text) return 0;
  const direct = Date.parse(text);
  if (!isNaN(direct)) return direct;
  const m = text.match(/(20\d{2}|\d{2})[.\-\/]\s*(\d{1,2})[.\-\/]\s*(\d{1,2})(?:\s*(오전|오후|AM|PM|am|pm)?\s*(\d{1,2})[:시]\s*(\d{1,2})?(?:[:분]\s*(\d{1,2}))?)?/);
  if (!m) return 0;
  let year = Number(m[1]);
  if (year < 100) year += 2000;
  const month = Number(m[2]) - 1;
  const day = Number(m[3]);
  const marker = String(m[4] || '').toLowerCase();
  let hour = Number(m[5] || 0) || 0;
  const minute = Number(m[6] || 0) || 0;
  const second = Number(m[7] || 0) || 0;
  if ((marker === '오후' || marker === 'pm') && hour < 12) hour += 12;
  if ((marker === '오전' || marker === 'am') && hour === 12) hour = 0;
  const d = new Date(year, month, day, hour, minute, second);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function checkPortalCustomerFolderDbP517_(webSs) {
  try {
    const names = getPortalCustomerFolderSheetNamesP517_();
    const folderSheet = webSs.getSheetByName(names.folderSheetName);
    const itemSheet = webSs.getSheetByName(names.itemSheetName);
    let status = 'ok';
    const messages = [];
    const details = [];
    const folderHeaders = (typeof PORTAL_MY_CUSTOMER_FOLDER_HEADERS_P497 !== 'undefined') ? PORTAL_MY_CUSTOMER_FOLDER_HEADERS_P497 : ['folderId','ownerEmail','parentFolderId','folderName','sortOrder','isDeleted','createdAt','updatedAt'];
    const itemHeaders = (typeof PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497 !== 'undefined') ? PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497 : ['itemId','ownerEmail','folderId','customerNo','rowNo','companyNameSnapshot','assignedUserSnapshot','sortOrder','memo','isDeleted','createdAt','updatedAt'];

    const folderStatus = checkPortalFolderSheetOneP517_(folderSheet, names.folderSheetName, folderHeaders, 500);
    const itemStatus = checkPortalFolderSheetOneP517_(itemSheet, names.itemSheetName, itemHeaders, 500);
    [folderStatus, itemStatus].forEach(function(r) {
      if (r.status === 'danger') status = 'danger';
      else if (r.status === 'warn' && status !== 'danger') status = 'warn';
      messages.push(r.message);
      if (r.detail) details.push(r.detail);
    });
    return { status: status, message: messages.join(' / '), detail: details.join('\n') };
  } catch (err) {
    return { status: 'warn', message: '고객분류 DB 상태 확인 실패', detail: getPortalSystemHealthErrorTextP517_(err) };
  }
}

function checkPortalFolderSheetOneP517_(sheet, sheetName, requiredHeaders, limit) {
  if (!sheet) return { status: 'warn', message: sheetName + ' 없음', detail: sheetName + ': 사용 전 생성되지 않았거나 DB 초기화가 필요할 수 있습니다.' };
  const headerStatus = getPortalSystemRequiredHeaderStatusP517_(sheet, requiredHeaders);
  const recent = getPortalSystemRecentRowsP517_(sheet, limit || 500, requiredHeaders.length || 1);
  const deletedCol = getPortalSystemColFromHeadersP517_(recent.headers, ['isDeleted', '삭제여부']);
  let deletedRecent = 0;
  if (deletedCol) {
    recent.rows.forEach(function(row) {
      const v = String(row[deletedCol - 1] || '').trim().toUpperCase();
      if (v === 'Y' || v === 'TRUE' || v === '삭제') deletedRecent++;
    });
  }
  let status = headerStatus.ok ? 'ok' : 'danger';
  if (status === 'ok' && recent.rows.length >= 200 && deletedRecent / Math.max(1, recent.rows.length) >= 0.7) status = 'warn';
  const activeRecent = Math.max(0, recent.rows.length - deletedRecent);
  return {
    status: status,
    message: sheetName + ' 전체 ' + recent.totalRows + '건 / 최근 활성 ' + activeRecent + '건 / 삭제 ' + deletedRecent + '건',
    detail: sheetName + ': ' + headerStatus.detail + ' / lastCol=' + sheet.getLastColumn()
  };
}

function checkPortalContactProfileDbP517_(webSs) {
  try {
    const sheetName = getPortalContactProfileSheetNameP517_();
    const sheet = webSs.getSheetByName(sheetName);
    if (!sheet) return { status: 'warn', message: sheetName + ' 없음: 담당자 프로필 기능 사용 전 생성되지 않았을 수 있음', detail: '담당자 프로필 저장 시 생성됩니다.' };
    const requiredHeaders = (PORTAL_CONFIG && Array.isArray(PORTAL_CONFIG.CONTACT_PROFILE_HEADERS)) ? PORTAL_CONFIG.CONTACT_PROFILE_HEADERS : [];
    const headerStatus = getPortalSystemRequiredHeaderStatusP517_(sheet, requiredHeaders);
    const recent = getPortalSystemRecentRowsP517_(sheet, 200, requiredHeaders.length || 1);
    const activeCol = getPortalSystemColFromHeadersP517_(recent.headers, ['활성여부', 'active']);
    let inactive = 0;
    if (activeCol) {
      recent.rows.forEach(function(row) {
        const v = String(row[activeCol - 1] || '').trim().toUpperCase();
        if (v === 'N' || v === 'FALSE' || v === '비활성') inactive++;
      });
    }
    return {
      status: headerStatus.ok ? 'ok' : 'danger',
      message: '전체 ' + recent.totalRows + '건 / 최근 ' + recent.rows.length + '행 기준 비활성 ' + inactive + '건',
      detail: headerStatus.detail + ' / lastCol=' + sheet.getLastColumn()
    };
  } catch (err) {
    return { status: 'warn', message: '담당자프로필_DB 상태 확인 실패', detail: getPortalSystemHealthErrorTextP517_(err) };
  }
}

function getPortalSystemHealthErrorTextP517_(err) {
  return String(err && err.message ? err.message : err || '').replace(/Exception:|ScriptError/gi, '').trim() || '알 수 없는 오류';
}

/***************************************
 * P503: 시스템 점검 화면 보수공사용 경량 API
 * - 화면 진입 시 전체 점검을 한 번에 실행하지 않고, 카드별로 가볍게 조회합니다.
 ***************************************/
function getPortalSystemHealthPermissionDebugP503() {
  const started = new Date();
  const userInfo = getPortalSystemHealthUserP208_();
  // 권한 디버그 카드는 시스템 점검 화면을 열 수 있는 계정만 상세 반환합니다.
  assertPortalSystemHealthAllowedP208_(userInfo);

  let session = {};
  try {
    if (typeof debugPortalCurrentSessionUser === 'function') session = debugPortalCurrentSessionUser() || {};
  } catch (err) {
    session = { ok: false, message: err && err.message ? err.message : String(err) };
  }

  let cache = {};
  try {
    if (typeof getPortalPermissionCacheStatus === 'function') cache = getPortalPermissionCacheStatus() || {};
  } catch (err2) {
    cache = { ok: false, message: err2 && err2.message ? err2.message : String(err2) };
  }

  let permission = null;
  try {
    const p = getPortalCurrentPermission();
    permission = p && p.permission ? p.permission : p;
  } catch (err3) {
    permission = { active: false, message: err3 && err3.message ? err3.message : String(err3) };
  }

  return {
    ok: true,
    checkedAt: getPortalSystemNowP208_(new Date()),
    elapsedMs: new Date().getTime() - started.getTime(),
    activeEmail: session.activeEmail || userInfo.email || '',
    effectiveEmail: session.effectiveEmail || '',
    permission: permission || {},
    cache: cache || {},
    session: session || {}
  };
}

function getPortalSystemHealthConnectionChecksP503() {
  const started = new Date();
  const userInfo = getPortalSystemHealthUserP208_();
  assertPortalSystemHealthAllowedP208_(userInfo);

  const checks = [];
  function add(name, status, message, detail) {
    checks.push({ name: name, status: status || 'info', message: message || '', detail: detail || '' });
  }

  let masterSs = null;
  let webSs = null;

  try {
    masterSs = SpreadsheetApp.openById(PORTAL_CONFIG.MASTER_SPREADSHEET_ID);
    add('마스터 스프레드시트', 'ok', '연결됨: ' + masterSs.getName(), PORTAL_CONFIG.MASTER_SPREADSHEET_ID);
  } catch (err) {
    add('마스터 스프레드시트', 'danger', '연결 실패: ' + getPortalSystemHealthErrorTextP503_(err), PORTAL_CONFIG.MASTER_SPREADSHEET_ID);
  }

  try {
    webSs = getPortalSystemHealthWebDbSpreadsheetP208_();
    add('웹앱 DB 스프레드시트', 'ok', '연결됨: ' + webSs.getName(), webSs.getId());
  } catch (err2) {
    add('웹앱 DB 스프레드시트', 'danger', '연결 실패: ' + getPortalSystemHealthErrorTextP503_(err2), '');
  }

  const releaseInfoP517 = getPortalReleaseInfoP517_();
  add('릴리즈 기준본', releaseInfoP517.exists ? (releaseInfoP517.businessLogicChanged ? 'warn' : 'ok') : 'warn', releaseInfoP517.message, releaseInfoP517.detail);

  const saveFlowFunctionCheckP518 = checkPortalSaveFlowFunctionsP518_();
  add('저장 핵심 함수', saveFlowFunctionCheckP518.status, saveFlowFunctionCheckP518.message, saveFlowFunctionCheckP518.detail);

  try {
    if (masterSs) {
      const sheet = masterSs.getSheetByName(PORTAL_CONFIG.MASTER_SHEET_NAME);
      add('마스터시트', sheet ? 'ok' : 'danger', sheet ? ('데이터 ' + Math.max(0, sheet.getLastRow() - PORTAL_CONFIG.DATA_START_ROW + 1) + '행') : '시트 없음', PORTAL_CONFIG.MASTER_SHEET_NAME);
    }
  } catch (err3) {
    add('마스터시트', 'warn', '확인 지연/실패: ' + getPortalSystemHealthErrorTextP503_(err3), '');
  }

  if (webSs) {
    const saveQueueSchemaCheckP518 = checkPortalSaveQueueSchemaP518_(webSs);
    add('저장큐_DB 스키마', saveQueueSchemaCheckP518.status, saveQueueSchemaCheckP518.message, saveQueueSchemaCheckP518.detail);

    ['권한_DB', '검색인덱스_DB', '고객분류_폴더_DB', '고객분류_고객_DB', '저장큐_DB', '변경큐_DB', '성능로그_DB', '담당자프로필_DB'].forEach(function(name) {
      try {
        const sheet = webSs.getSheetByName(name);
        add(name, sheet ? 'ok' : getPortalSystemHealthMissingDbStatusP517_(name), sheet ? (Math.max(0, sheet.getLastRow() - 1) + '행') : '시트 없음', sheet ? ('lastCol=' + sheet.getLastColumn()) : '신규/선택 기능 시트는 사용 전 생성되지 않았을 수 있습니다.');
      } catch (err4) {
        add(name, 'warn', '확인 실패: ' + getPortalSystemHealthErrorTextP503_(err4), '');
      }
    });
  }

  try {
    const version = getPortalSystemVersionInfoP208_();
    add('검색인덱스 상태', version.customerIndexDirty === 'Y' ? 'warn' : 'ok', version.customerIndexDirty === 'Y' ? ('재생성 필요: ' + (version.customerIndexDirtyReason || '사유 미기록')) : '정상', 'version=' + (version.customerIndexVersion || '') + ', builtAt=' + (version.customerIndexBuiltAt || ''));
  } catch (err5) {
    add('검색인덱스 상태', 'warn', '확인 실패: ' + getPortalSystemHealthErrorTextP503_(err5), '');
  }

  return {
    ok: true,
    checkedAt: getPortalSystemNowP208_(new Date()),
    elapsedMs: new Date().getTime() - started.getTime(),
    checks: checks
  };
}

function clearMyCustomerFolderCacheFromHealthP503() {
  const userInfo = getPortalSystemHealthUserP208_();
  assertPortalSystemHealthAllowedP208_(userInfo);

  const cleared = [];
  let permission = null;
  try {
    const p = getPortalCurrentPermission();
    permission = p && p.permission ? p.permission : p;
  } catch (err) {}

  const emailKey = String((permission && permission.email) || userInfo.email || '').trim().toLowerCase();
  try {
    if (typeof invalidateMyCustomerFolderCacheP497_ === 'function' && emailKey) {
      invalidateMyCustomerFolderCacheP497_(emailKey);
      cleared.push(emailKey);
    }
  } catch (err2) {}

  try {
    if (typeof invalidateMyCustomerFolderCacheP497_ === 'function' && typeof PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500 !== 'undefined') {
      invalidateMyCustomerFolderCacheP497_(PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500);
      cleared.push(PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500);
    }
  } catch (err3) {}

  return {
    ok: true,
    cleared: cleared,
    message: '나의 고객 폴더 서버 캐시를 초기화했습니다.'
  };
}

function getPortalSystemHealthErrorTextP503_(err) {
  return String(err && err.message ? err.message : err || '').replace(/Exception:|ScriptError/gi, '').trim() || '알 수 없는 오류';
}
