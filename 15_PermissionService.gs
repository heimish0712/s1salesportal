/***************************************
 * S1 Sales Portal - 15_PermissionService.gs
 * PATCH S: 권한_DB 기반 사용자 인식/권한 제어
 * - 로그인 이메일 기준 사용자 식별
 * - 관리자/서무는 기존 권한 유지
 * - 영업팀은 기본적으로 본인 고객만 조회
 * - 영업팀은 공지 작성/삭제 불가, 영업지원 완료처리 불가
 ***************************************/

const PORTAL_PERMISSION_SHEET_NAME = '권한_DB';
const PORTAL_PERMISSION_HEADERS = [
  '이메일',
  '이름',
  '부서',
  '직급',
  '역할',
  '권한등급',
  '영업담당자명',
  '기본고객범위',
  '전체고객열람',
  '공지작성수정',
  '공지삭제',
  '영업지원요청작성',
  '영업지원전체열람',
  '영업지원완료처리',
  '사용여부',
  '비고',
  '작업로그전체열람',
  '작업로그비교통계'
];

const PORTAL_PERMISSION_SEED_ROWS = [
  ['pangsw712@gmail.com', '이옥희', '영업팀', '차장', '영업담당자', 'SALES', '이옥희', 'OWN', 'Y', 'N', 'N', 'Y', 'Y', 'N', 'Y', '영업팀: 기본 본인 고객만 조회', 'N', 'N'],
  ['xnewspringx@gmail.com', '박새봄', '기획팀', '대리', '서무', 'ADMIN', '', 'ALL', 'Y', 'Y', 'Y', 'Y', 'Y', 'Y', 'Y', '관리자/서무 계정', 'Y', 'Y'],
  ['okhee1359@gmail.com', '이옥희', '영업팀', '차장', '영업담당자', 'SALES', '이옥희', 'OWN', 'Y', 'N', 'N', 'Y', 'Y', 'N', 'Y', '영업팀: 기본 본인 고객만 조회', 'N', 'N'],
  ['testsh260613@gmail.com', '테스트영업', '영업팀', '대리', '영업담당자', 'SALES', '테스트', 'OWN', 'Y', 'N', 'N', 'Y', 'Y', 'N', 'Y', '테스트용 영업담당자 계정', 'N', 'N'],
  ['bang@s1samsung.com', '방수원(관리자)', '기획팀', '대리', '서무', 'ADMIN', '', 'ALL', 'Y', 'Y', 'Y', 'Y', 'Y', 'Y', 'Y', '관리자/서무 계정', 'Y', 'Y'],
  ['testsh260602@gmail.com', '테스트영업', '영업팀', '대리', '영업담당자', 'SALES', '테스트', 'OWN', 'Y', 'N', 'N', 'Y', 'Y', 'N', 'Y', '테스트용 영업담당자 계정', 'N', 'N'],
  ['mhj842@gmail.com', '문형진', '기획팀', '책임', '총괄', 'ADMIN', '문형진', 'ALL', 'Y', 'Y', 'Y', 'Y', 'Y', 'Y', 'Y', '관리자/서무 계정', 'Y', 'Y'],
  ['chl882000@gmail.com', '최보람', '영업팀', '팀장', '영업담당자', 'SALES', '최보람', 'OWN', 'Y', 'N', 'N', 'Y', 'Y', 'N', 'Y', '영업팀: 기본 본인 고객만 조회', 'N', 'N'],
  ['sworkskim7922@gmail.com', '김경아', '영업팀', '차장', '영업담당자', 'SALES', '김경아', 'OWN', 'Y', 'N', 'N', 'Y', 'Y', 'N', 'Y', '영업팀: 기본 본인 고객만 조회', 'N', 'N'],
  ['seoha3383@gmail.com', '김서하', '영업팀', '대리', '영업담당자', 'SALES', '김서하', 'OWN', 'Y', 'N', 'N', 'Y', 'Y', 'N', 'Y', '영업팀: 기본 본인 고객만 조회', 'N', 'N']
];

function setupPortalPermissionSheet() {
  const sheet = ensurePortalPermissionSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet.getRange(2, 1, PORTAL_PERMISSION_SEED_ROWS.length, PORTAL_PERMISSION_HEADERS.length).setValues(PORTAL_PERMISSION_SEED_ROWS);
  } else {
    const emailCol = 1;
    const existing = sheet.getRange(2, emailCol, lastRow - 1, 1).getDisplayValues()
      .map(function(r) { return String(r[0] || '').trim().toLowerCase(); });
    const append = PORTAL_PERMISSION_SEED_ROWS.filter(function(r) {
      return existing.indexOf(String(r[0] || '').trim().toLowerCase()) < 0;
    });
    if (append.length) sheet.getRange(lastRow + 1, 1, append.length, PORTAL_PERMISSION_HEADERS.length).setValues(append);
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, PORTAL_PERMISSION_HEADERS.length);
  return { ok: true, sheetName: sheet.getName(), rows: Math.max(0, sheet.getLastRow() - 1), headers: PORTAL_PERMISSION_HEADERS };
}

function resetPortalPermissionSheetForPatchS() {
  const sheet = ensurePortalPermissionSheet_();
  sheet.clearContents();
  sheet.getRange(1, 1, 1, PORTAL_PERMISSION_HEADERS.length).setValues([PORTAL_PERMISSION_HEADERS]);
  sheet.getRange(2, 1, PORTAL_PERMISSION_SEED_ROWS.length, PORTAL_PERMISSION_HEADERS.length).setValues(PORTAL_PERMISSION_SEED_ROWS);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, PORTAL_PERMISSION_HEADERS.length).setFontWeight('bold').setBackground('#f2f4f7');
  sheet.autoResizeColumns(1, PORTAL_PERMISSION_HEADERS.length);
  return { ok: true, reset: true, sheetName: sheet.getName(), rows: PORTAL_PERMISSION_SEED_ROWS.length };
}

function getPortalCurrentPermission() {
  const perm = getPortalCurrentPermission_();
  return { ok: true, permission: sanitizePortalPermissionForClient_(perm) };
}

function getPortalPermissionAuditMatrix() {
  const sheet = ensurePortalPermissionSheet_();
  const lastRow = sheet.getLastRow();
  const rows = [];
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_PERMISSION_HEADERS.length).getDisplayValues();
    values.forEach(function(row, idx) {
      const p = buildPortalPermissionFromRow_(row, idx + 2, row[0]);
      rows.push({
        rowNo: idx + 2,
        email: p.email,
        name: p.name,
        role: p.role,
        level: p.level,
        salesRepName: p.salesRepName,
        defaultScope: p.defaultScope,
        active: p.active,
        canUseAdminHome: p.canUseAdminHome,
        canViewAllCustomers: p.canViewAllCustomers,
        canWriteNotice: p.canWriteNotice,
        canDeleteNotice: p.canDeleteNotice,
        canWriteSupport: p.canWriteSupport,
        canReadAllSupport: p.canReadAllSupport,
        canCompleteSupport: p.canCompleteSupport,
        canViewAllActivityLogs: p.canViewAllActivityLogs,
        canCompareActivityLogs: p.canCompareActivityLogs
      });
    });
  }
  return { ok: true, headers: PORTAL_PERMISSION_HEADERS, rows: rows };
}


function getPortalCurrentUserName_() {
  const p = sanitizePortalPermissionForClient_(getPortalCurrentPermission_());
  return p.name || p.email || '웹앱사용자';
}

function getPortalCurrentUserFullLabel_() {
  const p = sanitizePortalPermissionForClient_(getPortalCurrentPermission_());
  return p.displayName || p.name || p.email || '웹앱사용자';
}

function debugPortalCurrentSessionUser() {
  const activeEmail = getPortalActiveUserEmail_();
  const effectiveEmail = getPortalEffectiveUserEmail_();
  const selectedEmail = getPortalSessionEmail_();
  const tempKey = getPortalTemporaryActiveUserKey_();
  const perm = getPortalCurrentPermission_();
  return {
    ok: true,
    activeEmail: activeEmail,
    effectiveEmail: effectiveEmail,
    selectedEmail: selectedEmail,
    temporaryActiveUserKey: tempKey,
    executeAsOwnerDetected: !!effectiveEmail && !!activeEmail && effectiveEmail !== activeEmail,
    activeEmailBlank: !activeEmail,
    note: !activeEmail
      ? '접속자 이메일을 확인하지 못했습니다. Execute as me 환경에서는 외부 Gmail의 active user email이 빈값일 수 있습니다. 이 경우 effectiveUser로 대체하지 않고 권한을 차단합니다.'
      : '',
    permission: sanitizePortalPermissionForClient_(perm)
  };
}

function getPortalAuthDiagnostics() {
  return debugPortalCurrentSessionUser();
}

function ensurePortalPermissionSheet_() {
  const ss = getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(PORTAL_PERMISSION_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(PORTAL_PERMISSION_SHEET_NAME);

  const width = Math.max(PORTAL_PERMISSION_HEADERS.length, sheet.getLastColumn() || PORTAL_PERMISSION_HEADERS.length);
  const current = sheet.getRange(1, 1, 1, width).getDisplayValues()[0];
  PORTAL_PERMISSION_HEADERS.forEach(function(h, i) {
    if (String(current[i] || '').trim() !== h) sheet.getRange(1, i + 1).setValue(h);
  });
  sheet.getRange(1, 1, 1, PORTAL_PERMISSION_HEADERS.length).setFontWeight('bold').setBackground('#f2f4f7');
  return sheet;
}

function getPortalCurrentPermission_() {
  const email = getPortalSessionEmail_();
  const emailKey = String(email || '').trim().toLowerCase();
  const sheet = ensurePortalPermissionSheet_();
  if (sheet.getLastRow() < 2) setupPortalPermissionSheet();

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_PERMISSION_HEADERS.length).getDisplayValues();
    for (let i = 0; i < values.length; i++) {
      const p = buildPortalPermissionFromRow_(values[i], i + 2, email);
      if (p.emailKey && p.emailKey === emailKey && p.active) return p;
    }
  }

  return buildGuestPortalPermission_(email);
}

function buildPortalPermissionFromRow_(row, rowNo, sessionEmail) {
  const v = function(idx) { return String(row[idx] || '').trim(); };
  const email = v(0);
  const name = v(1);
  const dept = v(2);
  const rank = v(3);
  const role = v(4);
  const level = String(v(5) || '').toUpperCase() || 'GUEST';
  const salesRepName = v(6) || name;
  const defaultScope = String(v(7) || '').toUpperCase() || (level === 'SALES' ? 'OWN' : 'ALL');

  // PATCH P1-15: 권한 판정 중앙화.
  // 화면별 권한은 권한_DB 컬럼 의미를 분리해서 사용합니다.
  // - 전체고객열람: 고객 목록/검색에서 전체 고객을 볼 수 있는 권한
  // - 영업지원전체열람: 영업지원 요청 전체 목록 조회 권한
  // - 영업지원완료처리: 영업지원 처리/완료 입력 권한
  // - 작업로그전체열람/작업로그비교통계: 작업로그 전용 권한
  // SALES 계정에 전체고객열람=Y가 있어도 HOME/작업로그/완료처리는 관리자 권한으로 보지 않습니다.
  const baseAdmin = isPortalAdminPermissionValues_(level, role, rank, defaultScope);
  const canViewAllCustomers = baseAdmin || yn_(v(8)) || defaultScope === 'ALL';
  const canWriteNotice = baseAdmin || yn_(v(9));
  const canDeleteNotice = baseAdmin || yn_(v(10));
  const canWriteSupport = baseAdmin || yn_(v(11));
  const canReadAllSupport = baseAdmin || yn_(v(12));
  const canCompleteSupport = baseAdmin || yn_(v(13));
  const canViewAllActivityLogs = baseAdmin || yn_(v(16));
  const canCompareActivityLogs = canViewAllActivityLogs && (baseAdmin || yn_(v(17)));
  const canUseAdminHome = baseAdmin;

  return {
    rowNo: rowNo,
    email: email || sessionEmail || '',
    emailKey: String(email || '').trim().toLowerCase(),
    name: name || email || sessionEmail || '미등록 사용자',
    dept: dept,
    rank: rank,
    role: role,
    level: level,
    salesRepName: salesRepName,
    salesRepAliases: splitPortalPermissionAliases_(salesRepName),
    defaultScope: defaultScope,
    isAdmin: baseAdmin,
    canUseAdminHome: canUseAdminHome,
    canViewAllCustomers: canViewAllCustomers,
    canWriteNotice: canWriteNotice,
    canDeleteNotice: canDeleteNotice,
    canWriteSupport: canWriteSupport,
    canReadAllSupport: canReadAllSupport,
    canCompleteSupport: canCompleteSupport,
    active: String(v(14) || 'Y').toUpperCase() !== 'N',
    note: v(15),
    canViewAllActivityLogs: canViewAllActivityLogs,
    canCompareActivityLogs: canCompareActivityLogs
  };
}

function isPortalAdminPermissionValues_(level, role, rank, defaultScope) {
  level = String(level || '').toUpperCase();
  role = String(role || '');
  rank = String(rank || '');
  defaultScope = String(defaultScope || '').toUpperCase();
  if (level === 'ADMIN' || level === 'MANAGER') return true;
  if (defaultScope === 'ALL') return true;
  if (role.indexOf('서무') >= 0 || role.indexOf('총괄') >= 0 || role.indexOf('관리자') >= 0) return true;
  // 직급이 책임이어도 영업담당자 역할이면 관리자 권한을 주지 않습니다.
  if (rank.indexOf('책임') >= 0 && role.indexOf('영업담당자') < 0) return true;
  return false;
}

function canPortalUseAdminHome_(perm) {
  perm = perm || getPortalCurrentPermission_();
  return !!(perm && perm.active !== false && perm.canUseAdminHome);
}

function canPortalViewAllCustomerScope_(perm) {
  perm = perm || getPortalCurrentPermission_();
  return !!(perm && perm.active !== false && perm.canViewAllCustomers);
}

function canPortalReadAllSupport_(perm) {
  perm = perm || getPortalCurrentPermission_();
  return !!(perm && perm.active !== false && perm.canReadAllSupport);
}

function canPortalCompleteSupport_(perm) {
  perm = perm || getPortalCurrentPermission_();
  return !!(perm && perm.active !== false && perm.canCompleteSupport);
}

function canPortalViewAllActivityLogsByPermission_(perm) {
  perm = perm || getPortalCurrentPermission_();
  return !!(perm && perm.active !== false && perm.canViewAllActivityLogs);
}

function canPortalCompareActivityLogsByPermission_(perm) {
  perm = perm || getPortalCurrentPermission_();
  return !!(perm && perm.active !== false && perm.canCompareActivityLogs);
}

function buildGuestPortalPermission_(email) {
  const key = String(email || '').trim();
  return {
    rowNo: 0,
    email: key,
    emailKey: key.toLowerCase(),
    name: key || '미등록 사용자',
    dept: '',
    rank: '',
    role: '미등록',
    level: 'GUEST',
    salesRepName: '',
    salesRepAliases: [],
    defaultScope: 'OWN',
    isAdmin: false,
    canUseAdminHome: false,
    canViewAllCustomers: false,
    canWriteNotice: false,
    canDeleteNotice: false,
    canWriteSupport: false,
    canReadAllSupport: false,
    canCompleteSupport: false,
    active: false,
    note: '권한_DB에 등록되지 않은 사용자',
    canViewAllActivityLogs: false,
    canCompareActivityLogs: false
  };
}

function sanitizePortalPermissionForClient_(p) {
  p = p || buildGuestPortalPermission_('');
  const displayName = [p.dept, p.name, p.rank].filter(Boolean).join(' ') || p.name || p.email || '미등록 사용자';
  return {
    email: p.email || '',
    name: p.name || '',
    dept: p.dept || '',
    rank: p.rank || '',
    role: p.role || '',
    level: p.level || 'GUEST',
    displayName: displayName,
    salesRepName: p.salesRepName || '',
    defaultScope: p.defaultScope || 'OWN',
    isAdmin: !!p.isAdmin,
    canUseAdminHome: !!p.canUseAdminHome,
    canViewAllCustomers: !!p.canViewAllCustomers,
    canWriteNotice: !!p.canWriteNotice,
    canDeleteNotice: !!p.canDeleteNotice,
    canWriteSupport: !!p.canWriteSupport,
    canReadAllSupport: !!p.canReadAllSupport,
    canCompleteSupport: !!p.canCompleteSupport,
    canViewAllActivityLogs: !!p.canViewAllActivityLogs,
    canCompareActivityLogs: !!p.canCompareActivityLogs,
    active: !!p.active,
    note: p.note || ''
  };
}

function getPortalSessionEmail_() {
  // PATCH S-FIX4: 권한 판정은 접속자(active user) 이메일만 사용합니다.
  // Execute as me 환경에서 effective user는 배포자/소유자일 수 있으므로,
  // active user가 비어 있다고 effective user로 대체하면 모든 미식별 사용자가 관리자처럼 보일 수 있습니다.
  return getPortalActiveUserEmail_();
}

function getPortalActiveUserEmail_() {
  try { return String(Session.getActiveUser().getEmail() || '').trim(); } catch (err) { return ''; }
}

function getPortalEffectiveUserEmail_() {
  // 진단용입니다. 권한 판정에 사용하지 않습니다.
  try { return String(Session.getEffectiveUser().getEmail() || '').trim(); } catch (err) { return ''; }
}

function getPortalTemporaryActiveUserKey_() {
  try { return String(Session.getTemporaryActiveUserKey() || '').trim(); } catch (err) { return ''; }
}

function getPortalWebAppEntryAuth_() {
  const activeEmail = getPortalActiveUserEmail_();
  const effectiveEmail = getPortalEffectiveUserEmail_();
  if (!activeEmail) {
    return {
      ok: false,
      reason: 'ACTIVE_USER_EMAIL_EMPTY',
      message: '현재 접속자의 Google 이메일을 확인하지 못했습니다.',
      activeEmail: '',
      effectiveEmail: effectiveEmail,
      temporaryActiveUserKey: getPortalTemporaryActiveUserKey_()
    };
  }
  const perm = getPortalCurrentPermission_();
  if (!perm || !perm.active || String(perm.level || '').toUpperCase() === 'GUEST') {
    return {
      ok: false,
      reason: 'PERMISSION_DB_NOT_ALLOWED',
      message: '권한_DB에 등록되지 않았거나 사용 중지된 계정입니다.',
      activeEmail: activeEmail,
      effectiveEmail: effectiveEmail,
      permission: sanitizePortalPermissionForClient_(perm)
    };
  }
  return {
    ok: true,
    reason: 'OK',
    activeEmail: activeEmail,
    effectiveEmail: effectiveEmail,
    permission: sanitizePortalPermissionForClient_(perm)
  };
}

function buildPortalAccessDeniedHtml_(auth) {
  auth = auth || {};
  const email = String(auth.activeEmail || '').replace(/[<>&"]/g, function(ch) {
    return ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'}[ch] || ch);
  });
  const reason = String(auth.reason || '').replace(/[<>&"]/g, function(ch) {
    return ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'}[ch] || ch);
  });
  const message = String(auth.message || '접속 권한이 없습니다.').replace(/[<>&"]/g, function(ch) {
    return ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'}[ch] || ch);
  });
  return '<!DOCTYPE html><html><head><base target="_top"><meta charset="UTF-8">' +
    '<style>body{margin:0;background:#eef2f6;font-family:Malgun Gothic,Arial,sans-serif;color:#111827;}' +
    '.wrap{max-width:720px;margin:80px auto;background:#fff;border:1px solid #d7dde8;padding:28px 32px;border-radius:10px;box-shadow:0 10px 24px rgba(16,24,40,.06);}' +
    'h1{font-size:22px;margin:0 0 12px;font-weight:900}.msg{line-height:1.7;color:#344054}.code{margin-top:14px;padding:12px;background:#f8fafc;border:1px solid #e4e7ec;border-radius:8px;font-size:13px;color:#475467}' +
    '</style></head><body><div class="wrap"><h1>S1 Sales Portal 접속 제한</h1>' +
    '<div class="msg">' + message + '<br>관리자에게 권한_DB 등록 상태를 확인해 주세요.</div>' +
    '<div class="code">접속계정: ' + (email || '확인 불가') + '<br>사유: ' + reason + '</div>' +
    '</div></body></html>';
}

function yn_(value) {
  const text = String(value || '').trim().toUpperCase();
  return text === 'Y' || text === 'YES' || text === 'TRUE' || text === 'O' || text === '1' || text === '허용' || text === '가능';
}

function splitPortalPermissionAliases_(value) {
  return String(value || '').split(/[,:，、\/|\n]/).map(function(v) { return String(v || '').trim(); }).filter(Boolean);
}

function normalizePortalNameForPermission_(value) {
  return String(value || '').replace(/\s+/g, '').trim();
}

function getPortalCustomerSalesRepFromRow_(row) {
  row = row || {};
  let value = '';
  try { value = row.salesRep || row['영업담당자'] || row['견적담당'] || row['담당영업'] || ''; } catch (err) {}
  if (!value) {
    try { value = getCustomerMasterHeaderValueK2_(row, 'salesRep'); } catch (err) {}
  }
  return String(value || '').trim();
}

function isPortalCustomerRowAllowedForPermission_(row, perm) {
  perm = perm || getPortalCurrentPermission_();
  if (!perm.active) return false;
  if (perm.canViewAllCustomers || perm.isAdmin) return true;
  const salesRep = normalizePortalNameForPermission_(getPortalCustomerSalesRepFromRow_(row));
  const aliases = perm.salesRepAliases || [];
  if (!salesRep || !aliases.length) return false;
  return aliases.some(function(alias) {
    const a = normalizePortalNameForPermission_(alias);
    return a && (salesRep === a || salesRep.indexOf(a) >= 0 || a.indexOf(salesRep) >= 0);
  });
}

function filterPortalCustomerRowsByPermission_(rows) {
  rows = Array.isArray(rows) ? rows : [];
  const perm = getPortalCurrentPermission_();
  if (perm.canViewAllCustomers || perm.isAdmin) return rows;
  return rows.filter(function(row) { return isPortalCustomerRowAllowedForPermission_(row, perm); });
}

function assertPortalCanAccessCustomerObject_(obj, action) {
  const perm = getPortalCurrentPermission_();
  if (isPortalCustomerRowAllowedForPermission_(obj, perm)) return true;
  throw new Error((action || '고객 접근') + ' 권한이 없습니다. 현재 계정은 본인 담당 고객만 조회/수정할 수 있습니다.');
}

function assertPortalCanAccessCustomerTarget_(target, action) {
  target = target || {};
  if (target.obj) return assertPortalCanAccessCustomerObject_(target.obj, action);
  const sheet = target.sheet || (getMasterSpreadsheet_().getSheetByName(PORTAL_CONFIG.MASTER_SHEET_NAME));
  const rowNo = Number(target.rowNo) || 0;
  if (!sheet || !rowNo) throw new Error((action || '고객 접근') + ' 대상이 올바르지 않습니다.');
  const obj = readMasterRowObject_(sheet, rowNo);
  obj.__rowNo = rowNo;
  return assertPortalCanAccessCustomerObject_(obj, action);
}

function assertPortalCanWriteNotice_() {
  const perm = getPortalCurrentPermission_();
  if (perm.canWriteNotice) return true;
  throw new Error('공지사항 작성 권한이 없습니다.');
}

function assertPortalCanDeleteNotice_() {
  const perm = getPortalCurrentPermission_();
  if (perm.canDeleteNotice) return true;
  throw new Error('공지사항 삭제 권한이 없습니다.');
}

function assertPortalCanWriteSupport_() {
  const perm = getPortalCurrentPermission_();
  if (perm.canWriteSupport) return true;
  throw new Error('영업지원 요청 작성 권한이 없습니다.');
}

function assertPortalCanReadAllSupport_() {
  const perm = getPortalCurrentPermission_();
  if (canPortalReadAllSupport_(perm)) return true;
  throw new Error('영업지원 전체 열람 권한이 없습니다.');
}

function assertPortalCanCompleteSupport_() {
  const perm = getPortalCurrentPermission_();
  if (perm.canCompleteSupport) return true;
  throw new Error('영업지원 완료 처리 권한이 없습니다.');
}
