/***************************************
 * S1 Sales Portal - 17_ActivityLogService.gs
 * PATCH P1-10: 작업 로그 / 개인·관리자 조회
 * - 영업담당자: 본인 로그만 조회
 * - 관리자/서무/책임: 전체 조회 + 담당자별 비교 통계
 ***************************************/

const PORTAL_ACTIVITY_LOG_SHEET_NAME = '작업로그_DB';
const PORTAL_ACTIVITY_LOG_HEADERS = [
  '로그ID',
  '일시',
  '일자',
  '시각',
  '사용자이메일',
  '사용자명',
  '역할',
  '영업담당자명',
  '작업구분',
  '화면',
  '고객번호',
  '회사명',
  '마스터행',
  '요약',
  '상세JSON'
];

function appendPortalActivityLog_(entry) {
  entry = entry || {};
  try {
    const perm = getPortalCurrentPermission_();
    const now = new Date();
    const tz = Session.getScriptTimeZone();
    const rowNo = Number(entry.rowNo || entry.masterRow || 0) || '';
    let customerNo = String(entry.customerNo || '').trim();
    let company = String(entry.company || '').trim();

    if ((!customerNo || !company) && rowNo) {
      try {
        const snap = getPortalActivityCustomerSnapshot_(rowNo);
        if (!customerNo) customerNo = snap.customerNo || '';
        if (!company) company = snap.company || '';
      } catch (e) {}
    }

    const detail = entry.detail == null ? '' : (typeof entry.detail === 'string' ? entry.detail : JSON.stringify(entry.detail));
    const sheet = ensurePortalActivityLogSheet_();
    sheet.appendRow([
      Utilities.getUuid().slice(0, 12),
      now,
      Utilities.formatDate(now, tz, 'yyyy-MM-dd'),
      Utilities.formatDate(now, tz, 'HH:mm:ss'),
      perm.email || '',
      perm.name || perm.email || '미등록 사용자',
      perm.level || perm.role || '',
      perm.salesRepName || perm.name || '',
      String(entry.actionType || entry.type || '기타').trim() || '기타',
      String(entry.screen || '').trim(),
      customerNo,
      company,
      rowNo,
      String(entry.summary || '').trim(),
      detail
    ]);
    return true;
  } catch (err) {
    Logger.log('작업로그 기록 실패: ' + (err && err.stack || err));
    return false;
  }
}

function getPortalActivityLogData(filters) {
  filters = filters || {};
  const perm = getPortalCurrentPermission_();
  const canViewAll = canPortalViewAllActivityLogs_(perm);
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const dateFrom = normalizePortalActivityDateText_(filters.dateFrom) || today;
  const dateTo = normalizePortalActivityDateText_(filters.dateTo) || dateFrom;
  const actor = String(filters.actor || '').trim();
  const type = String(filters.type || '').trim();
  const keyword = String(filters.keyword || '').trim().toLowerCase();
  const limit = Math.min(Math.max(Number(filters.limit) || 300, 20), 1000);

  const sheet = ensurePortalActivityLogSheet_();
  const lastRow = sheet.getLastRow();
  let rows = [];
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_ACTIVITY_LOG_HEADERS.length).getDisplayValues();
    rows = values.map(function(row, idx) {
      return buildPortalActivityLogItem_(row, idx + 2);
    });
  }

  rows = rows.filter(function(item) {
    if (item.date < dateFrom || item.date > dateTo) return false;
    if (!canViewAll && !isPortalActivityItemMine_(item, perm)) return false;
    if (canViewAll && actor && item.actorName !== actor && item.actorEmail !== actor && item.salesRepName !== actor) return false;
    if (type && item.actionType !== type) return false;
    if (keyword) {
      const hay = [item.actorName, item.actionType, item.screen, item.customerNo, item.company, item.summary, item.detailText].join(' ').toLowerCase();
      if (hay.indexOf(keyword) < 0) return false;
    }
    return true;
  });

  rows.sort(function(a, b) {
    return String(b.atText || '').localeCompare(String(a.atText || '')) || (Number(b.rowNo) - Number(a.rowNo));
  });

  const canCompare = canPortalCompareActivityLogs_(perm);
  const allActors = canViewAll ? getPortalActivityActorOptions_(rows) : [];
  const stats = buildPortalActivityStats_(rows);
  const byUser = canCompare ? buildPortalActivityByUserStats_(rows) : [];
  const byType = buildPortalActivityByTypeStats_(rows);

  return {
    ok: true,
    canViewAll: canViewAll,
    canCompare: canCompare,
    currentUser: sanitizePortalPermissionForClient_(perm),
    filters: {
      dateFrom: dateFrom,
      dateTo: dateTo,
      actor: actor,
      type: type,
      keyword: keyword,
      limit: limit
    },
    actorOptions: allActors,
    typeOptions: getPortalActivityTypeOptions_(),
    stats: stats,
    byUser: byUser,
    byType: byType,
    rows: rows.slice(0, limit),
    totalMatched: rows.length
  };
}

function ensurePortalActivityLogSheet_() {
  const ss = getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(PORTAL_ACTIVITY_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PORTAL_ACTIVITY_LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, PORTAL_ACTIVITY_LOG_HEADERS.length).setValues([PORTAL_ACTIVITY_LOG_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, PORTAL_ACTIVITY_LOG_HEADERS.length).setFontWeight('bold').setBackground('#f2f4f7');
    return sheet;
  }
  ensureSheetHeaders_(sheet, PORTAL_ACTIVITY_LOG_HEADERS);
  return sheet;
}

function buildPortalActivityLogItem_(row, rowNo) {
  const v = function(i) { return String(row[i] || '').trim(); };
  return {
    rowNo: rowNo,
    id: v(0),
    atText: v(1),
    date: v(2),
    time: v(3),
    actorEmail: v(4),
    actorName: v(5),
    role: v(6),
    salesRepName: v(7),
    actionType: v(8),
    screen: v(9),
    customerNo: v(10),
    company: v(11),
    masterRow: v(12),
    summary: v(13),
    detailText: v(14)
  };
}

function canPortalViewAllActivityLogs_(perm) {
  // PATCH P1-15: 작업로그 권한은 권한_DB의 작업로그전체열람/관리자성 권한에서 이미 계산된 값만 사용합니다.
  // 전체고객열람은 여기서 절대 참조하지 않습니다.
  perm = perm || getPortalCurrentPermission_();
  if (typeof canPortalViewAllActivityLogsByPermission_ === 'function') return canPortalViewAllActivityLogsByPermission_(perm);
  return !!(perm && perm.active !== false && perm.canViewAllActivityLogs);
}

function canPortalCompareActivityLogs_(perm) {
  perm = perm || getPortalCurrentPermission_();
  if (typeof canPortalCompareActivityLogsByPermission_ === 'function') return canPortalCompareActivityLogsByPermission_(perm);
  return !!(perm && perm.active !== false && perm.canCompareActivityLogs);
}

function isPortalActivityItemMine_(item, perm) {
  perm = perm || getPortalCurrentPermission_();
  const email = String(perm.email || '').trim().toLowerCase();
  if (email && String(item.actorEmail || '').trim().toLowerCase() === email) return true;
  const names = [];
  if (perm.name) names.push(perm.name);
  if (perm.salesRepName) names.push(perm.salesRepName);
  (perm.salesRepAliases || []).forEach(function(x) { if (x) names.push(x); });
  const actorName = normalizePortalNameForPermission_(item.actorName || '');
  const salesRepName = normalizePortalNameForPermission_(item.salesRepName || '');
  return names.some(function(name) {
    const n = normalizePortalNameForPermission_(name || '');
    return n && (actorName === n || salesRepName === n || actorName.indexOf(n) >= 0 || salesRepName.indexOf(n) >= 0);
  });
}

function getPortalActivityActorOptions_(rows) {
  const map = {};
  rows.forEach(function(item) {
    const key = item.actorName || item.salesRepName || item.actorEmail || '';
    if (!key) return;
    map[key] = { value: key, label: key, email: item.actorEmail || '', role: item.role || '' };
  });

  try {
    const sheet = ensurePortalPermissionSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_PERMISSION_HEADERS.length).getDisplayValues();
      values.forEach(function(r) {
        const email = String(r[0] || '').trim();
        const name = String(r[1] || '').trim();
        const role = String(r[4] || '').trim();
        const level = String(r[5] || '').trim();
        const salesRep = String(r[6] || '').trim() || name;
        const active = String(r[14] || 'Y').toUpperCase() !== 'N';
        if (!active || !name) return;
        map[name] = { value: name, label: name + (role ? ' · ' + role : ''), email: email, role: level || role };
        if (salesRep && salesRep !== name) map[salesRep] = { value: salesRep, label: salesRep + ' · 영업담당자명', email: email, role: level || role };
      });
    }
  } catch (err) {}

  return Object.keys(map).sort().map(function(k) { return map[k]; });
}

function getPortalActivityTypeOptions_() {
  return ['고객정보수정', '메모수정', '컨택이력추가', '영업지원요청', '자료발송', '즐겨찾기', '계약종합관리', '공지사항', '오늘할일', '기타'];
}

function buildPortalActivityStats_(rows) {
  const customerMap = {};
  rows.forEach(function(item) {
    const key = item.customerNo || item.company || item.masterRow || '';
    if (key) customerMap[key] = true;
  });
  return {
    total: rows.length,
    uniqueCustomers: Object.keys(customerMap).length,
    contactCount: rows.filter(function(x){ return x.actionType === '컨택이력추가'; }).length,
    supportCount: rows.filter(function(x){ return x.actionType === '영업지원요청'; }).length,
    sendCount: rows.filter(function(x){ return x.actionType === '자료발송'; }).length,
    editCount: rows.filter(function(x){ return x.actionType === '고객정보수정' || x.actionType === '메모수정'; }).length
  };
}

function buildPortalActivityByUserStats_(rows) {
  const map = {};
  rows.forEach(function(item) {
    const key = item.actorName || item.salesRepName || item.actorEmail || '미확인';
    if (!map[key]) map[key] = { actor: key, total: 0, contact: 0, support: 0, send: 0, edit: 0, contract: 0, notice: 0, todo: 0 };
    const s = map[key];
    s.total += 1;
    if (item.actionType === '컨택이력추가') s.contact += 1;
    else if (item.actionType === '영업지원요청') s.support += 1;
    else if (item.actionType === '자료발송') s.send += 1;
    else if (item.actionType === '고객정보수정' || item.actionType === '메모수정') s.edit += 1;
    else if (item.actionType === '계약종합관리') s.contract += 1;
    else if (item.actionType === '공지사항') s.notice += 1;
    else if (item.actionType === '오늘할일') s.todo += 1;
  });
  return Object.keys(map).map(function(k) { return map[k]; }).sort(function(a, b) { return b.total - a.total; });
}

function buildPortalActivityByTypeStats_(rows) {
  const map = {};
  rows.forEach(function(item) { map[item.actionType] = (map[item.actionType] || 0) + 1; });
  return Object.keys(map).map(function(k) { return { type: k, count: map[k] }; }).sort(function(a, b) { return b.count - a.count; });
}

function normalizePortalActivityDateText_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(value || '').trim();
  const m = s.match(/(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return '';
}

function getPortalActivityCustomerSnapshot_(rowNo) {
  const sheet = getMasterSpreadsheet_().getSheetByName(PORTAL_CONFIG.MASTER_SHEET_NAME);
  if (!sheet || !rowNo) return {};
  const obj = readMasterRowObject_(sheet, Number(rowNo));
  return {
    customerNo: getMasterFieldValue_(obj, 'customerNo') || obj['고객번호'] || '',
    company: getCompanyValue_(obj) || ''
  };
}
