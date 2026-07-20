/***************************************
 * S1 Sales Portal - 14_FavoriteService.gs
 * PATCH O: 유저별 고객 즐겨찾기
 * 원칙: 즐겨찾기는 userKey 기준 개인별 적용. 검색인덱스_DB는 빠른 표시용으로만 사용.
 ***************************************/

const PORTAL_FAVORITE_SHEET_NAME = '고객즐겨찾기_DB';
const PORTAL_FAVORITE_HEADERS = [
  '즐겨찾기ID',
  '사용자키',
  '사용자명',
  '고객번호',
  '마스터행',
  '회사명스냅샷',
  '추가일시',
  '삭제여부',
  '삭제일시'
];

function getPortalCustomerFavoriteMap() {
  const user = getPortalFavoriteUser_();
  const rows = getPortalFavoriteRows_({ includeDeleted: false });
  const map = {};
  rows.forEach(function(r) {
    const key = makePortalFavoriteCustomerKey_(r.customerNo, r.rowNo);
    if (key) map[key] = true;
  });

  const response = {
    ok: true,
    map: map,
    total: rows.length,
    currentUserKey: user.key,
    currentUserLabel: user.label
  };

  // P549: 관리자/서무는 모든 사용자의 활성 즐겨찾기를 조회할 수 있습니다.
  // 단, map은 항상 현재 로그인 사용자의 즐겨찾기만 반환하여 별표 추가/해제가
  // 다른 사용자의 즐겨찾기를 건드리지 않도록 분리합니다.
  if (canPortalViewAllFavoritesP549_()) {
    const allRows = getPortalFavoriteRows_({ includeDeleted: false, allUsers: true });
    response.adminView = buildPortalFavoriteAdminViewP549_(allRows, user);
  } else {
    response.adminView = { allowed: false, entries: [], owners: [] };
  }
  return response;
}

function canPortalViewAllFavoritesP549_() {
  let perm = null;
  try { perm = getPortalCurrentPermission_(); } catch (err) {}
  return !!(perm && perm.active !== false && (perm.isAdmin || perm.canUseAdminHome || String(perm.level || '').toUpperCase() === 'ADMIN'));
}

function buildPortalFavoriteAdminViewP549_(rows, currentUser) {
  rows = Array.isArray(rows) ? rows : [];
  currentUser = currentUser || getPortalFavoriteUser_();
  const entryByRelation = {};

  rows.forEach(function(r) {
    const customerKey = makePortalFavoriteCustomerKey_(r.customerNo, r.rowNo);
    const userKey = String(r.userKey || '').trim();
    if (!customerKey || !userKey) return;
    const relationKey = userKey + '||' + customerKey;
    const previous = entryByRelation[relationKey];
    if (previous && String(previous.addedAt || '') >= String(r.addedAt || '')) return;
    entryByRelation[relationKey] = {
      userKey: userKey,
      userLabel: String(r.userLabel || userKey).trim() || userKey,
      customerKey: customerKey,
      customerNo: String(r.customerNo || '').trim(),
      rowNo: Number(r.rowNo) || 0,
      company: String(r.company || '').trim(),
      addedAt: String(r.addedAt || '').trim()
    };
  });

  const entries = Object.keys(entryByRelation).map(function(k) { return entryByRelation[k]; });
  entries.sort(function(a, b) {
    const byDate = String(b.addedAt || '').localeCompare(String(a.addedAt || ''));
    if (byDate) return byDate;
    return String(a.userLabel || '').localeCompare(String(b.userLabel || ''), 'ko');
  });

  const ownerMap = {};
  entries.forEach(function(entry) {
    const ownerKey = entry.userKey;
    if (!ownerMap[ownerKey]) {
      ownerMap[ownerKey] = {
        userKey: ownerKey,
        userLabel: entry.userLabel || ownerKey,
        customerKeys: {}
      };
    }
    ownerMap[ownerKey].customerKeys[entry.customerKey] = true;
  });

  const owners = Object.keys(ownerMap).map(function(k) {
    return {
      userKey: ownerMap[k].userKey,
      userLabel: ownerMap[k].userLabel,
      count: Object.keys(ownerMap[k].customerKeys || {}).length,
      isCurrentUser: String(ownerMap[k].userKey || '') === String(currentUser.key || '')
    };
  }).sort(function(a, b) {
    if (!!a.isCurrentUser !== !!b.isCurrentUser) return a.isCurrentUser ? -1 : 1;
    return String(a.userLabel || '').localeCompare(String(b.userLabel || ''), 'ko');
  });

  return {
    allowed: true,
    currentUserKey: currentUser.key,
    currentUserLabel: currentUser.label,
    totalRelations: entries.length,
    totalCustomers: Object.keys(entries.reduce(function(acc, e) { acc[e.customerKey] = true; return acc; }, {})).length,
    entries: entries,
    owners: owners
  };
}

function getPortalFavoriteCustomers() {
  const favRows = getPortalFavoriteRows_({ includeDeleted: false });
  const favMap = {};
  favRows.forEach(function(r) {
    const key = makePortalFavoriteCustomerKey_(r.customerNo, r.rowNo);
    if (key) favMap[key] = r;
  });

  const indexData = getCustomerSearchIndexData();
  const indexRows = indexData.rows || [];
  const result = [];

  indexRows.forEach(function(row) {
    const key = makePortalFavoriteCustomerKey_(row.customerNo, row.rowNo);
    if (!key || !favMap[key]) return;
    const item = Object.assign({}, row);
    item.favorite = true;
    item.favoriteAddedAt = favMap[key].addedAt || '';
    result.push(item);
  });

  // 인덱스에서 못 찾은 즐겨찾기는 스냅샷으로라도 표시합니다.
  favRows.forEach(function(fav) {
    const key = makePortalFavoriteCustomerKey_(fav.customerNo, fav.rowNo);
    if (!key) return;
    const exists = result.some(function(r) { return makePortalFavoriteCustomerKey_(r.customerNo, r.rowNo) === key; });
    if (!exists) {
      result.push({
        rowNo: fav.rowNo,
        customerNo: fav.customerNo,
        company: fav.company,
        status: '',
        salesRep: '',
        contact: '',
        phone: '',
        directPhone: '',
        email: '',
        vendor: '',
        finalQuote: '',
        memo: '',
        favorite: true,
        favoriteAddedAt: fav.addedAt || '',
        __source: 'favoriteSnapshot'
      });
    }
  });

  const map = {};
  favRows.forEach(function(r) {
    const key = makePortalFavoriteCustomerKey_(r.customerNo, r.rowNo);
    if (key) map[key] = true;
  });

  result.sort(function(a, b) {
    return String(b.favoriteAddedAt || '').localeCompare(String(a.favoriteAddedAt || ''));
  });

  return { ok: true, rows: result, map: map, total: result.length, source: '고객즐겨찾기_DB + 검색인덱스_DB' };
}

function togglePortalCustomerFavorite(payload) {
  // 하위 호환용: 기존 클라이언트 호출도 명시 상태 저장 함수로 처리합니다.
  return setPortalCustomerFavorite(payload || {});
}

function setPortalCustomerFavorite(payload) {
  // 기존 호출 호환용. P2-4 이후 클라이언트는 아래 lightweight 함수를 사용합니다.
  const result = writePortalCustomerFavoriteStateP204_(payload || {}, { lockLabel: 'favorite-set' });

  try {
    appendPortalActivityLog_({
      actionType: '즐겨찾기',
      screen: '고객 즐겨찾기',
      rowNo: result.rowNo,
      customerNo: result.customerNo,
      company: result.company,
      summary: result.favorite ? '즐겨찾기 추가' : '즐겨찾기 해제',
      detail: { favorite: result.favorite, action: result.action }
    });
  } catch (err) {}

  const mapRes = getPortalCustomerFavoriteMap();
  const listRes = getPortalFavoriteCustomers();
  return {
    ok: true,
    favorite: result.favorite,
    customerNo: result.customerNo,
    rowNo: result.rowNo,
    map: mapRes.map || {},
    favorites: listRes.rows || [],
    action: result.action
  };
}

/**
 * P2-4: 즐겨찾기 빠른 저장 전용.
 * - 화면은 이미 optimistic 처리했으므로 map/favorite list 전체를 다시 만들지 않습니다.
 * - 고객즐겨찾기_DB만 멱등 upsert/delete 처리합니다.
 */
function setPortalCustomerFavoriteStateP204(payload) {
  const result = writePortalCustomerFavoriteStateP204_(payload || {}, { lockLabel: 'favorite-state-p204' });
  try {
    appendPortalActivityLog_({
      actionType: '즐겨찾기',
      screen: '고객 즐겨찾기',
      rowNo: result.rowNo,
      customerNo: result.customerNo,
      company: result.company,
      summary: result.favorite ? '즐겨찾기 추가' : '즐겨찾기 해제',
      detail: { favorite: result.favorite, action: result.action, mode: 'lightweight' }
    });
  } catch (err) {}
  return result;
}

function writePortalCustomerFavoriteStateP204_(payload, options) {
  payload = payload || {};
  options = options || {};
  const user = getPortalFavoriteUser_();
  const customerNo = String(payload.customerNo || '').trim();
  const rowNo = Number(payload.rowNo) || 0;
  const company = String(payload.company || '').trim();
  const favorite = payload.favorite !== false;
  const key = makePortalFavoriteCustomerKey_(customerNo, rowNo);
  if (!key) throw new Error('즐겨찾기할 고객번호 또는 행 정보가 없습니다.');

  const writeResult = withPortalScriptLockP201_(options.lockLabel || 'favorite-state', function() {
    const sheet = ensurePortalFavoriteSheet_();
    const now = new Date();
    const nowText = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const lastRow = sheet.getLastRow();
    const matches = [];

    if (lastRow >= 2) {
      const width = Math.max(PORTAL_FAVORITE_HEADERS.length, sheet.getLastColumn());
      const values = sheet.getRange(2, 1, lastRow - 1, width).getDisplayValues();
      for (let i = 0; i < values.length; i++) {
        const rUser = String(values[i][1] || '').trim();
        const rCustomerNo = String(values[i][3] || '').trim();
        const rRowNo = Number(values[i][4]) || 0;
        if (rUser === user.key && makePortalFavoriteCustomerKey_(rCustomerNo, rRowNo) === key) {
          matches.push({ row: i + 2, deleted: String(values[i][7] || '').trim().toUpperCase() === 'Y' });
        }
      }
    }

    let action = 'noop';
    if (favorite) {
      const primary = matches.length ? matches[0].row : 0;
      if (primary) {
        sheet.getRange(primary, 2, 1, 8).setValues([[
          user.key,
          user.label,
          customerNo,
          rowNo || '',
          company,
          nowText,
          '',
          ''
        ]]);
        action = matches[0].deleted ? 'restore' : 'keep-on';
      } else {
        sheet.appendRow([
          Utilities.getUuid().slice(0, 10),
          user.key,
          user.label,
          customerNo,
          rowNo || '',
          company,
          nowText,
          '',
          ''
        ]);
        action = 'add';
      }
      matches.slice(1).forEach(function(m) {
        sheet.getRange(m.row, 8).setValue('Y');
        sheet.getRange(m.row, 9).setValue(nowText);
      });
    } else {
      if (matches.length) {
        matches.forEach(function(m) {
          sheet.getRange(m.row, 8).setValue('Y');
          sheet.getRange(m.row, 9).setValue(nowText);
        });
        action = 'off';
      } else {
        action = 'keep-off';
      }
    }
    SpreadsheetApp.flush();
    return { action: action, matched: matches.length, savedAt: nowText };
  }, { attempts: 4, waitMs: 350, sleepBaseMs: 120 });

  return {
    ok: true,
    favorite: favorite,
    customerNo: customerNo,
    rowNo: rowNo,
    company: company,
    key: key,
    action: writeResult && writeResult.action,
    savedAt: writeResult && writeResult.savedAt
  };
}

function ensurePortalFavoriteSheet_() {
  const ss = getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(PORTAL_FAVORITE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PORTAL_FAVORITE_SHEET_NAME);
    sheet.getRange(1, 1, 1, PORTAL_FAVORITE_HEADERS.length).setValues([PORTAL_FAVORITE_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, PORTAL_FAVORITE_HEADERS.length).setFontWeight('bold').setBackground('#f2f4f7');
  }
  const current = sheet.getRange(1, 1, 1, Math.max(PORTAL_FAVORITE_HEADERS.length, sheet.getLastColumn())).getDisplayValues()[0];
  PORTAL_FAVORITE_HEADERS.forEach(function(h, i) {
    if (String(current[i] || '').trim() !== h) sheet.getRange(1, i + 1).setValue(h);
  });
  return sheet;
}

function getPortalFavoriteRows_(options) {
  options = options || {};
  const user = getPortalFavoriteUser_();
  const allowAllUsers = !!options.allUsers && canPortalViewAllFavoritesP549_();
  const sheet = ensurePortalFavoriteSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_FAVORITE_HEADERS.length).getDisplayValues();
  return values.map(function(r) {
    return {
      id: String(r[0] || '').trim(),
      userKey: String(r[1] || '').trim(),
      userLabel: String(r[2] || '').trim(),
      customerNo: String(r[3] || '').trim(),
      rowNo: Number(r[4]) || 0,
      company: String(r[5] || '').trim(),
      addedAt: String(r[6] || '').trim(),
      deleted: String(r[7] || '').trim(),
      deletedAt: String(r[8] || '').trim()
    };
  }).filter(function(r) {
    // P549: allUsers는 서버에서 관리자 권한을 다시 확인한 경우에만 허용합니다.
    if (!allowAllUsers && r.userKey !== user.key) return false;
    if (!options.includeDeleted && String(r.deleted || '').toUpperCase() === 'Y') return false;
    return !!makePortalFavoriteCustomerKey_(r.customerNo, r.rowNo);
  });
}

function makePortalFavoriteCustomerKey_(customerNo, rowNo) {
  const cno = String(customerNo || '').trim();
  if (cno) return 'CNO:' + cno;
  rowNo = Number(rowNo) || 0;
  return rowNo ? ('ROW:' + rowNo) : '';
}

function getPortalFavoriteUser_() {
  let perm = null;
  try { perm = sanitizePortalPermissionForClient_(getPortalCurrentPermission_()); } catch (err) {}
  let email = perm && perm.email ? String(perm.email || '').trim() : '';
  if (!email) { try { email = String(Session.getActiveUser().getEmail() || '').trim(); } catch (err) {} }
  let label = perm && perm.displayName ? String(perm.displayName || '').trim() : '';
  if (!label) { try { label = getCurrentUserLabel_(); } catch (err) {} }
  const key = email || label || 'unknown-user';
  return { key: key, label: label || email || key };
}
