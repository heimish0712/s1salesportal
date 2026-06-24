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
  const rows = getPortalFavoriteRows_({ includeDeleted: false });
  const map = {};
  rows.forEach(function(r) {
    const key = makePortalFavoriteCustomerKey_(r.customerNo, r.rowNo);
    if (key) map[key] = true;
  });
  return { ok: true, map: map, total: rows.length };
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
  payload = payload || {};
  const user = getPortalFavoriteUser_();
  const customerNo = String(payload.customerNo || '').trim();
  const rowNo = Number(payload.rowNo) || 0;
  const company = String(payload.company || '').trim();
  const favorite = payload.favorite !== false;
  const key = makePortalFavoriteCustomerKey_(customerNo, rowNo);
  if (!key) throw new Error('즐겨찾기할 고객번호 또는 행 정보가 없습니다.');

  const writeResult = withPortalScriptLockP201_('favorite-set', function() {
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
      // 같은 사용자/고객 중복 row는 첫 row만 살리고 나머지는 삭제 처리합니다.
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
    return { action: action, matched: matches.length };
  }, { attempts: 5, waitMs: 500, sleepBaseMs: 180 });

  try {
    appendPortalActivityLog_({
      actionType: '즐겨찾기',
      screen: '고객 즐겨찾기',
      rowNo: rowNo,
      customerNo: customerNo,
      company: company,
      summary: favorite ? '즐겨찾기 추가' : '즐겨찾기 해제',
      detail: { favorite: favorite, action: writeResult && writeResult.action }
    });
  } catch (err) {}

  const mapRes = getPortalCustomerFavoriteMap();
  const listRes = getPortalFavoriteCustomers();
  return {
    ok: true,
    favorite: favorite,
    customerNo: customerNo,
    rowNo: rowNo,
    map: mapRes.map || {},
    favorites: listRes.rows || [],
    action: writeResult && writeResult.action
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
    if (r.userKey !== user.key) return false;
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
