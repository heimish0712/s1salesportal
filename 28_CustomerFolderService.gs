/***************************************
 * S1 Sales Portal - 28_CustomerFolderService.gs
 * P497: 나의 고객 폴더
 * - 유저별 개인 폴더 트리
 * - 고객분류 DB는 웹앱 시트에 저장
 * - 영업담당자는 본인 담당 고객만 분류 가능
 ***************************************/

const PORTAL_MY_CUSTOMER_FOLDER_SHEET_P497 = '고객분류_폴더_DB';
const PORTAL_MY_CUSTOMER_FOLDER_ITEM_SHEET_P497 = '고객분류_고객_DB';

const PORTAL_MY_CUSTOMER_FOLDER_HEADERS_P497 = [
  'folderId',
  'ownerEmail',
  'parentFolderId',
  'folderName',
  'sortOrder',
  'isDeleted',
  'createdAt',
  'updatedAt',
  'ownerKey',
  'ownerName',
  'loginEmail'
];

const PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497 = [
  'itemId',
  'ownerEmail',
  'folderId',
  'customerNo',
  'rowNo',
  'companyNameSnapshot',
  'assignedUserSnapshot',
  'sortOrder',
  'memo',
  'isDeleted',
  'createdAt',
  'updatedAt',
  'ownerKey',
  'ownerName',
  'loginEmail'
];

const PORTAL_MY_CUSTOMER_FOLDER_CACHE_PREFIX_P497 = 'MY_CUSTOMER_FOLDER_P497_';
const PORTAL_MY_CUSTOMER_FOLDER_CACHE_TTL_SEC_P497 = 300;
const PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497 = '나의 고객 폴더';
// P505: 내 폴더 소유 기준을 로그인 이메일이 아닌 권한_DB 기준 현재 프로필(ownerKey)로 분리
const PORTAL_MY_CUSTOMER_FOLDER_VERSION_P501 = 505;

// P500: 기본 폴더 / 내 폴더 분리
// 기본 자동 폴더는 DB에 고객을 저장하지 않고 고객검색 인덱스/마스터 기준으로 실시간 계산합니다.
const PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500 = '__PORTAL_ADMIN_REQUEST__';
const PORTAL_MY_CUSTOMER_FOLDER_SYSTEM_OWNER_P500 = '__PORTAL_SYSTEM__';
const PORTAL_MY_CUSTOMER_FOLDER_SYSTEM_MY_ALL_P500 = 'SYS_MY_ALL';
const PORTAL_MY_CUSTOMER_FOLDER_SYSTEM_REPS_P500 = ['김경아', '김서하', '이옥희', '최보람'];

function getMyCustomerFolderBundleP497(options) {
  options = options || {};
  const user = getMyCustomerFolderUserP497_();
  const cacheKey = makeMyCustomerFolderCacheKeyP497_(user.key);

  if (!options.force) {
    const cached = getMyCustomerFolderCachedBundleP497_(cacheKey);
    if (cached) return cached;
  }

  try {
    // 조회 경로에서는 DB 시트 생성/헤더 보정을 하지 않습니다.
    // P499: 웹앱 DB Spreadsheet open을 한 번만 수행해 메뉴/팝업 첫 로딩 체감 속도를 줄입니다.
    // 웹앱 DB가 순간적으로 timeout 나더라도 화면이 죽지 않도록 fallback bundle을 반환합니다.
    const bundleRowsP499 = readMyCustomerFolderBundleRowsFastP499_(user.key, { noEnsure: true, user: user });
    const folders = bundleRowsP499.folders;
    const items = bundleRowsP499.items;
    const result = buildMyCustomerFolderBundleP497_(user, folders, items, {
      source: PORTAL_MY_CUSTOMER_FOLDER_SHEET_P497 + ' + ' + PORTAL_MY_CUSTOMER_FOLDER_ITEM_SHEET_P497,
      dbReady: true
    });
    putMyCustomerFolderCachedBundleP497_(cacheKey, result);
    return result;
  } catch (err) {
    const cached = getMyCustomerFolderCachedBundleP497_(cacheKey);
    if (cached) {
      cached.dbReady = false;
      cached.fromCache = true;
      cached.dbWarning = '웹앱 DB 연결이 지연되어 직전 캐시로 표시 중입니다. 새로고침 또는 DB 초기화/복구를 눌러 주세요.';
      cached.dbError = getMyCustomerFolderErrorMessageP497_(err);
      cached.loadedAt = cached.loadedAt || getMyCustomerFolderNowTextP497_();
      return cached;
    }
    return buildMyCustomerFolderBundleP497_(user, [], [], {
      source: 'fallback-empty',
      dbReady: false,
      dbWarning: '웹앱 DB 연결이 지연되어 빈 화면으로 열었습니다. DB 초기화/복구 후 다시 새로고침해 주세요.',
      dbError: getMyCustomerFolderErrorMessageP497_(err)
    });
  }
}

function createMyCustomerFolderP497(payload) {
  payload = payload || {};
  const user = getMyCustomerFolderUserP497_();
  const folderName = normalizeMyCustomerFolderNameP497_(payload.folderName);
  const requestedScope = String(payload.scope || '').trim().toUpperCase();
  const ownerKey = requestedScope === 'ADMIN' ? PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500 : user.key;
  let parentFolderId = String(payload.parentFolderId || '').trim();
  if (!folderName) throw new Error('폴더명을 입력해 주세요.');
  if (ownerKey === PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500 && !isMyCustomerFolderAdminUserP497_(user.permission)) {
    throw new Error('기본 폴더는 관리자/서무만 만들 수 있습니다.');
  }
  if (isMyCustomerFolderSystemFolderIdP500_(parentFolderId)) parentFolderId = '';

  return withPortalScriptLockP201_('my-customer-folder-create-p497', function() {
    const folderSheet = ensureMyCustomerFolderSheetP497_();
    const nowText = getMyCustomerFolderNowTextP497_();
    const folders = readMyCustomerFolderRowsP497_(ownerKey, { includeDeleted: true });

    if (parentFolderId && !folders.some(function(f) { return f.folderId === parentFolderId && !f.isDeleted; })) {
      throw new Error('상위 폴더를 찾지 못했습니다. 화면을 새로고침한 뒤 다시 시도해 주세요.');
    }

    const sameName = folders.some(function(f) {
      return !f.isDeleted && String(f.parentFolderId || '') === parentFolderId && String(f.folderName || '').trim() === folderName;
    });
    if (sameName) throw new Error('같은 위치에 동일한 이름의 폴더가 이미 있습니다.');

    const maxOrder = folders
      .filter(function(f) { return !f.isDeleted && String(f.parentFolderId || '') === parentFolderId; })
      .reduce(function(max, f) { return Math.max(max, Number(f.sortOrder) || 0); }, 0);

    const folderId = makeMyCustomerFolderIdP497_(ownerKey === PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500 ? 'ADF' : 'FOL');
    folderSheet.appendRow(buildMyCustomerFolderFolderRowValuesP505_(user, {
      folderId: folderId,
      ownerEmail: ownerKey,
      parentFolderId: parentFolderId,
      folderName: folderName,
      sortOrder: maxOrder + 10,
      isDeleted: '',
      createdAt: nowText,
      updatedAt: nowText
    }));

    invalidateMyCustomerFolderActorCachesP500_(user, ownerKey);
    try {
      appendPortalActivityLog_({
        actionType: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        screen: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        summary: (ownerKey === PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500 ? '기본 폴더 생성: ' : '폴더 생성: ') + folderName,
        detail: { action: 'createFolder', folderId: folderId, parentFolderId: parentFolderId, folderName: folderName, scope: ownerKey === PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500 ? 'ADMIN' : 'PERSONAL' }
      });
    } catch (err) {}

    return Object.assign({ createdFolderId: folderId }, getMyCustomerFolderBundleP497({ force: true }));
  }, { attempts: 4, waitMs: 500, sleepBaseMs: 120 });
}

function renameMyCustomerFolderP497(payload) {
  payload = payload || {};
  const user = getMyCustomerFolderUserP497_();
  const folderId = String(payload.folderId || '').trim();
  const folderName = normalizeMyCustomerFolderNameP497_(payload.folderName);
  if (!folderId) throw new Error('수정할 폴더가 선택되지 않았습니다.');
  if (!folderName) throw new Error('폴더명을 입력해 주세요.');
  if (isMyCustomerFolderSystemFolderIdP500_(folderId)) throw new Error('기본 자동 폴더명은 수정할 수 없습니다.');

  return withPortalScriptLockP201_('my-customer-folder-rename-p497', function() {
    const sheet = ensureMyCustomerFolderSheetP497_();
    const width = PORTAL_MY_CUSTOMER_FOLDER_HEADERS_P497.length;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('폴더 정보를 찾지 못했습니다.');

    const values = sheet.getRange(2, 1, lastRow - 1, width).getDisplayValues();
    let targetIdx = -1;
    let target = null;
    const folders = [];
    values.forEach(function(row, idx) {
      const f = mapMyCustomerFolderRowP497_(row, idx + 2);
      const owner = getMyCustomerFolderRecordOwnerKeyP505_(f);
      if (owner !== user.key && owner !== PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500) return;
      folders.push(f);
      if (f.folderId === folderId && !f.isDeleted) { targetIdx = idx; target = f; }
    });
    if (!target) throw new Error('수정할 폴더를 찾지 못했습니다.');
    if (getMyCustomerFolderRecordOwnerKeyP505_(target) === PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500 && !isMyCustomerFolderAdminUserP497_(user.permission)) throw new Error('기본 폴더는 관리자/서무만 수정할 수 있습니다.');
    if (getMyCustomerFolderRecordOwnerKeyP505_(target) !== user.key && getMyCustomerFolderRecordOwnerKeyP505_(target) !== PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500) throw new Error('본인 폴더만 수정할 수 있습니다.');

    const sameName = folders.some(function(f) {
      return !f.isDeleted && f.folderId !== folderId && f.ownerEmail === target.ownerEmail && String(f.parentFolderId || '') === String(target.parentFolderId || '') && String(f.folderName || '').trim() === folderName;
    });
    if (sameName) throw new Error('같은 위치에 동일한 이름의 폴더가 이미 있습니다.');

    sheet.getRange(targetIdx + 2, 4).setValue(folderName);
    sheet.getRange(targetIdx + 2, 8).setValue(getMyCustomerFolderNowTextP497_());

    invalidateMyCustomerFolderActorCachesP500_(user, target.ownerEmail);
    try {
      appendPortalActivityLog_({
        actionType: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        screen: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        summary: '폴더명 변경: ' + target.folderName + ' → ' + folderName,
        detail: { action: 'renameFolder', folderId: folderId, before: target.folderName, after: folderName, ownerEmail: target.ownerEmail }
      });
    } catch (err) {}

    return Object.assign({ renamedFolderId: folderId }, getMyCustomerFolderBundleP497({ force: true }));
  }, { attempts: 4, waitMs: 500, sleepBaseMs: 120 });
}

function deleteMyCustomerFolderP497(payload) {
  payload = payload || {};
  const user = getMyCustomerFolderUserP497_();
  const folderId = String(payload.folderId || '').trim();
  if (!folderId) throw new Error('삭제할 폴더가 선택되지 않았습니다.');
  if (isMyCustomerFolderSystemFolderIdP500_(folderId)) throw new Error('기본 자동 폴더는 삭제할 수 없습니다.');

  return withPortalScriptLockP201_('my-customer-folder-delete-p497', function() {
    const folderSheet = ensureMyCustomerFolderSheetP497_();
    const itemSheet = ensureMyCustomerFolderItemSheetP497_();
    const foldersAll = readMyCustomerFolderRowsForActorP500_(user, { includeDeleted: true });
    const target = foldersAll.find(function(f) { return f.folderId === folderId && !f.isDeleted; });
    if (!target) throw new Error('삭제할 폴더를 찾지 못했습니다.');
    if (target.folderScope === 'ADMIN' && !isMyCustomerFolderAdminUserP497_(user.permission)) throw new Error('기본 폴더는 관리자/서무만 삭제할 수 있습니다.');
    if (target.folderScope === 'PERSONAL' && getMyCustomerFolderRecordOwnerKeyP505_(target) !== user.key) throw new Error('본인 폴더만 삭제할 수 있습니다.');

    const sameOwnerFolders = foldersAll.filter(function(f) { return f.ownerEmail === target.ownerEmail && !f.isDeleted; });
    const deleteIds = collectMyCustomerFolderDescendantIdsP497_(sameOwnerFolders, folderId);
    const nowText = getMyCustomerFolderNowTextP497_();

    const folderLastRow = folderSheet.getLastRow();
    if (folderLastRow >= 2) {
      const values = folderSheet.getRange(2, 1, folderLastRow - 1, PORTAL_MY_CUSTOMER_FOLDER_HEADERS_P497.length).getDisplayValues();
      values.forEach(function(row, idx) {
        const f = mapMyCustomerFolderRowP497_(row, idx + 2);
        if (f.ownerEmail === target.ownerEmail && deleteIds.indexOf(f.folderId) >= 0) {
          folderSheet.getRange(idx + 2, 6).setValue('Y');
          folderSheet.getRange(idx + 2, 8).setValue(nowText);
        }
      });
    }

    const itemLastRow = itemSheet.getLastRow();
    let removedItems = 0;
    if (itemLastRow >= 2) {
      const itemValues = itemSheet.getRange(2, 1, itemLastRow - 1, PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497.length).getDisplayValues();
      itemValues.forEach(function(row, idx) {
        const item = mapMyCustomerFolderItemRowP497_(row, idx + 2);
        if (item.ownerEmail === target.ownerEmail && !item.isDeleted && deleteIds.indexOf(item.folderId) >= 0) {
          itemSheet.getRange(idx + 2, 10).setValue('Y');
          itemSheet.getRange(idx + 2, 12).setValue(nowText);
          removedItems += 1;
        }
      });
    }

    invalidateMyCustomerFolderActorCachesP500_(user, target.ownerEmail);
    try {
      appendPortalActivityLog_({
        actionType: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        screen: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        summary: '폴더 삭제: ' + target.folderName,
        detail: { action: 'deleteFolder', folderId: folderId, folderName: target.folderName, deletedFolderCount: deleteIds.length, removedItems: removedItems, ownerEmail: target.ownerEmail }
      });
    } catch (err) {}

    return Object.assign({ deletedFolderId: folderId, deletedFolderIds: deleteIds, removedItems: removedItems }, getMyCustomerFolderBundleP497({ force: true }));
  }, { attempts: 4, waitMs: 500, sleepBaseMs: 120 });
}

function addCustomersToMyFolderP497(payload) {
  payload = payload || {};
  const user = getMyCustomerFolderUserP497_();
  const folderId = String(payload.folderId || '').trim();
  const customers = Array.isArray(payload.customers) ? payload.customers : [];
  if (!folderId) throw new Error('고객을 넣을 폴더가 선택되지 않았습니다.');
  if (!customers.length) throw new Error('추가할 고객이 없습니다.');

  return withPortalScriptLockP201_('my-customer-folder-add-items-p497', function() {
    const itemSheet = ensureMyCustomerFolderItemSheetP497_();
    const folder = readMyCustomerFolderRowsP497_(user.key).find(function(f) { return f.folderId === folderId; });
    if (!folder) throw new Error('고객을 넣을 폴더를 찾지 못했습니다.');

    const nowText = getMyCustomerFolderNowTextP497_();
    const existingItems = readMyCustomerFolderItemRowsP497_(user.key, { includeDeleted: true });
    const activeByFolderKey = {};
    const deletedByFolderKey = {};
    existingItems.forEach(function(item) {
      const key = makeMyCustomerFolderItemKeyP497_(item.folderId, item.customerNo, item.rowNo);
      if (!key) return;
      if (item.isDeleted) deletedByFolderKey[key] = item;
      else activeByFolderKey[key] = item;
    });

    let added = 0;
    let restored = 0;
    let skipped = 0;
    const addedRows = [];
    const targets = [];

    customers.forEach(function(c) {
      const target = assertCustomerTarget_({ customerNo: c && c.customerNo, rowNo: c && c.rowNo }, PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497 + ' 고객 추가', { readObject: true });
      assertMyCustomerFolderCanClassifyTargetP497_(target, user.permission);
      const snapshot = buildMyCustomerFolderCustomerSnapshotP497_(target);
      const key = makeMyCustomerFolderItemKeyP497_(folderId, snapshot.customerNo, snapshot.rowNo);
      if (!key || activeByFolderKey[key]) {
        skipped += 1;
        return;
      }
      targets.push({ key: key, target: target, snapshot: snapshot, restoreItem: deletedByFolderKey[key] || null });
    });

    targets.forEach(function(t) {
      const itemId = t.restoreItem && t.restoreItem.itemId ? t.restoreItem.itemId : makeMyCustomerFolderIdP497_('FIT');
      const sortOrder = getNextMyCustomerFolderItemSortOrderP497_(existingItems, folderId);
      if (t.restoreItem && t.restoreItem.rowNoInSheet) {
        itemSheet.getRange(t.restoreItem.rowNoInSheet, 1, 1, PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497.length).setValues([buildMyCustomerFolderItemRowValuesP505_(user, {
          itemId: itemId,
          ownerEmail: user.key,
          folderId: folderId,
          customerNo: t.snapshot.customerNo,
          rowNo: t.snapshot.rowNo || '',
          companyNameSnapshot: t.snapshot.company,
          assignedUserSnapshot: t.snapshot.salesRep,
          sortOrder: sortOrder,
          memo: '',
          isDeleted: '',
          createdAt: t.restoreItem.createdAt || nowText,
          updatedAt: nowText
        })]);
        restored += 1;
      } else {
        itemSheet.appendRow(buildMyCustomerFolderItemRowValuesP505_(user, {
          itemId: itemId,
          ownerEmail: user.key,
          folderId: folderId,
          customerNo: t.snapshot.customerNo,
          rowNo: t.snapshot.rowNo || '',
          companyNameSnapshot: t.snapshot.company,
          assignedUserSnapshot: t.snapshot.salesRep,
          sortOrder: sortOrder,
          memo: '',
          isDeleted: '',
          createdAt: nowText,
          updatedAt: nowText
        }));
        added += 1;
      }
      existingItems.push({ folderId: folderId, sortOrder: sortOrder, itemId: itemId, customerNo: t.snapshot.customerNo, rowNo: t.snapshot.rowNo });
      activeByFolderKey[t.key] = true;
      addedRows.push(Object.assign({ itemId: itemId }, t.snapshot));
    });

    invalidateMyCustomerFolderCacheP497_(user.key);
    try {
      appendPortalActivityLog_({
        actionType: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        screen: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        summary: '고객 폴더 추가: ' + folder.folderName + ' / ' + (added + restored) + '건',
        detail: { action: 'addCustomers', folderId: folderId, folderName: folder.folderName, added: added, restored: restored, skipped: skipped, customers: addedRows.slice(0, 20) }
      });
    } catch (err) {}

    return Object.assign({ added: added, restored: restored, skipped: skipped }, getMyCustomerFolderBundleP497({ force: true }));
  }, { attempts: 4, waitMs: 600, sleepBaseMs: 150 });
}

function removeCustomerFromMyFolderP497(payload) {
  payload = payload || {};
  const user = getMyCustomerFolderUserP497_();
  const itemId = String(payload.itemId || '').trim();
  const folderId = String(payload.folderId || '').trim();
  const customerNo = String(payload.customerNo || '').trim();
  const rowNo = Number(payload.rowNo) || 0;
  if (!itemId && !folderId) throw new Error('제거할 고객 항목을 찾지 못했습니다.');

  return withPortalScriptLockP201_('my-customer-folder-remove-item-p497', function() {
    const sheet = ensureMyCustomerFolderItemSheetP497_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('폴더 고객 목록이 비어 있습니다.');
    const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497.length).getDisplayValues();
    const isAdmin = isMyCustomerFolderAdminUserP497_(user.permission);
    const nowText = getMyCustomerFolderNowTextP497_();
    let removed = 0;
    let removedItem = null;
    let removedOwner = '';

    values.forEach(function(row, idx) {
      const item = mapMyCustomerFolderItemRowP497_(row, idx + 2);
      if (item.isDeleted) return;
      const itemOwner = getMyCustomerFolderRecordOwnerKeyP505_(item);
      const editableOwner = itemOwner === user.key || (itemOwner === PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500 && isAdmin);
      if (!editableOwner) return;
      let hit = false;
      if (itemId) hit = item.itemId === itemId;
      else if (folderId && item.folderId === folderId) {
        const sameCustomerNo = customerNo && String(item.customerNo || '') === customerNo;
        const sameRow = rowNo && Number(item.rowNo) === rowNo;
        hit = !!(sameCustomerNo || sameRow);
      }
      if (!hit) return;
      sheet.getRange(idx + 2, 10).setValue('Y');
      sheet.getRange(idx + 2, 12).setValue(nowText);
      removed += 1;
      removedItem = removedItem || item;
      removedOwner = item.ownerEmail || removedOwner;
    });

    if (!removed) throw new Error('제거할 고객 항목을 찾지 못했습니다. 화면을 새로고침한 뒤 다시 시도해 주세요.');

    invalidateMyCustomerFolderActorCachesP500_(user, removedOwner);
    try {
      appendPortalActivityLog_({
        actionType: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        screen: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        rowNo: removedItem && removedItem.rowNo,
        customerNo: removedItem && removedItem.customerNo,
        company: removedItem && removedItem.companyNameSnapshot,
        summary: '고객 폴더 제거: ' + (removedItem && removedItem.companyNameSnapshot || removedItem && removedItem.customerNo || '') + ' / ' + removed + '건',
        detail: { action: 'removeCustomer', itemId: itemId, folderId: folderId, customerNo: customerNo, rowNo: rowNo, removed: removed, ownerEmail: removedOwner }
      });
    } catch (err) {}

    return Object.assign({ removed: removed }, getMyCustomerFolderBundleP497({ force: true }));
  }, { attempts: 4, waitMs: 500, sleepBaseMs: 120 });
}

function searchMyCustomerFolderCandidatesP497(keyword, limit) {
  keyword = String(keyword || '').trim();
  limit = Math.max(1, Math.min(50, Number(limit) || 20));
  if (!keyword) return { ok: true, rows: [], total: 0, keyword: keyword };
  const perm = getPortalCurrentPermission_();
  const scope = isMyCustomerFolderAdminUserP497_(perm) ? 'ALL' : 'OWN';
  const res = searchCustomersPaged(keyword, 0, limit, 'customerNoDigitsDesc', scope);
  let rows = Array.isArray(res.rows) ? res.rows : [];
  if (!isMyCustomerFolderAdminUserP497_(perm)) {
    rows = rows.filter(function(row) { return canMyCustomerFolderClassifyIndexRowP497_(row, perm); });
  }
  return {
    ok: true,
    rows: rows.slice(0, limit),
    total: rows.length,
    keyword: keyword,
    source: '검색인덱스_DB'
  };
}

function setupMyCustomerFolderDbP497() {
  const folderSheet = ensureMyCustomerFolderSheetP497_();
  const itemSheet = ensureMyCustomerFolderItemSheetP497_();
  return {
    ok: true,
    folderSheetName: folderSheet.getName(),
    itemSheetName: itemSheet.getName(),
    folderHeaders: PORTAL_MY_CUSTOMER_FOLDER_HEADERS_P497,
    itemHeaders: PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497
  };
}

function getExistingMyCustomerFolderSheetP497_() {
  const ss = getWebAppDbSpreadsheet_();
  return ss.getSheetByName(PORTAL_MY_CUSTOMER_FOLDER_SHEET_P497) || null;
}

function getExistingMyCustomerFolderItemSheetP497_() {
  const ss = getWebAppDbSpreadsheet_();
  return ss.getSheetByName(PORTAL_MY_CUSTOMER_FOLDER_ITEM_SHEET_P497) || null;
}


function readMyCustomerFolderBundleRowsFastP499_(ownerEmail, options) {
  options = options || {};
  ownerEmail = String(ownerEmail || '').trim();
  const user = options.user || { key: ownerEmail, permission: getPortalCurrentPermission_() };
  const isAdmin = isMyCustomerFolderAdminUserP497_(user.permission);
  const ss = getWebAppDbSpreadsheet_();
  const folderSheet = ss.getSheetByName(PORTAL_MY_CUSTOMER_FOLDER_SHEET_P497) || null;
  const itemSheet = ss.getSheetByName(PORTAL_MY_CUSTOMER_FOLDER_ITEM_SHEET_P497) || null;
  const folders = [];
  const items = [];

  function includeStoredOwner_(record) {
    const storedOwner = getMyCustomerFolderRecordOwnerKeyP505_(record);
    return storedOwner === ownerEmail || storedOwner === PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500;
  }

  if (folderSheet) {
    const folderLastRow = folderSheet.getLastRow();
    if (folderLastRow >= 2) {
      const values = folderSheet.getRange(2, 1, folderLastRow - 1, PORTAL_MY_CUSTOMER_FOLDER_HEADERS_P497.length).getDisplayValues();
      values.forEach(function(row, idx) {
        const f = mapMyCustomerFolderRowP497_(row, idx + 2);
        if (!includeStoredOwner_(f)) return;
        if (!options.includeDeleted && f.isDeleted) return;
        if (f.folderId) folders.push(f);
      });
      folders.sort(function(a, b) {
        const as = getMyCustomerFolderScopeOrderP500_(a);
        const bs = getMyCustomerFolderScopeOrderP500_(b);
        const ao = Number(a.sortOrder) || 0;
        const bo = Number(b.sortOrder) || 0;
        return (as - bs) || (ao - bo) || String(a.folderName || '').localeCompare(String(b.folderName || ''));
      });
    }
  }

  if (itemSheet) {
    const itemLastRow = itemSheet.getLastRow();
    if (itemLastRow >= 2) {
      const values = itemSheet.getRange(2, 1, itemLastRow - 1, PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497.length).getDisplayValues();
      values.forEach(function(row, idx) {
        const item = mapMyCustomerFolderItemRowP497_(row, idx + 2);
        if (!includeStoredOwner_(item)) return;
        if (!options.includeDeleted && item.isDeleted) return;
        if (!isAdmin && !isMyCustomerFolderStoredItemAllowedForUserP504_(item, user)) return;
        if (item.itemId && item.folderId) items.push(item);
      });
      items.sort(function(a, b) {
        const ao = Number(a.sortOrder) || 0;
        const bo = Number(b.sortOrder) || 0;
        return (ao - bo) || String(b.createdAt || '').localeCompare(String(a.createdAt || ''));
      });
    }
  }

  return { folders: folders, items: items };
}

function buildMyCustomerFolderBundleP497_(user, folders, items, extra) {
  extra = extra || {};
  const isAdmin = isMyCustomerFolderAdminUserP497_(user.permission);
  const decoratedFolders = buildMyCustomerFolderFoldersForBundleP500_(user, folders);
  const decoratedItems = filterMyCustomerFolderItemsForBundleP500_(user, items);
  return {
    ok: true,
    folderVersion: PORTAL_MY_CUSTOMER_FOLDER_VERSION_P501,
    user: {
      ownerEmail: user.key,
      ownerKey: user.ownerKey || user.key,
      ownerName: user.ownerName || user.label,
      loginEmail: user.loginEmail || '',
      label: user.label,
      level: user.permission && user.permission.level || '',
      canClassifyAllCustomers: isAdmin,
      canEditBasicFolders: isAdmin
    },
    folders: decoratedFolders,
    items: decoratedItems,
    systemReps: PORTAL_MY_CUSTOMER_FOLDER_SYSTEM_REPS_P500.slice(),
    source: extra.source || (PORTAL_MY_CUSTOMER_FOLDER_SHEET_P497 + ' + ' + PORTAL_MY_CUSTOMER_FOLDER_ITEM_SHEET_P497),
    loadedAt: getMyCustomerFolderNowTextP497_(),
    dbReady: extra.dbReady !== false,
    dbWarning: String(extra.dbWarning || ''),
    dbError: String(extra.dbError || ''),
    fromCache: !!extra.fromCache
  };
}


function buildMyCustomerFolderFoldersForBundleP500_(user, folders) {
  folders = Array.isArray(folders) ? folders.slice() : [];
  const isAdmin = isMyCustomerFolderAdminUserP497_(user.permission);
  const virtualFolders = buildMyCustomerFolderVirtualFoldersP500_(isAdmin);
  const stored = folders.map(function(f) { return decorateMyCustomerFolderRowP500_(f, user); });
  return virtualFolders.concat(stored).sort(function(a, b) {
    const as = getMyCustomerFolderScopeOrderP500_(a);
    const bs = getMyCustomerFolderScopeOrderP500_(b);
    const ao = Number(a.sortOrder) || 0;
    const bo = Number(b.sortOrder) || 0;
    return (as - bs) || (ao - bo) || String(a.folderName || '').localeCompare(String(b.folderName || ''));
  });
}

function buildMyCustomerFolderVirtualFoldersP500_(isAdmin) {
  const list = [{
    rowNoInSheet: 0,
    folderId: PORTAL_MY_CUSTOMER_FOLDER_SYSTEM_MY_ALL_P500,
    ownerEmail: PORTAL_MY_CUSTOMER_FOLDER_SYSTEM_OWNER_P500,
    parentFolderId: '',
    folderName: '나의 전체 고객',
    sortOrder: 10,
    isDeleted: false,
    createdAt: '',
    updatedAt: '',
    folderScope: 'SYSTEM',
    folderType: 'VIRTUAL_MY_ALL',
    isSystemFolder: true,
    isVirtual: true,
    isEditable: false,
    canDrop: false,
    targetSalesRep: ''
  }];
  if (isAdmin) {
    PORTAL_MY_CUSTOMER_FOLDER_SYSTEM_REPS_P500.forEach(function(name, idx) {
      list.push({
        rowNoInSheet: 0,
        folderId: 'SYS_REP_ALL_' + name,
        ownerEmail: PORTAL_MY_CUSTOMER_FOLDER_SYSTEM_OWNER_P500,
        parentFolderId: '',
        folderName: name + ' 전체 고객',
        sortOrder: 20 + idx * 10,
        isDeleted: false,
        createdAt: '',
        updatedAt: '',
        folderScope: 'SYSTEM',
        folderType: 'VIRTUAL_REP_ALL',
        isSystemFolder: true,
        isVirtual: true,
        isEditable: false,
        canDrop: false,
        targetSalesRep: name
      });
    });
  }
  return list;
}

function decorateMyCustomerFolderRowP500_(folder, user) {
  folder = Object.assign({}, folder || {});
  const isAdmin = isMyCustomerFolderAdminUserP497_(user && user.permission);
  const owner = getMyCustomerFolderRecordOwnerKeyP505_(folder);
  const isAdminFolder = owner === PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500;
  folder.folderScope = isAdminFolder ? 'ADMIN' : 'PERSONAL';
  folder.folderType = isAdminFolder ? 'REQUEST' : 'USER_FOLDER';
  folder.isSystemFolder = false;
  folder.isVirtual = false;
  folder.isEditable = isAdminFolder ? isAdmin : true;
  folder.canDrop = folder.isEditable;
  return folder;
}

function filterMyCustomerFolderItemsForBundleP500_(user, items) {
  items = Array.isArray(items) ? items : [];
  const isAdmin = isMyCustomerFolderAdminUserP497_(user && user.permission);
  return items.filter(function(item) {
    if (!item || item.isDeleted) return false;
    if (getMyCustomerFolderRecordOwnerKeyP505_(item) === user.key) {
      // P504: 같은 이메일 계정이 ADMIN → SALES로 권한 전환된 경우,
      // 과거 개인 폴더에 담긴 타 영업담당 고객이 그대로 노출되지 않도록 담당자 필터를 다시 적용합니다.
      return isAdmin || isMyCustomerFolderStoredItemAllowedForUserP504_(item, user);
    }
    if (getMyCustomerFolderRecordOwnerKeyP505_(item) === PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500) {
      if (isAdmin) return true;
      return isMyCustomerFolderStoredItemAllowedForUserP504_(item, user);
    }
    return false;
  });
}

function isMyCustomerFolderStoredItemAllowedForUserP504_(item, user) {
  user = user || { permission: getPortalCurrentPermission_() };
  if (isMyCustomerFolderAdminUserP497_(user.permission)) return true;
  const pseudo = {
    salesRep: item && item.assignedUserSnapshot,
    salesRepName: item && item.assignedUserSnapshot,
    '영업담당자': item && item.assignedUserSnapshot,
    '견적담당': item && item.assignedUserSnapshot
  };
  return canMyCustomerFolderClassifyIndexRowP497_(pseudo, user.permission);
}

function getMyCustomerFolderScopeOrderP500_(folder) {
  const scope = String(folder && folder.folderScope || (folder && folder.ownerEmail === PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500 ? 'ADMIN' : 'PERSONAL')).toUpperCase();
  if (scope === 'SYSTEM') return 0;
  if (scope === 'ADMIN') return 1;
  return 2;
}

function readMyCustomerFolderRowsForActorP500_(user, options) {
  options = options || {};
  const sheet = options.noEnsure ? getExistingMyCustomerFolderSheetP497_() : ensureMyCustomerFolderSheetP497_();
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const isAdmin = isMyCustomerFolderAdminUserP497_(user && user.permission);
  const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_MY_CUSTOMER_FOLDER_HEADERS_P497.length).getDisplayValues();
  return values.map(function(row, idx) { return mapMyCustomerFolderRowP497_(row, idx + 2); }).filter(function(f) {
    if (!f.folderId) return false;
    if (!options.includeDeleted && f.isDeleted) return false;
    const owner = getMyCustomerFolderRecordOwnerKeyP505_(f);
    if (owner === user.key) return true;
    if (owner === PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500) return true;
    return false;
  }).map(function(f) { return decorateMyCustomerFolderRowP500_(f, user); }).sort(function(a, b) {
    const as = getMyCustomerFolderScopeOrderP500_(a);
    const bs = getMyCustomerFolderScopeOrderP500_(b);
    const ao = Number(a.sortOrder) || 0;
    const bo = Number(b.sortOrder) || 0;
    return (as - bs) || (ao - bo) || String(a.folderName || '').localeCompare(String(b.folderName || ''));
  });
}

function readMyCustomerFolderItemRowsForActorP500_(user, options) {
  options = options || {};
  const sheet = options.noEnsure ? getExistingMyCustomerFolderItemSheetP497_() : ensureMyCustomerFolderItemSheetP497_();
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const isAdmin = isMyCustomerFolderAdminUserP497_(user && user.permission);
  const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497.length).getDisplayValues();
  return values.map(function(row, idx) { return mapMyCustomerFolderItemRowP497_(row, idx + 2); }).filter(function(item) {
    if (!item.itemId || !item.folderId) return false;
    if (!options.includeDeleted && item.isDeleted) return false;
    if (getMyCustomerFolderRecordOwnerKeyP505_(item) === user.key) {
      return isAdmin || isMyCustomerFolderStoredItemAllowedForUserP504_(item, user);
    }
    if (getMyCustomerFolderRecordOwnerKeyP505_(item) === PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500) {
      if (isAdmin) return true;
      return isMyCustomerFolderStoredItemAllowedForUserP504_(item, user);
    }
    return false;
  }).sort(function(a, b) {
    const ao = Number(a.sortOrder) || 0;
    const bo = Number(b.sortOrder) || 0;
    return (ao - bo) || String(b.createdAt || '').localeCompare(String(a.createdAt || ''));
  });
}

function resolveMyCustomerFolderTargetFolderP500_(folderId, user) {
  folderId = String(folderId || '').trim();
  if (!folderId) throw new Error('대상 폴더가 선택되지 않았습니다.');
  if (isMyCustomerFolderSystemFolderIdP500_(folderId)) throw new Error('기본 자동 폴더에는 직접 저장할 수 없습니다. 내 폴더 또는 관리자 요청 폴더를 선택해 주세요.');
  const folders = readMyCustomerFolderRowsForActorP500_(user, { includeDeleted: false });
  const folder = folders.find(function(f) { return f.folderId === folderId && !f.isDeleted; });
  if (!folder) throw new Error('대상 폴더를 찾지 못했습니다.');
  if (folder.folderScope === 'ADMIN' && !isMyCustomerFolderAdminUserP497_(user.permission)) {
    throw new Error('기본 폴더는 관리자/서무만 수정할 수 있습니다.');
  }
  if (folder.folderScope === 'PERSONAL' && getMyCustomerFolderRecordOwnerKeyP505_(folder) !== user.key) {
    throw new Error('본인 폴더만 수정할 수 있습니다.');
  }
  return folder;
}

function isMyCustomerFolderSystemFolderIdP500_(folderId) {
  return String(folderId || '').indexOf('SYS_') === 0;
}

function invalidateMyCustomerFolderActorCachesP500_(user, folderOwner) {
  invalidateMyCustomerFolderCacheP497_(user && user.key);
  if (folderOwner && folderOwner !== (user && user.key)) invalidateMyCustomerFolderCacheP497_(folderOwner);
}

function getMyCustomerFolderCachedBundleP497_(cacheKey) {
  try {
    const cached = CacheService.getUserCache().get(cacheKey);
    if (!cached) return null;
    const parsed = JSON.parse(cached);
    return parsed && parsed.ok && Number(parsed.folderVersion || 0) >= PORTAL_MY_CUSTOMER_FOLDER_VERSION_P501 ? parsed : null;
  } catch (err) {}
  return null;
}

function putMyCustomerFolderCachedBundleP497_(cacheKey, result) {
  try {
    const payload = JSON.stringify(result);
    if (payload.length < 90000) CacheService.getUserCache().put(cacheKey, payload, PORTAL_MY_CUSTOMER_FOLDER_CACHE_TTL_SEC_P497);
  } catch (err) {}
}

function getMyCustomerFolderErrorMessageP497_(err) {
  const msg = err && err.message ? err.message : String(err || '');
  return msg || '알 수 없는 오류';
}

function ensureMyCustomerFolderSheetP497_() {
  const ss = getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(PORTAL_MY_CUSTOMER_FOLDER_SHEET_P497);
  if (!sheet) {
    sheet = ss.insertSheet(PORTAL_MY_CUSTOMER_FOLDER_SHEET_P497);
    sheet.getRange(1, 1, 1, PORTAL_MY_CUSTOMER_FOLDER_HEADERS_P497.length).setValues([PORTAL_MY_CUSTOMER_FOLDER_HEADERS_P497]);
    sheet.setFrozenRows(1);
  }
  ensureMyCustomerFolderHeadersP497_(sheet, PORTAL_MY_CUSTOMER_FOLDER_HEADERS_P497);
  return sheet;
}

function ensureMyCustomerFolderItemSheetP497_() {
  const ss = getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(PORTAL_MY_CUSTOMER_FOLDER_ITEM_SHEET_P497);
  if (!sheet) {
    sheet = ss.insertSheet(PORTAL_MY_CUSTOMER_FOLDER_ITEM_SHEET_P497);
    sheet.getRange(1, 1, 1, PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497.length).setValues([PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497]);
    sheet.setFrozenRows(1);
  }
  ensureMyCustomerFolderHeadersP497_(sheet, PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497);
  return sheet;
}

function ensureMyCustomerFolderHeadersP497_(sheet, headers) {
  const width = Math.max(sheet.getLastColumn() || 1, headers.length);
  const current = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function(v) { return String(v || '').trim(); });
  let changed = false;
  headers.forEach(function(h, i) {
    if (String(current[i] || '') !== h) {
      sheet.getRange(1, i + 1).setValue(h);
      changed = true;
    }
  });
  try { sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f2f4f7'); } catch (err) {}
  if (changed) SpreadsheetApp.flush();
}

function readMyCustomerFolderRowsP497_(ownerEmail, options) {
  options = options || {};
  ownerEmail = String(ownerEmail || '').trim();
  const sheet = options.noEnsure ? getExistingMyCustomerFolderSheetP497_() : ensureMyCustomerFolderSheetP497_();
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_MY_CUSTOMER_FOLDER_HEADERS_P497.length).getDisplayValues();
  return values.map(function(row, idx) {
    return mapMyCustomerFolderRowP497_(row, idx + 2);
  }).filter(function(f) {
    if (ownerEmail && f.ownerEmail !== ownerEmail) return false;
    if (!options.includeDeleted && f.isDeleted) return false;
    return !!f.folderId;
  }).sort(function(a, b) {
    const ao = Number(a.sortOrder) || 0;
    const bo = Number(b.sortOrder) || 0;
    return (ao - bo) || String(a.folderName || '').localeCompare(String(b.folderName || ''));
  });
}

function readMyCustomerFolderItemRowsP497_(ownerEmail, options) {
  options = options || {};
  ownerEmail = String(ownerEmail || '').trim();
  const sheet = options.noEnsure ? getExistingMyCustomerFolderItemSheetP497_() : ensureMyCustomerFolderItemSheetP497_();
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497.length).getDisplayValues();
  return values.map(function(row, idx) {
    return mapMyCustomerFolderItemRowP497_(row, idx + 2);
  }).filter(function(item) {
    if (ownerEmail && item.ownerEmail !== ownerEmail) return false;
    if (!options.includeDeleted && item.isDeleted) return false;
    return !!item.itemId && !!item.folderId;
  }).sort(function(a, b) {
    const ao = Number(a.sortOrder) || 0;
    const bo = Number(b.sortOrder) || 0;
    return (ao - bo) || String(b.createdAt || '').localeCompare(String(a.createdAt || ''));
  });
}

function mapMyCustomerFolderRowP497_(row, rowNoInSheet) {
  return {
    rowNoInSheet: rowNoInSheet || 0,
    folderId: String(row[0] || '').trim(),
    ownerEmail: String(row[1] || '').trim(),
    parentFolderId: String(row[2] || '').trim(),
    folderName: String(row[3] || '').trim(),
    sortOrder: Number(row[4]) || 0,
    isDeleted: String(row[5] || '').trim().toUpperCase() === 'Y',
    createdAt: String(row[6] || '').trim(),
    updatedAt: String(row[7] || '').trim(),
    ownerKey: String(row[8] || '').trim(),
    ownerName: String(row[9] || '').trim(),
    loginEmail: String(row[10] || '').trim()
  };
}

function mapMyCustomerFolderItemRowP497_(row, rowNoInSheet) {
  return {
    rowNoInSheet: rowNoInSheet || 0,
    itemId: String(row[0] || '').trim(),
    ownerEmail: String(row[1] || '').trim(),
    folderId: String(row[2] || '').trim(),
    customerNo: String(row[3] || '').trim(),
    rowNo: Number(row[4]) || 0,
    companyNameSnapshot: String(row[5] || '').trim(),
    assignedUserSnapshot: String(row[6] || '').trim(),
    sortOrder: Number(row[7]) || 0,
    memo: String(row[8] || '').trim(),
    isDeleted: String(row[9] || '').trim().toUpperCase() === 'Y',
    createdAt: String(row[10] || '').trim(),
    updatedAt: String(row[11] || '').trim(),
    ownerKey: String(row[12] || '').trim(),
    ownerName: String(row[13] || '').trim(),
    loginEmail: String(row[14] || '').trim()
  };
}

function getMyCustomerFolderUserP497_() {
  const perm = getPortalCurrentPermission_();
  if (!perm || perm.active === false || String(perm.level || '').toUpperCase() === 'GUEST') {
    throw new Error('나의 고객 폴더를 사용할 권한이 없습니다. 권한_DB 등록 상태를 확인해 주세요.');
  }
  const profile = getMyCustomerFolderEffectiveOwnerProfileP505_(perm);
  const key = profile.ownerKey;
  if (!key) throw new Error('현재 사용자를 식별하지 못했습니다. 다시 접속해 주세요.');
  const label = perm.displayName || perm.name || perm.email || key;
  return {
    key: key,
    ownerKey: key,
    ownerName: profile.ownerName,
    loginEmail: profile.loginEmail,
    legacyOwnerEmail: profile.loginEmail,
    label: label,
    permission: perm
  };
}

function getMyCustomerFolderEffectiveOwnerProfileP505_(perm) {
  perm = perm || getPortalCurrentPermission_();
  const loginEmail = String(perm && perm.email || '').trim().toLowerCase();
  const level = String(perm && perm.level || '').trim().toUpperCase() || 'USER';
  const isSales = level === 'SALES';
  const ownerName = String((isSales && (perm.salesRepName || perm.name)) || perm.name || perm.salesRepName || perm.displayName || loginEmail || '').trim();
  const subject = ownerName || loginEmail || 'UNKNOWN';
  const safeSubject = makeMyCustomerFolderOwnerKeySegmentP505_(subject);
  const safeLevel = makeMyCustomerFolderOwnerKeySegmentP505_(level || 'USER');
  return {
    ownerKey: ('PROFILE:' + safeLevel + ':' + safeSubject).slice(0, 180),
    ownerName: ownerName || subject,
    loginEmail: loginEmail
  };
}

function makeMyCustomerFolderOwnerKeySegmentP505_(value) {
  value = String(value || '').trim();
  if (!value) return 'UNKNOWN';
  try { value = normalizePortalNameForPermission_(value) || value; } catch (err) {}
  value = value.replace(/\s+/g, '');
  value = value.replace(/[^0-9A-Za-z가-힣@._+-]/g, '_');
  return value || 'UNKNOWN';
}

function getMyCustomerFolderRecordOwnerKeyP505_(record) {
  record = record || {};
  return String(record.ownerKey || record.ownerEmail || '').trim();
}

function buildMyCustomerFolderOwnerMetaP505_(user, ownerKey) {
  ownerKey = String(ownerKey || (user && user.key) || '').trim();
  const isAdminOwner = ownerKey === PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500;
  return {
    ownerKey: ownerKey,
    ownerName: isAdminOwner ? '기본 요청 폴더' : String(user && (user.ownerName || user.label) || '').trim(),
    loginEmail: String(user && user.loginEmail || '').trim()
  };
}

function buildMyCustomerFolderFolderRowValuesP505_(user, data) {
  data = data || {};
  const owner = String(data.ownerEmail || data.ownerKey || (user && user.key) || '').trim();
  const meta = buildMyCustomerFolderOwnerMetaP505_(user, owner);
  return [
    data.folderId || '',
    owner,
    data.parentFolderId || '',
    data.folderName || '',
    data.sortOrder || '',
    data.isDeleted || '',
    data.createdAt || '',
    data.updatedAt || '',
    meta.ownerKey,
    meta.ownerName,
    meta.loginEmail
  ];
}

function buildMyCustomerFolderItemRowValuesP505_(user, data) {
  data = data || {};
  const owner = String(data.ownerEmail || data.ownerKey || (user && user.key) || '').trim();
  const meta = buildMyCustomerFolderOwnerMetaP505_(user, owner);
  return [
    data.itemId || '',
    owner,
    data.folderId || '',
    data.customerNo || '',
    Number(data.rowNo) || '',
    data.companyNameSnapshot || '',
    data.assignedUserSnapshot || '',
    data.sortOrder || '',
    data.memo || '',
    data.isDeleted || '',
    data.createdAt || '',
    data.updatedAt || '',
    meta.ownerKey,
    meta.ownerName,
    meta.loginEmail
  ];
}

function isMyCustomerFolderAdminUserP497_(perm) {
  perm = perm || getPortalCurrentPermission_();
  const level = String(perm && perm.level || '').toUpperCase();
  return !!(perm && perm.active !== false && (perm.isAdmin || perm.canUseAdminHome || level === 'ADMIN'));
}

function assertMyCustomerFolderCanClassifyTargetP497_(target, perm) {
  perm = perm || getPortalCurrentPermission_();
  if (isMyCustomerFolderAdminUserP497_(perm)) return true;
  const obj = target && target.obj ? target.obj : null;
  if (!obj) throw new Error('고객 권한 확인에 필요한 마스터 데이터를 읽지 못했습니다.');
  if (canMyCustomerFolderClassifyIndexRowP497_(obj, perm)) return true;
  throw new Error('본인 담당 고객만 나의 고객 폴더에 분류할 수 있습니다.');
}

function canMyCustomerFolderClassifyIndexRowP497_(row, perm) {
  perm = perm || getPortalCurrentPermission_();
  if (isMyCustomerFolderAdminUserP497_(perm)) return true;
  const salesRep = normalizePortalNameForPermission_(getPortalCustomerSalesRepFromRow_(row));
  const aliases = perm.salesRepAliases || splitPortalPermissionAliases_(perm.salesRepName || perm.name || '');
  if (!salesRep || !aliases.length) return false;
  return aliases.some(function(alias) {
    const a = normalizePortalNameForPermission_(alias);
    return a && (salesRep === a || salesRep.indexOf(a) >= 0 || a.indexOf(salesRep) >= 0);
  });
}

function buildMyCustomerFolderCustomerSnapshotP497_(target) {
  const obj = target && target.obj || {};
  return {
    customerNo: String(target && target.customerNo || getCustomerIndexObjectValueK2_(obj, 'customerNo') || obj['고객번호'] || '').trim(),
    rowNo: Number(target && target.rowNo) || Number(obj.__rowNo) || 0,
    company: String(getCompanyValue_(obj) || obj['회사명'] || obj['고객사명'] || '').trim(),
    salesRep: String(getPortalCustomerSalesRepFromRow_(obj) || '').trim()
  };
}

function normalizeMyCustomerFolderNameP497_(value) {
  return String(value || '').replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80);
}

function makeMyCustomerFolderIdP497_(prefix) {
  prefix = String(prefix || 'ID').replace(/[^A-Z0-9]/gi, '').toUpperCase() || 'ID';
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmssSSS');
  return prefix + '-' + stamp + '-' + Utilities.getUuid().slice(0, 8);
}

function getMyCustomerFolderNowTextP497_() {
  try { return getPortalNowTextP201_(new Date()); } catch (err) {}
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function makeMyCustomerFolderCacheKeyP497_(ownerEmail) {
  return (PORTAL_MY_CUSTOMER_FOLDER_CACHE_PREFIX_P497 + 'V' + PORTAL_MY_CUSTOMER_FOLDER_VERSION_P501 + '_' + String(ownerEmail || '').trim().toLowerCase().replace(/[^a-z0-9@._+-]/g, '_')).slice(0, 240);
}

function invalidateMyCustomerFolderCacheP497_(ownerEmail) {
  try { CacheService.getUserCache().remove(makeMyCustomerFolderCacheKeyP497_(ownerEmail)); } catch (err) {}
}

function collectMyCustomerFolderDescendantIdsP497_(folders, rootId) {
  folders = Array.isArray(folders) ? folders : [];
  rootId = String(rootId || '').trim();
  const ids = [];
  const visit = function(id) {
    if (!id || ids.indexOf(id) >= 0) return;
    ids.push(id);
    folders.forEach(function(f) {
      if (String(f.parentFolderId || '') === id) visit(f.folderId);
    });
  };
  visit(rootId);
  return ids;
}

function makeMyCustomerFolderItemKeyP497_(folderId, customerNo, rowNo) {
  folderId = String(folderId || '').trim();
  if (!folderId) return '';
  const cno = String(customerNo || '').trim();
  if (cno) return folderId + '|CNO:' + cno;
  rowNo = Number(rowNo) || 0;
  return rowNo ? (folderId + '|ROW:' + rowNo) : '';
}

function getNextMyCustomerFolderItemSortOrderP497_(items, folderId) {
  items = Array.isArray(items) ? items : [];
  folderId = String(folderId || '').trim();
  const max = items.filter(function(item) {
    return String(item.folderId || '') === folderId && !item.isDeleted;
  }).reduce(function(acc, item) {
    return Math.max(acc, Number(item.sortOrder) || 0);
  }, 0);
  return max + 10;
}


/***************************************
 * P499: 나의 고객 폴더 2차 - 빠른 추가/복사/이동
 * - 고객검색/상세모달 폴더 선택 팝업에서 사용
 * - 저장 성공 후 전체 bundle 재조회 없이 delta만 반환하여 체감 속도 개선
 ***************************************/
function addCustomersToMyFolderFastP499(payload) {
  payload = payload || {};
  const user = getMyCustomerFolderUserP497_();
  const folderId = String(payload.folderId || '').trim();
  const customers = Array.isArray(payload.customers) ? payload.customers : [];
  if (!folderId) throw new Error('고객을 넣을 폴더가 선택되지 않았습니다.');
  if (!customers.length) throw new Error('추가할 고객이 없습니다.');

  return withPortalScriptLockP201_('my-customer-folder-add-fast-p499', function() {
    const itemSheet = ensureMyCustomerFolderItemSheetP497_();
    const folder = resolveMyCustomerFolderTargetFolderP500_(folderId, user);
    const targetOwner = folder.ownerEmail;

    const nowText = getMyCustomerFolderNowTextP497_();
    const existingItems = readMyCustomerFolderItemRowsP497_(targetOwner, { includeDeleted: true });
    const activeByFolderKey = {};
    const deletedByFolderKey = {};
    existingItems.forEach(function(item) {
      const key = makeMyCustomerFolderItemKeyP497_(item.folderId, item.customerNo, item.rowNo);
      if (!key) return;
      if (item.isDeleted) deletedByFolderKey[key] = item;
      else activeByFolderKey[key] = item;
    });

    let added = 0;
    let restored = 0;
    let skipped = 0;
    let nextSort = getNextMyCustomerFolderItemSortOrderP497_(existingItems, folderId);
    const newRows = [];
    const savedItems = [];
    const restoredRowWrites = [];

    customers.forEach(function(c) {
      const target = assertCustomerTarget_({ customerNo: c && c.customerNo, rowNo: c && c.rowNo }, PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497 + ' 고객 추가', { readObject: true });
      assertMyCustomerFolderCanClassifyTargetP497_(target, user.permission);
      const snapshot = buildMyCustomerFolderCustomerSnapshotP497_(target);
      const key = makeMyCustomerFolderItemKeyP497_(folderId, snapshot.customerNo, snapshot.rowNo);
      if (!key || activeByFolderKey[key]) {
        skipped += 1;
        return;
      }

      const restoreItem = deletedByFolderKey[key] || null;
      const itemId = restoreItem && restoreItem.itemId ? restoreItem.itemId : makeMyCustomerFolderIdP497_('FIT');
      const sortOrder = nextSort;
      nextSort += 10;
      const rowValues = buildMyCustomerFolderItemRowValuesP505_(user, {
        itemId: itemId,
        ownerEmail: targetOwner,
        folderId: folderId,
        customerNo: snapshot.customerNo,
        rowNo: snapshot.rowNo || '',
        companyNameSnapshot: snapshot.company,
        assignedUserSnapshot: snapshot.salesRep,
        sortOrder: sortOrder,
        memo: '',
        isDeleted: '',
        createdAt: restoreItem && restoreItem.createdAt ? restoreItem.createdAt : nowText,
        updatedAt: nowText
      });

      const saved = {
        itemId: itemId,
        ownerEmail: targetOwner,
        folderId: folderId,
        customerNo: snapshot.customerNo,
        rowNo: Number(snapshot.rowNo) || 0,
        companyNameSnapshot: snapshot.company,
        assignedUserSnapshot: snapshot.salesRep,
        sortOrder: sortOrder,
        memo: '',
        isDeleted: false,
        createdAt: rowValues[10],
        updatedAt: nowText
      };

      if (restoreItem && restoreItem.rowNoInSheet) {
        restoredRowWrites.push({ rowNoInSheet: restoreItem.rowNoInSheet, rowValues: rowValues });
        restored += 1;
      } else {
        newRows.push(rowValues);
        added += 1;
      }
      savedItems.push(saved);
      activeByFolderKey[key] = saved;
      existingItems.push(saved);
    });

    restoredRowWrites.forEach(function(w) {
      itemSheet.getRange(w.rowNoInSheet, 1, 1, PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497.length).setValues([w.rowValues]);
    });
    if (newRows.length) {
      const startRow = Math.max(2, itemSheet.getLastRow() + 1);
      itemSheet.getRange(startRow, 1, newRows.length, PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497.length).setValues(newRows);
    }

    if (added || restored) invalidateMyCustomerFolderActorCachesP500_(user, targetOwner);
    try {
      if (added || restored || skipped) appendPortalActivityLog_({
        actionType: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        screen: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        summary: '고객 폴더 빠른 추가: ' + folder.folderName + ' / ' + (added + restored) + '건',
        detail: { action: 'addCustomersFastP500', folderId: folderId, folderName: folder.folderName, added: added, restored: restored, skipped: skipped, customers: savedItems.slice(0, 20), ownerEmail: targetOwner }
      });
    } catch (err) {}

    return {
      ok: true,
      fastDelta: true,
      folderId: folderId,
      added: added,
      restored: restored,
      skipped: skipped,
      savedItems: savedItems,
      loadedAt: nowText
    };
  }, { attempts: 2, waitMs: 300, sleepBaseMs: 80 });
}

function transferMyCustomerFolderItemsP499(payload) {
  payload = payload || {};
  const user = getMyCustomerFolderUserP497_();
  const mode = String(payload.mode || '').toLowerCase() === 'move' ? 'move' : 'copy';
  const toFolderId = String(payload.toFolderId || '').trim();
  const itemIds = (Array.isArray(payload.itemIds) ? payload.itemIds : [payload.itemId]).map(function(v) { return String(v || '').trim(); }).filter(Boolean);
  if (!toFolderId) throw new Error('대상 폴더가 선택되지 않았습니다.');
  if (!itemIds.length) throw new Error('복사/이동할 고객이 선택되지 않았습니다.');

  return withPortalScriptLockP201_('my-customer-folder-transfer-p499', function() {
    const itemSheet = ensureMyCustomerFolderItemSheetP497_();
    const toFolder = resolveMyCustomerFolderTargetFolderP500_(toFolderId, user);
    const targetOwner = toFolder.ownerEmail;
    const isAdmin = isMyCustomerFolderAdminUserP497_(user.permission);

    const nowText = getMyCustomerFolderNowTextP497_();
    const items = readMyCustomerFolderItemRowsForActorP500_(user, { includeDeleted: true });
    const itemIdSet = {};
    itemIds.forEach(function(id) { itemIdSet[id] = true; });
    const sourceItems = items.filter(function(item) { return itemIdSet[item.itemId] && !item.isDeleted; });
    if (!sourceItems.length) throw new Error('복사/이동할 고객 항목을 찾지 못했습니다.');

    if (mode === 'move') {
      sourceItems.forEach(function(source) {
        const sourceOwner = getMyCustomerFolderRecordOwnerKeyP505_(source);
        const editableSource = sourceOwner === user.key || (sourceOwner === PORTAL_MY_CUSTOMER_FOLDER_ADMIN_OWNER_P500 && isAdmin);
        if (!editableSource) throw new Error('관리자가 만든 기본 폴더 항목은 이동할 수 없습니다. 내 폴더로 복사만 가능합니다.');
      });
    }

    const targetOwnerItems = readMyCustomerFolderItemRowsP497_(targetOwner, { includeDeleted: true });
    const activeByTargetKey = {};
    const deletedByTargetKey = {};
    targetOwnerItems.forEach(function(item) {
      const key = makeMyCustomerFolderItemKeyP497_(item.folderId, item.customerNo, item.rowNo);
      if (!key) return;
      if (item.isDeleted) deletedByTargetKey[key] = item;
      else activeByTargetKey[key] = item;
    });

    let nextSort = getNextMyCustomerFolderItemSortOrderP497_(targetOwnerItems, toFolderId);
    const newRows = [];
    const upsertItems = [];
    const deletedItemIds = [];
    let copied = 0;
    let moved = 0;
    let skipped = 0;
    let duplicateSkipped = 0;

    function markDeleted(item) {
      if (!item || !item.rowNoInSheet) return;
      itemSheet.getRange(item.rowNoInSheet, 10).setValue('Y');
      itemSheet.getRange(item.rowNoInSheet, 12).setValue(nowText);
      deletedItemIds.push(item.itemId);
    }

    sourceItems.forEach(function(source) {
      if (String(source.folderId || '') === toFolderId && String(source.ownerEmail || '') === targetOwner) {
        skipped += 1;
        return;
      }
      const targetKey = makeMyCustomerFolderItemKeyP497_(toFolderId, source.customerNo, source.rowNo);
      const activeTarget = activeByTargetKey[targetKey];
      const deletedTarget = deletedByTargetKey[targetKey];
      const sortOrder = nextSort;
      nextSort += 10;

      if (activeTarget) {
        duplicateSkipped += 1;
        if (mode === 'move') {
          markDeleted(source);
          moved += 1;
        }
        return;
      }

      if (mode === 'move' && source.rowNoInSheet && !deletedTarget) {
        const movedItem = {
          itemId: source.itemId,
          ownerEmail: targetOwner,
          folderId: toFolderId,
          customerNo: source.customerNo,
          rowNo: Number(source.rowNo) || 0,
          companyNameSnapshot: source.companyNameSnapshot,
          assignedUserSnapshot: source.assignedUserSnapshot,
          sortOrder: sortOrder,
          memo: source.memo || '',
          isDeleted: false,
          createdAt: source.createdAt || nowText,
          updatedAt: nowText
        };
        itemSheet.getRange(source.rowNoInSheet, 1, 1, PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497.length).setValues([buildMyCustomerFolderItemRowValuesP505_(user, {
          itemId: movedItem.itemId,
          ownerEmail: movedItem.ownerEmail,
          folderId: movedItem.folderId,
          customerNo: movedItem.customerNo,
          rowNo: movedItem.rowNo || '',
          companyNameSnapshot: movedItem.companyNameSnapshot,
          assignedUserSnapshot: movedItem.assignedUserSnapshot,
          sortOrder: movedItem.sortOrder,
          memo: movedItem.memo,
          isDeleted: '',
          createdAt: movedItem.createdAt,
          updatedAt: movedItem.updatedAt
        })]);
        upsertItems.push(movedItem);
        moved += 1;
        activeByTargetKey[targetKey] = movedItem;
        return;
      }

      const itemId = deletedTarget && deletedTarget.itemId ? deletedTarget.itemId : makeMyCustomerFolderIdP497_('FIT');
      const targetItem = {
        itemId: itemId,
        ownerEmail: targetOwner,
        folderId: toFolderId,
        customerNo: source.customerNo,
        rowNo: Number(source.rowNo) || 0,
        companyNameSnapshot: source.companyNameSnapshot,
        assignedUserSnapshot: source.assignedUserSnapshot,
        sortOrder: sortOrder,
        memo: source.memo || '',
        isDeleted: false,
        createdAt: deletedTarget && deletedTarget.createdAt ? deletedTarget.createdAt : nowText,
        updatedAt: nowText
      };
      const rowValues = buildMyCustomerFolderItemRowValuesP505_(user, {
        itemId: targetItem.itemId,
        ownerEmail: targetItem.ownerEmail,
        folderId: targetItem.folderId,
        customerNo: targetItem.customerNo,
        rowNo: targetItem.rowNo || '',
        companyNameSnapshot: targetItem.companyNameSnapshot,
        assignedUserSnapshot: targetItem.assignedUserSnapshot,
        sortOrder: targetItem.sortOrder,
        memo: targetItem.memo,
        isDeleted: '',
        createdAt: targetItem.createdAt,
        updatedAt: targetItem.updatedAt
      });
      if (deletedTarget && deletedTarget.rowNoInSheet) {
        itemSheet.getRange(deletedTarget.rowNoInSheet, 1, 1, PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497.length).setValues([rowValues]);
      } else {
        newRows.push(rowValues);
      }
      upsertItems.push(targetItem);
      activeByTargetKey[targetKey] = targetItem;
      if (mode === 'move') {
        markDeleted(source);
        moved += 1;
      } else {
        copied += 1;
      }
    });

    if (newRows.length) {
      const startRow = Math.max(2, itemSheet.getLastRow() + 1);
      itemSheet.getRange(startRow, 1, newRows.length, PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497.length).setValues(newRows);
    }

    if (copied || moved || deletedItemIds.length) invalidateMyCustomerFolderActorCachesP500_(user, targetOwner);
    try {
      appendPortalActivityLog_({
        actionType: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        screen: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        summary: '고객 폴더 ' + (mode === 'move' ? '이동' : '복사') + ': ' + (copied + moved) + '건',
        detail: { action: 'transferItemsP500', mode: mode, toFolderId: toFolderId, copied: copied, moved: moved, skipped: skipped, duplicateSkipped: duplicateSkipped, itemIds: itemIds.slice(0, 50), targetOwner: targetOwner }
      });
    } catch (err) {}

    return {
      ok: true,
      fastDelta: true,
      mode: mode,
      toFolderId: toFolderId,
      copied: copied,
      moved: moved,
      skipped: skipped,
      duplicateSkipped: duplicateSkipped,
      upsertItems: upsertItems,
      deletedItemIds: deletedItemIds,
      loadedAt: nowText
    };
  }, { attempts: 2, waitMs: 300, sleepBaseMs: 80 });
}
