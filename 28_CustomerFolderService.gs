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
  'updatedAt'
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
  'updatedAt'
];

const PORTAL_MY_CUSTOMER_FOLDER_CACHE_PREFIX_P497 = 'MY_CUSTOMER_FOLDER_P497_';
const PORTAL_MY_CUSTOMER_FOLDER_CACHE_TTL_SEC_P497 = 300;
const PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497 = '나의 고객 폴더';

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
    const bundleRowsP499 = readMyCustomerFolderBundleRowsFastP499_(user.key, { noEnsure: true });
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
  const parentFolderId = String(payload.parentFolderId || '').trim();
  if (!folderName) throw new Error('폴더명을 입력해 주세요.');

  return withPortalScriptLockP201_('my-customer-folder-create-p497', function() {
    const folderSheet = ensureMyCustomerFolderSheetP497_();
    const nowText = getMyCustomerFolderNowTextP497_();
    const folders = readMyCustomerFolderRowsP497_(user.key, { includeDeleted: true });

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

    const folderId = makeMyCustomerFolderIdP497_('FOL');
    folderSheet.appendRow([
      folderId,
      user.key,
      parentFolderId,
      folderName,
      maxOrder + 10,
      '',
      nowText,
      nowText
    ]);

    invalidateMyCustomerFolderCacheP497_(user.key);
    try {
      appendPortalActivityLog_({
        actionType: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        screen: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        summary: '폴더 생성: ' + folderName,
        detail: { action: 'createFolder', folderId: folderId, parentFolderId: parentFolderId, folderName: folderName }
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
      if (f.ownerEmail !== user.key) return;
      folders.push(f);
      if (f.folderId === folderId && !f.isDeleted) { targetIdx = idx; target = f; }
    });
    if (!target) throw new Error('수정할 폴더를 찾지 못했습니다.');

    const sameName = folders.some(function(f) {
      return !f.isDeleted && f.folderId !== folderId && String(f.parentFolderId || '') === String(target.parentFolderId || '') && String(f.folderName || '').trim() === folderName;
    });
    if (sameName) throw new Error('같은 위치에 동일한 이름의 폴더가 이미 있습니다.');

    sheet.getRange(targetIdx + 2, 4).setValue(folderName);
    sheet.getRange(targetIdx + 2, 8).setValue(getMyCustomerFolderNowTextP497_());

    invalidateMyCustomerFolderCacheP497_(user.key);
    try {
      appendPortalActivityLog_({
        actionType: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        screen: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        summary: '폴더명 변경: ' + target.folderName + ' → ' + folderName,
        detail: { action: 'renameFolder', folderId: folderId, before: target.folderName, after: folderName }
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

  return withPortalScriptLockP201_('my-customer-folder-delete-p497', function() {
    const folderSheet = ensureMyCustomerFolderSheetP497_();
    const itemSheet = ensureMyCustomerFolderItemSheetP497_();
    const foldersAll = readMyCustomerFolderRowsP497_(user.key, { includeDeleted: true });
    const target = foldersAll.find(function(f) { return f.folderId === folderId && !f.isDeleted; });
    if (!target) throw new Error('삭제할 폴더를 찾지 못했습니다.');

    const deleteIds = collectMyCustomerFolderDescendantIdsP497_(foldersAll.filter(function(f) { return !f.isDeleted; }), folderId);
    const nowText = getMyCustomerFolderNowTextP497_();

    const folderLastRow = folderSheet.getLastRow();
    if (folderLastRow >= 2) {
      const values = folderSheet.getRange(2, 1, folderLastRow - 1, PORTAL_MY_CUSTOMER_FOLDER_HEADERS_P497.length).getDisplayValues();
      values.forEach(function(row, idx) {
        const f = mapMyCustomerFolderRowP497_(row, idx + 2);
        if (f.ownerEmail === user.key && deleteIds.indexOf(f.folderId) >= 0) {
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
        if (item.ownerEmail === user.key && !item.isDeleted && deleteIds.indexOf(item.folderId) >= 0) {
          itemSheet.getRange(idx + 2, 10).setValue('Y');
          itemSheet.getRange(idx + 2, 12).setValue(nowText);
          removedItems += 1;
        }
      });
    }

    invalidateMyCustomerFolderCacheP497_(user.key);
    try {
      appendPortalActivityLog_({
        actionType: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        screen: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        summary: '폴더 삭제: ' + target.folderName,
        detail: { action: 'deleteFolder', folderId: folderId, folderName: target.folderName, deletedFolderCount: deleteIds.length, removedItems: removedItems }
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
        itemSheet.getRange(t.restoreItem.rowNoInSheet, 1, 1, PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497.length).setValues([[
          itemId,
          user.key,
          folderId,
          t.snapshot.customerNo,
          t.snapshot.rowNo || '',
          t.snapshot.company,
          t.snapshot.salesRep,
          sortOrder,
          '',
          '',
          t.restoreItem.createdAt || nowText,
          nowText
        ]]);
        restored += 1;
      } else {
        itemSheet.appendRow([
          itemId,
          user.key,
          folderId,
          t.snapshot.customerNo,
          t.snapshot.rowNo || '',
          t.snapshot.company,
          t.snapshot.salesRep,
          sortOrder,
          '',
          '',
          nowText,
          nowText
        ]);
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
    const nowText = getMyCustomerFolderNowTextP497_();
    let removed = 0;
    let removedItem = null;

    values.forEach(function(row, idx) {
      const item = mapMyCustomerFolderItemRowP497_(row, idx + 2);
      if (item.ownerEmail !== user.key || item.isDeleted) return;
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
    });

    if (!removed) throw new Error('제거할 고객 항목을 찾지 못했습니다. 화면을 새로고침한 뒤 다시 시도해 주세요.');

    invalidateMyCustomerFolderCacheP497_(user.key);
    try {
      appendPortalActivityLog_({
        actionType: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        screen: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        rowNo: removedItem && removedItem.rowNo,
        customerNo: removedItem && removedItem.customerNo,
        company: removedItem && removedItem.companyNameSnapshot,
        summary: '폴더 고객 제거: ' + (removedItem && removedItem.companyNameSnapshot || ''),
        detail: { action: 'removeCustomer', itemId: itemId, folderId: folderId, removed: removed }
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
  const ss = getWebAppDbSpreadsheet_();
  const folderSheet = ss.getSheetByName(PORTAL_MY_CUSTOMER_FOLDER_SHEET_P497) || null;
  const itemSheet = ss.getSheetByName(PORTAL_MY_CUSTOMER_FOLDER_ITEM_SHEET_P497) || null;
  const folders = [];
  const items = [];

  if (folderSheet) {
    const folderLastRow = folderSheet.getLastRow();
    if (folderLastRow >= 2) {
      const values = folderSheet.getRange(2, 1, folderLastRow - 1, PORTAL_MY_CUSTOMER_FOLDER_HEADERS_P497.length).getDisplayValues();
      values.forEach(function(row, idx) {
        const f = mapMyCustomerFolderRowP497_(row, idx + 2);
        if (ownerEmail && f.ownerEmail !== ownerEmail) return;
        if (!options.includeDeleted && f.isDeleted) return;
        if (f.folderId) folders.push(f);
      });
      folders.sort(function(a, b) {
        const ao = Number(a.sortOrder) || 0;
        const bo = Number(b.sortOrder) || 0;
        return (ao - bo) || String(a.folderName || '').localeCompare(String(b.folderName || ''));
      });
    }
  }

  if (itemSheet) {
    const itemLastRow = itemSheet.getLastRow();
    if (itemLastRow >= 2) {
      const values = itemSheet.getRange(2, 1, itemLastRow - 1, PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497.length).getDisplayValues();
      values.forEach(function(row, idx) {
        const item = mapMyCustomerFolderItemRowP497_(row, idx + 2);
        if (ownerEmail && item.ownerEmail !== ownerEmail) return;
        if (!options.includeDeleted && item.isDeleted) return;
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
  return {
    ok: true,
    user: {
      ownerEmail: user.key,
      label: user.label,
      level: user.permission && user.permission.level || '',
      canClassifyAllCustomers: isMyCustomerFolderAdminUserP497_(user.permission)
    },
    folders: Array.isArray(folders) ? folders : [],
    items: Array.isArray(items) ? items : [],
    source: extra.source || (PORTAL_MY_CUSTOMER_FOLDER_SHEET_P497 + ' + ' + PORTAL_MY_CUSTOMER_FOLDER_ITEM_SHEET_P497),
    loadedAt: getMyCustomerFolderNowTextP497_(),
    dbReady: extra.dbReady !== false,
    dbWarning: String(extra.dbWarning || ''),
    dbError: String(extra.dbError || ''),
    fromCache: !!extra.fromCache
  };
}

function getMyCustomerFolderCachedBundleP497_(cacheKey) {
  try {
    const cached = CacheService.getUserCache().get(cacheKey);
    if (!cached) return null;
    const parsed = JSON.parse(cached);
    return parsed && parsed.ok ? parsed : null;
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
    updatedAt: String(row[7] || '').trim()
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
    updatedAt: String(row[11] || '').trim()
  };
}

function getMyCustomerFolderUserP497_() {
  const perm = getPortalCurrentPermission_();
  if (!perm || perm.active === false || String(perm.level || '').toUpperCase() === 'GUEST') {
    throw new Error('나의 고객 폴더를 사용할 권한이 없습니다. 권한_DB 등록 상태를 확인해 주세요.');
  }
  const email = String(perm.email || '').trim();
  const key = email || String(perm.name || '').trim();
  if (!key) throw new Error('현재 사용자를 식별하지 못했습니다. 다시 접속해 주세요.');
  const label = perm.displayName || perm.name || email || key;
  return { key: key, label: label, permission: perm };
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
  return (PORTAL_MY_CUSTOMER_FOLDER_CACHE_PREFIX_P497 + String(ownerEmail || '').trim().toLowerCase().replace(/[^a-z0-9@._+-]/g, '_')).slice(0, 240);
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
    const folder = readMyCustomerFolderRowsP497_(user.key).find(function(f) { return f.folderId === folderId && !f.isDeleted; });
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
      const rowValues = [
        itemId,
        user.key,
        folderId,
        snapshot.customerNo,
        snapshot.rowNo || '',
        snapshot.company,
        snapshot.salesRep,
        sortOrder,
        '',
        '',
        restoreItem && restoreItem.createdAt ? restoreItem.createdAt : nowText,
        nowText
      ];

      const saved = {
        itemId: itemId,
        ownerEmail: user.key,
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

    if (added || restored) invalidateMyCustomerFolderCacheP497_(user.key);
    try {
      if (added || restored || skipped) appendPortalActivityLog_({
        actionType: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        screen: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        summary: '고객 폴더 빠른 추가: ' + folder.folderName + ' / ' + (added + restored) + '건',
        detail: { action: 'addCustomersFastP499', folderId: folderId, folderName: folder.folderName, added: added, restored: restored, skipped: skipped, customers: savedItems.slice(0, 20) }
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
    const folders = readMyCustomerFolderRowsP497_(user.key);
    const toFolder = folders.find(function(f) { return f.folderId === toFolderId && !f.isDeleted; });
    if (!toFolder) throw new Error('대상 폴더를 찾지 못했습니다.');

    const nowText = getMyCustomerFolderNowTextP497_();
    const items = readMyCustomerFolderItemRowsP497_(user.key, { includeDeleted: true });
    const itemIdSet = {};
    itemIds.forEach(function(id) { itemIdSet[id] = true; });
    const activeItems = items.filter(function(item) { return item.ownerEmail === user.key && !item.isDeleted; });
    const sourceItems = activeItems.filter(function(item) { return itemIdSet[item.itemId]; });
    if (!sourceItems.length) throw new Error('복사/이동할 고객 항목을 찾지 못했습니다.');

    const activeByTargetKey = {};
    const deletedByTargetKey = {};
    items.forEach(function(item) {
      if (item.ownerEmail !== user.key) return;
      const key = makeMyCustomerFolderItemKeyP497_(item.folderId, item.customerNo, item.rowNo);
      if (!key) return;
      if (item.isDeleted) deletedByTargetKey[key] = item;
      else activeByTargetKey[key] = item;
    });

    let nextSort = getNextMyCustomerFolderItemSortOrderP497_(items, toFolderId);
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
      if (String(source.folderId || '') === toFolderId) {
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
          ownerEmail: user.key,
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
        itemSheet.getRange(source.rowNoInSheet, 1, 1, PORTAL_MY_CUSTOMER_FOLDER_ITEM_HEADERS_P497.length).setValues([[
          movedItem.itemId,
          movedItem.ownerEmail,
          movedItem.folderId,
          movedItem.customerNo,
          movedItem.rowNo || '',
          movedItem.companyNameSnapshot,
          movedItem.assignedUserSnapshot,
          movedItem.sortOrder,
          movedItem.memo,
          '',
          movedItem.createdAt,
          movedItem.updatedAt
        ]]);
        upsertItems.push(movedItem);
        moved += 1;
        activeByTargetKey[targetKey] = movedItem;
        return;
      }

      const itemId = deletedTarget && deletedTarget.itemId ? deletedTarget.itemId : makeMyCustomerFolderIdP497_('FIT');
      const targetItem = {
        itemId: itemId,
        ownerEmail: user.key,
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
      const rowValues = [
        targetItem.itemId,
        targetItem.ownerEmail,
        targetItem.folderId,
        targetItem.customerNo,
        targetItem.rowNo || '',
        targetItem.companyNameSnapshot,
        targetItem.assignedUserSnapshot,
        targetItem.sortOrder,
        targetItem.memo,
        '',
        targetItem.createdAt,
        targetItem.updatedAt
      ];
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

    if (copied || moved || deletedItemIds.length) invalidateMyCustomerFolderCacheP497_(user.key);
    try {
      appendPortalActivityLog_({
        actionType: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        screen: PORTAL_MY_CUSTOMER_FOLDER_SCREEN_NAME_P497,
        summary: '고객 폴더 ' + (mode === 'move' ? '이동' : '복사') + ': ' + (copied + moved) + '건',
        detail: { action: 'transferItemsP499', mode: mode, toFolderId: toFolderId, copied: copied, moved: moved, skipped: skipped, duplicateSkipped: duplicateSkipped, itemIds: itemIds.slice(0, 50) }
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
