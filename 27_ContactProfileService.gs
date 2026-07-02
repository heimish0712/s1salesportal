/***************************************
 * S1 Sales Portal - 27_ContactProfileService.gs
 * P484: 고객별 담당자 프로필 DB
 ***************************************/

function ensurePortalContactProfileSheetP484_(ss) {
  ss = ss || getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(PORTAL_CONFIG.CONTACT_PROFILE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PORTAL_CONFIG.CONTACT_PROFILE_SHEET_NAME);
    sheet.getRange(1, 1, 1, PORTAL_CONFIG.CONTACT_PROFILE_HEADERS.length)
      .setValues([PORTAL_CONFIG.CONTACT_PROFILE_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, PORTAL_CONFIG.CONTACT_PROFILE_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#f2f4f7');
    sheet.autoResizeColumns(1, PORTAL_CONFIG.CONTACT_PROFILE_HEADERS.length);
    return sheet;
  }
  ensureSheetHeaders_(sheet, PORTAL_CONFIG.CONTACT_PROFILE_HEADERS);
  return sheet;
}

function normalizePortalContactProfileNameP484_(value) {
  let text = String(value == null ? '' : value).trim();
  if (text.length >= 4 && /님$/.test(text)) text = text.slice(0, -1).trim();
  return text;
}

function normalizePortalContactProfileBooleanP484_(value) {
  if (value === true) return true;
  const text = String(value == null ? '' : value).trim().toUpperCase();
  return text === 'TRUE' || text === 'Y' || text === 'YES' || text === '1' || text === '대표';
}

function makePortalContactProfileIdP484_(customerNo) {
  return 'CP-' + String(customerNo || 'UNKNOWN').replace(/[^0-9A-Za-z가-힣_-]/g, '') + '-' + Utilities.getUuid().slice(0, 8) + '-' + new Date().getTime();
}

function getPortalContactProfileHeaderMapP484_(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), PORTAL_CONFIG.CONTACT_PROFILE_HEADERS.length)).getDisplayValues()[0];
  return buildHeaderMapFromHeaders_(headers, 1);
}

function getPortalContactProfileColP484_(headerMap, header) {
  return findFirstExistingHeaderCol_(headerMap, [header]);
}

function getPortalContactProfileValueP484_(row, headerMap, header) {
  const col = getPortalContactProfileColP484_(headerMap, header);
  return col ? String(row[col - 1] || '').trim() : '';
}

function getContactProfilesForCustomerP484_(customerNo, rowNo) {
  customerNo = (typeof normalizeCustomerNoForKey_ === 'function') ? normalizeCustomerNoForKey_(customerNo || '') : String(customerNo || '').trim();
  rowNo = Number(rowNo) || 0;
  if (!customerNo && !rowNo) return [];
  const sheet = ensurePortalContactProfileSheetP484_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = Math.max(sheet.getLastColumn(), PORTAL_CONFIG.CONTACT_PROFILE_HEADERS.length);
  const headerMap = getPortalContactProfileHeaderMapP484_(sheet);
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  const rows = [];
  values.forEach(function(row, idx) {
    const rowCustomerNo = getPortalContactProfileValueP484_(row, headerMap, '고객번호');
    const rowMaster = Number(String(getPortalContactProfileValueP484_(row, headerMap, '마스터행')).replace(/[^0-9]/g, '')) || 0;
    const active = getPortalContactProfileValueP484_(row, headerMap, '활성여부');
    if (active && !normalizePortalContactProfileBooleanP484_(active)) return;
    if (customerNo && ((typeof normalizeCustomerNoForKey_ === 'function') ? normalizeCustomerNoForKey_(rowCustomerNo) : rowCustomerNo) !== customerNo) return;
    if (!customerNo && rowNo && rowMaster !== rowNo) return;
    rows.push({
      profileId: getPortalContactProfileValueP484_(row, headerMap, '프로필ID'),
      customerNo: rowCustomerNo,
      company: getPortalContactProfileValueP484_(row, headerMap, '회사명'),
      rowNo: rowMaster,
      contactName: normalizePortalContactProfileNameP484_(getPortalContactProfileValueP484_(row, headerMap, '담당자명')),
      phone: getPortalContactProfileValueP484_(row, headerMap, '대표번호'),
      directPhone: getPortalContactProfileValueP484_(row, headerMap, '직통번호'),
      email: getPortalContactProfileValueP484_(row, headerMap, '이메일'),
      representative: normalizePortalContactProfileBooleanP484_(getPortalContactProfileValueP484_(row, headerMap, '대표담당자여부')),
      active: active ? normalizePortalContactProfileBooleanP484_(active) : true,
      createdAt: getPortalContactProfileValueP484_(row, headerMap, '작성일시'),
      createdBy: getPortalContactProfileValueP484_(row, headerMap, '작성자'),
      updatedAt: getPortalContactProfileValueP484_(row, headerMap, '수정일시'),
      updatedBy: getPortalContactProfileValueP484_(row, headerMap, '수정자'),
      __sheetRow: idx + 2
    });
  });
  rows.sort(function(a, b) {
    if (!!a.representative !== !!b.representative) return a.representative ? -1 : 1;
    return Number(a.__sheetRow || 0) - Number(b.__sheetRow || 0);
  });
  return rows;
}

function getCustomerContactProfilesP484(payloadOrRowNo, customerNo) {
  const payload = (payloadOrRowNo && typeof payloadOrRowNo === 'object') ? payloadOrRowNo : { rowNo: payloadOrRowNo, customerNo: customerNo };
  const target = assertCustomerTarget_(payload, '담당자 프로필 조회', { readObject: true });
  assertPortalCanAccessCustomerTarget_(target, '담당자 프로필 조회');
  return {
    ok: true,
    rowNo: target.rowNo,
    customerNo: target.customerNo,
    profiles: getContactProfilesForCustomerP484_(target.customerNo, target.rowNo)
  };
}

function normalizeContactProfilePayloadListP484_(profiles, customerNo, company, rowNo, representativeProfileId) {
  profiles = Array.isArray(profiles) ? profiles : [];
  const seen = {};
  const normalized = [];
  profiles.forEach(function(profile) {
    profile = profile || {};
    let profileId = String(profile.profileId || '').trim();
    const isTemp = !profileId || /^TMP[-_]/i.test(profileId);
    if (isTemp) profileId = makePortalContactProfileIdP484_(customerNo);
    if (seen[profileId]) return;
    seen[profileId] = true;
    const contactName = normalizePortalContactProfileNameP484_(profile.contactName || profile.contact || profile.name || profile['담당자명']);
    const phone = String(profile.phone || profile.mainPhone || profile['대표번호'] || '').trim();
    const directPhone = String(profile.directPhone || profile.mobile || profile['직통번호'] || '').trim();
    const email = String(profile.email || profile['이메일'] || '').trim();
    if (!contactName && !phone && !directPhone && !email) return;
    normalized.push({
      profileId: profileId,
      oldProfileId: String(profile.profileId || '').trim(),
      customerNo: customerNo,
      company: company,
      rowNo: rowNo,
      contactName: contactName,
      phone: phone,
      directPhone: directPhone,
      email: email,
      representative: false,
      active: true
    });
  });
  let repId = String(representativeProfileId || '').trim();
  const tempRep = normalized.find(function(p) { return p.oldProfileId && p.oldProfileId === repId; });
  if (tempRep) repId = tempRep.profileId;
  if (!repId && normalized.length) repId = normalized[0].profileId;
  normalized.forEach(function(p) { p.representative = p.profileId === repId; });
  return { profiles: normalized, representativeProfileId: repId };
}

function saveCustomerContactProfilesP484(payload) {
  payload = payload || {};
  const target = assertCustomerTarget_(payload, '담당자 프로필 저장', { readObject: true });
  assertPortalCanAccessCustomerTarget_(target, '담당자 프로필 저장');
  const rowNo = target.rowNo;
  const customerNo = target.customerNo;
  const company = getCompanyValue_(target.obj || {}) || String(payload.company || '').trim();
  const actor = (typeof getPortalCurrentUserName_ === 'function' ? getPortalCurrentUserName_() : (typeof getCurrentUserLabel_ === 'function' ? getCurrentUserLabel_() : ''));
  const now = new Date();
  const normalized = normalizeContactProfilePayloadListP484_(payload.profiles || [], customerNo, company, rowNo, payload.representativeProfileId);
  const profiles = normalized.profiles;
  if (!profiles.length) return { ok: true, rowNo: rowNo, customerNo: customerNo, profiles: [], representativeProfileId: '', message: '저장할 담당자 프로필이 없습니다.' };

  const sheet = ensurePortalContactProfileSheetP484_();
  const headerMap = getPortalContactProfileHeaderMapP484_(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), PORTAL_CONFIG.CONTACT_PROFILE_HEADERS.length);
  const existing = {};
  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    rows.forEach(function(row, idx) {
      const id = getPortalContactProfileValueP484_(row, headerMap, '프로필ID');
      const rowCustomerNo = getPortalContactProfileValueP484_(row, headerMap, '고객번호');
      if (id && ((typeof normalizeCustomerNoForKey_ === 'function') ? normalizeCustomerNoForKey_(rowCustomerNo) : rowCustomerNo) === customerNo) existing[id] = { rowNo: idx + 2, row: row };
    });
  }

  const col = function(header) { return getPortalContactProfileColP484_(headerMap, header); };
  const writeRow = function(sheetRow, profile, createdAt, createdBy) {
    const row = new Array(lastCol).fill('');
    const set = function(header, value) { const c = col(header); if (c) row[c - 1] = value; };
    set('프로필ID', profile.profileId);
    set('고객번호', customerNo);
    set('회사명', company);
    set('마스터행', rowNo);
    set('담당자명', profile.contactName);
    set('대표번호', profile.phone);
    set('직통번호', profile.directPhone);
    set('이메일', profile.email);
    set('대표담당자여부', profile.representative ? 'TRUE' : 'FALSE');
    set('활성여부', 'TRUE');
    set('작성일시', createdAt || now);
    set('작성자', createdBy || actor);
    set('수정일시', now);
    set('수정자', actor);
    sheet.getRange(sheetRow, 1, 1, lastCol).setValues([row]);
  };

  profiles.forEach(function(profile) {
    const found = existing[profile.profileId];
    if (found) {
      const createdAt = getPortalContactProfileValueP484_(found.row, headerMap, '작성일시') || now;
      const createdBy = getPortalContactProfileValueP484_(found.row, headerMap, '작성자') || actor;
      writeRow(found.rowNo, profile, createdAt, createdBy);
    } else {
      writeRow(sheet.getLastRow() + 1, profile, now, actor);
    }
  });

  // 동일 고객의 기존 활성 프로필 중 이번 payload에 없는 행은 유지하되 대표담당자 여부만 해제합니다.
  const repId = normalized.representativeProfileId;
  const savedIds = {};
  profiles.forEach(function(p) { savedIds[p.profileId] = true; });
  if (lastRow >= 2) {
    const repCol = col('대표담당자여부');
    const customerCol = col('고객번호');
    const profileCol = col('프로필ID');
    if (repCol && customerCol && profileCol) {
      const updatedLastRow = sheet.getLastRow();
      const rows = sheet.getRange(2, 1, updatedLastRow - 1, lastCol).getDisplayValues();
      rows.forEach(function(row, idx) {
        const rowCustomerNo = String(row[customerCol - 1] || '').trim();
        const id = String(row[profileCol - 1] || '').trim();
        if (((typeof normalizeCustomerNoForKey_ === 'function') ? normalizeCustomerNoForKey_(rowCustomerNo) : rowCustomerNo) !== customerNo || savedIds[id]) return;
        if (String(row[repCol - 1] || '').trim().toUpperCase() === 'TRUE') {
          sheet.getRange(idx + 2, repCol).setValue('FALSE');
        }
      });
    }
  }

  SpreadsheetApp.flush();
  return {
    ok: true,
    rowNo: rowNo,
    customerNo: customerNo,
    representativeProfileId: repId,
    profiles: getContactProfilesForCustomerP484_(customerNo, rowNo),
    message: '담당자 프로필을 저장했습니다.'
  };
}
