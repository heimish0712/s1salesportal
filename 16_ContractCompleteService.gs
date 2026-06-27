/***************************************
 * S1 Sales Portal - 16_ContractCompleteService.gs
 * v69: 수주확정/계약완료 시트 조회 화면
 * - 원본 위치: MASTER_SPREADSHEET_ID 파일의 `수주확정/계약완료` 시트
 * - 목적: 계약 완료/수주 확정 리스트를 포털 메뉴에서 조회하고 행 펼침 상세를 표시
 ***************************************/

const PORTAL_CONTRACT_COMPLETE_SHEET_NAME_V69 = '수주확정/계약완료';
const PORTAL_CONTRACT_COMPLETE_SHEET_NAME_FALLBACKS_P250 = ['수주확정/계약완료', '수주확정계약완료'];
const PORTAL_CONTRACT_COMPLETE_CACHE_KEY_V69 = 'PORTAL_CONTRACT_COMPLETE_LIST_V69';
const PORTAL_CONTRACT_COMPLETE_CACHE_SECONDS_V69 = 180;

const PORTAL_CONTRACT_COMPLETE_FIELDS_V69 = [
  { key: 'contractNo', label: '계약번호', headers: ['계약번호'] },
  { key: 'customerNo', label: '고객번호', headers: ['고객번호'] },
  { key: 'contractDate', label: '계약일자', headers: ['계약일자'] },
  { key: 'vendorSentDate', label: '수행사발송일자', headers: ['수행사발송일자', '수행사 발송일자'] },
  { key: 'businessRegSaved', label: '사업자등록증 저장', headers: ['사업자등록증 저장', '사업자등록증저장'] },
  { key: 'orderMailSaved', label: '발주메일 저장', headers: ['발주메일 저장', '발주 메일 저장', '발주메일저장'] },
  { key: 'contractSaved', label: '계약서 저장', headers: ['계약서 저장', '계약서저장'] },
  { key: 'region', label: '지역', headers: ['지역'] },
  { key: 'city', label: '도시', headers: ['도시', '시군구', '시/군'] },
  { key: 'referrer', label: '제보자', headers: ['제보자'] },
  { key: 'contractRep', label: '계약담당자', headers: ['계약담당자', '계약 담당자'] },
  { key: 'company', label: '고객사명', headers: ['고객사명', '회사명', '고객명'] },
  { key: 'contactName', label: '담당자 성함', headers: ['담당자 성함', '담당자성함', '담당자 이름', '담당자'] },
  { key: 'phone', label: '전화번호', headers: ['전화번호', '대표전화번호', '연락처'] },
  { key: 'email', label: '이메일 주소', headers: ['이메일 주소', '이메일주소', '이메일'] },
  { key: 'area', label: '연면적', headers: ['연면적'] },
  { key: 'grade', label: '관리등급', headers: ['관리등급', '관리 등급'] },
  { key: 'contractPrice', label: '계약가', headers: ['계약가', '계약금액', '최종 견적가', '최종견적가'] },
  { key: 'vat', label: 'VAT', headers: ['VAT', '부가세'] },
  { key: 'vendor', label: '수행사', headers: ['수행사'] },
  { key: 'businessNo', label: '사업자등록번호', headers: ['사업자등록번호', '사업자 등록 번호', '사업자번호'] },
  { key: 'representativeName', label: '대표자명', headers: ['대표자명', '대표자'] },
  { key: 'businessType', label: '업종', headers: ['업종', '업태', '종목', '업태/종목'] },
  { key: 'customerAddress', label: '고객사 주소', headers: ['고객사 주소', '고객사주소', '주소'] },
  { key: 'contractPeriod', label: '계약기간', headers: ['계약기간'] },
  { key: 'appointment', label: '비상주선임', headers: ['비상주선임', '비상주 선임', '관리자 선임 여부'] },
  { key: 'maintenance', label: '유지점검', headers: ['유지점검', '유지 점검'] },
  { key: 'performance', label: '성능점검', headers: ['성능점검', '성능 점검'] },
  { key: 'billingMemo', label: '청구 등 메모', headers: ['청구 등 메모', '청구등메모', '메모', '비고'] },
  { key: 'appointmentDue', label: '선임예정일', headers: ['선임예정일', '선임 예정일'] },
  { key: 'firstMaintenanceDue', label: '첫 유지점검예정일', headers: ['첫 유지점검예정일', '첫 유지점검 예정일', '유지점검예정일'] },
  { key: 'performanceDue', label: '성능점검예정일', headers: ['성능점검예정일', '성능점검 예정일'] },
  { key: 'appointmentDone', label: '선임완료여부', headers: ['선임완료여부', '선임 완료 여부'] },
  { key: 'maintenanceDone', label: '유지점검완료여부', headers: ['유지점검완료여부', '유지점검 완료 여부'] },
  { key: 'performanceDone', label: '성능점검완료여부', headers: ['성능점검완료여부', '성능점검 완료 여부'] }
];



function getContractCompleteSheetV69_(ss) {
  ss = ss || getMasterSpreadsheet_();
  const names = (typeof PORTAL_CONTRACT_COMPLETE_SHEET_NAME_FALLBACKS_P250 !== 'undefined' && PORTAL_CONTRACT_COMPLETE_SHEET_NAME_FALLBACKS_P250)
    ? PORTAL_CONTRACT_COMPLETE_SHEET_NAME_FALLBACKS_P250
    : [PORTAL_CONTRACT_COMPLETE_SHEET_NAME_V69, '수주확정계약완료'];
  for (let i = 0; i < names.length; i++) {
    const name = String(names[i] || '').trim();
    if (!name) continue;
    const sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
  }
  return null;
}

function clearContractCompleteCacheV69_() {
  try { CacheService.getScriptCache().remove(PORTAL_CONTRACT_COMPLETE_CACHE_KEY_V69); } catch (err) {}
}

function listContractCompleteRowsV69(options) {
  options = options || {};
  const force = options.force === true;
  const cache = CacheService.getScriptCache();

  if (!force) {
    try {
      const cached = cache.get(PORTAL_CONTRACT_COMPLETE_CACHE_KEY_V69);
      if (cached) {
        const parsed = JSON.parse(cached);
        parsed.fromCache = true;
        return parsed;
      }
    } catch (err) {}
  }

  const ss = getMasterSpreadsheet_();
  const sheet = getContractCompleteSheetV69_(ss);
  if (!sheet) {
    throw new Error('마스터 파일에서 `' + PORTAL_CONTRACT_COMPLETE_SHEET_NAME_V69 + '` 시트를 찾지 못했습니다.');
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return {
      ok: true,
      sheetName: PORTAL_CONTRACT_COMPLETE_SHEET_NAME_V69,
      total: 0,
      rows: [],
      fromCache: false,
      loadedAt: formatContractCompleteLoadedAtV69_(new Date())
    };
  }

  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = values[0].map(function(h) { return String(h || '').trim(); });
  const headerMap = buildContractCompleteHeaderMapV69_(headers);
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const item = buildContractCompleteRowV69_(row, headerMap, i + 1);
    if (!String(item.contractNo || item.customerNo || item.company || '').trim()) continue;
    rows.push(item);
  }

  rows.sort(function(a, b) {
    const ca = parseContractNumberV69_(a.contractNo);
    const cb = parseContractNumberV69_(b.contractNo);
    if (!isNaN(ca) && !isNaN(cb) && ca !== cb) return cb - ca;
    if (!isNaN(ca) && isNaN(cb)) return -1;
    if (isNaN(ca) && !isNaN(cb)) return 1;
    return Number(b.rowNo || 0) - Number(a.rowNo || 0);
  });

  const result = {
    ok: true,
    sheetName: PORTAL_CONTRACT_COMPLETE_SHEET_NAME_V69,
    total: rows.length,
    rows: rows,
    fromCache: false,
    loadedAt: formatContractCompleteLoadedAtV69_(new Date()),
    headers: headers
  };

  try {
    const json = JSON.stringify(result);
    if (json.length < 90000) {
      cache.put(PORTAL_CONTRACT_COMPLETE_CACHE_KEY_V69, json, PORTAL_CONTRACT_COMPLETE_CACHE_SECONDS_V69);
    }
  } catch (err) {}

  return result;
}

function buildContractCompleteHeaderMapV69_(headers) {
  const map = { exact: {}, normalized: {} };
  (headers || []).forEach(function(header, idx) {
    const text = String(header || '').trim();
    if (!text) return;
    map.exact[text] = idx;
    map.normalized[normalizeContractCompleteHeaderV69_(text)] = idx;
  });
  return map;
}

function buildContractCompleteRowV69_(row, headerMap, rowNo) {
  const item = { rowNo: rowNo };
  PORTAL_CONTRACT_COMPLETE_FIELDS_V69.forEach(function(def) {
    item[def.key] = getContractCompleteCellV69_(row, headerMap, def.headers || [def.label]);
  });
  item.regionCity = [item.region, item.city].filter(function(v) { return String(v || '').trim(); }).join(' ');
  item.documentsSavedSummary = buildContractCompleteSavedSummaryV69_(item);
  item.serviceDoneSummary = buildContractCompleteDoneSummaryV69_(item);
  item.searchText = buildContractCompleteSearchTextV69_(item);
  return item;
}

function getContractCompleteCellV69_(row, headerMap, candidates) {
  candidates = Array.isArray(candidates) ? candidates : [candidates];
  for (let i = 0; i < candidates.length; i++) {
    const h = String(candidates[i] || '').trim();
    if (!h) continue;
    if (Object.prototype.hasOwnProperty.call(headerMap.exact, h)) {
      const v = row[headerMap.exact[h]];
      if (String(v || '').trim() !== '') return v;
    }
    const key = normalizeContractCompleteHeaderV69_(h);
    if (Object.prototype.hasOwnProperty.call(headerMap.normalized, key)) {
      const v2 = row[headerMap.normalized[key]];
      if (String(v2 || '').trim() !== '') return v2;
    }
  }
  return '';
}

function normalizeContractCompleteHeaderV69_(header) {
  return String(header || '').replace(/[\s\n\r\t\/·_()（）\-]+/g, '').trim().toLowerCase();
}

function buildContractCompleteSavedSummaryV69_(item) {
  const parts = [];
  parts.push('사업자등록증 ' + (String(item.businessRegSaved || '').trim() || '-'));
  parts.push('발주메일 ' + (String(item.orderMailSaved || '').trim() || '-'));
  parts.push('계약서 ' + (String(item.contractSaved || '').trim() || '-'));
  return parts.join(' / ');
}

function buildContractCompleteDoneSummaryV69_(item) {
  const parts = [];
  parts.push('선임 ' + (String(item.appointmentDone || '').trim() || '-'));
  parts.push('유지 ' + (String(item.maintenanceDone || '').trim() || '-'));
  parts.push('성능 ' + (String(item.performanceDone || '').trim() || '-'));
  return parts.join(' / ');
}

function buildContractCompleteSearchTextV69_(item) {
  const values = [];
  PORTAL_CONTRACT_COMPLETE_FIELDS_V69.forEach(function(def) {
    values.push(item[def.key] || '');
  });
  values.push(item.regionCity || '');
  values.push(item.documentsSavedSummary || '');
  values.push(item.serviceDoneSummary || '');
  return values.join(' ').toLowerCase();
}

function parseContractNumberV69_(value) {
  const raw = String(value == null ? '' : value).replace(/[^0-9.\-]/g, '');
  return raw ? Number(raw) : NaN;
}

function formatContractCompleteLoadedAtV69_(date) {
  return Utilities.formatDate(date || new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}




/**
 * STEP13/P250: 고객 상세검색 펼침 영역의 `발주하기` 버튼에서 호출합니다.
 * - 동일 고객번호가 이미 수주확정/계약완료 시트에 있으면 새 행을 만들지 않고 기존 계약번호를 반환합니다.
 * - 신규 발주 확정 시 계약번호는 현재 시트의 최대 계약번호 + 1로 순차 부여합니다.
 * - 마스터시트 진행현황은 `발주완료`로 보정하여 고객검색/필터와 발주여부 표시가 어긋나지 않게 합니다.
 */
function createPortalCustomerOrderFromSearchP250(payload) {
  payload = payload || {};
  return runPortalContractOrderLockedP250_(function() {
    const target = assertCustomerTarget_({
      rowNo: payload.rowNo,
      customerNo: payload.customerNo
    }, '발주 처리', { readObject: true });
    assertPortalCanAccessCustomerTarget_(target, '발주 처리');

    const ss = getMasterSpreadsheet_();
    const completeSheet = getContractCompleteSheetV69_(ss);
    if (!completeSheet) {
      throw new Error('영업관리대장에서 `' + PORTAL_CONTRACT_COMPLETE_SHEET_NAME_V69 + '` 시트를 찾지 못했습니다.');
    }

    const lastCol = Math.max(completeSheet.getLastColumn(), PORTAL_CONTRACT_COMPLETE_FIELDS_V69.length);
    const headers = completeSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
    const headerMap = buildContractCompleteHeaderMapV69_(headers);
    const customerNo = String(target.customerNo || getMasterFieldValue_(target.obj, 'customerNo') || '').trim();
    const company = String(getCompanyValue_(target.obj) || '').trim();

    const existing = findContractCompleteExistingOrderP250_(completeSheet, headerMap, customerNo, company);
    if (existing && existing.contractNo) {
      markCustomerMasterStatusOrderedP250_(target);
      return buildPortalCustomerOrderResultP250_(target, existing, true);
    }

    const nextContractNo = getNextContractNumberP250_(completeSheet, headerMap);
    const rowObject = buildContractCompleteAppendObjectFromCustomerP250_(target, nextContractNo);
    const targetRow = getNextContractCompleteAppendRowP250_(completeSheet, headerMap);
    if (targetRow > completeSheet.getMaxRows()) completeSheet.insertRowsAfter(completeSheet.getMaxRows(), targetRow - completeSheet.getMaxRows());

    const rowValues = new Array(lastCol).fill('');
    PORTAL_CONTRACT_COMPLETE_FIELDS_V69.forEach(function(def) {
      const idx = getContractCompleteFieldColumnIndexV69_(headerMap, def.key);
      if (idx < 0 || idx >= lastCol) return;
      rowValues[idx] = rowObject[def.key] == null ? '' : rowObject[def.key];
    });

    completeSheet.getRange(targetRow, 1, 1, lastCol).setValues([rowValues]);
    try {
      if (targetRow > 2) completeSheet.getRange(targetRow - 1, 1, 1, lastCol).copyTo(completeSheet.getRange(targetRow, 1, 1, lastCol), { formatOnly: true });
    } catch (styleErr) {}

    markCustomerMasterStatusOrderedP250_(target);
    clearContractCompleteCacheV69_();
    try { markPortalMasterDataChangedP201_('발주확정 customerNo=' + customerNo + ', contractNo=' + nextContractNo); } catch (err) {}

    const displayRow = completeSheet.getRange(targetRow, 1, 1, lastCol).getDisplayValues()[0];
    const appended = buildContractCompleteRowV69_(displayRow, headerMap, targetRow);
    try {
      appendPortalActivityLog_({
        actionType: '발주확정',
        screen: '고객 상세 검색',
        customerNo: customerNo,
        company: company,
        summary: '발주 확정 등록 #' + nextContractNo,
        detail: { contractNo: nextContractNo, masterRowNo: target.rowNo, contractRowNo: targetRow }
      });
    } catch (logErr) {}

    return buildPortalCustomerOrderResultP250_(target, appended, false);
  });
}

function runPortalContractOrderLockedP250_(callback) {
  if (typeof withPortalScriptLockP201_ === 'function') {
    return withPortalScriptLockP201_('contract-order-create', callback, { attempts: 5, waitMs: 900, sleepBaseMs: 220 });
  }
  return callback();
}

function buildPortalCustomerOrderResultP250_(target, orderRow, alreadyExists) {
  orderRow = orderRow || {};
  const orderInfo = {
    exists: true,
    contractNo: String(orderRow.contractNo || '').trim(),
    rowNo: Number(orderRow.rowNo) || 0,
    company: String(orderRow.company || getCompanyValue_(target && target.obj) || '').trim(),
    customerNo: String(orderRow.customerNo || (target && target.customerNo) || '').trim(),
    alreadyExists: !!alreadyExists
  };
  return {
    ok: true,
    alreadyExists: !!alreadyExists,
    customerNo: orderInfo.customerNo,
    rowNo: target && target.rowNo,
    company: orderInfo.company,
    contractNo: orderInfo.contractNo,
    orderInfo: orderInfo,
    message: orderInfo.company + '의 발주번호(계약번호)는 ' + orderInfo.contractNo + '번 입니다.'
  };
}

function findContractCompleteExistingOrderP250_(sheet, headerMap, customerNo, company) {
  customerNo = String(customerNo || '').trim();
  company = normalizeContractOrderCompanyKeyP250_(company);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const item = buildContractCompleteRowV69_(row, headerMap, i + 2);
    const itemCustomerNo = String(item.customerNo || '').trim();
    if (customerNo && itemCustomerNo && itemCustomerNo === customerNo) return item;
    if (!customerNo && company && normalizeContractOrderCompanyKeyP250_(item.company) === company) return item;
  }
  return null;
}

function getNextContractNumberP250_(sheet, headerMap) {
  const idx = getContractCompleteFieldColumnIndexV69_(headerMap, 'contractNo');
  const lastRow = sheet.getLastRow();
  let maxNo = 0;
  if (idx >= 0 && lastRow >= 2) {
    const values = sheet.getRange(2, idx + 1, lastRow - 1, 1).getDisplayValues();
    values.forEach(function(row) {
      const n = parseContractNumberV69_(row && row[0]);
      if (!isNaN(n) && n > maxNo) maxNo = n;
    });
  }
  return maxNo + 1;
}

function getNextContractCompleteAppendRowP250_(sheet, headerMap) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return 2;
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  let lastDataOffset = -1;
  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const item = buildContractCompleteRowV69_(row, headerMap, i + 2);
    if (String(item.contractNo || item.customerNo || item.company || '').trim()) {
      lastDataOffset = i;
      break;
    }
  }
  return lastDataOffset >= 0 ? lastDataOffset + 3 : 2;
}

function buildContractCompleteAppendObjectFromCustomerP250_(target, contractNo) {
  const obj = target && target.obj ? target.obj : {};
  const address = getMasterFieldValue_(obj, 'address');
  const parsed = parseContractOrderRegionCityP250_(address, getMasterFieldValue_(obj, 'region'));
  const appointment = getMasterFieldValue_(obj, 'appointment');
  const contractUnit = getMasterFieldValue_(obj, 'contractUnit');
  const result = {
    contractNo: contractNo,
    customerNo: getMasterFieldValue_(obj, 'customerNo'),
    contractDate: new Date(),
    vendorSentDate: '',
    businessRegSaved: '',
    orderMailSaved: '',
    contractSaved: '',
    region: parsed.region,
    city: parsed.city,
    referrer: '',
    contractRep: getMasterFieldValue_(obj, 'salesRep'),
    company: getCompanyValue_(obj),
    contactName: getMasterFieldValue_(obj, 'contact'),
    phone: joinUniqueContractOrderValuesP250_([getMasterFieldValue_(obj, 'phone'), getMasterFieldValue_(obj, 'directPhone')], ' / '),
    email: getMasterFieldValue_(obj, 'email'),
    area: toContractOrderNumberOrTextP250_(getMasterFieldValue_(obj, 'area')),
    grade: getMasterFieldValue_(obj, 'grade'),
    contractPrice: toContractOrderNumberOrTextP250_(getMasterFieldValue_(obj, 'finalQuote')),
    vat: getMasterFieldValue_(obj, 'vat'),
    vendor: normalizeContractOrderVendorP250_(getMasterFieldValue_(obj, 'vendor')),
    businessNo: getMasterFieldValue_(obj, 'businessNo'),
    representativeName: getMasterFieldValue_(obj, 'representativeName'),
    businessType: buildContractOrderBusinessTypeP250_(obj),
    customerAddress: address,
    contractPeriod: getContractOrderExplicitPeriodP250_(obj),
    appointment: normalizeContractOrderAppointmentMonthsP250_(appointment, contractUnit),
    maintenance: parseContractOrderCountP250_(getMasterFieldValue_(obj, 'maintenance')),
    performance: parseContractOrderCountP250_(getMasterFieldValue_(obj, 'performance')),
    billingMemo: '',
    appointmentDue: '',
    firstMaintenanceDue: '',
    performanceDue: '',
    appointmentDone: '',
    maintenanceDone: '',
    performanceDone: ''
  };
  return result;
}

function markCustomerMasterStatusOrderedP250_(target) {
  try {
    if (!target || !target.sheet || !target.rowNo) return;
    const sheet = target.sheet;
    const headerMap = getHeaderMap_(sheet);
    const statusCol = findMasterFieldCol_(headerMap, 'status');
    if (!statusCol) return;
    const current = String(sheet.getRange(target.rowNo, statusCol).getDisplayValue() || '').trim();
    if (current === '발주완료' || current === '계약완료') return;
    sheet.getRange(target.rowNo, statusCol).setValue('발주완료');
    if (typeof updateCustomerSearchIndexRowFastByPatch_ === 'function') {
      try { updateCustomerSearchIndexRowFastByPatch_(target.rowNo, target.customerNo, { status: '발주완료' }); } catch (idxErr) {}
    }
  } catch (err) {}
}

function getPortalCustomerOrderInfoP250(payload) {
  payload = payload || {};
  const target = assertCustomerTarget_({
    rowNo: payload.rowNo,
    customerNo: payload.customerNo
  }, '발주여부 조회', { readObject: true });
  assertPortalCanAccessCustomerTarget_(target, '발주여부 조회');
  return getPortalCustomerOrderInfoByTargetP250_(target);
}

function getPortalCustomerOrderInfoByTargetP250_(target) {
  const customerNo = String((target && target.customerNo) || '').trim();
  const company = String((target && target.obj && getCompanyValue_(target.obj)) || '').trim();
  const rows = listContractCompleteRowsV69({ force: false }).rows || [];
  const companyKey = normalizeContractOrderCompanyKeyP250_(company);
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || {};
    if (customerNo && String(row.customerNo || '').trim() === customerNo) {
      return { exists: true, contractNo: String(row.contractNo || '').trim(), rowNo: Number(row.rowNo) || 0, customerNo: customerNo, company: row.company || company };
    }
    if (!customerNo && companyKey && normalizeContractOrderCompanyKeyP250_(row.company) === companyKey) {
      return { exists: true, contractNo: String(row.contractNo || '').trim(), rowNo: Number(row.rowNo) || 0, customerNo: row.customerNo || '', company: row.company || company };
    }
  }
  return { exists: false, contractNo: '', rowNo: 0, customerNo: customerNo, company: company };
}

function normalizeContractOrderCompanyKeyP250_(value) {
  return String(value || '')
    .replace(/\s+/g, '')
    .replace(/주식회사/g, '')
    .replace(/[()（）]/g, '')
    .replace(/㈜/g, '')
    .trim()
    .toLowerCase();
}

function normalizeContractOrderVendorP250_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const norm = raw.replace(/\s+/g, '').toLowerCase();
  if (norm === 'kj' || norm.indexOf('케이제이') >= 0 || norm.indexOf('기술사') >= 0) return '케이제이';
  return raw;
}

function toContractOrderNumberOrTextP250_(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return '';
  const numText = raw.replace(/[,원㎡\s]/g, '');
  if (/^-?\d+(\.\d+)?$/.test(numText)) return Number(numText);
  return raw;
}

function parseContractOrderCountP250_(value) {
  const m = String(value || '').match(/(\d+(?:\.\d+)?)/);
  return m ? Number(m[1]) : '';
}

function normalizeContractOrderAppointmentMonthsP250_(appointment, contractUnit) {
  const app = String(appointment || '').trim();
  if (!app || app === '미선임' || app === '해당없음') return '';
  const months = parseContractOrderCountP250_(contractUnit);
  return months || 12;
}

function buildContractOrderBusinessTypeP250_(obj) {
  const businessType = getValueByHeaderCandidates_(obj, ['업종', '업태/종목', '업태 및 종목']);
  const kind = getValueByHeaderCandidates_(obj, ['업태']);
  const item = getValueByHeaderCandidates_(obj, ['종목']);
  if (businessType) return businessType;
  return joinUniqueContractOrderValuesP250_([kind, item], '/');
}

function getContractOrderExplicitPeriodP250_(obj) {
  return getValueByHeaderCandidates_(obj, ['계약기간', '계약 기간', '계약기간\n(상세 작성 요구한 경우만)']);
}

function joinUniqueContractOrderValuesP250_(values, delimiter) {
  const seen = {};
  const out = [];
  (values || []).forEach(function(value) {
    const text = String(value || '').trim();
    if (!text || seen[text]) return;
    seen[text] = true;
    out.push(text);
  });
  return out.join(delimiter || ' / ');
}

function parseContractOrderRegionCityP250_(address, fallbackRegion) {
  const addr = String(address || '').trim();
  const fb = String(fallbackRegion || '').replace(/권$/,'').trim();
  const tokens = addr.split(/\s+/).filter(Boolean);
  let region = fb || '';
  let city = '';
  if (tokens.length) {
    const first = tokens[0];
    const second = tokens[1] || '';
    if (/서울/.test(first)) { region = '수도권'; city = '서울'; }
    else if (/인천/.test(first)) { region = '수도권'; city = '인천'; }
    else if (/경기/.test(first)) { region = '경기'; city = second.replace(/시|군|구$/,''); }
    else if (/충청|대전|세종/.test(first)) { region = '충청'; city = /대전|세종/.test(first) ? first.replace(/광역시|특별자치시|특별자치도|시$/,'') : second.replace(/시|군|구$/,''); }
    else if (/전라|광주/.test(first)) { region = '호남'; city = /광주/.test(first) ? '광주' : second.replace(/시|군|구$/,''); }
    else if (/경상|부산|울산|대구/.test(first)) { region = '부울경'; city = /부산|울산|대구/.test(first) ? first.replace(/광역시|시$/,'') : second.replace(/시|군|구$/,''); }
    else if (/강원/.test(first)) { region = '강원'; city = second.replace(/시|군|구$/,''); }
    else if (/제주/.test(first)) { region = '제주'; city = second.replace(/시|군|구$/,'') || '제주'; }
  }
  return { region: region, city: city };
}

function updateContractCompleteRowV69(payload) {
  payload = payload || {};
  const rowNo = Number(payload.rowNo) || 0;
  const values = payload.values || {};
  if (rowNo < 2) throw new Error('계약종합관리 저장 대상 행 번호가 올바르지 않습니다.');

  const ss = getMasterSpreadsheet_();
  const sheet = getContractCompleteSheetV69_(ss);
  if (!sheet) throw new Error('마스터 파일에서 `' + PORTAL_CONTRACT_COMPLETE_SHEET_NAME_V69 + '` 시트를 찾지 못했습니다.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (rowNo > lastRow) throw new Error('저장 대상 행을 찾지 못했습니다. rowNo=' + rowNo);

  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const headerMap = buildContractCompleteHeaderMapV69_(headers);
  const rowValues = sheet.getRange(rowNo, 1, 1, lastCol).getValues()[0];

  const editableKeys = {
    appointmentDue: true,
    firstMaintenanceDue: true,
    performanceDue: true,
    appointmentDone: true,
    maintenanceDone: true,
    performanceDone: true,
    businessRegSaved: true,
    orderMailSaved: true,
    contractSaved: true,
    businessNo: true,
    representativeName: true,
    businessType: true,
    billingMemo: true
  };

  Object.keys(values).forEach(function(key) {
    if (!editableKeys[key]) return;
    const idx = getContractCompleteFieldColumnIndexV69_(headerMap, key);
    if (idx < 0 || idx >= lastCol) return;
    rowValues[idx] = values[key] == null ? '' : values[key];
  });

  sheet.getRange(rowNo, 1, 1, lastCol).setValues([rowValues]);
  clearContractCompleteCacheV69_()

  const displayRow = sheet.getRange(rowNo, 1, 1, lastCol).getDisplayValues()[0];
  const updatedRow = buildContractCompleteRowV69_(displayRow, headerMap, rowNo);
  try {
    appendPortalActivityLog_({
      actionType: '계약종합관리',
      screen: '계약종합관리',
      customerNo: updatedRow.customerNo || '',
      company: updatedRow.company || updatedRow.customerName || '',
      summary: '계약종합관리 저장',
      detail: { rowNo: rowNo, values: values }
    });
  } catch (err) {}
  return { ok: true, rowNo: rowNo, updatedRow: updatedRow };
}

function getContractCompleteFieldColumnIndexV69_(headerMap, key) {
  const def = PORTAL_CONTRACT_COMPLETE_FIELDS_V69.find(function(item) { return item && item.key === key; });
  if (!def) return -1;
  const candidates = def.headers || [def.label];
  for (let i = 0; i < candidates.length; i++) {
    const h = String(candidates[i] || '').trim();
    if (!h) continue;
    if (Object.prototype.hasOwnProperty.call(headerMap.exact, h)) return headerMap.exact[h];
    const normalized = normalizeContractCompleteHeaderV69_(h);
    if (Object.prototype.hasOwnProperty.call(headerMap.normalized, normalized)) return headerMap.normalized[normalized];
  }
  return -1;
}

