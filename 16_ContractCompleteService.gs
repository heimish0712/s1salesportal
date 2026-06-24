/***************************************
 * S1 Sales Portal - 16_ContractCompleteService.gs
 * v69: 수주확정/계약완료 시트 조회 화면
 * - 원본 위치: MASTER_SPREADSHEET_ID 파일의 `수주확정/계약완료` 시트
 * - 목적: 계약 완료/수주 확정 리스트를 포털 메뉴에서 조회하고 행 펼침 상세를 표시
 ***************************************/

const PORTAL_CONTRACT_COMPLETE_SHEET_NAME_V69 = '수주확정/계약완료';
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
  const sheet = ss.getSheetByName(PORTAL_CONTRACT_COMPLETE_SHEET_NAME_V69);
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


function updateContractCompleteRowV69(payload) {
  payload = payload || {};
  const rowNo = Number(payload.rowNo) || 0;
  const values = payload.values || {};
  if (rowNo < 2) throw new Error('계약종합관리 저장 대상 행 번호가 올바르지 않습니다.');

  const ss = getMasterSpreadsheet_();
  const sheet = ss.getSheetByName(PORTAL_CONTRACT_COMPLETE_SHEET_NAME_V69);
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
  try { CacheService.getScriptCache().remove(PORTAL_CONTRACT_COMPLETE_CACHE_KEY_V69); } catch (err) {}

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

