/***************************************
 * S1 Sales Portal - 04_MasterRepository.gs
 * PATCH Q: 마스터시트 헤더 resolver 중앙화
 * 기준일: 2026-06-21
 ***************************************/

function getMasterObjects_() {
  const ss = getMasterSpreadsheet_();
  const sheet = ss.getSheetByName(PORTAL_CONFIG.MASTER_SHEET_NAME);
  if (!sheet) throw new Error('마스터시트(신규)를 찾지 못했습니다.');

  const headerMap = getHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < PORTAL_CONFIG.DATA_START_ROW) return [];

  const values = sheet
    .getRange(PORTAL_CONFIG.DATA_START_ROW, 1, lastRow - PORTAL_CONFIG.DATA_START_ROW + 1, lastCol)
    .getDisplayValues();

  const headers = getHeaderMapDisplayHeaders_(headerMap);

  return values
    .map(function(row, idx) {
      const obj = { __rowNo: PORTAL_CONFIG.DATA_START_ROW + idx };
      headers.forEach(function(header) {
        const col = headerMap[header];
        if (col) obj[header] = row[col - 1] || '';
      });
      return obj;
    })
    .filter(function(obj) { return String(getCompanyValue_(obj) || getMasterFieldValue_(obj, 'customerNo') || '').trim() !== ''; });
}

function readMasterRowObject_(sheet, rowNo) {
  const headerMap = getHeaderMap_(sheet);
  const values = sheet.getRange(rowNo, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const obj = {};
  getHeaderMapDisplayHeaders_(headerMap).forEach(function(header) {
    const col = headerMap[header];
    if (col) obj[header] = values[col - 1] || '';
  });
  return obj;
}

function normalizeMasterHeaderKey_(header) {
  return String(header == null ? '' : header)
    .replace(/[\s\n\r\t]+/g, '')
    .replace(/[()（）\[\]{}·ㆍ_\-\/\\]/g, '')
    .trim()
    .toLowerCase();
}

function attachHeaderMapMeta_(map, displayHeaders, normalizedMap) {
  try {
    Object.defineProperty(map, '__headers', { value: displayHeaders || [], enumerable: false, configurable: true });
    Object.defineProperty(map, '__normalized', { value: normalizedMap || {}, enumerable: false, configurable: true });
  } catch (err) {
    map.__headers = displayHeaders || [];
    map.__normalized = normalizedMap || {};
  }
  return map;
}

function getHeaderMapDisplayHeaders_(headerMap) {
  if (headerMap && Array.isArray(headerMap.__headers)) return headerMap.__headers.slice();
  return Object.keys(headerMap || {}).filter(function(k) { return String(k).indexOf('__') !== 0; });
}

function buildHeaderMapFromHeaders_(headers, baseIndex) {
  const map = {};
  const normalized = {};
  const displayHeaders = [];
  headers = headers || [];
  baseIndex = Number(baseIndex) || 1;

  headers.forEach(function(raw, i) {
    const h = String(raw || '').trim();
    if (!h) return;
    const col = baseIndex === 0 ? i : i + baseIndex;
    if (!map[h]) {
      map[h] = col;
      displayHeaders.push(h);
    }
    const nk = normalizeMasterHeaderKey_(h);
    if (nk && !normalized[nk]) normalized[nk] = col;
  });

  return attachHeaderMapMeta_(map, displayHeaders, normalized);
}

function getHeaderMap_(sheet) {
  const headers = sheet
    .getRange(PORTAL_CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  return buildHeaderMapFromHeaders_(headers, 1);
}

function getSimpleHeaderMapFromRow_(sheet, rowNo) {
  const headers = sheet.getRange(rowNo, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  return buildHeaderMapFromHeaders_(headers, 1);
}

function findFirstExistingHeaderCol_(headerMap, candidates) {
  candidates = candidates || [];
  for (const header of candidates) {
    const h = String(header || '').trim();
    if (!h) continue;
    if (headerMap && headerMap[h]) return headerMap[h];
  }

  const normalized = (headerMap && headerMap.__normalized) || {};
  for (const header of candidates) {
    const nk = normalizeMasterHeaderKey_(header);
    if (nk && normalized[nk]) return normalized[nk];
  }
  return 0;
}

function ensureMasterColumn_(sheet, headerMap, headerName) {
  const col = findFirstExistingHeaderCol_(headerMap, [headerName]);
  if (col) return col;
  const newCol = sheet.getLastColumn() + 1;
  sheet.getRange(PORTAL_CONFIG.HEADER_ROW, newCol).setValue(headerName);
  sheet.getRange(PORTAL_CONFIG.HEADER_ROW, newCol)
    .setFontWeight('bold')
    .setBackground('#f2f4f7');
  return newCol;
}

function getCompanyValue_(obj) {
  return getMasterFieldValue_(obj, 'company');
}

function getMemoValueFromObj_(obj) {
  return getMasterFieldValue_(obj, 'memo');
}

function getStatusValueFromObj_(obj) {
  return getMasterFieldValue_(obj, 'status');
}

function getObjectValueByHeaderCandidate_(obj, candidate) {
  obj = obj || {};
  const h = String(candidate || '').trim();
  if (!h) return '';

  if (Object.prototype.hasOwnProperty.call(obj, h)) {
    const exact = String(obj[h] == null ? '' : obj[h]).trim();
    if (exact) return exact;
  }

  const target = normalizeMasterHeaderKey_(h);
  if (!target) return '';
  const keys = Object.keys(obj);
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (normalizeMasterHeaderKey_(key) !== target) continue;
    const value = String(obj[key] == null ? '' : obj[key]).trim();
    if (value) return value;
  }
  return '';
}

function getValueByHeaderCandidates_(obj, candidates) {
  candidates = candidates || [];
  for (const header of candidates) {
    const value = getObjectValueByHeaderCandidate_(obj, header);
    if (value) return value;
  }
  return '';
}

function cellByHeader_(row, map, header) {
  if (!map) return '';
  let pos = map[header];
  if (pos == null && map.__normalized) pos = map.__normalized[normalizeMasterHeaderKey_(header)];
  if (pos == null) return '';

  // 기존 컨택이력/오늘할일 서비스는 header map을 0-base로 직접 만들고,
  // getHeaderMap_()이 만든 마스터시트 map은 1-base입니다. 둘 다 호환합니다.
  const idx = map.__headers ? Number(pos) - 1 : Number(pos);
  return idx >= 0 ? (row[idx] || '') : '';
}

function cellByHeaderIndex_(row, map, headers) {
  headers = Array.isArray(headers) ? headers : [headers];
  for (const h of headers || []) {
    if (map && map[h] != null) return row[map[h]] || '';
  }
  const normalized = (map && map.__normalized) || {};
  for (const h of headers || []) {
    const idx = normalized[normalizeMasterHeaderKey_(h)];
    if (idx != null) return row[idx] || '';
  }
  return '';
}

function cellByIndexHeader_(row, map, header) {
  let idx = map && Object.prototype.hasOwnProperty.call(map, header) ? map[header] : -1;
  if ((idx == null || idx < 0) && map && map.__normalized) idx = map.__normalized[normalizeMasterHeaderKey_(header)];
  return idx >= 0 ? String(row[idx] || '').trim() : '';
}

function ensureSheetHeaders_(sheet, requiredHeaders) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const current = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h || '').trim());
  const existingMap = buildHeaderMapFromHeaders_(current, 1);

  let col = lastCol;
  requiredHeaders.forEach(function(header) {
    if (!findFirstExistingHeaderCol_(existingMap, [header])) {
      col += 1;
      sheet.getRange(1, col).setValue(header);
      sheet.getRange(1, col).setFontWeight('bold').setBackground('#f2f4f7');
      existingMap[header] = col;
      if (existingMap.__normalized) existingMap.__normalized[normalizeMasterHeaderKey_(header)] = col;
    }
  });
  sheet.setFrozenRows(1);
}

function getMasterHeaderMappingDiagnostics() {
  const sheet = getMasterSheet_();
  const headerMap = getHeaderMap_(sheet);
  const rows = [];
  Object.keys(PORTAL_MASTER_FIELD_SCHEMA || {}).forEach(function(key) {
    const def = PORTAL_MASTER_FIELD_SCHEMA[key] || {};
    const col = findFirstExistingHeaderCol_(headerMap, def.headers || []);
    rows.push({
      key: key,
      label: def.label || key,
      headers: (def.headers || []).join(' | '),
      matchedCol: col,
      matched: !!col
    });
  });
  return { ok: true, sheetName: sheet.getName(), rows: rows };
}

// =========================
// v56 PATCH A: 고객 식별자 resolver
// - 원칙: 고객번호를 기준 식별자로 우선 사용하고, rowNo는 보조 위치값으로만 사용합니다.
// - 1차 패치에서는 기존 rowNo 호출도 호환하되, customerNo+rowNo가 함께 들어오면 불일치를 차단합니다.
// =========================
function normalizeCustomerNoForKey_(value) {
  return String(value == null ? '' : value).trim();
}

function getMasterSheet_() {
  const ss = getMasterSpreadsheet_();
  const sheet = ss.getSheetByName(PORTAL_CONFIG.MASTER_SHEET_NAME);
  if (!sheet) throw new Error('마스터시트(신규)를 찾지 못했습니다.');
  return sheet;
}

function getCustomerNoFromMasterRow_(sheet, rowNo) {
  rowNo = Number(rowNo) || 0;
  if (!sheet || !rowNo || rowNo < PORTAL_CONFIG.DATA_START_ROW || rowNo > sheet.getLastRow()) return '';
  const headerMap = getHeaderMap_(sheet);
  const col = findFirstExistingHeaderCol_(headerMap, ['고객번호']);
  if (!col) return '';
  return normalizeCustomerNoForKey_(sheet.getRange(rowNo, col).getDisplayValue());
}

function findMasterRowNoByCustomerNoSafe_(masterSheet, customerNo) {
  customerNo = normalizeCustomerNoForKey_(customerNo);
  if (!customerNo) return 0;
  masterSheet = masterSheet || getMasterSheet_();
  const headerMap = getHeaderMap_(masterSheet);
  const col = findFirstExistingHeaderCol_(headerMap, ['고객번호']);
  if (!col) return 0;
  const lastRow = masterSheet.getLastRow();
  if (lastRow < PORTAL_CONFIG.DATA_START_ROW) return 0;

  const values = masterSheet
    .getRange(PORTAL_CONFIG.DATA_START_ROW, col, lastRow - PORTAL_CONFIG.DATA_START_ROW + 1, 1)
    .getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    if (normalizeCustomerNoForKey_(values[i][0]) === customerNo) return PORTAL_CONFIG.DATA_START_ROW + i;
  }
  return 0;
}

function resolveCustomerTarget_(payloadOrCustomerNo, fallbackRowNo, options) {
  options = options || {};

  const payload = (payloadOrCustomerNo && typeof payloadOrCustomerNo === 'object')
    ? payloadOrCustomerNo
    : { customerNo: payloadOrCustomerNo, rowNo: fallbackRowNo };

  const actionLabel = String(options.actionLabel || '고객 처리').trim();
  const sheet = options.sheet || getMasterSheet_();
  const inputCustomerNo = normalizeCustomerNoForKey_(payload.customerNo || payload.customerNumber || payload['고객번호']);
  const inputRowNo = Number(payload.rowNo || payload.masterRowNo || fallbackRowNo) || 0;
  const hasValidInputRowNo = inputRowNo >= PORTAL_CONFIG.DATA_START_ROW && inputRowNo <= sheet.getLastRow();

  let resolvedRowNo = 0;
  let resolvedBy = '';
  let rowCustomerNo = '';

  if (inputCustomerNo) {
    resolvedRowNo = findMasterRowNoByCustomerNoSafe_(sheet, inputCustomerNo);
    resolvedBy = 'customerNo';

    if (!resolvedRowNo) {
      throw new Error(actionLabel + ' 대상 고객번호를 마스터시트에서 찾지 못했습니다: ' + inputCustomerNo);
    }

    if (hasValidInputRowNo) {
      rowCustomerNo = getCustomerNoFromMasterRow_(sheet, inputRowNo);
      if (rowCustomerNo && rowCustomerNo !== inputCustomerNo) {
        throw new Error(actionLabel + ' 대상 불일치: 화면 행의 고객번호(' + rowCustomerNo + ')와 요청 고객번호(' + inputCustomerNo + ')가 다릅니다. 새로고침 후 다시 시도하세요.');
      }
    }
  } else if (hasValidInputRowNo) {
    resolvedRowNo = inputRowNo;
    resolvedBy = 'rowNo_fallback';
    rowCustomerNo = getCustomerNoFromMasterRow_(sheet, resolvedRowNo);

    if (options.requireCustomerNo && !rowCustomerNo) {
      throw new Error(actionLabel + ' 대상 고객번호가 없습니다. 고객번호를 먼저 보정하세요.');
    }
  } else {
    throw new Error(actionLabel + ' 대상 고객을 식별하지 못했습니다. 고객번호 또는 행 번호가 필요합니다.');
  }

  const obj = options.readObject === false ? null : readMasterRowObject_(sheet, resolvedRowNo);
  const resolvedCustomerNo = normalizeCustomerNoForKey_(inputCustomerNo || rowCustomerNo || (obj && obj['고객번호']));

  return {
    ok: true,
    sheet: sheet,
    rowNo: resolvedRowNo,
    inputRowNo: inputRowNo,
    customerNo: resolvedCustomerNo,
    inputCustomerNo: inputCustomerNo,
    rowCustomerNo: rowCustomerNo || resolvedCustomerNo,
    resolvedBy: resolvedBy,
    obj: obj
  };
}

function assertCustomerTarget_(payloadOrCustomerNo, actionLabel, options) {
  options = options || {};
  options.actionLabel = actionLabel || options.actionLabel || '고객 처리';
  return resolveCustomerTarget_(payloadOrCustomerNo, null, options);
}
