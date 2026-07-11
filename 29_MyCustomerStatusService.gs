/***************************************
 * S1 Sales Portal - 29_MyCustomerStatusService.gs
 * P530: 나의 고객 현황 - 보수적 상태 분석 정책 + 보호 상태 우선
 * - 내 고객 기준: 마스터시트 영업담당자 = 현재 로그인 사용자의 영업담당자명
 * - 마스터시트 원본 메모 + 컨택이력_DB + TM 컨택 내용 + 자료발송 스냅샷을 통합 분석
 * - 수주실패/발주완료/계약완료는 영업담당자 판단 우선 보호 상태로 취급
 ***************************************/

const MY_CUSTOMER_STATUS_ANALYZER_VERSION_P528 = 'P530_CONSERVATIVE_STATUS_PROTECTION_V3';
const MY_CUSTOMER_STATUS_ANALYSIS_SHEET_P529 = '고객현황분석_DB';
const MY_CUSTOMER_STATUS_BLOCKING_STATUSES_P528 = ['수주실패', '발주완료', '계약완료'];
const MY_CUSTOMER_STATUS_ACTIVE_STATUSES_P528 = ['견적제출완료', '장기 추진건', '고객 설득 중', '발주완료', '!!상태지정필요!!'];
const MY_CUSTOMER_STATUS_ANALYSIS_HEADERS_P529 = [
  '분석일시', '분석ID', '고객번호', 'rowNo', '회사명', '영업담당자', '현재상태', '추천상태',
  '판정유형', '신뢰도', '우선순위등급', '불일치여부', '분석소스', '최근연락일', '최근이벤트출처',
  '최근이벤트요약', '강한긍정키워드', '강한부정키워드', '계약진행키워드', '자료발송키워드',
  '장기추진키워드', '데이터누락키워드', '가능성점수', '위험태그', '추천액션', '근거JSON',
  '이벤트JSON', 'memoHash', 'contactHistoryHash', 'sendHash', '분석버전', '검토상태', '검토자', '검토일시', '검토메모',
  '상태보호여부', '상태변경추천허용', '계약완료구분', '타사계약여부', '추천노출여부', '분석주의등급', '보수판정사유'
];

function getMyCustomerStatusDashboardP528(options) {
  options = options || {};
  const started = new Date();
  const perm = getPortalCurrentPermission_();
  const aliases = buildMyCustomerStatusOwnerAliasesP528_(perm);
  const indexData = getCustomerSearchIndexData(perm);
  const sourceRows = Array.isArray(indexData.rows) ? indexData.rows : [];
  const ownRowsRaw = sourceRows.filter(function(row) { return isMyCustomerStatusOwnRowP528_(row, aliases); });
  const now = new Date();

  const masterMap = getMyCustomerMasterRowMapP529_(ownRowsRaw);
  const ownRows = ownRowsRaw.map(function(row) { return enrichMyCustomerStatusRowFromMasterP529_(row, masterMap); });
  const contactMap = getMyCustomerContactHistoryMapP529_(ownRows);

  const analyzed = ownRows.map(function(row) {
    return buildMyCustomerStatusAnalyzedRowP528_(row, now, contactMap);
  }).filter(function(row) { return row && row.rowNo; });

  const statusCounts = {};
  analyzed.forEach(function(row) {
    const key = row.status || '(공란)';
    statusCounts[key] = (statusCounts[key] || 0) + 1;
  });

  const mismatchRows = analyzed.filter(function(row) { return !!row.analysis.mismatch; });
  const recentRows = analyzed.filter(function(row) { return row.analysis.lastContactDays != null && row.analysis.lastContactDays <= 7; })
    .sort(sortMyCustomerStatusByRecentP528_);
  const noContactRows = analyzed.filter(function(row) { return !!row.analysis.longNoContact; })
    .sort(sortMyCustomerStatusByStaleP528_);
  const sentNoFollowRows = analyzed.filter(function(row) { return !!row.analysis.sentNoFollow; })
    .sort(sortMyCustomerStatusByStaleP528_);
  const highPotentialRows = analyzed.filter(function(row) { return row.analysis.potentialScore >= 6; })
    .sort(function(a, b) { return (b.analysis.potentialScore || 0) - (a.analysis.potentialScore || 0); });
  const dataMissingRows = analyzed.filter(function(row) { return !!row.analysis.dataMissing; })
    .sort(function(a, b) { return (b.analysis.missingFields || []).length - (a.analysis.missingFields || []).length; });
  const terminalRows = analyzed.filter(function(row) { return !!row.analysis.terminalCandidate; })
    .sort(sortMyCustomerStatusByRecentP528_);

  const activeCount = analyzed.filter(function(row) { return isMyCustomerStatusActiveP528_(row.status); }).length;
  const needStatusCount = analyzed.filter(function(row) { return isMyCustomerStatusNeedStatusP528_(row.status); }).length;

  let analysisDbResult = { saved: false, rows: 0, sheetName: MY_CUSTOMER_STATUS_ANALYSIS_SHEET_P529 };
  if (options.persist !== false) {
    try {
      analysisDbResult = saveMyCustomerStatusAnalysisRowsP529_(analyzed, perm, aliases, started);
    } catch (err) {
      analysisDbResult = {
        saved: false,
        rows: 0,
        sheetName: MY_CUSTOMER_STATUS_ANALYSIS_SHEET_P529,
        error: err && err.message ? err.message : String(err || '')
      };
      try { Logger.log('고객현황분석_DB 저장 실패: ' + (err && err.stack || err)); } catch (e) {}
    }
  }

  return {
    ok: true,
    version: MY_CUSTOMER_STATUS_ANALYZER_VERSION_P528,
    generatedAt: formatMyCustomerStatusDateTimeP528_(started),
    elapsedMs: new Date().getTime() - started.getTime(),
    owner: {
      name: perm.salesRepName || perm.name || '',
      displayName: perm.displayName || perm.name || perm.email || '',
      email: perm.email || '',
      aliases: aliases
    },
    index: {
      version: indexData.version || '',
      builtAt: indexData.builtAt || '',
      dirty: !!indexData.dirty,
      sourceTotal: sourceRows.length,
      ownTotal: analyzed.length
    },
    analysisDb: analysisDbResult,
    statusOptions: (PORTAL_CONFIG.STATUS_OPTIONS || []).slice(),
    statusCounts: objectToSortedStatusCountArrayP528_(statusCounts),
    cards: {
      total: analyzed.length,
      active: activeCount,
      needStatus: needStatusCount,
      mismatch: mismatchRows.length,
      recent7: recentRows.length,
      noContact14: noContactRows.length,
      sentNoFollow: sentNoFollowRows.length,
      highPotential: highPotentialRows.length,
      dataMissing: dataMissingRows.length,
      terminal: terminalRows.length
    },
    rows: analyzed.slice(0, 1500),
    lists: {
      mismatch: mismatchRows.slice(0, 300),
      recent7: recentRows.slice(0, 300),
      noContact14: noContactRows.slice(0, 300),
      sentNoFollow: sentNoFollowRows.slice(0, 300),
      highPotential: highPotentialRows.slice(0, 300),
      dataMissing: dataMissingRows.slice(0, 300),
      terminal: terminalRows.slice(0, 300)
    }
  };
}

function updateMyCustomerStatusFromAnalysisP528(payload) {
  payload = payload || {};
  const rowNo = Number(payload.rowNo || 0) || 0;
  const customerNo = normalizeCustomerNoForKey_(payload.customerNo || '');
  const newStatus = String(payload.newStatus || '').trim();
  const expectedStatus = String(payload.expectedStatus || '').trim();
  if (!rowNo || !customerNo) throw new Error('고객 행/고객번호가 없어 상태를 수정할 수 없습니다.');
  if (!newStatus) throw new Error('변경할 진행현황을 선택해 주세요.');
  const allowed = (PORTAL_CONFIG.STATUS_OPTIONS || []).map(function(s) { return String(s || '').trim(); });
  if (allowed.indexOf(newStatus) < 0) throw new Error('허용되지 않은 진행현황입니다: ' + newStatus);

  const perm = getPortalCurrentPermission_();
  const aliases = buildMyCustomerStatusOwnerAliasesP528_(perm);
  const detail = getCustomerDetail(rowNo);
  const detailCustomerNo = normalizeCustomerNoForKey_(detail && detail.customerNo || '');
  if (detailCustomerNo !== customerNo) throw new Error('고객번호가 일치하지 않습니다. 화면을 새로고침한 뒤 다시 시도해 주세요.');
  if (!isMyCustomerStatusOwnRowP528_(detail, aliases)) throw new Error('내 담당 고객만 나의 고객 현황에서 상태를 수정할 수 있습니다.');
  const currentStatus = String((detail && detail.status) || '').trim();
  if (isMyCustomerStatusProtectedStatusP530_(currentStatus)) {
    throw new Error('수주실패/발주완료/계약완료 상태는 나의 고객 현황 자동추천으로 변경하지 않습니다. 실제 수정이 필요하면 고객상세에서 직접 확인 후 수정해 주세요.');
  }
  if (expectedStatus && currentStatus !== expectedStatus) {
    throw new Error('현재 진행현황이 화면에 표시된 값과 다릅니다. 현재값: ' + (currentStatus || '(공란)'));
  }

  const res = saveCustomerPatchFastP473({
    rowNo: rowNo,
    customerNo: customerNo,
    values: { status: newStatus },
    expectedValues: { status: currentStatus },
    clientSaveSource: 'myCustomerStatus.analysisStatusApply.P530',
    source: 'myCustomerStatus.analysisStatusApply.P530',
    clientOperationId: String(payload.clientOperationId || makeMyCustomerStatusOperationIdP528_(customerNo, rowNo)),
    thinSave: true,
    fastMode: true,
    noSynchronousRefresh: true
  });
  return Object.assign({}, res || {}, {
    ok: true,
    rowNo: rowNo,
    customerNo: customerNo,
    oldStatus: currentStatus,
    newStatus: newStatus,
    source: 'myCustomerStatus.analysisStatusApply.P530'
  });
}

function buildMyCustomerStatusOwnerAliasesP528_(perm) {
  perm = perm || {};
  const source = [];
  if (perm.salesRepAliases && perm.salesRepAliases.length) {
    perm.salesRepAliases.forEach(function(v) { source.push(v); });
  }
  source.push(perm.salesRepName || '');
  source.push(perm.name || '');
  source.push(perm.displayName || '');
  const seen = {};
  const out = [];
  source.forEach(function(raw) {
    splitPortalPermissionAliases_(raw).forEach(function(part) {
      [part, stripPortalParentheticalNameP528_(part)].forEach(function(v) {
        v = String(v || '').trim();
        if (!v) return;
        const key = normalizeMyCustomerStatusNameP528_(v);
        if (!key || seen[key]) return;
        seen[key] = true;
        out.push(v);
      });
    });
  });
  return out;
}

function stripPortalParentheticalNameP528_(value) {
  return String(value || '').replace(/\([^)]*\)/g, '').replace(/（[^）]*）/g, '').trim();
}

function normalizeMyCustomerStatusNameP528_(value) {
  return String(value || '').replace(/\([^)]*\)/g, '').replace(/（[^）]*）/g, '').replace(/\s+/g, '').trim();
}

function isMyCustomerStatusOwnRowP528_(row, aliases) {
  row = row || {};
  aliases = Array.isArray(aliases) ? aliases : [];
  const salesRep = normalizeMyCustomerStatusNameP528_(getPortalCustomerSalesRepFromRow_(row));
  if (!salesRep || !aliases.length) return false;
  return aliases.some(function(alias) {
    const a = normalizeMyCustomerStatusNameP528_(alias);
    return a && (salesRep === a || salesRep.indexOf(a) >= 0 || a.indexOf(salesRep) >= 0);
  });
}

function getMyCustomerMasterRowMapP529_(ownRows) {
  const map = { byRowNo: {}, byCustomerNo: {} };
  try {
    const needRowNo = {};
    const needCustomerNo = {};
    (ownRows || []).forEach(function(row) {
      if (row && row.rowNo) needRowNo[String(row.rowNo)] = true;
      const cno = normalizeCustomerNoForKey_(row && row.customerNo || '');
      if (cno) needCustomerNo[cno] = true;
    });
    if (!Object.keys(needRowNo).length && !Object.keys(needCustomerNo).length) return map;
    const masterObjects = getMasterObjects_();
    masterObjects.forEach(function(obj) {
      const rn = String(obj.__rowNo || '');
      const cno = normalizeCustomerNoForKey_(getCustomerListValue_(obj, 'customerNo') || obj['고객번호'] || '');
      if (rn && needRowNo[rn]) map.byRowNo[rn] = obj;
      if (cno && needCustomerNo[cno]) map.byCustomerNo[cno] = obj;
    });
  } catch (err) {
    try { Logger.log('나의 고객 현황 마스터 원본 메모 조회 실패: ' + (err && err.stack || err)); } catch (e) {}
  }
  return map;
}

function enrichMyCustomerStatusRowFromMasterP529_(indexRow, masterMap) {
  indexRow = indexRow || {};
  masterMap = masterMap || { byRowNo: {}, byCustomerNo: {} };
  const rn = String(indexRow.rowNo || '');
  const cno = normalizeCustomerNoForKey_(indexRow.customerNo || '');
  const obj = (rn && masterMap.byRowNo[rn]) || (cno && masterMap.byCustomerNo[cno]) || null;
  if (!obj) return indexRow;
  const out = Object.assign({}, indexRow);
  out.__source = 'master+index';
  out.__summary = false;
  out.__masterBacked = true;
  out.customerNo = getCustomerListValue_(obj, 'customerNo') || out.customerNo;
  out.orderNo = getCustomerListValue_(obj, 'orderNo') || out.orderNo;
  out.company = getCustomerListValue_(obj, 'company') || out.company;
  out.salesRep = getCustomerListValue_(obj, 'salesRep') || out.salesRep;
  out.status = getCustomerListValue_(obj, 'status') || out.status;
  out.customerRank = getCustomerListValue_(obj, 'customerRank') || out.customerRank;
  out.contact = getCustomerListValue_(obj, 'contact') || out.contact;
  out.phone = getCustomerListValue_(obj, 'phone') || out.phone;
  out.directPhone = getCustomerListValue_(obj, 'directPhone') || out.directPhone;
  out.email = getCustomerListValue_(obj, 'email') || out.email;
  out.vendor = getCustomerListValue_(obj, 'vendor') || out.vendor;
  out.finalQuote = getCustomerListValue_(obj, 'finalQuote') || out.finalQuote;
  out.memo = getCustomerListValue_(obj, 'memo') || out.memo || '';
  out.fullMemo = out.memo;
  out.address = getCustomerListValue_(obj, 'address') || out.address;
  out.fullAddress = getCustomerListValue_(obj, 'address') || out.fullAddress;
  out.area = getCustomerIndexObjectValueK2_(obj, 'area') || out.area;
  out.grade = getCustomerIndexObjectValueK2_(obj, 'grade') || out.grade;
  out.buildingType = getCustomerIndexObjectValueK2_(obj, 'buildingType') || out.buildingType;
  out.contractUnit = getCustomerIndexObjectValueK2_(obj, 'contractUnit') || out.contractUnit;
  out.appointment = getCustomerIndexObjectValueK2_(obj, 'appointment') || out.appointment;
  out.maintenance = getCustomerIndexObjectValueK2_(obj, 'maintenance') || out.maintenance;
  out.performance = getCustomerIndexObjectValueK2_(obj, 'performance') || out.performance;
  out.vat = getCustomerIndexObjectValueK2_(obj, 'vat') || out.vat;
  out.discountRate = getCustomerIndexObjectValueK2_(obj, 'discountRate') || out.discountRate;
  out.specialTerms = getCustomerIndexObjectValueK2_(obj, 'specialTerms') || out.specialTerms;
  out.lastSent = getCustomerIndexObjectValueK2_(obj, 'lastSent') || out.lastSent;
  out.sendCount = getCustomerIndexObjectValueK2_(obj, ['발송횟수', '발송 횟수']) || out.sendCount;
  out.sendStatus = getCustomerIndexObjectValueK2_(obj, ['발송상태', '발송 상태']) || out.sendStatus;
  out.sentAt = getCustomerIndexObjectValueK2_(obj, 'sentAt') || out.sentAt;
  out.memoInferredStatus = getCustomerIndexObjectValueK2_(obj, ['메모상 추측 상태값', '메모상추측상태값', '메모 추측 상태값']) || out.memoInferredStatus;
  out.statusMatch = getCustomerIndexObjectValueK2_(obj, ['상태값 일치 여부', '상태값일치여부']) || out.statusMatch;
  out.tmProgressStatus = getCustomerIndexObjectValueK2_(obj, 'tmProgressStatus') || out.tmProgressStatus;
  out.tmContactContent = getCustomerIndexObjectValueK2_(obj, 'tmContactContent') || out.tmContactContent;
  return out;
}

function getMyCustomerContactHistoryMapP529_(ownRows) {
  const result = { byCustomerNo: {}, byRowNo: {}, count: 0 };
  const customerKeys = {};
  const rowKeys = {};
  (ownRows || []).forEach(function(row) {
    const cno = normalizeCustomerNoForKey_(row && row.customerNo || '');
    const rn = String(row && row.rowNo || '').trim();
    if (cno) customerKeys[cno] = true;
    if (rn) rowKeys[rn] = true;
  });
  if (!Object.keys(customerKeys).length && !Object.keys(rowKeys).length) return result;
  try {
    const ss = getWebAppDbSpreadsheet_();
    const sheet = ss.getSheetByName(PORTAL_CONFIG.CONTACT_HISTORY_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return result;
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
    const map = {};
    headers.forEach(function(h, i) { if (h) map[h] = i; });
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();
    values.forEach(function(row, i) {
      const cno = normalizeCustomerNoForKey_(cellByHeaderIndex_(row, map, ['고객번호']));
      const rn = String(cellByHeaderIndex_(row, map, ['마스터행', 'rowNo']) || '').trim();
      if (!(cno && customerKeys[cno]) && !(rn && rowKeys[rn])) return;
      const createdAt = cellByHeaderIndex_(row, map, ['작성일시', '일시', '등록일시']);
      const content = cellByHeaderIndex_(row, map, ['컨택내용', '메모']);
      const note = cellByHeaderIndex_(row, map, ['특이사항', '비고']);
      const status = cellByHeaderIndex_(row, map, ['계약진행상태', '통화결과']);
      const nextAction = cellByHeaderIndex_(row, map, ['다음액션']);
      const method = cellByHeaderIndex_(row, map, ['연락수단', '컨택방식']);
      const recordType = cellByHeaderIndex_(row, map, ['기록구분']) || '컨택이력';
      const author = cellByHeaderIndex_(row, map, ['작성자']);
      const text = [status, content, note, nextAction].map(function(v) { return String(v || '').trim(); }).filter(Boolean).join(' / ');
      if (!text && !createdAt) return;
      const ev = makeMyCustomerStatusEventP529_({
        source: 'contactHistory',
        sourceLabel: '컨택이력_DB',
        date: parseLoosePortalDateP528_(createdAt) || parseLoosePortalDateP528_(text),
        text: '[' + recordType + (method ? '/' + method : '') + '] ' + text,
        rawText: text,
        actor: author,
        order: 100000 + i
      });
      if (cno) {
        if (!result.byCustomerNo[cno]) result.byCustomerNo[cno] = [];
        result.byCustomerNo[cno].push(ev);
      }
      if (rn) {
        if (!result.byRowNo[rn]) result.byRowNo[rn] = [];
        result.byRowNo[rn].push(ev);
      }
      result.count += 1;
    });
  } catch (err) {
    try { Logger.log('나의 고객 현황 컨택이력_DB 조회 실패: ' + (err && err.stack || err)); } catch (e) {}
  }
  return result;
}

function buildMyCustomerStatusAnalyzedRowP528_(row, now, contactMap) {
  row = row || {};
  contactMap = contactMap || { byCustomerNo: {}, byRowNo: {} };
  const memo = String(row.fullMemo || row.memo || '');
  const status = String(row.status || '').trim();
  const sentAt = String(row.sentAt || row.lastSent || '');
  const customerNo = normalizeCustomerNoForKey_(row.customerNo || '');
  const rowNoKey = String(row.rowNo || '');
  const contactEvents = [].concat(
    customerNo && contactMap.byCustomerNo[customerNo] ? contactMap.byCustomerNo[customerNo] : [],
    rowNoKey && contactMap.byRowNo[rowNoKey] ? contactMap.byRowNo[rowNoKey] : []
  );
  const dedupedContactEvents = dedupeMyCustomerEventsP529_(contactEvents);
  const events = buildMyCustomerCombinedEventsP529_(row, memo, dedupedContactEvents);
  const latestEvent = getLatestMyCustomerMemoEventP528_(events);
  const latestDate = latestEvent && latestEvent.date ? latestEvent.date : parseLoosePortalDateP528_(sentAt);
  const lastContactDays = latestDate ? daysBetweenPortalDatesP528_(latestDate, now) : null;
  const analysis = classifyMyCustomerMemoP528_(row, events, latestEvent, lastContactDays, now);
  const latestText = latestEvent && latestEvent.text ? latestEvent.text : getLastMemoLineP528_(memo);
  const missingFields = getMyCustomerMissingFieldsP528_(row);
  analysis.missingFields = missingFields;
  analysis.dataMissing = missingFields.length > 0;
  analysis.events = events.slice(0, 40).map(function(ev) { return eventForClientP529_(ev); });
  analysis.memoHash = hashMyCustomerStatusTextP529_(memo);
  analysis.contactHistoryHash = hashMyCustomerStatusTextP529_(JSON.stringify(dedupedContactEvents.map(function(ev) { return eventForHashP529_(ev); })));
  analysis.sendHash = hashMyCustomerStatusTextP529_([row.lastSent, row.sentAt, row.sendStatus, row.sendCount].join('|'));

  return {
    rowNo: Number(row.rowNo || 0) || 0,
    customerNo: String(row.customerNo || ''),
    company: String(row.company || ''),
    salesRep: String(row.salesRep || ''),
    status: status,
    vendor: String(row.vendor || ''),
    contact: String(row.contact || ''),
    phone: String(row.phone || ''),
    directPhone: String(row.directPhone || ''),
    email: String(row.email || ''),
    finalQuote: String(row.finalQuote || ''),
    lastSent: String(row.lastSent || ''),
    sentAt: String(row.sentAt || ''),
    sendStatus: String(row.sendStatus || ''),
    memoSummary: shortenMyCustomerStatusTextP528_(latestText || memo, 260),
    lastContactDate: latestDate ? formatMyCustomerStatusDateP528_(latestDate) : '',
    analysis: analysis
  };
}

function buildMyCustomerCombinedEventsP529_(row, memo, contactEvents) {
  const events = [];
  extractMyCustomerMemoEventsP528_(memo).forEach(function(ev) {
    events.push(makeMyCustomerStatusEventP529_({
      source: 'masterMemo',
      sourceLabel: '마스터메모',
      date: ev.date,
      text: ev.text,
      rawText: ev.text,
      order: ev.index || 0
    }));
  });
  String(row.tmContactContent || '').split(/\r?\n+/).map(function(v) { return String(v || '').trim(); }).filter(Boolean).forEach(function(line, idx) {
    events.push(makeMyCustomerStatusEventP529_({
      source: 'tmContact',
      sourceLabel: 'TM컨택내용',
      date: parseLoosePortalDateP528_(line),
      text: '[TM] ' + line,
      rawText: line,
      order: 30000 + idx
    }));
  });
  (contactEvents || []).forEach(function(ev, idx) {
    ev.order = ev.order || (50000 + idx);
    events.push(ev);
  });
  const sendText = [row.sendStatus, row.lastSent, row.sentAt, row.sendCount ? ('발송횟수 ' + row.sendCount) : ''].filter(Boolean).join(' / ');
  if (sendText) {
    events.push(makeMyCustomerStatusEventP529_({
      source: 'sendLog',
      sourceLabel: '자료발송',
      date: parseLoosePortalDateP528_(row.sentAt || row.lastSent || sendText),
      text: '[자료발송] ' + sendText,
      rawText: sendText,
      order: 90000
    }));
  }
  return sortMyCustomerEventsDescP529_(dedupeMyCustomerEventsP529_(events));
}

function makeMyCustomerStatusEventP529_(data) {
  data = data || {};
  const text = String(data.text || data.rawText || '').trim();
  const ev = {
    source: String(data.source || 'unknown'),
    sourceLabel: String(data.sourceLabel || data.source || '이력'),
    date: data.date || null,
    text: text,
    rawText: String(data.rawText || text || ''),
    actor: String(data.actor || ''),
    order: Number(data.order || 0) || 0
  };
  ev.signal = getMyCustomerEventSignalsP529_(ev.text);
  return ev;
}

function classifyMyCustomerMemoP528_(row, events, latestEvent, lastContactDays, now) {
  row = row || {};
  events = Array.isArray(events) ? events : [];
  const currentStatus = String(row.status || '').trim();
  const stateProtected = isMyCustomerStatusProtectedStatusP530_(currentStatus);
  const sentDone = !!(String(row.sendStatus || '').indexOf('발송완료') >= 0 || row.lastSent || row.sentAt);
  const signalSummary = summarizeMyCustomerSignalsP529_(events);

  let recommended = '';
  let mismatchType = '';
  let confidence = 0;
  let reason = '';
  let nextAction = '';
  let priorityRank = '';
  let terminalCandidate = false;
  let analysisCautionLevel = '';
  let conservativeReason = '';

  const ownComplete = signalSummary.complete.latest;
  const thirdPartyContract = signalSummary.thirdPartyContract.latest;
  const unknownComplete = signalSummary.unknownComplete.latest;
  const fail = signalSummary.fail.latest;
  const failLike = getLatestMyCustomerSignalP530_([thirdPartyContract, fail]);
  const order = signalSummary.order.latest;
  const long = signalSummary.long.latest;
  const quote = signalSummary.quote.latest;
  const active = signalSummary.active.latest;
  const data = signalSummary.data.latest;
  const contractCompleteType = getMyCustomerContractCompleteTypeP530_(signalSummary);
  const hasThirdPartyContract = !!thirdPartyContract;

  if (stateProtected) {
    recommended = currentStatus;
    confidence = 98;
    priorityRank = '0_상태보호';
    terminalCandidate = true;
    mismatchType = currentStatus + ' 보호';
    analysisCautionLevel = hasThirdPartyContract && currentStatus === '계약완료' ? '참고확인' : '보호';
    conservativeReason = currentStatus + '은/는 영업담당자가 직접 판단했을 가능성이 높은 보호 상태입니다. 메모/TM 이력만으로 상태변경 추천을 하지 않습니다.';
    if (currentStatus === '수주실패' && (failLike || hasThirdPartyContract)) {
      mismatchType = '수주실패 정합 / 종결상태 보호';
      reason = hasThirdPartyContract ? '타사 계약/타사 결정 계열 이력이 있어 현재 수주실패 상태와 정합합니다.' : '거절/진행불가 계열 이력이 있어 현재 수주실패 상태와 정합합니다.';
      nextAction = '추가 영업 대상에서 제외하거나 장기 재접촉 여부만 참고하세요.';
    } else if (currentStatus === '발주완료') {
      mismatchType = '발주완료 보호';
      reason = '발주완료는 주요 진행 상태이므로 현재 상태를 우선 신뢰합니다.';
      nextAction = '계약완료 전환 여부는 담당자가 직접 확인해 주세요.';
    } else if (currentStatus === '계약완료') {
      mismatchType = '계약완료 보호';
      reason = hasThirdPartyContract ? '타사 계약 표현이 감지되었지만 현재 계약완료 상태는 자동으로 뒤집지 않습니다. 필요 시 참고 확인만 하세요.' : '계약완료는 주요 종결 상태이므로 현재 상태를 우선 신뢰합니다.';
      nextAction = '계약완료 자료/계약완료 시트와 대조가 필요할 때만 수동 확인하세요.';
    }
  } else if (thirdPartyContract) {
    recommended = '수주실패';
    confidence = 93;
    mismatchType = currentStatus === '수주실패' ? '수주실패 정합' : '수주실패 전환 의심';
    reason = '최신 유효 이력에 타사 계약/타사 결정/다른 업체 선정 신호가 있습니다. 타사 계약완료는 우리 계약완료가 아니라 수주실패 신호입니다.';
    nextAction = currentStatus === '수주실패' ? '추가 영업 대상에서 제외하거나 장기 재접촉 여부만 검토하세요.' : '수주실패 처리 여부를 확인하세요.';
    priorityRank = '2_타사계약수주실패';
    terminalCandidate = true;
    analysisCautionLevel = '높음';
  } else if (fail && !hasPositiveSignalAfterP529_(signalSummary, fail)) {
    recommended = '수주실패';
    confidence = 88;
    mismatchType = currentStatus === '수주실패' ? '수주실패 정합' : '수주실패 전환 의심';
    reason = '최신 유효 이력에 거절/진행불가 신호가 있습니다.';
    nextAction = currentStatus === '수주실패' ? '추가 영업 대상에서 제외하거나 장기 재접촉 여부만 검토하세요.' : '수주실패 처리 여부를 확인하세요.';
    priorityRank = '2_수주실패';
    terminalCandidate = true;
  } else if (ownComplete && !isMyCustomerSignalAfterP529_(failLike, ownComplete)) {
    recommended = '계약완료';
    confidence = 88;
    mismatchType = currentStatus === '계약완료' ? '계약완료 정합' : '계약완료 의심';
    reason = '당사 계약/계약완료 처리/계약서 저장 등 우리 계약완료로 볼 수 있는 강한 신호가 있습니다.';
    nextAction = '계약완료 자료와 마스터 상태가 일치하는지 담당자가 직접 확인하세요.';
    priorityRank = '1_우리계약완료';
    terminalCandidate = true;
    analysisCautionLevel = '보수확인';
  } else if (order && !isMyCustomerSignalAfterP529_(failLike, order)) {
    recommended = '발주완료';
    confidence = 78;
    mismatchType = currentStatus === '발주완료' ? '발주완료 정합' : '단계 상향 참고';
    reason = '최근 이력에 용역신청서/계약서/사업자등록증/발주 신호가 있습니다. 단, 발주완료는 주요 상태이므로 보수적으로 확인이 필요합니다.';
    nextAction = '계약서류 취합 및 발주 완료 여부를 확인하세요.';
    priorityRank = '4_발주계약진행';
    analysisCautionLevel = '보수확인';
  } else if (long) {
    recommended = '장기 추진건';
    confidence = 80;
    mismatchType = currentStatus === '장기 추진건' ? '장기추진 정합' : '장기추진 전환 의심';
    reason = '최근 이력에 예산/내년/추후/보류 신호가 있습니다.';
    nextAction = '다음 재접촉 시점을 지정하세요.';
    priorityRank = '3_장기추진';
  } else if (quote || sentDone) {
    // 자료발송은 보조근거입니다. 종결/부정/계약진행 신호를 뒤집지 않습니다.
    recommended = '견적제출완료';
    confidence = quote ? 73 : 66;
    mismatchType = currentStatus === '견적제출완료' ? '견적제출완료 정합' : '견적제출완료 참고';
    reason = quote ? '최근 이력에 견적/자료 발송 신호가 있습니다.' : '발송상태 또는 마지막발송 값이 있습니다.';
    nextAction = '견적서 발송 후 후속 연락을 진행하세요.';
    priorityRank = '5_견적제출';
  } else if (active) {
    recommended = '고객 설득 중';
    confidence = 67;
    mismatchType = currentStatus === '고객 설득 중' ? '진행중 정합' : '진행중 상태 참고';
    reason = '최근 이력에 검토/비교/재확인/담당자 전달 신호가 있습니다.';
    nextAction = '고객 반응을 확인하고 다음 액션을 남기세요.';
    priorityRank = '6_설득중';
  } else if (unknownComplete) {
    recommended = currentStatus || '!!상태지정필요!!';
    confidence = 58;
    mismatchType = '계약완료 주체불명';
    reason = '계약완료 표현은 있으나 타사 계약인지 당사 계약인지 불명확하여 상태변경 추천을 보류합니다.';
    nextAction = '계약 주체를 확인하세요.';
    priorityRank = '7_주체불명계약완료';
    analysisCautionLevel = '확인필요';
  } else if (data) {
    recommended = currentStatus || '!!상태지정필요!!';
    confidence = 62;
    mismatchType = '데이터확인 필요';
    reason = '이력에 중복/주소/연면적/연락처 확인 신호가 있습니다.';
    nextAction = '고객정보를 먼저 정리하세요.';
    priorityRank = '8_데이터확인';
  }

  const stateChangeAllowed = !stateProtected && isMyCustomerStatusStateChangeCandidateP530_(currentStatus);
  const recommendationVisible = !!(stateChangeAllowed && recommended && recommended !== currentStatus && confidence >= 75);
  const mismatch = !!recommendationVisible;
  const activeStatus = isMyCustomerStatusActiveP528_(currentStatus);
  const longNoContact = activeStatus && (lastContactDays == null || lastContactDays >= 14);
  let sentNoFollow = false;
  if (sentDone && activeStatus) {
    if (lastContactDays == null) sentNoFollow = true;
    else if (lastContactDays >= 3 && !order && !failLike && !ownComplete && !thirdPartyContract) sentNoFollow = true;
  }
  const potentialScore = stateProtected ? 0 : calculateMyCustomerPotentialScoreP528_(row, signalSummary.flags, lastContactDays, recommended);
  const tags = [];
  Object.keys(signalSummary.flags).forEach(function(k) { if (signalSummary.flags[k]) tags.push(k); });
  if (longNoContact) tags.push('장기미접촉');
  if (sentNoFollow) tags.push('자료발송후미후속');
  if (terminalCandidate || stateProtected) tags.push('종결성신호');
  if (stateProtected) tags.push('상태보호');
  if (hasThirdPartyContract) tags.push('타사계약');
  if (unknownComplete) tags.push('계약주체불명');

  const latestText = latestEvent && latestEvent.text ? latestEvent.text : '';
  return {
    recommendedStatus: recommended,
    confidence: confidence,
    priorityRank: priorityRank,
    mismatch: mismatch,
    mismatchType: mismatch || recommended === currentStatus || stateProtected ? mismatchType : '',
    terminalCandidate: terminalCandidate || stateProtected,
    reason: reason || conservativeReason,
    nextAction: nextAction,
    lastContactDays: lastContactDays,
    longNoContact: longNoContact,
    sentNoFollow: sentNoFollow,
    potentialScore: potentialScore,
    tags: tags,
    latestMemoDate: latestEvent && latestEvent.date ? formatMyCustomerStatusDateP528_(latestEvent.date) : '',
    latestMemoText: shortenMyCustomerStatusTextP528_(latestText, 220),
    latestEventSource: latestEvent ? latestEvent.sourceLabel : '',
    signalSummary: signalSummary,
    matchedKeywords: signalSummary.matchedKeywords,
    sourceLabels: signalSummary.sourceLabels,
    statusProtected: stateProtected,
    stateChangeAllowed: stateChangeAllowed,
    canApplyRecommendation: recommendationVisible,
    recommendationVisible: recommendationVisible,
    contractCompleteType: contractCompleteType,
    thirdPartyContract: hasThirdPartyContract,
    analysisCautionLevel: analysisCautionLevel,
    conservativeReason: conservativeReason || (stateProtected ? reason : '')
  };
}

function isMyCustomerStatusProtectedStatusP530_(status) {
  status = String(status || '').trim();
  return MY_CUSTOMER_STATUS_BLOCKING_STATUSES_P528.indexOf(status) >= 0;
}

function isMyCustomerStatusStateChangeCandidateP530_(status) {
  status = String(status || '').trim();
  if (isMyCustomerStatusProtectedStatusP530_(status)) return false;
  if (!status) return true;
  if (status.indexOf('상태지정') >= 0) return true;
  return ['고객 설득 중', '견적제출완료', '장기 추진건', '데이터확인필요'].indexOf(status) >= 0;
}

function getLatestMyCustomerSignalP530_(events) {
  let best = null;
  (events || []).forEach(function(ev) {
    if (!ev) return;
    if (!best || compareMyCustomerEventsP529_(ev, best) > 0) best = ev;
  });
  return best;
}

function getMyCustomerContractCompleteTypeP530_(summary) {
  summary = summary || {};
  if (summary.thirdPartyContract && summary.thirdPartyContract.latest) return '타사계약완료';
  if (summary.complete && summary.complete.latest) return '우리계약완료';
  if (summary.unknownComplete && summary.unknownComplete.latest) return '계약완료_주체불명';
  return '';
}

function getMyCustomerEventSignalsP529_(text) {
  text = normalizeMyCustomerSignalTextP529_(text);
  const defs = getMyCustomerSignalDefsP529_();
  const out = { complete: [], thirdPartyContract: [], unknownComplete: [], fail: [], order: [], long: [], quote: [], active: [], data: [] };
  Object.keys(defs).forEach(function(k) {
    defs[k].forEach(function(rule) {
      let hit = false;
      if (rule.re) hit = rule.re.test(text);
      else if (rule.kw) hit = text.indexOf(normalizeMyCustomerSignalTextP529_(rule.kw)) >= 0;
      if (hit) out[k].push(rule.label || rule.kw || String(rule.re));
    });
  });
  return out;
}

function getMyCustomerSignalDefsP529_() {
  return {
    // complete는 우리 계약완료로 볼 수 있는 강한 신호만 둡니다. 단순 "계약완료"는 unknownComplete로 분리합니다.
    complete: [
      { kw: '당사와 계약' }, { kw: '에스원과 계약' }, { kw: '우리랑 계약' }, { kw: '우리와 계약' },
      { kw: '계약완료 처리' }, { kw: '계약 완료 처리' }, { kw: '계약완료 시트' }, { kw: '계약 완료 시트' },
      { kw: '계약서 저장' }, { kw: '계약서 취합 완료' }, { kw: '발주번호' }, { kw: '계약 체결 완료' },
      { re: /(일신|삼구|kj|케이제이|에스원).{0,12}(계약완료|계약 완료|계약체결|계약 체결)/, label: '우리 수행사 계약완료' }
    ],
    thirdPartyContract: [
      { kw: '타사계약완료', label: '타사 계약완료' }, { kw: '타사 계약완료', label: '타사 계약완료' }, { kw: '타사 계약 완료', label: '타사 계약완료' },
      { kw: '타업체계약완료', label: '타업체 계약완료' }, { kw: '타업체 계약완료', label: '타업체 계약완료' }, { kw: '타업체와 계약완료', label: '타업체 계약완료' },
      { kw: '다른업체와 계약완료', label: '다른업체 계약완료' }, { kw: '다른 업체와 계약 완료', label: '다른업체 계약완료' },
      { kw: '타사로 결정', label: '타사로 결정' }, { kw: '타사 결정', label: '타사 결정' }, { kw: '타업체로 결정', label: '타업체로 결정' }, { kw: '타업체 결정', label: '타업체 결정' },
      { kw: '다른 업체로 결정', label: '다른 업체로 결정' }, { kw: '다른곳으로 결정', label: '다른 곳으로 결정' }, { kw: '다른 곳으로 결정', label: '다른 곳으로 결정' },
      { kw: '기존 업체 유지', label: '기존업체 유지' }, { kw: '기존업체 유지', label: '기존업체 유지' }, { kw: '금액 낮은 곳으로 계약', label: '저가 타업체 계약' },
      { kw: '장애인사업장하고 진행', label: '타 기관 진행' }, { kw: '업체 선정 완료', label: '타 업체 선정 완료' }, { kw: '선정 완료', label: '업체 선정 완료' },
      { re: /타(사|업체).{0,12}(결정|계약|진행|완료|선정)/, label: '타사 결정/계약' },
      { re: /(다른|기존).{0,6}(업체|곳).{0,12}(계약|진행|결정|유지|완료)/, label: '다른 업체 계약/결정' },
      { re: /(가격|금액).{0,10}(낮|저렴).{0,12}(곳|업체).{0,12}(계약|진행|결정)/, label: '저가 업체 계약/결정' }
    ],
    unknownComplete: [
      { re: /(^|[^타사타업체다른기존])계약\s*완료(?!\s*처리|\s*시트)/, label: '계약완료 주체불명' },
      { kw: '계약완료', label: '계약완료 주체불명' }, { kw: '계약 완료', label: '계약완료 주체불명' }
    ],
    fail: [
      { kw: '수주실패' }, { kw: '거절' }, { kw: '가격차이가 많이 나서 어렵', label: '가격차이로 어려움' }, { kw: '가격 차이가 많이 나서 어렵', label: '가격차이로 어려움' },
      { kw: '진행 어렵' }, { kw: '진행어렵' }, { kw: '어렵다고' }, { kw: '불가' }, { kw: '안한다고' }, { kw: '안 한다고' },
      { kw: '하지 말' }, { kw: '연락하지' }, { kw: '대상 아님' }, { kw: '대상아님' }, { kw: '폐업' }, { kw: '취소' },
      { re: /(가격|금액).{0,12}(차이|비싸).{0,16}(어렵|불가|안)/, label: '가격 이슈로 진행불가' }
    ],
    order: [
      { kw: '용역신청서' }, { kw: '계약서 요청' }, { kw: '계약서류' }, { kw: '계약 진행' }, { kw: '계약진행' },
      { kw: '사업자등록증' }, { kw: '사업자 등록증' }, { kw: '발주' }, { kw: '진행하시기로' }, { kw: '계약하시기로' },
      { kw: '서류 받' }, { kw: '서류 요청' }, { kw: '취합' }
    ],
    long: [
      { kw: '내년' }, { kw: '예산' }, { kw: '하반기' }, { kw: '상반기' }, { kw: '추후' }, { kw: '나중' }, { kw: '장기' }, { kw: '보류' }, { kw: '재검토' }
    ],
    quote: [
      { kw: '견적서 발송' }, { kw: '견적서발송' }, { kw: '견적 발송' }, { kw: '견적발송' }, { kw: '자료발송' }, { kw: '자료 발송' },
      { kw: '단가표' }, { kw: '비교견적' }, { kw: '수기견적' }, { kw: '메일 발송' }, { kw: '발송완료' }, { kw: '견적서 재발송' }, { kw: '재견적' }
    ],
    active: [
      { kw: '검토중' }, { kw: '검토 중' }, { kw: '비교중' }, { kw: '비교 중' }, { kw: '타업체와 비교' }, { kw: '타사 비교' },
      { kw: '담당자 전달' }, { kw: '내부 검토' }, { kw: '확인 후' }, { kw: '재확인' }, { kw: '연락 예정' }, { kw: '연락준다고' },
      { kw: '가격 조율' }, { kw: '본사' }, { kw: '회신' }, { kw: '상담' }
    ],
    data: [
      { kw: '중복' }, { kw: '전화번호 오류' }, { kw: '메일 오류' }, { kw: '주소 확인' }, { kw: '연면적 확인' }, { kw: '확인 필요' }, { kw: '정보 확인' }, { kw: '번호 오류' }
    ]
  };
}

function normalizeMyCustomerSignalTextP529_(text) {
  return String(text || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function summarizeMyCustomerSignalsP529_(events) {
  const categories = ['complete', 'thirdPartyContract', 'unknownComplete', 'fail', 'order', 'long', 'quote', 'active', 'data'];
  const summary = {
    flags: {},
    matchedKeywords: {},
    sourceLabels: [],
    complete: { latest: null, count: 0, keywords: [] },
    thirdPartyContract: { latest: null, count: 0, keywords: [] },
    unknownComplete: { latest: null, count: 0, keywords: [] },
    fail: { latest: null, count: 0, keywords: [] },
    order: { latest: null, count: 0, keywords: [] },
    long: { latest: null, count: 0, keywords: [] },
    quote: { latest: null, count: 0, keywords: [] },
    active: { latest: null, count: 0, keywords: [] },
    data: { latest: null, count: 0, keywords: [] }
  };
  const sourceSeen = {};
  (events || []).forEach(function(ev) {
    if (ev && ev.sourceLabel && !sourceSeen[ev.sourceLabel]) {
      sourceSeen[ev.sourceLabel] = true;
      summary.sourceLabels.push(ev.sourceLabel);
    }
    categories.forEach(function(cat) {
      const hits = ev && ev.signal && ev.signal[cat] || [];
      if (!hits.length) return;
      summary.flags[cat] = true;
      summary[cat].count += 1;
      hits.forEach(function(h) {
        if (summary[cat].keywords.indexOf(h) < 0) summary[cat].keywords.push(h);
      });
      if (!summary[cat].latest || compareMyCustomerEventsP529_(ev, summary[cat].latest) > 0) summary[cat].latest = ev;
    });
  });
  categories.forEach(function(cat) {
    summary.matchedKeywords[cat] = summary[cat].keywords.slice(0, 12);
  });
  return summary;
}

function hasPositiveSignalAfterP529_(summary, baseEvent) {
  if (!baseEvent) return false;
  const positives = [];
  ['complete', 'order', 'quote', 'active'].forEach(function(cat) {
    if (summary && summary[cat] && summary[cat].latest) positives.push(summary[cat].latest);
  });
  return positives.some(function(ev) {
    if (!isMyCustomerSignalAfterP529_(ev, baseEvent)) return false;
    // 단순 과거 자료발송보다 종결 이후의 재견적/재진행/계약 진행 신호만 살립니다.
    const txt = normalizeMyCustomerSignalTextP529_(ev.text || '');
    if (ev.source === 'sendLog' && txt.indexOf('재') < 0) return false;
    return true;
  });
}

function isMyCustomerSignalAfterP529_(a, b) {
  if (!a || !b) return false;
  return compareMyCustomerEventsP529_(a, b) > 0;
}

function compareMyCustomerEventsP529_(a, b) {
  const at = a && a.date ? a.date.getTime() : 0;
  const bt = b && b.date ? b.date.getTime() : 0;
  if (at !== bt) return at - bt;
  return (Number(a && a.order || 0) || 0) - (Number(b && b.order || 0) || 0);
}

function calculateMyCustomerPotentialScoreP528_(row, signals, lastContactDays, recommended) {
  signals = signals || {};
  let score = 0;
  if (lastContactDays != null && lastContactDays <= 7) score += 2;
  if (row.lastSent || row.sentAt || String(row.sendStatus || '').indexOf('발송완료') >= 0) score += 2;
  if (row.finalQuote) score += 1;
  if (row.email) score += 1;
  if (row.contact) score += 1;
  if (signals.order) score += 3;
  if (signals.quote || signals.active) score += 1;
  if (signals.fail || signals.thirdPartyContract) score -= 6;
  if (signals.complete || signals.unknownComplete) score -= 2;
  if (recommended === '발주완료') score += 2;
  if (recommended === '수주실패') score -= 4;
  if (lastContactDays != null && lastContactDays >= 30) score -= 2;
  return Math.max(0, score);
}

function getMyCustomerMissingFieldsP528_(row) {
  row = row || {};
  const miss = [];
  if (!String(row.contact || '').trim()) miss.push('담당자');
  if (!String(row.phone || row.directPhone || '').trim()) miss.push('전화/직통');
  if (!String(row.email || '').trim()) miss.push('이메일');
  if (!String(row.finalQuote || '').trim()) miss.push('최종견적가');
  if (!String(row.vendor || '').trim()) miss.push('수행사');
  return miss;
}

function isMyCustomerStatusNeedStatusP528_(status) {
  status = String(status || '').trim();
  return !status || status.indexOf('상태지정') >= 0;
}

function isMyCustomerStatusActiveP528_(status) {
  status = String(status || '').trim();
  if (!status) return true;
  if (MY_CUSTOMER_STATUS_BLOCKING_STATUSES_P528.indexOf(status) >= 0) return false;
  return true;
}

function extractMyCustomerMemoEventsP528_(memo) {
  memo = String(memo || '');
  if (!memo) return [];
  const lines = memo.split(/\r?\n+/).map(function(line) { return String(line || '').trim(); }).filter(Boolean);
  const events = [];
  lines.forEach(function(line, idx) {
    const d = parseLoosePortalDateP528_(line);
    events.push({ index: idx, text: line, date: d });
  });
  return events;
}

function getLatestMyCustomerMemoEventP528_(events) {
  events = Array.isArray(events) ? events : [];
  let best = null;
  events.forEach(function(ev) {
    if (!best) { best = ev; return; }
    if (compareMyCustomerEventsP529_(ev, best) >= 0) best = ev;
  });
  return best;
}

function getRecentMemoTextP528_(events, memo, limit) {
  events = Array.isArray(events) ? events : [];
  limit = Math.max(1, Number(limit || 5) || 5);
  if (events.length) return events.slice(0, limit).map(function(ev) { return ev.text; }).join('\n');
  return String(memo || '').split(/\r?\n+/).slice(-limit).join('\n');
}

function getLastMemoLineP528_(memo) {
  const lines = String(memo || '').split(/\r?\n+/).map(function(v) { return String(v || '').trim(); }).filter(Boolean);
  return lines.length ? lines[lines.length - 1] : '';
}

function parseLoosePortalDateP528_(text) {
  text = String(text || '');
  let m = text.match(/(20\d{2})\s*[.\-\/년]\s*(\d{1,2})\s*[.\-\/월]\s*(\d{1,2})/);
  if (m) return safePortalDateP528_(Number(m[1]), Number(m[2]), Number(m[3]));
  m = text.match(/(?:^|[^\d])(\d{2})\s*[.\-\/]\s*(\d{1,2})\s*[.\-\/]\s*(\d{1,2})/);
  if (m) return safePortalDateP528_(2000 + Number(m[1]), Number(m[2]), Number(m[3]));
  m = text.match(/(20\d{2})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일/);
  if (m) return safePortalDateP528_(Number(m[1]), Number(m[2]), Number(m[3]));
  return null;
}

function safePortalDateP528_(year, month, day) {
  if (!year || !month || !day) return null;
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  return new Date(year, month - 1, day);
}

function daysBetweenPortalDatesP528_(a, b) {
  if (!a || !b) return null;
  const da = new Date(a.getFullYear(), a.getMonth(), a.getDate()).getTime();
  const db = new Date(b.getFullYear(), b.getMonth(), b.getDate()).getTime();
  return Math.floor((db - da) / 86400000);
}

function hasAnyKeywordP528_(text, keywords) {
  text = String(text || '').toLowerCase().replace(/\s+/g, ' ');
  return (keywords || []).some(function(k) {
    return text.indexOf(String(k || '').toLowerCase()) >= 0;
  });
}

function shortenMyCustomerStatusTextP528_(text, maxLen) {
  text = String(text || '').replace(/\s+/g, ' ').trim();
  maxLen = Number(maxLen || 160) || 160;
  return text.length > maxLen ? text.slice(0, maxLen - 1) + '…' : text;
}

function formatMyCustomerStatusDateP528_(date) {
  if (!date) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatMyCustomerStatusDateTimeP528_(date) {
  if (!date) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function objectToSortedStatusCountArrayP528_(obj) {
  obj = obj || {};
  const order = (PORTAL_CONFIG.STATUS_OPTIONS || []).concat(['(공란)']);
  const used = {};
  const out = [];
  order.forEach(function(k) {
    if (obj[k]) { out.push({ status: k, count: obj[k] }); used[k] = true; }
  });
  Object.keys(obj).sort().forEach(function(k) {
    if (!used[k]) out.push({ status: k, count: obj[k] });
  });
  return out;
}

function sortMyCustomerStatusByRecentP528_(a, b) {
  const ad = a && a.analysis && a.analysis.lastContactDays;
  const bd = b && b.analysis && b.analysis.lastContactDays;
  if (ad == null && bd == null) return 0;
  if (ad == null) return 1;
  if (bd == null) return -1;
  return ad - bd;
}

function sortMyCustomerStatusByStaleP528_(a, b) {
  const ad = a && a.analysis && a.analysis.lastContactDays;
  const bd = b && b.analysis && b.analysis.lastContactDays;
  if (ad == null && bd == null) return 0;
  if (ad == null) return -1;
  if (bd == null) return 1;
  return bd - ad;
}

function makeMyCustomerStatusOperationIdP528_(customerNo, rowNo) {
  return 'MY_STATUS_P530_' + String(customerNo || '') + '_' + String(rowNo || '') + '_' + String(Date.now());
}

function sortMyCustomerEventsDescP529_(events) {
  return (events || []).slice().sort(function(a, b) { return compareMyCustomerEventsP529_(b, a); });
}

function dedupeMyCustomerEventsP529_(events) {
  const seen = {};
  const out = [];
  (events || []).forEach(function(ev) {
    if (!ev) return;
    const key = [ev.source, formatMyCustomerStatusDateP528_(ev.date), shortenMyCustomerStatusTextP528_(ev.text, 120)].join('|');
    if (seen[key]) return;
    seen[key] = true;
    out.push(ev);
  });
  return out;
}

function eventForClientP529_(ev) {
  return {
    source: ev.source,
    sourceLabel: ev.sourceLabel,
    date: ev.date ? formatMyCustomerStatusDateP528_(ev.date) : '',
    text: shortenMyCustomerStatusTextP528_(ev.text, 260),
    actor: ev.actor || '',
    keywords: flattenMyCustomerSignalKeywordsP529_(ev.signal)
  };
}

function eventForHashP529_(ev) {
  return {
    source: ev.source,
    date: ev.date ? formatMyCustomerStatusDateP528_(ev.date) : '',
    text: ev.text || '',
    actor: ev.actor || ''
  };
}

function flattenMyCustomerSignalKeywordsP529_(signal) {
  const out = [];
  Object.keys(signal || {}).forEach(function(k) {
    (signal[k] || []).forEach(function(v) { if (out.indexOf(v) < 0) out.push(v); });
  });
  return out.slice(0, 10);
}

function hashMyCustomerStatusTextP529_(text) {
  text = String(text || '');
  if (!text) return '';
  try {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
    return bytes.map(function(b) { const v = (b < 0 ? b + 256 : b).toString(16); return v.length === 1 ? '0' + v : v; }).join('').slice(0, 24);
  } catch (err) {
    return String(text.length) + '_' + String(text.charCodeAt(0) || 0) + '_' + String(text.charCodeAt(text.length - 1) || 0);
  }
}

function ensureMyCustomerStatusAnalysisSheetP529_() {
  const ss = getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(MY_CUSTOMER_STATUS_ANALYSIS_SHEET_P529);
  if (!sheet) {
    sheet = ss.insertSheet(MY_CUSTOMER_STATUS_ANALYSIS_SHEET_P529);
    sheet.getRange(1, 1, 1, MY_CUSTOMER_STATUS_ANALYSIS_HEADERS_P529.length).setValues([MY_CUSTOMER_STATUS_ANALYSIS_HEADERS_P529]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, MY_CUSTOMER_STATUS_ANALYSIS_HEADERS_P529.length).setFontWeight('bold').setBackground('#f2f4f7');
    sheet.autoResizeColumns(1, Math.min(MY_CUSTOMER_STATUS_ANALYSIS_HEADERS_P529.length, 12));
    return sheet;
  }
  ensureSheetHeaders_(sheet, MY_CUSTOMER_STATUS_ANALYSIS_HEADERS_P529);
  return sheet;
}

function saveMyCustomerStatusAnalysisRowsP529_(rows, perm, aliases, analyzedAt) {
  rows = Array.isArray(rows) ? rows : [];
  const sheet = ensureMyCustomerStatusAnalysisSheetP529_();
  const lastCol = Math.max(sheet.getLastColumn(), MY_CUSTOMER_STATUS_ANALYSIS_HEADERS_P529.length);
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const headerMap = {};
  headers.forEach(function(h, i) { if (h) headerMap[h] = i; });
  const ownerAliases = {};
  (aliases || []).forEach(function(a) { const k = normalizeMyCustomerStatusNameP528_(a); if (k) ownerAliases[k] = true; });
  const ownerName = String((perm && (perm.salesRepName || perm.name || perm.displayName)) || '').trim();
  const ownerKey = normalizeMyCustomerStatusNameP528_(ownerName);
  if (ownerKey) ownerAliases[ownerKey] = true;

  const existing = [];
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    values.forEach(function(r) {
      const rep = String(r[headerMap['영업담당자']] || '').trim();
      const rk = normalizeMyCustomerStatusNameP528_(rep);
      if (rk && ownerAliases[rk]) return;
      existing.push(r);
    });
  }

  const nowText = formatMyCustomerStatusDateTimeP528_(analyzedAt || new Date());
  const newRows = rows.map(function(row) { return buildMyCustomerAnalysisDbRowP529_(row, nowText, headers); });
  const combined = existing.concat(newRows);
  if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  if (combined.length) sheet.getRange(2, 1, combined.length, lastCol).setValues(combined);
  return {
    saved: true,
    rows: newRows.length,
    totalRows: combined.length,
    sheetName: MY_CUSTOMER_STATUS_ANALYSIS_SHEET_P529,
    spreadsheetUrl: getWebAppDbSpreadsheet_().getUrl(),
    savedAt: nowText
  };
}

function buildMyCustomerAnalysisDbRowP529_(row, analyzedAtText, headers) {
  row = row || {};
  const a = row.analysis || {};
  const s = a.signalSummary || {};
  const latestEvent = (a.events && a.events[0]) || {};
  const evidence = {
    reason: a.reason || '',
    nextAction: a.nextAction || '',
    sourceLabels: a.sourceLabels || [],
    matchedKeywords: a.matchedKeywords || {},
    latestEvent: latestEvent
  };
  const sourceLabels = (a.sourceLabels || []).join(', ');
  const rec = {
    '분석일시': analyzedAtText,
    '분석ID': 'MCSA-' + String(row.customerNo || '') + '-' + String(row.rowNo || '') + '-' + String(a.memoHash || '').slice(0, 8),
    '고객번호': row.customerNo || '',
    'rowNo': row.rowNo || '',
    '회사명': row.company || '',
    '영업담당자': row.salesRep || '',
    '현재상태': row.status || '',
    '추천상태': a.recommendedStatus || '',
    '판정유형': a.mismatchType || '',
    '신뢰도': a.confidence || '',
    '우선순위등급': a.priorityRank || '',
    '불일치여부': a.mismatch ? 'Y' : 'N',
    '분석소스': sourceLabels,
    '최근연락일': row.lastContactDate || a.latestMemoDate || '',
    '최근이벤트출처': a.latestEventSource || latestEvent.sourceLabel || '',
    '최근이벤트요약': a.latestMemoText || row.memoSummary || '',
    '강한긍정키워드': [].concat((a.matchedKeywords && a.matchedKeywords.complete) || [], (a.matchedKeywords && a.matchedKeywords.order) || [], (a.matchedKeywords && a.matchedKeywords.active) || []).join(', '),
    '강한부정키워드': ((a.matchedKeywords && a.matchedKeywords.fail) || []).join(', '),
    '계약진행키워드': ((a.matchedKeywords && a.matchedKeywords.order) || []).join(', '),
    '자료발송키워드': ((a.matchedKeywords && a.matchedKeywords.quote) || []).join(', '),
    '장기추진키워드': ((a.matchedKeywords && a.matchedKeywords.long) || []).join(', '),
    '데이터누락키워드': (a.missingFields || []).join(', '),
    '가능성점수': a.potentialScore || 0,
    '위험태그': (a.tags || []).join(', '),
    '추천액션': a.nextAction || '',
    '근거JSON': safeStringifyMyCustomerStatusP529_(evidence, 12000),
    '이벤트JSON': safeStringifyMyCustomerStatusP529_((a.events || []).slice(0, 30), 30000),
    'memoHash': a.memoHash || '',
    'contactHistoryHash': a.contactHistoryHash || '',
    'sendHash': a.sendHash || '',
    '분석버전': MY_CUSTOMER_STATUS_ANALYZER_VERSION_P528,
    '검토상태': '',
    '검토자': '',
    '검토일시': '',
    '검토메모': '',
    '상태보호여부': a.statusProtected ? 'Y' : 'N',
    '상태변경추천허용': a.stateChangeAllowed ? 'Y' : 'N',
    '계약완료구분': a.contractCompleteType || '',
    '타사계약여부': a.thirdPartyContract ? 'Y' : 'N',
    '추천노출여부': a.recommendationVisible ? 'Y' : 'N',
    '분석주의등급': a.analysisCautionLevel || '',
    '보수판정사유': a.conservativeReason || ''
  };
  return headers.map(function(h) { return rec[h] == null ? '' : rec[h]; });
}

function safeStringifyMyCustomerStatusP529_(value, maxLen) {
  let s = '';
  try { s = JSON.stringify(value || {}); } catch (err) { s = String(value || ''); }
  maxLen = Number(maxLen || 20000) || 20000;
  return s.length > maxLen ? s.slice(0, maxLen - 1) + '…' : s;
}
