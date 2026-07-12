/***************************************
 * S1 Sales Portal - 29_MyCustomerStatusService.gs
 * P539: 나의 고객 현황 - 분석DB 사전생성/빠른 조회 구조
 * - 내 고객 기준: 마스터시트 영업담당자 = 현재 로그인 사용자의 영업담당자명
 * - 마스터시트 원본 메모 + 컨택이력_DB + TM 컨택 내용 + 자료발송 스냅샷을 통합 분석
 * - 수주실패/발주완료/계약완료는 영업담당자 판단 우선 보호 상태로 취급
 ***************************************/

const MY_CUSTOMER_STATUS_ANALYZER_VERSION_P528 = 'P541_FAST_LOAD_MEASURED_THIN_COLUMNS';
const MY_CUSTOMER_STATUS_ANALYSIS_SHEET_P529 = '고객현황분석_DB';
const MY_CUSTOMER_STATUS_ANALYSIS_JOB_SHEET_P539 = '고객현황분석작업_DB';
const MY_CUSTOMER_STATUS_ANALYSIS_JOB_PROCESS_LIMIT_P539 = 200;
const MY_CUSTOMER_STATUS_ANALYSIS_JOB_HEADERS_P539 = [
  '등록일시', '수정일시', '작업ID', '요청자', '범위', '상태', '전체건수', '처리건수', '성공건수',
  '스킵건수', '실패건수', 'nextOffset', 'limit', 'ownerAliasesJson', 'ownerEmail', 'scopeReason',
  'resultJson', '마지막오류', '완료일시'
];
const MY_CUSTOMER_STATUS_ANALYSIS_WRITE_BATCH_P532 = 120;
const MY_CUSTOMER_STATUS_DB_EVIDENCE_JSON_LIMIT_P532 = 3500;
const MY_CUSTOMER_STATUS_DB_EVENTS_JSON_LIMIT_P532 = 5000;
const MY_CUSTOMER_STATUS_EXPLICIT_ADMIN_EMAILS_P532 = ['bang@s1samsung.com', 'xnewspringx@gmail.com', 'mhj842@gmail.com'];
const MY_CUSTOMER_STATUS_BLOCKING_STATUSES_P528 = ['수주실패', '발주완료', '계약완료'];
const MY_CUSTOMER_STATUS_HARD_PROTECTED_STATUSES_P536 = ['수주실패', '계약완료'];
const MY_CUSTOMER_STATUS_ACTIVE_STATUSES_P528 = ['견적제출완료', '장기 추진건', '고객 설득 중', '발주완료', '!!상태지정필요!!'];
const MY_CUSTOMER_STATUS_ANALYSIS_HEADERS_P529 = [
  '분석일시', '분석ID', '고객번호', 'rowNo', '회사명', '영업담당자', '현재상태', '추천상태',
  '판정유형', '신뢰도', '우선순위등급', '불일치여부', '분석소스', '최근연락일', '최근이벤트출처',
  '최근이벤트요약', '강한긍정키워드', '강한부정키워드', '계약진행키워드', '자료발송키워드',
  '장기추진키워드', '데이터누락키워드', '가능성점수', '위험태그', '추천액션', '근거JSON',
  '이벤트JSON', 'memoHash', 'contactHistoryHash', 'sendHash', '분석버전', '검토상태', '검토자', '검토일시', '검토메모',
  '상태보호여부', '상태변경추천허용', '계약완료구분', '타사계약여부', '추천노출여부', '분석주의등급', '보수판정사유',
  '상태변경추천상태', '상태변경추천여부', '상태변경추천등급', '업무인사이트유형', '업무인사이트요약',
  '최신전체이벤트요약', '최신판정이벤트요약', '최신판정이벤트일자', '최신판정이벤트출처',
  '이벤트분류요약', '제외이벤트요약', '상태추천차단사유'
];


// P541: 첫 화면은 표/카드/정렬에 필요한 얇은 컬럼만 읽습니다.
// 상세 검증용 키워드/hash/JSON/검토 메모는 더블클릭 상세 또는 DB 검증용으로 분리합니다.
const MY_CUSTOMER_STATUS_FAST_DB_HEADERS_P540 = [
  '분석일시', '고객번호', 'rowNo', '회사명', '영업담당자', '현재상태', '추천상태',
  '신뢰도', '데이터누락키워드', '가능성점수', '추천액션', '상태보호여부', '추천노출여부',
  '상태변경추천상태', '상태변경추천여부', '상태변경추천등급',
  '업무인사이트유형', '업무인사이트요약', '최근연락일',
  '최신판정이벤트요약', '최신판정이벤트일자', '최신판정이벤트출처',
  '상태추천차단사유'
];
const MY_CUSTOMER_STATUS_HEAVY_DB_HEADERS_P540 = ['근거JSON', '이벤트JSON', 'rawResultJson', 'payloadJson', 'candidateJson', 'resultJson'];


function getMyCustomerStatusDashboardP528(options) {
  options = options || {};
  // P539: 기본 화면 진입은 분석을 새로 하지 않고 고객현황분석_DB만 빠르게 읽습니다.
  // 분석DB 갱신은 start/process job 함수로 분리합니다.
  if (options.directAnalyze === true || options.forceAnalyze === true) {
    return getMyCustomerStatusDashboardFreshP539_(Object.assign({}, options, { persist: options.persist !== false }));
  }
  return getMyCustomerStatusDashboardFromDbP539_(options);
}

function getMyCustomerStatusDashboardFreshP539_(options) {
  options = options || {};
  const started = new Date();
  let stage = 'start';
  try {
    stage = 'permission';
    const perm = getPortalCurrentPermission_();
    const aliases = buildMyCustomerStatusOwnerAliasesP528_(perm);
    const scopeInfo = resolveMyCustomerStatusScopeP532_(perm);
    const isAllScope = !!scopeInfo.isAllScope;
    const now = new Date();

    stage = 'masterSource';
    const sourceInfo = getMyCustomerStatusSourceRowsP532_(perm, aliases, isAllScope);
    const scopedRows = sourceInfo.rows || [];

    stage = 'contactHistory';
    const contactMap = getMyCustomerContactHistoryMapP529_(scopedRows);

    stage = 'contractCompleteLookup';
    const contractCompleteMap = getMyCustomerContractCompleteMapP536_();

    stage = 'analysis';
    const analyzed = scopedRows.map(function(row) {
      return buildMyCustomerStatusAnalyzedRowP528_(row, now, contactMap, contractCompleteMap);
    }).filter(function(row) { return row && row.rowNo; });

    const statusCounts = {};
    analyzed.forEach(function(row) {
      const key = row.status || '(공란)';
      statusCounts[key] = (statusCounts[key] || 0) + 1;
    });

    const stateChangeRows = analyzed.filter(function(row) { return !!(row.analysis && row.analysis.statusChangeRecommendation); });
    const recentRows = analyzed.filter(function(row) { return row.analysis.lastContactDays != null && row.analysis.lastContactDays <= 7; });
    const noContactRows = analyzed.filter(function(row) { return !!row.analysis.longNoContact; });
    const sentNoFollowRows = analyzed.filter(function(row) { return !!row.analysis.sentNoFollow || row.analysis.insightType === '자료발송후미후속' || row.analysis.insightType === '견적후속필요'; });
    const highPotentialRows = analyzed.filter(function(row) { return row.analysis.potentialScore >= 6; });
    const dataMissingRows = analyzed.filter(function(row) { return !!row.analysis.dataMissing; });
    const terminalRows = analyzed.filter(function(row) { return !!row.analysis.terminalCandidate || !!row.analysis.statusProtected; });
    const contactIssueRows = analyzed.filter(function(row) { return row.analysis.insightType === '연락장애/담당자확인'; });
    const contractCheckRows = analyzed.filter(function(row) { return row.analysis.insightType === '계약/발주확인필요'; });
    const longReconnectRows = analyzed.filter(function(row) { return row.analysis.insightType === '장기재접촉'; });

    const activeCount = analyzed.filter(function(row) { return isMyCustomerStatusActiveP528_(row.status); }).length;
    const needStatusCount = analyzed.filter(function(row) { return isMyCustomerStatusNeedStatusP528_(row.status); }).length;

    let analysisDbResult = { saved: false, rows: 0, sheetName: MY_CUSTOMER_STATUS_ANALYSIS_SHEET_P529 };
    if (options.persist !== false) {
      stage = 'persistAnalysisDb';
      try {
        analysisDbResult = saveMyCustomerStatusAnalysisRowsP529_(analyzed, perm, aliases, started, {
          allScope: isAllScope,
          scopeInfo: scopeInfo,
          sourceInfo: sourceInfo
        });
      } catch (err) {
        analysisDbResult = {
          saved: false,
          rows: 0,
          sheetName: MY_CUSTOMER_STATUS_ANALYSIS_SHEET_P529,
          stage: 'persistAnalysisDb',
          error: err && err.message ? err.message : String(err || '')
        };
        try { Logger.log('고객현황분석_DB 저장 실패: ' + (err && err.stack || err)); } catch (e) {}
      }
    }

    stage = 'buildClientResponse';
    const clientRows = analyzed.map(buildMyCustomerStatusClientRowP531_);
    return {
      ok: true,
      version: MY_CUSTOMER_STATUS_ANALYZER_VERSION_P528,
      generatedAt: formatMyCustomerStatusDateTimeP528_(started),
      elapsedMs: new Date().getTime() - started.getTime(),
      owner: {
        name: isAllScope ? '전체 고객' : (perm.salesRepName || perm.name || ''),
        displayName: isAllScope ? ((perm.displayName || perm.name || perm.email || '') + ' · ADMIN 전체') : (perm.displayName || perm.name || perm.email || ''),
        email: perm.email || '',
        aliases: aliases,
        scope: isAllScope ? 'ALL' : 'OWN',
        isAdmin: !!isAllScope,
        scopeReason: scopeInfo.reason || '',
        scopeDebug: scopeInfo.debug || ''
      },
      index: {
        version: sourceInfo.indexVersion || '',
        builtAt: sourceInfo.indexBuiltAt || '',
        dirty: !!sourceInfo.indexDirty,
        sourceTotal: sourceInfo.rawTotal || scopedRows.length,
        ownTotal: analyzed.length,
        scopedTotal: analyzed.length,
        scope: isAllScope ? 'ALL' : 'OWN',
        sourceType: sourceInfo.sourceType || 'MASTER',
        sourceMessage: sourceInfo.message || ''
      },
      analysisDb: analysisDbResult,
      statusOptions: (PORTAL_CONFIG.STATUS_OPTIONS || []).slice(),
      statusCounts: objectToSortedStatusCountArrayP528_(statusCounts),
      cards: {
        total: analyzed.length,
        active: activeCount,
        needStatus: needStatusCount,
        mismatch: stateChangeRows.length,
        recent7: recentRows.length,
        noContact14: noContactRows.length,
        sentNoFollow: sentNoFollowRows.length,
        highPotential: highPotentialRows.length,
        dataMissing: dataMissingRows.length,
        terminal: terminalRows.length,
        contactIssue: contactIssueRows.length,
        contractCheck: contractCheckRows.length,
        longReconnect: longReconnectRows.length
      },
      rows: clientRows,
      lists: {},
      responseMode: 'SLIM_ROWS_ONLY_P532_MASTER_SOURCE',
      ai: (typeof getMyCustomerAiAnalysisHealthP537 === 'function' ? getMyCustomerAiAnalysisHealthP537({ light: true }) : null)
    };
  } catch (err) {
    const msg = err && err.message ? err.message : String(err || '');
    return {
      ok: false,
      version: MY_CUSTOMER_STATUS_ANALYZER_VERSION_P528,
      stage: stage,
      error: msg,
      message: msg,
      stack: err && err.stack ? String(err.stack).slice(0, 2000) : '',
      generatedAt: formatMyCustomerStatusDateTimeP528_(started),
      elapsedMs: new Date().getTime() - started.getTime()
    };
  }
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
  if (!isMyCustomerStatusRowAllowedP531_(detail, perm, aliases)) throw new Error('나의 고객 현황에서 수정할 수 있는 고객이 아닙니다. 관리자/admin은 전체 고객, 영업담당자는 본인 담당 고객만 수정할 수 있습니다.');
  const currentStatus = String((detail && detail.status) || '').trim();
  const contractCompletionInfoP536 = getMyCustomerContractCompletionInfoP536_(detail, getMyCustomerContractCompleteMapP536_());
  if (!isMyCustomerStatusAllowedTransitionP536_(currentStatus, newStatus, {
    contractCompleteReady: !!(contractCompletionInfoP536 && contractCompletionInfoP536.ready),
    strongFailure: newStatus === '수주실패'
  })) {
    throw new Error('나의 고객 현황 자동추천으로 허용되지 않는 상태 변경입니다. 현재값: ' + (currentStatus || '(공란)') + ', 변경값: ' + newStatus + '. 발주완료→계약완료는 계약번호 기준 서류 저장 3종이 모두 저장일 때만 가능합니다.');
  }
  if (expectedStatus && currentStatus !== expectedStatus) {
    throw new Error('현재 진행현황이 화면에 표시된 값과 다릅니다. 현재값: ' + (currentStatus || '(공란)'));
  }

  const res = saveCustomerPatchFastP473({
    rowNo: rowNo,
    customerNo: customerNo,
    values: { status: newStatus },
    expectedValues: { status: currentStatus },
    clientSaveSource: 'myCustomerStatus.analysisStatusApply.P534',
    source: 'myCustomerStatus.analysisStatusApply.P534',
    clientOperationId: String(payload.clientOperationId || makeMyCustomerStatusOperationIdP528_(customerNo, rowNo)),
    thinSave: true,
    fastMode: true,
    noSynchronousRefresh: true
  });
  if (res && (res.queuedFallbackP473 || res.conflictP473 || res.applied === false)) {
    return Object.assign({}, res || {}, {
      ok: false,
      rowNo: rowNo,
      customerNo: customerNo,
      oldStatus: currentStatus,
      newStatus: newStatus,
      source: 'myCustomerStatus.analysisStatusApply.P534',
      fastApply: false,
      noFullReanalysis: true,
      message: res.message || '상태 저장이 즉시 완료되지 않았습니다. 저장큐/충돌 상태를 확인해 주세요.'
    });
  }
  const analysisDbUpdate = markMyCustomerStatusAnalysisAppliedP534_(rowNo, customerNo, currentStatus, newStatus, perm);
  return Object.assign({}, res || {}, {
    ok: true,
    rowNo: rowNo,
    customerNo: customerNo,
    oldStatus: currentStatus,
    newStatus: newStatus,
    analysisDbUpdate: analysisDbUpdate,
    source: 'myCustomerStatus.analysisStatusApply.P534',
    fastApply: true,
    noFullReanalysis: true
  });
}

function markMyCustomerStatusAnalysisAppliedP534_(rowNo, customerNo, oldStatus, newStatus, perm) {
  try {
    const ss = getWebAppDbSpreadsheet_();
    const sheet = ss.getSheetByName(MY_CUSTOMER_STATUS_ANALYSIS_SHEET_P529);
    if (!sheet || sheet.getLastRow() < 2) return { ok: false, reason: 'analysisSheetEmpty' };
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
    const idx = {};
    headers.forEach(function(h, i) { if (h) idx[h] = i; });
    const rowIdx = idx['rowNo'];
    const customerIdx = idx['고객번호'];
    if (rowIdx == null && customerIdx == null) return { ok: false, reason: 'missingKeyHeaders' };
    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    const nowText = formatMyCustomerStatusDateTimeP528_(new Date());
    const reviewer = String((perm && (perm.name || perm.displayName || perm.email)) || Session.getActiveUser().getEmail() || '').trim();
    let updated = 0;
    const targetCustomer = normalizeCustomerNoForKey_(customerNo);
    const targetRow = Number(rowNo || 0) || 0;

    function setIfExists(arr, header, value) {
      if (idx[header] != null) arr[idx[header]] = value;
    }

    values.forEach(function(arr, i) {
      const rNo = Number(arr[rowIdx] || 0) || 0;
      const cNo = normalizeCustomerNoForKey_(arr[customerIdx] || '');
      const rowMatches = targetRow && rNo === targetRow;
      const customerMatches = targetCustomer && cNo === targetCustomer;
      if (!rowMatches && !customerMatches) return;
      if (targetRow && rNo && rNo !== targetRow) return;
      if (targetCustomer && cNo && cNo !== targetCustomer) return;

      setIfExists(arr, '현재상태', newStatus);
      setIfExists(arr, '추천상태', newStatus);
      setIfExists(arr, '불일치여부', 'N');
      setIfExists(arr, '검토상태', '적용완료');
      setIfExists(arr, '검토자', reviewer);
      setIfExists(arr, '검토일시', nowText);
      setIfExists(arr, '검토메모', '나의 고객 현황 추천상태 적용: ' + (oldStatus || '(공란)') + ' → ' + newStatus);
      setIfExists(arr, '상태변경추천상태', '');
      setIfExists(arr, '상태변경추천여부', 'N');
      setIfExists(arr, '상태변경추천등급', '');
      setIfExists(arr, '상태변경추천허용', 'N');
      setIfExists(arr, '추천노출여부', 'N');
      setIfExists(arr, '상태보호여부', isMyCustomerStatusProtectedStatusP530_(newStatus) ? 'Y' : 'N');
      setIfExists(arr, '상태추천차단사유', '사용자가 추천상태를 적용하여 즉시 반영됨');
      setIfExists(arr, '보수판정사유', 'P534 fast apply: 전체 재분석 없이 해당 고객 상태만 반영');
      sheet.getRange(2 + i, 1, 1, lastCol).setValues([arr]);
      updated++;
    });
    return { ok: true, updated: updated, sheetName: MY_CUSTOMER_STATUS_ANALYSIS_SHEET_P529, updatedAt: nowText };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err || '') };
  }
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

function getMyCustomerStatusSourceRowsP532_(perm, aliases, isAllScope) {
  const info = {
    rows: [],
    rawTotal: 0,
    sourceType: 'MASTER',
    message: '',
    indexVersion: '',
    indexBuiltAt: '',
    indexDirty: false
  };
  try {
    if (typeof getCustomerSearchIndexMeta === 'function') {
      const meta = getCustomerSearchIndexMeta();
      info.indexVersion = meta && meta.version || '';
      info.indexBuiltAt = meta && meta.builtAt || '';
      info.indexDirty = !!(meta && meta.dirty);
    }
  } catch (err) {
    info.message = '검색인덱스 메타 조회 실패: ' + (err && err.message ? err.message : String(err || ''));
  }

  const masterObjects = getMasterObjects_();
  info.rawTotal = masterObjects.length;
  const rows = masterObjects.map(function(obj) {
    return makeMyCustomerStatusRowFromMasterObjectP532_(obj);
  }).filter(function(row) {
    return row && row.rowNo && normalizeCustomerNoForKey_(row.customerNo || '');
  });

  if (isAllScope) {
    info.rows = rows;
    info.message = (info.message ? info.message + ' / ' : '') + '마스터시트 전체 기준 분석';
    return info;
  }

  info.rows = rows.filter(function(row) { return isMyCustomerStatusOwnRowP528_(row, aliases); });
  info.message = (info.message ? info.message + ' / ' : '') + '마스터시트 영업담당자 기준 분석';
  return info;
}

function makeMyCustomerStatusRowFromMasterObjectP532_(obj) {
  obj = obj || {};
  const memo = getCustomerListValue_(obj, 'memo') || '';
  const customerNo = getCustomerListValue_(obj, 'customerNo') || obj['고객번호'] || '';
  const row = {
    __source: 'master',
    __summary: false,
    __masterBacked: true,
    rowNo: Number(obj.__rowNo || 0) || obj.__rowNo || '',
    customerNo: customerNo,
    orderNo: getCustomerListValue_(obj, 'orderNo') || '',
    company: getCustomerListValue_(obj, 'company') || '',
    salesRep: getCustomerListValue_(obj, 'salesRep') || '',
    status: getCustomerListValue_(obj, 'status') || '',
    customerRank: getCustomerListValue_(obj, 'customerRank') || '',
    contact: getCustomerListValue_(obj, 'contact') || '',
    phone: getCustomerListValue_(obj, 'phone') || '',
    directPhone: getCustomerListValue_(obj, 'directPhone') || '',
    email: getCustomerListValue_(obj, 'email') || '',
    vendor: getCustomerListValue_(obj, 'vendor') || '',
    finalQuote: getCustomerListValue_(obj, 'finalQuote') || '',
    memo: memo,
    fullMemo: memo,
    memoSummary: shortenMyCustomerStatusTextP528_(getRecentMemoTextP528_(extractMyCustomerMemoEventsP528_(memo), memo, 260) || memo, 260),
    address: getCustomerListValue_(obj, 'address') || '',
    fullAddress: getCustomerListValue_(obj, 'address') || '',
    area: getCustomerIndexObjectValueK2_(obj, 'area') || '',
    grade: getCustomerIndexObjectValueK2_(obj, 'grade') || '',
    buildingType: getCustomerIndexObjectValueK2_(obj, 'buildingType') || '',
    contractUnit: getCustomerIndexObjectValueK2_(obj, 'contractUnit') || '',
    appointment: getCustomerIndexObjectValueK2_(obj, 'appointment') || '',
    maintenance: getCustomerIndexObjectValueK2_(obj, 'maintenance') || '',
    performance: getCustomerIndexObjectValueK2_(obj, 'performance') || '',
    vat: getCustomerIndexObjectValueK2_(obj, 'vat') || '',
    discountRate: getCustomerIndexObjectValueK2_(obj, 'discountRate') || '',
    specialTerms: getCustomerIndexObjectValueK2_(obj, 'specialTerms') || '',
    lastSent: getCustomerIndexObjectValueK2_(obj, 'lastSent') || '',
    sendCount: getCustomerIndexObjectValueK2_(obj, ['발송횟수', '발송 횟수']) || '',
    sendStatus: getCustomerIndexObjectValueK2_(obj, ['발송상태', '발송 상태']) || '',
    sentAt: getCustomerIndexObjectValueK2_(obj, 'sentAt') || '',
    memoInferredStatus: getCustomerIndexObjectValueK2_(obj, ['메모상 추측 상태값', '메모상추측상태값', '메모 추측 상태값']) || '',
    statusMatch: getCustomerIndexObjectValueK2_(obj, ['상태값 일치 여부', '상태값일치여부']) || '',
    tmProgressStatus: getCustomerIndexObjectValueK2_(obj, 'tmProgressStatus') || '',
    tmContactContent: getCustomerIndexObjectValueK2_(obj, 'tmContactContent') || ''
  };
  return row;
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
  // P538: 컨택이력_DB의 실제 연결키는 `고객번호`입니다.
  // `마스터행`은 과거 고객번호 공란 데이터에만 쓰는 legacy fallback이며,
  // 고객번호가 있는 컨택이력을 rowNo 기준으로 다른 고객에게 붙이면 안 됩니다.
  const result = { byCustomerNo: {}, byRowNo: {}, count: 0, excludedCustomerNoMismatch: 0, legacyRowFallbackCount: 0 };
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
      const customerNoMatched = !!(cno && customerKeys[cno]);
      const legacyRowMatched = !!(!cno && rn && rowKeys[rn]);
      if (!customerNoMatched && !legacyRowMatched) {
        if (cno && rn && rowKeys[rn]) result.excludedCustomerNoMismatch += 1;
        return;
      }
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
      if (customerNoMatched) {
        if (!result.byCustomerNo[cno]) result.byCustomerNo[cno] = [];
        result.byCustomerNo[cno].push(ev);
        result.count += 1;
        return;
      }
      if (legacyRowMatched) {
        if (!result.byRowNo[rn]) result.byRowNo[rn] = [];
        result.byRowNo[rn].push(ev);
        result.legacyRowFallbackCount += 1;
        result.count += 1;
      }
    });
  } catch (err) {
    try { Logger.log('나의 고객 현황 컨택이력_DB 조회 실패: ' + (err && err.stack || err)); } catch (e) {}
  }
  return result;
}


function getMyCustomerContractCompleteMapP536_() {
  const out = { byContractNo: {}, count: 0, error: '' };
  try {
    if (typeof listContractCompleteRowsV69 !== 'function') {
      out.error = 'listContractCompleteRowsV69 함수가 없어 수주확정/계약완료 시트를 조회하지 못했습니다.';
      return out;
    }
    const res = listContractCompleteRowsV69({ force: false });
    const rows = (res && res.rows) || [];
    rows.forEach(function(item) {
      const key = normalizeMyCustomerContractNoP536_(item && item.contractNo);
      if (!key) return;
      if (!out.byContractNo[key]) {
        out.byContractNo[key] = item;
        out.count += 1;
      }
    });
  } catch (err) {
    out.error = err && err.message ? err.message : String(err || '');
    try { Logger.log('나의 고객 현황 수주확정/계약완료 lookup 실패: ' + (err && err.stack || err)); } catch (e) {}
  }
  return out;
}

function normalizeMyCustomerContractNoP536_(value) {
  return String(value == null ? '' : value).replace(/\u00a0/g, ' ').replace(/,/g, '').replace(/\s+/g, '').trim();
}

function isMyCustomerContractDocSavedP536_(value) {
  const v = String(value == null ? '' : value).replace(/\s+/g, '').trim();
  return v === '저장' || v === 'Y' || v === 'TRUE' || v === '완료' || v === 'O' || v === '○';
}

function getMyCustomerContractCompletionInfoP536_(row, contractCompleteMap) {
  row = row || {};
  contractCompleteMap = contractCompleteMap || { byContractNo: {}, error: '' };
  const orderNo = normalizeMyCustomerContractNoP536_(row.orderNo || row['발주번호'] || row.contractNo || '');
  const customerNo = normalizeCustomerNoForKey_(row.customerNo || row['고객번호'] || '');
  const info = {
    orderNo: orderNo,
    contractNo: '',
    contractRowNo: '',
    customerNo: customerNo,
    matched: false,
    ready: false,
    businessRegSaved: '',
    contractSaved: '',
    appointmentReportSaved: '',
    customerNoMatches: '',
    error: contractCompleteMap.error || '',
    summary: ''
  };
  if (!orderNo) {
    info.summary = '마스터시트 발주번호가 없어 수주확정/계약완료 시트와 매칭할 수 없습니다.';
    return info;
  }
  const item = contractCompleteMap.byContractNo && contractCompleteMap.byContractNo[orderNo];
  if (!item) {
    info.summary = '수주확정/계약완료 시트에서 계약번호 ' + orderNo + ' 행을 찾지 못했습니다.';
    return info;
  }
  info.matched = true;
  info.contractNo = String(item.contractNo || orderNo || '').trim();
  info.contractRowNo = item.rowNo || '';
  const itemCustomerNo = normalizeCustomerNoForKey_(item.customerNo || '');
  info.customerNoMatches = (!customerNo || !itemCustomerNo || customerNo === itemCustomerNo) ? 'Y' : 'N';
  info.businessRegSaved = String(item.businessRegSaved || '').trim();
  info.contractSaved = String(item.contractSaved || '').trim();
  info.appointmentReportSaved = String(item.orderMailSaved || item.appointmentReportSaved || '').trim();
  const docsReady = isMyCustomerContractDocSavedP536_(info.businessRegSaved) &&
    isMyCustomerContractDocSavedP536_(info.contractSaved) &&
    isMyCustomerContractDocSavedP536_(info.appointmentReportSaved);
  info.ready = docsReady && info.customerNoMatches !== 'N';
  if (info.customerNoMatches === 'N') {
    info.summary = '계약번호 ' + orderNo + '는 일치하지만 고객번호가 다릅니다. 자동 계약완료 추천을 보류합니다.';
  } else if (info.ready) {
    info.summary = '계약번호 ' + orderNo + '의 사업자등록증 저장·계약서 저장·선임신고서 저장이 모두 저장입니다.';
  } else {
    const missing = [];
    if (!isMyCustomerContractDocSavedP536_(info.businessRegSaved)) missing.push('사업자등록증 저장=' + (info.businessRegSaved || '-'));
    if (!isMyCustomerContractDocSavedP536_(info.contractSaved)) missing.push('계약서 저장=' + (info.contractSaved || '-'));
    if (!isMyCustomerContractDocSavedP536_(info.appointmentReportSaved)) missing.push('선임신고서 저장=' + (info.appointmentReportSaved || '-'));
    info.summary = '계약번호 ' + orderNo + ' 필수서류 미완료: ' + missing.join(', ');
  }
  return info;
}

function buildMyCustomerStatusAnalyzedRowP528_(row, now, contactMap, contractCompleteMap) {
  row = row || {};
  contactMap = contactMap || { byCustomerNo: {}, byRowNo: {} };
  contractCompleteMap = contractCompleteMap || { byContractNo: {} };
  const memo = String(row.fullMemo || row.memo || '');
  const status = String(row.status || '').trim();
  const sentAt = String(row.sentAt || row.lastSent || '');
  const customerNo = normalizeCustomerNoForKey_(row.customerNo || '');
  const rowNoKey = String(row.rowNo || '');
  const contactEventsByCustomerNo = customerNo && contactMap.byCustomerNo[customerNo] ? contactMap.byCustomerNo[customerNo] : [];
  const legacyContactEventsByRowNo = (!contactEventsByCustomerNo.length && rowNoKey && contactMap.byRowNo[rowNoKey]) ? contactMap.byRowNo[rowNoKey] : [];
  const contactEvents = [].concat(contactEventsByCustomerNo, legacyContactEventsByRowNo);
  const dedupedContactEvents = dedupeMyCustomerEventsP529_(contactEvents);
  const events = buildMyCustomerCombinedEventsP529_(row, memo, dedupedContactEvents);
  const latestAnyEvent = getLatestMyCustomerMemoEventP528_(events);
  const latestDecisionEvent = getLatestMyCustomerDecisionEventP531_(events);
  const latestDate = latestDecisionEvent && latestDecisionEvent.date ? latestDecisionEvent.date : (latestAnyEvent && latestAnyEvent.date ? latestAnyEvent.date : parseLoosePortalDateP528_(sentAt));
  const lastContactDays = latestDate ? daysBetweenPortalDatesP528_(latestDate, now) : null;
  const contractCompletionInfo = getMyCustomerContractCompletionInfoP536_(row, contractCompleteMap);
  const analysis = classifyMyCustomerMemoP528_(row, events, latestAnyEvent, latestDecisionEvent, lastContactDays, now, contractCompletionInfo);
  const latestText = latestDecisionEvent && latestDecisionEvent.text ? latestDecisionEvent.text : (latestAnyEvent && latestAnyEvent.text ? latestAnyEvent.text : getLastMemoLineP528_(memo));
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
    orderNo: String(row.orderNo || ''),
    contractCompletionInfo: contractCompletionInfo,
    memoSummary: shortenMyCustomerStatusTextP528_(latestText || memo, 220),
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

function classifyMyCustomerMemoP528_(row, events, latestAnyEvent, latestDecisionEvent, lastContactDays, now, contractCompletionInfo) {
  row = row || {};
  events = Array.isArray(events) ? events : [];
  events.forEach(function(ev) { enrichMyCustomerEventClassP531_(ev); });
  contractCompletionInfo = contractCompletionInfo || {};

  const currentStatus = String(row.status || '').trim();
  const statusLabel = currentStatus || '!!상태지정필요!!';
  const hardProtected = isMyCustomerStatusHardProtectedP536_(currentStatus);
  const orderStatus = currentStatus === '발주완료';
  const contractReady = !!contractCompletionInfo.ready;
  const stateProtected = hardProtected || (orderStatus && !contractReady);
  const sentDone = !!(String(row.sendStatus || '').indexOf('발송완료') >= 0 || row.lastSent || row.sentAt);

  const decisionEvents = events.filter(function(ev) { return ev && ev.eventClass === 'CUSTOMER_INTENT'; });
  const statusEvents = events.filter(function(ev) { return ev && ev.eventClass !== 'SYSTEM_META'; });
  const signalSummary = summarizeMyCustomerSignalsP529_(statusEvents);
  const decisionSummary = summarizeMyCustomerSignalsP529_(decisionEvents);

  const thirdPartyContract = decisionSummary.thirdPartyContract.latest;
  const fail = decisionSummary.fail.latest;
  const contactBlocker = decisionSummary.contactBlocker.latest || signalSummary.contactBlocker.latest;
  const ownComplete = decisionSummary.complete.latest;
  const unknownComplete = decisionSummary.unknownComplete.latest;
  const order = decisionSummary.order.latest;
  const long = decisionSummary.long.latest;
  const quote = decisionSummary.quote.latest || signalSummary.quote.latest;
  const active = decisionSummary.active.latest;
  const data = signalSummary.data.latest;
  const failLike = getLatestMyCustomerSignalP530_([thirdPartyContract, fail]);
  const hasThirdPartyContract = !!thirdPartyContract;
  const latestDecision = latestDecisionEvent || getLatestMyCustomerDecisionEventP531_(events);
  const latestAny = latestAnyEvent || getLatestMyCustomerMemoEventP528_(events);

  let statusChangeRecommendedStatus = currentStatus || '';
  let statusChangeRecommendation = false;
  let statusChangeRecommendationGrade = '';
  let statusRecommendationBlockedReason = '';
  let confidence = 0;
  let priorityRank = '';
  let insightType = '';
  let insightSummary = '';
  let reason = '';
  let nextAction = '';
  let terminalCandidate = false;
  let analysisCautionLevel = '';
  let conservativeReason = '';

  function setInsight(type, summary, action, conf, priority) {
    if (!insightType) insightType = type || '';
    if (!insightSummary) insightSummary = summary || '';
    if (!nextAction) nextAction = action || '';
    if (!confidence && conf) confidence = conf;
    if (!priorityRank && priority) priorityRank = priority;
  }
  function setStatusRecommendation(target, grade, summary, action, conf, priority) {
    statusChangeRecommendedStatus = target || currentStatus || '';
    statusChangeRecommendation = !!target && target !== currentStatus;
    statusChangeRecommendationGrade = grade || '';
    insightType = grade || insightType;
    insightSummary = summary || insightSummary;
    reason = summary || reason;
    nextAction = action || nextAction;
    confidence = conf || confidence;
    priorityRank = priority || priorityRank;
  }
  function recommendIfAllowed(target, grade, summary, action, conf, priority, extra) {
    extra = extra || {};
    if (isMyCustomerStatusAllowedTransitionP536_(currentStatus, target, {
      contractCompleteReady: contractReady,
      strongFailure: !!(extra.strongFailure || failLike),
      orderNo: row.orderNo || '',
      contractCompletionInfo: contractCompletionInfo
    })) {
      setStatusRecommendation(target, grade, summary, action, conf, priority);
      return true;
    }
    statusChangeRecommendedStatus = currentStatus;
    statusRecommendationBlockedReason = statusRecommendationBlockedReason || getMyCustomerStatusTransitionBlockReasonP536_(currentStatus, target, contractCompletionInfo);
    setInsight(extra.insightType || grade || '상태확인필요', summary, action, Math.min(conf || 70, 80), priority);
    return false;
  }

  if (hardProtected) {
    statusChangeRecommendedStatus = currentStatus;
    statusChangeRecommendation = false;
    confidence = 98;
    priorityRank = '0_상태보호';
    terminalCandidate = true;
    statusRecommendationBlockedReason = currentStatus + '은/는 영업담당자가 직접 판단한 주요 상태이므로 나의 고객 현황 자동추천으로 변경하지 않습니다.';
    conservativeReason = statusRecommendationBlockedReason;
    if (currentStatus === '수주실패' && (hasThirdPartyContract || fail)) {
      setInsight('종결상태보호', hasThirdPartyContract ? '타사 계약/타사 결정 계열 이력이 있어 현재 수주실패 상태와 정합합니다.' : '거절/진행불가 계열 이력이 있어 현재 수주실패 상태와 정합합니다.', '추가 영업 대상에서 제외하거나 장기 재접촉 여부만 참고하세요.', 98, '0_수주실패보호');
    } else if (currentStatus === '계약완료') {
      setInsight('종결상태보호', '계약완료는 수주 확정 및 필수서류 취합 완료 상태이므로 현재 상태를 우선 신뢰합니다.', '필요 시 계약종합관리/수주확정 시트와 수동 대조하세요.', 98, '0_계약완료보호');
    } else {
      setInsight('종결상태보호', statusRecommendationBlockedReason, '필요 시 고객상세에서 직접 확인하세요.', 98, '0_상태보호');
    }
  } else if (orderStatus) {
    terminalCandidate = true;
    if (contractReady) {
      setStatusRecommendation('계약완료', '계약완료 전환 추천', '현재 발주완료 상태이며, 수주확정/계약완료 시트에서 계약번호 ' + (contractCompletionInfo.contractNo || row.orderNo || '-') + '의 사업자등록증 저장·계약서 저장·선임신고서 저장이 모두 저장 상태입니다.', '계약완료 적용 후 계약완료 고객으로 관리하세요.', 96, '1_발주완료계약완료');
    } else {
      statusChangeRecommendedStatus = currentStatus;
      statusChangeRecommendation = false;
      confidence = 94;
      priorityRank = '0_발주완료보호';
      statusRecommendationBlockedReason = '발주완료는 수주 확정 상태입니다. 계약완료 추천은 수주확정/계약완료 시트의 사업자등록증 저장·계약서 저장·선임신고서 저장이 모두 저장일 때만 허용합니다.';
      setInsight('발주완료서류확인', contractCompletionInfo.summary || '발주완료 상태입니다. 필수서류 저장 3종이 모두 저장인지 확인하세요.', '필수서류가 모두 저장되면 계약완료 전환 대상입니다.', 94, '0_발주완료보호');
    }
  } else if (thirdPartyContract) {
    terminalCandidate = true;
    recommendIfAllowed('수주실패', '수주실패 전환 추천', '최신 유효 이력에 타사 계약/타사 결정/다른 업체 선정 신호가 있습니다. 타사 계약완료는 우리 계약완료가 아니라 수주실패 신호입니다.', '수주실패 처리 여부를 확인한 뒤 적용하세요.', 94, '2_타사계약수주실패', { insightType: '타사선정/수주실패확인', strongFailure: true });
  } else if (fail && !hasPositiveSignalAfterP529_(decisionSummary, fail)) {
    terminalCandidate = true;
    recommendIfAllowed('수주실패', '수주실패 전환 추천', '최신 유효 이력에 명확한 거절/진행불가 신호가 있습니다.', '수주실패 처리 여부를 확인한 뒤 적용하세요.', 90, '2_수주실패', { insightType: '수주실패확인', strongFailure: true });
  } else if (contactBlocker) {
    statusChangeRecommendedStatus = currentStatus;
    statusRecommendationBlockedReason = '전화연결/담당자확인/직통번호 안내불가류는 수주실패가 아니라 연락장애로 분류합니다.';
    setInsight('연락장애/담당자확인', '연락 경로 또는 담당자 확인 문제가 감지되었습니다. 수주실패로 자동 추천하지 않습니다.', '담당자/직통번호/관리사무소 경로를 보강한 뒤 재연락하세요.', 76, '6_연락장애');
  } else if (long) {
    recommendIfAllowed('장기 추진건', '장기 추진건 전환 추천', '최신 유효 이력에 내년/예산/금년도 대상 아님/추후 재접촉 등 장기 추진 신호가 있습니다.', '장기 추진건으로 정리하고 재접촉 시점을 메모나 오늘 할 일에 남기세요.', 82, '3_장기추진', { insightType: '장기재접촉' });
  } else if (ownComplete || order) {
    statusChangeRecommendedStatus = currentStatus;
    statusRecommendationBlockedReason = '발주완료는 사람이 확정해야 하는 수주 상태이므로 나의 고객 현황에서 자동 추천하지 않습니다. 단, 발주완료→계약완료만 서류 저장 기준으로 추천합니다.';
    analysisCautionLevel = '확인필요';
    setInsight('계약/발주확인필요', ownComplete ? '당사 계약/계약완료 처리로 보일 수 있는 신호가 있습니다. 단, 계약완료는 발주완료 상태에서 필수서류 3종 저장이 확인될 때만 추천합니다.' : '용역신청서/계약서류/사업자등록증/발주 신호가 있습니다. 발주완료 여부는 담당자가 직접 확정해야 합니다.', '계약/발주 자료를 확인하고 필요한 경우 고객상세에서 직접 상태를 수정하세요.', ownComplete ? 86 : 78, ownComplete ? '1_우리계약확인' : '4_발주계약확인');
    terminalCandidate = !!ownComplete;
  } else if ((quote || sentDone) && !failLike && !ownComplete && !order && !thirdPartyContract) {
    if (quote && isMyCustomerQuoteLatestEnoughForStatusP531_(quote, failLike, ownComplete, order, thirdPartyContract)) {
      recommendIfAllowed('견적제출완료', '견적제출완료 전환 추천', '견적/자료 발송 이력이 있어 영업팀의 후속 컨택이 필요한 상태입니다.', '견적제출완료 적용 후 후속 연락 일정을 잡으세요.', 78, '5_견적제출상태추천', { insightType: '자료발송후미후속' });
    } else {
      statusChangeRecommendedStatus = currentStatus;
      statusRecommendationBlockedReason = '자료발송/견적발송은 확인됐지만 상태전이 기준에 맞지 않아 후속연락 인사이트로만 표시합니다.';
      setInsight('자료발송후미후속', '견적/자료 발송 이력이 있습니다. 발송 후 후속 연락 여부를 확인하세요.', '견적 확인 연락 또는 재발송 필요 여부를 확인하세요.', quote ? 72 : 66, '5_견적후속');
    }
  } else if (active) {
    recommendIfAllowed('고객 설득 중', '고객 설득 중 전환 추천', '검토/비교/재확인/담당자 전달 등 영업팀 추가 컨택이 필요한 진행중 신호가 있습니다.', '고객 설득 중으로 정리하고 다음 컨택 액션을 남기세요.', 70, '6_고객설득중', { insightType: '진행중추적' });
  } else if (unknownComplete) {
    statusChangeRecommendedStatus = currentStatus;
    statusRecommendationBlockedReason = '계약완료 표현은 있으나 타사 계약인지 당사 계약인지 불명확하여 상태변경 추천을 보류합니다.';
    analysisCautionLevel = '확인필요';
    setInsight('계약주체확인필요', '계약완료 표현은 있으나 계약 주체가 불명확합니다.', '타사 계약인지 당사 계약인지 확인하세요.', 60, '7_주체불명계약완료');
  } else if (data) {
    statusChangeRecommendedStatus = currentStatus;
    setInsight('데이터확인필요', '중복/주소/연면적/연락처 등 데이터 확인 신호가 있습니다.', '고객정보를 먼저 정리하세요.', 62, '8_데이터확인');
  } else {
    statusChangeRecommendedStatus = currentStatus;
    setInsight('', '', '', 0, '');
  }

  const transitionAllowed = isMyCustomerStatusAllowedTransitionP536_(currentStatus, statusChangeRecommendedStatus, {
    contractCompleteReady: contractReady,
    strongFailure: !!failLike,
    contractCompletionInfo: contractCompletionInfo
  });
  const recommendationVisible = !!(statusChangeRecommendation && transitionAllowed && statusChangeRecommendedStatus && statusChangeRecommendedStatus !== currentStatus && confidence >= 70 && isMyCustomerStatusAllowedAutoTargetP536_(statusChangeRecommendedStatus));
  if (!recommendationVisible && statusChangeRecommendation) {
    statusChangeRecommendation = false;
    statusRecommendationBlockedReason = statusRecommendationBlockedReason || getMyCustomerStatusTransitionBlockReasonP536_(currentStatus, statusChangeRecommendedStatus, contractCompletionInfo);
    statusChangeRecommendedStatus = currentStatus;
  }

  const activeStatus = isMyCustomerStatusActiveP528_(currentStatus);
  const longNoContact = activeStatus && (lastContactDays == null || lastContactDays >= 14);
  let sentNoFollow = false;
  if (sentDone && activeStatus && !hardProtected) {
    if (lastContactDays == null) sentNoFollow = true;
    else if (lastContactDays >= 3 && !order && !failLike && !ownComplete && !thirdPartyContract) sentNoFollow = true;
  }
  const potentialScore = stateProtected ? 0 : calculateMyCustomerPotentialScoreP528_(row, signalSummary.flags, lastContactDays, recommendationVisible ? statusChangeRecommendedStatus : currentStatus);
  const tags = [];
  Object.keys(signalSummary.flags).forEach(function(k) { if (signalSummary.flags[k]) tags.push(k); });
  if (longNoContact) tags.push('장기미접촉');
  if (sentNoFollow) tags.push('자료발송후미후속');
  if (terminalCandidate || stateProtected) tags.push('종결성신호');
  if (stateProtected) tags.push('상태보호');
  if (hasThirdPartyContract) tags.push('타사계약');
  if (unknownComplete) tags.push('계약주체불명');
  if (contactBlocker) tags.push('연락장애');
  if (contractReady) tags.push('계약완료전환가능');

  const latestAnyText = latestAny && latestAny.text ? latestAny.text : '';
  const latestDecisionText = latestDecision && latestDecision.text ? latestDecision.text : '';
  const excludedEvents = events.filter(function(ev) { return ev && ev.eventClass === 'SYSTEM_META'; }).slice(0, 5);
  const eventClassSummary = summarizeMyCustomerEventClassesP531_(events);
  const transitionBlockReason = statusRecommendationBlockedReason || (!recommendationVisible && statusChangeRecommendedStatus && statusChangeRecommendedStatus !== currentStatus ? getMyCustomerStatusTransitionBlockReasonP536_(currentStatus, statusChangeRecommendedStatus, contractCompletionInfo) : '');

  return {
    recommendedStatus: recommendationVisible ? statusChangeRecommendedStatus : statusLabel,
    statusChangeRecommendedStatus: recommendationVisible ? statusChangeRecommendedStatus : '',
    statusChangeRecommendation: recommendationVisible,
    statusChangeRecommendationGrade: recommendationVisible ? statusChangeRecommendationGrade : '',
    confidence: confidence,
    priorityRank: priorityRank,
    mismatch: recommendationVisible,
    mismatchType: recommendationVisible ? statusChangeRecommendationGrade : '',
    terminalCandidate: terminalCandidate || stateProtected,
    reason: insightSummary || reason || conservativeReason || '',
    insightType: insightType || '',
    insightSummary: insightSummary || '',
    nextAction: nextAction || '',
    lastContactDays: lastContactDays,
    longNoContact: longNoContact,
    sentNoFollow: sentNoFollow,
    potentialScore: potentialScore,
    tags: tags,
    latestMemoDate: latestDecision && latestDecision.date ? formatMyCustomerStatusDateP528_(latestDecision.date) : (latestAny && latestAny.date ? formatMyCustomerStatusDateP528_(latestAny.date) : ''),
    latestMemoText: shortenMyCustomerStatusTextP528_(latestDecisionText || latestAnyText, 220),
    latestEventSource: latestDecision ? latestDecision.sourceLabel : (latestAny ? latestAny.sourceLabel : ''),
    latestAnyEventText: shortenMyCustomerStatusTextP528_(latestAnyText, 220),
    latestAnyEventSource: latestAny ? latestAny.sourceLabel : '',
    latestDecisionEventText: shortenMyCustomerStatusTextP528_(latestDecisionText, 220),
    latestDecisionEventDate: latestDecision && latestDecision.date ? formatMyCustomerStatusDateP528_(latestDecision.date) : '',
    latestDecisionEventSource: latestDecision ? latestDecision.sourceLabel : '',
    signalSummary: signalSummary,
    matchedKeywords: signalSummary.matchedKeywords,
    sourceLabels: signalSummary.sourceLabels,
    statusProtected: stateProtected,
    stateChangeAllowed: transitionAllowed,
    canApplyRecommendation: recommendationVisible,
    recommendationVisible: recommendationVisible,
    contractCompleteType: contractReady ? '발주완료_서류3종저장완료' : getMyCustomerContractCompleteTypeP530_(decisionSummary),
    contractCompletionInfo: contractCompletionInfo,
    thirdPartyContract: hasThirdPartyContract,
    analysisCautionLevel: analysisCautionLevel,
    conservativeReason: conservativeReason || '',
    statusRecommendationBlockedReason: transitionBlockReason || '',
    eventClassSummary: eventClassSummary,
    excludedEventSummary: excludedEvents.map(function(ev) { return shortenMyCustomerStatusTextP528_(ev.text, 120); }).join(' / '),
    latestAnyEvent: eventForClientP529_(latestAny || {}),
    latestDecisionEvent: eventForClientP529_(latestDecision || {})
  };
}

function isMyCustomerStatusProtectedStatusP530_(status) {
  status = String(status || '').trim();
  return MY_CUSTOMER_STATUS_BLOCKING_STATUSES_P528.indexOf(status) >= 0;
}

function isMyCustomerStatusHardProtectedP536_(status) {
  status = String(status || '').trim();
  return MY_CUSTOMER_STATUS_HARD_PROTECTED_STATUSES_P536.indexOf(status) >= 0;
}

function isMyCustomerStatusStateChangeCandidateP530_(status) {
  status = String(status || '').trim();
  if (isMyCustomerStatusHardProtectedP536_(status)) return false;
  if (!status) return true;
  if (status.indexOf('상태지정') >= 0) return true;
  return ['고객 설득 중', '견적제출완료', '장기 추진건', '발주완료'].indexOf(status) >= 0;
}

function isMyCustomerStatusAllowedTransitionP536_(currentStatus, targetStatus, context) {
  currentStatus = String(currentStatus || '').trim();
  targetStatus = String(targetStatus || '').trim();
  context = context || {};
  if (!targetStatus || targetStatus === currentStatus) return false;
  if (targetStatus === '발주완료') return false;
  if (currentStatus === '수주실패' || currentStatus === '계약완료') return false;
  if (currentStatus === '발주완료') return targetStatus === '계약완료' && !!context.contractCompleteReady;
  const needStatus = !currentStatus || currentStatus.indexOf('상태지정') >= 0;
  if (needStatus) return ['수주실패', '장기 추진건', '견적제출완료', '고객 설득 중'].indexOf(targetStatus) >= 0;
  if (currentStatus === '고객 설득 중') return ['수주실패', '장기 추진건', '견적제출완료'].indexOf(targetStatus) >= 0;
  if (currentStatus === '견적제출완료') return ['수주실패', '장기 추진건'].indexOf(targetStatus) >= 0;
  if (currentStatus === '장기 추진건') return targetStatus === '수주실패' && !!context.strongFailure;
  return false;
}

function isMyCustomerStatusAllowedAutoTargetP536_(status) {
  status = String(status || '').trim();
  return ['수주실패', '견적제출완료', '장기 추진건', '고객 설득 중', '계약완료'].indexOf(status) >= 0;
}

function getMyCustomerStatusTransitionBlockReasonP536_(currentStatus, targetStatus, contractCompletionInfo) {
  currentStatus = String(currentStatus || '').trim();
  targetStatus = String(targetStatus || '').trim();
  if (!targetStatus) return '';
  if (targetStatus === '발주완료') return '발주완료는 영업담당자가 수주 확정 후 수기로 지정하는 상태이므로 자동추천하지 않습니다.';
  if (currentStatus === '수주실패' || currentStatus === '계약완료') return currentStatus + '은/는 보호 상태라 자동추천으로 변경하지 않습니다.';
  if (currentStatus === '발주완료' && targetStatus === '계약완료') {
    return contractCompletionInfo && contractCompletionInfo.summary ? contractCompletionInfo.summary : '발주완료→계약완료는 수주확정/계약완료 시트의 사업자등록증 저장·계약서 저장·선임신고서 저장이 모두 저장일 때만 허용합니다.';
  }
  if (currentStatus === '장기 추진건' && targetStatus !== '수주실패') return '장기 추진건은 명확한 수주실패 신호 외에는 자동 상태변경하지 않습니다.';
  return '현재상태에서 해당 추천상태로 자동 전환하는 기준에 맞지 않아 업무 인사이트로만 표시합니다.';
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
  const out = {
    complete: [], thirdPartyContract: [], unknownComplete: [], fail: [], contactBlocker: [],
    order: [], long: [], quote: [], active: [], data: []
  };
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
      { kw: '타업체 선정', label: '타업체 선정' }, { kw: '타사 선정', label: '타사 선정' },
      { kw: '다른 업체로 결정', label: '다른 업체로 결정' }, { kw: '다른곳으로 결정', label: '다른 곳으로 결정' }, { kw: '다른 곳으로 결정', label: '다른 곳으로 결정' },
      { kw: '기존 업체 유지', label: '기존업체 유지' }, { kw: '기존업체 유지', label: '기존업체 유지' }, { kw: '금액 낮은 곳으로 계약', label: '저가 타업체 계약' },
      { kw: '장애인사업장하고 진행', label: '타 기관 진행' }, { kw: '업체 선정 완료', label: '타 업체 선정 완료' },
      { re: /타(사|업체).{0,12}(결정|계약|진행|완료|선정)/, label: '타사 결정/계약' },
      { re: /(다른|기존).{0,6}(업체|곳).{0,12}(계약|진행|결정|유지|완료)/, label: '다른 업체 계약/결정' },
      { re: /(가격|금액).{0,10}(낮|저렴).{0,12}(곳|업체).{0,12}(계약|진행|결정)/, label: '저가 업체 계약/결정' }
    ],
    unknownComplete: [
      { kw: '계약완료', label: '계약완료 주체불명' }, { kw: '계약 완료', label: '계약완료 주체불명' }
    ],
    fail: [
      { kw: '수주실패' }, { kw: '계약 취소' }, { kw: '계약취소' }, { kw: '진행 취소' },
      { kw: '계약 못한다고', label: '계약 못한다고 함' }, { kw: '계약못한다고', label: '계약 못한다고 함' },
      { kw: '진행 안한다고', label: '진행 안 한다고 함' }, { kw: '진행 안 한다고', label: '진행 안 한다고 함' },
      { kw: '안한다고', label: '고객 안 한다고 함' }, { kw: '안 한다고', label: '고객 안 한다고 함' },
      { kw: '필요없다고', label: '필요없다고 함' }, { kw: '필요 없다고', label: '필요없다고 함' },
      { kw: '하지 말', label: '진행 거절' }, { kw: '연락하지', label: '연락 거절' },
      { kw: '대상 아님' }, { kw: '대상아님' }, { kw: '폐업' },
      { re: /(계약|진행|제안).{0,8}(거절|불가|안함|안 함)/, label: '계약/진행 거절' }
    ],
    contactBlocker: [
      { kw: '전화연결 불가' }, { kw: '전화 연결 불가' }, { kw: '연결 불가' }, { kw: '전화 안받' }, { kw: '전화 안 받' },
      { kw: '직원 연결 불가' }, { kw: '직통번호 안내불가' }, { kw: '직통번호 안내 불가' }, { kw: '안내 불가' },
      { kw: '담당자 부재' }, { kw: '자리비움' }, { kw: '자리 비움' }, { kw: '내용 잘 모름' }, { kw: '잘 모른다고' },
      { kw: '연락처 안내 거절' }, { kw: '관리사무소 번호 문의하니 거절' }, { kw: '메일 확인 불가' }, { kw: '담당자 확인 필요' }
    ],
    order: [
      { kw: '용역신청서 회신' }, { kw: '용역신청서 받' }, { kw: '계약서 요청' }, { kw: '계약서류 요청' }, { kw: '계약서류 회신' },
      { kw: '계약 진행' }, { kw: '계약진행' }, { kw: '사업자등록증 요청' }, { kw: '사업자 등록증 요청' }, { kw: '발주 요청' }, { kw: '발주 예정' },
      { kw: '진행하시기로' }, { kw: '계약하시기로' }, { kw: '서류 받' }, { kw: '서류 요청' }, { kw: '취합' }
    ],
    long: [
      { kw: '내년' }, { kw: '내년도' }, { kw: '2027년' }, { kw: '올해 대상 아님' }, { kw: '금년도 대상 아님' },
      { kw: '예산 없음' }, { kw: '예산없' }, { kw: '예산 잡히면' }, { kw: '예산 편성' },
      { kw: '하반기 이후' }, { kw: '추후 검토' }, { kw: '내년에 다시 연락' }, { kw: '장기' }, { kw: '보류' }, { kw: '재검토' }
    ],
    quote: [
      { kw: '견적서 발송' }, { kw: '견적서발송' }, { kw: '견적 발송' }, { kw: '견적발송' }, { kw: '자료발송' }, { kw: '자료 발송' },
      { kw: '단가표' }, { kw: '비교견적' }, { kw: '수기견적' }, { kw: '메일 발송' }, { kw: '발송완료' }, { kw: '견적서 재발송' }, { kw: '재견적' }
    ],
    active: [
      { kw: '검토중' }, { kw: '검토 중' }, { kw: '비교중' }, { kw: '비교 중' }, { kw: '타업체와 비교' }, { kw: '타사 비교' },
      { kw: '담당자 전달' }, { kw: '내부 검토' }, { kw: '확인 후' }, { kw: '재확인' }, { kw: '연락 예정' }, { kw: '연락준다고' },
      { kw: '가격 조율' }, { kw: '가격차이' }, { kw: '가격 차이' }, { kw: '금액 차이' }, { kw: '본사' }, { kw: '회신' }, { kw: '상담' }
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
  const categories = ['complete', 'thirdPartyContract', 'unknownComplete', 'fail', 'contactBlocker', 'order', 'long', 'quote', 'active', 'data'];
  const summary = {
    flags: {},
    matchedKeywords: {},
    sourceLabels: [],
    complete: { latest: null, count: 0, keywords: [] },
    thirdPartyContract: { latest: null, count: 0, keywords: [] },
    unknownComplete: { latest: null, count: 0, keywords: [] },
    fail: { latest: null, count: 0, keywords: [] },
    contactBlocker: { latest: null, count: 0, keywords: [] },
    order: { latest: null, count: 0, keywords: [] },
    long: { latest: null, count: 0, keywords: [] },
    quote: { latest: null, count: 0, keywords: [] },
    active: { latest: null, count: 0, keywords: [] },
    data: { latest: null, count: 0, keywords: [] }
  };
  (events || []).forEach(function(ev) {
    if (!ev) return;
    enrichMyCustomerEventClassP531_(ev);
    if (ev.sourceLabel && summary.sourceLabels.indexOf(ev.sourceLabel) < 0) summary.sourceLabels.push(ev.sourceLabel);
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
  const ad = coerceMyCustomerStatusDateP535_(a && a.date);
  const bd = coerceMyCustomerStatusDateP535_(b && b.date);
  const at = ad ? ad.getTime() : 0;
  const bt = bd ? bd.getTime() : 0;
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
  const ad = coerceMyCustomerStatusDateP535_(a);
  const bd = coerceMyCustomerStatusDateP535_(b);
  if (!ad || !bd) return null;
  const da = new Date(ad.getFullYear(), ad.getMonth(), ad.getDate()).getTime();
  const db = new Date(bd.getFullYear(), bd.getMonth(), bd.getDate()).getTime();
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

function coerceMyCustomerStatusDateP535_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value;
  if (typeof value === 'number') {
    const numDate = new Date(value);
    if (!isNaN(numDate.getTime())) return numDate;
    return null;
  }

  const text = String(value || '').trim();
  if (!text) return null;

  function makeDate_(year, month, day, hour, minute, second, ampm) {
    year = Number(year || 0);
    month = Number(month || 0);
    day = Number(day || 0);
    if (!year || !month || !day) return null;
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;
    let h = Number(hour || 0) || 0;
    const mi = Number(minute || 0) || 0;
    const se = Number(second || 0) || 0;
    ampm = String(ampm || '').trim();
    if (ampm === '오후' && h < 12) h += 12;
    if (ampm === '오전' && h === 12) h = 0;
    return new Date(year, month - 1, day, h, mi, se);
  }

  let m = text.match(/^(20\d{2})[-.\/]\s*(\d{1,2})[-.\/]\s*(\d{1,2})(?:[ T]+(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?)?(?:.*)?$/);
  if (m) return makeDate_(m[1], m[2], m[3], m[4], m[5], m[6], '');

  m = text.match(/^(20\d{2})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})\s*(?:일)?\s*(오전|오후)?\s*(?:(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?)?(?:.*)?$/);
  if (m) return makeDate_(m[1], m[2], m[3], m[5], m[6], m[7], m[4]);

  const loose = parseLoosePortalDateP528_(text);
  if (loose) return loose;

  if (/^20\d{2}-\d{1,2}-\d{1,2}T/.test(text)) {
    const iso = new Date(text);
    if (!isNaN(iso.getTime())) return iso;
  }
  return null;
}

function formatMyCustomerStatusDateP528_(value) {
  if (!value) return '';
  const date = coerceMyCustomerStatusDateP535_(value);
  if (date) return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(value || '').trim();
}

function formatMyCustomerStatusDateTimeP528_(value) {
  if (!value) return '';
  const date = coerceMyCustomerStatusDateP535_(value);
  if (date) return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  return String(value || '').trim();
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
  return 'MY_STATUS_P531_' + String(customerNo || '') + '_' + String(rowNo || '') + '_' + String(Date.now());
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

function saveMyCustomerStatusAnalysisRowsP529_(rows, perm, aliases, analyzedAt, options) {
  rows = Array.isArray(rows) ? rows : [];
  options = options || {};
  const sheet = ensureMyCustomerStatusAnalysisSheetP529_();
  const lastCol = Math.max(sheet.getLastColumn(), MY_CUSTOMER_STATUS_ANALYSIS_HEADERS_P529.length);
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const headerMap = {};
  headers.forEach(function(h, i) { if (h) headerMap[h] = i; });
  const allScope = !!options.allScope || canMyCustomerStatusViewAllP531_(perm);
  const ownerAliases = {};
  (aliases || []).forEach(function(a) { const k = normalizeMyCustomerStatusNameP528_(a); if (k) ownerAliases[k] = true; });
  const ownerName = String((perm && (perm.salesRepName || perm.name || perm.displayName)) || '').trim();
  const ownerKey = normalizeMyCustomerStatusNameP528_(ownerName);
  if (ownerKey) ownerAliases[ownerKey] = true;

  const existing = [];
  const lastRow = sheet.getLastRow();
  if (!allScope && lastRow >= 2) {
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
  const combined = allScope ? newRows : existing.concat(newRows);

  if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  const batchSize = MY_CUSTOMER_STATUS_ANALYSIS_WRITE_BATCH_P532 || 120;
  let written = 0;
  for (let start = 0; start < combined.length; start += batchSize) {
    const chunk = combined.slice(start, start + batchSize);
    if (!chunk.length) continue;
    sheet.getRange(2 + start, 1, chunk.length, lastCol).setValues(chunk);
    written += chunk.length;
    try { SpreadsheetApp.flush(); } catch (err) {}
  }
  return {
    saved: true,
    rows: newRows.length,
    totalRows: combined.length,
    written: written,
    batchSize: batchSize,
    sheetName: MY_CUSTOMER_STATUS_ANALYSIS_SHEET_P529,
    spreadsheetUrl: getWebAppDbSpreadsheet_().getUrl(),
    savedAt: nowText,
    scope: allScope ? 'ALL' : 'OWN',
    sourceType: options.sourceInfo && options.sourceInfo.sourceType || 'MASTER'
  };
}


function buildMyCustomerAnalysisDbRowP529_(row, analyzedAtText, headers) {
  row = row || {};
  const a = row.analysis || {};
  const latestEvent = (a.events && a.events[0]) || {};
  const evidence = {
    reason: a.reason || '',
    insightType: a.insightType || '',
    insightSummary: a.insightSummary || '',
    nextAction: shortenMyCustomerStatusTextP528_(a.nextAction || '', 180),
    sourceLabels: a.sourceLabels || [],
    matchedKeywords: a.matchedKeywords || {},
    latestAnyEvent: a.latestAnyEvent || latestEvent,
    latestDecisionEvent: a.latestDecisionEvent || {},
    contractCompletionInfo: a.contractCompletionInfo || {}
  };
  const sourceLabels = (a.sourceLabels || []).join(', ');
  const statusRecommendation = a.statusChangeRecommendation ? 'Y' : 'N';
  const recommendedForDb = a.statusChangeRecommendation ? a.statusChangeRecommendedStatus : (row.status || '');
  const rec = {
    '분석일시': analyzedAtText,
    '분석ID': 'MCSA-' + String(row.customerNo || '') + '-' + String(row.rowNo || '') + '-' + String(a.memoHash || '').slice(0, 8),
    '고객번호': row.customerNo || '',
    'rowNo': row.rowNo || '',
    '회사명': row.company || '',
    '영업담당자': row.salesRep || '',
    '현재상태': row.status || '',
    '추천상태': recommendedForDb,
    '판정유형': a.statusChangeRecommendation ? (a.statusChangeRecommendationGrade || '상태변경 추천') : (a.insightType || ''),
    '신뢰도': a.confidence || '',
    '우선순위등급': a.priorityRank || '',
    '불일치여부': a.statusChangeRecommendation ? 'Y' : 'N',
    '분석소스': sourceLabels,
    '최근연락일': row.lastContactDate || a.latestMemoDate || '',
    '최근이벤트출처': a.latestDecisionEventSource || a.latestEventSource || latestEvent.sourceLabel || '',
    '최근이벤트요약': a.latestDecisionEventText || a.latestMemoText || row.memoSummary || '',
    '강한긍정키워드': [].concat((a.matchedKeywords && a.matchedKeywords.complete) || [], (a.matchedKeywords && a.matchedKeywords.order) || [], (a.matchedKeywords && a.matchedKeywords.active) || []).join(', '),
    '강한부정키워드': [].concat((a.matchedKeywords && a.matchedKeywords.thirdPartyContract) || [], (a.matchedKeywords && a.matchedKeywords.fail) || []).join(', '),
    '계약진행키워드': ((a.matchedKeywords && a.matchedKeywords.order) || []).join(', '),
    '자료발송키워드': ((a.matchedKeywords && a.matchedKeywords.quote) || []).join(', '),
    '장기추진키워드': ((a.matchedKeywords && a.matchedKeywords.long) || []).join(', '),
    '데이터누락키워드': (a.missingFields || []).join(', '),
    '가능성점수': a.potentialScore || 0,
    '위험태그': (a.tags || []).join(', '),
    '추천액션': a.nextAction || '',
    '근거JSON': safeStringifyMyCustomerStatusP529_(evidence, MY_CUSTOMER_STATUS_DB_EVIDENCE_JSON_LIMIT_P532),
    '이벤트JSON': safeStringifyMyCustomerStatusP529_((a.events || []).slice(0, 8).map(eventForClientP529_), MY_CUSTOMER_STATUS_DB_EVENTS_JSON_LIMIT_P532),
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
    '보수판정사유': a.conservativeReason || '',
    '상태변경추천상태': a.statusChangeRecommendedStatus || '',
    '상태변경추천여부': statusRecommendation,
    '상태변경추천등급': a.statusChangeRecommendationGrade || '',
    '업무인사이트유형': a.insightType || '',
    '업무인사이트요약': a.insightSummary || '',
    '최신전체이벤트요약': a.latestAnyEventText || '',
    '최신판정이벤트요약': a.latestDecisionEventText || '',
    '최신판정이벤트일자': a.latestDecisionEventDate || '',
    '최신판정이벤트출처': a.latestDecisionEventSource || '',
    '이벤트분류요약': a.eventClassSummary || '',
    '제외이벤트요약': a.excludedEventSummary || '',
    '상태추천차단사유': a.statusRecommendationBlockedReason || ''
  };
  return headers.map(function(h) { return rec[h] == null ? '' : rec[h]; });
}


function getMyCustomerStatusDashboardFromDbP539_(options) {
  options = options || {};
  const started = new Date();
  let stage = 'start';
  const perf = { readMs: 0, mapMs: 0, buildMs: 0, totalMs: 0 };
  try {
    stage = 'permission';
    const perm = getPortalCurrentPermission_();
    const aliases = buildMyCustomerStatusOwnerAliasesP528_(perm);
    const scopeInfo = resolveMyCustomerStatusScopeP532_(perm);
    const isAllScope = !!scopeInfo.isAllScope;

    stage = 'readAnalysisDb';
    const readStarted = new Date();
    const dbResult = readMyCustomerStatusAnalysisDbRowsP539_(perm, aliases, isAllScope);
    perf.readMs = new Date().getTime() - readStarted.getTime();

    stage = 'mapClientRows';
    const mapStarted = new Date();
    const clientRows = dbResult.rows.map(buildMyCustomerStatusClientRowFromDbP539_);
    perf.mapMs = new Date().getTime() - mapStarted.getTime();

    stage = 'buildFastResponse';
    const buildStarted = new Date();
    const response = buildMyCustomerStatusDashboardResponseFromClientRowsP539_(clientRows, {
      started: started,
      perm: perm,
      aliases: aliases,
      scopeInfo: scopeInfo,
      isAllScope: isAllScope,
      dbResult: dbResult,
      sourceType: 'ANALYSIS_DB_FAST_COLUMNS',
      sourceMessage: dbResult.message || ''
    });
    perf.buildMs = new Date().getTime() - buildStarted.getTime();
    perf.totalMs = new Date().getTime() - started.getTime();
    response.perf = Object.assign({}, perf, {
      rawRows: dbResult.rawRows || 0,
      scopedRows: dbResult.scopedRows || 0,
      selectedColumnCount: dbResult.selectedColumnCount || 0,
      blockCount: dbResult.blockCount || 0
    });
    logMyCustomerStatusPerfP540_('myCustomerStatus.dbFastLoad.end', perf.totalMs, '', {
      stage: stage,
      scope: isAllScope ? 'ALL' : 'OWN',
      isAdmin: !!isAllScope,
      rawRows: dbResult.rawRows || 0,
      scopedRows: dbResult.scopedRows || 0,
      selectedColumnCount: dbResult.selectedColumnCount || 0,
      blockCount: dbResult.blockCount || 0,
      missingHeaders: dbResult.missingHeaders || [],
      skippedHeavyHeaders: dbResult.skippedHeavyHeaders || [],
      readMs: perf.readMs,
      mapMs: perf.mapMs,
      buildMs: perf.buildMs
    });
    return response;
  } catch (err) {
    const msg = err && err.message ? err.message : String(err || '');
    const elapsed = new Date().getTime() - started.getTime();
    logMyCustomerStatusPerfP540_('myCustomerStatus.dbFastLoad.fail', elapsed, msg, {
      stage: stage,
      error: msg,
      stack: err && err.stack ? String(err.stack).slice(0, 1200) : ''
    });
    return {
      ok: false,
      version: MY_CUSTOMER_STATUS_ANALYZER_VERSION_P528,
      stage: stage,
      error: msg,
      message: msg,
      stack: err && err.stack ? String(err.stack).slice(0, 2000) : '',
      generatedAt: formatMyCustomerStatusDateTimeP528_(started),
      elapsedMs: elapsed,
      responseMode: 'FAST_DB_LOAD_FAILED_P540'
    };
  }
}

function readMyCustomerStatusAnalysisDbRowsP539_(perm, aliases, isAllScope) {
  const out = {
    ok: true,
    sheetName: MY_CUSTOMER_STATUS_ANALYSIS_SHEET_P529,
    rows: [],
    rawRows: 0,
    scopedRows: 0,
    latestAnalyzedAt: '',
    message: '',
    error: '',
    selectedColumnCount: 0,
    selectedHeaders: [],
    missingHeaders: [],
    blockCount: 0,
    skippedHeavyHeaders: []
  };
  const ss = getWebAppDbSpreadsheet_();
  const sheet = ss.getSheetByName(MY_CUSTOMER_STATUS_ANALYSIS_SHEET_P529);
  if (!sheet || sheet.getLastRow() < 2) {
    out.ok = true;
    out.message = '고객현황분석_DB가 없거나 분석 결과가 없습니다. 분석DB 갱신을 먼저 실행해 주세요.';
    return out;
  }
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const readResult = readMyCustomerStatusAnalysisFastColumnsP540_(sheet, headers, 2, lastRow - 1);
  const records = readResult.records || [];
  const aliasKeys = {};
  (aliases || []).forEach(function(a) { const k = normalizeMyCustomerStatusNameP528_(a); if (k) aliasKeys[k] = true; });
  out.rawRows = records.length;
  out.selectedColumnCount = readResult.selectedColumnCount || 0;
  out.selectedHeaders = readResult.selectedHeaders || [];
  out.missingHeaders = readResult.missingHeaders || [];
  out.blockCount = readResult.blockCount || 0;
  out.skippedHeavyHeaders = MY_CUSTOMER_STATUS_HEAVY_DB_HEADERS_P540.filter(function(h) { return headers.indexOf(h) >= 0 && out.selectedHeaders.indexOf(h) < 0; });

  records.forEach(function(rec) {
    const repKey = normalizeMyCustomerStatusNameP528_(rec['영업담당자'] || '');
    if (!isAllScope) {
      if (!repKey || !aliasKeys[repKey]) return;
    }
    out.rows.push(rec);
    const at = String(rec['분석일시'] || '').trim();
    if (at && (!out.latestAnalyzedAt || String(at) > String(out.latestAnalyzedAt))) out.latestAnalyzedAt = at;
  });
  out.scopedRows = out.rows.length;
  out.message = '고객현황분석_DB 기준 P541 얇은 컬럼 빠른 조회';
  return out;
}

function readMyCustomerStatusAnalysisFastColumnsP540_(sheet, headers, startRow, rowCount) {
  const result = {
    records: [],
    selectedColumnCount: 0,
    selectedHeaders: [],
    missingHeaders: [],
    blockCount: 0
  };
  rowCount = Math.max(0, Number(rowCount || 0) || 0);
  if (!sheet || !rowCount) return result;
  headers = (headers || []).map(function(h) { return String(h || '').trim(); });
  const headerToIndex = {};
  headers.forEach(function(h, i) { if (h && headerToIndex[h] == null) headerToIndex[h] = i; });
  const selectedIndexes = [];
  const seen = {};
  (MY_CUSTOMER_STATUS_FAST_DB_HEADERS_P540 || []).forEach(function(h) {
    const idx = headerToIndex[h];
    if (idx == null) {
      result.missingHeaders.push(h);
      return;
    }
    if (seen[idx]) return;
    seen[idx] = true;
    selectedIndexes.push(idx);
    result.selectedHeaders.push(h);
  });
  selectedIndexes.sort(function(a, b) { return a - b; });
  result.selectedColumnCount = selectedIndexes.length;
  const records = Array.from({ length: rowCount }, function() { return {}; });
  if (!selectedIndexes.length) {
    result.records = records;
    return result;
  }
  const blocks = [];
  let blockStart = selectedIndexes[0];
  let prev = selectedIndexes[0];
  for (let i = 1; i < selectedIndexes.length; i++) {
    const idx = selectedIndexes[i];
    if (idx === prev + 1) {
      prev = idx;
      continue;
    }
    blocks.push({ start: blockStart, end: prev });
    blockStart = idx;
    prev = idx;
  }
  blocks.push({ start: blockStart, end: prev });
  result.blockCount = blocks.length;
  blocks.forEach(function(block) {
    const width = block.end - block.start + 1;
    const values = sheet.getRange(startRow, block.start + 1, rowCount, width).getDisplayValues();
    for (let r = 0; r < rowCount; r++) {
      for (let c = 0; c < width; c++) {
        const header = headers[block.start + c];
        if (header) records[r][header] = values[r][c];
      }
    }
  });
  result.records = records;
  return result;
}

function buildMyCustomerStatusClientRowFromDbP539_(rec) {
  rec = rec || {};
  function yn(v) { return String(v || '').trim().toUpperCase() === 'Y'; }
  function n(v) { const x = Number(String(v || '').replace(/,/g, '')); return isNaN(x) ? 0 : x; }
  function split(v) { return String(v || '').split(/[,|]/).map(function(x) { return String(x || '').trim(); }).filter(Boolean); }
  const status = String(rec['현재상태'] || '').trim();
  const recommended = String(rec['추천상태'] || status || '').trim();
  const changeStatus = String(rec['상태변경추천상태'] || '').trim();
  const changeYn = yn(rec['상태변경추천여부']) && yn(rec['추천노출여부']) && !!changeStatus;
  const lastContact = String(rec['최근연락일'] || rec['최신판정이벤트일자'] || '').trim();
  const lastContactDays = lastContact ? daysBetweenPortalDatesP528_(coerceMyCustomerStatusDateP535_(lastContact), new Date()) : null;
  const missingFields = split(rec['데이터누락키워드'] || '');
  const tags = split(rec['위험태그'] || '');
  const sourceLabels = split(rec['분석소스'] || '');
  const matchedKeywords = {
    thirdPartyContract: split(rec['타사계약여부'] === 'Y' ? '타사계약' : ''),
    fail: split(rec['강한부정키워드'] || ''),
    complete: split(rec['강한긍정키워드'] || ''),
    order: split(rec['계약진행키워드'] || ''),
    quote: split(rec['자료발송키워드'] || ''),
    long: split(rec['장기추진키워드'] || ''),
    data: missingFields
  };
  const insightType = String(rec['업무인사이트유형'] || rec['판정유형'] || '').trim();
  const insightSummary = String(rec['업무인사이트요약'] || rec['추천액션'] || '').trim();
  return {
    rowNo: n(rec['rowNo']),
    customerNo: String(rec['고객번호'] || '').trim(),
    company: String(rec['회사명'] || '').trim(),
    salesRep: String(rec['영업담당자'] || '').trim(),
    status: status,
    vendor: '',
    contact: '',
    phone: '',
    directPhone: '',
    email: '',
    finalQuote: '',
    lastSent: '',
    sentAt: '',
    sendStatus: '',
    orderNo: '',
    memoSummary: shortenMyCustomerStatusTextP528_(String(rec['최근이벤트요약'] || rec['최신판정이벤트요약'] || rec['업무인사이트요약'] || ''), 180),
    lastContactDate: lastContact,
    analysis: {
      recommendedStatus: recommended || status,
      statusChangeRecommendedStatus: changeYn ? changeStatus : '',
      statusChangeRecommendation: changeYn,
      statusChangeRecommendationGrade: String(rec['상태변경추천등급'] || rec['판정유형'] || '').trim(),
      mismatch: changeYn,
      mismatchType: changeYn ? String(rec['상태변경추천등급'] || '').trim() : '',
      confidence: n(rec['신뢰도']),
      priorityRank: String(rec['우선순위등급'] || '').trim(),
      insightType: insightType,
      insightSummary: shortenMyCustomerStatusTextP528_(insightSummary, 180),
      reason: shortenMyCustomerStatusTextP528_(String(rec['보수판정사유'] || insightSummary || ''), 180),
      nextAction: shortenMyCustomerStatusTextP528_(String(rec['추천액션'] || ''), 180),
      lastContactDays: lastContactDays,
      longNoContact: (lastContactDays == null || Number(lastContactDays) >= 14) && ['수주실패','발주완료','계약완료'].indexOf(status) < 0,
      sentNoFollow: insightType === '자료발송후미후속' || insightType === '견적후속필요',
      potentialScore: n(rec['가능성점수']),
      tags: tags,
      latestMemoText: shortenMyCustomerStatusTextP528_(String(rec['최근이벤트요약'] || ''), 220),
      latestEventSource: String(rec['최근이벤트출처'] || '').trim(),
      latestAnyEventText: shortenMyCustomerStatusTextP528_(String(rec['최신전체이벤트요약'] || rec['최근이벤트요약'] || ''), 220),
      latestAnyEventSource: String(rec['최근이벤트출처'] || '').trim(),
      latestDecisionEventText: shortenMyCustomerStatusTextP528_(String(rec['최신판정이벤트요약'] || rec['최근이벤트요약'] || ''), 220),
      latestDecisionEventDate: String(rec['최신판정이벤트일자'] || lastContact || '').trim(),
      latestDecisionEventSource: String(rec['최신판정이벤트출처'] || rec['최근이벤트출처'] || '').trim(),
      matchedKeywords: matchedKeywords,
      sourceLabels: sourceLabels,
      missingFields: missingFields,
      dataMissing: missingFields.length > 0 || insightType === '데이터확인필요',
      statusProtected: yn(rec['상태보호여부']),
      stateChangeAllowed: yn(rec['상태변경추천허용']),
      canApplyRecommendation: changeYn,
      recommendationVisible: changeYn,
      terminalCandidate: yn(rec['상태보호여부']) || ['수주실패','발주완료','계약완료'].indexOf(status) >= 0,
      contractCompleteType: String(rec['계약완료구분'] || '').trim(),
      contractCompletionInfo: {},
      thirdPartyContract: yn(rec['타사계약여부']),
      analysisCautionLevel: String(rec['분석주의등급'] || '').trim(),
      conservativeReason: String(rec['보수판정사유'] || '').trim(),
      statusRecommendationBlockedReason: String(rec['상태추천차단사유'] || '').trim()
    }
  };
}

function buildMyCustomerStatusDashboardResponseFromClientRowsP539_(clientRows, ctx) {
  clientRows = Array.isArray(clientRows) ? clientRows : [];
  ctx = ctx || {};
  const started = ctx.started || new Date();
  const perm = ctx.perm || {};
  const scopeInfo = ctx.scopeInfo || { isAllScope: false };
  const isAllScope = !!ctx.isAllScope;
  const statusCounts = {};
  clientRows.forEach(function(row) {
    const key = row.status || '(공란)';
    statusCounts[key] = (statusCounts[key] || 0) + 1;
  });
  const cards = computeMyCustomerStatusCardsFromClientRowsP539_(clientRows);
  return {
    ok: true,
    version: MY_CUSTOMER_STATUS_ANALYZER_VERSION_P528,
    generatedAt: formatMyCustomerStatusDateTimeP528_(started),
    elapsedMs: new Date().getTime() - started.getTime(),
    owner: {
      name: isAllScope ? '전체 고객' : (perm.salesRepName || perm.name || ''),
      displayName: isAllScope ? ((perm.displayName || perm.name || perm.email || '') + ' · ADMIN 전체') : (perm.displayName || perm.name || perm.email || ''),
      email: perm.email || '',
      aliases: ctx.aliases || [],
      scope: isAllScope ? 'ALL' : 'OWN',
      isAdmin: !!isAllScope,
      scopeReason: scopeInfo.reason || '',
      scopeDebug: scopeInfo.debug || ''
    },
    index: {
      sourceTotal: (ctx.dbResult && ctx.dbResult.rawRows) || clientRows.length,
      ownTotal: clientRows.length,
      scopedTotal: clientRows.length,
      scope: isAllScope ? 'ALL' : 'OWN',
      sourceType: ctx.sourceType || 'ANALYSIS_DB',
      sourceMessage: ctx.sourceMessage || ''
    },
    analysisDb: {
      loaded: true,
      saved: false,
      rows: clientRows.length,
      rawRows: (ctx.dbResult && ctx.dbResult.rawRows) || 0,
      sheetName: MY_CUSTOMER_STATUS_ANALYSIS_SHEET_P529,
      latestAnalyzedAt: (ctx.dbResult && ctx.dbResult.latestAnalyzedAt) || '',
      message: (ctx.dbResult && ctx.dbResult.message) || '고객현황분석_DB 빠른 조회',
      fastLoad: true,
      selectedColumnCount: (ctx.dbResult && ctx.dbResult.selectedColumnCount) || 0,
      blockCount: (ctx.dbResult && ctx.dbResult.blockCount) || 0,
      missingHeaders: (ctx.dbResult && ctx.dbResult.missingHeaders) || [],
      skippedHeavyHeaders: (ctx.dbResult && ctx.dbResult.skippedHeavyHeaders) || []
    },
    job: null,
    asyncHealthPending: true,
    statusOptions: (PORTAL_CONFIG.STATUS_OPTIONS || []).slice(),
    statusCounts: objectToSortedStatusCountArrayP528_(statusCounts),
    cards: cards,
    rows: clientRows,
    lists: {},
    responseMode: 'FAST_ANALYSIS_DB_THIN_COLUMNS_P541',
    ai: null
  };
}

function getMyCustomerStatusAsyncHealthP540() {
  const started = new Date();
  let stage = 'start';
  try {
    stage = 'jobStatus';
    const job = getMyCustomerStatusAnalysisJobStatusP539({ light: true });
    stage = 'aiStatus';
    const ai = (typeof getMyCustomerAiAnalysisHealthP537 === 'function' ? getMyCustomerAiAnalysisHealthP537({ light: true }) : null);
    const elapsed = new Date().getTime() - started.getTime();
    logMyCustomerStatusPerfP540_('myCustomerStatus.asyncHealth.end', elapsed, '', {
      jobOk: !job || job.ok !== false,
      aiOk: !ai || ai.ok !== false,
      jobTrigger: job && job.triggerStatus || '',
      aiTrigger: ai && ai.triggerStatus || ''
    });
    return { ok: true, job: job, ai: ai, elapsedMs: elapsed, generatedAt: formatMyCustomerStatusDateTimeP528_(started) };
  } catch (err) {
    const msg = err && err.message ? err.message : String(err || '');
    const elapsed = new Date().getTime() - started.getTime();
    logMyCustomerStatusPerfP540_('myCustomerStatus.asyncHealth.fail', elapsed, msg, { stage: stage });
    return { ok: false, stage: stage, error: msg, elapsedMs: elapsed };
  }
}

function logMyCustomerStatusPerfP540_(eventName, durationMs, error, detail) {
  try {
    if (typeof appendPortalServerPerfLogP460_ !== 'function') return;
    appendPortalServerPerfLogP460_({
      event: eventName,
      durationMs: Number(durationMs || 0) || 0,
      page: 'myCustomerStatus',
      status: error ? 'error' : 'ok',
      error: error || '',
      detail: detail || {}
    });
  } catch (err) {
    try { Logger.log('P540 my customer status perf log failed: ' + (err && err.stack || err)); } catch (e) {}
  }
}

function computeMyCustomerStatusCardsFromClientRowsP539_(rows) {
  const cards = { total: rows.length, active: 0, needStatus: 0, mismatch: 0, recent7: 0, noContact14: 0, sentNoFollow: 0, highPotential: 0, dataMissing: 0, terminal: 0, contactIssue: 0, contractCheck: 0, longReconnect: 0 };
  rows.forEach(function(row) {
    const st = String(row.status || '').trim();
    const a = row.analysis || {};
    if (['수주실패','발주완료','계약완료'].indexOf(st) < 0) cards.active++;
    if (!st || st.indexOf('상태지정') >= 0) cards.needStatus++;
    if (a.statusChangeRecommendation) cards.mismatch++;
    if (a.lastContactDays != null && Number(a.lastContactDays) <= 7) cards.recent7++;
    if (a.longNoContact) cards.noContact14++;
    if (a.sentNoFollow || a.insightType === '자료발송후미후속' || a.insightType === '견적후속필요') cards.sentNoFollow++;
    if (Number(a.potentialScore || 0) >= 6) cards.highPotential++;
    if (a.dataMissing) cards.dataMissing++;
    if (a.terminalCandidate || a.statusProtected) cards.terminal++;
    if (a.insightType === '연락장애/담당자확인') cards.contactIssue++;
    if (a.insightType === '계약/발주확인필요') cards.contractCheck++;
    if (a.insightType === '장기재접촉') cards.longReconnect++;
  });
  return cards;
}

function ensureMyCustomerStatusAnalysisJobSheetP539_() {
  const ss = getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(MY_CUSTOMER_STATUS_ANALYSIS_JOB_SHEET_P539);
  if (!sheet) {
    sheet = ss.insertSheet(MY_CUSTOMER_STATUS_ANALYSIS_JOB_SHEET_P539);
    sheet.getRange(1, 1, 1, MY_CUSTOMER_STATUS_ANALYSIS_JOB_HEADERS_P539.length).setValues([MY_CUSTOMER_STATUS_ANALYSIS_JOB_HEADERS_P539]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, MY_CUSTOMER_STATUS_ANALYSIS_JOB_HEADERS_P539.length).setFontWeight('bold').setBackground('#f2f4f7');
    return sheet;
  }
  ensureSheetHeaders_(sheet, MY_CUSTOMER_STATUS_ANALYSIS_JOB_HEADERS_P539);
  return sheet;
}

function startMyCustomerStatusAnalysisJobP539(options) {
  options = options || {};
  const perm = getPortalCurrentPermission_();
  const aliases = buildMyCustomerStatusOwnerAliasesP528_(perm);
  const scopeInfo = resolveMyCustomerStatusScopeP532_(perm);
  const requestedScope = String(options.scope || '').toUpperCase();
  const allScope = requestedScope === 'ALL' ? !!scopeInfo.isAllScope : (requestedScope === 'OWN' ? false : !!scopeInfo.isAllScope);
  if (requestedScope === 'ALL' && !scopeInfo.isAllScope) throw new Error('전체 고객 분석DB 갱신은 admin/관리자만 실행할 수 있습니다.');
  const sourceInfo = getMyCustomerStatusSourceRowsP532_(perm, aliases, allScope);
  const total = (sourceInfo.rows || []).length;
  const sheet = ensureMyCustomerStatusAnalysisJobSheetP539_();
  const nowText = formatMyCustomerStatusDateTimeP528_(new Date());
  const jobId = 'MCSA-JOB-' + Utilities.getUuid().slice(0, 8) + '-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const requester = String(perm.email || perm.name || Session.getActiveUser().getEmail() || '').trim();
  const limit = Math.max(20, Math.min(500, Number(options.limit || MY_CUSTOMER_STATUS_ANALYSIS_JOB_PROCESS_LIMIT_P539) || MY_CUSTOMER_STATUS_ANALYSIS_JOB_PROCESS_LIMIT_P539));
  const rec = {
    '등록일시': nowText,
    '수정일시': nowText,
    '작업ID': jobId,
    '요청자': requester,
    '범위': allScope ? 'ALL' : 'OWN',
    '상태': total ? 'QUEUED' : 'DONE',
    '전체건수': total,
    '처리건수': 0,
    '성공건수': 0,
    '스킵건수': 0,
    '실패건수': 0,
    'nextOffset': 0,
    'limit': limit,
    'ownerAliasesJson': safeStringifyMyCustomerStatusP529_(aliases || [], 5000),
    'ownerEmail': String(perm.email || '').trim(),
    'scopeReason': scopeInfo.reason || '',
    'resultJson': safeStringifyMyCustomerStatusP529_({ sourceTotal: sourceInfo.rawTotal || 0, scopedTotal: total, createdBy: requester }, 3000),
    '마지막오류': '',
    '완료일시': total ? '' : nowText
  };
  const rowValues = MY_CUSTOMER_STATUS_ANALYSIS_JOB_HEADERS_P539.map(function(h) { return rec[h] == null ? '' : rec[h]; });
  sheet.appendRow(rowValues);
  return { ok: true, jobId: jobId, scope: rec['범위'], total: total, status: rec['상태'], limit: limit, sheetName: MY_CUSTOMER_STATUS_ANALYSIS_JOB_SHEET_P539, message: '분석DB 갱신 작업을 생성했습니다. 처리 버튼 또는 5분 트리거로 chunk 처리하세요.' };
}

function getMyCustomerStatusAnalysisJobRowsP539_() {
  const sheet = ensureMyCustomerStatusAnalysisJobSheetP539_();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const idx = {};
  headers.forEach(function(h, i) { if (h) idx[h] = i; });
  const rows = [];
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    values.forEach(function(arr, i) {
      const rec = { __rowNo: 2 + i, __arr: arr, __headers: headers, __idx: idx, __sheet: sheet };
      headers.forEach(function(h, j) { if (h) rec[h] = arr[j]; });
      rows.push(rec);
    });
  }
  return rows;
}

function processMyCustomerStatusAnalysisJobP539(options) {
  options = options || {};
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return { ok: false, error: '고객현황분석작업_DB 처리 lock을 획득하지 못했습니다. 잠시 후 다시 시도해 주세요.' };
  try {
    const jobs = getMyCustomerStatusAnalysisJobRowsP539_();
    let job = null;
    const requestedJobId = String(options.jobId || '').trim();
    for (let i = jobs.length - 1; i >= 0; i--) {
      const st = String(jobs[i]['상태'] || '').trim();
      if (requestedJobId && jobs[i]['작업ID'] !== requestedJobId) continue;
      if (['QUEUED','RUNNING','RETRY'].indexOf(st) >= 0) { job = jobs[i]; break; }
    }
    if (!job) return { ok: true, processed: 0, done: 0, message: '처리할 고객현황 분석 작업이 없습니다.' };
    return processMyCustomerStatusAnalysisJobRowP539_(job, options);
  } finally {
    try { lock.releaseLock(); } catch (err) {}
  }
}

function processMyCustomerStatusAnalysisJobTriggerP539() {
  return processMyCustomerStatusAnalysisJobP539({ trigger: true, limit: MY_CUSTOMER_STATUS_ANALYSIS_JOB_PROCESS_LIMIT_P539 });
}

function processMyCustomerStatusAnalysisJobRowP539_(job, options) {
  options = options || {};
  const sheet = job.__sheet;
  const headers = job.__headers;
  const idx = job.__idx;
  function setJob(header, value) {
    if (idx[header] == null) return;
    sheet.getRange(job.__rowNo, idx[header] + 1).setValue(value == null ? '' : value);
  }
  const now = new Date();
  const nowText = formatMyCustomerStatusDateTimeP528_(now);
  const limit = Math.max(10, Math.min(500, Number(options.limit || job['limit'] || MY_CUSTOMER_STATUS_ANALYSIS_JOB_PROCESS_LIMIT_P539) || MY_CUSTOMER_STATUS_ANALYSIS_JOB_PROCESS_LIMIT_P539));
  setJob('상태', 'RUNNING');
  setJob('수정일시', nowText);
  setJob('마지막오류', '');

  try {
    const isAllScope = String(job['범위'] || '').toUpperCase() === 'ALL';
    const aliases = parseMyCustomerStatusJsonP539_(job['ownerAliasesJson'], []);
    const sourceInfo = getMyCustomerStatusSourceRowsForJobP539_(isAllScope, aliases);
    const rows = sourceInfo.rows || [];
    const total = rows.length;
    const offset = Math.max(0, Number(job['nextOffset'] || 0) || 0);
    const chunk = rows.slice(offset, offset + limit);
    const contactMap = getMyCustomerContactHistoryMapP529_(chunk);
    const contractCompleteMap = getMyCustomerContractCompleteMapP536_();
    const analyzedAt = new Date();
    const analyzed = [];
    let failed = 0;
    chunk.forEach(function(row) {
      try { analyzed.push(buildMyCustomerStatusAnalyzedRowP528_(row, analyzedAt, contactMap, contractCompleteMap)); }
      catch (err) { failed += 1; try { Logger.log('고객현황 chunk 분석 실패: ' + (err && err.stack || err)); } catch (e) {} }
    });
    const upsert = upsertMyCustomerStatusAnalysisRowsP539_(analyzed, formatMyCustomerStatusDateTimeP528_(analyzedAt));
    const processed = offset + chunk.length;
    const done = processed >= total;
    const prevSuccess = Number(job['성공건수'] || 0) || 0;
    const prevFail = Number(job['실패건수'] || 0) || 0;
    setJob('전체건수', total);
    setJob('처리건수', processed);
    setJob('성공건수', prevSuccess + analyzed.length);
    setJob('실패건수', prevFail + failed);
    setJob('nextOffset', done ? total : processed);
    setJob('수정일시', nowText);
    setJob('상태', done ? 'DONE' : 'QUEUED');
    setJob('resultJson', safeStringifyMyCustomerStatusP529_({ lastChunk: chunk.length, analyzed: analyzed.length, failed: failed, upsert: upsert, total: total, processed: processed }, 5000));
    if (done) setJob('완료일시', nowText);
    return { ok: true, jobId: job['작업ID'], status: done ? 'DONE' : 'QUEUED', total: total, processed: processed, chunk: chunk.length, analyzed: analyzed.length, failed: failed, upsert: upsert, nextOffset: done ? total : processed };
  } catch (err) {
    const msg = err && err.message ? err.message : String(err || '');
    setJob('상태', 'RETRY');
    setJob('수정일시', nowText);
    setJob('마지막오류', msg);
    return { ok: false, jobId: job['작업ID'], error: msg, stack: err && err.stack ? String(err.stack).slice(0, 1500) : '' };
  }
}

function getMyCustomerStatusSourceRowsForJobP539_(isAllScope, aliases) {
  const info = { rows: [], rawTotal: 0, sourceType: 'MASTER_JOB', message: '' };
  const masterObjects = getMasterObjects_();
  info.rawTotal = masterObjects.length;
  const allRows = masterObjects.map(function(obj) { return makeMyCustomerStatusRowFromMasterObjectP532_(obj); }).filter(function(row) {
    return row && row.rowNo && normalizeCustomerNoForKey_(row.customerNo || '');
  });
  if (isAllScope) {
    info.rows = allRows;
    info.message = '고객현황분석작업_DB: 전체 고객';
  } else {
    info.rows = allRows.filter(function(row) { return isMyCustomerStatusOwnRowP528_(row, aliases || []); });
    info.message = '고객현황분석작업_DB: 영업담당자 범위';
  }
  return info;
}

function upsertMyCustomerStatusAnalysisRowsP539_(rows, analyzedAtText) {
  rows = Array.isArray(rows) ? rows : [];
  const sheet = ensureMyCustomerStatusAnalysisSheetP529_();
  const lastCol = Math.max(sheet.getLastColumn(), MY_CUSTOMER_STATUS_ANALYSIS_HEADERS_P529.length);
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const idx = {};
  headers.forEach(function(h, i) { if (h) idx[h] = i; });
  const existingMap = {};
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    values.forEach(function(arr, i) {
      const key = makeMyCustomerStatusAnalysisKeyP539_(arr[idx['rowNo']], arr[idx['고객번호']]);
      if (key) existingMap[key] = 2 + i;
    });
  }
  const appendRows = [];
  let updated = 0;
  rows.forEach(function(row) {
    const arr = buildMyCustomerAnalysisDbRowP529_(row, analyzedAtText, headers);
    const key = makeMyCustomerStatusAnalysisKeyP539_(row.rowNo, row.customerNo);
    const sheetRow = key && existingMap[key];
    if (sheetRow) {
      sheet.getRange(sheetRow, 1, 1, lastCol).setValues([arr]);
      updated += 1;
    } else {
      appendRows.push(arr);
    }
  });
  if (appendRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, lastCol).setValues(appendRows);
  }
  try { SpreadsheetApp.flush(); } catch (err) {}
  return { updated: updated, appended: appendRows.length, total: rows.length, sheetName: MY_CUSTOMER_STATUS_ANALYSIS_SHEET_P529 };
}

function makeMyCustomerStatusAnalysisKeyP539_(rowNo, customerNo) {
  const rn = String(rowNo || '').trim();
  const cn = normalizeCustomerNoForKey_(customerNo || '');
  if (!rn && !cn) return '';
  return rn + ':' + cn;
}

function parseMyCustomerStatusJsonP539_(value, fallback) {
  try { return JSON.parse(String(value || '')); } catch (err) { return fallback; }
}

function getMyCustomerStatusAnalysisJobStatusP539(options) {
  options = options || {};
  try {
    const rows = getMyCustomerStatusAnalysisJobRowsP539_();
    const byStatus = {};
    let latest = null;
    rows.forEach(function(r) {
      const st = String(r['상태'] || '(공란)').trim();
      byStatus[st] = (byStatus[st] || 0) + 1;
      if (!latest || String(r['등록일시'] || '') > String(latest['등록일시'] || '')) latest = r;
    });
    const triggers = ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction && t.getHandlerFunction() === 'processMyCustomerStatusAnalysisJobTriggerP539'; });
    return {
      ok: true,
      sheetName: MY_CUSTOMER_STATUS_ANALYSIS_JOB_SHEET_P539,
      totalJobs: rows.length,
      byStatus: byStatus,
      triggerCount: triggers.length,
      triggerStatus: triggers.length === 1 ? '정상' : (triggers.length > 1 ? '중복 ' + triggers.length + '개' : '미설치'),
      latest: latest ? {
        jobId: latest['작업ID'] || '', scope: latest['범위'] || '', status: latest['상태'] || '', total: Number(latest['전체건수'] || 0) || 0,
        processed: Number(latest['처리건수'] || 0) || 0, success: Number(latest['성공건수'] || 0) || 0, fail: Number(latest['실패건수'] || 0) || 0,
        nextOffset: Number(latest['nextOffset'] || 0) || 0, updatedAt: latest['수정일시'] || '', error: latest['마지막오류'] || ''
      } : null
    };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err || '') };
  }
}

function installMyCustomerStatusAnalysisJobTriggerP539() {
  const fn = 'processMyCustomerStatusAnalysisJobTriggerP539';
  const existing = ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction && t.getHandlerFunction() === fn; });
  if (!existing.length) ScriptApp.newTrigger(fn).timeBased().everyMinutes(5).create();
  return { ok: true, functionName: fn, existing: existing.length, installed: !existing.length };
}



function canMyCustomerStatusViewAllP531_(perm) {
  return !!resolveMyCustomerStatusScopeP532_(perm || getPortalCurrentPermission_()).isAllScope;
}

function resolveMyCustomerStatusScopeP532_(perm) {
  perm = perm || {};
  const email = String(perm.email || perm.emailKey || '').trim().toLowerCase();
  const level = String(perm.level || '').trim().toUpperCase();
  const role = String(perm.role || '').trim();
  const rank = String(perm.rank || '').trim();
  const scope = String(perm.defaultScope || '').trim().toUpperCase();
  const nameText = [perm.name, perm.displayName, perm.salesRepName, perm.note].map(function(v) { return String(v || ''); }).join(' ');
  const debug = [email, level, role, rank, scope, nameText].join(' | ');
  if (!perm || perm.active === false) return { isAllScope: false, reason: 'inactive', debug: debug };
  if (perm.isAdmin || perm.canUseAdminHome) return { isAllScope: true, reason: 'adminFlag', debug: debug };
  if (level === 'ADMIN' || level === 'MANAGER') return { isAllScope: true, reason: 'adminLevel', debug: debug };
  if (role.indexOf('서무') >= 0 || role.indexOf('총괄') >= 0 || role.indexOf('관리자') >= 0) return { isAllScope: true, reason: 'adminRole', debug: debug };
  if (nameText.indexOf('관리자') >= 0 || nameText.indexOf('서무') >= 0 || nameText.indexOf('총괄') >= 0) return { isAllScope: true, reason: 'adminNameOrNote', debug: debug };
  if (MY_CUSTOMER_STATUS_EXPLICIT_ADMIN_EMAILS_P532.indexOf(email) >= 0) return { isAllScope: true, reason: 'explicitAdminEmail', debug: debug };
  if (rank.indexOf('책임') >= 0 && role.indexOf('영업담당자') < 0) return { isAllScope: true, reason: 'adminRank', debug: debug };
  // 전체고객열람(Y)만으로는 나의 고객 현황을 전체 고객 범위로 열지 않습니다.
  // 단, SALES가 아닌 계정의 기본범위 ALL은 관리자성 계정으로 간주합니다.
  if (scope === 'ALL' && level !== 'SALES' && role.indexOf('영업담당자') < 0) return { isAllScope: true, reason: 'adminDefaultScope', debug: debug };
  return { isAllScope: false, reason: 'salesOwnOnly', debug: debug };
}

function debugMyCustomerStatusScopeP532() {
  const perm = getPortalCurrentPermission_();
  const aliases = buildMyCustomerStatusOwnerAliasesP528_(perm);
  const scopeInfo = resolveMyCustomerStatusScopeP532_(perm);
  let sourceInfo = null;
  let error = '';
  try {
    sourceInfo = getMyCustomerStatusSourceRowsP532_(perm, aliases, !!scopeInfo.isAllScope);
  } catch (err) {
    error = err && err.message ? err.message : String(err || '');
  }
  return {
    ok: true,
    version: MY_CUSTOMER_STATUS_ANALYZER_VERSION_P528,
    permission: {
      email: perm.email || '',
      name: perm.name || '',
      displayName: perm.displayName || '',
      role: perm.role || '',
      level: perm.level || '',
      rank: perm.rank || '',
      defaultScope: perm.defaultScope || '',
      salesRepName: perm.salesRepName || '',
      isAdmin: !!perm.isAdmin,
      canUseAdminHome: !!perm.canUseAdminHome,
      canViewAllCustomers: !!perm.canViewAllCustomers,
      canCompleteSupport: !!perm.canCompleteSupport,
      canWriteNotice: !!perm.canWriteNotice,
      active: perm.active !== false
    },
    aliases: aliases,
    scopeInfo: scopeInfo,
    sourceInfo: sourceInfo ? {
      rowCount: sourceInfo.rows ? sourceInfo.rows.length : 0,
      rawTotal: sourceInfo.rawTotal || 0,
      sourceType: sourceInfo.sourceType || '',
      message: sourceInfo.message || '',
      indexVersion: sourceInfo.indexVersion || '',
      indexBuiltAt: sourceInfo.indexBuiltAt || '',
      indexDirty: !!sourceInfo.indexDirty
    } : null,
    error: error
  };
}

function isMyCustomerStatusRowAllowedP531_(row, perm, aliases) {
  if (canMyCustomerStatusViewAllP531_(perm)) return true;
  return isMyCustomerStatusOwnRowP528_(row, aliases || buildMyCustomerStatusOwnerAliasesP528_(perm));
}

function isMyCustomerStatusStrictStateChangeBaseP531_(status) {
  status = String(status || '').trim();
  if (!status) return true;
  if (status.indexOf('상태지정') >= 0) return true;
  return ['고객 설득 중', '견적제출완료', '장기 추진건', '발주완료'].indexOf(status) >= 0;
}

function isMyCustomerStatusAllowedAutoTargetP531_(status) {
  return isMyCustomerStatusAllowedAutoTargetP536_(status);
}

function isMyCustomerQuoteLatestEnoughForStatusP531_(quote, failLike, ownComplete, order, thirdPartyContract) {
  if (!quote) return false;
  const blockers = [failLike, ownComplete, order, thirdPartyContract].filter(Boolean);
  return !blockers.some(function(ev) { return isMyCustomerSignalAfterP529_(ev, quote); });
}

function enrichMyCustomerEventClassP531_(ev) {
  if (!ev) return ev;
  if (ev.eventClass) return ev;
  const cls = classifyMyCustomerEventClassP531_(ev);
  ev.eventClass = cls.eventClass;
  ev.eventClassReason = cls.reason;
  return ev;
}

function classifyMyCustomerEventClassP531_(ev) {
  ev = ev || {};
  const source = String(ev.source || '');
  const text = normalizeMyCustomerSignalTextP529_(ev.text || ev.rawText || '');
  const systemMeta = [
    '중복 삭제', '데이터 병합', '기존 담당자별 시트', '담당자별 시트', '이관', '삭제 고객번호',
    '대표전화번호', '담당자 이메일 주소', '연면적', '정보 확인해보시고', 'tm 콜 원하시면',
    '상태지정필요로', '법정동코드 보정', '주소 보정', '폴더', 'renamed'
  ];
  for (let i = 0; i < systemMeta.length; i++) {
    if (text.indexOf(normalizeMyCustomerSignalTextP529_(systemMeta[i])) >= 0) {
      return { eventClass: 'SYSTEM_META', reason: systemMeta[i] };
    }
  }
  if (source === 'sendLog') return { eventClass: 'SEND_LOG', reason: '자료발송 스냅샷' };
  const sig = ev.signal || getMyCustomerEventSignalsP529_(text);
  if ((sig.thirdPartyContract && sig.thirdPartyContract.length) || (sig.fail && sig.fail.length) ||
      (sig.complete && sig.complete.length) || (sig.unknownComplete && sig.unknownComplete.length) ||
      (sig.order && sig.order.length) || (sig.long && sig.long.length) || (sig.active && sig.active.length) ||
      (sig.contactBlocker && sig.contactBlocker.length)) {
    return { eventClass: 'CUSTOMER_INTENT', reason: '고객의사/영업진행 신호' };
  }
  if ((sig.quote && sig.quote.length)) return { eventClass: 'SEND_LOG', reason: '견적/자료발송 신호' };
  if ((sig.data && sig.data.length)) return { eventClass: 'DATA_NOTE', reason: '데이터확인 신호' };
  return { eventClass: 'UNKNOWN', reason: '' };
}

function getLatestMyCustomerDecisionEventP531_(events) {
  let best = null;
  (events || []).forEach(function(ev) {
    if (!ev) return;
    enrichMyCustomerEventClassP531_(ev);
    if (ev.eventClass !== 'CUSTOMER_INTENT') return;
    if (!best || compareMyCustomerEventsP529_(ev, best) >= 0) best = ev;
  });
  return best;
}

function summarizeMyCustomerEventClassesP531_(events) {
  const counts = {};
  (events || []).forEach(function(ev) {
    enrichMyCustomerEventClassP531_(ev);
    const k = ev.eventClass || 'UNKNOWN';
    counts[k] = (counts[k] || 0) + 1;
  });
  return Object.keys(counts).sort().map(function(k) { return k + ':' + counts[k]; }).join(', ');
}

function buildMyCustomerStatusClientRowP531_(row) {
  row = row || {};
  const a = row.analysis || {};
  return {
    rowNo: row.rowNo || 0,
    customerNo: row.customerNo || '',
    company: row.company || '',
    salesRep: row.salesRep || '',
    status: row.status || '',
    vendor: row.vendor || '',
    contact: row.contact || '',
    phone: row.phone || '',
    directPhone: row.directPhone || '',
    email: row.email || '',
    finalQuote: row.finalQuote || '',
    lastSent: row.lastSent || '',
    sentAt: row.sentAt || '',
    sendStatus: row.sendStatus || '',
    orderNo: row.orderNo || '',
    memoSummary: shortenMyCustomerStatusTextP528_(row.memoSummary || '', 180),
    lastContactDate: row.lastContactDate || '',
    analysis: {
      recommendedStatus: a.recommendedStatus || row.status || '',
      statusChangeRecommendedStatus: a.statusChangeRecommendedStatus || '',
      statusChangeRecommendation: !!a.statusChangeRecommendation,
      statusChangeRecommendationGrade: a.statusChangeRecommendationGrade || '',
      mismatch: !!a.statusChangeRecommendation,
      mismatchType: a.statusChangeRecommendationGrade || '',
      confidence: a.confidence || 0,
      priorityRank: a.priorityRank || '',
      insightType: a.insightType || '',
      insightSummary: shortenMyCustomerStatusTextP528_(a.insightSummary || a.reason || '', 180),
      reason: shortenMyCustomerStatusTextP528_(a.reason || a.insightSummary || '', 180),
      nextAction: shortenMyCustomerStatusTextP528_(a.nextAction || '', 180),
      lastContactDays: a.lastContactDays,
      longNoContact: !!a.longNoContact,
      sentNoFollow: !!a.sentNoFollow,
      potentialScore: a.potentialScore || 0,
      tags: (a.tags || []).slice(0, 12),
      latestMemoText: shortenMyCustomerStatusTextP528_(a.latestMemoText || '', 220),
      latestEventSource: a.latestEventSource || '',
      latestAnyEventText: shortenMyCustomerStatusTextP528_(a.latestAnyEventText || '', 220),
      latestAnyEventSource: a.latestAnyEventSource || '',
      latestDecisionEventText: shortenMyCustomerStatusTextP528_(a.latestDecisionEventText || '', 220),
      latestDecisionEventDate: a.latestDecisionEventDate || '',
      latestDecisionEventSource: a.latestDecisionEventSource || '',
      matchedKeywords: reduceMyCustomerMatchedKeywordsForClientP531_(a.matchedKeywords || {}),
      sourceLabels: (a.sourceLabels || []).slice(0, 4),
      missingFields: (a.missingFields || []).slice(0, 8),
      dataMissing: !!a.dataMissing,
      statusProtected: !!a.statusProtected,
      stateChangeAllowed: !!a.stateChangeAllowed,
      canApplyRecommendation: !!a.canApplyRecommendation,
      recommendationVisible: !!a.recommendationVisible,
      terminalCandidate: !!a.terminalCandidate,
      contractCompleteType: a.contractCompleteType || '',
      contractCompletionInfo: a.contractCompletionInfo || {},
      thirdPartyContract: !!a.thirdPartyContract,
      analysisCautionLevel: a.analysisCautionLevel || '',
      conservativeReason: a.conservativeReason || '',
      statusRecommendationBlockedReason: a.statusRecommendationBlockedReason || ''
    }
  };
}

function reduceMyCustomerMatchedKeywordsForClientP531_(matched) {
  const out = {};
  ['thirdPartyContract','fail','contactBlocker','complete','unknownComplete','order','quote','long','active','data'].forEach(function(k) {
    if (matched && matched[k] && matched[k].length) out[k] = matched[k].slice(0, 5);
  });
  return out;
}

function safeStringifyMyCustomerStatusP529_(value, maxLen) {
  let s = '';
  try { s = JSON.stringify(value || {}); } catch (err) { s = String(value || ''); }
  maxLen = Number(maxLen || 20000) || 20000;
  return s.length > maxLen ? s.slice(0, maxLen - 1) + '…' : s;
}
