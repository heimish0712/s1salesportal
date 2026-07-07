/***************************************
 * S1 Sales Portal - 23_FocusCustomerService.gs
 * STEP19 / P290: 중점고객관리 메뉴
 * - 마스터/검색인덱스 기준 고객 자동점수 산정
 * - A/B/C/보류/제외 추천등급 및 오늘 우선 처리 고객 목록 제공
 * - 권한_DB 기준 고객 범위는 기존 고객검색 권한 필터를 그대로 사용
 ***************************************/

const PORTAL_FOCUS_CACHE_PREFIX_P290 = 'PORTAL_FOCUS_CUSTOMERS_P290_';
const PORTAL_FOCUS_CACHE_TTL_SEC_P290 = 180;
const PORTAL_FOCUS_DEFAULT_PAGE_SIZE_P290 = 20;
const PORTAL_FOCUS_MAX_RETURN_ROWS_P290 = 100;

function getFocusCustomerDashboardP290(options) {
  options = options || {};
  const perm = getPortalCurrentPermission_();
  if (!perm || perm.active === false) throw new Error('중점고객관리 권한이 없습니다. 로그인 계정 권한을 확인하세요.');

  const keyword = normalizeFocusTextP290_(options.keyword || '');
  const rankFilter = String(options.rank || options.grade || 'ALL').trim() || 'ALL';
  const salesRepFilter = normalizeFocusFilterSentinelP291_(options.salesRep || 'ALL');
  const statusFilter = normalizeFocusFilterSentinelP291_(options.status || 'ALL');
  const priorityOnly = !!options.priorityOnly;
  const includeExcluded = !!options.includeExcluded;
  const page = Math.max(0, Number(options.page) || 0);
  const pageSize = Math.max(10, Math.min(PORTAL_FOCUS_MAX_RETURN_ROWS_P290, Number(options.pageSize) || PORTAL_FOCUS_DEFAULT_PAGE_SIZE_P290));
  const sortMode = String(options.sortMode || 'priority').trim();
  const metricFilter = String(options.metricFilter || 'ALL').trim() || 'ALL';
  const cacheKey = makeFocusCustomerCacheKeyP290_(perm, {
    keyword: keyword,
    rank: rankFilter,
    salesRep: salesRepFilter,
    status: statusFilter,
    priorityOnly: priorityOnly,
    includeExcluded: includeExcluded,
    page: page,
    pageSize: pageSize,
    sortMode: sortMode,
    metricFilter: metricFilter
  });

  const cached = readFocusCustomerCacheP290_(cacheKey);
  if (cached) {
    cached.fromCache = true;
    return cached;
  }

  const indexData = getCustomerSearchIndexData(perm);
  let rows = Array.isArray(indexData.rows) ? indexData.rows.slice() : [];
  rows = rows.map(function(row) { return buildFocusCustomerRowP290_(row); });

  const allStats = buildFocusCustomerStatsP290_(rows);
  const salesRepOptions = uniqueFocusOptionsP290_(rows, 'salesRep');
  const statusOptions = uniqueFocusOptionsP290_(rows, 'status');

  rows = rows.filter(function(row) {
    if (!includeExcluded && row.finalRank === '제외') return false;
    if (priorityOnly && !row.priorityTarget) return false;
    if (rankFilter && rankFilter !== 'ALL' && row.finalRank !== rankFilter) return false;
    if (!doesFocusCustomerMatchMetricP514_(row, metricFilter)) return false;
    if (salesRepFilter && salesRepFilter !== 'ALL' && normalizeFocusTextP290_(row.salesRep) !== salesRepFilter) return false;
    if (statusFilter && statusFilter !== 'ALL' && normalizeFocusTextP290_(row.status) !== statusFilter) return false;
    if (keyword && !doesFocusCustomerMatchKeywordP290_(row, keyword)) return false;
    return true;
  });

  rows = sortFocusCustomerRowsP290_(rows, sortMode);
  const total = rows.length;
  const start = page * pageSize;
  const pageRows = rows.slice(start, start + pageSize);
  const todayRows = sortFocusCustomerRowsP290_(rows.filter(function(r) { return r.priorityTarget; }), 'priority').slice(0, 12);

  const result = {
    ok: true,
    rows: pageRows,
    todayRows: todayRows,
    total: total,
    page: page,
    pageSize: pageSize,
    start: start,
    end: Math.min(start + pageRows.length, total),
    hasPrev: page > 0,
    hasNext: start + pageSize < total,
    stats: allStats,
    filterStats: buildFocusCustomerStatsP290_(rows),
    options: {
      salesReps: salesRepOptions,
      statuses: statusOptions,
      ranks: ['ALL', 'A급', 'B급', 'C급', '보류', '제외'],
      buckets: buildFocusBucketOptionsP514_()
    },
    keyword: keyword,
    rank: rankFilter,
    salesRep: salesRepFilter,
    status: statusFilter,
    priorityOnly: priorityOnly,
    includeExcluded: includeExcluded,
    sortMode: sortMode,
    metricFilter: metricFilter,
    source: '검색인덱스_DB+상태판정/업무분류',
    indexVersion: indexData.version || '',
    indexBuiltAt: indexData.builtAt || '',
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    fromCache: false
  };
  writeFocusCustomerCacheP290_(cacheKey, result);
  return result;
}

function clearFocusCustomerCacheP290() {
  PropertiesService.getScriptProperties().setProperty('PORTAL_FOCUS_CACHE_BUST_P290', String(Date.now()));
  return { ok: true, clearedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') };
}

function makeFocusCustomerCacheKeyP290_(perm, opts) {
  const props = PropertiesService.getScriptProperties();
  const bust = props.getProperty('PORTAL_FOCUS_CACHE_BUST_P290') || '0';
  const indexVersion = props.getProperty('CUSTOMER_SEARCH_INDEX_VERSION') || '';
  const indexDirty = props.getProperty('CUSTOMER_SEARCH_INDEX_DIRTY') || 'N';
  const indexDirtyAt = props.getProperty('CUSTOMER_SEARCH_INDEX_DIRTY_AT') || '';
  const masterVersion = props.getProperty('PORTAL_MASTER_DATA_VERSION') || '';
  const email = String(perm && perm.email || '').trim().toLowerCase();
  const scope = [perm && perm.level, perm && perm.defaultScope, perm && perm.canViewAllCustomers ? 'ALL' : 'OWN', perm && perm.salesRepName].join('|');
  const raw = JSON.stringify({
    email: email,
    scope: scope,
    opts: opts || {},
    indexVersion: indexVersion,
    indexDirty: indexDirty,
    indexDirtyAt: indexDirtyAt,
    masterVersion: masterVersion,
    bust: bust
  });
  const digest = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw)).slice(0, 30);
  return PORTAL_FOCUS_CACHE_PREFIX_P290 + digest;
}

function readFocusCustomerCacheP290_(key) {
  try {
    const raw = CacheService.getScriptCache().get(key);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function writeFocusCustomerCacheP290_(key, value) {
  try {
    const payload = JSON.stringify(value || {});
    if (payload.length > 90000) return;
    CacheService.getScriptCache().put(key, payload, PORTAL_FOCUS_CACHE_TTL_SEC_P290);
  } catch (err) {}
}

function buildFocusCustomerRowP290_(row) {
  row = row || {};
  const score = scoreFocusCustomerP290_(row);
  const finalRank = chooseFocusFinalRankP290_(row.customerRank, score.autoRank);
  const nextAction = suggestFocusNextActionP290_(row, score, finalRank);
  const priorityTarget = isFocusPriorityTargetP290_(row, score, finalRank);
  const text = buildFocusCombinedTextP290_(row);
  const memoInferredStatus = String(row.memoInferredStatus || '').trim();
  const statusMatch = normalizeFocusStatusMatchLabelP514_(row.statusMatch, row.status, memoInferredStatus);
  const tmProgressStatus = String(row.tmProgressStatus || '').trim();
  const tmContactContent = String(row.tmContactContent || '').trim();
  const focusBucket = score.focusBucket || 'TODAY';
  const bucketMeta = getFocusBucketMetaP514_(focusBucket);
  return {
    rowNo: Number(row.rowNo) || 0,
    customerNo: String(row.customerNo || '').trim(),
    company: String(row.company || '').trim(),
    salesRep: String(row.salesRep || '').trim(),
    status: String(row.status || '').trim(),
    memoInferredStatus: memoInferredStatus,
    effectiveStatus: score.effectiveStatus || memoInferredStatus || String(row.status || '').trim(),
    statusMatch: statusMatch,
    statusMismatch: !!score.statusMismatch,
    customerRank: String(row.customerRank || '').trim(),
    finalRank: finalRank,
    autoRank: score.autoRank,
    focusBucket: focusBucket,
    focusBucketLabel: bucketMeta.label,
    focusBucketTone: bucketMeta.tone,
    latestSignal: score.latestSignal || '',
    score: score.score,
    priorityScore: score.priorityScore,
    priorityTarget: priorityTarget,
    priorityLabel: priorityTarget ? '오늘 처리' : '',
    area: String(row.area || '').trim(),
    areaNumber: parseFocusNumberP290_(row.area),
    grade: String(row.grade || '').trim(),
    buildingType: String(row.buildingType || '').trim(),
    contact: String(row.contact || '').trim(),
    phone: String(row.phone || '').trim(),
    directPhone: String(row.directPhone || '').trim(),
    email: String(row.email || '').trim(),
    region: String(row.region || '').trim(),
    vendor: String(row.vendor || '').trim(),
    finalQuote: String(row.finalQuote || '').trim(),
    memo: String(row.memo || '').trim(),
    memoPreview: shortenFocusTextP290_(row.memo || '', 120),
    tmProgressStatus: tmProgressStatus,
    tmContactContent: tmContactContent,
    tmPreview: shortenFocusTextP290_(tmContactContent || tmProgressStatus, 80),
    longNoContactTransferred: String(row.longNoContactTransferred || '').trim(),
    areaCheckNeeded: normalizeFocusBooleanP514_(row.areaCheckNeeded),
    addressCheckNeeded: normalizeFocusBooleanP514_(row.addressCheckNeeded),
    addressNormalizeStatus: String(row.addressNormalizeStatus || '').trim(),
    addressNormalizeNote: String(row.addressNormalizeNote || '').trim(),
    dataIssueLevel: score.dataIssueLevel || '',
    sendCount: String(row.sendCount || '').trim(),
    sendStatus: String(row.sendStatus || '').trim(),
    address: String(row.fullAddress || row.address || '').trim(),
    lastSent: String(row.lastSent || '').trim(),
    sentAt: String(row.sentAt || '').trim(),
    nextAction: nextAction,
    reasons: score.reasons.slice(0, 8),
    deductions: score.deductions.slice(0, 5),
    exclusionReason: score.exclusionReason || '',
    searchText: text,
    sortKey: [focusBucket, finalRank, score.priorityScore, score.score, row.customerNo || ''].join('|')
  };
}

function scoreFocusCustomerP290_(row) {
  row = row || {};
  const reasons = [];
  const deductions = [];
  const text = buildFocusCombinedTextP290_(row);
  const statusRaw = String(row.status || '').trim();
  const memoStatusRaw = String(row.memoInferredStatus || '').trim();
  const status = normalizeFocusTextP290_(statusRaw);
  const memoStatus = normalizeFocusTextP290_(memoStatusRaw);
  const buildingType = normalizeFocusTextP290_(row.buildingType);
  const area = parseFocusNumberP290_(row.area);
  let score = 0;
  let priorityScore = 0;
  let exclusionReason = '';

  const statusMismatch = isFocusStatusMismatchP514_(statusRaw, memoStatusRaw, row.statusMatch);
  const tmSignal = getFocusTmSignalP514_(row);
  const dataIssue = getFocusDataIssueLevelP514_(row, text);
  const closed = isFocusClosedSignalP514_(row, text);
  const longTerm = isFocusLongTermSignalP514_(row, text);
  const quoteNoResponse = isFocusQuoteNoResponseP340_(row);
  const explicitExcludedType = (buildingType.indexOf('공동주택') >= 0 || buildingType.indexOf('학교') >= 0 || text.indexOf('초중고') >= 0 || text.indexOf('초등학교') >= 0 || text.indexOf('중학교') >= 0 || text.indexOf('고등학교') >= 0);

  let focusBucket = 'TODAY';
  if (statusMismatch) {
    focusBucket = 'STATUS_REVIEW';
    reasons.push('현재상태와 메모상 추측 상태값 불일치');
    priorityScore += 70;
  } else if (closed || explicitExcludedType) {
    focusBucket = 'CLOSED';
    exclusionReason = explicitExcludedType ? '제외대상 건물유형' : '종결/제외 신호';
    reasons.push(exclusionReason);
  } else if (dataIssue) {
    focusBucket = 'DATA_FIX';
    reasons.push(dataIssue === 'high' ? '주소/연면적 등 핵심 데이터 확인 필요' : '데이터 보강 필요');
    priorityScore += 50;
  } else if (longTerm) {
    focusBucket = 'LONG_TERM';
    reasons.push('장기추진/예약 관리 신호');
    priorityScore += 20;
  } else if (tmSignal === 'success') {
    focusBucket = 'TM_SUCCESS_FOLLOWUP';
    reasons.push('TM 성공/검토중 후속관리 필요');
    priorityScore += 55;
  } else if (quoteNoResponse) {
    focusBucket = 'QUOTE_NO_RESPONSE';
    reasons.push('견적 발송 후 후속 확인 필요');
    priorityScore += 60;
  } else {
    focusBucket = 'TODAY';
    reasons.push('영업 우선 관리 대상');
    priorityScore += 40;
  }

  if (focusBucket === 'CLOSED') {
    return {
      score: 0,
      priorityScore: 0,
      autoRank: '제외',
      focusBucket: focusBucket,
      effectiveStatus: memoStatusRaw || statusRaw,
      statusMismatch: statusMismatch,
      latestSignal: summarizeFocusLatestSignalP514_(row, text),
      dataIssueLevel: dataIssue,
      reasons: reasons,
      deductions: [],
      exclusionReason: exclusionReason
    };
  }

  // 1) 법 적용 및 연면적 우선도
  if (area >= 60000) { score += 30; priorityScore += 30; reasons.push('연면적 60,000㎡ 이상/특급 우선도'); }
  else if (area >= 30000) { score += 28; priorityScore += 28; reasons.push('연면적 30,000㎡ 이상/고급 우선도'); }
  else if (area >= 15000) { score += 25; priorityScore += 25; reasons.push('연면적 15,000㎡ 이상/중급 우선도'); }
  else if (area >= 10000) { score += 22; priorityScore += 22; reasons.push('연면적 10,000㎡ 이상 올해 대상'); }
  else if (area >= 5000) { score += 10; priorityScore += 10; reasons.push('연면적 5,000㎡ 이상 장기 관리군'); }
  else if (area > 0) { score += 2; deductions.push('연면적 5,000㎡ 미만'); }
  else { score += 5; deductions.push('연면적 미확인'); }

  // 2) 고객정보 완성도
  if (row.phone || row.directPhone) { score += 4; reasons.push('전화번호 확보'); }
  else { score -= 4; deductions.push('전화번호 미확인'); }
  if (row.contact) { score += 4; reasons.push('담당자 확인'); }
  else { score -= 5; deductions.push('담당자 미확인'); }
  if (row.email) { score += 4; reasons.push('이메일 확보'); }
  else { score -= 3; deductions.push('이메일 미확인'); }
  if (row.fullAddress || row.address) { score += 2; }
  if (area > 0) { score += 1; }

  // 3) 진행 신호
  if (hasAnyFocusKeywordP290_(text, ['견적요청', '견적받', '견적서요청', '견적서발송', '견적발송', '견적제출', '견적재발송'])) { score += 8; priorityScore += 8; reasons.push('견적 요청/발송 신호'); }
  if (hasAnyFocusKeywordP290_(text, ['자료요청', '자료발송', '안내문', '법령자료', '제안서', '단가표'])) { score += 4; priorityScore += 4; reasons.push('자료/안내문 요청 또는 발송'); }
  if (hasAnyFocusKeywordP290_(text, ['재발송', '다시보내', '다시 보내', '메일확인', '확인콜'])) { score += 5; priorityScore += 5; reasons.push('재발송/메일 확인 신호'); }
  if (hasAnyFocusKeywordP290_(text, ['긍정', '호의적', '문의', '관심', '검토하겠다', '검토중', '업체선정전', '아직업체선정전'])) { score += 5; priorityScore += 6; reasons.push('검토중/긍정 반응'); }
  if (hasAnyFocusKeywordP290_(text, ['예산', '금액', '단가', '할인', '비싸', '가격', '견적가', '네고'])) { score += 5; priorityScore += 5; reasons.push('예산/가격/네고 언급'); }
  if (hasAnyFocusKeywordP290_(text, ['재연락', '연락준', '결정시점', '이번주', '다음주', '월말', '월요일', '목요일'])) { score += 5; priorityScore += 5; reasons.push('후속시점 단서'); }
  if (hasAnyFocusKeywordP290_(text, ['계약', '발주', '품의', '나라장터', '용역신청서', '사업자등록증'])) { score += 8; priorityScore += 10; reasons.push('계약/발주/서류 절차 신호'); }
  if (hasAnyFocusKeywordP290_(text, ['수행사', '방문', '미팅', '수정견적', '비교견적'])) { score += 5; priorityScore += 6; reasons.push('수행사/미팅/비교견적 구체화'); }

  // 4) 상태/신호 기반 보정
  if (status.indexOf('고객설득중') >= 0 || memoStatus.indexOf('고객설득중') >= 0) { score += 8; priorityScore += 8; reasons.push('고객 설득 중'); }
  if (status.indexOf('견적제출완료') >= 0 || memoStatus.indexOf('견적제출완료') >= 0) { score += 6; priorityScore += 8; reasons.push('견적 제출 완료'); }
  if (status.indexOf('발주완료') >= 0 || memoStatus.indexOf('발주완료') >= 0) { score += 10; priorityScore += 12; reasons.push('발주/계약 서류 후속 필요'); }
  if (tmSignal === 'fail') { score -= 12; deductions.push('TM 실패/거절 신호'); }
  if (hasAnyFocusKeywordP290_(text, ['타사계약', '타업체', '다른업체', '업체선정완료', '계약완료라고하심'])) { score -= 12; deductions.push('타사계약/업체선정완료 언급'); }
  if (hasAnyFocusKeywordP290_(text, ['자체선임'])) {
    if (hasAnyFocusKeywordP290_(text, ['점검위탁', '성능점검', '유지점검'])) { score -= 4; deductions.push('자체선임이나 점검위탁 가능성'); }
    else { score -= 10; deductions.push('자체선임 예정'); }
  }
  if (hasAnyFocusKeywordP290_(text, ['장기무응답', '부재중', '연락안됨', '연락 안됨', '전화넘김'])) { score -= 8; deductions.push('부재/무응답 신호'); }

  score = Math.max(0, Math.min(100, Math.round(score)));
  let autoRank = score >= 75 ? 'A급' : score >= 55 ? 'B급' : score >= 35 ? 'C급' : '보류';
  if (focusBucket === 'DATA_FIX' || focusBucket === 'STATUS_REVIEW') autoRank = score >= 55 ? 'C급' : '보류';
  if (focusBucket === 'LONG_TERM') autoRank = score >= 65 ? 'C급' : '보류';

  return {
    score: score,
    priorityScore: priorityScore + score,
    autoRank: autoRank,
    focusBucket: focusBucket,
    effectiveStatus: memoStatusRaw || statusRaw,
    statusMismatch: statusMismatch,
    latestSignal: summarizeFocusLatestSignalP514_(row, text),
    dataIssueLevel: dataIssue,
    reasons: reasons,
    deductions: deductions,
    exclusionReason: ''
  };
}

function chooseFocusFinalRankP290_(manualRank, autoRank) {
  if (autoRank === '제외') return '제외';
  const rank = String(manualRank || '').trim();
  if (['A급', 'B급', 'C급', '보류', '제외'].indexOf(rank) >= 0) return rank;
  return autoRank || '보류';
}

function suggestFocusNextActionP290_(row, score, finalRank) {
  row = row || {};
  score = score || {};
  const bucket = score.focusBucket || 'TODAY';
  const text = buildFocusCombinedTextP290_(row);
  if (bucket === 'CLOSED') return '종결/제외 확인';
  if (bucket === 'STATUS_REVIEW') return '현재상태와 메모추측 상태값 정리';
  if (bucket === 'DATA_FIX') return '주소·연면적·담당자 데이터 보강';
  if (bucket === 'LONG_TERM') return '장기추진 예약일/재컨택월 지정';
  if (bucket === 'TM_SUCCESS_FOLLOWUP') return 'TM 성공콜 인계 후 영업 후속콜';
  if (bucket === 'QUOTE_NO_RESPONSE') return '견적 발송 후 확인콜';
  if (hasAnyFocusKeywordP290_(text, ['발주', '계약', '품의', '나라장터', '용역신청서'])) return '계약/발주 서류 진행 확인';
  if (hasAnyFocusKeywordP290_(text, ['재발송', '다시보내', '메일확인'])) return '자료 재발송 후 확인콜';
  if (hasAnyFocusKeywordP290_(text, ['견적', '비교견적', '수정견적', '가격', '네고'])) return '견적 조건 확인 및 후속콜';
  if (!row.contact || !row.email) return '담당자/이메일 보강';
  if (finalRank === 'A급') return '1~3일 내 후속조치';
  if (finalRank === 'B급') return '7~14일 내 재컨택';
  if (finalRank === 'C급') return '월 1회 장기관리';
  return '정보 보강 후 분류 확정';
}

function isFocusPriorityTargetP290_(row, score, finalRank) {
  score = score || {};
  const bucket = score.focusBucket || 'TODAY';
  return ['TODAY', 'QUOTE_NO_RESPONSE', 'TM_SUCCESS_FOLLOWUP', 'STATUS_REVIEW', 'DATA_FIX'].indexOf(bucket) >= 0;
}


function isFocusQuoteNoResponseP340_(row) {
  row = row || {};
  const text = buildFocusCombinedTextP290_(row);
  const sent = hasAnyFocusKeywordP290_(text, ['견적발송', '견적제출', '견적서발송', '견적재발송', '재견적']) || !!row.lastSent || !!row.sentAt || Number(row.sendCount || 0) > 0;
  const response = hasAnyFocusKeywordP290_(text, ['확인', '검토', '회신', '재연락', '전화오', '통화함', '결정', '계약', '발주', '수주실패', '타사계약']);
  const closed = isFocusClosedSignalP514_(row, text);
  return !!sent && !response && !closed;
}

function buildFocusCustomerStatsP290_(rows) {
  const stats = {
    total: 0,
    a: 0,
    b: 0,
    c: 0,
    pending: 0,
    excluded: 0,
    today: 0,
    delayed: 0,
    quoteNoResponse: 0,
    statusReview: 0,
    dataFix: 0,
    tmSuccessFollowup: 0,
    longTerm: 0,
    closed: 0,
    byRank: {},
    byBucket: {}
  };
  buildFocusBucketOptionsP514_().forEach(function(b) { stats.byBucket[b.value] = 0; });
  (rows || []).forEach(function(r) {
    stats.total++;
    const rank = r.finalRank || '보류';
    stats.byRank[rank] = (stats.byRank[rank] || 0) + 1;
    if (rank === 'A급') stats.a++;
    else if (rank === 'B급') stats.b++;
    else if (rank === 'C급') stats.c++;
    else if (rank === '제외') stats.excluded++;
    else stats.pending++;
    if (r.priorityTarget) stats.today++;
    if (isFocusQuoteNoResponseP340_(r) || r.focusBucket === 'QUOTE_NO_RESPONSE') stats.quoteNoResponse++;
    const bucket = r.focusBucket || 'TODAY';
    stats.byBucket[bucket] = (stats.byBucket[bucket] || 0) + 1;
    if (bucket === 'STATUS_REVIEW') stats.statusReview++;
    if (bucket === 'DATA_FIX') stats.dataFix++;
    if (bucket === 'TM_SUCCESS_FOLLOWUP') stats.tmSuccessFollowup++;
    if (bucket === 'LONG_TERM') stats.longTerm++;
    if (bucket === 'CLOSED') stats.closed++;
  });
  return stats;
}

function sortFocusCustomerRowsP290_(rows, sortMode) {
  rows = Array.isArray(rows) ? rows.slice() : [];
  const mode = String(sortMode || 'priority').trim();
  return rows.sort(function(a, b) {
    if (mode === 'scoreDesc') return (Number(b.score) || 0) - (Number(a.score) || 0) || compareFocusCustomerNoP290_(b, a);
    if (mode === 'areaDesc') return (Number(b.areaNumber) || 0) - (Number(a.areaNumber) || 0) || (Number(b.score) || 0) - (Number(a.score) || 0);
    if (mode === 'customerNoDesc') return compareFocusCustomerNoP290_(b, a);
    if (mode === 'companyAsc') return String(a.company || '').localeCompare(String(b.company || ''), 'ko', { numeric: true, sensitivity: 'base' });
    return getFocusBucketSortWeightP514_(b.focusBucket) - getFocusBucketSortWeightP514_(a.focusBucket) ||
      (Number(b.priorityScore) || 0) - (Number(a.priorityScore) || 0) ||
      (Number(b.score) || 0) - (Number(a.score) || 0) ||
      compareFocusCustomerNoP290_(b, a);
  });
}

function compareFocusCustomerNoP290_(a, b) {
  return (Number(String(a && a.customerNo || '').replace(/\D/g, '')) || 0) - (Number(String(b && b.customerNo || '').replace(/\D/g, '')) || 0);
}

function uniqueFocusOptionsP290_(rows, key) {
  const seen = {};
  const out = [];
  (rows || []).forEach(function(r) {
    const v = String(r && r[key] || '').trim();
    if (!v || seen[v]) return;
    seen[v] = true;
    out.push(v);
  });
  return out.sort(function(a, b) { return String(a).localeCompare(String(b), 'ko', { numeric: true, sensitivity: 'base' }); });
}

function doesFocusCustomerMatchKeywordP290_(row, keyword) {
  keyword = normalizeFocusTextP290_(keyword);
  if (!keyword) return true;
  return normalizeFocusTextP290_(row.searchText || buildFocusCombinedTextP290_(row)).indexOf(keyword) >= 0;
}

function buildFocusCombinedTextP290_(row) {
  row = row || {};
  return normalizeFocusTextP290_([
    row.customerNo, row.company, row.salesRep, row.status, row.memoInferredStatus, row.statusMatch, row.customerRank, row.finalRank, row.focusBucketLabel,
    row.contact, row.phone, row.directPhone, row.email, row.region, row.vendor, row.grade,
    row.buildingType, row.area, row.finalQuote, row.memo, row.tmProgressStatus, row.tmContactContent,
    row.longNoContactTransferred, row.areaCheckNeeded, row.addressCheckNeeded, row.addressNormalizeStatus, row.addressNormalizeNote,
    row.sendStatus, row.sendCount, row.address, row.fullAddress, row.lastSent, row.sentAt
  ].join(' '));
}

function hasAnyFocusKeywordP290_(text, keywords) {
  text = normalizeFocusTextP290_(text);
  return (keywords || []).some(function(k) { return text.indexOf(normalizeFocusTextP290_(k)) >= 0; });
}


function doesFocusCustomerMatchMetricP514_(row, metricFilter) {
  const metric = String(metricFilter || 'ALL').trim() || 'ALL';
  if (metric === 'ALL') return true;
  if (metric === 'quoteNoResponse') return row.focusBucket === 'QUOTE_NO_RESPONSE' || isFocusQuoteNoResponseP340_(row);
  if (metric.indexOf('bucket:') === 0) return String(row.focusBucket || '') === metric.slice(7);
  if (metric === 'statusReview') return row.focusBucket === 'STATUS_REVIEW';
  if (metric === 'dataFix') return row.focusBucket === 'DATA_FIX';
  if (metric === 'tmSuccessFollowup') return row.focusBucket === 'TM_SUCCESS_FOLLOWUP';
  if (metric === 'longTerm') return row.focusBucket === 'LONG_TERM';
  if (metric === 'closed') return row.focusBucket === 'CLOSED';
  return true;
}

function buildFocusBucketOptionsP514_() {
  return [
    { value: 'TODAY', label: '오늘 처리', tone: 'today', order: 700 },
    { value: 'QUOTE_NO_RESPONSE', label: '견적 후 미확인', tone: 'quote', order: 650 },
    { value: 'TM_SUCCESS_FOLLOWUP', label: 'TM 성공 후속', tone: 'tm', order: 620 },
    { value: 'STATUS_REVIEW', label: '상태값 정리 필요', tone: 'status', order: 600 },
    { value: 'DATA_FIX', label: '데이터 보강 필요', tone: 'data', order: 550 },
    { value: 'LONG_TERM', label: '장기추진 예약', tone: 'long', order: 300 },
    { value: 'CLOSED', label: '종결/제외', tone: 'closed', order: 0 }
  ];
}

function getFocusBucketMetaP514_(bucket) {
  const list = buildFocusBucketOptionsP514_();
  for (let i = 0; i < list.length; i++) if (list[i].value === bucket) return list[i];
  return list[0];
}

function getFocusBucketSortWeightP514_(bucket) {
  return Number(getFocusBucketMetaP514_(bucket).order) || 0;
}

function normalizeFocusStatusCanonicalP514_(value) {
  const n = normalizeFocusTextP290_(value);
  if (!n) return '';
  if (n.indexOf('상태지정필요') >= 0) return '상태지정필요';
  if (n.indexOf('계약완료') >= 0) return '계약완료';
  if (n.indexOf('발주완료') >= 0 || n.indexOf('수주성공') >= 0 || n.indexOf('계약서취합완료') >= 0) return '발주완료';
  if (n.indexOf('수주실패') >= 0 || n.indexOf('영업종료') >= 0 || n.indexOf('거절') >= 0) return '수주실패';
  if (n.indexOf('장기미접촉') >= 0) return '장기미접촉';
  if (n.indexOf('장기추진') >= 0 || n.indexOf('장기') >= 0) return '장기추진건';
  if (n.indexOf('데이터확인필요') >= 0 || n.indexOf('데이터') >= 0) return '데이터확인필요';
  if (n.indexOf('견적제출완료') >= 0 || n.indexOf('견적') >= 0) return '견적제출완료';
  if (n.indexOf('고객설득중') >= 0 || n.indexOf('영업중') >= 0 || n.indexOf('검토중') >= 0) return '고객설득중';
  return n;
}

function normalizeFocusStatusMatchLabelP514_(statusMatch, status, memoStatus) {
  const raw = String(statusMatch || '').trim().toUpperCase();
  if (raw === 'O' || raw === 'OK' || raw === 'TRUE' || raw === 'Y') return 'O';
  if (raw === 'X' || raw === 'FALSE' || raw === 'N') return 'X';
  if (!memoStatus) return '';
  return normalizeFocusStatusCanonicalP514_(status) === normalizeFocusStatusCanonicalP514_(memoStatus) ? 'O' : 'X';
}

function isFocusStatusMismatchP514_(status, memoStatus, statusMatch) {
  const label = normalizeFocusStatusMatchLabelP514_(statusMatch, status, memoStatus);
  if (label === 'X') return true;
  if (label === 'O') return false;
  const a = normalizeFocusStatusCanonicalP514_(status);
  const b = normalizeFocusStatusCanonicalP514_(memoStatus);
  return !!(a && b && a !== b);
}

function normalizeFocusBooleanP514_(value) {
  const raw = String(value == null ? '' : value).trim();
  const n = normalizeFocusTextP290_(raw);
  if (!n) return false;
  return ['true', '1', 'y', 'yes', 'o', '필요', '확인필요'].indexOf(n) >= 0;
}

function getFocusTmSignalP514_(row) {
  row = row || {};
  const text = normalizeFocusTextP290_([row.tmProgressStatus, row.tmContactContent].join(' '));
  if (!text) return '';
  if (hasAnyFocusKeywordP290_(text, ['성공', '검토중', '업체선정전', '견적서', '재발송', '요청', '관심'])) return 'success';
  if (hasAnyFocusKeywordP290_(text, ['실패', '거절', '타사계약', '계약완료', '하지마', '필요없'])) return 'fail';
  return '';
}

function getFocusDataIssueLevelP514_(row, text) {
  row = row || {};
  text = normalizeFocusTextP290_(text || buildFocusCombinedTextP290_(row));
  const memoStatus = normalizeFocusStatusCanonicalP514_(row.memoInferredStatus);
  const areaNeeded = normalizeFocusBooleanP514_(row.areaCheckNeeded) || !parseFocusNumberP290_(row.area);
  const addressNeeded = normalizeFocusBooleanP514_(row.addressCheckNeeded) || hasAnyFocusKeywordP290_(row.addressNormalizeStatus || '', ['확인필요', '공란']) || hasAnyFocusKeywordP290_(row.addressNormalizeNote || '', ['확인필요', '공란', '주소 형식 불충분', '메모/복수주소']);
  const contactNeeded = !(row.contact || row.directPhone || row.phone || row.email);
  if (memoStatus === '데이터확인필요') return 'high';
  if (areaNeeded && addressNeeded) return 'high';
  if (areaNeeded || addressNeeded || contactNeeded) return 'medium';
  if (hasAnyFocusKeywordP290_(text, ['주소확인필요', '연면적확인필요', '데이터확인필요'])) return 'medium';
  return '';
}

function isFocusClosedSignalP514_(row, text) {
  row = row || {};
  text = normalizeFocusTextP290_(text || buildFocusCombinedTextP290_(row));
  const status = normalizeFocusStatusCanonicalP514_(row.status);
  const memoStatus = normalizeFocusStatusCanonicalP514_(row.memoInferredStatus);
  if (status === '수주실패' || status === '계약완료' || memoStatus === '수주실패' || memoStatus === '계약완료') return true;
  return hasAnyFocusKeywordP290_(text, ['타사계약완료', '업체선정완료', '계약완료라고하심', '대상아님', '폐업', '중복삭제', '중복으로삭제', '진행하지마세요']);
}

function isFocusLongTermSignalP514_(row, text) {
  row = row || {};
  text = normalizeFocusTextP290_(text || buildFocusCombinedTextP290_(row));
  const status = normalizeFocusStatusCanonicalP514_(row.status);
  const memoStatus = normalizeFocusStatusCanonicalP514_(row.memoInferredStatus);
  if (status === '장기추진건' || status === '장기미접촉' || memoStatus === '장기추진건' || memoStatus === '장기미접촉') return true;
  return hasAnyFocusKeywordP290_(text, ['내년', '11월', '12월', '하반기', '예산', '장기', '나중에', '추후', '연말', '자체선임후', '유지성능만']);
}

function summarizeFocusLatestSignalP514_(row, text) {
  row = row || {};
  if (row.tmProgressStatus) return 'TM ' + String(row.tmProgressStatus).trim();
  if (row.memoInferredStatus) return '메모추측 ' + String(row.memoInferredStatus).trim();
  if (row.sendStatus || row.sentAt || row.lastSent) return ['발송', row.sendStatus, row.sentAt || row.lastSent].filter(Boolean).join(' ');
  if (hasAnyFocusKeywordP290_(text, ['견적재발송', '견적서발송', '견적제출'])) return '견적 발송/제출 신호';
  if (hasAnyFocusKeywordP290_(text, ['검토중', '업체선정전'])) return '검토중 신호';
  return '';
}

function normalizeFocusFilterSentinelP291_(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return 'ALL';
  if (raw === 'ALL' || raw.toUpperCase() === 'ALL') return 'ALL';
  return normalizeFocusTextP290_(raw);
}

function normalizeFocusTextP290_(value) {
  return String(value == null ? '' : value).replace(/[\s\n\r\t]+/g, '').trim().toLowerCase();
}

function parseFocusNumberP290_(value) {
  const raw = String(value == null ? '' : value).replace(/[₩￦원,㎡\s]/g, '').trim();
  const n = Number(raw);
  return isNaN(n) ? 0 : n;
}

function shortenFocusTextP290_(value, maxLen) {
  const text = String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
  maxLen = Number(maxLen) || 120;
  return text.length > maxLen ? text.slice(0, maxLen - 1) + '…' : text;
}
