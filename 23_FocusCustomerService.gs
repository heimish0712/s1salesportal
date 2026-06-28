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
  const cacheKey = makeFocusCustomerCacheKeyP290_(perm, {
    keyword: keyword,
    rank: rankFilter,
    salesRep: salesRepFilter,
    status: statusFilter,
    priorityOnly: priorityOnly,
    includeExcluded: includeExcluded,
    page: page,
    pageSize: pageSize,
    sortMode: sortMode
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
      ranks: ['ALL', 'A급', 'B급', 'C급', '보류', '제외']
    },
    keyword: keyword,
    rank: rankFilter,
    salesRep: salesRepFilter,
    status: statusFilter,
    priorityOnly: priorityOnly,
    includeExcluded: includeExcluded,
    sortMode: sortMode,
    source: '검색인덱스_DB+자동점수',
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
  const email = String(perm && perm.email || '').trim().toLowerCase();
  const scope = [perm && perm.level, perm && perm.defaultScope, perm && perm.canViewAllCustomers ? 'ALL' : 'OWN', perm && perm.salesRepName].join('|');
  const raw = JSON.stringify({ email: email, scope: scope, opts: opts || {}, indexVersion: indexVersion, bust: bust });
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
  return {
    rowNo: Number(row.rowNo) || 0,
    customerNo: String(row.customerNo || '').trim(),
    company: String(row.company || '').trim(),
    salesRep: String(row.salesRep || '').trim(),
    status: String(row.status || '').trim(),
    customerRank: String(row.customerRank || '').trim(),
    finalRank: finalRank,
    autoRank: score.autoRank,
    score: score.score,
    priorityScore: score.priorityScore,
    priorityTarget: priorityTarget,
    priorityLabel: priorityTarget ? '오늘 우선' : '',
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
    address: String(row.fullAddress || row.address || '').trim(),
    lastSent: String(row.lastSent || '').trim(),
    sentAt: String(row.sentAt || '').trim(),
    nextAction: nextAction,
    reasons: score.reasons.slice(0, 8),
    deductions: score.deductions.slice(0, 5),
    exclusionReason: score.exclusionReason || '',
    searchText: text,
    sortKey: [finalRank, score.priorityScore, score.score, row.customerNo || ''].join('|')
  };
}

function scoreFocusCustomerP290_(row) {
  row = row || {};
  const reasons = [];
  const deductions = [];
  const text = buildFocusCombinedTextP290_(row);
  const status = normalizeFocusTextP290_(row.status);
  const buildingType = normalizeFocusTextP290_(row.buildingType);
  const area = parseFocusNumberP290_(row.area);
  let score = 0;
  let priorityScore = 0;
  let exclusionReason = '';

  const explicitExclude = [
    ['전화번호오류', '전화번호 오류'], ['번호오류', '전화번호 오류'], ['대상아님', '대상 아님'], ['대상아님', '대상 아님'],
    ['폐업', '폐업'], ['중복', '중복'], ['명시적거절', '명시적 거절'], ['거절', '거절'], ['수주실패', '수주실패'],
    ['계약완료', '이미 계약완료'], ['발주완료', '이미 발주완료']
  ];
  for (let i = 0; i < explicitExclude.length; i++) {
    if (text.indexOf(normalizeFocusTextP290_(explicitExclude[i][0])) >= 0 || status.indexOf(normalizeFocusTextP290_(explicitExclude[i][0])) >= 0) {
      exclusionReason = explicitExclude[i][1];
      break;
    }
  }
  if (!exclusionReason && (buildingType.indexOf('공동주택') >= 0 || buildingType.indexOf('학교') >= 0 || text.indexOf('초중고') >= 0 || text.indexOf('초등학교') >= 0 || text.indexOf('중학교') >= 0 || text.indexOf('고등학교') >= 0)) {
    exclusionReason = '제외대상 건물유형';
  }

  if (exclusionReason) {
    return { score: 0, priorityScore: 0, autoRank: '제외', reasons: [exclusionReason], deductions: [], exclusionReason: exclusionReason };
  }

  // 1) 법 적용 및 연면적 우선도: 30점
  if (area >= 60000) { score += 30; priorityScore += 30; reasons.push('연면적 60,000㎡ 이상/특급 우선도'); }
  else if (area >= 30000) { score += 28; priorityScore += 28; reasons.push('연면적 30,000㎡ 이상/고급 우선도'); }
  else if (area >= 15000) { score += 25; priorityScore += 25; reasons.push('연면적 15,000㎡ 이상/중급 우선도'); }
  else if (area >= 10000) { score += 22; priorityScore += 22; reasons.push('연면적 10,000㎡ 이상 올해 대상'); }
  else if (area >= 5000) { score += 10; priorityScore += 10; reasons.push('연면적 5,000㎡ 이상 내년 관리군'); }
  else if (area > 0) { score += 2; reasons.push('연면적 5,000㎡ 미만/우선도 낮음'); }
  else { score += 5; reasons.push('연면적 미확인'); }

  // 2) 고객정보 완성도: 15점
  if (row.phone || row.directPhone) { score += 4; reasons.push('전화번호 확보'); }
  if (row.contact) { score += 4; reasons.push('담당자 확인'); }
  if (row.email) { score += 4; reasons.push('이메일 확보'); }
  if (row.fullAddress || row.address) { score += 2; reasons.push('주소 확인'); }
  if (area > 0) { score += 1; reasons.push('연면적 수치 확인'); }

  // 3) 관심 신호: 20점
  if (hasAnyFocusKeywordP290_(text, ['견적요청', '견적받', '견적서요청', '견적서발송', '견적발송', '견적제출'])) { score += 8; priorityScore += 8; reasons.push('견적 요청/견적 발송 신호'); }
  if (hasAnyFocusKeywordP290_(text, ['자료요청', '자료발송', '안내문', '법령자료', '제안서', '단가표'])) { score += 4; priorityScore += 4; reasons.push('자료/안내문 요청 또는 발송'); }
  if (hasAnyFocusKeywordP290_(text, ['담당자전달', '전달하', '전달요청'])) { score += 3; reasons.push('담당자 전달 신호'); }
  if (hasAnyFocusKeywordP290_(text, ['메일확인', '확인했다', '재발송', '다시보내', '다시 보내'])) { score += 3; priorityScore += 2; reasons.push('메일 확인/재발송 신호'); }
  if (hasAnyFocusKeywordP290_(text, ['긍정', '호의적', '문의', '관심', '검토하겠다', '검토중'])) { score += 2; reasons.push('긍정 반응/추가 문의'); }

  // 4) 진행 구체성: 25점
  if (hasAnyFocusKeywordP290_(text, ['예산', '금액', '단가', '할인', '비싸', '가격', '견적가'])) { score += 5; priorityScore += 5; reasons.push('예산/금액/가격 언급'); }
  if (hasAnyFocusKeywordP290_(text, ['검토일정', '재연락', '연락준', '결정시점', '결정', '시점', '이번주', '다음주', '월말'])) { score += 5; priorityScore += 5; reasons.push('검토/결정/재연락 시점 언급'); }
  if (hasAnyFocusKeywordP290_(text, ['결재권자', '관리주체', '담당자확인', '팀장', '부장', '과장', '실장'])) { score += 4; reasons.push('담당자/관리주체 구체화'); }
  if (hasAnyFocusKeywordP290_(text, ['계약', '발주', '품의', '내부검토', '선정', '업체선정'])) { score += 6; priorityScore += 8; reasons.push('계약/발주/선정 절차 언급'); }
  if (hasAnyFocusKeywordP290_(text, ['수행사', '방문', '미팅', '수정견적', '비교견적'])) { score += 5; priorityScore += 6; reasons.push('수행사/미팅/수정견적 구체화'); }

  // 5) 후속관리 가능성: 10점
  if (hasAnyFocusKeywordP290_(text, ['다음액션', '재확인', '재컨택', '전화하기', '다시연락'])) { score += 4; priorityScore += 4; reasons.push('다음액션 존재'); }
  if (hasAnyFocusKeywordP290_(text, ['오늘', '내일', '이번주', '다음주', '월말', '까지', '오전', '오후'])) { score += 4; priorityScore += 4; reasons.push('후속기한/시점 단서'); }
  if (row.memo) { score += 2; reasons.push('최근 컨택/판단근거 메모 존재'); }

  // 감점
  if (hasAnyFocusKeywordP290_(text, ['타사계약', '타업체', '타사', '다른곳', '다른 곳'])) { score -= 10; deductions.push('타사계약/타업체 언급 -10'); }
  if (hasAnyFocusKeywordP290_(text, ['자체선임'])) {
    if (hasAnyFocusKeywordP290_(text, ['점검위탁', '성능점검', '유지점검'])) { score -= 5; deductions.push('자체선임이나 점검위탁 가능성 -5'); }
    else { score -= 10; deductions.push('자체선임 예정 -10'); }
  }
  if (!row.contact) { score -= 5; deductions.push('담당자 미확인 -5'); }
  if (hasAnyFocusKeywordP290_(text, ['장기무응답', '부재중', '연락안됨', '연락 안됨'])) { score -= 10; deductions.push('장기 무응답/부재 -10'); }

  score = Math.max(0, Math.min(100, Math.round(score)));
  const autoRank = score >= 75 ? 'A급' : score >= 55 ? 'B급' : score >= 35 ? 'C급' : '보류';
  return { score: score, priorityScore: priorityScore + score, autoRank: autoRank, reasons: reasons, deductions: deductions, exclusionReason: '' };
}

function chooseFocusFinalRankP290_(manualRank, autoRank) {
  const rank = String(manualRank || '').trim();
  if (['A급', 'B급', 'C급', '보류', '제외'].indexOf(rank) >= 0) return rank;
  return autoRank || '보류';
}

function suggestFocusNextActionP290_(row, score, finalRank) {
  const text = buildFocusCombinedTextP290_(row);
  if (finalRank === '제외') return '일반관리 제외';
  if (hasAnyFocusKeywordP290_(text, ['발주', '계약', '품의', '선정'])) return '계약/발주 일정 확인';
  if (hasAnyFocusKeywordP290_(text, ['재발송', '다시보내', '메일확인'])) return '자료 재발송 후 확인콜';
  if (hasAnyFocusKeywordP290_(text, ['견적', '비교견적', '수정견적'])) return '견적 조건 확인 및 후속콜';
  if (!row.contact || !row.email) return '담당자/이메일 보강';
  if (finalRank === 'A급') return '1~3일 내 후속조치';
  if (finalRank === 'B급') return '7~14일 내 재컨택';
  if (finalRank === 'C급') return '월 1회 장기관리';
  return '정보 보강 후 등급 확정';
}

function isFocusPriorityTargetP290_(row, score, finalRank) {
  const text = buildFocusCombinedTextP290_(row);
  const status = normalizeFocusTextP290_(row.status);
  const area = parseFocusNumberP290_(row.area);
  if (finalRank === 'A급') return true;
  if (area >= 10000 && hasAnyFocusKeywordP290_(text, ['견적발송', '견적제출', '견적서발송', '견적제출완료'])) return true;
  if (status.indexOf('고객설득중') >= 0) return true;
  if (finalRank === 'B급' && hasAnyFocusKeywordP290_(text, ['재연락', '재확인', '다음주', '이번주', '검토중'])) return true;
  if (finalRank === '보류' && (!row.contact || !row.phone && !row.directPhone || !row.email)) return true;
  return false;
}

function buildFocusCustomerStatsP290_(rows) {
  const stats = { total: 0, a: 0, b: 0, c: 0, pending: 0, excluded: 0, today: 0, delayed: 0, quoteNoResponse: 0, byRank: {} };
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
    if (hasAnyFocusKeywordP290_(buildFocusCombinedTextP290_(r), ['견적발송', '견적제출', '견적서발송']) && !hasAnyFocusKeywordP290_(buildFocusCombinedTextP290_(r), ['확인', '검토', '회신', '재연락'])) stats.quoteNoResponse++;
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
    return (Number(b.priorityScore) || 0) - (Number(a.priorityScore) || 0) || (Number(b.score) || 0) - (Number(a.score) || 0) || compareFocusCustomerNoP290_(b, a);
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
    row.customerNo, row.company, row.salesRep, row.status, row.customerRank, row.finalRank,
    row.contact, row.phone, row.directPhone, row.email, row.region, row.vendor, row.grade,
    row.buildingType, row.area, row.finalQuote, row.memo, row.address, row.fullAddress, row.lastSent, row.sentAt
  ].join(' '));
}

function hasAnyFocusKeywordP290_(text, keywords) {
  text = normalizeFocusTextP290_(text);
  return (keywords || []).some(function(k) { return text.indexOf(normalizeFocusTextP290_(k)) >= 0; });
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
