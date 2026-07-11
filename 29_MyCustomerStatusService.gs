/***************************************
 * S1 Sales Portal - 29_MyCustomerStatusService.gs
 * P528: 나의 고객 현황 - 규칙 기반 1차 분석
 * - 내 고객 기준: 마스터시트/검색인덱스의 영업담당자 = 현재 로그인 사용자의 영업담당자명
 * - GPT 연동 전 단계: 최신 메모 이벤트 + 키워드 기반 추천/불일치/후속관리 후보 추출
 ***************************************/

const MY_CUSTOMER_STATUS_ANALYZER_VERSION_P528 = 'P528_RULE_BASED_MEMO_ANALYZER_V1';
const MY_CUSTOMER_STATUS_BLOCKING_STATUSES_P528 = ['수주실패', '계약완료'];
const MY_CUSTOMER_STATUS_ACTIVE_STATUSES_P528 = ['견적제출완료', '장기 추진건', '고객 설득 중', '발주완료', '!!상태지정필요!!'];

function getMyCustomerStatusDashboardP528(options) {
  options = options || {};
  const started = new Date();
  const perm = getPortalCurrentPermission_();
  const aliases = buildMyCustomerStatusOwnerAliasesP528_(perm);
  const indexData = getCustomerSearchIndexData(perm);
  const sourceRows = Array.isArray(indexData.rows) ? indexData.rows : [];
  const ownRows = sourceRows.filter(function(row) { return isMyCustomerStatusOwnRowP528_(row, aliases); });
  const now = new Date();

  const analyzed = ownRows.map(function(row) {
    return buildMyCustomerStatusAnalyzedRowP528_(row, now);
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

  const activeCount = analyzed.filter(function(row) { return isMyCustomerStatusActiveP528_(row.status); }).length;
  const needStatusCount = analyzed.filter(function(row) { return isMyCustomerStatusNeedStatusP528_(row.status); }).length;

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
      dataMissing: dataMissingRows.length
    },
    rows: analyzed.slice(0, 1500),
    lists: {
      mismatch: mismatchRows.slice(0, 300),
      recent7: recentRows.slice(0, 300),
      noContact14: noContactRows.slice(0, 300),
      sentNoFollow: sentNoFollowRows.slice(0, 300),
      highPotential: highPotentialRows.slice(0, 300),
      dataMissing: dataMissingRows.slice(0, 300)
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
  if (expectedStatus && currentStatus !== expectedStatus) {
    throw new Error('현재 진행현황이 화면에 표시된 값과 다릅니다. 현재값: ' + (currentStatus || '(공란)'));
  }

  const res = saveCustomerPatchFastP473({
    rowNo: rowNo,
    customerNo: customerNo,
    values: { status: newStatus },
    expectedValues: { status: currentStatus },
    clientSaveSource: 'myCustomerStatus.analysisStatusApply.P528',
    source: 'myCustomerStatus.analysisStatusApply.P528',
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
    source: 'myCustomerStatus.analysisStatusApply.P528'
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

function buildMyCustomerStatusAnalyzedRowP528_(row, now) {
  row = row || {};
  const memo = String(row.memo || '');
  const status = String(row.status || '').trim();
  const sentAt = String(row.sentAt || row.lastSent || '');
  const events = extractMyCustomerMemoEventsP528_(memo);
  const latestEvent = getLatestMyCustomerMemoEventP528_(events);
  const latestDate = latestEvent && latestEvent.date ? latestEvent.date : parseLoosePortalDateP528_(sentAt);
  const lastContactDays = latestDate ? daysBetweenPortalDatesP528_(latestDate, now) : null;
  const analysis = classifyMyCustomerMemoP528_(row, events, latestEvent, lastContactDays, now);
  const latestText = latestEvent && latestEvent.text ? latestEvent.text : getLastMemoLineP528_(memo);
  const missingFields = getMyCustomerMissingFieldsP528_(row);
  analysis.missingFields = missingFields;
  analysis.dataMissing = missingFields.length > 0;
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
    memoSummary: shortenMyCustomerStatusTextP528_(latestText || memo, 220),
    lastContactDate: latestDate ? formatMyCustomerStatusDateP528_(latestDate) : '',
    analysis: analysis
  };
}

function classifyMyCustomerMemoP528_(row, events, latestEvent, lastContactDays, now) {
  row = row || {};
  const currentStatus = String(row.status || '').trim();
  const memo = String(row.memo || '');
  const latestText = String(latestEvent && latestEvent.text || getLastMemoLineP528_(memo) || '');
  const recentText = getRecentMemoTextP528_(events, memo, 5);
  const basisText = (latestText + '\n' + recentText).toLowerCase();
  const fullText = memo.toLowerCase();
  const signals = {
    complete: hasAnyKeywordP528_(basisText, ['계약완료', '계약 완료', '계약서 저장', '계약 완료 처리']),
    order: hasAnyKeywordP528_(basisText, ['용역신청서', '계약서', '사업자등록증', '사업자 등록증', '발주', '계약 진행', '진행하시기로', '서류 받', '서류 요청']),
    fail: hasAnyKeywordP528_(basisText, ['거절', '타사 계약', '타사계약', '기존 업체', '기존업체', '유지하기로', '진행 어려', '진행어려', '하지 말', '연락하지', '대상 아님', '대상아님', '폐업']),
    long: hasAnyKeywordP528_(basisText, ['내년', '예산', '하반기', '상반기', '추후', '나중', '장기', '보류', '재검토']),
    quote: hasAnyKeywordP528_(basisText, ['견적서 발송', '견적서발송', '견적 발송', '견적발송', '자료발송', '자료 발송', '단가표', '비교견적', '메일 발송', '발송완료']),
    active: hasAnyKeywordP528_(basisText, ['검토중', '검토 중', '비교중', '비교 중', '담당자 전달', '내부 검토', '확인 후', '재확인', '연락 예정', '가격 조율', '본사', '회신', '전화 예정']),
    data: hasAnyKeywordP528_(basisText + '\n' + fullText.slice(-800), ['중복', '전화번호 오류', '메일 오류', '주소 확인', '연면적 확인', '확인 필요', '정보 확인', '번호 오류'])
  };

  let recommended = '';
  let mismatchType = '';
  let confidence = 0;
  let reason = '';
  let nextAction = '';
  const sentDone = !!(String(row.sendStatus || '').indexOf('발송완료') >= 0 || row.lastSent || row.sentAt);

  if (signals.complete) {
    recommended = '계약완료'; confidence = 92; mismatchType = '계약완료 의심'; reason = '최근 메모에 계약완료/계약서 저장 신호가 있습니다.'; nextAction = '계약완료 여부를 확인하고 상태를 정리하세요.';
  } else if (signals.fail) {
    recommended = '수주실패'; confidence = 88; mismatchType = '수주실패 의심'; reason = '최근 메모에 거절/타사계약/진행불가 신호가 있습니다.'; nextAction = '수주실패 처리 또는 보류 사유를 확인하세요.';
  } else if (signals.order) {
    recommended = '발주완료'; confidence = 83; mismatchType = '단계 상향 의심'; reason = '최근 메모에 용역신청서/계약서/사업자등록증/발주 신호가 있습니다.'; nextAction = '계약서류 취합 여부를 확인하세요.';
  } else if (signals.long) {
    recommended = '장기 추진건'; confidence = 78; mismatchType = '장기추진 전환 의심'; reason = '최근 메모에 예산/추후/내년/보류 신호가 있습니다.'; nextAction = '다음 재접촉 시점을 잡아 두세요.';
  } else if (signals.quote || sentDone) {
    recommended = '견적제출완료'; confidence = signals.quote ? 76 : 68; mismatchType = '견적제출완료 의심'; reason = signals.quote ? '최근 메모에 견적/자료 발송 신호가 있습니다.' : '발송상태 또는 마지막발송 값이 있습니다.'; nextAction = '견적서 발송 후 후속 연락을 진행하세요.';
  } else if (signals.active) {
    recommended = '고객 설득 중'; confidence = 66; mismatchType = '진행중 상태 의심'; reason = '최근 메모에 검토/비교/재확인/담당자 전달 신호가 있습니다.'; nextAction = '고객 반응을 확인하고 다음 액션을 남기세요.';
  }

  const mismatch = !!(recommended && recommended !== currentStatus && confidence >= 65);
  const active = isMyCustomerStatusActiveP528_(currentStatus);
  const longNoContact = active && (lastContactDays == null || lastContactDays >= 14);
  let sentNoFollow = false;
  if (sentDone && active) {
    if (lastContactDays == null) sentNoFollow = true;
    else if (lastContactDays >= 3 && !signals.order && !signals.fail && !signals.complete) sentNoFollow = true;
  }
  const potentialScore = calculateMyCustomerPotentialScoreP528_(row, signals, lastContactDays, recommended);
  const tags = [];
  Object.keys(signals).forEach(function(k) { if (signals[k]) tags.push(k); });
  if (longNoContact) tags.push('장기미접촉');
  if (sentNoFollow) tags.push('자료발송후미후속');

  return {
    recommendedStatus: recommended,
    confidence: confidence,
    mismatch: mismatch,
    mismatchType: mismatch ? mismatchType : '',
    reason: reason,
    nextAction: nextAction,
    lastContactDays: lastContactDays,
    longNoContact: longNoContact,
    sentNoFollow: sentNoFollow,
    potentialScore: potentialScore,
    tags: tags,
    latestMemoDate: latestEvent && latestEvent.date ? formatMyCustomerStatusDateP528_(latestEvent.date) : '',
    latestMemoText: shortenMyCustomerStatusTextP528_(latestText, 180)
  };
}

function calculateMyCustomerPotentialScoreP528_(row, signals, lastContactDays, recommended) {
  let score = 0;
  if (lastContactDays != null && lastContactDays <= 7) score += 2;
  if (row.lastSent || row.sentAt || String(row.sendStatus || '').indexOf('발송완료') >= 0) score += 2;
  if (row.finalQuote) score += 1;
  if (row.email) score += 1;
  if (row.contact) score += 1;
  if (signals.order) score += 3;
  if (signals.quote || signals.active) score += 1;
  if (signals.fail) score -= 4;
  if (recommended === '발주완료') score += 2;
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
    if (ev.date && (!best.date || ev.date.getTime() >= best.date.getTime())) best = ev;
    else if (!best.date && ev.index > best.index) best = ev;
  });
  return best;
}

function getRecentMemoTextP528_(events, memo, limit) {
  events = Array.isArray(events) ? events : [];
  limit = Math.max(1, Number(limit || 5) || 5);
  if (events.length) return events.slice(Math.max(0, events.length - limit)).map(function(ev) { return ev.text; }).join('\n');
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
  return 'MY_STATUS_P528_' + String(customerNo || '') + '_' + String(rowNo || '') + '_' + String(Date.now());
}
