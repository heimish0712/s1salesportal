/***************************************
 * S1 Sales Portal - 30_MyCustomerAiAnalysisService.gs
 * P537: 나의 고객 현황 - GPT 애매건 분석 큐/DB
 * - 규칙 기반 분석 결과 중 애매건만 OpenAI API로 보조 분석합니다.
 * - GPT 결과는 바로 상태값을 바꾸지 않고 고객AI분석_DB에 저장합니다.
 * - 최종 상태변경 허용 여부는 서버 업무규칙(P536 전이표)로 다시 검증합니다.
 ***************************************/

const MY_CUSTOMER_AI_ANALYZER_VERSION_P537 = 'P544_AI_PARSE_STABLE_CANDIDATE_SLIM';
const MY_CUSTOMER_AI_PROMPT_VERSION_P537 = 'P544_STATUS_RULES_SLIM_V2';
const MY_CUSTOMER_AI_ANALYSIS_SHEET_P537 = '고객AI분석_DB';
const MY_CUSTOMER_AI_QUEUE_SHEET_P537 = '고객AI분석큐_DB';
const MY_CUSTOMER_AI_DEFAULT_QUEUE_LIMIT_P537 = 120;
const MY_CUSTOMER_AI_DEFAULT_PROCESS_LIMIT_P537 = 10;
const MY_CUSTOMER_AI_MAX_RETRY_P537 = 3;
const MY_CUSTOMER_AI_DEFAULT_MODEL_P537 = 'gpt-4.1-mini';
const MY_CUSTOMER_AI_API_URL_P537 = 'https://api.openai.com/v1/responses';

const MY_CUSTOMER_AI_ANALYSIS_HEADERS_P537 = [
  '분석일시', '분석ID', '고객번호', 'rowNo', '회사명', '영업담당자', '현재상태', '규칙추천상태',
  'GPT추천상태', '최종추천상태', 'decision', 'recommendationAllowed', 'serverAllowed', 'confidence',
  'falsePositiveRisk', 'insightType', 'summary', 'reason', 'nextAction', 'evidenceJson', 'blockedReason',
  'inputHash', 'promptVersion', 'model', 'tokensInput', 'tokensOutput', 'estimatedCostUsd',
  'rawRequestJson', 'rawResponseJson', '검토상태', '검토자', '검토일시', '검토메모'
];

const MY_CUSTOMER_AI_QUEUE_HEADERS_P537 = [
  '등록일시', '수정일시', '작업ID', '고객번호', 'rowNo', '회사명', '영업담당자', '현재상태',
  '상태', '우선순위', '시도횟수', 'inputHash', 'candidateJson', 'resultJson', '마지막오류',
  '적용일시', '모델', '분석ID'
];

function ensureMyCustomerAiAnalysisSheetP537_() {
  const ss = getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(MY_CUSTOMER_AI_ANALYSIS_SHEET_P537);
  if (!sheet) {
    sheet = ss.insertSheet(MY_CUSTOMER_AI_ANALYSIS_SHEET_P537);
    sheet.getRange(1, 1, 1, MY_CUSTOMER_AI_ANALYSIS_HEADERS_P537.length).setValues([MY_CUSTOMER_AI_ANALYSIS_HEADERS_P537]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, MY_CUSTOMER_AI_ANALYSIS_HEADERS_P537.length).setFontWeight('bold').setBackground('#eef6ff');
    sheet.autoResizeColumns(1, Math.min(MY_CUSTOMER_AI_ANALYSIS_HEADERS_P537.length, 14));
    return sheet;
  }
  ensureSheetHeaders_(sheet, MY_CUSTOMER_AI_ANALYSIS_HEADERS_P537);
  return sheet;
}

function ensureMyCustomerAiQueueSheetP537_() {
  const ss = getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(MY_CUSTOMER_AI_QUEUE_SHEET_P537);
  if (!sheet) {
    sheet = ss.insertSheet(MY_CUSTOMER_AI_QUEUE_SHEET_P537);
    sheet.getRange(1, 1, 1, MY_CUSTOMER_AI_QUEUE_HEADERS_P537.length).setValues([MY_CUSTOMER_AI_QUEUE_HEADERS_P537]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, MY_CUSTOMER_AI_QUEUE_HEADERS_P537.length).setFontWeight('bold').setBackground('#fff7ed');
    sheet.autoResizeColumns(1, Math.min(MY_CUSTOMER_AI_QUEUE_HEADERS_P537.length, 14));
    return sheet;
  }
  ensureSheetHeaders_(sheet, MY_CUSTOMER_AI_QUEUE_HEADERS_P537);
  return sheet;
}

function setupMyCustomerAiAnalysisSheetsP537() {
  const analysisSheet = ensureMyCustomerAiAnalysisSheetP537_();
  const queueSheet = ensureMyCustomerAiQueueSheetP537_();
  return {
    ok: true,
    analysisSheet: MY_CUSTOMER_AI_ANALYSIS_SHEET_P537,
    queueSheet: MY_CUSTOMER_AI_QUEUE_SHEET_P537,
    url: getWebAppDbSpreadsheet_().getUrl(),
    analysisRows: Math.max(0, analysisSheet.getLastRow() - 1),
    queueRows: Math.max(0, queueSheet.getLastRow() - 1)
  };
}

function getMyCustomerAiAnalysisHealthP537(options) {
  options = options || {};
  try {
    const analysisSheet = ensureMyCustomerAiAnalysisSheetP537_();
    const queueSheet = ensureMyCustomerAiQueueSheetP537_();
    const queueSummary = summarizeMyCustomerAiSheetByStatusP537_(queueSheet, '상태');
    const resultSummary = summarizeMyCustomerAiSheetByStatusP537_(analysisSheet, 'decision');
    const triggerCount = ScriptApp.getProjectTriggers().filter(function(t) {
      return t.getHandlerFunction && t.getHandlerFunction() === 'processMyCustomerAiAnalysisQueueTriggerP537';
    }).length;
    return {
      ok: true,
      version: MY_CUSTOMER_AI_ANALYZER_VERSION_P537,
      analysisSheet: MY_CUSTOMER_AI_ANALYSIS_SHEET_P537,
      queueSheet: MY_CUSTOMER_AI_QUEUE_SHEET_P537,
      spreadsheetUrl: getWebAppDbSpreadsheet_().getUrl(),
      queue: queueSummary,
      analysis: resultSummary,
      triggerCount: triggerCount,
      triggerStatus: triggerCount === 0 ? '미설치' : (triggerCount === 1 ? '정상' : '중복주의')
    };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err || '') };
  }
}

function summarizeMyCustomerAiSheetByStatusP537_(sheet, statusHeader) {
  const out = { total: 0, byStatus: {}, lastAt: '', recentError: '' };
  if (!sheet || sheet.getLastRow() < 2) return out;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const map = {};
  headers.forEach(function(h, i) { if (h) map[h] = i; });
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();
  out.total = values.length;
  const statusIdx = map[statusHeader];
  const modIdx = map['수정일시'] != null ? map['수정일시'] : map['분석일시'];
  const errIdx = map['마지막오류'] != null ? map['마지막오류'] : map['blockedReason'];
  values.forEach(function(row) {
    const st = String(statusIdx == null ? '' : row[statusIdx] || '').trim() || '(공란)';
    out.byStatus[st] = (out.byStatus[st] || 0) + 1;
    const at = String(modIdx == null ? '' : row[modIdx] || '').trim();
    if (at) out.lastAt = at;
    const er = String(errIdx == null ? '' : row[errIdx] || '').trim();
    if (er) out.recentError = shortenMyCustomerStatusTextP528_(er, 160);
  });
  return out;
}

function createMyCustomerAiAnalysisQueueP537(options) {
  options = options || {};
  const limit = Math.max(1, Math.min(500, Number(options.limit || MY_CUSTOMER_AI_DEFAULT_QUEUE_LIMIT_P537) || MY_CUSTOMER_AI_DEFAULT_QUEUE_LIMIT_P537));
  const includeProtected = !!options.includeProtected;
  const perm = getPortalCurrentPermission_();
  const aliases = buildMyCustomerStatusOwnerAliasesP528_(perm);
  const scopeInfo = resolveMyCustomerStatusScopeP532_(perm);
  const analysisSheet = ensureMyCustomerStatusAnalysisSheetP529_();
  const queueSheet = ensureMyCustomerAiQueueSheetP537_();
  ensureMyCustomerAiAnalysisSheetP537_();

  const analysisRows = readMyCustomerStatusAnalysisDbRowsP537_(analysisSheet);
  const existingHashes = getMyCustomerAiExistingInputHashMapP537_(queueSheet);
  const candidates = [];
  analysisRows.forEach(function(row) {
    if (!scopeInfo.isAllScope && !isMyCustomerAiAnalysisRowAllowedP537_(row, aliases)) return;
    const candidate = buildMyCustomerAiCandidateFromAnalysisRowP537_(row, includeProtected);
    if (!candidate || !candidate.inputHash) return;
    if (existingHashes[candidate.inputHash]) return;
    candidates.push(candidate);
  });
  candidates.sort(function(a, b) {
    if ((b.priority || 0) !== (a.priority || 0)) return (b.priority || 0) - (a.priority || 0);
    return String(a.company || '').localeCompare(String(b.company || ''), 'ko');
  });

  const selected = candidates.slice(0, limit);
  const nowText = formatMyCustomerStatusDateTimeP528_(new Date());
  const model = getMyCustomerAiModelP537_();
  const headers = queueSheet.getRange(1, 1, 1, queueSheet.getLastColumn()).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const rows = selected.map(function(c) {
    const rec = {
      '등록일시': nowText,
      '수정일시': nowText,
      '작업ID': 'MCAIQ-' + c.customerNo + '-' + c.rowNo + '-' + String(c.inputHash).slice(0, 10),
      '고객번호': c.customerNo || '',
      'rowNo': c.rowNo || '',
      '회사명': c.company || '',
      '영업담당자': c.salesRep || '',
      '현재상태': c.currentStatus || '',
      '상태': 'QUEUED',
      '우선순위': c.priority || 0,
      '시도횟수': 0,
      'inputHash': c.inputHash || '',
      'candidateJson': safeStringifyMyCustomerAiP537_(c, 6500),
      'resultJson': '',
      '마지막오류': '',
      '적용일시': '',
      '모델': model,
      '분석ID': c.analysisId || ''
    };
    return headers.map(function(h) { return rec[h] == null ? '' : rec[h]; });
  });
  if (rows.length) queueSheet.getRange(queueSheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  return {
    ok: true,
    created: rows.length,
    candidates: candidates.length,
    skippedDuplicate: candidates.length - selected.length,
    limit: limit,
    scope: scopeInfo.isAllScope ? 'ALL' : 'OWN',
    queueSheet: MY_CUSTOMER_AI_QUEUE_SHEET_P537,
    analysisSheet: MY_CUSTOMER_AI_ANALYSIS_SHEET_P537,
    url: getWebAppDbSpreadsheet_().getUrl()
  };
}

function readMyCustomerStatusAnalysisDbRowsP537_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();
  return values.map(function(r, i) {
    const obj = { __rowIndex: i + 2 };
    headers.forEach(function(h, c) { if (h) obj[h] = r[c]; });
    return obj;
  });
}

function isMyCustomerAiAnalysisRowAllowedP537_(row, aliases) {
  const rep = normalizeMyCustomerStatusNameP528_(row && row['영업담당자'] || '');
  if (!rep) return false;
  return (aliases || []).some(function(a) {
    const key = normalizeMyCustomerStatusNameP528_(a);
    return key && (rep === key || rep.indexOf(key) >= 0 || key.indexOf(rep) >= 0);
  });
}

function getMyCustomerAiExistingInputHashMapP537_(queueSheet) {
  const out = {};
  if (!queueSheet || queueSheet.getLastRow() < 2) return out;
  const lastCol = queueSheet.getLastColumn();
  const headers = queueSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const map = {};
  headers.forEach(function(h, i) { if (h) map[h] = i; });
  const values = queueSheet.getRange(2, 1, queueSheet.getLastRow() - 1, lastCol).getDisplayValues();
  values.forEach(function(row) {
    const hash = String(row[map['inputHash']] || '').trim();
    const status = String(row[map['상태']] || '').trim();
    if (!hash) return;
    if (['QUEUED','RUNNING','DONE','RETRY','RETRY_API'].indexOf(status) >= 0) out[hash] = true;
  });
  return out;
}

function buildMyCustomerAiCandidateFromAnalysisRowP537_(row, includeProtected) {
  row = row || {};
  const currentStatus = String(row['현재상태'] || '').trim();
  const protectedYn = String(row['상태보호여부'] || '').trim().toUpperCase() === 'Y';
  if (protectedYn && !includeProtected && currentStatus !== '발주완료') return null;
  const stateChangeYn = String(row['상태변경추천여부'] || row['추천노출여부'] || '').trim().toUpperCase() === 'Y';
  const confidence = Number(row['신뢰도'] || 0) || 0;
  const insightType = String(row['업무인사이트유형'] || '').trim();
  const currentNeedsStatus = !currentStatus || currentStatus.indexOf('상태지정') >= 0;
  const caution = String(row['분석주의등급'] || '').trim();
  const hasMixedSignals = !!String(row['강한부정키워드'] || row['계약진행키워드'] || row['장기추진키워드'] || '').trim();
  let priority = 0;
  let reason = '';
  if (stateChangeYn) { priority = 100; reason = '규칙 기반 상태변경 추천 후보'; }
  else if (currentNeedsStatus && hasMixedSignals) { priority = 82; reason = '상태지정필요 + 중요 키워드'; }
  else if (confidence >= 40 && confidence <= 75 && hasMixedSignals) { priority = 72; reason = '중간 신뢰도 + 혼합 신호'; }
  else if (['계약/발주확인필요','장기재접촉','자료발송후미후속','견적후속필요','타사비교','연락장애/담당자확인'].indexOf(insightType) >= 0 && confidence <= 80) { priority = 55; reason = '업무 인사이트 검수 후보'; }
  else if (caution && caution !== 'LOW') { priority = 50; reason = '분석주의등급 검수 후보'; }
  if (!priority) return null;

  const evidence = parseJsonSafeMyCustomerAiP537_(row['근거JSON']) || {};
  const events = parseJsonSafeMyCustomerAiP537_(row['이벤트JSON']);
  const compactEvents = Array.isArray(events) ? events
    .filter(function(ev) {
      const txt = ev && (ev.text || ev.summary || '');
      return txt && !isMyCustomerAiSystemMetaTextP544_(txt);
    })
    .slice(0, 5)
    .map(function(ev) {
      return slimMyCustomerAiEventP544_(ev, 180);
    }) : [];
  const contractInfo = evidence && evidence.contractCompletionInfo ? evidence.contractCompletionInfo : {};
  const candidate = {
    analysisId: row['분석ID'] || '',
    customerNo: row['고객번호'] || '',
    rowNo: row['rowNo'] || '',
    company: row['회사명'] || '',
    salesRep: row['영업담당자'] || '',
    currentStatus: currentStatus,
    ruleRecommendedStatus: row['상태변경추천상태'] || row['추천상태'] || currentStatus,
    ruleStateChangeRecommended: stateChangeYn,
    confidence: confidence,
    insightType: insightType,
    insightSummary: row['업무인사이트요약'] || '',
    reason: row['보수판정사유'] || row['판정유형'] || reason,
    latestDecisionEvent: {
      date: row['최신판정이벤트일자'] || '',
      source: row['최신판정이벤트출처'] || row['최근이벤트출처'] || '',
      text: row['최신판정이벤트요약'] || row['최근이벤트요약'] || ''
    },
    latestAnyEvent: { text: row['최신전체이벤트요약'] || '' },
    keywords: {
      negative: row['강한부정키워드'] || '',
      order: row['계약진행키워드'] || '',
      quote: row['자료발송키워드'] || '',
      long: row['장기추진키워드'] || '',
      data: row['데이터누락키워드'] || ''
    },
    events: compactEvents,
    contractCompletionInfo: {
      ready: !!contractInfo.ready,
      summary: shortenMyCustomerStatusTextP528_(contractInfo.summary || '', 200),
      orderNo: contractInfo.orderNo || '',
      contractNo: contractInfo.contractNo || ''
    },
    priority: priority,
    queueReason: reason
  };
  const skipReason = shouldSkipMyCustomerAiCandidateP544_(candidate);
  if (skipReason) return null;
  candidate.inputHash = hashMyCustomerStatusTextP529_(safeStringifyMyCustomerAiP537_({
    v: MY_CUSTOMER_AI_PROMPT_VERSION_P537,
    customerNo: candidate.customerNo,
    rowNo: candidate.rowNo,
    currentStatus: candidate.currentStatus,
    ruleRecommendedStatus: candidate.ruleRecommendedStatus,
    insightType: candidate.insightType,
    latestDecisionEvent: candidate.latestDecisionEvent,
    events: candidate.events,
    keywords: candidate.keywords,
    contractReady: candidate.contractCompletionInfo && candidate.contractCompletionInfo.ready
  }, 20000));
  return candidate;
}

function processMyCustomerAiAnalysisQueueTriggerP537() {
  return processMyCustomerAiAnalysisQueueP537({ limit: MY_CUSTOMER_AI_DEFAULT_PROCESS_LIMIT_P537, trigger: true });
}

function processMyCustomerAiAnalysisQueueP537(options) {
  options = options || {};
  const limit = Math.max(1, Math.min(50, Number(options.limit || MY_CUSTOMER_AI_DEFAULT_PROCESS_LIMIT_P537) || MY_CUSTOMER_AI_DEFAULT_PROCESS_LIMIT_P537));
  const apiKey = getMyCustomerAiApiKeyP537_();
  if (!apiKey) throw new Error('Script Properties에 OPENAI_API_KEY를 설정해야 GPT 보조 분석을 실행할 수 있습니다.');

  const queueSheet = ensureMyCustomerAiQueueSheetP537_();
  const analysisSheet = ensureMyCustomerAiAnalysisSheetP537_();
  if (queueSheet.getLastRow() < 2) return { ok: true, processed: 0, message: '처리할 AI 분석 큐가 없습니다.' };

  const qHeaders = queueSheet.getRange(1, 1, 1, queueSheet.getLastColumn()).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const qMap = {};
  qHeaders.forEach(function(h, i) { if (h) qMap[h] = i; });
  const values = queueSheet.getRange(2, 1, queueSheet.getLastRow() - 1, qHeaders.length).getDisplayValues();
  const targets = [];
  values.forEach(function(row, i) {
    if (targets.length >= limit) return;
    const st = String(row[qMap['상태']] || '').trim();
    if (['QUEUED','RETRY','RETRY_API'].indexOf(st) < 0) return;
    targets.push({ sheetRow: i + 2, values: row });
  });
  if (!targets.length) return { ok: true, processed: 0, message: 'QUEUED/RETRY 상태의 AI 분석 큐가 없습니다.' };

  const result = { ok: true, processed: 0, done: 0, retry: 0, fail: 0, errors: [] };
  targets.forEach(function(job) {
    const row = job.values.slice();
    const sheetRow = job.sheetRow;
    const nowText = formatMyCustomerStatusDateTimeP528_(new Date());
    const attempts = (Number(row[qMap['시도횟수']] || 0) || 0) + 1;
    setMyCustomerAiQueueRowStatusP537_(queueSheet, qMap, qHeaders.length, sheetRow, row, 'RUNNING', attempts, '', '');
    try {
      const candidate = parseJsonSafeMyCustomerAiP537_(row[qMap['candidateJson']]);
      if (!candidate || !candidate.customerNo) throw new Error('candidateJson이 올바르지 않습니다.');
      const request = buildMyCustomerAiOpenAiRequestP537_(candidate);
      const apiRes = callOpenAiForMyCustomerStatusP537_(apiKey, request);
      const ai = normalizeMyCustomerAiApiResultP537_(apiRes);
      const validated = validateMyCustomerAiRecommendationP537_(candidate, ai);
      const saveRes = appendMyCustomerAiAnalysisResultP537_(analysisSheet, candidate, ai, validated, request, apiRes);
      setMyCustomerAiQueueRowStatusP537_(queueSheet, qMap, qHeaders.length, sheetRow, row, 'DONE', attempts, safeStringifyMyCustomerAiP537_({ ai: ai, validated: validated, save: saveRes }, 8000), '');
      result.done++;
    } catch (err) {
      const msg = err && err.message ? err.message : String(err || '');
      const nextStatus = getMyCustomerAiQueueErrorStatusP544_(err, attempts);
      setMyCustomerAiQueueRowStatusP537_(queueSheet, qMap, qHeaders.length, sheetRow, row, nextStatus, attempts, '', msg);
      if (String(nextStatus).indexOf('FAIL') === 0) result.fail++; else result.retry++;
      result.errors.push({ row: sheetRow, status: nextStatus, error: msg });
    }
    result.processed++;
  });
  return result;
}

function setMyCustomerAiQueueRowStatusP537_(sheet, map, width, rowNo, row, status, attempts, resultJson, error) {
  const arr = row.slice();
  const nowText = formatMyCustomerStatusDateTimeP528_(new Date());
  function set(header, value) { if (map[header] != null) arr[map[header]] = value; }
  set('수정일시', nowText);
  set('상태', status);
  if (attempts != null) set('시도횟수', attempts);
  if (resultJson != null) set('resultJson', resultJson);
  if (error != null) set('마지막오류', error);
  if (status === 'DONE') set('적용일시', nowText);
  sheet.getRange(rowNo, 1, 1, width).setValues([arr]);
  try { SpreadsheetApp.flush(); } catch (err) {}
}

function buildMyCustomerAiOpenAiRequestP537_(candidate) {
  const schema = getMyCustomerAiStructuredSchemaP537_();
  const promptInput = buildMyCustomerAiPromptInputP537_(candidate);
  return {
    model: getMyCustomerAiModelP537_(),
    max_output_tokens: Number(PropertiesService.getScriptProperties().getProperty('OPENAI_CUSTOMER_AI_MAX_OUTPUT_TOKENS') || 900) || 900,
    input: [
      { role: 'system', content: buildMyCustomerAiSystemPromptP537_() },
      { role: 'user', content: JSON.stringify(promptInput) }
    ],
    text: {
      format: {
        type: 'json_schema',
        name: 'customer_status_ai_analysis',
        strict: true,
        schema: schema
      }
    }
  };
}

function buildMyCustomerAiSystemPromptP537_() {
  return [
    '당신은 S1 Sales Portal의 영업 상태 분석 보조자입니다.',
    '목표는 상태값을 자동으로 바꾸는 것이 아니라, 규칙 기반 분석의 애매건을 검토하고 오탐을 줄이는 것입니다.',
    '상태 정의:',
    '- 견적제출완료: 견적 제출 후 영업팀이 추가 컨택 영업해야 하는 상태.',
    '- 장기 추진건: 금년도 대상이 아니거나 기타 사유로 내년/추후 컨택해야 하는 상태.',
    '- 고객 설득 중: 영업팀이 추가 컨택 영업중인 상태.',
    '- 발주완료: 수주 확정됐지만 필수 서류 미취합.',
    '- 계약완료: 수주 확정 + 필수 서류 취합 완료.',
    '- 수주실패: 확실한 거절/타사 확정/진행 불가.',
    '절대 규칙:',
    '1. 어떤 경우에도 발주완료를 추천하지 마세요.',
    '2. 수주실패/계약완료는 보호 상태이므로 다른 상태로 변경 추천하지 마세요.',
    '3. 발주완료는 보호 상태지만, 서버 규칙이 contractCompletionInfo.ready=true라고 제공한 경우에만 계약완료 전환은 허용됩니다. 이 경우 서버 판단을 존중하세요.',
    '4. 계약완료 추천은 메모만으로 판단하지 말고 서버가 제공한 contractCompletionInfo.ready=true일 때만 동의할 수 있습니다.',
    '5. 타사 계약완료는 우리 계약완료가 아니라 수주실패 신호입니다.',
    '6. 전화연결 불가/담당자 부재/직통번호 안내 불가는 수주실패가 아니라 연락장애입니다.',
    '7. 과거 키워드보다 최신 고객 의사 이벤트를 우선하세요.',
    '8. 애매하면 STATUS_CHANGE가 아니라 HUMAN_REVIEW 또는 INSIGHT_ONLY를 선택하세요.',
    '출력은 반드시 JSON Schema에 맞는 JSON만 반환하고, 긴 근거문/원문 인용을 반복하지 마세요.'
  ].join('\n');
}

function buildMyCustomerAiPromptInputP537_(candidate) {
  const contractInfo = candidate.contractCompletionInfo || {};
  return {
    promptVersion: MY_CUSTOMER_AI_PROMPT_VERSION_P537,
    customer: {
      customerNo: candidate.customerNo || '',
      rowNo: candidate.rowNo || '',
      company: candidate.company || '',
      salesRep: candidate.salesRep || '',
      currentStatus: candidate.currentStatus || ''
    },
    ruleAnalysis: {
      ruleRecommendedStatus: candidate.ruleRecommendedStatus || '',
      ruleStateChangeRecommended: !!candidate.ruleStateChangeRecommended,
      confidence: candidate.confidence || 0,
      insightType: candidate.insightType || '',
      insightSummary: shortenMyCustomerStatusTextP528_(candidate.insightSummary || '', 180),
      reason: shortenMyCustomerStatusTextP528_(candidate.reason || '', 180),
      queueReason: candidate.queueReason || ''
    },
    latestDecisionEvent: slimMyCustomerAiEventP544_(candidate.latestDecisionEvent, 180),
    latestAnyEvent: slimMyCustomerAiEventP544_(candidate.latestAnyEvent, 160),
    recentEvents: (candidate.events || []).slice(0, 5).map(function(ev) { return slimMyCustomerAiEventP544_(ev, 180); }),
    contractCompletionInfo: {
      ready: !!contractInfo.ready,
      summary: shortenMyCustomerStatusTextP528_(contractInfo.summary || '', 180),
      orderNo: contractInfo.orderNo || '',
      contractNo: contractInfo.contractNo || ''
    },
    allowedTransitionReminder: {
      neverRecommend: ['발주완료'],
      protectedStatuses: ['수주실패', '계약완료'],
      orderCompleteException: '현재상태가 발주완료이고 contractCompletionInfo.ready=true이면 계약완료 전환만 허용',
      uncertainShouldBe: 'HUMAN_REVIEW 또는 INSIGHT_ONLY'
    }
  };
}

function getMyCustomerAiStructuredSchemaP537_() {
  return {
    type: 'object',
    additionalProperties: false,
    required: ['decision','recommendedStatus','recommendationAllowed','confidence','falsePositiveRisk','insightType','summary','reason','blockedReason','nextAction'],
    properties: {
      decision: { type: 'string', enum: ['STATUS_CHANGE','INSIGHT_ONLY','KEEP_CURRENT','HUMAN_REVIEW'] },
      recommendedStatus: { type: 'string', enum: ['고객 설득 중','견적제출완료','장기 추진건','수주실패','계약완료','현재상태유지'] },
      recommendationAllowed: { type: 'boolean' },
      confidence: { type: 'number', minimum: 0, maximum: 100 },
      falsePositiveRisk: { type: 'string', enum: ['LOW','MEDIUM','HIGH'] },
      insightType: { type: 'string', enum: ['자료발송후미후속','연락장애','장기재접촉','타사비교','타사선정','계약발주확인필요','데이터확인필요','진행중추적','없음'] },
      summary: { type: 'string' },
      reason: { type: 'string' },
      blockedReason: { type: 'string' },
      nextAction: { type: 'string' }
    }
  };
}

function callOpenAiForMyCustomerStatusP537_(apiKey, request) {
  const started = new Date();
  const res = UrlFetchApp.fetch(MY_CUSTOMER_AI_API_URL_P537, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(request)
  });
  const code = res.getResponseCode();
  const text = res.getContentText() || '';
  let json = null;
  try { json = JSON.parse(text); } catch (err) {}
  if (code < 200 || code >= 300) {
    const msg = json && json.error && (json.error.message || json.error.code) || text.slice(0, 500) || ('HTTP ' + code);
    const e = new Error('OpenAI API 오류(' + code + '): ' + msg);
    e.aiStatus = (code === 429 || code >= 500) ? 'RETRY_API' : 'FAIL_API';
    e.retryable = (code === 429 || code >= 500);
    throw e;
  }
  if (!json) {
    const e = new Error('OpenAI API 응답 JSON 파싱 실패');
    e.aiStatus = 'FAIL_API';
    e.retryable = false;
    throw e;
  }
  json.__elapsedMsP537 = new Date().getTime() - started.getTime();
  return json;
}

function normalizeMyCustomerAiApiResultP537_(apiRes) {
  const text = extractOpenAiResponseTextP537_(apiRes);
  if (!text) {
    const e = new Error('OpenAI 응답에서 output_text를 찾지 못했습니다.');
    e.aiStatus = 'FAIL_PARSE';
    e.retryable = false;
    throw e;
  }
  const parsed = parseOpenAiStructuredJsonP544_(text);
  parsed.__rawText = text;
  parsed.__usage = apiRes && apiRes.usage || {};
  return parsed;
}

function extractOpenAiResponseTextP537_(apiRes) {
  if (!apiRes) return '';
  if (apiRes.output_text) return String(apiRes.output_text || '');
  const out = [];
  (apiRes.output || []).forEach(function(item) {
    (item.content || []).forEach(function(c) {
      if (c && c.text) out.push(c.text);
      if (c && c.type === 'output_text' && c.text) out.push(c.text);
    });
  });
  if (out.length) return out.join('\n');
  try {
    const msg = apiRes.choices && apiRes.choices[0] && apiRes.choices[0].message && apiRes.choices[0].message.content;
    if (msg) return String(msg);
  } catch (err) {}
  return '';
}

function validateMyCustomerAiRecommendationP537_(candidate, ai) {
  candidate = candidate || {};
  ai = ai || {};
  const currentStatus = String(candidate.currentStatus || '').trim();
  let target = String(ai.recommendedStatus || '').trim();
  if (!target || target === '현재상태유지') target = currentStatus;
  let serverAllowed = !!ai.recommendationAllowed && ai.decision === 'STATUS_CHANGE';
  const reasons = [];
  if (target === '발주완료') { serverAllowed = false; reasons.push('발주완료는 AI/자동추천 금지'); }
  if (['수주실패','계약완료'].indexOf(currentStatus) >= 0 && target !== currentStatus) { serverAllowed = false; reasons.push('보호 상태 변경 금지: ' + currentStatus); }
  if (currentStatus === '발주완료' && target !== '계약완료' && target !== currentStatus) { serverAllowed = false; reasons.push('발주완료는 계약완료만 추천 가능'); }
  if (String(ai.falsePositiveRisk || '').toUpperCase() === 'HIGH') { serverAllowed = false; reasons.push('오탐위험 HIGH'); }
  if (target && target !== currentStatus) {
    try {
      const detail = getCustomerDetail(Number(candidate.rowNo || 0) || candidate.rowNo);
      const contractInfo = getMyCustomerContractCompletionInfoP536_(detail || {}, getMyCustomerContractCompleteMapP536_());
      const ok = isMyCustomerStatusAllowedTransitionP536_(currentStatus, target, {
        contractCompleteReady: !!(contractInfo && contractInfo.ready),
        strongFailure: target === '수주실패'
      });
      if (!ok) { serverAllowed = false; reasons.push(getMyCustomerStatusTransitionBlockReasonP536_(currentStatus, target, contractInfo)); }
    } catch (err) {
      serverAllowed = false;
      reasons.push('서버 전이 검증 실패: ' + (err && err.message ? err.message : String(err || '')));
    }
  } else {
    serverAllowed = false;
    reasons.push('현재상태 유지 또는 추천상태 없음');
  }
  return {
    serverAllowed: serverAllowed,
    finalStatus: serverAllowed ? target : currentStatus,
    blockedReason: reasons.join(' / ') || String(ai.blockedReason || ''),
    targetStatus: target
  };
}

function appendMyCustomerAiAnalysisResultP537_(sheet, candidate, ai, validated, request, apiRes) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(function(h) { return String(h || '').trim(); });
  const usage = ai.__usage || apiRes.usage || {};
  const tokensIn = usage.input_tokens || usage.prompt_tokens || 0;
  const tokensOut = usage.output_tokens || usage.completion_tokens || 0;
  const model = getMyCustomerAiModelP537_();
  const cost = estimateMyCustomerAiCostP537_(model, tokensIn, tokensOut);
  const nowText = formatMyCustomerStatusDateTimeP528_(new Date());
  const analysisId = 'MCAI-' + String(candidate.customerNo || '') + '-' + String(candidate.rowNo || '') + '-' + String(candidate.inputHash || '').slice(0, 8);
  const rec = {
    '분석일시': nowText,
    '분석ID': analysisId,
    '고객번호': candidate.customerNo || '',
    'rowNo': candidate.rowNo || '',
    '회사명': candidate.company || '',
    '영업담당자': candidate.salesRep || '',
    '현재상태': candidate.currentStatus || '',
    '규칙추천상태': candidate.ruleRecommendedStatus || '',
    'GPT추천상태': ai.recommendedStatus || '',
    '최종추천상태': validated.finalStatus || candidate.currentStatus || '',
    'decision': ai.decision || '',
    'recommendationAllowed': ai.recommendationAllowed ? 'Y' : 'N',
    'serverAllowed': validated.serverAllowed ? 'Y' : 'N',
    'confidence': ai.confidence || '',
    'falsePositiveRisk': ai.falsePositiveRisk || '',
    'insightType': ai.insightType || '',
    'summary': ai.summary || '',
    'reason': ai.reason || '',
    'nextAction': ai.nextAction || '',
    'evidenceJson': safeStringifyMyCustomerAiP537_(ai.evidence || [], 1200),
    'blockedReason': validated.blockedReason || ai.blockedReason || '',
    'inputHash': candidate.inputHash || '',
    'promptVersion': MY_CUSTOMER_AI_PROMPT_VERSION_P537,
    'model': model,
    'tokensInput': tokensIn,
    'tokensOutput': tokensOut,
    'estimatedCostUsd': cost,
    'rawRequestJson': safeStringifyMyCustomerAiP537_(redactMyCustomerAiRequestForLogP537_(request), 7000),
    'rawResponseJson': safeStringifyMyCustomerAiP537_(apiRes, 9000),
    '검토상태': '',
    '검토자': '',
    '검토일시': '',
    '검토메모': ''
  };
  const row = headers.map(function(h) { return rec[h] == null ? '' : rec[h]; });
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([row]);
  return { ok: true, analysisId: analysisId, sheetName: MY_CUSTOMER_AI_ANALYSIS_SHEET_P537 };
}

function getMyCustomerAiQueueErrorStatusP544_(err, attempts) {
  const status = err && err.aiStatus ? String(err.aiStatus) : '';
  if (status === 'FAIL_PARSE' || status === 'FAIL_SCHEMA' || status === 'FAIL_API') return status;
  if (status === 'RETRY_API') return attempts >= MY_CUSTOMER_AI_MAX_RETRY_P537 ? 'FAIL_API' : 'RETRY_API';
  const msg = err && err.message ? String(err.message) : String(err || '');
  if (/JSON 결과 파싱 실패|output_text를 찾지 못했습니다|스키마|schema/i.test(msg)) return 'FAIL_PARSE';
  if (/OpenAI API 오류\((429|5\d\d)\)/.test(msg)) return attempts >= MY_CUSTOMER_AI_MAX_RETRY_P537 ? 'FAIL_API' : 'RETRY_API';
  return attempts >= MY_CUSTOMER_AI_MAX_RETRY_P537 ? 'FAIL' : 'RETRY';
}

function parseOpenAiStructuredJsonP544_(text) {
  let src = String(text || '').trim();
  src = src.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();
  let parsed = null;
  try { parsed = JSON.parse(src); } catch (err1) {
    const start = src.indexOf('{');
    const end = src.lastIndexOf('}');
    if (start >= 0 && end > start) {
      try { parsed = JSON.parse(src.slice(start, end + 1)); } catch (err2) {}
    }
  }
  if (!parsed) {
    const e = new Error('OpenAI JSON 결과 파싱 실패: ' + src.slice(0, 700));
    e.aiStatus = 'FAIL_PARSE';
    e.retryable = false;
    throw e;
  }
  return parsed;
}

function slimMyCustomerAiEventP544_(ev, maxLen) {
  ev = ev || {};
  return {
    source: ev.sourceLabel || ev.source || '',
    date: ev.dateText || ev.date || '',
    text: shortenMyCustomerStatusTextP528_(ev.text || ev.summary || '', maxLen || 180)
  };
}

function isMyCustomerAiSystemMetaTextP544_(text) {
  text = String(text || '');
  return /(기존\s*담당자별\s*시트|마스터시트로\s*이관|이관\.|중복\s*삭제|데이터\s*병합|정보\s*확인해보시고|TM\s*콜\s*원하시면|삭제\s*고객번호|대표전화번호|담당자\s*이메일\s*주소|연면적)/i.test(text);
}

function isMyCustomerAiContactBlockerOnlyTextP544_(text) {
  text = String(text || '');
  if (!text) return false;
  const hasBlocker = /(부재|전화\s*연결\s*불가|전화통화\s*불가|연락\s*안됨|통화\s*안됨|자리\s*비움|안내\s*불가|담당자\s*부재|연락처\s*안내\s*거절|통화종료)/.test(text);
  const hasStrongDecision = /(타사|타업체|다른\s*업체|기존\s*업체|계약\s*완료|계약완료|안\s*한다고|필요\s*없다고|거절|내년|예산|견적\s*요청|견적서|자료\s*발송|발송\s*완료|검토중|결정)/.test(text);
  return hasBlocker && !hasStrongDecision;
}

function hasMyCustomerAiMaterialSendTextP544_(text) {
  text = String(text || '');
  if (/테스트\s*발송|내게\s*테스트|나에게\s*테스트/.test(text)) return false;
  return /(자료\s*발송|발송\s*완료|견적서\s*발송|견적\s*발송|견적요청|고객발송)/.test(text) && /(견적서|견적|자료\s*발송|고객발송|발송\s*완료)/.test(text);
}

function shouldSkipMyCustomerAiCandidateP544_(candidate) {
  candidate = candidate || {};
  const latestDecisionText = String(candidate.latestDecisionEvent && candidate.latestDecisionEvent.text || '').trim();
  const latestAnyText = String(candidate.latestAnyEvent && candidate.latestAnyEvent.text || '').trim();
  const combinedLatest = [latestDecisionText, latestAnyText].join('\n');
  const ruleStatus = String(candidate.ruleRecommendedStatus || '').trim();
  if (isMyCustomerAiSystemMetaTextP544_(combinedLatest)) return 'SYSTEM_META_LATEST_EVENT';
  if (isMyCustomerAiContactBlockerOnlyTextP544_(combinedLatest)) return 'CONTACT_BLOCKER_ONLY_LATEST_EVENT';
  if (ruleStatus === '견적제출완료') {
    const latestHasMaterialSend = hasMyCustomerAiMaterialSendTextP544_(latestDecisionText) || hasMyCustomerAiMaterialSendTextP544_(latestAnyText);
    if (!latestHasMaterialSend && !latestDecisionText) return 'WEAK_PAST_QUOTE_ONLY';
  }
  return '';
}

function installMyCustomerAiAnalysisQueueTriggerP537() {
  const fn = 'processMyCustomerAiAnalysisQueueTriggerP537';
  const existing = ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction && t.getHandlerFunction() === fn; });
  if (!existing.length) ScriptApp.newTrigger(fn).timeBased().everyMinutes(5).create();
  return { ok: true, functionName: fn, existing: existing.length, installed: !existing.length };
}

function getMyCustomerAiApiKeyP537_() {
  return String(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY') || '').trim();
}

function getMyCustomerAiModelP537_() {
  return String(PropertiesService.getScriptProperties().getProperty('OPENAI_CUSTOMER_AI_MODEL') || MY_CUSTOMER_AI_DEFAULT_MODEL_P537).trim();
}

function estimateMyCustomerAiCostP537_(model, inputTokens, outputTokens) {
  const props = PropertiesService.getScriptProperties();
  let inRate = Number(props.getProperty('OPENAI_CUSTOMER_AI_INPUT_USD_PER_1M') || 0) || 0;
  let outRate = Number(props.getProperty('OPENAI_CUSTOMER_AI_OUTPUT_USD_PER_1M') || 0) || 0;
  const m = String(model || '').toLowerCase();
  if (!inRate && m.indexOf('luna') >= 0) inRate = 1;
  if (!outRate && m.indexOf('luna') >= 0) outRate = 6;
  if (!inRate && m.indexOf('terra') >= 0) inRate = 2.5;
  if (!outRate && m.indexOf('terra') >= 0) outRate = 15;
  const cost = ((Number(inputTokens || 0) * inRate) + (Number(outputTokens || 0) * outRate)) / 1000000;
  return cost ? Math.round(cost * 1000000) / 1000000 : '';
}

function parseJsonSafeMyCustomerAiP537_(value) {
  const text = String(value || '').trim();
  if (!text) return null;
  try { return JSON.parse(text); } catch (err) { return null; }
}

function safeStringifyMyCustomerAiP537_(value, maxLen) {
  let s = '';
  try { s = JSON.stringify(value == null ? {} : value); } catch (err) { s = String(value || ''); }
  maxLen = Number(maxLen || 0) || 0;
  if (maxLen && s.length > maxLen) return s.slice(0, maxLen) + '...TRUNCATED';
  return s;
}

function redactMyCustomerAiRequestForLogP537_(request) {
  const copy = JSON.parse(JSON.stringify(request || {}));
  return copy;
}
