/***************************************
 * S1 Sales Portal - 30_MyCustomerAiAnalysisService.gs
 * P547: 나의 고객 현황 - AI 독립판정 / 근거 ID / 오탐 차단
 * - P537 공개 함수명은 기존 호출 호환성을 위해 유지합니다.
 * - 규칙 추천값은 큐 선별·사후 대조에만 쓰고 OpenAI 입력에서는 제외합니다.
 * - GPT 결과는 바로 상태값을 바꾸지 않고 서버 증거/전이 규칙으로 재검증합니다.
 ***************************************/

const MY_CUSTOMER_AI_ANALYZER_VERSION_P537 = 'P547_AI_INDEPENDENT_EVIDENCE_GUARD';
const MY_CUSTOMER_AI_PROMPT_VERSION_P537 = 'P547_STATUS_INDEPENDENT_EVIDENCE_V3';
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
  'falsePositiveRisk', 'insightType', 'summary', 'reason', 'nextAction', 'evidenceJson', 'evidenceIdsJson',
  'ambiguityFlagsJson', 'ruleAiMismatch', 'serverDecision', 'blockedReason', 'inputHash', 'promptVersion',
  'model', 'tokensInput', 'tokensOutput', 'estimatedCostUsd', 'rawRequestJson', 'rawResponseJson',
  '검토상태', '검토자', '검토일시', '검토메모'
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
  const parsedEvents = parseJsonSafeMyCustomerAiP537_(row['이벤트JSON']);
  const fallbackEvents = [
    {
      source: row['최신판정이벤트출처'] || row['최근이벤트출처'] || '',
      date: row['최신판정이벤트일자'] || '',
      text: row['최신판정이벤트요약'] || row['최근이벤트요약'] || '',
      declaredLatest: true
    },
    {
      source: row['최신전체이벤트출처'] || '',
      date: row['최신전체이벤트일자'] || '',
      text: row['최신전체이벤트요약'] || '',
      declaredLatest: true
    }
  ];
  const evidencePack = buildMyCustomerAiEvidencePackP547_(Array.isArray(parsedEvents) ? parsedEvents : [], fallbackEvents);
  const contractInfo = evidence && evidence.contractCompletionInfo ? evidence.contractCompletionInfo : {};
  const candidate = {
    analysisId: row['분석ID'] || '',
    customerNo: row['고객번호'] || '',
    rowNo: row['rowNo'] || '',
    company: row['회사명'] || '',
    salesRep: row['영업담당자'] || '',
    currentStatus: currentStatus,

    // 규칙 결과는 큐 우선순위 및 사후 대조/로그에만 사용합니다.
    // OpenAI 프롬프트에는 전달하지 않아 독립 판단을 보장합니다.
    ruleRecommendedStatus: row['상태변경추천상태'] || row['추천상태'] || currentStatus,
    ruleStateChangeRecommended: stateChangeYn,
    confidence: confidence,
    insightType: insightType,
    insightSummary: row['업무인사이트요약'] || '',
    reason: row['보수판정사유'] || row['판정유형'] || reason,

    latestDecisionEvent: evidencePack.latestDecisionEvent || {},
    latestAnyEvent: evidencePack.latestAnyEvent || {},
    events: evidencePack.events.slice(0, 8),
    ambiguityFlags: evidencePack.ambiguityFlags,
    dateCoverage: evidencePack.dateCoverage,
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
  if (!candidate.events.length && !(candidate.currentStatus === '발주완료' && candidate.contractCompletionInfo.ready)) return null;

  candidate.inputHash = hashMyCustomerStatusTextP529_(safeStringifyMyCustomerAiP537_({
    v: MY_CUSTOMER_AI_PROMPT_VERSION_P537,
    customerNo: candidate.customerNo,
    rowNo: candidate.rowNo,
    currentStatus: candidate.currentStatus,
    events: candidate.events,
    ambiguityFlags: candidate.ambiguityFlags,
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
      const parsedCandidate = parseJsonSafeMyCustomerAiP537_(row[qMap['candidateJson']]);
      if (!parsedCandidate || !parsedCandidate.customerNo) throw new Error('candidateJson이 올바르지 않습니다.');
      const candidate = upgradeMyCustomerAiCandidateP547_(parsedCandidate);
      const request = buildMyCustomerAiOpenAiRequestP537_(candidate);
      const apiRes = callOpenAiForMyCustomerStatusP537_(apiKey, request);
      const ai = normalizeMyCustomerAiApiResultP537_(apiRes, candidate);
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
    '당신은 S1 Sales Portal의 영업 상태 독립 검증자입니다.',
    '입력에는 규칙 엔진의 추천상태나 신뢰도가 제공되지 않습니다. evidenceEvents만 보고 독립적으로 판단하세요.',
    '목표는 자동변경을 많이 만드는 것이 아니라 오탐을 차단하고 사람이 검토할 애매건을 정확히 분리하는 것입니다.',
    '상태 정의:',
    '- 견적제출완료: 실제 견적/자료가 제출되었고 영업팀의 추가 컨택이 필요한 상태.',
    '- 장기 추진건: 금년도 대상이 아니거나 내년/추후/예산 편성 후 재접촉해야 하는 상태.',
    '- 고객 설득 중: 현재 영업 협의와 설득이 진행 중이나 견적 제출만으로 상태를 확정하기 어려운 상태.',
    '- 발주완료: 수주가 확정됐지만 필수 서류 미취합. 절대 추천하지 마세요.',
    '- 계약완료: 수주 확정 + 필수 서류 취합 완료. contractCompletionInfo.ready=true인 발주완료 고객만 추천할 수 있습니다.',
    '- 수주실패: 명확한 거절, 타사 선정/계약 확정, 진행 불가가 확인된 상태.',
    '판정 원칙:',
    '1. 현재상태가 수주실패 또는 계약완료인 경우에만 보호상태입니다. 추천하려는 목표상태가 수주실패라는 이유로 보호상태라고 해석하지 마세요.',
    '2. 타업체 견적을 받음, 가격 비교, 네고 문의, 견적가가 다양함은 타사 선정이 아닙니다. 이런 경우 타사비교 또는 진행중추적으로 분류하세요.',
    '3. 타사/타업체 선정, 타사로 결정, 타업체와 계약, 기존업체 유지가 명확하게 확인될 때만 수주실패 후보가 될 수 있습니다.',
    '4. 자체적으로 진행, 관리업체 있음, 전화가 끊김처럼 범위나 의도가 불명확한 최신 이력은 HUMAN_REVIEW로 처리하세요.',
    '5. 정상 통화에서 검토 중이거나 견적 재발송을 요청한 경우 연락장애로 분류하지 마세요.',
    '6. 과거 자료발송보다 최신 고객 의사 이벤트를 우선하고, 상충하는 이벤트가 있으면 HUMAN_REVIEW로 처리하세요.',
    '7. STATUS_CHANGE는 목표상태를 직접 지지하는 날짜가 있는 evidenceId가 최소 1개이고 중대한 ambiguityFlag가 없을 때만 선택하세요.',
    '8. 날짜가 없거나 근거가 약하면 HUMAN_REVIEW 또는 INSIGHT_ONLY를 선택하세요.',
    '9. STATUS_CHANGE가 아니면 recommendedStatus=현재상태유지, recommendationAllowed=false로 출력하세요.',
    '10. evidenceIds에는 실제 판단에 사용한 E01 형식의 ID만 1~4개 넣으세요. 입력에 없는 ID를 만들지 마세요.',
    'confidence는 독립적으로 산정하고 외부 점수를 추측하거나 복제하지 마세요.',
    '출력은 JSON Schema에 맞는 JSON만 반환하세요.'
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
    evidenceSummary: {
      latestEventId: candidate.latestAnyEvent && candidate.latestAnyEvent.evidenceId || '',
      latestDecisionEventId: candidate.latestDecisionEvent && candidate.latestDecisionEvent.evidenceId || '',
      ambiguityFlags: candidate.ambiguityFlags || [],
      dateCoverage: candidate.dateCoverage || {}
    },
    evidenceEvents: (candidate.events || []).slice(0, 8).map(function(ev) {
      return {
        evidenceId: ev.evidenceId || '',
        source: ev.sourceLabel || ev.source || '',
        date: ev.date || '',
        text: shortenMyCustomerStatusTextP528_(ev.text || '', 220)
      };
    }),
    contractCompletionInfo: {
      ready: !!contractInfo.ready,
      summary: shortenMyCustomerStatusTextP528_(contractInfo.summary || '', 180),
      orderNo: contractInfo.orderNo || '',
      contractNo: contractInfo.contractNo || ''
    },
    transitionConstraints: {
      neverRecommend: ['발주완료'],
      protectedOnlyWhenCurrentStatusIs: ['수주실패', '계약완료'],
      orderCompleteException: '현재상태가 발주완료이고 contractCompletionInfo.ready=true인 경우에만 계약완료 추천 가능',
      uncertainDecision: 'HUMAN_REVIEW 또는 INSIGHT_ONLY'
    }
  };
}

function getMyCustomerAiStructuredSchemaP537_() {
  return {
    type: 'object',
    additionalProperties: false,
    required: ['decision','recommendedStatus','recommendationAllowed','confidence','falsePositiveRisk','insightType','summary','reason','blockedReason','nextAction','evidenceIds','ambiguityFlags'],
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
      nextAction: { type: 'string' },
      evidenceIds: {
        type: 'array',
        items: { type: 'string' }
      },
      ambiguityFlags: {
        type: 'array',
        items: {
          type: 'string',
          enum: ['DATE_MISSING','CONFLICTING_EVENTS','SELF_HANDLED_AMBIGUOUS','EXISTING_VENDOR_AMBIGUOUS','CALL_ENDED_AMBIGUOUS','THIRD_PARTY_COMPARISON_ONLY','SCOPE_UNCLEAR','EVIDENCE_WEAK']
        }
      }
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

function normalizeMyCustomerAiApiResultP537_(apiRes, candidate) {
  const text = extractOpenAiResponseTextP537_(apiRes);
  if (!text) {
    const e = new Error('OpenAI 응답에서 output_text를 찾지 못했습니다.');
    e.aiStatus = 'FAIL_PARSE';
    e.retryable = false;
    throw e;
  }
  const parsed = parseOpenAiStructuredJsonP544_(text);
  const normalized = normalizeMyCustomerAiDecisionP547_(candidate || {}, parsed || {});
  normalized.__rawText = text;
  normalized.__usage = apiRes && apiRes.usage || {};
  return normalized;
}

function extractOpenAiResponseTextP537_(apiRes) {
  if (!apiRes) return '';
  if (apiRes.output_text) return String(apiRes.output_text || '');

  const out = [];
  const seen = {};
  function pushText_(value) {
    const text = String(value || '').trim();
    if (!text) return;
    // P545: Responses API output_text items can expose both `type=output_text` and `text`.
    // Do not append the same JSON body twice, because `{"..."}\n{"..."}` is invalid JSON.
    if (seen[text]) return;
    seen[text] = true;
    out.push(text);
  }

  (apiRes.output || []).forEach(function(item) {
    (item.content || []).forEach(function(c) {
      if (!c) return;
      if (c.type === 'output_text' && c.text) {
        pushText_(c.text);
        return;
      }
      if (c.text) {
        pushText_(c.text);
        return;
      }
      if (typeof c === 'string') {
        pushText_(c);
      }
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
  const ambiguityFlags = uniqueMyCustomerAiStringsP547_([].concat(candidate.ambiguityFlags || [], ai.ambiguityFlags || []));
  const evidence = Array.isArray(ai.evidence) ? ai.evidence : [];

  if (ai.decision !== 'STATUS_CHANGE') {
    serverAllowed = false;
    reasons.push('AI 판단이 ' + (ai.decision || '미지정') + '이므로 상태변경하지 않음');
  }
  if (!evidence.length) {
    serverAllowed = false;
    reasons.push('판단 근거 evidenceId 없음');
  }
  if (serverAllowed && !evidence.some(function(ev) { return !!String(ev && ev.date || '').trim(); })) {
    serverAllowed = false;
    reasons.push('상태변경을 지지하는 날짜 있는 근거 없음');
  }
  if (hasMyCustomerAiHardAmbiguityP547_(ambiguityFlags)) {
    serverAllowed = false;
    reasons.push('최신 이력의 의미/범위가 불명확함: ' + ambiguityFlags.join(','));
  }
  if (target === '수주실패' && ambiguityFlags.indexOf('THIRD_PARTY_COMPARISON_ONLY') >= 0) {
    serverAllowed = false;
    reasons.push('타업체 견적/가격 비교만으로 수주실패 전환 금지');
  }
  if (target === '발주완료') {
    serverAllowed = false;
    reasons.push('발주완료는 AI/자동추천 금지');
  }
  if (['수주실패','계약완료'].indexOf(currentStatus) >= 0 && target !== currentStatus) {
    serverAllowed = false;
    reasons.push('현재 보호 상태 변경 금지: ' + currentStatus);
  }
  if (currentStatus === '발주완료' && target !== '계약완료' && target !== currentStatus) {
    serverAllowed = false;
    reasons.push('발주완료는 계약완료만 추천 가능');
  }
  if (String(ai.falsePositiveRisk || '').toUpperCase() === 'HIGH') {
    serverAllowed = false;
    reasons.push('오탐위험 HIGH');
  }
  if (target && target !== currentStatus) {
    try {
      const detail = getCustomerDetail(Number(candidate.rowNo || 0) || candidate.rowNo);
      const liveContractInfo = getMyCustomerContractCompletionInfoP536_(detail || {}, getMyCustomerContractCompleteMapP536_());
      const contractInfo = liveContractInfo || candidate.contractCompletionInfo || {};
      const ok = isMyCustomerStatusAllowedTransitionP536_(currentStatus, target, {
        contractCompleteReady: !!(contractInfo && contractInfo.ready),
        strongFailure: target === '수주실패'
      });
      if (!ok) {
        serverAllowed = false;
        reasons.push(getMyCustomerStatusTransitionBlockReasonP536_(currentStatus, target, contractInfo));
      }
    } catch (err) {
      serverAllowed = false;
      reasons.push('서버 전이 검증 실패: ' + (err && err.message ? err.message : String(err || '')));
    }
  } else {
    serverAllowed = false;
    reasons.push('현재상태 유지 또는 추천상태 없음');
  }

  const ruleMismatch = getMyCustomerAiRuleMismatchP547_(candidate, ai);
  return {
    serverAllowed: serverAllowed,
    finalStatus: serverAllowed ? target : currentStatus,
    blockedReason: reasons.join(' / ') || String(ai.blockedReason || ''),
    targetStatus: target,
    ambiguityFlags: ambiguityFlags,
    ruleAiMismatch: ruleMismatch,
    serverDecision: serverAllowed ? 'ALLOW_STATUS_CHANGE' : (ai.decision === 'HUMAN_REVIEW' ? 'HUMAN_REVIEW' : 'KEEP_CURRENT')
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
    'evidenceJson': safeStringifyMyCustomerAiP537_(ai.evidence || [], 2200),
    'evidenceIdsJson': safeStringifyMyCustomerAiP537_(ai.evidenceIds || [], 500),
    'ambiguityFlagsJson': safeStringifyMyCustomerAiP537_(validated.ambiguityFlags || ai.ambiguityFlags || [], 800),
    'ruleAiMismatch': validated.ruleAiMismatch || '',
    'serverDecision': validated.serverDecision || '',
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
    const firstObject = extractFirstJsonObjectMyCustomerAiP545_(src);
    if (firstObject) {
      try { parsed = JSON.parse(firstObject); } catch (err2) {}
    }

    // Last-resort fallback for old logs or non-standard wrappers.
    if (!parsed) {
      const start = src.indexOf('{');
      const end = src.lastIndexOf('}');
      if (start >= 0 && end > start) {
        try { parsed = JSON.parse(src.slice(start, end + 1)); } catch (err3) {}
      }
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

function extractFirstJsonObjectMyCustomerAiP545_(src) {
  src = String(src || '');
  const start = src.indexOf('{');
  if (start < 0) return '';

  let depth = 0;
  let inString = false;
  let escapeNext = false;

  for (let i = start; i < src.length; i++) {
    const ch = src.charAt(i);

    if (inString) {
      if (escapeNext) {
        escapeNext = false;
      } else if (ch === '\\') {
        escapeNext = true;
      } else if (ch === '"') {
        inString = false;
      }
      continue;
    }

    if (ch === '"') {
      inString = true;
      continue;
    }
    if (ch === '{') depth++;
    if (ch === '}') {
      depth--;
      if (depth === 0) {
        return src.slice(start, i + 1);
      }
    }
  }
  return '';
}

function upgradeMyCustomerAiCandidateP547_(candidate) {
  candidate = candidate || {};
  const alreadyPrepared = Array.isArray(candidate.events) && candidate.events.length && candidate.events.every(function(ev) { return !!(ev && ev.evidenceId); });
  if (alreadyPrepared) {
    candidate.ambiguityFlags = uniqueMyCustomerAiStringsP547_(candidate.ambiguityFlags || []);
    return candidate;
  }
  const pack = buildMyCustomerAiEvidencePackP547_(candidate.events || [], [
    Object.assign({}, candidate.latestDecisionEvent || {}, { declaredLatest: true }),
    Object.assign({}, candidate.latestAnyEvent || {}, { declaredLatest: true })
  ]);
  candidate.events = pack.events.slice(0, 8);
  candidate.latestDecisionEvent = pack.latestDecisionEvent || {};
  candidate.latestAnyEvent = pack.latestAnyEvent || {};
  candidate.ambiguityFlags = pack.ambiguityFlags || [];
  candidate.dateCoverage = pack.dateCoverage || {};
  return candidate;
}

function buildMyCustomerAiEvidencePackP547_(events, fallbackEvents) {
  const combined = [].concat(Array.isArray(events) ? events : [], Array.isArray(fallbackEvents) ? fallbackEvents : []);
  const dedupedByKey = {};
  combined.forEach(function(ev, idx) {
    const normalized = normalizeMyCustomerAiEvidenceEventP547_(ev, idx);
    if (!normalized || !normalized.text || isMyCustomerAiSystemMetaTextP544_(normalized.text)) return;
    const key = normalizeMyCustomerAiEventTextKeyP547_(normalized.text);
    if (!key) return;
    const existing = dedupedByKey[key];
    if (!existing) {
      dedupedByKey[key] = normalized;
    } else if (isMyCustomerAiEvidenceEventBetterP547_(normalized, existing)) {
      normalized.__declaredLatest = !!(normalized.__declaredLatest || existing.__declaredLatest);
      dedupedByKey[key] = normalized;
    } else {
      existing.__declaredLatest = !!(existing.__declaredLatest || normalized.__declaredLatest);
    }
  });

  const normalizedEvents = Object.keys(dedupedByKey).map(function(k) { return dedupedByKey[k]; });
  normalizedEvents.sort(function(a, b) {
    const ad = getMyCustomerAiEventDateMsP547_(a);
    const bd = getMyCustomerAiEventDateMsP547_(b);
    if (!!a.__declaredLatest !== !!b.__declaredLatest && (!ad || !bd)) return a.__declaredLatest ? -1 : 1;
    if (ad !== bd) return bd - ad;
    if ((b.__sourcePriority || 0) !== (a.__sourcePriority || 0)) return (b.__sourcePriority || 0) - (a.__sourcePriority || 0);
    return (b.__inputOrder || 0) - (a.__inputOrder || 0);
  });

  const outEvents = normalizedEvents.slice(0, 12).map(function(ev, idx) {
    return {
      evidenceId: 'E' + String(idx + 1).padStart(2, '0'),
      source: ev.source || '',
      sourceLabel: ev.sourceLabel || ev.source || '',
      date: ev.date || '',
      text: ev.text || '',
      actor: ev.actor || ''
    };
  });
  const latestAny = outEvents.length ? outEvents[0] : {};
  const latestDecision = outEvents.filter(isMyCustomerAiDecisionEvidenceP547_)[0] || latestAny || {};
  const dated = outEvents.filter(function(ev) { return !!String(ev.date || '').trim(); }).length;
  const ambiguityFlags = getMyCustomerAiDeterministicFlagsP547_(outEvents, latestAny, latestDecision);
  return {
    events: outEvents,
    latestAnyEvent: latestAny,
    latestDecisionEvent: latestDecision,
    ambiguityFlags: ambiguityFlags,
    dateCoverage: {
      total: outEvents.length,
      dated: dated,
      undated: Math.max(0, outEvents.length - dated),
      latestDate: latestAny && latestAny.date || '',
      latestEventHasDate: !!(latestAny && latestAny.date)
    }
  };
}

function normalizeMyCustomerAiEvidenceEventP547_(ev, inputOrder) {
  ev = ev || {};
  const text = shortenMyCustomerStatusTextP528_(ev.text || ev.summary || ev.rawText || '', 260);
  if (!text) return null;
  const rawDate = ev.dateText || ev.date || '';
  const parsedDate = coerceMyCustomerStatusDateP535_(rawDate) || parseLoosePortalDateP528_(text);
  const dateText = parsedDate ? formatMyCustomerStatusDateP528_(parsedDate) : '';
  const source = String(ev.source || '').trim();
  const sourceLabel = String(ev.sourceLabel || ev.source || '').trim();
  return {
    source: source,
    sourceLabel: sourceLabel,
    date: dateText,
    text: text,
    actor: String(ev.actor || '').trim(),
    __sourcePriority: getMyCustomerAiSourcePriorityP547_(source, sourceLabel),
    __inputOrder: Number(inputOrder || 0) || 0,
    __declaredLatest: !!ev.declaredLatest
  };
}

function getMyCustomerAiSourcePriorityP547_(source, sourceLabel) {
  const text = [source, sourceLabel].join(' ').toLowerCase();
  if (/컨택이력|contact/.test(text)) return 50;
  if (/tm/.test(text)) return 45;
  if (/마스터|memo/.test(text)) return 30;
  if (/자료발송|send/.test(text)) return 20;
  return 10;
}

function normalizeMyCustomerAiEventTextKeyP547_(text) {
  text = String(text || '').toLowerCase().trim();
  text = text.replace(/^\s*(?:\[[^\]]+\]\s*)+/, '');
  text = text.replace(/^\s*(?:20)?\d{2}\s*[.\-/]\s*\d{1,2}\s*[.\-/]\s*\d{1,2}\s*[.]?\s*/, '');
  text = text.replace(/\([^()]{0,20}\)\s*$/, '');
  return text.replace(/[\s.,;:!?()[\]{}<>…'\"“”‘’·~`@#$%^&*+=|\\/_\-]+/g, '');
}

function isMyCustomerAiEvidenceEventBetterP547_(a, b) {
  const ad = getMyCustomerAiEventDateMsP547_(a);
  const bd = getMyCustomerAiEventDateMsP547_(b);
  if (!!ad !== !!bd) return !!ad;
  if (ad !== bd) return ad > bd;
  return (a.__sourcePriority || 0) > (b.__sourcePriority || 0);
}

function getMyCustomerAiEventDateMsP547_(ev) {
  const d = coerceMyCustomerStatusDateP535_(ev && ev.date);
  return d ? d.getTime() : 0;
}

function isMyCustomerAiDecisionEvidenceP547_(ev) {
  const text = String(ev && ev.text || '');
  if (!text) return false;
  if (hasMyCustomerAiMaterialSendTextP544_(text) && !/(검토|결정|진행|계약|선정|거절|자체|관리업체|타사|타업체|기존업체|내년|하반기|예산|전화|통화)/.test(text)) return false;
  return /(검토|결정|진행|계약|선정|거절|안\s*한다고|필요\s*없|자체|관리업체|타사|타업체|다른\s*업체|기존\s*업체|내년|하반기|예산|전화|통화|부재|연락)/.test(text);
}

function getMyCustomerAiDeterministicFlagsP547_(events, latestAny, latestDecision) {
  const flags = [];
  const latestText = [latestAny && latestAny.text || '', latestDecision && latestDecision.text || ''].join('\n');
  const allText = (events || []).map(function(ev) { return ev.text || ''; }).join('\n');
  if (!(events || []).some(function(ev) { return !!String(ev.date || '').trim(); }) || (latestAny && !String(latestAny.date || '').trim())) flags.push('DATE_MISSING');
  if (/(자체적으로\s*진행|자체\s*진행)/.test(latestText) && !/(관리자|선임|유지보수|성능점검|공사|점검)/.test(latestText)) {
    flags.push('SELF_HANDLED_AMBIGUOUS');
    flags.push('SCOPE_UNCLEAR');
  }
  if (/(관리\s*업체\s*(있음|있다|있다고)|기존\s*업체\s*(있음|있다))/.test(latestText) && !/(유지|계약|선정|결정|진행|교체|변경|계속)/.test(latestText)) {
    flags.push('EXISTING_VENDOR_AMBIGUOUS');
    flags.push('SCOPE_UNCLEAR');
  }
  if (/(전화|통화).{0,12}(끊|종료)|바로\s*끊/.test(latestText)) flags.push('CALL_ENDED_AMBIGUOUS');
  const hasComparison = /(타사|타업체|다른\s*업체).{0,20}(견적|가격|금액|비교|네고)|(견적|가격|금액).{0,20}(타사|타업체|다른\s*업체)/.test(allText);
  const hasSelection = isMyCustomerAiThirdPartySelectionTextP547_(allText);
  if (hasComparison && !hasSelection) flags.push('THIRD_PARTY_COMPARISON_ONLY');
  if (hasSelection) {
    const selectedIndex = (events || []).findIndex(function(ev) { return isMyCustomerAiThirdPartySelectionTextP547_(ev.text || ''); });
    const newerActive = (events || []).slice(0, Math.max(0, selectedIndex)).some(function(ev) {
      return /(재견적|견적.{0,8}(요청|재발송)|검토\s*중|네고|가격\s*조율|진행\s*문의)/.test(ev.text || '');
    });
    if (newerActive) flags.push('CONFLICTING_EVENTS');
  }
  return uniqueMyCustomerAiStringsP547_(flags);
}

function isMyCustomerAiThirdPartySelectionTextP547_(text) {
  text = String(text || '');
  const uncertain = /(언제\s*(결정|선정)|(결정|선정).{0,10}(모르|못|미정|안\s*됨|전)|아직.{0,14}(결정|선정).{0,8}(못|안)|검토\s*중|비교\s*중)/.test(text);
  if (uncertain) return false;
  return /(타사|타업체|다른\s*업체|다른\s*곳).{0,16}(선정\s*(완료|함|했다|됨)|결정\s*(완료|함|했다|됨)|계약\s*(완료|체결|함|했다|진행)|진행하기로)|(기존\s*업체).{0,12}(유지|계속|계약)/.test(text);
}

function normalizeMyCustomerAiDecisionP547_(candidate, ai) {
  candidate = candidate || {};
  ai = ai || {};
  const allowedDecisions = ['STATUS_CHANGE','INSIGHT_ONLY','KEEP_CURRENT','HUMAN_REVIEW'];
  const allowedStatuses = ['고객 설득 중','견적제출완료','장기 추진건','수주실패','계약완료','현재상태유지'];
  const allowedRisks = ['LOW','MEDIUM','HIGH'];
  const allowedInsights = ['자료발송후미후속','연락장애','장기재접촉','타사비교','타사선정','계약발주확인필요','데이터확인필요','진행중추적','없음'];
  const evidenceById = {};
  (candidate.events || []).forEach(function(ev) { if (ev && ev.evidenceId) evidenceById[ev.evidenceId] = ev; });

  let decision = allowedDecisions.indexOf(String(ai.decision || '')) >= 0 ? String(ai.decision) : 'HUMAN_REVIEW';
  let recommendedStatus = allowedStatuses.indexOf(String(ai.recommendedStatus || '')) >= 0 ? String(ai.recommendedStatus) : '현재상태유지';
  let recommendationAllowed = !!ai.recommendationAllowed;
  let confidence = Math.max(0, Math.min(100, Number(ai.confidence || 0) || 0));
  let falsePositiveRisk = allowedRisks.indexOf(String(ai.falsePositiveRisk || '').toUpperCase()) >= 0 ? String(ai.falsePositiveRisk).toUpperCase() : 'HIGH';
  let insightType = allowedInsights.indexOf(String(ai.insightType || '')) >= 0 ? String(ai.insightType) : '데이터확인필요';
  let blockedReason = String(ai.blockedReason || '').trim();
  const modelEvidenceIds = uniqueMyCustomerAiStringsP547_(Array.isArray(ai.evidenceIds) ? ai.evidenceIds : []);
  const evidenceIds = modelEvidenceIds.filter(function(id) { return !!evidenceById[id]; }).slice(0, 4);
  let ambiguityFlags = uniqueMyCustomerAiStringsP547_([].concat(candidate.ambiguityFlags || [], Array.isArray(ai.ambiguityFlags) ? ai.ambiguityFlags : []));
  const guardReasons = [];

  if (evidenceIds.length !== modelEvidenceIds.length) {
    ambiguityFlags.push('EVIDENCE_WEAK');
    guardReasons.push('입력에 없는 evidenceId 제거');
  }
  if (!evidenceIds.length) {
    ambiguityFlags.push('EVIDENCE_WEAK');
    decision = 'HUMAN_REVIEW';
    recommendationAllowed = false;
    recommendedStatus = '현재상태유지';
    falsePositiveRisk = 'HIGH';
    confidence = Math.min(confidence, 55);
    guardReasons.push('유효한 판단 근거 없음');
  }
  const evidence = evidenceIds.map(function(id) { return evidenceById[id]; });
  if (decision === 'STATUS_CHANGE' && !evidence.some(function(ev) { return !!String(ev && ev.date || '').trim(); })) {
    ambiguityFlags.push('DATE_MISSING');
    decision = 'HUMAN_REVIEW';
    recommendationAllowed = false;
    recommendedStatus = '현재상태유지';
    falsePositiveRisk = 'HIGH';
    confidence = Math.min(confidence, 55);
    guardReasons.push('날짜가 확인된 상태변경 근거 없음');
  }
  if (decision === 'STATUS_CHANGE' && !recommendationAllowed) {
    decision = 'HUMAN_REVIEW';
    recommendedStatus = '현재상태유지';
    falsePositiveRisk = 'HIGH';
    confidence = Math.min(confidence, 55);
    guardReasons.push('STATUS_CHANGE와 recommendationAllowed 값 모순');
  }
  if (decision !== 'STATUS_CHANGE') {
    recommendationAllowed = false;
    recommendedStatus = '현재상태유지';
  }
  if (hasMyCustomerAiHardAmbiguityP547_(ambiguityFlags) && decision === 'STATUS_CHANGE') {
    decision = 'HUMAN_REVIEW';
    recommendationAllowed = false;
    recommendedStatus = '현재상태유지';
    falsePositiveRisk = 'HIGH';
    confidence = Math.min(confidence, 55);
    guardReasons.push('최신 이력 의미 또는 범위 불명확');
  }
  if (recommendedStatus === '수주실패' && ambiguityFlags.indexOf('THIRD_PARTY_COMPARISON_ONLY') >= 0) {
    decision = 'HUMAN_REVIEW';
    recommendationAllowed = false;
    recommendedStatus = '현재상태유지';
    falsePositiveRisk = 'HIGH';
    confidence = Math.min(confidence, 55);
    insightType = '타사비교';
    guardReasons.push('타업체 견적/가격 비교는 타사 선정이 아님');
  }
  if (['수주실패','계약완료'].indexOf(String(candidate.currentStatus || '').trim()) >= 0) {
    decision = 'KEEP_CURRENT';
    recommendationAllowed = false;
    recommendedStatus = '현재상태유지';
    guardReasons.push('현재상태가 보호 상태임');
  }
  if (String(candidate.currentStatus || '').trim() === '발주완료') {
    if (!(recommendedStatus === '계약완료' && candidate.contractCompletionInfo && candidate.contractCompletionInfo.ready)) {
      decision = 'KEEP_CURRENT';
      recommendationAllowed = false;
      recommendedStatus = '현재상태유지';
      guardReasons.push('발주완료→계약완료 서류 조건 미충족 또는 다른 상태 추천');
    }
  }
  insightType = normalizeMyCustomerAiInsightTypeP547_(candidate, insightType, recommendedStatus, ambiguityFlags);
  ambiguityFlags = uniqueMyCustomerAiStringsP547_(ambiguityFlags);
  if (guardReasons.length) blockedReason = [blockedReason, guardReasons.join(' / ')].filter(Boolean).join(' / ');

  return {
    decision: decision,
    recommendedStatus: recommendedStatus,
    recommendationAllowed: recommendationAllowed,
    confidence: confidence,
    falsePositiveRisk: falsePositiveRisk,
    insightType: insightType,
    summary: shortenMyCustomerStatusTextP528_(ai.summary || '', 260),
    reason: shortenMyCustomerStatusTextP528_(ai.reason || '', 360),
    blockedReason: shortenMyCustomerStatusTextP528_(blockedReason, 360),
    nextAction: shortenMyCustomerStatusTextP528_(ai.nextAction || '', 260),
    evidenceIds: evidenceIds,
    evidence: evidence,
    ambiguityFlags: ambiguityFlags
  };
}

function normalizeMyCustomerAiInsightTypeP547_(candidate, insightType, recommendedStatus, ambiguityFlags) {
  const allText = (candidate.events || []).map(function(ev) { return ev.text || ''; }).join('\n');
  const hasContactBlocker = /(부재|전화\s*연결\s*불가|전화\s*안\s*받|연락\s*안\s*됨|자리\s*비움|담당자\s*부재|안내\s*불가)/.test(allText);
  const hasContractSignal = /(계약서|용역신청서|사업자등록증|발주|계약\s*진행|계약완료)/.test(allText);
  const hasMaterialSend = hasMyCustomerAiMaterialSendTextP544_(allText);
  if (recommendedStatus === '계약완료') return '계약발주확인필요';
  if (recommendedStatus === '수주실패' && isMyCustomerAiThirdPartySelectionTextP547_(allText)) return '타사선정';
  if ((ambiguityFlags || []).indexOf('THIRD_PARTY_COMPARISON_ONLY') >= 0) return '타사비교';
  if ((ambiguityFlags || []).some(function(f) { return ['SELF_HANDLED_AMBIGUOUS','EXISTING_VENDOR_AMBIGUOUS','SCOPE_UNCLEAR'].indexOf(f) >= 0; })) return '데이터확인필요';
  if (insightType === '연락장애' && !hasContactBlocker) return hasMaterialSend ? '자료발송후미후속' : '진행중추적';
  if (insightType === '계약발주확인필요' && !hasContractSignal) return hasMaterialSend ? '자료발송후미후속' : '진행중추적';
  return insightType || '없음';
}

function hasMyCustomerAiHardAmbiguityP547_(flags) {
  return (flags || []).some(function(f) {
    return ['DATE_MISSING','CONFLICTING_EVENTS','SELF_HANDLED_AMBIGUOUS','EXISTING_VENDOR_AMBIGUOUS','CALL_ENDED_AMBIGUOUS','SCOPE_UNCLEAR','EVIDENCE_WEAK'].indexOf(f) >= 0;
  });
}

function getMyCustomerAiRuleMismatchP547_(candidate, ai) {
  const ruleTarget = String(candidate.ruleRecommendedStatus || candidate.currentStatus || '').trim();
  const aiTarget = ai.decision === 'STATUS_CHANGE' ? String(ai.recommendedStatus || '').trim() : String(candidate.currentStatus || '').trim();
  if (!!candidate.ruleStateChangeRecommended !== (ai.decision === 'STATUS_CHANGE')) {
    return candidate.ruleStateChangeRecommended ? '규칙=변경추천 / AI=보류' : '규칙=유지 / AI=변경추천';
  }
  if (candidate.ruleStateChangeRecommended && ruleTarget !== aiTarget) return '추천상태 불일치: 규칙=' + ruleTarget + ' / AI=' + aiTarget;
  return '';
}

function uniqueMyCustomerAiStringsP547_(values) {
  const out = [];
  const seen = {};
  (values || []).forEach(function(v) {
    const key = String(v || '').trim();
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push(key);
  });
  return out;
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
  return /(기존\s*담당자별\s*시트|마스터시트로\s*이관|이관|중복\s*삭제|데이터\s*병합|정보\s*확인해보시고|TM\s*콜\s*원하시면|삭제\s*고객번호|대표전화번호|담당자\s*이메일\s*주소|연면적)/i.test(text);
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
