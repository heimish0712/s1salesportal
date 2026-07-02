/***************************************
 * S1 Sales Portal - 01_Schema.gs
 * PATCH Q: 마스터시트 헤더 매핑 중앙화
 * 기준일: 2026-06-21
 ***************************************/

/**
 * PATCH Q 원칙
 * - 마스터시트 헤더 후보는 이 파일의 PORTAL_MASTER_FIELD_SCHEMA에서만 관리합니다.
 * - 각 서비스/화면은 되도록 key(customerNo, company, grade 등)만 사용합니다.
 * - 실제 헤더명 차이(공백, 줄바꿈, 일부 후보명)는 04_MasterRepository.gs의 normalized header resolver가 처리합니다.
 */
const PORTAL_MASTER_FIELD_SCHEMA = {
  customerNo: { label: '고객번호', headers: ['고객번호', '고객 번호'], type: 'text' },
  orderNo: { label: '발주번호', headers: ['발주번호', '계약번호', '계약 번호'], type: 'text' },
  company: { label: '회사명', headers: ['회사명', '건물명', '건물명 (회사명)', '고객사명'], type: 'text' },
  salesRep: { label: '영업담당자', headers: ['영업담당자', '견적담당', '담당영업'], type: 'text' },
  firstRegisteredAt: { label: '마스터시트 최초등록일', headers: ['마스터시트\n최초등록일', '마스터시트 최초등록일', '최초등록일', '등록일'], type: 'text' },
  region: { label: '지역구분', headers: ['지역구분', '지역'], type: 'text' },
  vendor: { label: '수행사', headers: ['수행사', '최종수행사', '최종 수행사'], type: 'select', options: ['', 'KJ', '일신', '삼구'] },
  grade: { label: '관리등급', headers: ['관리등급', '관리 등급', '등급', '연면적유형', '연면적 유형'], type: 'text' },
  status: { label: '현재 영업 진행 상황', headers: ['현재 영업 진행 상황', '계약진행상황', '계약진행상태', '영업상태', '상태'], type: 'select', optionsSource: 'statusOptions' },
  customerRank: { label: '고객등급', headers: ['고객등급', '고객 등급', '영업등급', '고객관리등급', '고객 관리 등급'], type: 'select', optionsSource: 'customerRankOptions' },

  contact: { label: '고객사 담당자', headers: ['고객사 담당자', '담당자', '담당자명'], type: 'text' },
  phone: { label: '대표전화번호', headers: ['대표전화번호', '전화번호', '대표 전화번호'], type: 'text' },
  directPhone: { label: '직통번호', headers: ['직통번호', '직통번호or휴대폰번호', '직통번호/휴대폰번호', '직통번호 or 휴대폰번호', '휴대폰번호'], type: 'text' },
  email: { label: '담당자 이메일 주소', headers: ['담당자 이메일 주소', '담당자 이메일', '이메일주소', '이메일 주소', '이메일'], type: 'email' },
  address: { label: '고객사 상세 주소', headers: ['고객사 상세 주소', '도로명주소', '지번주소', '주소', '상세주소'], type: 'text', wide: true },
  memo: { label: '마스터시트 메모', headers: ['메모', '비고', '영업메모', '마스터시트 메모'], type: 'textarea', wide: true },
  longNoContactTransferred: { label: '장기미접촉 이관 여부', headers: ['장기미접촉 이관 여부', '장기미접촉이관여부', '장기 미접촉 이관 여부'], type: 'text' },
  tmProgressStatus: { label: 'TM 진행 현황', headers: ['TM 진행 현황', 'TM진행현황', 'TM 진행현황', 'TM진행 현황'], type: 'text' },
  tmContactContent: { label: 'TM 컨택 내용', headers: ['TM 컨택 내용', 'TM컨택내용', 'TM 컨택내용', 'TM컨택 내용'], type: 'textarea', wide: true },

  area: { label: '연면적', headers: ['연면적', '연면적(㎡)', '연면적㎡', '면적'], type: 'numberText' },
  buildingType: { label: '건물 유형', headers: ['건물 유형', '건물유형'], type: 'select', options: ['', '기업', '공동주택', '학교', '공공', '기타'] },
  finalQuote: { label: '최종 견적가', headers: ['최종 견적가', '최종견적가', '최종단가', '최종 견적금액', '견적금액'], type: 'money' },
  contractUnit: { label: '계약단위', headers: ['계약단위', '계약 단위'], type: 'select', options: buildPortalContractUnitOptionsP280_() },
  contractStartDate: { label: '계약시작일', headers: ['계약시작일', '계약 시작일'], type: 'dateText' },
  contractEndDate: { label: '계약종료일', headers: ['계약종료일', '계약 종료일'], type: 'dateText' },
  appointment: { label: '관리자 선임 여부', headers: ['관리자\n선임 여부', '관리자 선임 여부', '관리자선임여부', '선임 여부', '관리자선임'], type: 'select', options: ['', '선임', '비선임', '해당없음'] },
  maintenance: { label: '유지점검', headers: ['유지점검', '유지 점검'], type: 'select', options: buildPortalInspectionCountOptionsP280_(12) },
  performance: { label: '성능점검', headers: ['성능점검', '성능 점검'], type: 'select', options: buildPortalInspectionCountOptionsP280_(12) },
  vat: { label: '부가세', headers: ['부가세', 'VAT', '부가세 여부'], type: 'select', options: ['', '별도', '포함'] },
  discountRate: { label: '할인율', headers: ['할인율(%)', '할인률(%)', '할인율', '할인률', '할인율 %'], type: 'numberText' },
  specialTerms: { label: '용역신청서특약사항', headers: ['용역신청서특약사항', '용역신청서 특약사항', '특약사항', '계약기간\n(상세 작성 요구한 경우만)'], type: 'textarea', wide: true },
  s1Referrer: { label: '에스원제보자', headers: ['제보자', '에스원제보자', '에스원 제보자', 'S1제보자', 'S1 제보자'], type: 'text' },

  lastSent: { label: '마지막발송', headers: ['마지막발송', '마지막 발송', '최근발송', '최근 발송'], type: 'text' },
  sentAt: { label: '발송일시', headers: ['발송일시', '발송 일시', '최근발송일시', '최근 발송일시'], type: 'text' },

  businessNo: { label: '사업자등록번호', headers: ['사업자등록번호', '사업자 등록 번호', '사업자번호', '등록번호'], type: 'text' },
  businessLegalName: { label: '법인명', headers: ['사업자등록증상 법인명', '계약 당사자(사업자등록증상 법인명)', '계약 당사자', '법인명', '상호', '업체명'], type: 'text' },
  representativeName: { label: '대표자명', headers: ['대표자명', '대표자', '대표', '대표자 성명'], type: 'text' },
  businessAddress: { label: '사업자등록증상 법인 주소', headers: ['사업자등록증상 법인 주소', '사업자등록증상 법인주소', '사업자등록증 주소', '사업자등록증상 주소', '법인 주소', '법인주소'], type: 'text', wide: true },
  businessRegistrationReceived: { label: '사업자등록증', headers: ['사업자등록증', '사업자등록증 수취 여부', '사업자등록증 수취여부', '사업자등록증수취여부', '사업자등록증요청', '사업자등록증 요청'], type: 'booleanMark' },
  serviceApplicationReceived: { label: '용역신청서', headers: ['용역신청서', '용역신청서 수취 여부', '용역신청서 수취여부', '용역신청서수취여부', '용역신청서요청', '용역신청서 요청'], type: 'booleanMark' },
  appointmentReportReceived: { label: '선임신고서', headers: ['선임신고서', '선임신고서 수취 여부', '선임신고서 수취여부', '선임신고서수취여부', '선임신고서 및 위임장 요청', '선임신고서 및 위임장', '선임신고서·위임장'], type: 'booleanMark' }
};


function buildPortalContractUnitOptionsP280_() {
  const opts = [{ value: '', label: '-' }];
  for (let i = 1; i <= 12; i++) opts.push({ value: String(i), label: i + '개월' });
  return opts;
}

function buildPortalInspectionCountOptionsP280_(maxCount) {
  maxCount = Number(maxCount) || 12;
  const opts = [{ value: '', label: '-' }];
  for (let i = 0; i <= maxCount; i++) opts.push({ value: String(i), label: i + '회' });
  return opts;
}

function getMasterFieldSchema_() {
  return PORTAL_MASTER_FIELD_SCHEMA;
}

function getMasterFieldDef_(key) {
  key = String(key || '').trim();
  return PORTAL_MASTER_FIELD_SCHEMA[key] || null;
}

function masterFieldHeaders_(key, fallback) {
  const def = getMasterFieldDef_(key);
  if (def && def.headers && def.headers.length) return def.headers.slice();
  return Array.isArray(fallback) ? fallback.slice() : (fallback ? [fallback] : []);
}

function masterFieldLabel_(key, fallback) {
  const def = getMasterFieldDef_(key);
  return (def && def.label) || fallback || key;
}

function portalFieldDef_(key, overrides) {
  const base = getMasterFieldDef_(key) || { label: key, headers: [key], type: 'text' };
  const def = Object.assign({}, base, overrides || {});
  def.key = key;
  def.headers = masterFieldHeaders_(key, def.headers || []);
  def.label = def.label || masterFieldLabel_(key, key);
  if (!def.type) def.type = base.type || 'text';
  return def;
}

function getMasterFieldValue_(obj, key) {
  const headers = masterFieldHeaders_(key);
  if (!headers.length) return '';
  return getValueByHeaderCandidates_(obj || {}, headers);
}

function findMasterFieldCol_(headerMap, key) {
  return findFirstExistingHeaderCol_(headerMap, masterFieldHeaders_(key));
}

function ensureMasterFieldColumn_(sheet, headerMap, key) {
  const headers = masterFieldHeaders_(key);
  return ensureMasterColumn_(sheet, headerMap, headers[0] || masterFieldLabel_(key, key));
}

function getMasterFieldSchemaForClient() {
  const result = [];
  Object.keys(PORTAL_MASTER_FIELD_SCHEMA).forEach(function(key) {
    const def = PORTAL_MASTER_FIELD_SCHEMA[key] || {};
    result.push({ key: key, label: def.label || key, headers: (def.headers || []).slice(), type: def.type || 'text' });
  });
  return result;
}

const PORTAL_SEND_FILE_DEFINITIONS = [
  { key: 'quote', label: '견적서요청', headers: ['견적서요청'] },
  { key: 'serviceApplication', label: '용역신청서요청', headers: ['용역신청서요청'] },
  { key: 'appointmentDoc', label: '선임신고서 및 위임장 요청', headers: ['선임신고서 및 위임장 요청'] },
  { key: 'termsGuide', label: '약관, 안내문, 법령요약', headers: ['약관, 안내문, 법령요약', '안내문, 법령요약', '안내문, 법령', '약관/안내문/법령요약'] },
  { key: 'contractorInfo', label: '수행사 정보', headers: ['수행사정보', '수행사 정보'] },
  { key: 'sampleReport', label: '샘플보고서', headers: ['샘플보고서'] },
  { key: 'compareQuote', label: '비교견적서', headers: ['비교견적서'] },
  { key: 'serviceStandardContract', label: '수행사 용역표준계약서', headers: ['수행사 용역표준계약서', '수행사 용역표준계약', '용역표준계약서', '표준계약서'] }
];

const PORTAL_DETAIL_FIELDS = {
  basic: [
    portalFieldDef_('customerNo', { editable: false }),
    portalFieldDef_('company', { editable: true }),
    portalFieldDef_('salesRep', { editable: true }),
    portalFieldDef_('firstRegisteredAt', { editable: false }),
    portalFieldDef_('region', { editable: true }),
    portalFieldDef_('vendor', { editable: true }),
    portalFieldDef_('grade', { editable: false }),
    portalFieldDef_('status', { editable: true, optionsSource: 'statusOptions' }),
    portalFieldDef_('customerRank', { editable: true, optionsSource: 'customerRankOptions' }),
    portalFieldDef_('contact', { editable: true }),
    portalFieldDef_('phone', { editable: true }),
    portalFieldDef_('directPhone', { editable: true }),
    portalFieldDef_('email', { editable: true }),
    portalFieldDef_('address', { editable: true, wide: true }),
    portalFieldDef_('memo', { editable: true, wide: true })
  ],
  contract: [
    portalFieldDef_('area', { editable: true }),
    portalFieldDef_('buildingType', { editable: true }),
    portalFieldDef_('finalQuote', { editable: false }),
    portalFieldDef_('contractUnit', { editable: true }),
    portalFieldDef_('contractStartDate', { editable: true }),
    portalFieldDef_('contractEndDate', { editable: true }),
    portalFieldDef_('s1Referrer', { editable: true }),
    portalFieldDef_('appointment', { editable: true }),
    portalFieldDef_('maintenance', { editable: true }),
    portalFieldDef_('performance', { editable: true }),
    portalFieldDef_('vat', { editable: true }),
    portalFieldDef_('discountRate', { editable: true }),
    portalFieldDef_('specialTerms', { editable: true, wide: true })
  ],
  extra: [
    portalFieldDef_('businessRegistrationReceived', { editable: true }),
    portalFieldDef_('serviceApplicationReceived', { editable: true }),
    portalFieldDef_('appointmentReportReceived', { editable: true }),
    portalFieldDef_('businessNo', { editable: true }),
    portalFieldDef_('businessLegalName', { editable: true }),
    portalFieldDef_('representativeName', { editable: true }),
    portalFieldDef_('businessAddress', { editable: true, wide: true })
  ]
};

const PORTAL_CUSTOMER_LIST_FIELD_HEADERS = {
  customerNo: masterFieldHeaders_('customerNo'),
  orderNo: masterFieldHeaders_('orderNo'),
  salesRep: masterFieldHeaders_('salesRep'),
  company: masterFieldHeaders_('company'),
  status: masterFieldHeaders_('status'),
  customerRank: masterFieldHeaders_('customerRank'),
  contact: masterFieldHeaders_('contact'),
  phone: masterFieldHeaders_('phone'),
  directPhone: masterFieldHeaders_('directPhone'),
  email: masterFieldHeaders_('email'),
  vendor: masterFieldHeaders_('vendor'),
  finalQuote: masterFieldHeaders_('finalQuote'),
  memo: masterFieldHeaders_('memo'),
  address: masterFieldHeaders_('address'),
  lastSent: masterFieldHeaders_('lastSent'),
  sentAt: masterFieldHeaders_('sentAt'),
  firstRegisteredAt: masterFieldHeaders_('firstRegisteredAt'),
  region: masterFieldHeaders_('region'),
  area: masterFieldHeaders_('area'),
  grade: masterFieldHeaders_('grade'),
  buildingType: masterFieldHeaders_('buildingType'),
  contractUnit: masterFieldHeaders_('contractUnit'),
  contractStartDate: masterFieldHeaders_('contractStartDate'),
  contractEndDate: masterFieldHeaders_('contractEndDate'),
  s1Referrer: masterFieldHeaders_('s1Referrer'),
  appointment: masterFieldHeaders_('appointment'),
  maintenance: masterFieldHeaders_('maintenance'),
  performance: masterFieldHeaders_('performance'),
  vat: masterFieldHeaders_('vat'),
  discountRate: masterFieldHeaders_('discountRate'),
  specialTerms: masterFieldHeaders_('specialTerms'),
  businessNo: masterFieldHeaders_('businessNo'),
  businessLegalName: masterFieldHeaders_('businessLegalName'),
  representativeName: masterFieldHeaders_('representativeName')
};

function getCustomerListValue_(obj, key) {
  obj = obj || {};
  key = String(key || '').trim();
  if (getMasterFieldDef_(key)) return getMasterFieldValue_(obj, key);
  const headers = PORTAL_CUSTOMER_LIST_FIELD_HEADERS && PORTAL_CUSTOMER_LIST_FIELD_HEADERS[key];
  if (headers && headers.length) return getValueByHeaderCandidates_(obj, headers);
  if (Object.prototype.hasOwnProperty.call(obj, key)) return String(obj[key] || '').trim();
  return '';
}

function buildStatusOptions_(currentStatus) {
  const options = (PORTAL_CONFIG.STATUS_OPTIONS || []).slice();
  const current = String(currentStatus || '').trim();
  if (current && options.indexOf(current) < 0) options.unshift(current);
  return options;
}

function buildCustomerRankOptions_(currentRank) {
  const options = (PORTAL_CONFIG.CUSTOMER_RANK_OPTIONS || ['', 'A급', 'B급', 'C급', '보류', '제외']).slice();
  const current = String(currentRank || '').trim();
  if (current && options.indexOf(current) < 0) options.unshift(current);
  return options;
}
