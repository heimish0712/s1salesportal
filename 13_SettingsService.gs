/***************************************
 * S1 Sales Portal - 13_SettingsService.gs
 * PATCH F: 설정_DB 기반 진행현황 옵션/색상 관리
 * 원칙: 상태값/칩색상은 코드 하드코딩이 아니라 설정_DB에서 운영자가 수정한다.
 ***************************************/

const PORTAL_SETTINGS_SHEET_NAME_F = '설정_DB';
const PORTAL_SETTINGS_HEADERS_F = ['설정구분', '값', '표시명', '배경색', '글자색', '순서', '사용여부', '설명'];
const PORTAL_SETTING_TYPE_STATUS_F = 'CONTRACT_STATUS';

const PORTAL_DEFAULT_STATUS_SETTINGS_F = [
  ['CONTRACT_STATUS', '수주실패', '수주실패', '#fecaca', '#991b1b', 10, 'Y', '수주 실패/실패 확정'],
  ['CONTRACT_STATUS', '견적제출완료', '견적제출완료', '#bae6fd', '#075985', 20, 'Y', '견적 제출 완료'],
  ['CONTRACT_STATUS', '장기 추진건', '장기 추진건', '#dcfce7', '#166534', 30, 'Y', '장기 추진 대상'],
  ['CONTRACT_STATUS', '고객 설득 중', '고객 설득 중', '#bbf7d0', '#166534', 40, 'Y', '고객 설득/협의 진행 중'],
  ['CONTRACT_STATUS', '발주완료', '발주완료', '#fef9c3', '#713f12', 50, 'Y', '발주 완료'],
  ['CONTRACT_STATUS', '계약완료', '계약완료', '#fff200', '#111827', 60, 'Y', '계약 완료'],
  ['CONTRACT_STATUS', '!!상태지정필요!!', '!!상태지정필요!!', '#e5e7eb', '#374151', 70, 'Y', '상태 보정 필요'],

  // 구버전/기타 상태값 호환용. 드롭다운에서는 숨기고 기존 데이터 칩 색상만 유지합니다.
  ['CONTRACT_STATUS', '영업종료', '영업종료', '#fecaca', '#991b1b', 900, 'N', '구버전 호환: 수주실패로 통합 권장'],
  ['CONTRACT_STATUS', '미선택', '미선택', '#f3f4f6', '#6b7280', 901, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '표준견적 발송요청', '표준견적 발송요청', '#bae6fd', '#075985', 902, 'N', '구버전 호환: 견적제출완료로 통합 권장'],
  ['CONTRACT_STATUS', '후속상담 완료', '후속상담 완료', '#bbf7d0', '#166534', 903, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '영업자 컨택 중', '영업자 컨택 중', '#bbf7d0', '#166534', 904, 'N', '구버전 호환: 고객 설득 중으로 통합 권장'],
  ['CONTRACT_STATUS', '계약서류 발송완료', '계약서류 발송완료', '#bae6fd', '#075985', 905, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '고객 내부검토중', '고객 내부검토중', '#dcfce7', '#166534', 906, 'N', '구버전 호환: 장기 추진건으로 통합 권장'],
  ['CONTRACT_STATUS', '수주확정', '수주확정', '#fff200', '#111827', 907, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '계약서 취합 완료', '계약서 취합 완료', '#fef9c3', '#713f12', 908, 'N', '구버전 호환: 발주완료로 통합 권장'],
  ['CONTRACT_STATUS', '영업팀 2차콜 필요', '영업팀 2차콜 필요', '#dbeafe', '#1d4ed8', 909, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '영업담당자 컨택 중', '영업담당자 컨택 중', '#bbf7d0', '#166534', 910, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '영업팀 수주 성공', '영업팀 수주 성공', '#fff200', '#111827', 911, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '영업종료(거절)', '영업종료(거절)', '#fecaca', '#991b1b', 912, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '계약중도취소', '계약중도취소', '#fecaca', '#991b1b', 913, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '미컨택', '미컨택', '#f3f4f6', '#4b5563', 914, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', 'TM 배분 완료', 'TM 배분 완료', '#fff34d', '#5f4b00', 915, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', 'TM 성공콜', 'TM 성공콜', '#fde047', '#713f12', 916, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '영업팀 2차콜 이후 영업 중', '영업팀 2차콜 이후 영업 중', '#bfdbfe', '#1e40af', 917, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '1차컨택대상', '1차컨택대상', '#f3f4f6', '#4b5563', 918, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '통화부재', '통화부재', '#fef3c7', '#92400e', 919, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '자료요청', '자료요청', '#dbeafe', '#1d4ed8', 920, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '견적발송', '견적발송', '#bae6fd', '#075985', 921, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '재컨택필요', '재컨택필요', '#bfdbfe', '#1e40af', 922, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '확인필요', '확인필요', '#e5e7eb', '#374151', 923, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '계약검토중', '계약검토중', '#e9d5ff', '#6b21a8', 924, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '계약서류발송', '계약서류발송', '#e9d5ff', '#6b21a8', 925, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '보류', '보류', '#f3f4f6', '#4b5563', 926, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '거절', '거절', '#fecaca', '#991b1b', 927, 'N', '구버전 호환'],
  ['CONTRACT_STATUS', '오류/제외', '오류/제외', '#e5e7eb', '#374151', 928, 'N', '구버전 호환']
];

function getPortalSettings() {
  const sheet = ensurePortalSettingsSheet_();
  const statusOptions = readPortalStatusSettings_(sheet);
  return {
    ok: true,
    source: PORTAL_SETTINGS_SHEET_NAME_F,
    updatedAt: new Date().toISOString(),
    statusOptions: statusOptions
  };
}

function setupPortalSettingsSheet() {
  const sheet = ensurePortalSettingsSheet_();
  SpreadsheetApp.flush();
  return {
    ok: true,
    sheetName: sheet.getName(),
    message: '설정_DB 시트를 준비했습니다.',
    statusOptions: readPortalStatusSettings_(sheet)
  };
}

function ensurePortalSettingsSheet_() {
  const ss = getWebAppDbSpreadsheet_();
  let sheet = ss.getSheetByName(PORTAL_SETTINGS_SHEET_NAME_F);
  if (!sheet) {
    sheet = ss.insertSheet(PORTAL_SETTINGS_SHEET_NAME_F);
    sheet.getRange(1, 1, 1, PORTAL_SETTINGS_HEADERS_F.length).setValues([PORTAL_SETTINGS_HEADERS_F]);
    sheet.setFrozenRows(1);
    seedPortalSettingsSheet_(sheet);
    return sheet;
  }

  const lastCol = Math.max(PORTAL_SETTINGS_HEADERS_F.length, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  let headerChanged = false;
  PORTAL_SETTINGS_HEADERS_F.forEach(function(h, i) {
    if (!String(headers[i] || '').trim()) {
      sheet.getRange(1, i + 1).setValue(h);
      headerChanged = true;
    }
  });

  if (sheet.getLastRow() < 2) seedPortalSettingsSheet_(sheet);
  syncPortalStatusSettingsToLatestDefaults_(sheet);
  if (headerChanged) SpreadsheetApp.flush();
  return sheet;
}

function seedPortalSettingsSheet_(sheet) {
  if (!sheet) return;
  if (sheet.getLastRow() >= 2) return;
  sheet.getRange(2, 1, PORTAL_DEFAULT_STATUS_SETTINGS_F.length, PORTAL_SETTINGS_HEADERS_F.length).setValues(PORTAL_DEFAULT_STATUS_SETTINGS_F);
  try {
    sheet.getRange(1, 1, 1, PORTAL_SETTINGS_HEADERS_F.length).setFontWeight('bold').setBackground('#e8eef7');
    sheet.autoResizeColumns(1, PORTAL_SETTINGS_HEADERS_F.length);
  } catch (err) {}
}


// PATCH P0-1K: 기존 설정_DB가 있어도 오늘 회의 기준 진행현황 목록으로 동기화합니다.
// CONTRACT_STATUS 구간만 최신 기본값으로 재작성하고, 다른 설정 구분은 보존합니다.
function syncPortalStatusSettingsToLatestDefaults_(sheet) {
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(PORTAL_SETTINGS_HEADERS_F.length, sheet.getLastColumn());
  let nonStatusRows = [];

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    nonStatusRows = values.filter(function(row) {
      return String(row[0] || '').trim() !== PORTAL_SETTING_TYPE_STATUS_F && row.some(function(v) { return String(v || '').trim(); });
    }).map(function(row) {
      const next = row.slice(0, PORTAL_SETTINGS_HEADERS_F.length);
      while (next.length < PORTAL_SETTINGS_HEADERS_F.length) next.push('');
      return next;
    });
  }

  const nextRows = PORTAL_DEFAULT_STATUS_SETTINGS_F.concat(nonStatusRows);
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
  if (nextRows.length) {
    sheet.getRange(2, 1, nextRows.length, PORTAL_SETTINGS_HEADERS_F.length).setValues(nextRows);
  }
  try {
    sheet.autoResizeColumns(1, PORTAL_SETTINGS_HEADERS_F.length);
  } catch (err) {}
}

function readPortalStatusSettings_(sheet) {
  sheet = sheet || ensurePortalSettingsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return mapPortalStatusSettingsRows_(PORTAL_DEFAULT_STATUS_SETTINGS_F);

  const values = sheet.getRange(2, 1, lastRow - 1, PORTAL_SETTINGS_HEADERS_F.length).getDisplayValues();
  const rows = values.filter(function(row) {
    return String(row[0] || '').trim() === PORTAL_SETTING_TYPE_STATUS_F && String(row[1] || row[2] || '').trim();
  });

  if (!rows.length) return mapPortalStatusSettingsRows_(PORTAL_DEFAULT_STATUS_SETTINGS_F);
  return mapPortalStatusSettingsRows_(rows);
}

function mapPortalStatusSettingsRows_(rows) {
  return (rows || []).map(function(row) {
    return {
      type: String(row[0] || '').trim(),
      value: String(row[1] || row[2] || '').trim(),
      label: String(row[2] || row[1] || '').trim(),
      bg: normalizePortalHexColor_(row[3], '#f2f4f7'),
      fg: normalizePortalHexColor_(row[4], '#344054'),
      order: Number(row[5] || 9999) || 9999,
      active: String(row[6] || 'Y').trim().toUpperCase() !== 'N',
      description: String(row[7] || '').trim()
    };
  }).sort(function(a, b) {
    return (a.order || 9999) - (b.order || 9999);
  });
}

function normalizePortalHexColor_(value, fallback) {
  const text = String(value || '').trim();
  if (/^#[0-9a-fA-F]{6}$/.test(text)) return text;
  if (/^[0-9a-fA-F]{6}$/.test(text)) return '#' + text;
  return fallback;
}
