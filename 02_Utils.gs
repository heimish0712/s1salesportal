/***************************************
 * S1 Sales Portal - 02_Utils.gs
 * 분리일: 2026-06-19
 * 원칙: 기능 변경 없이 최신 단일 파일을 물리적으로 분리
 ***************************************/

function containsAny_(value, keywords) {
  const text = String(value || '');
  return (keywords || []).some(k => text.indexOf(k) >= 0);
}

function formatDateText_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  }
  return String(value);
}

function formatDateTimeText_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy. MM. dd HH:mm');
  }
  return String(value || '');
}

function formatMemoDateTime_(value) {
  const date = parsePortalDateTime_(value);
  if (!date) return String(value || '').trim();
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yy.MM.dd HH:mm');
}

function parsePortalDateTime_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]') return value;

  const text = String(value || '').trim();
  let m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})[T\s](\d{1,2}):(\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]), 0);

  m = text.match(/^(\d{2})[.\/-](\d{1,2})[.\/-](\d{1,2})\s+(\d{1,2}):(\d{1,2})/);
  if (m) return new Date(2000 + Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]), 0);

  m = text.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})\s+(\d{1,2}):(\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]), 0);

  const d = new Date(text);
  if (!isNaN(d.getTime())) return d;
  return null;
}

function shortenAddressForList_(address) {
  const text = String(address || '').trim().replace(/\s+/g, ' ');
  if (!text) return '';

  const parts = text.split(' ').filter(Boolean);
  if (!parts.length) return '';

  if (parts[0].indexOf('세종') >= 0) {
    const idx = parts.findIndex(function(p, i) {
      return i > 0 && /(동|읍|면|리)$/.test(p);
    });
    return parts.slice(0, idx >= 0 ? idx + 1 : Math.min(parts.length, 2)).join(' ');
  }

  const guIdx = parts.findIndex(function(p, i) { return i > 0 && /구$/.test(p); });
  if (guIdx >= 0) return parts.slice(0, guIdx + 1).join(' ');

  const gunIdx = parts.findIndex(function(p, i) { return i > 0 && /군$/.test(p); });
  if (gunIdx >= 0) return parts.slice(0, gunIdx + 1).join(' ');

  const dongIdx = parts.findIndex(function(p, i) {
    return i > 0 && /(동|읍|면)$/.test(p);
  });
  if (dongIdx >= 0) return parts.slice(0, dongIdx + 1).join(' ');

  return parts.slice(0, Math.min(parts.length, 2)).join(' ');
}

function getCurrentUserLabel_() {
  const candidates = [];

  try {
    candidates.push(String(Session.getActiveUser().getEmail() || '').trim());
  } catch (err) {
    Logger.log('Session.getActiveUser 이메일 확인 실패: ' + (err && err.message || err));
  }

  try {
    candidates.push(String(Session.getEffectiveUser().getEmail() || '').trim());
  } catch (err) {
    Logger.log('Session.getEffectiveUser 이메일 확인 실패: ' + (err && err.message || err));
  }

  const email = candidates.filter(Boolean)[0] || '';
  const map = PORTAL_CONFIG.USER_DISPLAY_NAME_MAP || {};
  if (email && map[email]) return map[email];
  if (email) return email.split('@')[0];
  return '웹앱사용자';
}

function shortUuid_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 8);
}

function normalizeGrade_(grade) {
  return String(grade || '').replace(/\s+/g, '').replace(/\(.*?\)/g, '').trim();
}

function parseMoney_(value) {
  if (value == null) return 0;
  const text = String(value).replace(/[^0-9.\-]/g, '');
  const n = Number(text);
  return isNaN(n) ? 0 : n;
}

function parseCount_(value) {
  const m = String(value || '').match(/\d+/);
  return m ? Number(m[0]) : 0;
}

function roundDownToUnit_(value, unit) {
  value = Number(value) || 0;
  unit = Number(unit) || 1;
  if (value >= 0) return Math.floor(value / unit) * unit;
  return Math.ceil(value / unit) * unit;
}

function formatWon_(n) {
  n = Number(n) || 0;
  return '₩' + Math.round(n).toLocaleString('ko-KR');
}

// =========================
// P1-4: 주소/지역구분 자동 규격화
// - 주소는 수기 입력을 유지하되 저장 시 축약 행정구역명을 표준명으로 보정합니다.
// - 지역구분은 회의 기준 5개 권역만 자동 분류합니다.
//   수도권 / 강원권 / 충청권 / 부울경권 / 대구경북권
// =========================
function normalizePortalAddress_(address) {
  let text = String(address == null ? '' : address).trim().replace(/\s+/g, ' ');
  if (!text) return '';

  const replacements = [
    [/^서울시(?=\s|$)/, '서울특별시'],
    [/^서울(?=\s|$)/, '서울특별시'],
    [/^부산시(?=\s|$)/, '부산광역시'],
    [/^부산(?=\s|$)/, '부산광역시'],
    [/^대구시(?=\s|$)/, '대구광역시'],
    [/^대구(?=\s|$)/, '대구광역시'],
    [/^인천시(?=\s|$)/, '인천광역시'],
    [/^인천(?=\s|$)/, '인천광역시'],
    [/^광주시(?=\s|$)/, '광주광역시'],
    [/^광주(?=\s|$)/, '광주광역시'],
    [/^대전시(?=\s|$)/, '대전광역시'],
    [/^대전(?=\s|$)/, '대전광역시'],
    [/^울산시(?=\s|$)/, '울산광역시'],
    [/^울산(?=\s|$)/, '울산광역시'],
    [/^세종시(?=\s|$)/, '세종특별자치시'],
    [/^세종(?=\s|$)/, '세종특별자치시'],
    [/^경기(?=\s|$)/, '경기도'],
    [/^강원도(?=\s|$)/, '강원특별자치도'],
    [/^강원(?=\s|$)/, '강원특별자치도'],
    [/^충남(?=\s|$)/, '충청남도'],
    [/^충북(?=\s|$)/, '충청북도'],
    [/^충청남도(?=\s|$)/, '충청남도'],
    [/^충청북도(?=\s|$)/, '충청북도'],
    [/^경남(?=\s|$)/, '경상남도'],
    [/^경북(?=\s|$)/, '경상북도'],
    [/^경상남도(?=\s|$)/, '경상남도'],
    [/^경상북도(?=\s|$)/, '경상북도'],
    [/^전남(?=\s|$)/, '전라남도'],
    [/^전북(?=\s|$)/, '전북특별자치도'],
    [/^전라남도(?=\s|$)/, '전라남도'],
    [/^전라북도(?=\s|$)/, '전북특별자치도'],
    [/^제주도(?=\s|$)/, '제주특별자치도'],
    [/^제주(?=\s|$)/, '제주특별자치도']
  ];

  for (let i = 0; i < replacements.length; i++) {
    const pair = replacements[i];
    if (pair[0].test(text)) {
      text = text.replace(pair[0], pair[1]);
      break;
    }
  }
  return text.trim().replace(/\s+/g, ' ');
}

function normalizePortalRegionKeyword_(value) {
  return String(value == null ? '' : value)
    .trim()
    .replace(/\s+/g, '')
    .replace(/특별자치도|특별자치시|특별시|광역시|자치도|도|시/g, '');
}

function derivePortalRegionGroup_(addressOrRegion) {
  const raw = String(addressOrRegion == null ? '' : addressOrRegion).trim();
  if (!raw) return '';
  const compact = normalizePortalRegionKeyword_(raw);

  if (/^(수도권|서울|경기|인천)/.test(compact) || /서울특별시|경기도|인천광역시/.test(raw)) return '수도권';
  if (/^(강원권|강원)/.test(compact) || /강원특별자치도|강원도/.test(raw)) return '강원권';
  if (/^(충청권|충청|충남|충북|대전|세종)/.test(compact) || /충청남도|충청북도|대전광역시|세종특별자치시/.test(raw)) return '충청권';
  if (/^(부울경권|부울경|부산|울산|경남)/.test(compact) || /부산광역시|울산광역시|경상남도/.test(raw)) return '부울경권';
  if (/^(대구경북권|대경권|대구경북|대구|경북)/.test(compact) || /대구광역시|경상북도/.test(raw)) return '대구경북권';

  return '';
}

function normalizePortalRegionGroup_(regionValue, addressValue) {
  const fromAddress = derivePortalRegionGroup_(addressValue);
  if (fromAddress) return fromAddress;

  const fromRegion = derivePortalRegionGroup_(regionValue);
  if (fromRegion) return fromRegion;

  // 회의 기준 5개 권역에 없는 지역(호남/제주 등)은 임의 변환하지 않고 기존 입력값을 보존합니다.
  return String(regionValue == null ? '' : regionValue).trim();
}

function normalizePortalCustomerLocationValues_(values) {
  const result = Object.assign({}, values || {});
  const hasAddress = Object.prototype.hasOwnProperty.call(result, 'address');
  const hasRegion = Object.prototype.hasOwnProperty.call(result, 'region');

  const normalizedAddress = hasAddress ? normalizePortalAddress_(result.address) : '';
  if (hasAddress) result.address = normalizedAddress;

  const derivedRegion = normalizePortalRegionGroup_(hasRegion ? result.region : '', normalizedAddress || result.address || '');
  if (derivedRegion) {
    result.region = derivedRegion;
  } else if (hasRegion) {
    result.region = String(result.region == null ? '' : result.region).trim();
  }

  return result;
}

