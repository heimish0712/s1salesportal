
/***************************************
 * S1 Sales Portal - 19_RequestGuardService.gs
 * P2-6: 중복 요청/연속 클릭 방지 공통 가드
 ***************************************/

function makePortalRequestGuardKeyP26_(type, requestId) {
  type = String(type || 'REQ').replace(/[^A-Za-z0-9_\-]/g, '_').slice(0, 40) || 'REQ';
  requestId = String(requestId || '').replace(/[^A-Za-z0-9_\-:.]/g, '_').slice(0, 120);
  return 'PORTAL_REQ_GUARD__' + type + '__' + requestId;
}

function beginPortalIdempotentRequestP26_(type, requestId, options) {
  requestId = String(requestId || '').trim();
  if (!requestId) return { enabled: false, duplicate: false, key: '' };
  options = options || {};
  var runningTtlMs = Number(options.runningTtlMs || 10 * 60 * 1000);
  var doneTtlMs = Number(options.doneTtlMs || 30 * 60 * 1000);
  var now = Date.now();
  var key = makePortalRequestGuardKeyP26_(type, requestId);
  return withPortalScriptLockP201_('request-guard-' + type, function() {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(key);
    if (raw) {
      try {
        var state = JSON.parse(raw);
        var age = now - Number(state.ts || 0);
        if (state.status === 'RUNNING' && age < runningTtlMs) {
          return { enabled: true, duplicate: true, running: true, key: key, state: state };
        }
        if (state.status === 'DONE' && age < doneTtlMs) {
          return { enabled: true, duplicate: true, done: true, key: key, state: state };
        }
      } catch (e) {}
    }
    props.setProperty(key, JSON.stringify({ status: 'RUNNING', ts: now, type: type, requestId: requestId }));
    return { enabled: true, duplicate: false, started: true, key: key, requestId: requestId };
  }, { waitMs: 1200 });
}

function finishPortalIdempotentRequestP26_(guard, resultSummary) {
  if (!guard || !guard.enabled || !guard.key || guard.duplicate) return;
  try {
    PropertiesService.getScriptProperties().setProperty(guard.key, JSON.stringify({
      status: 'DONE',
      ts: Date.now(),
      result: resultSummary || {}
    }));
  } catch (e) {}
}

function failPortalIdempotentRequestP26_(guard) {
  if (!guard || !guard.enabled || !guard.key || guard.duplicate) return;
  try { PropertiesService.getScriptProperties().deleteProperty(guard.key); } catch (e) {}
}

function throwPortalDuplicateRunningP26_(label) {
  throw new Error((label || '요청') + '이 이미 처리 중입니다. 화면의 진행상태를 확인해 주세요.');
}
