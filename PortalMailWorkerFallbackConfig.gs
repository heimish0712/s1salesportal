/***************************************
 * S1 Sales Portal - Mail Worker 고정 설정 fallback v83
 *
 * 사용법:
 * 1) 이 파일을 포털 Apps Script 프로젝트에 추가합니다.
 * 2) 아래 WEBAPP_URL / SHARED_SECRET에 값을 1번만 넣습니다.
 * 3) 11_MailBridgeService.gs v83과 함께 쓰면 Script Properties가 비어도 자동 복구됩니다.
 *
 * 보안상 이 파일을 외부 공유하지 마세요.
 ***************************************/
var PORTAL_MAIL_WORKER_FALLBACK_CONFIG_V83 = {
  WEBAPP_URL: 'https://script.google.com/macros/s/AKfycbw0fZ5qmA5ABTQZibqAOdHMEJkSRus99TjvhRmZCuDOkHLBelc2B36tDZAqvWYCjYIQ/exec',       // 예: https://script.google.com/macros/s/AKfycb.../exec
  SHARED_SECRET: 'dladmlanswkduf'     // 포털/메일 Worker 양쪽에 동일하게 넣는 secret
};
