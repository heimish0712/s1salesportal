/***************************************
 * S1 Sales Portal - Code.gs
 * 분리일: 2026-06-19
 * 원칙: 기능 변경 없이 최신 단일 파일을 물리적으로 분리
 ***************************************/

/***************************************
 * PATCH S-FIX4: 권한_DB 기반 웹앱 입구 차단형 doGet
 * 사용법: 기존 doGet(e) 함수만 이 블록으로 교체하세요.
 * include(filename)는 기존 함수가 있으면 중복 추가하지 마세요.
 ***************************************/
function doGet(e) {
  const action = String(e && e.parameter && e.parameter.action || '').trim();

  if (action === 'mailAutoMultiDownload') {
    return serveMailAutoMultiDownload_(e);
  }

  const auth = getPortalWebAppEntryAuth_();
  if (!auth.ok) {
    return HtmlService
      .createHtmlOutput(buildPortalAccessDeniedHtml_(auth))
      .setTitle('S1 Sales Portal - Access denied')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.portalBootPermission = auth.permission;

  return template
    .evaluate()
    .setTitle('S1 Sales Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

