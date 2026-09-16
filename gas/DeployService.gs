/**
 * 웹 앱 자동 배포. Apps Script API(script.googleapis.com)를 스스로 호출해
 * "배포 → 새 배포 → 웹 앱" 수동 절차 없이 버튼 클릭 한 번으로 배포/재배포한다.
 * ScriptApp.getOAuthToken()이 appsscript.json에 선언된 scope 범위 내에서만
 * 동작하므로, script.deployments / script.projects가 반드시 선언돼 있어야 한다.
 */

function autoDeployWebApp_(description) {
  var scriptId = ScriptApp.getScriptId();
  var token = ScriptApp.getOAuthToken();
  var base = 'https://script.googleapis.com/v1/projects/' + scriptId;
  var desc = description || ('자동 배포 ' + nowStamp_());

  var verResp = UrlFetchApp.fetch(base + '/versions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ description: desc }),
    muteHttpExceptions: true
  });
  var verData = JSON.parse(verResp.getContentText());
  if (verResp.getResponseCode() >= 300) {
    throw new Error('버전 생성 실패: ' + (verData.error ? verData.error.message : verResp.getContentText()));
  }
  var versionNumber = verData.versionNumber;

  var listResp = UrlFetchApp.fetch(base + '/deployments', {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  var listData = JSON.parse(listResp.getContentText());
  var existing = null;
  (listData.deployments || []).forEach(function (d) {
    var hasWebApp = (d.entryPoints || []).some(function (ep) { return ep.entryPointType === 'WEB_APP'; });
    if (hasWebApp) existing = d;
  });

  var deployment;
  if (existing) {
    var updResp = UrlFetchApp.fetch(base + '/deployments/' + existing.deploymentId, {
      method: 'put',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({
        deploymentConfig: { scriptId: scriptId, versionNumber: versionNumber, manifestFileName: 'appsscript', description: desc }
      }),
      muteHttpExceptions: true
    });
    deployment = JSON.parse(updResp.getContentText());
    if (updResp.getResponseCode() >= 300) {
      throw new Error('배포 갱신 실패: ' + (deployment.error ? deployment.error.message : updResp.getContentText()));
    }
  } else {
    var createResp = UrlFetchApp.fetch(base + '/deployments', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ versionNumber: versionNumber, manifestFileName: 'appsscript', description: desc }),
      muteHttpExceptions: true
    });
    deployment = JSON.parse(createResp.getContentText());
    if (createResp.getResponseCode() >= 300) {
      throw new Error('배포 생성 실패: ' + (deployment.error ? deployment.error.message : createResp.getContentText()));
    }
  }

  var webAppEntry = (deployment.entryPoints || []).filter(function (ep) { return ep.entryPointType === 'WEB_APP'; })[0];
  var url = webAppEntry ? webAppEntry.webApp.url : '';
  if (!url) throw new Error('배포는 되었으나 웹 앱 URL을 가져오지 못했습니다. Apps Script 편집기에서 확인해 주세요.');

  setScriptProp_('WEB_APP_URL', url);
  return { url: url, deploymentId: deployment.deploymentId, versionNumber: versionNumber };
}

// 설정 마법사(SetupWizard.html)의 "지금 배포하기" 버튼이 호출하는 공개 엔드포인트
function deployNow() {
  try {
    var res = autoDeployWebApp_();
    var adminToken = getScriptProp_(ADMIN_TOKEN_KEY);
    return ok_({ url: res.url, adminUrl: adminToken ? (res.url + '?admin=' + adminToken) : '' });
  } catch (err) {
    return fail_('DEPLOY_ERROR', err.message);
  }
}
