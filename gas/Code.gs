/**
 * 진입점(doGet), 관리자 메뉴(onOpen), 스프레드시트 직접 편집 경고(onEdit)
 */

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  template.adminToken = (e && e.parameter && e.parameter.admin) || '';
  return template.evaluate()
    .setTitle('교직원 급식비 신청')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('급식비 관리')
    .addItem('초기 설정(1회)', 'runInitialSetup')
    .addItem('구조 점검', 'menuValidateStructure_')
    .addItem('직원ID 부여', 'menuAssignEmployeeIds_')
    .addItem('직원 동기화', 'menuSyncEmployees_')
    .addSeparator()
    .addItem('월 생성', 'menuGenerateMonth_')
    .addItem('월 마감(다음 달로 전환)', 'menuCloseMonth_')
    .addSeparator()
    .addItem('개인별내역 재생성', 'menuRebuildPersonal_')
    .addItem('테스트 메일(본인)', 'menuSendTestMail_')
    .addItem('관리자 토큰 재발급', 'menuRegenerateAdminToken_')
    .addToUi();
}

function menuValidateStructure_() {
  var res = validateSheetStructure();
  SpreadsheetApp.getUi().alert(res.ok ? '문제 없음' : ('경고:\n' + res.message));
}

function menuAssignEmployeeIds_() {
  var res = assignEmployeeIds();
  SpreadsheetApp.getUi().alert('직원ID ' + res.data.assigned + '건 부여됨');
}

function menuSyncEmployees_() {
  var res = syncEmployeesToMealSheet_();
  SpreadsheetApp.getUi().alert('직원 동기화 완료: ' + res.data.synced + '명');
}

function menuGenerateMonth_() {
  var ui = SpreadsheetApp.getUi();
  var yearResp = ui.prompt('월 생성', '년도를 입력하세요 (예: 2026)', ui.ButtonSet.OK_CANCEL);
  if (yearResp.getSelectedButton() !== ui.Button.OK) return;
  var monthResp = ui.prompt('월 생성', '월을 입력하세요 (1~12)', ui.ButtonSet.OK_CANCEL);
  if (monthResp.getSelectedButton() !== ui.Button.OK) return;
  var res = generateMonth(Number(yearResp.getResponseText()), Number(monthResp.getResponseText()));
  ui.alert('월 생성 완료: ' + res.data.year + '년 ' + res.data.month + '월 (말일 ' + res.data.lastDay + '일)');
}

function menuCloseMonth_() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert('월 마감', '현재 월 데이터를 보관(아카이브)하고 다음 달로 전환합니다. 계속할까요?', ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  var res = closeMonth();
  ui.alert('월 마감 완료. 다음 월: ' + res.data.next.year + '년 ' + res.data.next.month + '월');
}

function menuRebuildPersonal_() {
  var res = rebuildPersonalStatement();
  SpreadsheetApp.getUi().alert('개인별내역 재생성 완료: ' + res.data.rows + '명, 합계 ' + res.data.totalAmount + '원');
}

function menuSendTestMail_() {
  var res = sendTestMailToSelf();
  SpreadsheetApp.getUi().alert('테스트 메일 발송: ' + res.data.to);
}

function menuRegenerateAdminToken_() {
  var token = Utilities.getUuid().replace(/-/g, '');
  setScriptProp_(ADMIN_TOKEN_KEY, token);
  var url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('관리자 토큰이 재발급되었습니다.\n관리자 URL:\n' + (url || '(먼저 웹앱으로 배포하세요)') + '?admin=' + token);
}

// 학교급식 시트의 년/월 셀을 직접 수정했을 때 경고(값을 되돌리지는 않음)
function onEdit(e) {
  try {
    var sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAMES.MEAL) return;
    var a1 = e.range.getA1Notation();
    if (a1 === MEAL_YEAR_CELL || a1 === MEAL_MONTH_CELL) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        '월 전환은 메뉴 "급식비 관리 > 월 마감"을 사용하세요. 기존 신청자료가 남아있을 수 있습니다.',
        '경고', 8
      );
    }
  } catch (err) {
    // onEdit 트리거는 예외를 조용히 무시(사용자 편집 흐름을 막지 않음)
  }
}
