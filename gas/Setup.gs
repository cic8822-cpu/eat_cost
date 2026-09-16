/**
 * Phase 1 초기 설정: 데이터 시트 열 추가, 설정/제출로그 시트 생성, 학교급식 AM~AP 헤더,
 * 직원ID 부여, 직원 동기화를 한 번에 수행한다.
 * 최초 1회는 스크립트 편집기에서 직접 실행해 권한 승인을 받아야 한다(원격 무인 실행 불가).
 */

function runInitialSetup() {
  var log = [];

  ensureDataSheetHeaders_(log);
  ensureSettingsSheet_(log);
  ensureLogSheet_(log);
  ensureMealHeaders_(log);

  var idRes = assignEmployeeIds();
  log.push('직원ID 부여: ' + idRes.data.assigned + '건');

  var syncRes = syncEmployeesToMealSheet_();
  log.push('학교급식 동기화: ' + syncRes.data.synced + '명');

  var billRes = rebuildPersonalStatement();
  log.push('개인별내역 재생성: ' + billRes.data.rows + '명, 합계 ' + billRes.data.totalAmount + '원');

  if (!getScriptProp_(ADMIN_TOKEN_KEY)) {
    var token = Utilities.getUuid().replace(/-/g, '');
    setScriptProp_(ADMIN_TOKEN_KEY, token);
    log.push('관리자 토큰 신규 발급: ' + token);
  } else {
    log.push('관리자 토큰: 기존 값 유지');
  }

  var check = validateSheetStructure();
  log.push('구조 점검: ' + (check.ok ? '문제 없음' : ('경고 있음 -> ' + check.message)));

  var summary = log.join('\n');
  Logger.log(summary);
  try { SpreadsheetApp.getUi().alert('초기 설정 완료\n\n' + summary); } catch (e) {
    // 스크립트 편집기에서 실행하면 Ui가 없어 예외가 나므로 무시(Logger.log로 대체 확인)
  }
  return ok_({ log: log });
}

function ensureDataSheetHeaders_(log) {
  var sheet = getSheet_(SHEET_NAMES.DATA);
  var header = sheet.getRange(1, 1, 1, 6).getValues()[0];
  var expected = ['순', '직급', '성명', '이메일', '직원ID', '재직여부'];
  var changed = false;
  for (var i = 0; i < expected.length; i++) {
    if (trimStr_(header[i]) !== expected[i]) {
      sheet.getRange(1, i + 1).setValue(expected[i]);
      changed = true;
    }
  }
  if (changed) {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var activeCol = sheet.getRange(2, DATA_COL.ACTIVE, lastRow - 1, 1).getValues();
      var toFill = [];
      for (var r = 0; r < activeCol.length; r++) {
        toFill.push([trimStr_(activeCol[r][0]) || 'Y']);
      }
      sheet.getRange(2, DATA_COL.ACTIVE, toFill.length, 1).setValues(toFill);
    }
  }
  log.push('데이터 시트 헤더: ' + (changed ? '갱신함(직원ID/재직여부 추가)' : '이미 정상'));
}

function ensureSettingsSheet_(log) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  var created = false;
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
    created = true;
  }

  var kvHeader = sheet.getRange(1, 1, 1, 2).getValues()[0];
  if (trimStr_(kvHeader[0]) !== 'KEY') {
    sheet.getRange(1, 1, 1, 2).setValues([['KEY', 'VALUE']]);
  }

  var defaults = [
    ['SCHOOL_NAME', ''],
    ['BANK_NAME', ''],
    ['PAYMENT_DEADLINE', ''],
    ['ADMIN_EMAIL', ''],
    ['MAIL_SUBJECT_TEMPLATE', '[학교급식] {년도}년 {월}월 급식 신청 및 급식비 안내']
  ];
  var existingKeys = {};
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function (r) {
      if (trimStr_(r[0])) existingKeys[trimStr_(r[0])] = true;
    });
  }
  var appendRows = defaults.filter(function (d) { return !existingKeys[d[0]]; });
  if (appendRows.length > 0) {
    sheet.getRange(Math.max(lastRow + 1, 2), 1, appendRows.length, 2).setValues(appendRows);
  }

  var dayHeader = sheet.getRange(1, 4, 1, 3).getValues()[0];
  if (trimStr_(dayHeader[0]) !== '날짜') {
    sheet.getRange(1, 4, 1, 3).setValues([['날짜', '상태', '사유']]);
  }

  log.push('설정 시트: ' + (created ? '신규 생성' : '기존 유지') + ', 운영설정 기본키 ' + appendRows.length + '건 추가');
}

function ensureLogSheet_(log) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_NAMES.LOG);
  var created = false;
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.LOG);
    created = true;
  }
  var header = sheet.getRange(1, 1, 1, 14).getValues()[0];
  if (trimStr_(header[0]) !== 'timestamp') {
    sheet.getRange(1, 1, 1, 14).setValues([[
      'timestamp', '년도', '월', '직원ID', '직급', '성명', '선택날짜',
      '급식일수', '급식단가', '급식비', '구분', '비고', '메일상태', '접속자식별'
    ]]);
  }
  log.push('제출로그 시트: ' + (created ? '신규 생성' : '기존 유지'));
}

// 웹앱이 로드될 때(getInitialData) 매번 가볍게 호출되는 자가복구 로직.
// 관리자가 마법사/초기설정 메뉴를 따로 실행하지 않고 데이터 시트만 채운 뒤
// 바로 배포해도 직원 목록이 정상적으로 뜨도록, 구조 확인·직원ID 부여·동기화를
// 자동으로 처리한다. 구조 확인(시트/헤더 생성)은 한 번만 하면 되므로
// STRUCTURE_READY 플래그로 캐시하고, 직원ID 부여·동기화는 매번 가볍게 재확인한다.
function ensureCoreStructureSilently_() {
  if (!getScriptProp_('STRUCTURE_READY')) {
    var log = [];
    ensureDataSheetHeaders_(log);
    ensureSettingsSheet_(log);
    ensureLogSheet_(log);
    ensureMealHeaders_(log);
    setScriptProp_('STRUCTURE_READY', 'true');
  }
  assignEmployeeIds();
  if (mealSheetNeedsSync_()) {
    syncEmployeesToMealSheet_();
  }
}

function setSettingValue_(key, value) {
  var sheet = getSheet_(SHEET_NAMES.SETTINGS);
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (trimStr_(keys[i][0]) === key) {
        sheet.getRange(2 + i, 2).setValue(value);
        return;
      }
    }
  }
  sheet.appendRow([key, value]);
}

// 다른 학교가 시트를 복사한 뒤 1회 실행하는 안내형 설정 마법사(Setup.html)의 서버 로직.
// runInitialSetup과 달리 학교명/계좌/급식단가/대상월을 사용자 입력값으로 즉시 반영한다.
function runGuidedSetup(formData) {
  try {
    formData = formData || {};
    var log = [];

    ensureDataSheetHeaders_(log);
    ensureSettingsSheet_(log);
    ensureLogSheet_(log);
    ensureMealHeaders_(log);

    setSettingValue_('SCHOOL_NAME', trimStr_(formData.schoolName));
    setSettingValue_('BANK_NAME', trimStr_(formData.bankName));
    setSettingValue_('PAYMENT_DEADLINE', trimStr_(formData.paymentDeadline));
    setSettingValue_('ADMIN_EMAIL', trimStr_(formData.adminEmail));

    setScriptProp_(BANK_ACCOUNT_KEY, trimStr_(formData.bankAccount));
    setScriptProp_(ACCOUNT_HOLDER_KEY, trimStr_(formData.accountHolder));

    var unitPrice = Number(formData.unitPrice);
    if (!unitPrice || unitPrice <= 0) return fail_('INVALID_INPUT', '급식단가를 올바르게 입력해 주세요.');
    var year = Number(formData.year);
    var month = Number(formData.month);
    if (!year || !month || month < 1 || month > 12) return fail_('INVALID_INPUT', '년도/월을 올바르게 입력해 주세요.');

    var meal = getSheet_(SHEET_NAMES.MEAL);
    meal.getRange(MEAL_UNIT_PRICE_CELL).setValue(unitPrice);

    var genRes = generateMonth(year, month);
    log.push('월 생성: ' + genRes.data.year + '년 ' + genRes.data.month + '월 (급식단가 ' + unitPrice + '원)');

    var idRes = assignEmployeeIds();
    log.push('직원ID 부여: ' + idRes.data.assigned + '건');

    var billRes = rebuildPersonalStatement();
    log.push('개인별내역 재생성: ' + billRes.data.rows + '명');

    var token = getScriptProp_(ADMIN_TOKEN_KEY);
    if (!token) {
      token = Utilities.getUuid().replace(/-/g, '');
      setScriptProp_(ADMIN_TOKEN_KEY, token);
      log.push('관리자 토큰 신규 발급');
    }

    setScriptProp_('SETUP_DONE', nowStamp_());

    var check = validateSheetStructure();

    return ok_({
      log: log,
      adminToken: token,
      structureOk: check.ok,
      structureMessage: check.ok ? '' : check.message
    });
  } catch (err) {
    return fail_('SETUP_ERROR', err.message);
  }
}

function ensureMealHeaders_(log) {
  var sheet = getSheet_(SHEET_NAMES.MEAL);
  var current = sheet.getRange(MEAL_HEADER_ROW, MEAL_COL.SUBMITTED_AT, 1, 4).getValues()[0];
  var expected = ['제출일시', '비고', '메일상태', '직원ID'];
  var changed = false;
  for (var i = 0; i < expected.length; i++) {
    if (trimStr_(current[i]) !== expected[i]) {
      sheet.getRange(MEAL_HEADER_ROW, MEAL_COL.SUBMITTED_AT + i).setValue(expected[i]);
      changed = true;
    }
  }
  log.push('학교급식 AM~AP 헤더: ' + (changed ? '갱신함' : '이미 정상'));
}
