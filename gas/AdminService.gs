/**
 * 관리자 전용 조회(토큰 필요) 및 시트 구조 자기진단
 */

function checkAdminToken_(token) {
  var real = getScriptProp_(ADMIN_TOKEN_KEY);
  return !!real && String(token) === real;
}

// 관리자 상세표 (토큰 일치 시에만 반환, 이메일 열은 포함하지 않음)
function getAdminDetail(token) {
  if (!checkAdminToken_(token)) return fail_('FORBIDDEN', '접근 권한이 없습니다.');
  try {
    var meal = getSheet_(SHEET_NAMES.MEAL);
    var lastRow = meal.getLastRow();
    var rows = [];
    if (lastRow >= MEAL_DATA_START_ROW) {
      var vals = meal.getRange(MEAL_DATA_START_ROW, MEAL_COL.ORDER, lastRow - MEAL_DATA_START_ROW + 1, MEAL_COL.MAIL_STATUS - MEAL_COL.ORDER + 1).getValues();
      vals.forEach(function (r) {
        var name = trimStr_(r[MEAL_COL.NAME - 1]);
        if (!name) return;
        rows.push({
          position: r[MEAL_COL.POSITION - 1],
          name: name,
          days: Number(r[MEAL_COL.TOTAL_DAYS - 1]) || 0,
          unitPrice: Number(r[MEAL_COL.UNIT_PRICE - 1]) || 0,
          amount: Number(r[MEAL_COL.AMOUNT - 1]) || 0,
          submittedAt: r[MEAL_COL.SUBMITTED_AT - 1],
          note: r[MEAL_COL.NOTE - 1],
          mailStatus: r[MEAL_COL.MAIL_STATUS - 1]
        });
      });
    }
    return ok_({ rows: rows });
  } catch (err) {
    return fail_('SHEET_STRUCTURE_ERROR', err.message);
  }
}

// 관리자 메뉴: 시트 구조/설정값 자기진단
function validateSheetStructure() {
  var problems = [];
  var ss = getSpreadsheet_();
  Object.keys(SHEET_NAMES).forEach(function (k) {
    if (!ss.getSheetByName(SHEET_NAMES[k])) problems.push('시트 없음: ' + SHEET_NAMES[k]);
  });
  if (problems.length > 0) return fail_('SHEET_STRUCTURE_ERROR', problems.join(', '));

  var dataSheet = getSheet_(SHEET_NAMES.DATA);
  var header = dataSheet.getRange(1, 1, 1, 6).getValues()[0].map(trimStr_);
  var expected = ['순', '직급', '성명', '이메일', '직원ID', '재직여부'];
  for (var i = 0; i < expected.length; i++) {
    if (header[i] !== expected[i]) {
      problems.push('데이터 시트 헤더 불일치(' + colLetter_(i + 1) + '열): 실제="' + header[i] + '" 기대="' + expected[i] + '"');
    }
  }

  var rows = readDataSheetRows_();
  rows.forEach(function (r) {
    if (!r.email) {
      problems.push('이메일 없음: ' + r.name);
    } else if (!isValidEmail_(r.email)) {
      problems.push('이메일 형식 오류: ' + r.name);
    }
  });

  if (!getScriptProp_(ADMIN_TOKEN_KEY)) problems.push('ADMIN_TOKEN 미설정 (메뉴: 관리자 토큰 재발급)');
  if (!getScriptProp_(BANK_ACCOUNT_KEY)) problems.push('BANK_ACCOUNT 미설정 (프로젝트 설정 > 스크립트 속성)');
  if (!getScriptProp_(ACCOUNT_HOLDER_KEY)) problems.push('ACCOUNT_HOLDER 미설정 (프로젝트 설정 > 스크립트 속성)');

  if (problems.length === 0) return ok_({ message: '문제 없음' });
  return fail_('VALIDATION_WARNING', problems.join('\n'), { problems: problems });
}
