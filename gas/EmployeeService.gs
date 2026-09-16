/**
 * 직원 마스터(데이터 시트) 관리: 목록 조회, 직원ID 부여, 학교급식 시트 동기화
 */

function readDataSheetRows_() {
  var sheet = getSheet_(SHEET_NAMES.DATA);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    var name = trimStr_(r[DATA_COL.NAME - 1]);
    if (!name) continue;
    rows.push({
      sheetRow: i + 2,
      order: r[DATA_COL.ORDER - 1],
      position: trimStr_(r[DATA_COL.POSITION - 1]),
      name: name,
      email: trimStr_(r[DATA_COL.EMAIL - 1]),
      empId: trimStr_(r[DATA_COL.EMP_ID - 1]),
      active: trimStr_(r[DATA_COL.ACTIVE - 1]) || 'Y'
    });
  }
  return rows;
}

// 관리자 메뉴: 직원ID가 비어 있는 행에만 자동 부여 (기존 ID는 변경하지 않음)
function assignEmployeeIds() {
  var sheet = getSheet_(SHEET_NAMES.DATA);
  var rows = readDataSheetRows_();
  var maxSeq = 0;
  rows.forEach(function (r) {
    var m = /^EMP(\d+)$/.exec(r.empId);
    if (m) maxSeq = Math.max(maxSeq, Number(m[1]));
  });
  var assigned = 0;
  rows.forEach(function (r) {
    if (!r.empId) {
      maxSeq++;
      var newId = 'EMP' + ('000' + maxSeq).slice(-3);
      sheet.getRange(r.sheetRow, DATA_COL.EMP_ID).setValue(newId);
      assigned++;
    }
  });
  SpreadsheetApp.flush();
  return ok_({ assigned: assigned });
}

function findEmployeeById_(rows, empId) {
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].empId === empId) return rows[i];
  }
  return null;
}

// 동명이인 구분용 표시 이름: "직급 · 성명" (중복 시 뒤에 번호 접미)
function buildDisplayNames_(rows) {
  var countByKey = {};
  rows.forEach(function (r) {
    var key = r.position + '·' + r.name;
    countByKey[key] = (countByKey[key] || 0) + 1;
  });
  var seenKey = {};
  return rows.map(function (r) {
    var key = r.position + '·' + r.name;
    var label = r.position + ' · ' + r.name;
    if (countByKey[key] > 1) {
      seenKey[key] = (seenKey[key] || 0) + 1;
      label += ' (' + seenKey[key] + ')';
    }
    return { id: r.empId, position: r.position, name: r.name, displayName: label, active: r.active };
  });
}

// 클라이언트에 내려줄 직원 목록 (이메일 비노출, 재직중인 직원만)
function getEmployeesForClient_() {
  var rows = readDataSheetRows_();
  var display = buildDisplayNames_(rows);
  return display
    .filter(function (d) { return d.active !== 'N' && d.id; })
    .map(function (d) { return { id: d.id, position: d.position, displayName: d.displayName }; });
}

// 학교급식 시트를 데이터 시트와 직원ID 기준으로 동기화 (행 삭제/추가/정렬에도 신청 데이터 보존)
function syncEmployeesToMealSheet_() {
  var meal = getSheet_(SHEET_NAMES.MEAL);
  var rows = readDataSheetRows_().filter(function (r) { return r.active !== 'N' && r.empId; });

  var lastMealRow = meal.getLastRow();
  var existing = {};
  if (lastMealRow >= MEAL_DATA_START_ROW) {
    var vals = meal.getRange(
      MEAL_DATA_START_ROW, MEAL_COL.DAY_START,
      lastMealRow - MEAL_DATA_START_ROW + 1,
      MEAL_COL.DAY_COUNT + 4 // 일자 31칸 + AM,AN,AO,AP
    ).getValues();
    for (var i = 0; i < vals.length; i++) {
      var row = vals[i];
      var empId = trimStr_(row[MEAL_COL.DAY_COUNT + 3]); // AP
      if (empId) existing[empId] = row;
    }
  }

  if (lastMealRow >= MEAL_DATA_START_ROW) {
    meal.getRange(MEAL_DATA_START_ROW, MEAL_COL.ORDER, lastMealRow - MEAL_DATA_START_ROW + 1, MEAL_COL.EMP_ID).clearContent();
  }

  rows.forEach(function (r, idx) {
    var targetRow = MEAL_DATA_START_ROW + idx;
    meal.getRange(targetRow, MEAL_COL.ORDER).setValue(idx + 1);
    meal.getRange(targetRow, MEAL_COL.POSITION).setValue(r.position);
    meal.getRange(targetRow, MEAL_COL.NAME).setValue(r.name);
    meal.getRange(targetRow, MEAL_COL.EMAIL).setValue(r.email);
    meal.getRange(targetRow, MEAL_COL.TOTAL_DAYS).setFormula('=SUM(' +
      colLetter_(MEAL_COL.DAY_START) + targetRow + ':' + colLetter_(MEAL_COL.DAY_START + MEAL_COL.DAY_COUNT - 1) + targetRow + ')');
    meal.getRange(targetRow, MEAL_COL.UNIT_PRICE).setFormula('=$F$1');
    meal.getRange(targetRow, MEAL_COL.AMOUNT).setFormula('=' + colLetter_(MEAL_COL.TOTAL_DAYS) + targetRow + '*' + colLetter_(MEAL_COL.UNIT_PRICE) + targetRow);

    var prev = existing[r.empId];
    if (prev) {
      var dayValues = prev.slice(0, MEAL_COL.DAY_COUNT);
      meal.getRange(targetRow, MEAL_COL.DAY_START, 1, MEAL_COL.DAY_COUNT).setValues([dayValues]);
      meal.getRange(targetRow, MEAL_COL.SUBMITTED_AT).setValue(prev[MEAL_COL.DAY_COUNT]);
      meal.getRange(targetRow, MEAL_COL.NOTE).setValue(prev[MEAL_COL.DAY_COUNT + 1]);
      meal.getRange(targetRow, MEAL_COL.MAIL_STATUS).setValue(prev[MEAL_COL.DAY_COUNT + 2]);
    }
    meal.getRange(targetRow, MEAL_COL.EMP_ID).setValue(r.empId);
  });

  var newLastRow = MEAL_DATA_START_ROW + rows.length - 1;
  if (lastMealRow > newLastRow) {
    meal.getRange(newLastRow + 1, MEAL_COL.ORDER, lastMealRow - newLastRow, MEAL_COL.EMP_ID).clearContent();
  }

  SpreadsheetApp.flush();
  return ok_({ synced: rows.length });
}
