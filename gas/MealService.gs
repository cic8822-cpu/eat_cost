/**
 * 급식 신청 핵심 로직: 기간/가능일 조회, 초기 데이터, 제출 조회/저장, 월 생성/마감
 */

function getPeriod_() {
  var meal = getSheet_(SHEET_NAMES.MEAL);
  var year = Number(meal.getRange(MEAL_YEAR_CELL).getValue());
  var month = Number(meal.getRange(MEAL_MONTH_CELL).getValue());
  var unitPrice = Number(meal.getRange(MEAL_UNIT_PRICE_CELL).getValue());
  // 학교급식 B1/D1이 비어있는 채(마법사·월 생성 실행 전) 배포되는 경우를 대비한 기본값
  var today = new Date();
  if (!year) year = today.getFullYear();
  if (!month || month < 1 || month > 12) month = today.getMonth() + 1;
  return { year: year, month: month, unitPrice: unitPrice || 0 };
}

function getSettingsMap_() {
  var sheet = getSheet_(SHEET_NAMES.SETTINGS);
  var lastRow = sheet.getLastRow();
  var map = {};
  if (lastRow >= SETTINGS_KV_RANGE.START_ROW) {
    var vals = sheet.getRange(SETTINGS_KV_RANGE.START_ROW, SETTINGS_KV_RANGE.COL_KEY, lastRow - SETTINGS_KV_RANGE.START_ROW + 1, 2).getValues();
    vals.forEach(function (r) {
      var key = trimStr_(r[0]);
      if (key) map[key] = trimStr_(r[1]);
    });
  }
  return map;
}

function getAvailableDaysMap_(year, month) {
  var sheet = getSheet_(SHEET_NAMES.SETTINGS);
  var lastRow = sheet.getLastRow();
  var map = {};
  if (lastRow >= SETTINGS_DAY_RANGE.START_ROW) {
    var vals = sheet.getRange(
      SETTINGS_DAY_RANGE.START_ROW, SETTINGS_DAY_RANGE.COL_DATE,
      lastRow - SETTINGS_DAY_RANGE.START_ROW + 1, 3
    ).getValues();
    vals.forEach(function (r) {
      var key = trimStr_(r[0]);
      if (!key) return;
      var p = parseDateKey_(key);
      if (p.y === year && p.m === month) {
        map[key] = { status: trimStr_(r[1]) || DAY_STATUS.OK, reason: trimStr_(r[2]) };
      }
    });
  }
  return map;
}

function buildDayList_(year, month) {
  var lastDay = lastDayOfMonth_(year, month);
  var availMap = getAvailableDaysMap_(year, month);
  var days = [];
  for (var d = 1; d <= lastDay; d++) {
    var key = toDateKey_(year, month, d);
    var info = availMap[key];
    var status = info ? info.status : (isWeekend_(year, month, d) ? DAY_STATUS.BLOCK : DAY_STATUS.OK);
    days.push({ key: key, label: dayLabel_(key), status: status, note: info ? info.reason : '' });
  }
  return days;
}

function getSubmissionSummary_() {
  var meal = getSheet_(SHEET_NAMES.MEAL);
  var lastRow = meal.getLastRow();
  if (lastRow < MEAL_DATA_START_ROW) return { total: 0, submitted: 0 };
  var vals = meal.getRange(MEAL_DATA_START_ROW, MEAL_COL.SUBMITTED_AT, lastRow - MEAL_DATA_START_ROW + 1, 1).getValues();
  var submitted = 0;
  vals.forEach(function (r) { if (trimStr_(r[0])) submitted++; });
  return { total: vals.length, submitted: submitted };
}

function getInitialData() {
  try {
    ensureCoreStructureSilently_();
    var period = getPeriod_();
    var employees = getEmployeesForClient_();
    var days = buildDayList_(period.year, period.month);
    var summary = getSubmissionSummary_();
    var settings = getSettingsMap_();
    return ok_({
      period: {
        year: period.year, month: period.month, unitPrice: period.unitPrice,
        title: period.year + '년 ' + period.month + '월 교직원 급식비 징수를 위한 급식일수 조사'
      },
      employees: employees,
      days: days,
      summary: summary,
      schoolName: settings.SCHOOL_NAME || '',
      paymentDeadline: settings.PAYMENT_DEADLINE || ''
    });
  } catch (err) {
    return fail_('SHEET_STRUCTURE_ERROR', err.message);
  }
}

function getMySubmission(empId) {
  try {
    var period = getPeriod_();
    var meal = getSheet_(SHEET_NAMES.MEAL);
    var lastRow = meal.getLastRow();
    if (lastRow < MEAL_DATA_START_ROW) return ok_({ dateKeys: [], note: '' });
    var vals = meal.getRange(MEAL_DATA_START_ROW, MEAL_COL.DAY_START, lastRow - MEAL_DATA_START_ROW + 1, MEAL_COL.DAY_COUNT + 4).getValues();
    for (var i = 0; i < vals.length; i++) {
      var row = vals[i];
      var rowEmpId = trimStr_(row[MEAL_COL.DAY_COUNT + 3]);
      if (rowEmpId === empId) {
        var dateKeys = [];
        for (var d = 0; d < MEAL_COL.DAY_COUNT; d++) {
          if (row[d] === 1 || row[d] === '1') dateKeys.push(toDateKey_(period.year, period.month, d + 1));
        }
        return ok_({ dateKeys: dateKeys, note: trimStr_(row[MEAL_COL.DAY_COUNT + 1]), submittedAt: trimStr_(row[MEAL_COL.DAY_COUNT]) });
      }
    }
    return ok_({ dateKeys: [], note: '' });
  } catch (err) {
    return fail_('SHEET_STRUCTURE_ERROR', err.message);
  }
}

function saveSubmission(payload) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    locked = lock.tryLock(30000);
    if (!locked) return fail_('LOCK_TIMEOUT', '다른 사용자가 제출 중입니다. 잠시 후 다시 시도해 주세요.');

    var empId = trimStr_(payload && payload.employeeId);
    var reqYear = Number(payload && payload.year);
    var reqMonth = Number(payload && payload.month);
    var dateKeys = (payload && payload.dateKeys) || [];
    var note = escapeCell_(payload && payload.note);

    if (!empId) return fail_('INVALID_EMPLOYEE', '직원을 선택해 주세요.');

    var period = getPeriod_();
    if (reqYear !== period.year || reqMonth !== period.month) {
      return fail_('MONTH_CHANGED', '급식 대상 월이 변경되었습니다. 새로고침 후 다시 시도해 주세요.');
    }

    var dataRows = readDataSheetRows_();
    var emp = findEmployeeById_(dataRows, empId);
    if (!emp || emp.active === 'N') return fail_('INVALID_EMPLOYEE', '유효하지 않은 직원입니다.');

    var uniqueKeys = Array.from(new Set(dateKeys)).sort();

    var availMap = getAvailableDaysMap_(period.year, period.month);
    var lastDay = lastDayOfMonth_(period.year, period.month);
    for (var i = 0; i < uniqueKeys.length; i++) {
      var key = uniqueKeys[i];
      var p = parseDateKey_(key);
      if (p.y !== period.year || p.m !== period.month || p.d < 1 || p.d > lastDay) {
        return fail_('INVALID_DATE', '유효하지 않은 날짜입니다: ' + key);
      }
      var info = availMap[key];
      var status = info ? info.status : (isWeekend_(p.y, p.m, p.d) ? DAY_STATUS.BLOCK : DAY_STATUS.OK);
      if (status === DAY_STATUS.BLOCK) {
        return fail_('NOT_AVAILABLE_DAY', '급식이 불가능한 날짜입니다: ' + dayLabel_(key));
      }
    }

    if (note && note.length > 200) return fail_('NOTE_TOO_LONG', '비고는 200자 이내로 입력해 주세요.');

    var meal = getSheet_(SHEET_NAMES.MEAL);
    var lastRow = meal.getLastRow();
    var targetRow = findMealRowByEmpId_(meal, lastRow, empId);
    if (targetRow === -1) {
      syncEmployeesToMealSheet_();
      lastRow = meal.getLastRow();
      targetRow = findMealRowByEmpId_(meal, lastRow, empId);
    }
    if (targetRow === -1) return fail_('INVALID_EMPLOYEE', '학교급식 시트에서 직원을 찾을 수 없습니다.');

    var prevSubmittedAt = trimStr_(meal.getRange(targetRow, MEAL_COL.SUBMITTED_AT).getValue());
    var isUpdate = !!prevSubmittedAt;

    var keySet = {};
    uniqueKeys.forEach(function (k) { keySet[k] = true; });
    var dayValues = [];
    for (var d2 = 1; d2 <= MEAL_COL.DAY_COUNT; d2++) {
      if (d2 <= lastDay) {
        var dk = toDateKey_(period.year, period.month, d2);
        dayValues.push(keySet[dk] ? 1 : '');
      } else {
        dayValues.push('');
      }
    }
    meal.getRange(targetRow, MEAL_COL.DAY_START, 1, MEAL_COL.DAY_COUNT).setValues([dayValues]);
    meal.getRange(targetRow, MEAL_COL.SUBMITTED_AT).setValue(nowStamp_());
    meal.getRange(targetRow, MEAL_COL.NOTE).setValue(note);
    SpreadsheetApp.flush();

    var days = uniqueKeys.length;
    var amount = days * period.unitPrice;

    rebuildPersonalStatement();

    appendSubmissionLog_({
      year: period.year, month: period.month, empId: empId,
      position: emp.position, name: emp.name,
      dateKeys: uniqueKeys, days: days, unitPrice: period.unitPrice, amount: amount,
      kind: isUpdate ? '수정' : '신규', note: note, mailStatus: '',
      viewer: getViewerId_()
    });

    lock.releaseLock();
    locked = false;

    var mailResult = sendResultMail_(emp, { days: days, unitPrice: period.unitPrice, amount: amount, year: period.year, month: period.month });
    meal.getRange(targetRow, MEAL_COL.MAIL_STATUS).setValue(mailResult.status);
    SpreadsheetApp.flush();

    return ok_({ days: days, unitPrice: period.unitPrice, amount: amount, emailStatus: mailResult.status, isUpdate: isUpdate });
  } catch (err) {
    return fail_('UNKNOWN_ERROR', err.message);
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e2) {} }
  }
}

function findMealRowByEmpId_(meal, lastRow, empId) {
  if (lastRow < MEAL_DATA_START_ROW) return -1;
  var idCol = meal.getRange(MEAL_DATA_START_ROW, MEAL_COL.EMP_ID, lastRow - MEAL_DATA_START_ROW + 1, 1).getValues();
  for (var r = 0; r < idCol.length; r++) {
    if (trimStr_(idCol[r][0]) === empId) return MEAL_DATA_START_ROW + r;
  }
  return -1;
}

function appendSubmissionLog_(entry) {
  var sheet = getSheet_(SHEET_NAMES.LOG);
  sheet.appendRow([
    nowStamp_(), entry.year, entry.month, entry.empId, entry.position, entry.name,
    entry.dateKeys.join(','), entry.days, entry.unitPrice, entry.amount,
    entry.kind, entry.note, entry.mailStatus, entry.viewer
  ]);
}

function generateMonth(year, month) {
  var meal = getSheet_(SHEET_NAMES.MEAL);
  meal.getRange(MEAL_YEAR_CELL).setValue(year);
  meal.getRange(MEAL_MONTH_CELL).setValue(month);

  var lastDay = lastDayOfMonth_(year, month);
  var headerVals = [];
  for (var i = 0; i < MEAL_COL.DAY_COUNT; i++) {
    var d = i + 1;
    headerVals.push(d <= lastDay ? d : '');
  }
  meal.getRange(MEAL_HEADER_ROW, MEAL_COL.DAY_START, 1, MEAL_COL.DAY_COUNT).setValues([headerVals]);

  ensureSettingsDaySheet_(year, month, lastDay);
  syncEmployeesToMealSheet_();
  SpreadsheetApp.flush();
  return ok_({ year: year, month: month, lastDay: lastDay });
}

function ensureSettingsDaySheet_(year, month, lastDay) {
  var sheet = getSheet_(SHEET_NAMES.SETTINGS);
  var lastRow = sheet.getLastRow();
  var existingKeys = {};
  if (lastRow >= SETTINGS_DAY_RANGE.START_ROW) {
    var vals = sheet.getRange(SETTINGS_DAY_RANGE.START_ROW, SETTINGS_DAY_RANGE.COL_DATE, lastRow - SETTINGS_DAY_RANGE.START_ROW + 1, 1).getValues();
    vals.forEach(function (r) { if (trimStr_(r[0])) existingKeys[trimStr_(r[0])] = true; });
  }
  var appendRows = [];
  for (var d = 1; d <= lastDay; d++) {
    var key = toDateKey_(year, month, d);
    if (existingKeys[key]) continue;
    var status = isWeekend_(year, month, d) ? DAY_STATUS.BLOCK : DAY_STATUS.OK;
    appendRows.push([key, status, '']);
  }
  if (appendRows.length > 0) {
    var startRow = Math.max(lastRow + 1, SETTINGS_DAY_RANGE.START_ROW);
    sheet.getRange(startRow, SETTINGS_DAY_RANGE.COL_DATE, appendRows.length, 3).setValues(appendRows);
  }
}

function closeMonth() {
  var meal = getSheet_(SHEET_NAMES.MEAL);
  var personal = getSheet_(SHEET_NAMES.PERSONAL);
  var period = getPeriod_();
  var archiveMealName = SHEET_NAMES.MEAL + '_' + period.year + '-' + pad2_(period.month);
  var archivePersonalName = SHEET_NAMES.PERSONAL + '_' + period.year + '-' + pad2_(period.month);

  var ss = getSpreadsheet_();
  if (ss.getSheetByName(archiveMealName)) ss.deleteSheet(ss.getSheetByName(archiveMealName));
  if (ss.getSheetByName(archivePersonalName)) ss.deleteSheet(ss.getSheetByName(archivePersonalName));

  freezeSheetValues_(meal.copyTo(ss).setName(archiveMealName));
  freezeSheetValues_(personal.copyTo(ss).setName(archivePersonalName));

  var lastRow = meal.getLastRow();
  if (lastRow >= MEAL_DATA_START_ROW) {
    meal.getRange(MEAL_DATA_START_ROW, MEAL_COL.DAY_START, lastRow - MEAL_DATA_START_ROW + 1, MEAL_COL.DAY_COUNT + 4).clearContent();
  }

  var nextMonth = period.month + 1, nextYear = period.year;
  if (nextMonth > 12) { nextMonth = 1; nextYear++; }
  generateMonth(nextYear, nextMonth);

  return ok_({ archived: [archiveMealName, archivePersonalName], next: { year: nextYear, month: nextMonth } });
}

function freezeSheetValues_(sheet) {
  var range = sheet.getDataRange();
  range.setValues(range.getValues());
}

// 학교급식 시트의 제출 데이터(급식일 체크·제출일시·비고·메일상태)만 초기화.
// 직원 목록(A~D)과 직원ID(AP)는 보존한다. 년/월을 바꿔도 지워지지 않던 잔존 데이터를
// 담당자가 수동으로 깨끗이 비우기 위한 용도(archive 없이 즉시 삭제, 되돌릴 수 없음).
function resetSubmissions() {
  var meal = getSheet_(SHEET_NAMES.MEAL);
  var lastRow = meal.getLastRow();
  var clearedRows = 0;
  if (lastRow >= MEAL_DATA_START_ROW) {
    clearedRows = lastRow - MEAL_DATA_START_ROW + 1;
    meal.getRange(
      MEAL_DATA_START_ROW, MEAL_COL.DAY_START,
      clearedRows, MEAL_COL.MAIL_STATUS - MEAL_COL.DAY_START + 1 // H ~ AO (AP 직원ID는 보존)
    ).clearContent();
  }
  SpreadsheetApp.flush();

  var billRes = rebuildPersonalStatement();
  return ok_({ clearedRows: clearedRows, personal: billRes.data });
}
