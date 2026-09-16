/**
 * 개인별내역 시트 재생성 (학교급식 시트 기준으로 전체 다시 씀)
 */

function rebuildPersonalStatement() {
  var meal = getSheet_(SHEET_NAMES.MEAL);
  var personal = getSheet_(SHEET_NAMES.PERSONAL);
  var period = getPeriod_();

  personal.getRange(PERSONAL_TITLE_ROW, 1).setValue(period.year + '년 ' + period.month + '월 교직원급식비 징수내역');

  var lastRow = meal.getLastRow();
  var rows = [];
  if (lastRow >= MEAL_DATA_START_ROW) {
    var vals = meal.getRange(MEAL_DATA_START_ROW, MEAL_COL.ORDER, lastRow - MEAL_DATA_START_ROW + 1, MEAL_COL.NOTE - MEAL_COL.ORDER + 1).getValues();
    vals.forEach(function (r) {
      var name = trimStr_(r[MEAL_COL.NAME - 1]);
      if (!name) return;
      rows.push({
        position: r[MEAL_COL.POSITION - 1],
        name: name,
        days: Number(r[MEAL_COL.TOTAL_DAYS - 1]) || 0,
        unitPrice: Number(r[MEAL_COL.UNIT_PRICE - 1]) || 0,
        amount: Number(r[MEAL_COL.AMOUNT - 1]) || 0,
        note: r[MEAL_COL.NOTE - 1]
      });
    });
  }

  var maxRows = personal.getMaxRows();
  var clearRows = Math.max(maxRows - PERSONAL_DATA_START_ROW + 1, rows.length + 1);
  try { personal.getRange(PERSONAL_DATA_START_ROW, 1, clearRows, 7).breakApart(); } catch (e) {}
  personal.getRange(PERSONAL_DATA_START_ROW, 1, clearRows, 7).clearContent();

  var out = rows.map(function (r, idx) {
    return [idx + 1, r.position, r.name, r.days, r.unitPrice, r.amount, r.note];
  });
  if (out.length > 0) {
    personal.getRange(PERSONAL_DATA_START_ROW, 1, out.length, 7).setValues(out);
    personal.getRange(PERSONAL_DATA_START_ROW, PERSONAL_COL.AMOUNT, out.length, 1).setNumberFormat('#,##0');
  }

  var totalRow = PERSONAL_DATA_START_ROW + out.length;
  var totalDays = rows.reduce(function (s, r) { return s + r.days; }, 0);
  var totalAmount = rows.reduce(function (s, r) { return s + r.amount; }, 0);
  personal.getRange(totalRow, 1).setValue('교직원 계');
  personal.getRange(totalRow, 1, 1, 3).merge();
  personal.getRange(totalRow, PERSONAL_COL.DAYS).setValue(totalDays);
  personal.getRange(totalRow, PERSONAL_COL.AMOUNT).setValue(totalAmount);
  personal.getRange(totalRow, PERSONAL_COL.AMOUNT).setNumberFormat('#,##0');

  return ok_({ rows: out.length, totalDays: totalDays, totalAmount: totalAmount });
}
