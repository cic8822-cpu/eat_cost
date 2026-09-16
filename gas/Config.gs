/**
 * 전역 설정 상수 - 시트명, 열 인덱스, 헤더 구조를 한 곳에서 관리합니다.
 * 시트 구조가 바뀌면 이 파일만 수정하면 됩니다.
 */

var SHEET_NAMES = {
  DATA: '데이터',
  MEAL: '학교급식',
  PERSONAL: '개인별내역',
  SETTINGS: '설정',
  LOG: '제출로그'
};

// 데이터 시트 열 (1-base)
var DATA_COL = {
  ORDER: 1,     // A 순
  POSITION: 2,  // B 직급
  NAME: 3,      // C 성명
  EMAIL: 4,     // D 이메일
  EMP_ID: 5,    // E 직원ID
  ACTIVE: 6     // F 재직여부(Y/N)
};

// 학교급식 시트 열 (1-base)
var MEAL_COL = {
  ORDER: 1,
  POSITION: 2,
  NAME: 3,
  EMAIL: 4,
  TOTAL_DAYS: 5,
  UNIT_PRICE: 6,
  AMOUNT: 7,
  DAY_START: 8,
  DAY_COUNT: 31,
  SUBMITTED_AT: 39,
  NOTE: 40,
  MAIL_STATUS: 41,
  EMP_ID: 42
};

var MEAL_HEADER_ROW = 2;
var MEAL_DATA_START_ROW = 3;

var MEAL_YEAR_CELL = 'B1';
var MEAL_MONTH_CELL = 'D1';
var MEAL_UNIT_PRICE_CELL = 'F1';

var PERSONAL_TITLE_ROW = 1;
var PERSONAL_HEADER_ROW = 3;
var PERSONAL_DATA_START_ROW = 4;

var PERSONAL_COL = {
  ORDER: 1, POSITION: 2, NAME: 3, DAYS: 4, UNIT_PRICE: 5, AMOUNT: 6, NOTE: 7
};

var SETTINGS_KV_RANGE = { COL_KEY: 1, COL_VALUE: 2, START_ROW: 2 };
var SETTINGS_DAY_RANGE = { COL_DATE: 4, COL_STATUS: 5, COL_REASON: 6, START_ROW: 2 };

var MAIL_STATUS = { SENT: 'SENT', NO_EMAIL: 'NO_EMAIL', FAILED: 'FAILED', QUOTA: 'QUOTA' };
var DAY_STATUS = { OK: '가능', BLOCK: '불가', CAUTION: '주의' };

var ADMIN_TOKEN_KEY = 'ADMIN_TOKEN';
var BANK_ACCOUNT_KEY = 'BANK_ACCOUNT';
var ACCOUNT_HOLDER_KEY = 'ACCOUNT_HOLDER';

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('SHEET_NOT_FOUND:' + name);
  return sheet;
}

function getScriptProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}

function setScriptProp_(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}
