/**
 * 날짜/문자열/응답 포맷 공용 헬퍼
 */

function pad2_(n) { return (n < 10 ? '0' : '') + n; }

function toDateKey_(y, m, d) {
  return y + '-' + pad2_(m) + '-' + pad2_(d);
}

function parseDateKey_(key) {
  var parts = String(key).split('-');
  return { y: Number(parts[0]), m: Number(parts[1]), d: Number(parts[2]) };
}

var WEEKDAY_LABEL_ = ['일', '월', '화', '수', '목', '금', '토'];

function dayLabel_(dateKey) {
  var p = parseDateKey_(dateKey);
  var date = new Date(p.y, p.m - 1, p.d);
  var w = WEEKDAY_LABEL_[date.getDay()];
  return p.m + '/' + p.d + '(' + w + ')';
}

function isWeekend_(y, m, d) {
  var day = new Date(y, m - 1, d).getDay();
  return day === 0 || day === 6;
}

function lastDayOfMonth_(y, m) {
  return new Date(y, m, 0).getDate();
}

function formatWon_(n) {
  return Number(n || 0).toLocaleString('ko-KR') + '원';
}

function trimStr_(v) {
  return (v === null || v === undefined) ? '' : String(v).trim();
}

function isValidEmail_(v) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(trimStr_(v));
}

// 비고 등 사용자 입력이 시트 수식으로 해석되지 않도록 이스케이프 + 길이 제한
function escapeCell_(text) {
  var s = trimStr_(text);
  if (s.length > 200) s = s.substring(0, 200);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function ok_(data) {
  return { ok: true, code: 'OK', message: '', data: data || {} };
}

function fail_(code, message, data) {
  return { ok: false, code: code, message: message || '', data: data || {} };
}

function nowStamp_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

function colLetter_(col) {
  var letter = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

function getViewerId_() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email) return email;
  } catch (e) {}
  try {
    return Session.getTemporaryActiveUserKey();
  } catch (e2) {
    return 'unknown';
  }
}
