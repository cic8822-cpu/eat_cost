/**
 * 안내 메일 발송 (MailApp). 계좌/예금주는 Script Properties, 나머지는 설정 시트에서 읽음.
 * 발송 실패해도 저장 결과에는 영향을 주지 않는다(saveSubmission이 상태만 기록).
 */

function buildMailBody_(emp, result) {
  var settings = getSettingsMap_();
  var bank = getScriptProp_(BANK_ACCOUNT_KEY) || '(계좌 미입력)';
  var holder = getScriptProp_(ACCOUNT_HOLDER_KEY) || '(예금주 미입력)';
  var bankName = settings.BANK_NAME || '';
  var deadline = settings.PAYMENT_DEADLINE || '(입금기한 미입력)';
  var schoolName = settings.SCHOOL_NAME || '';

  var lines = [];
  lines.push(emp.name + '님, 안녕하세요.');
  lines.push('');
  if (result.days > 0) {
    lines.push(result.year + '년 ' + result.month + '월 급식 신청이 접수되었습니다.');
    lines.push('- 급식일수: ' + result.days + '일');
    lines.push('- 급식단가: ' + formatWon_(result.unitPrice));
    lines.push('- 급식비: ' + formatWon_(result.amount));
    lines.push('');
    lines.push('아래 계좌로 ' + deadline + '까지 입금해 주시기 바랍니다.');
    lines.push('- 입금계좌: ' + bankName + ' ' + bank + ' (예금주: ' + holder + ')');
  } else {
    lines.push(result.year + '년 ' + result.month + '월은 급식 신청 내역이 없습니다(0일).');
  }
  lines.push('');
  lines.push('감사합니다.');
  if (schoolName) lines.push(schoolName + ' 급식 담당자 드림');

  return lines.join('\n');
}

function sendResultMail_(emp, result) {
  if (!emp.email || !isValidEmail_(emp.email)) return { status: MAIL_STATUS.NO_EMAIL };

  var quota = MailApp.getRemainingDailyQuota();
  if (quota <= 0) return { status: MAIL_STATUS.QUOTA };

  var settings = getSettingsMap_();
  var subjectTemplate = settings.MAIL_SUBJECT_TEMPLATE || '[학교급식] {년도}년 {월}월 급식 신청 및 급식비 안내';
  var subject = subjectTemplate.replace('{년도}', result.year).replace('{월}', result.month);
  var body = buildMailBody_(emp, result);

  try {
    MailApp.sendEmail(emp.email, subject, body);
    return { status: MAIL_STATUS.SENT };
  } catch (err) {
    return { status: MAIL_STATUS.FAILED };
  }
}

// 관리자 메뉴: 본인에게 테스트 메일 발송
function sendTestMailToSelf() {
  var me = Session.getEffectiveUser().getEmail();
  var period = getPeriod_();
  MailApp.sendEmail(me, '[학교급식] 테스트 메일', buildMailBody_(
    { name: '테스트', email: me },
    { days: 10, unitPrice: period.unitPrice, amount: 10 * period.unitPrice, year: period.year, month: period.month }
  ));
  return ok_({ to: me });
}
