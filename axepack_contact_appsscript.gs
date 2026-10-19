/**
 * AXEpack 상담 폼 수신 — Google Apps Script 웹앱 (구글 시트 기록 + 알림 메일)
 *
 * 배포 절차 (jewook@axeworks.xyz 계정)
 *  1. 구글 시트를 새로 만든다. 예: "AXEpack 상담 문의"
 *  2. 확장 프로그램 → Apps Script → 이 파일 내용을 Code.gs에 붙여넣고 저장
 *  3. 함수 선택 창에서 setup 선택 → 실행 → 권한 승인 (한 번만)
 *  4. 배포 → 새 배포 → 유형 "웹 앱"
 *       실행 계정: 나 / 액세스 권한이 있는 사용자: 모든 사용자
 *     → 배포 → 웹 앱 URL(.../exec)을 복사해서 전달
 *  5. 코드를 고친 뒤에는 배포 → 배포 관리 → 편집 → 새 버전으로 올려야 같은 URL에 반영된다.
 *
 * 페이지에는 키나 토큰이 없다. 스크립트가 소유자 권한으로 시트와 메일을 다룬다.
 */

const NOTIFY_TO = 'hello@axeworks.xyz,jewook@axeworks.xyz';
const SHEET_NAME = '문의';
const MAX_LEN = 2000;
const TYPES = {
  env: 'AI 환경 구축 검토',
  cap: '조직의 AI 활용 역량 개선',
  poc: '기존 PoC의 운영 전환',
  phy: '제조 현장 피지컬 AX 검토',
};
const HEADERS = ['접수시각', '검토 내용', '대상 업무', '추진 단계', '현재 환경', '회사명', '이름', '이메일', '언어', '보던 섹션', '준비도', '참고 사례', '둘러본 경로', '페이지', 'UTM', '유입'];

/** 편집기에서 한 번 실행: 이 시트의 ID를 저장하고 '문의' 탭과 헤더를 만든다. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d || typeof d !== 'object') return reply({ ok: false, error: 'bad_request' });

    // 허니팟: 사람은 비워 두는 칸. 채워져 있으면 봇이므로 저장하지 않고 성공처럼 응답한다.
    if (d.hp_check) return reply({ ok: true });

    const v = k => (typeof d[k] === 'string' ? d[k].trim() : '');
    const required = ['company', 'name', 'email', 'work'];
    const tooLong = Object.keys(d).some(k => typeof d[k] === 'string' && d[k].length > MAX_LEN);
    const valid = required.every(k => v(k)) && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v('email')) && d.agree === 'on' && Object.prototype.hasOwnProperty.call(TYPES, v('type')) && !tooLong;
    if (!valid) return reply({ ok: false, error: 'invalid' });

    const row = [
      Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
      TYPES[v('type')], v('work'), v('stage'), v('env'),
      v('company'), v('name'), v('email'), v('lang'),
      v('context'), v('ready'), v('case'), v('journey'), v('page'), v('utm'), v('referrer'),
    ];

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getSheetByName(SHEET_NAME);
      const range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, HEADERS.length);
      // 일반 텍스트 서식으로 먼저 지정해 두면 '=' 로 시작하는 입력도 수식으로 해석되지 않는다.
      range.setNumberFormat('@');
      range.setValues([row]);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

    // 알림 메일은 실패해도 접수 자체는 성공으로 본다. 실패 내용은 Apps Script '실행' 로그에서 확인한다.
    try {
      const body = HEADERS.map((h, i) => h + ': ' + row[i]).join('\n');
      MailApp.sendEmail(NOTIFY_TO, '[AXEpack 상담] ' + row[5].replace(/[\r\n]+/g, ' ') + ' · ' + row[6].replace(/[\r\n]+/g, ' '), body, { replyTo: v('email'), name: 'AXEworks 상담 폼' });
    } catch (mailErr) {
      Logger.log('mail failed: ' + mailErr);
    }
    return reply({ ok: true });
  } catch (err) {
    Logger.log('doPost failed: ' + err);
    return reply({ ok: false, error: 'server' });
  }
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
