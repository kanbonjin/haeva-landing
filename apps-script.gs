/**
 * 해바 사전예약 → 구글 시트 저장
 * script.google.com 독립 프로젝트. setup() 한 번 실행(권한 허용) 후
 * 배포 > 새 배포 > 웹 앱 (실행: 나 / 액세스: 모든 사용자) 으로 배포.
 */
const SPREADSHEET_ID = '1xb0i4BBJnSBQBGN-1lgnvEJl8uINn1y4tX3B53PxETo';
const SHEET_NAME = '신청';
const HEADER = ['신청 시각', '이메일', '관심 취미', '유입 경로'];

// 처음 한 번 실행: 파일 이름 정리 + 머리글 만들기 (권한 허용 창이 뜸)
function setup() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ss.rename('해바 사전예약');
  ss.setSpreadsheetTimeZone('Asia/Seoul'); // 신청 시각을 한국 시간으로 표시
  getSheet();
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const email = String(data.email || '').trim().toLowerCase().slice(0, 200);
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return reply({ ok: false, reason: 'email' });

    const sheet = getSheet();
    // 같은 이메일이 이미 있으면 다시 쌓지 않음 (중복 신청 방지)
    const dup = sheet.getRange('B:B').createTextFinder(email).matchEntireCell(true).findNext();
    if (dup) return reply({ ok: true, dup: true });

    sheet.appendRow([
      new Date(),
      safe(email),
      safe(String(data.hobbies || '').slice(0, 200)),
      safe(String(data.ref || '').slice(0, 300)),
    ]);
    return reply({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADER);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// = + - @ 로 시작하는 값이 수식으로 실행되지 않게 막음
function safe(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
