// ============================================================
//  LOG.in — 청년 레벨업 지원 사업 신청서
//  시트명: LL_레벨업신청
// ============================================================
//
//  [배포 방법]
//  1. Google Sheet 새로 만들기 → URL에서 ID 복사
//     https://docs.google.com/spreadsheets/d/【ID】/edit
//  2. script.google.com → 새 프로젝트 → Code.gs에 이 파일 전체 붙여넣기
//  3. SPREADSHEET_ID 에 복사한 ID 입력 후 저장
//  4. 배포 > 새 배포 > 웹앱
//     실행 계정: 나 / 액세스: 모든 사용자(익명 포함)
//  5. 웹앱 URL 복사 → levelup.html 의 GAS_URL 변수에 붙여넣기
// ============================================================

var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
var SHEET_NAME     = 'LL_레벨업신청';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      var headers = [
        '제출 일시',
        '이름', '만 나이', '신분',
        '거주지 (시/도)', '거주지 상세', '전화번호', '이메일',
        '현재 상태', '기타 상태 입력',
        '희망 직무 / 관심 분야', '전공',
        '준비 이력',
        '참여 동기',
        '필요한 지원 항목',
        '개인정보 동의',
        'User-Agent', '타임스탬프'
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight('bold')
           .setBackground('#2BFF7E');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 150);
      sheet.setColumnWidth(13, 280);
      sheet.setColumnWidth(14, 280);
    }

    var statusVal = data.status || '';
    if (statusVal === 'E: 기타' && data.status_other) {
      statusVal = '기타: ' + data.status_other;
    }

    sheet.appendRow([
      new Date(),
      data.name             || '',
      data.age              || '',
      data.student_status   || '',
      data.region           || '',
      data.residence_detail || '',
      data.phone            || '',
      data.email            || '',
      statusVal,
      data.status_other     || '',
      data.job_target       || '',
      data.edu              || '',
      data.history           || '',
      data.motivation        || '',
      data.support            || '',
      data.agree_privacy ? 'O' : 'X',
      data.userAgent        || '',
      data.timestamp        || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput('LOG.in 청년 레벨업 지원 사업 신청서 엔드포인트 정상 동작 중.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function testConnection() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  Logger.log('연결 성공: ' + sheet.getName() + ' / 현재 행 수: ' + sheet.getLastRow());
}
