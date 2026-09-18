// ============================================================
//  LOGIN LAB — 나에게 딱 맞는 전공·진로 찾기 테스트 (전공.html)
//  시트명: 전공테스트 신청자
// ============================================================
//
//  [배포 방법 — 새 시트 + 새 스크립트로 독립 운영]
//  1. https://sheets.new 에서 새 구글 시트 생성 → URL에서 ID 복사
//     https://docs.google.com/spreadsheets/d/【여기가 ID】/edit
//  2. 그 시트에서 상단 메뉴 확장 프로그램 > Apps Script 클릭 (새 프로젝트가 열립니다)
//  3. Code.gs 내용을 전부 지우고 이 파일 전체를 붙여넣기
//  4. 아래 SPREADSHEET_ID 자리에 1번에서 복사한 ID 붙여넣고 저장 (Cmd+S)
//  5. 배포 > 새 배포 > 유형: 웹 앱
//     - 실행 계정: 나
//     - 액세스 권한: 모든 사용자 (익명 포함) ← 반드시 이걸로! "Google 계정이 있는 사용자" 아님
//  6. 배포 후 뜨는 웹 앱 URL 복사 → 전공.html 의 CONFIG.SUBMIT_ENDPOINT 값에 붙여넣기
// ============================================================

var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
var SHEET_NAME     = '전공테스트 신청자';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      var headers = [
        '제출 일시',
        '이름', '나이', '연락처',
        '거주 지역', '거주 시/군/구', '현재 상황',
        '고민되는 것', '신청계기',
        '테스트 결과', '결과 키',
        '문항별 응답'
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight('bold')
           .setBackground('#0A9E78')
           .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 150);
      sheet.setColumnWidth(8, 300);
      sheet.setColumnWidth(12, 260);
    }

    sheet.appendRow([
      data.submittedAt ? new Date(data.submittedAt) : new Date(),
      data.name       || '',
      data.age         || '',
      data.phone        || '',
      data.region        || '',
      data.district       || '',
      data.status          || '',
      data.concern          || '',
      data.source            || '',
      data.resultType        || '',
      data.resultKey          || '',
      Array.isArray(data.answers) ? data.answers.join(', ') : ''
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
    .createTextOutput('나에게 딱 맞는 전공·진로 찾기 테스트 엔드포인트 정상 동작 중.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function testConnection() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  Logger.log('연결 성공: ' + sheet.getName() + ' / 현재 행 수: ' + sheet.getLastRow());
}
