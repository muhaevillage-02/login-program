// ============================================================
//  LOGIN LAB — 도전스토리 공모전 참가 신청 (challenge-story.html)
//  시트명: 도전스토리 신청자
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
//  6. 배포 후 뜨는 웹 앱 URL 복사 → challenge-story.html 의 APPS_SCRIPT_URL 값에 붙여넣기
// ============================================================

var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
var SHEET_NAME     = '도전스토리 신청자';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      var headers = [
        '제출 일시',
        '이름', '출생연도', '성별',
        '거주 지역', '거주 시/군/구',
        '직업', '대학명',
        '전화번호', '이메일',
        '신청계기',
        '도전 스토리',
        '어려움 카테고리', '앞으로의 어려움',
        '개인정보 동의', 'User-Agent'
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight('bold')
           .setBackground('#FFE600')
           .setFontColor('#1F1B0C');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 150);
      sheet.setColumnWidth(12, 320);
      sheet.setColumnWidth(14, 320);
    }

    sheet.appendRow([
      data.submittedAt ? new Date(data.submittedAt) : new Date(),
      data.name         || '',
      data.birthYear     || '',
      data.gender          || '',
      data.region            || '',
      data.district            || '',
      data.job                   || '',
      data.university              || '',
      data.phone                     || '',
      data.email                       || '',
      data.motivation                    || '',
      data.story                           || '',
      data.difficultyCategories              || '',
      data.difficultyStory                     || '',
      data.agreePrivacy                          || '',
      data.userAgent                               || ''
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
    .createTextOutput('도전스토리 공모전 신청 엔드포인트 정상 동작 중.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function testConnection() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  Logger.log('연결 성공: ' + sheet.getName() + ' / 현재 행 수: ' + sheet.getLastRow());
}
