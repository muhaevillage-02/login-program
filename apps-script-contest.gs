// ============================================================
//  LOGIN LAB — 대학생 창업 아이디어 공모전 신청서
//  시트명: LL_공모전신청
// ============================================================
//
//  [배포 방법]
//  1. Google Sheet 새로 만들기 → URL에서 ID 복사
//     https://docs.google.com/spreadsheets/d/【ID】/edit
//  2. script.google.com → 새 프로젝트 → Code.gs에 이 파일 전체 붙여넣기
//  3. SPREADSHEET_ID 에 복사한 ID 입력 후 저장
//  4. 배포 > 새 배포 > 웹앱
//     실행 계정: 나 / 액세스: 모든 사용자(익명 포함)
//  5. 웹앱 URL 복사 → contest.html 의 GAS_URL 변수에 붙여넣기
// ============================================================

var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
var SHEET_NAME     = 'LL_공모전신청';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      var headers = [
        '제출 일시',
        '이름', '생년월일', '신분', '대학교/학과',
        '거주지 (시/도)', '거주지 상세', '전화번호', '이메일',
        '지원 부문',
        'Q1. 청년들이 겪는 문제',
        'Q2. 핵심 솔루션 아이디어',
        'Q3. 실행 방안',
        '포트폴리오 링크',
        '개인정보 동의', '콘텐츠 활용 동의', '뉴스레터 수신',
        'User-Agent', '타임스탬프'
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight('bold')
           .setBackground('#C8E600');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 150);
      sheet.setColumnWidth(11, 300);
      sheet.setColumnWidth(12, 300);
      sheet.setColumnWidth(13, 300);
    }

    sheet.appendRow([
      new Date(),
      data.name             || '',
      data.birth            || '',
      data.status           || '',
      data.university       || '',
      data.region           || '',
      data.residence_detail || '',
      data.phone            || '',
      data.email            || '',
      data.category         || '',
      data.q1               || '',
      data.q2               || '',
      data.idea             || '',
      data.portfolio_link   || '',
      data.agree_privacy    ? 'O' : 'X',
      data.agree_use        ? 'O' : 'X',
      data.agree_marketing  ? 'O' : 'X',
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
    .createTextOutput('LOGIN LAB 공모전 신청서 엔드포인트 정상 동작 중.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function testConnection() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  Logger.log('연결 성공: ' + sheet.getName() + ' / 현재 행 수: ' + sheet.getLastRow());
}
