// ============================================================
//  LOGIN LAB — 2026 대학생 지원 프로젝트 신청서 (career-coaching.html)
//  시트명: 2026 대학생 지원 프로젝트
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
//  6. 배포 후 뜨는 웹 앱 URL 복사 → career-coaching.html 의 GAS_URL 변수에 붙여넣기
// ============================================================

var SPREADSHEET_ID = '1SEY01okuRKKgiC1D2muL4W2Y-Ni6sO_n9KdgqGmDTiM';
var SHEET_NAME     = '2026 대학생 지원 프로젝트';

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
        '학교 · 학년', '전공',
        '고민 유형', '고민이 시작된 상황', '구체적으로 힘든 점', '기대되는 지원 항목',
        '개인정보 동의',
        'User-Agent', '타임스탬프'
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight('bold')
           .setBackground('#1FD1B0');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 150);
      sheet.setColumnWidth(12, 300);
      sheet.setColumnWidth(13, 300);
    }

    sheet.appendRow([
      new Date(),
      data.name             || '',
      data.age               || '',
      data.student_status    || '',
      data.region             || '',
      data.residence_detail  || '',
      data.phone              || '',
      data.email               || '',
      data.school              || '',
      data.major               || '',
      data.worry_type          || '',
      data.worry_situation     || '',
      data.worry_difficulty    || '',
      data.hope                || '',
      data.agree_privacy ? 'O' : 'X',
      data.userAgent          || '',
      data.timestamp          || ''
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
    .createTextOutput('2026 대학생 지원 프로젝트 신청서 엔드포인트 정상 동작 중.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function testConnection() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  Logger.log('연결 성공: ' + sheet.getName() + ' / 현재 행 수: ' + sheet.getLastRow());
}
