// ============================================================
//  LOGIN LAB — 대학생 창업 1:1 동행 사업 신청서 응답 수집
// ============================================================
//
//  배포 방법:
//  1. 새 Google Sheet 를 만들고 URL 에서 스프레드시트 ID 복사
//     예) https://docs.google.com/spreadsheets/d/【여기가 ID】/edit
//  2. 아래 SPREADSHEET_ID 에 붙여넣기
//  3. https://script.google.com 에서 새 프로젝트 생성
//  4. 이 파일 전체를 Code.gs 에 붙여넣고 저장
//  5. 상단 메뉴 > 배포 > 새 배포
//     - 유형: 웹 앱
//     - 실행 계정: 나
//     - 액세스 권한: 모든 사용자 (익명 포함)
//  6. 배포 후 웹 앱 URL 복사
//  7. apply.html 의 GAS_URL 변수에 붙여넣기
// ============================================================

var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';  // ← 여기에 시트 ID 입력
var SHEET_NAME     = 'LL_창업지원신청';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // 헤더 자동 생성
    if (sheet.getLastRow() === 0) {
      var headers = [
        '제출 일시',
        '현재 상태',
        '참여 동기',
        '이름', '만 나이', '신분', '거주지', '거주지 상세', '전화번호',
        '희망 직무 / 관심 분야', '전공',
        '시도 이력',
        '희망 지원 유형',
        '개인정보 동의',
        'User-Agent', '타임스탬프'
      ];
      sheet.appendRow(headers);

      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold').setBackground('#C8E600');
      sheet.setFrozenRows(1);

      // 열 너비 조정
      sheet.setColumnWidth(1, 150);   // 제출 일시
      sheet.setColumnWidth(2, 180);   // 현재 상태
      sheet.setColumnWidth(3, 280);   // 참여 동기
      sheet.setColumnWidth(4, 80);    // 이름
      sheet.setColumnWidth(5, 60);    // 나이
      sheet.setColumnWidth(6, 120);   // 신분
      sheet.setColumnWidth(7, 100);   // 거주지
      sheet.setColumnWidth(8, 160);   // 거주지 상세
      sheet.setColumnWidth(9, 130);   // 전화번호
      sheet.setColumnWidth(10, 180);  // 희망 직무
      sheet.setColumnWidth(11, 150);  // 전공
      sheet.setColumnWidth(12, 280);  // 시도 이력
      sheet.setColumnWidth(13, 160);  // 지원 유형
    }

    // 상태 값 처리 (기타 선택 시 직접 입력값 우선)
    var statusValue = data.status || '';
    if (statusValue === 'E: 기타' && data.status_other) {
      statusValue = '기타: ' + data.status_other;
    }

    sheet.appendRow([
      new Date(),
      statusValue,
      data.motivation       || '',
      data.name             || '',
      data.age              || '',
      data.jobType          || '',
      data.region           || '',
      data.residence_detail || '',
      data.phone            || '',
      data.job_target       || '',
      data.edu              || '',
      data.history          || '',
      data.support          || '',
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
    .createTextOutput('LOGIN LAB 창업 지원 신청서 엔드포인트 정상 동작 중.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// Apps Script 에디터에서 직접 실행해 연결 확인
function testConnection() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  Logger.log('연결 성공: ' + sheet.getName() + ' / 현재 행 수: ' + sheet.getLastRow());
}
