// ============================================================
//  LOG.in — 공용 신청서 접수 스크립트
//  하나의 웹앱 배포로 여러 폼을 처리, data.form_type 값에 따라
//  같은 스프레드시트 내 서로 다른 탭에 저장합니다.
//
//  - levelup.html          → form_type 없음 또는 'levelup'  → 시트 'LL_레벨업신청'
//  - story-contest.html    → form_type: 'story_contest'     → 시트 '도전 스토리'
//  - march-letter-form.html→ form_type: 'essay_contest'     → 시트 '대학생 에세이 공모전'
//  - major-worry.html      → form_type: 'major_worry'       → 시트 '2026 대학생 지원 프로그램'
// ============================================================
//
//  [배포 방법 — 기존에 levelup용으로 배포한 프로젝트를 그대로 사용]
//  1. script.google.com 에서 levelup 배포에 쓰던 그 프로젝트를 엽니다.
//  2. Code.gs 내용을 이 파일 전체로 교체 후 저장.
//  3. 배포 > 배포 관리 > 기존 배포 옆 연필(✏) 아이콘 클릭
//     → 버전: 새 버전 → 배포
//     (※ '새 배포'가 아니라 '기존 배포 수정'이어야 웹앱 URL이 그대로 유지됩니다)
//  4. 웹앱 URL은 기존과 동일하므로 story-contest.html 쪽 GAS_URL도
//     levelup.html과 같은 값으로 맞추면 됩니다 (이미 반영되어 있음).
// ============================================================

var SPREADSHEET_ID = '1054PSBS0-wxBPydMqiykXk6IUjrCmh5QqPxbf9vzaTU';

var FORMS = {
  levelup: {
    sheetName: 'LL_레벨업신청',
    headers: [
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
    ],
    toRow: function (data) {
      var statusVal = data.status || '';
      if (statusVal === 'E: 기타' && data.status_other) {
        statusVal = '기타: ' + data.status_other;
      }
      return [
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
        data.history          || '',
        data.motivation       || '',
        data.support           || '',
        data.agree_privacy ? 'O' : 'X',
        data.userAgent        || '',
        data.timestamp        || ''
      ];
    }
  },

  story_contest: {
    sheetName: '도전 스토리',
    headers: [
      '제출 일시',
      '이름', '만 나이', '신분',
      '거주지 (시/도)', '거주지 상세', '전화번호', '이메일',
      '도전 분야',
      '도전 스토리',
      '개인정보 동의',
      'User-Agent', '타임스탬프'
    ],
    toRow: function (data) {
      return [
        new Date(),
        data.name             || '',
        data.age              || '',
        data.student_status   || '',
        data.region           || '',
        data.residence_detail || '',
        data.phone            || '',
        data.email            || '',
        data.story_category   || '',
        data.story            || '',
        data.agree_privacy ? 'O' : 'X',
        data.userAgent        || '',
        data.timestamp        || ''
      ];
    }
  },

  essay_contest: {
    sheetName: '대학생 에세이 공모전',
    headers: [
      '제출 일시',
      '이름', '나이', '학교 · 학년', '전공', '연락처',
      '거주 시/도', '거주 구/군',
      '고민 키워드', '갓생 점수',
      '작년의 나에게 (에세이)', '지금 가장 큰 고민', '공개 여부',
      'User-Agent', '타임스탬프'
    ],
    toRow: function (data) {
      return [
        new Date(),
        data.name    || '',
        data.age     || '',
        data.school  || '',
        data.major   || '',
        data.phone   || '',
        data.region  || '',
        data.gu      || '',
        data.keywords || '',
        data.score   || '',
        data.letter  || '',
        data.secret  || '',
        data.isAnon  || '',
        data.userAgent || '',
        data.timestamp  || ''
      ];
    }
  },

  major_worry: {
    sheetName: '2026 대학생 지원 프로그램',
    headers: [
      '제출 일시',
      '이름', '만 나이', '신분',
      '거주지 (시/도)', '거주지 상세', '전화번호', '이메일',
      '학교 · 학년', '전공',
      '고민 유형', '고민이 시작된 상황', '구체적으로 힘든 점', '듣고 싶은 조언',
      '개인정보 동의',
      'User-Agent', '타임스탬프'
    ],
    toRow: function (data) {
      return [
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
        data.worry_advice        || '',
        data.agree_privacy ? 'O' : 'X',
        data.userAgent          || '',
        data.timestamp          || ''
      ];
    }
  }
};

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var formKey = FORMS[data.form_type] ? data.form_type : 'levelup';
    var form = FORMS[formKey];

    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(form.sheetName) || ss.insertSheet(form.sheetName);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(form.headers);
      sheet.getRange(1, 1, 1, form.headers.length)
           .setFontWeight('bold')
           .setBackground('#00CCFA');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 150);
    }

    sheet.appendRow(form.toRow(data));

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
    .createTextOutput('LOG.in 공용 신청서 접수 엔드포인트 정상 동작 중.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function testConnection() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.keys(FORMS).forEach(function (key) {
    var form = FORMS[key];
    var sheet = ss.getSheetByName(form.sheetName) || ss.insertSheet(form.sheetName);
    Logger.log(key + ' 연결 성공: ' + sheet.getName() + ' / 현재 행 수: ' + sheet.getLastRow());
  });
}
