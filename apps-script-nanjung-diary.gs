/**
 * 난중일기 공모전 신청서 → 구글 시트 자동 연동용 Apps Script
 *
 * 헤더 행을 직접 입력할 필요 없음 — 첫 응답이 도착하면 스크립트가 알아서 만든다.
 *
 * 사용법 (총 3단계)
 * 1) 새 구글 스프레드시트를 하나 연다 (sheets.new).
 * 2) 상단 메뉴 [확장 프로그램] > [Apps Script] 를 열고, 기본 코드를 지운 뒤
 *    이 파일 내용 전체를 붙여넣고 저장한다.
 *    이어서 우측 상단 [배포] > [새 배포] > 유형(톱니바퀴)에서 [웹 앱] 선택 →
 *    실행할 사용자: 나 / 액세스 권한: 전체 → [배포].
 *    권한 승인 화면이 뜨면 본인 계정으로 승인 (경고가 뜨면 [고급] > [이동]).
 * 3) 배포 후 나오는 "웹 앱 URL"을 신청서 코드로 보내주면, 이후부터는
 *    신청서가 제출될 때마다 이 시트에 자동으로 한 줄씩 쌓인다.
 *
 * 이 3단계는 구글 계정 소유자만 할 수 있는 인증/배포 절차라 본인이 직접 진행해야 하며,
 * 그 외 나머지(헤더 생성, 데이터 기록, 서식)는 전부 스크립트가 자동으로 처리한다.
 *
 * 코드를 수정한 뒤에는 [배포] > [배포 관리] > 연필 아이콘 > 버전: 새 버전 > [배포]
 * 를 눌러야 같은 URL로 변경 사항이 반영된다.
 */
var HEADERS = [
  "접수시각", "이름", "출생연도", "성별", "지역", "시군구", "상태", "대학명",
  "전화번호", "이메일", "신청계기", "일기제목", "언제이야기", "마음날씨", "먹구름",
  "일기내용", "글자수", "나를지켜준우산", "개인정보동의", "제출시각(클라이언트)", "User-Agent"
];

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  var data = JSON.parse(e.postData.contents);

  sheet.appendRow([
    new Date(),
    data.name || "",
    data.birthYear || "",
    data.gender || "",
    data.region || "",
    data.district || "",
    data.job || "",
    data.university || "",
    data.phone || "",
    data.email || "",
    data.motivation || "",
    data.diaryTitle || "",
    data.diaryWhen || "",
    data.mood || "",
    data.clouds || "",
    data.diary || "",
    data.diaryLength || "",
    data.umbrella || "",
    data.agreePrivacy || "",
    data.submittedAt || "",
    data.userAgent || ""
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ result: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}
