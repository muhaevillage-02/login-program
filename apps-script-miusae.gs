/**
 * 미우새 청년 성장 공모전 신청서 → 구글 시트 자동 연동용 Apps Script
 *
 * - 헤더 행은 첫 응답이 도착할 때 자동으로 만들어진다.
 * - 같은 연락처로 이미 접수된 건이 있으면 저장하지 않고 {code:'already'} 를 돌려준다.
 * - 필수 항목이 비어 있으면 {code:'invalid'} 를 돌려준다.
 * - 성공 시 {ok:true, receiptNo} 를 돌려준다. (신청서 화면이 이 응답을 기다린다)
 *
 * 코드를 수정한 뒤에는 [배포] > [배포 관리] > 연필 > 버전: 새 버전 > [배포] 를 눌러야
 * 같은 URL로 반영된다.
 */
var HEADERS = [
  "접수시각", "접수번호", "이름", "생년월일", "연락처", "거주지역(시·도)", "시·군·구", "읍·면·동",
  "지금의 나", "알게 된 경로", "개인정보동의", "작품 제목",
  "Q1 답변", "Q2 답변", "Q3 답변", "총 글자수", "Q1 글자수", "Q2 글자수", "Q3 글자수",
  "AI 미사용 확인", "비공개 확인", "수정불가 확인", "제출시각(KST)"
];

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    var a = d.answers || {};
    var c = d.confirmations || {};
    if (!d.name || !d.phone || !d.birth || !d.title || !a.q1 || !a.q2 || !a.q3 || !d.privacyConsent) {
      return out({ ok: false, code: "invalid" });
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    // 연락처 중복 확인 (5번째 열)
    var last = sheet.getLastRow();
    if (last > 1) {
      var phones = sheet.getRange(2, 5, last - 1, 1).getDisplayValues();
      for (var i = 0; i < phones.length; i++) {
        if (String(phones[i][0]).replace(/\D/g, "") === String(d.phone).replace(/\D/g, "")) {
          return out({ ok: false, code: "already" });
        }
      }
    }

    // 전화번호가 숫자로 변환되지 않도록 문자열로 저장
    var row = [
      new Date(), d.receiptNo || "", d.name, "'" + d.birth, "'" + d.phone,
      d.regionSido || "", d.regionSigungu || "", d.regionDong || "",
      d.status || "", d.source || "", d.privacyConsent ? "동의" : "", d.title,
      a.q1, a.q2, a.q3, d.charCount || "",
      (d.charCounts || {}).q1 || "", (d.charCounts || {}).q2 || "", (d.charCounts || {}).q3 || "",
      c.noAI ? "확인" : "", c.notPublic ? "확인" : "", c.noEditAfterSubmit ? "확인" : "",
      d.submittedAtKST || ""
    ];
    sheet.appendRow(row);
    return out({ ok: true, receiptNo: d.receiptNo });
  } catch (err) {
    return out({ ok: false, code: "error", message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 배포 주소를 브라우저로 열어 확인할 때 보이는 문구
function doGet() {
  return out({ ok: true, service: "miusae-apply" });
}
