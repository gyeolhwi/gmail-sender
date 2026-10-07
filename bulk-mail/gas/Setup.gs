/**
 * 설치 — 없는 시트만 만들고 예시 데이터를 넣는다. 이미 있는 시트는 건드리지 않는다.
 */
function setup() {
  var made = [];
  if (makeSheet_(SHEETS.CONFIG, ['항목', '값', '설명'], [
    ['보내는사람_이름', '', '받는 사람에게 보이는 이름 (비우면 계정 기본값)'],
    ['첨부파일_링크', '', 'Drive 파일 링크 — 모든 메일에 첨부 (비우면 첨부 없음)'],
    ['예약_일시', '', '예: 2026-10-05 09:00 → 메뉴 ⏰ 예약하기 (비우고 누르면 예약 취소)'],
    ['예약_템플릿', '기본안내', '예약 발송에 쓸 템플릿 이름']
  ])) made.push(SHEETS.CONFIG);

  if (makeSheet_(SHEETS.CUSTOMER, [COL.CHECK, COL.EMAIL, '이름', '웹사이트', COL.STATUS, COL.SENT_AT], [
    [false, 'customer1@example.com', '홍길동', 'example.com', '', ''],
    [false, 'customer2@example.com', '', 'sample.co.kr', '', '']
  ])) {
    sheet_(SHEETS.CUSTOMER).getRange('A2:A1000').insertCheckboxes();
    made.push(SHEETS.CUSTOMER);
  }

  if (makeSheet_(SHEETS.TEMPLATE, ['템플릿명', '제목', '본문'], [
    ['기본안내', '[안내] {{이름|웹사이트}}님께 알려드립니다',
      '안녕하세요, {{이름|웹사이트}}님.\n\n안내드릴 내용을 여기에 작성하세요.\n\n감사합니다.']
  ])) {
    sheet_(SHEETS.TEMPLATE).getRange('C:C').setWrap(true);
    sheet_(SHEETS.TEMPLATE).setColumnWidth(3, 500);
    made.push(SHEETS.TEMPLATE);
  }

  if (makeSheet_(SHEETS.LOG, ['일시', '이메일', '템플릿', '상태', '오류'], [])) made.push(SHEETS.LOG);

  alert_(made.length ? '설치 완료 ✅ 생성: ' + made.join(', ') +
    '\n\n1) [고객]에 이메일 입력 (열은 자유롭게 추가 → {{열이름}}으로 사용)\n2) [템플릿] 작성\n3) 체크 → 👀 미리보기 → 📨 발송'
    : '이미 설치되어 있습니다.');
}

// 시트가 없을 때만 헤더+예시로 생성. 생성했으면 true
function makeSheet_(name, head, rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(name)) return false;
  var sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  sh.setFrozenRows(1);
  return true;
}
