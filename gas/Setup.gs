/**
 * 설치/초기화 — 데이터 시트 생성 + 샘플 시드. 양식은 원본(자료입력/세금계산서양식)을 그대로 사용.
 *  setupSheets()   : 최초 1회 전체 설치 (데이터 시트 초기화 — 신규 설치용)
 *  setupUpgrades() : 기존 데이터 그대로 두고 정기항목/드롭다운/금액수식 + 자료입력 숨김
 */
function setupSheets() {
  if (!formExists_()) { alertImportNeeded_(); return; }
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('전체 설치 (초기화)', '거래처·정기항목·거래내역이 샘플로 초기화됩니다.\n기존 입력 데이터가 있으면 [양식·경량화 갱신]을 쓰세요.\n\n그래도 전체 초기화할까요?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  seedConfig_();
  seedCustomers_();
  seedRecurring_();
  seedTransactions_();
  initLog_();
  unifyToOneSheet();   // 자료입력 → 세금계산서양식 한 시트로 통합
  reorderSheets_();
  alert_('설치 완료 ✅\n데이터 시트 생성 + 양식을 한 시트(세금계산서양식)로 통합했습니다.\n다음으로 [② POC 테스트]를 실행하세요.');
}

// 기존 데이터(거래처/거래내역)는 그대로 두고 양식 통합 + 경량화 기능만 적용
function setupUpgrades() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.FORM)) { alertImportNeeded_(); return; }
  if (!ss.getSheetByName(SHEETS.RECUR)) seedRecurring_();
  var txSh = ss.getSheetByName(SHEETS.TX);
  if (txSh) { applyCustomerDropdown_(txSh, 1); applyStatusDropdown_(txSh); setupTxFormulas_(txSh, 2, 299); }
  if (getConfig()['발송시각'] === undefined) setConfigValue_('발송시각', 9);  // 기존 설정에 없으면 추가
  applyConfigValidation_();   // 발송일/발송시각 등 입력 형식(드롭다운) 강제
  unifyToOneSheet();   // 자료입력 남아있으면 한 시트로 통합 (이미 통합됐으면 무시)
  reorderSheets_();
  alert_('갱신 완료 ✅\n· 한 시트(세금계산서양식)로 통합 — 자료입력 탭 제거\n· 정기항목 + 거래처ID 드롭다운 + 금액 자동계산\n\n원본 파란 양식 그대로 PDF가 출력됩니다.');
}

// 원본 양식(입력·출력 시트)이 워크북에 있는지 확인
function formExists_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return !!(ss.getSheetByName(SHEETS.INPUT) && ss.getSheetByName(SHEETS.FORM));
}
function alertImportNeeded_() {
  alert_('먼저 원본 양식을 가져오세요.\n\n파일 → 가져오기 → 업로드 → [거래명세서 양식.xlsx] 선택\n→ 가져오기 위치: "새 시트 삽입" → 가져오기\n\n( 자료입력 / 세금계산서양식 탭이 생긴 뒤 다시 실행하세요 )');
}

// 양식을 한 시트로 통합: 자료입력 값을 세금계산서양식 인쇄영역 밖(행+40)으로 옮기고
// 수식의 '자료입력'!XXn → XX(n+40) 재연결 후 자료입력 탭 삭제. 보이는 양식은 그대로.
function unifyToOneSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var form = ss.getSheetByName(SHEETS.FORM);
  var input = ss.getSheetByName(SHEETS.INPUT);
  if (!form) { alert_('세금계산서양식 시트가 없습니다. 먼저 양식을 가져오세요.'); return; }
  if (!input) return;   // 이미 통합됨 → 무시

  var SHIFT = ROW_SHIFT;   // 40 (Code.gs 상수)
  var need = 21 + SHIFT;
  if (form.getMaxRows() < need) form.insertRowsAfter(form.getMaxRows(), need - form.getMaxRows());

  // 1) 자료입력 입력영역(A1:I21) 값을 양식 시트의 (행+SHIFT) 위치로 복사
  var vals = input.getRange(1, 1, 21, 9).getValues();
  form.getRange(1 + SHIFT, 1, vals.length, vals[0].length).setValues(vals);

  // 2) 양식 본문 수식의 '자료입력' 참조를 같은 시트의 (행+SHIFT) 셀로 재연결 (해당 셀만)
  var fs = form.getRange(1, 1, 24, 33).getFormulas();   // A1:AG24
  for (var r = 0; r < fs.length; r++) {
    for (var c = 0; c < fs[r].length; c++) {
      var f = fs[r][c];
      if (f && f.indexOf('자료입력') >= 0) {
        form.getRange(r + 1, c + 1).setFormula(
          f.replace(/'?자료입력'?!\$?([A-Z]{1,3})\$?(\d+)/g, function (m, col, row) { return col + (parseInt(row, 10) + SHIFT); })
        );
      }
    }
  }

  // 3) 입력영역 숨김 + 자료입력 탭 삭제
  form.hideRows(1 + SHIFT, 21);
  ss.deleteSheet(input);
}

function createOrClear_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  else sh.clear();
  return sh;
}

function reorderSheets_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var order = [SHEETS.CONFIG, SHEETS.CUSTOMER, SHEETS.RECUR, SHEETS.TX, SHEETS.FORM, SHEETS.LOG];  // 자료입력은 통합으로 제거됨
    order.forEach(function (name, i) {
      var sh = ss.getSheetByName(name);
      if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
    });
  } catch (e) { /* 정렬 실패 무시 */ }
}

// 사용자 이메일을 추가 권한 없이 안전하게 조회(불가하면 빈 값)
function safeUserEmail_() {
  try { return Session.getEffectiveUser().getEmail() || ''; }
  catch (e1) { try { return Session.getActiveUser().getEmail() || ''; } catch (e2) { return ''; } }
}

// ----- 설정 -----
function seedConfig_() {
  var sh = createOrClear_(SHEETS.CONFIG);
  var body = '{담당자}님 안녕하세요, {공급자_상호}입니다.\n\n'
    + '{청구월} 거래명세서를 첨부드립니다.\n'
    + '- 공급가액: {공급가액}원\n- 부가세: {세액}원\n- 합계(청구금액): {발행금액}원\n\n'
    + '입금계좌: {공급자_입금계좌}\n\n감사합니다.\n{공급자_상호} 드림';
  var rows = [
    ['키', '값'],
    ['공급자_상호', '예시상사'],                       // ← 본인 회사명으로
    ['공급자_사업자번호', '000-00-00000'],              // ← 본인 사업자번호
    ['공급자_대표자', '홍길동'],                        // ← 대표자명
    ['공급자_주소', '서울특별시 ○○구 ○○대로 000'],     // ← 사업장 주소
    ['공급자_업태', '서비스'],
    ['공급자_종목', '소프트웨어 개발'],
    ['공급자_은행', '○○은행'],                          // ← 은행명 (양식 합계금액 우측 윗줄)
    ['공급자_계좌', '000-00-000000 홍길동'],             // ← 계좌번호 + 예금주 (양식 아랫줄)
    ['발송일', 10],
    ['발송시각', 9],
    ['청구월_기준', '당월'],
    ['발송_활성화', 'FALSE'],
    ['메일_제목', '[{공급자_상호}] {청구월} 거래명세서'],
    ['메일_본문', body],
    ['관리자_이메일', safeUserEmail_()],
    ['테스트_수신메일', 'test@example.com'],            // ← 본인 테스트 수신 메일
    ['보관폴더명', '거래명세서_보관']
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange('A1:B1').setFontWeight('bold').setBackground('#e8eef7');
  sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 460);
  sh.getRange(1, 1, rows.length, 2).setVerticalAlignment('top');
  applyConfigValidation_();
}

// 설정값 입력 형식 강제 (드롭다운 + 안내) — 발송일/발송시각/활성화/청구월
function applyConfigValidation_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
  if (!sh) return;
  var keys = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  function rowOf(k) { for (var i = 0; i < keys.length; i++) if (String(keys[i][0]).trim() === k) return i + 1; return 0; }
  function listRule(arr) { return SpreadsheetApp.newDataValidation().requireValueInList(arr, true).setAllowInvalid(false).build(); }
  var days = []; for (var d = 1; d <= 31; d++) days.push(d);
  var hours = []; for (var h = 0; h <= 23; h++) hours.push(h);
  var r;
  if ((r = rowOf('발송일'))) sh.getRange(r, 2).setDataValidation(listRule(days)).setNote('매월 며칠에 발송할지 (1~31). 목록에서만 선택 가능.');
  if ((r = rowOf('발송시각'))) sh.getRange(r, 2).setDataValidation(listRule(hours)).setNote('몇 시에 발송할지 (0~23, 24시간제).\n예: 9 = 오전 9시, 18 = 오후 6시, 23 = 밤 11시');
  if ((r = rowOf('발송_활성화'))) sh.getRange(r, 2).setDataValidation(listRule(['TRUE', 'FALSE'])).setNote('TRUE = 실제 발송 / FALSE = 생성만(보류)');
  if ((r = rowOf('청구월_기준'))) sh.getRange(r, 2).setDataValidation(listRule(['당월', '전월'])).setNote('당월 = 이번 달분 / 전월 = 지난 달분');
}

// ----- 거래처 (실제 22곳 + POC 테스터) -----
function seedCustomers_() {
  var sh = createOrClear_(SHEETS.CUSTOMER);
  var head = ['거래처ID', '상호', '이메일', '담당자', '사업자번호', '대표자', '주소', '업태', '종목', '파일명규칙', '발송여부', '비고'];
  // 예시 거래처 (본인 데이터로 교체하세요)
  var base = [
    ['C001', '예시거래처1', 'customer1@example.com', '담당자1'],
    ['C002', '예시거래처2', 'customer2@example.com', '담당자2'],
    ['C003', '예시거래처3', 'customer3@example.com', '담당자3']
  ];
  var rows = [head];
  base.forEach(function (b) { rows.push([b[0], b[1], b[2], b[3], '', '', '', '', '', '', true, '']); });
  // 테스터 (POC/테스트용 — 공급받는자 정보 전부 채운 예시)
  rows.push(['C999', '테스터상사', 'test@example.com', 'tester',
    '123-45-67890', '김테스터', '서울특별시 ○○구 ○○대로 000', '도소매', '전자상거래', '', true, 'POC 테스트용']);
  sh.getRange(1, 1, rows.length, head.length).setValues(rows);
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#e8eef7');
  sh.setFrozenRows(1);
  sh.setColumnWidth(2, 160); sh.setColumnWidth(3, 200);
}

// ----- 정기항목 (매달 청구할 내용 — 사용자가 주로 관리하는 시트) -----
function seedRecurring_() {
  var sh = createOrClear_(SHEETS.RECUR);
  var head = ['거래처ID', '상호(자동)', '품목', '규격', '수량', '단가', '사용여부'];
  var rows = [head,
    ['C001', '', '정기 호스팅', '월', 1, 30000, true],
    ['C002', '', '월 유지보수', '월', 1, 50000, true],
    ['C003', '', '월 사용료', '월', 1, 33000, true]
  ];
  sh.getRange(1, 1, rows.length, head.length).setValues(rows);
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#e8eef7');
  // 상호 자동표시 (거래처ID 입력하면 거래처 시트에서 상호를 끌어옴)
  var f = [];
  for (var r = 2; r <= 200; r++) {
    f.push(['=IF($A' + r + '="","",IFERROR(VLOOKUP($A' + r + ',거래처!$A:$B,2,0),"❓ID없음"))']);
  }
  sh.getRange(2, 2, 199, 1).setFormulas(f);
  sh.getRange(2, 6, 199, 1).setNumberFormat('#,##0');   // 단가(F열)
  sh.setColumnWidth(2, 160); sh.setColumnWidth(3, 150);
  sh.setFrozenRows(1);
  applyCustomerDropdown_(sh, 1);
  sh.getRange('A1').setNote('★ 매달 청구할 내용을 거래처별로 "한 번만" 적어두세요.\n\n여기 적어두면 [📧 이번 달 발송하기] 누를 때 명세서가 자동으로 만들어져 발송됩니다.\n\n- 거래처ID: 칸 클릭 → 목록에서 선택 (옆에 상호 자동표시)\n- 사용여부: TRUE = 매달 보냄 / FALSE = 이번엔 제외');
}

// ----- 거래내역 (샘플: 입력열만, 금액은 수식 자동계산) -----
function seedTransactions_() {
  var sh = createOrClear_(SHEETS.TX);
  var head = ['거래처ID', '청구월', '거래일자', '품목', '규격', '수량', '단가', '공급가액', '세액', '발행금액', '발송상태'];
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e8eef7');
  var m = currentBillMonth_(getConfigSafe_());
  var samples = [
    ['C999', m, m + '-10', '테스트 상품 A', 'EA', 2, 15000],
    ['C999', m, m + '-10', '테스트 상품 B', '식', 1, 40000],
    ['C001', m, m + '-15', '예시 품목 A', '', 1, 25000],
    ['C002', m, m + '-15', '예시 품목 B', '월', 1, 50000]
  ];
  sh.getRange(2, 1, samples.length, 7).setValues(samples);
  sh.getRange(2, 11, samples.length, 1).setValues(samples.map(function () { return ['미발송']; }));
  setupTxFormulas_(sh, 2, 299);
  sh.getRange(2, 6, 299, 5).setNumberFormat('#,##0');
  sh.setColumnWidth(4, 140);
  sh.setFrozenRows(1);
  applyCustomerDropdown_(sh, 1);
  applyStatusDropdown_(sh);
  sh.getRange('A1').setNote('보통 직접 안 건드려도 됩니다.\n[📧 이번 달 발송하기]를 누르면 정기항목 기준으로 여기 자동으로 채워집니다.\n일회성(이번 달만) 청구가 있으면 그때만 한 줄 직접 추가하세요.');
}

// 발송상태 드롭다운 (미발송/보류) — 특정 행을 "보류"로 두면 발송 제외
function applyStatusDropdown_(sh) {
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['미발송', '보류'], true).setAllowInvalid(true).build();
  sh.getRange(2, 11, 500, 1).setDataValidation(rule);
  sh.getRange('K1').setNote('미발송 = 이번에 보냄\n보류 = 이번엔 안 보냄 (정기항목이어도 다시 안 생김)\n발송완료 = 이미 보냄 (시스템이 자동 표시)');
}

// 거래내역 금액 자동계산 수식
function setupTxFormulas_(sh, startRow, count) {
  if (count <= 0) return;
  var f = [];
  for (var k = 0; k < count; k++) {
    var r = startRow + k;
    f.push(['=IF($F' + r + '="","",$F' + r + '*$G' + r + ')',
            '=IF($H' + r + '="","",ROUND($H' + r + '*0.1,0))',
            '=IF($H' + r + '="","",$H' + r + '+$I' + r + ')']);
  }
  sh.getRange(startRow, 8, count, 3).setFormulas(f);
}

// 거래처ID 드롭다운
function applyCustomerDropdown_(sh, col) {
  var custSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CUSTOMER);
  if (!custSh) return;
  var idRange = custSh.getRange(2, 1, Math.max(custSh.getMaxRows() - 1, 1), 1);
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(idRange, true).setAllowInvalid(true).build();
  sh.getRange(2, col, 500, 1).setDataValidation(rule);
}

// ----- 발송로그 -----
function initLog_() {
  var sh = createOrClear_(SHEETS.LOG);
  var head = ['발송일시', '청구월', '거래처ID', '상호', '이메일', '상태', 'PDF링크', '오류메시지'];
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e8eef7');
  sh.setFrozenRows(1);
}

// 설정 시트가 아직 없을 수도 있는 시점(시드 순서) 대비
function getConfigSafe_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(SHEETS.CONFIG)) return getConfig();
  return { '청구월_기준': '당월' };
}
