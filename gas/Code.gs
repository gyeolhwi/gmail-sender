/**
 * 거래명세서 자동 발송 시스템 (Google Apps Script)
 * - 원본 양식 그대로 사용: '자료입력'에 값 기입 → '세금계산서양식'(파란 양식)이 자동 렌더링 → PDF
 * - 매월 10일 자동, 당월 기준, 거래처별 다중 품목(최대 4), 정기항목 자동생성, Drive 보관
 *
 * 사전: 거래명세서 양식.xlsx 를 가져와 [자료입력]/[세금계산서양식] 탭이 있어야 함.
 */

// ===== 상수 =====================================================
var SHEETS = {
  CONFIG: '설정', CUSTOMER: '거래처', RECUR: '정기항목', TX: '거래내역',
  INPUT: '자료입력',          // 스크립트가 값을 써넣는 입력 시트 (숨김)
  FORM: '세금계산서양식',      // 원본 파란 양식 (수식 자동 렌더링) → PDF
  LOG: '발송로그'
};
var MAX_ITEMS = 4;           // 원본 양식의 품목 행 수
var VAT_RATE = 0.1;          // 부가세 10% (별도)
var TZ = 'Asia/Seoul';

// 통합 후: 입력 셀은 '세금계산서양식' 시트의 인쇄영역 밖(원래 행 + 40, 숨김)에 위치.
// 양식의 68개 수식이 이 셀들을 참조해 파란 양식·자릿수 칸을 그대로 렌더링한다.
var ROW_SHIFT = 40;
var IN = {
  SUP_NAME: 'C45', SUP_REG: 'C46', SUP_CEO: 'C47', SUP_ADDR: 'C48', SUP_BIZ: 'C49', SUP_ITEM: 'C50',
  RCV_NAME: 'H45', RCV_REG: 'H46', RCV_CEO: 'H47', RCV_ADDR: 'H48', RCV_BIZ: 'H49', RCV_ITEM: 'H50', RCV_NOTE: 'H51',
  ISSUE_DATE: 'C55',
  ITEM_FIRST_ROW: 57,        // 품목 1행 (C:품목 D:규격 E:수량 F:단가 G:공급가액 H:부가세 I:발행금액)
  SUM_ROW: 61                // 합계금액 행 (G:공급가액합 H:부가세합 I:발행금액합)
};

// ===== 메뉴 =====================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📨 거래명세서')
    .addItem('👀 미리보기 (이번 달)', 'previewMonthly')
    .addItem('📧 이번 달 발송하기', 'runMonthlyConfirm')
    .addSeparator()
    .addItem(autoSendStatusLabel_(), 'showAutoSendStatus')
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙ 관리 (가끔만)')
      .addItem('자동발송 켜기 (설정의 발송일/시각)', 'createMonthlyTrigger')
      .addItem('자동발송 끄기', 'deleteTriggers')
      .addItem('테스트: 1분 후 자동발송 (나에게)', 'scheduleAutoTest')
      .addSeparator()
      .addItem('설치 / 갱신', 'setupUpgrades')
      .addItem('테스트 발송 (나에게 1건)', 'runPOC'))
    .addToUi();
}

// 메뉴에 표시할 자동발송 상태 (설정 시트의 캐시값을 읽음 — onOpen 권한 제약 회피)
function autoSendStatusLabel_() {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
    if (!sh) return '자동발송 상태 보기';
    var rows = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === '자동발송_상태') return '자동발송: ' + rows[i][1];
    }
    return '자동발송 상태 보기';
  } catch (e) { return '자동발송 상태 보기'; }
}

// ===== 설정/데이터 로더 ==========================================
function getConfig() {
  var rows = sheet_(SHEETS.CONFIG).getDataRange().getValues();
  var cfg = {};
  for (var i = 1; i < rows.length; i++) if (rows[i][0] !== '') cfg[String(rows[i][0]).trim()] = rows[i][1];
  return cfg;
}

function loadCustomers() {
  var rows = sheet_(SHEETS.CUSTOMER).getDataRange().getValues();
  var idx = headerIndex_(rows[0]);
  var map = {};
  for (var i = 1; i < rows.length; i++) {
    var id = String(rows[i][idx['거래처ID']]).trim();
    if (!id) continue;
    map[id] = {
      id: id, name: rows[i][idx['상호']], email: String(rows[i][idx['이메일']]).trim(),
      manager: rows[i][idx['담당자']], reg: rows[i][idx['사업자번호']], ceo: rows[i][idx['대표자']],
      addr: rows[i][idx['주소']],
      biz: idx['업태'] != null ? rows[i][idx['업태']] : '',
      item: idx['종목'] != null ? rows[i][idx['종목']] : '',
      fileRule: idx['파일명규칙'] != null ? rows[i][idx['파일명규칙']] : '',
      send: idx['발송여부'] != null ? rows[i][idx['발송여부']] : true
    };
  }
  return map;
}

function loadTransactions() {
  var rows = sheet_(SHEETS.TX).getDataRange().getValues();
  var idx = headerIndex_(rows[0]);
  var list = [];
  for (var i = 1; i < rows.length; i++) {
    var custId = String(rows[i][idx['거래처ID']]).trim();
    if (!custId) continue;
    var qty = Number(rows[i][idx['수량']]) || 0;
    var price = Number(rows[i][idx['단가']]) || 0;
    var supply = qty * price, vat = Math.round(supply * VAT_RATE);
    list.push({
      row: i + 1, custId: custId, billMonth: normMonth_(rows[i][idx['청구월']]),
      date: rows[i][idx['거래일자']], item: rows[i][idx['품목']], spec: rows[i][idx['규격']],
      qty: qty, price: price, supply: supply, vat: vat, total: supply + vat,
      status: String(rows[i][idx['발송상태']] || '').trim()
    });
  }
  return list;
}

function groupByCustomer(txs, billMonth, onlyUnsent) {
  var groups = {};
  for (var i = 0; i < txs.length; i++) {
    var t = txs[i];
    if (t.billMonth !== billMonth) continue;
    if (onlyUnsent && (t.status === '발송완료' || t.status === '보류' || t.status === '제외')) continue;
    (groups[t.custId] = groups[t.custId] || []).push(t);
  }
  return groups;
}

// ===== 핵심: 거래처 1곳 처리 ====================================
function processCustomer(custId, items, config, customer, options) {
  options = options || {};
  var billMonth = items[0].billMonth;
  if (!customer) throw new Error('거래처 마스터에 없음: ' + custId);
  if (items.length > MAX_ITEMS) throw new Error('품목 ' + items.length + '개로 최대 ' + MAX_ITEMS + '개 초과 (' + customer.name + ')');

  fillTemplate_(config, customer, items, billMonth);
  var totals = sumItems_(items);
  var pdf = exportTemplateAsPdf_().setName(buildFileName_(customer, billMonth));
  var pdfUrl = archivePdf_(pdf, config, billMonth);

  var recipient = options.recipientOverride || customer.email;
  var doSend = !options.noSend && (options.forceSend || String(config['발송_활성화']).toUpperCase() === 'TRUE');
  if (doSend && !recipient) throw new Error('이메일 없음: ' + customer.name);
  if (doSend) sendStatementMail_(recipient, customer, pdf, config, totals, billMonth, options.subjectPrefix);

  return { status: doSend ? '성공' : '생성만(발송보류)', pdfUrl: pdfUrl, totals: totals, sentTo: doSend ? recipient : '' };
}

// ===== 양식 채우기 (세금계산서양식 숨김 입력영역에 값 기입 → 수식 자동 렌더링) ====
function fillTemplate_(config, customer, items, billMonth) {
  var sh = sheet_(SHEETS.FORM);

  sh.getRange(IN.SUP_NAME).setValue(config['공급자_상호'] || '');
  sh.getRange(IN.SUP_REG).setValue(formatBizNo_(config['공급자_사업자번호']));
  sh.getRange(IN.SUP_CEO).setValue(config['공급자_대표자'] || '');
  sh.getRange(IN.SUP_ADDR).setValue(config['공급자_주소'] || '');
  sh.getRange(IN.SUP_BIZ).setValue(config['공급자_업태'] || '');
  sh.getRange(IN.SUP_ITEM).setValue(config['공급자_종목'] || '');

  sh.getRange(IN.RCV_NAME).setValue(customer.name || '');
  sh.getRange(IN.RCV_REG).setValue(formatBizNo_(customer.reg));
  sh.getRange(IN.RCV_CEO).setValue(customer.ceo || '');
  sh.getRange(IN.RCV_ADDR).setValue(customer.addr || '');
  sh.getRange(IN.RCV_BIZ).setValue(customer.biz || '');
  sh.getRange(IN.RCV_ITEM).setValue(customer.item || '');
  sh.getRange(IN.RCV_NOTE).setValue('');

  sh.getRange(IN.ISSUE_DATE).setValue(new Date());

  var first = IN.ITEM_FIRST_ROW;
  sh.getRange(first, 3, MAX_ITEMS, 7).clearContent();
  for (var i = 0; i < items.length; i++) {
    var t = items[i], r = first + i;
    sh.getRange(r, 3).setValue(t.item);
    sh.getRange(r, 4).setValue(t.spec || '');
    sh.getRange(r, 5).setValue(t.qty);
    sh.getRange(r, 6).setValue(t.price);
    sh.getRange(r, 7).setValue(t.supply);
    sh.getRange(r, 8).setValue(t.vat);
    sh.getRange(r, 9).setValue(t.total);
  }
  var s = sumItems_(items);
  sh.getRange(IN.SUM_ROW, 7).setValue(s.supply);
  sh.getRange(IN.SUM_ROW, 8).setValue(s.vat);
  sh.getRange(IN.SUM_ROW, 9).setValue(s.total);

  SpreadsheetApp.flush();   // 수식 재계산 후 PDF 추출 보장
}

// ===== 세금계산서양식 → PDF (원본 파란 양식 그대로) =============
function exportTemplateAsPdf_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sheet_(SHEETS.FORM);
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?'
    + 'format=pdf'
    + '&gid=' + sheet.getSheetId()
    + '&portrait=true&size=A4&fitw=true'
    + '&gridlines=false&printtitle=false&sheetnames=false'
    + '&pagenumbers=false&fzr=false'
    + '&top_margin=0.3&bottom_margin=0.3&left_margin=0.3&right_margin=0.3'
    + '&range=A1:AG24';
  var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
  return resp.getBlob();
}

// ===== Drive 보관 ===============================================
function archivePdf_(blob, config, billMonth) {
  var root = getOrCreateFolder_(DriveApp.getRootFolder(), config['보관폴더명'] || '거래명세서_보관');
  var yearF = getOrCreateFolder_(root, billMonth.split('-')[0]);
  return getOrCreateFolder_(yearF, billMonth).createFile(blob).getUrl();
}
function getOrCreateFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// ===== Gmail 발송 ===============================================
function sendStatementMail_(recipient, customer, pdf, config, totals, billMonth, subjectPrefix) {
  var map = {
    '청구월': billMonth, '담당자': customer.manager || customer.name, '상호': customer.name,
    '공급가액': won_(totals.supply), '세액': won_(totals.vat), '발행금액': won_(totals.total),
    '공급자_상호': config['공급자_상호'] || '', '공급자_입금계좌': config['공급자_입금계좌'] || ''
  };
  var subject = (subjectPrefix || '') + replaceTokens_(config['메일_제목'] || '[{공급자_상호}] {청구월} 거래명세서', map);
  var body = replaceTokens_(config['메일_본문'] || '{청구월} 거래명세서를 첨부드립니다.', map);
  GmailApp.sendEmail(recipient, subject, body, { name: config['공급자_상호'] || '거래명세서', attachments: [pdf] });
}

// ===== 로그/상태 ===============================================
function writeLog_(billMonth, customer, status, pdfUrl, errMsg) {
  sheet_(SHEETS.LOG).appendRow([
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    billMonth, customer ? customer.id : '', customer ? customer.name : '',
    customer ? customer.email : '', status, pdfUrl || '', errMsg || ''
  ]);
}
function markSent_(items) {
  var sh = sheet_(SHEETS.TX);
  var statusCol = headerIndex_(sh.getDataRange().getValues()[0])['발송상태'] + 1;
  for (var i = 0; i < items.length; i++) sh.getRange(items[i].row, statusCol).setValue('발송완료');
}

// ===== 실행 진입점 =============================================
function runPOC() {
  ensureUnified_();   // 자료입력 탭이 남아있으면 자동으로 한 시트 통합
  var config = getConfig();
  var customers = loadCustomers();
  var billMonth = currentBillMonth_(config);
  var groups = groupByCustomer(loadTransactions(), billMonth, false);

  var ids = Object.keys(groups);
  if (!ids.length) { alert_('당월(' + billMonth + ') 거래내역이 없습니다. 먼저 거래내역을 입력하세요.'); return; }
  var custId = groups['C999'] ? 'C999' : (groups['C001'] ? 'C001' : ids[0]);
  var testTo = config['테스트_수신메일'] || config['관리자_이메일'];
  if (!testTo) { alert_('설정 시트의 [테스트_수신메일] 또는 [관리자_이메일]을 먼저 입력하세요.'); return; }

  try {
    var res = processCustomer(custId, groups[custId], config, customers[custId], {
      recipientOverride: testTo, forceSend: true, subjectPrefix: '[POC테스트] '
    });
    writeLog_(billMonth, customers[custId], 'POC-' + res.status, res.pdfUrl, 'POC 수신: ' + testTo);
    alert_('POC 완료 ✅\n거래처: ' + customers[custId].name
      + '\n청구월: ' + billMonth + '\n합계: ₩' + won_(res.totals.total)
      + '\n발송(테스트): ' + testTo + '\nPDF 보관: ' + res.pdfUrl
      + '\n\n메일함과 Drive의 [거래명세서_보관] 폴더를 확인하세요.');
  } catch (e) {
    writeLog_(billMonth, customers[custId], 'POC-실패', '', e.message);
    alert_('POC 실패 ❌\n' + e.message);
  }
}

/** ②-1 대량발송 테스트: 테스터 5명을 즉석 생성해 실제 파이프라인으로 일괄 발송 (전부 테스트메일) */
function runBulkTest() {
  ensureUnified_();
  var config = getConfig();
  var billMonth = currentBillMonth_(config);
  var testTo = config['테스트_수신메일'] || config['관리자_이메일'];
  if (!testTo) { alert_('설정 시트의 [테스트_수신메일] 또는 [관리자_이메일]을 먼저 입력하세요.'); return; }

  var N = 5, ok = 0, fail = 0, errs = [];
  for (var i = 1; i <= N; i++) {
    var t = {
      id: 'B' + i, name: '테스터상사' + i, email: testTo, manager: 'tester' + i,
      reg: i + '11-11-1111' + i, ceo: '김테스터' + i, addr: '서울 강남구 테헤란로 ' + (i * 100) + '길 ' + i,
      biz: '도소매', item: '전자상거래', fileRule: '', send: true
    };
    var qty = i, price = 10000 + i * 5000, supply = qty * price, vat = Math.round(supply * VAT_RATE);
    var items = [{
      billMonth: billMonth, date: billMonth + '-10', item: '테스트 상품 ' + i, spec: 'EA',
      qty: qty, price: price, supply: supply, vat: vat, total: supply + vat
    }];
    try {
      var res = processCustomer(t.id, items, config, t, {
        recipientOverride: testTo, forceSend: true, subjectPrefix: '[BULK ' + i + '/' + N + ' ' + t.name + '] '
      });
      writeLog_(billMonth, t, 'BULK-성공', res.pdfUrl, '수신: ' + testTo);
      ok++;
    } catch (e) {
      writeLog_(billMonth, t, 'BULK-실패', '', e.message);
      errs.push(t.name + ': ' + e.message);
      fail++;
    }
  }
  alert_('대량발송 테스트 완료\n총 ' + N + '건 → 성공 ' + ok + ' / 실패 ' + fail
    + '\n수신: ' + testTo + '\n\n메일함에서 제목 [BULK ...] ' + ok + '통을 확인하세요.'
    + (errs.length ? '\n\n오류:\n' + errs.join('\n') : ''));
}

/** 👀 미리보기: 정기항목 자동채움 → 당월 미발송 PDF를 발송 없이 생성(Drive 저장) */
function previewMonthly() {
  ensureUnified_();
  var config = getConfig();
  generateMonthlyCore_(config);   // 정기항목 → 거래내역 자동 채움
  var customers = loadCustomers();
  var billMonth = currentBillMonth_(config);
  var groups = groupByCustomer(loadTransactions(), billMonth, true);
  var ids = Object.keys(groups);
  if (!ids.length) { alert_('당월(' + billMonth + ') 미발송 건이 없습니다.'); return; }

  var ok = 0, fail = 0, errs = [];
  for (var i = 0; i < ids.length; i++) {
    var id = ids[i], c = customers[id];
    try {
      if (!c) throw new Error('거래처 목록에 없는 ID');
      processCustomer(id, groups[id], config, c, { noSend: true });
      ok++;
    } catch (e) { fail++; errs.push((c ? c.name : id) + ': ' + e.message); }
  }
  alert_('미리보기 생성 완료 (발송 안 함) 👀\nPDF ' + ok + '건 → Drive [' + (config['보관폴더명'] || '거래명세서_보관') + '] 폴더에서 확인하세요.'
    + (fail ? '\n\n실패 ' + fail + ':\n' + errs.join('\n') : ''));
}

/** 📧 발송 전 확인창 (실수 방지): 정기항목 자동채움 → 확인 시 runMonthly */
function runMonthlyConfirm() {
  ensureUnified_();
  var config = getConfig();
  generateMonthlyCore_(config);   // 정기항목 → 거래내역 자동 채움
  var customers = loadCustomers();
  var billMonth = currentBillMonth_(config);
  var groups = groupByCustomer(loadTransactions(), billMonth, true);
  var ids = Object.keys(groups);
  if (!ids.length) { alert_('당월(' + billMonth + ') 발송할 미발송 건이 없습니다.'); return; }

  var noEmail = [], unknown = [];
  ids.forEach(function (id) {
    var c = customers[id];
    if (!c) { unknown.push(id); return; }
    if (!c.email) noEmail.push(c.name);
  });
  var active = String(config['발송_활성화']).toUpperCase() === 'TRUE';
  var warn = '';
  if (unknown.length) warn += '\n⚠️ 거래처 목록에 없는 ID: ' + unknown.join(', ');
  if (noEmail.length) warn += '\n⚠️ 이메일 없는 거래처(건너뜀): ' + noEmail.join(', ');

  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(billMonth + ' 거래명세서 발송',
    '대상 ' + ids.length + '곳에 ' + (active ? '실제 메일을 발송합니다.' : '(발송보류 모드 — PDF만 생성)') + warn + '\n\n계속할까요?',
    ui.ButtonSet.YES_NO);
  if (resp === ui.Button.YES) runMonthly();
}

function runMonthly() {
  ensureUnified_();   // 자료입력 탭이 남아있으면 자동으로 한 시트 통합
  var config = getConfig();
  generateMonthlyCore_(config);   // 자동발송도 정기항목 → 거래내역 자동 채움
  var customers = loadCustomers();
  var billMonth = currentBillMonth_(config);
  var groups = groupByCustomer(loadTransactions(), billMonth, true);

  var ids = Object.keys(groups);
  var ok = 0, fail = 0, skip = 0;
  for (var i = 0; i < ids.length; i++) {
    var custId = ids[i], cust = customers[custId];
    try {
      if (!cust) throw new Error('거래처 마스터에 없음');
      if (String(cust.send).toUpperCase() === 'FALSE') { skip++; writeLog_(billMonth, cust, '제외', '', '발송여부=FALSE'); continue; }
      var res = processCustomer(custId, groups[custId], config, cust, {});
      if (res.sentTo) { markSent_(groups[custId]); ok++; }
      writeLog_(billMonth, cust, res.status, res.pdfUrl, '');
    } catch (e) {
      fail++;
      writeLog_(billMonth, cust || { id: custId, name: '', email: '' }, '실패', '', e.message);
    }
  }
  notifyAdmin_(config, billMonth, ok, fail, skip, ids.length);
}

function notifyAdmin_(config, billMonth, ok, fail, skip, total) {
  var admin = config['관리자_이메일'];
  if (!admin) return;
  var msg = billMonth + ' 거래명세서 발송 결과\n\n총 ' + total + '곳 중\n - 성공: ' + ok + '\n - 실패: ' + fail + '\n - 제외: ' + skip
    + '\n\n자세한 내역은 [발송로그] 시트를 확인하세요.';
  GmailApp.sendEmail(admin, '[거래명세서] ' + billMonth + ' 발송결과 (성공 ' + ok + '/실패 ' + fail + ')', msg);
}

// ===== 정기항목 → 이번 달 거래내역 자동채움 ======================
// 정기항목에 적어둔 내용을 당월 거래내역으로 채운다(중복은 건너뜀). preview/발송에서 자동 호출.
function generateMonthlyCore_(config) {
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.RECUR)) return { added: 0, skipped: 0 };
  var billMonth = currentBillMonth_(config);
  var recur = loadRecurring_();
  if (!recur.length) return { added: 0, skipped: 0 };
  var txSh = sheet_(SHEETS.TX);
  var day = ('0' + (Number(config['발송일']) || 10)).slice(-2);
  var plan = planMonthlyRows_(recur, existingTxKeys_(txSh, billMonth), billMonth, billMonth + '-' + day);
  if (plan.rows.length) {
    var startRow = nextEmptyDataRow_(txSh);
    txSh.getRange(startRow, 1, plan.rows.length, 7).setValues(plan.rows);
    txSh.getRange(startRow, 11, plan.rows.length, 1).setValues(plan.rows.map(function () { return ['미발송']; }));
    setupTxFormulas_(txSh, startRow, plan.rows.length);
  }
  return { added: plan.rows.length, skipped: plan.skipped };
}

function loadRecurring_() {
  var rows = sheet_(SHEETS.RECUR).getDataRange().getValues();
  var idx = headerIndex_(rows[0]);
  var list = [];
  for (var i = 1; i < rows.length; i++) {
    var id = String(rows[i][idx['거래처ID']]).trim();
    if (!id) continue;
    if (String(rows[i][idx['사용여부']]).toUpperCase() === 'FALSE') continue;
    list.push({ custId: id, item: rows[i][idx['품목']], spec: rows[i][idx['규격']], qty: Number(rows[i][idx['수량']]) || 0, price: Number(rows[i][idx['단가']]) || 0 });
  }
  return list;
}
function planMonthlyRows_(recur, existing, billMonth, dateStr) {
  var rows = [], skipped = 0;
  for (var i = 0; i < recur.length; i++) {
    var r = recur[i];
    if (existing.has(r.custId + '|' + r.item)) { skipped++; continue; }
    rows.push([r.custId, billMonth, dateStr, r.item, r.spec || '', r.qty, r.price]);
  }
  return { rows: rows, skipped: skipped };
}
function existingTxKeys_(sh, billMonth) {
  var rows = sh.getDataRange().getValues();
  var idx = headerIndex_(rows[0]);
  var set = new Set();
  for (var i = 1; i < rows.length; i++) {
    var id = String(rows[i][idx['거래처ID']]).trim();
    if (!id) continue;
    if (normMonth_(rows[i][idx['청구월']]) !== billMonth) continue;
    set.add(id + '|' + rows[i][idx['품목']]);
  }
  return set;
}
function nextEmptyDataRow_(sh) {
  var col = sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), 1).getValues();
  for (var i = 0; i < col.length; i++) if (col[i][0] === '' || col[i][0] === null) return i + 2;
  return col.length + 2;
}

// ===== 트리거(자동발송) ========================================
function createMonthlyTrigger() {
  var config = getConfig();
  var day = Number(config['발송일']), hour = Number(config['발송시각']);
  if (!(Number.isInteger(day) && day >= 1 && day <= 31)) {
    alert_('⚠️ [설정]의 발송일이 올바르지 않습니다.\n1~31 사이 숫자여야 합니다. (현재: "' + config['발송일'] + '")\n설정 시트에서 목록으로 다시 선택하세요.'); return;
  }
  if (!(Number.isInteger(hour) && hour >= 0 && hour <= 23)) {
    alert_('⚠️ [설정]의 발송시각이 올바르지 않습니다.\n0~23 사이 숫자(24시간제)여야 합니다. (현재: "' + config['발송시각'] + '")\n설정 시트에서 목록으로 다시 선택하세요.'); return;
  }
  deleteTriggers_();   // 기존 정리 후 재등록
  ScriptApp.newTrigger('runMonthly').timeBased().onMonthDay(day).atHour(hour).create();
  setConfigValue_('자동발송_상태', '🟢 켜짐 (매월 ' + day + '일 ' + hour + '시)');
  alert_('자동발송 켜짐 ✅\n매월 ' + day + '일 ' + hour + '시에 자동 발송됩니다.\n\n시간을 바꾸려면 [설정] 시트의 발송일/발송시각을 고친 뒤 다시 [자동발송 켜기]를 누르세요.');
}

function deleteTriggers() {
  deleteTriggers_();
  setConfigValue_('자동발송_상태', '⚪ 꺼짐');
  alert_('자동발송 꺼짐 ⚪\n이제 자동으로는 발송되지 않습니다. (수동 발송은 그대로 가능)');
}

// 자동발송/테스트 트리거 모두 제거 (내부용)
function deleteTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var h = t.getHandlerFunction();
    if (h === 'runMonthly' || h === 'runAutoSendTest_') ScriptApp.deleteTrigger(t);
  });
}

// 실제 트리거 존재 여부로 상태 확인 (권위 있는 값) + 설정 캐시 갱신
function showAutoSendStatus() {
  var on = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'runMonthly'; });
  var config = getConfig();
  var day = Number(config['발송일']) || 10;
  var hour = Number(config['발송시각']); if (isNaN(hour)) hour = 9;
  setConfigValue_('자동발송_상태', on ? '🟢 켜짐 (매월 ' + day + '일 ' + hour + '시)' : '⚪ 꺼짐');
  alert_('자동발송 상태\n\n' + (on ? '🟢 켜짐 — 매월 ' + day + '일 ' + hour + '시에 발송' : '⚪ 꺼짐')
    + '\n\n(메뉴 항목과 [설정] 시트 [자동발송_상태] 칸에서도 확인됩니다)');
}

// 테스트: 약 1분 후 자동발송을 1회 실행 (전부 테스트메일로 — 실고객 X)
function scheduleAutoTest() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'runAutoSendTest_') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('runAutoSendTest_').timeBased().after(60 * 1000).create();
  alert_('테스트 예약 ✅\n약 1분 후 자동발송이 1회 실행됩니다.\n전부 [테스트메일]로만 가고 실고객엔 안 갑니다.\n\n1~2분 뒤 테스트 메일함과 [발송로그]를 확인하세요.');
}

// (트리거 전용) 자동발송 동작 확인용 — 샘플 1건을 테스트메일로 발송 + 작동 알림
function runAutoSendTest_() {
  var config = getConfig();
  var testTo = config['테스트_수신메일'] || config['관리자_이메일'];
  var billMonth = currentBillMonth_(config);
  try {
    generateMonthlyCore_(config);
    var customers = loadCustomers();
    var groups = groupByCustomer(loadTransactions(), billMonth, true);
    var ids = Object.keys(groups), done = false;
    for (var i = 0; i < ids.length && !done; i++) {
      var c = customers[ids[i]];
      if (!c) continue;
      processCustomer(ids[i], groups[ids[i]], config, c, { recipientOverride: testTo, forceSend: true, subjectPrefix: '[자동발송테스트] ' });
      done = true;
    }
    if (testTo) GmailApp.sendEmail(testTo, '[자동발송테스트] ✅ 예약 트리거 작동 확인',
      '예약된 자동발송 트리거가 정상 실행되었습니다.\n시각: ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')
      + (done ? '\n샘플 명세서 1건도 함께 발송했습니다.' : '\n(당월 거래내역이 없어 샘플 발송은 생략)'));
    writeLog_(billMonth, { id: 'TEST', name: '자동발송테스트', email: testTo }, '자동테스트-성공', '', '트리거 작동');
  } catch (e) {
    writeLog_(billMonth, { id: 'TEST', name: '자동발송테스트', email: testTo }, '자동테스트-실패', '', e.message);
  } finally {
    ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'runAutoSendTest_') ScriptApp.deleteTrigger(t); });
  }
}

// ===== 유틸 ====================================================
function sheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('시트 없음: ' + name + ' — 양식을 가져왔는지/설치했는지 확인하세요.');
  return sh;
}
// '자료입력' 탭이 남아있으면(미통합) 한 시트로 자동 통합
function ensureUnified_() {
  if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.INPUT)) unifyToOneSheet();
}
// 설정 시트의 key 값을 갱신(없으면 추가)
function setConfigValue_(key, value) {
  var sh = sheet_(SHEETS.CONFIG);
  var rows = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  }
  sh.appendRow([key, value]);
}
function headerIndex_(head) {
  var idx = {};
  for (var i = 0; i < head.length; i++) idx[String(head[i]).trim()] = i;
  return idx;
}
function sumItems_(items) {
  var s = { supply: 0, vat: 0, total: 0 };
  items.forEach(function (t) { s.supply += t.supply; s.vat += t.vat; s.total += t.total; });
  return s;
}
function buildFileName_(customer, billMonth) {
  var base = customer.fileRule || (billMonth + '_거래명세서_' + customer.name);
  return String(base).replace(/[\/\\:*?"<>|]/g, '_') + '.pdf';
}
function replaceTokens_(text, map) {
  return String(text).replace(/\{([^{}]+)\}/g, function (m, k) { return (map[k] != null) ? map[k] : m; });
}
function won_(n) {
  n = Math.round(Number(n) || 0);
  return (n < 0 ? '-' : '') + String(Math.abs(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
// 사업자번호를 양식이 기대하는 하이픈 형식(XXX-XX-XXXXX)으로 정규화. 10자리 아니면 원본 유지.
function formatBizNo_(v) {
  var d = String(v == null ? '' : v).replace(/\D/g, '');
  if (d.length === 10) return d.slice(0, 3) + '-' + d.slice(3, 5) + '-' + d.slice(5);
  return String(v == null ? '' : v);
}
function currentBillMonth_(config) {
  var now = new Date();
  if (String(config['청구월_기준']).indexOf('전월') >= 0) now.setMonth(now.getMonth() - 1);
  return Utilities.formatDate(now, TZ, 'yyyy-MM');
}
function normMonth_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM');
  var m = String(v).trim().match(/(\d{4})[-.\/](\d{1,2})/);
  return m ? (m[1] + '-' + ('0' + m[2]).slice(-2)) : String(v).trim();
}
function alert_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}
