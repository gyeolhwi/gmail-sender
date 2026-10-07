/**
 * 일괄메일 발송 (Google Apps Script)
 * - [고객] 시트에서 체크한 행에게 [템플릿] 시트의 양식으로 메일 발송
 * - 본문의 {{열이름}} 은 고객 행 값으로 치환, {{이름|웹사이트}} 는 앞 값이 비면 뒤 값 사용
 */

// ===== 상수 =====================================================
var SHEETS = { CONFIG: '설정', CUSTOMER: '고객', TEMPLATE: '템플릿', LOG: '발송로그' };
var COL = { CHECK: '선택', EMAIL: '이메일', STATUS: '상태', SENT_AT: '발송일시' };  // 고객 시트 고정 열
var TZ = 'Asia/Seoul';

// ===== 메뉴 =====================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📧 일괄메일')
    .addItem('👀 미리보기', 'preview')
    .addItem('📨 발송', 'sendNow')
    .addItem('⏰ 예약하기', 'schedule')
    .addSeparator()
    .addItem('⚙ 설치', 'setup')
    .addToUi();
}

// ===== 메뉴 동작 =================================================
// 체크된 첫 번째 고객 기준으로 치환 결과를 보여준다 (발송 안 함)
function preview() {
  var tpl = askTemplate_();
  if (!tpl) return;
  var rows = checkedRows_();
  if (!rows.length) { alert_('[고객] 시트에서 보낼 사람을 체크하세요.'); return; }
  var r = rows[0];
  var unknown = unknownKeys_(tpl.subject + tpl.body, r.data);
  alert_('체크된 고객: ' + rows.length + '명\n' +
    '(아래는 첫 번째 고객 ' + r.data[COL.EMAIL] + ' 기준)\n\n' +
    '제목: ' + render_(tpl.subject, r.data) + '\n\n' + render_(tpl.body, r.data) +
    (unknown.length ? '\n\n⚠ 고객 시트에 없는 열: ' + unknown.join(', ') : ''));
}

function sendNow() {
  var tpl = askTemplate_();
  if (!tpl) return;
  var n = checkedRows_().length;
  if (!n) { alert_('[고객] 시트에서 보낼 사람을 체크하세요.'); return; }
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('발송 확인', '템플릿 [' + tpl.name + '] 로 ' + n + '명에게 발송합니다.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var res = sendChecked_(tpl);
  alert_('발송 완료 ✅  성공 ' + res.ok + ' / 실패 ' + res.fail + '\n자세한 내용은 [발송로그] 시트를 보세요.');
}

// [설정]의 예약_일시·예약_템플릿으로 1회성 예약. 예약_일시를 비우고 누르면 예약 취소.
function schedule() {
  deleteScheduleTriggers_();
  var config = getConfig_();
  var at = config['예약_일시'];
  if (!at) { alert_('예약이 취소되었습니다. (예약 없음)'); return; }
  if (!(at instanceof Date) || at <= new Date()) { alert_('[설정]의 예약_일시를 미래 날짜·시각으로 입력하세요.\n예: 2026-10-05 09:00'); return; }
  if (!findTemplate_(config['예약_템플릿'])) { alert_('[설정]의 예약_템플릿 이름이 [템플릿] 시트에 없습니다.'); return; }
  ScriptApp.newTrigger('runScheduled').timeBased().at(at).create();
  alert_('예약 완료 ⏰ ' + fmt_(at) + '\n템플릿 [' + config['예약_템플릿'] + ']\n\n그 시각에 [고객] 시트에서 체크되어 있는 사람에게 발송됩니다.');
}

// 예약 트리거가 실행하는 함수 (UI 없음)
function runScheduled() {
  deleteScheduleTriggers_();
  var tpl = findTemplate_(getConfig_()['예약_템플릿']);
  if (tpl) sendChecked_(tpl);
}

// ===== 발송 =====================================================
function sendChecked_(tpl) {
  var config = getConfig_();
  var options = {};
  if (config['보내는사람_이름']) options.name = config['보내는사람_이름'];
  if (config['첨부파일_링크']) options.attachments = [attachment_(config['첨부파일_링크'])];

  var sh = sheet_(SHEETS.CUSTOMER);
  var head = headers_(sh);
  var res = { ok: 0, fail: 0 };
  checkedRows_().forEach(function (r) {
    var email = String(r.data[COL.EMAIL]).trim();
    var status = '발송완료', err = '';
    try {
      if (!isEmail_(email)) throw new Error('이메일 형식 오류');
      MailApp.sendEmail(email, render_(tpl.subject, r.data), render_(tpl.body, r.data), options);
      res.ok++;
    } catch (e) {
      status = '실패'; err = e.message; res.fail++;
    }
    var now = new Date();
    sh.getRange(r.row, head[COL.STATUS] + 1).setValue(status);
    sh.getRange(r.row, head[COL.SENT_AT] + 1).setValue(now);
    sh.getRange(r.row, head[COL.CHECK] + 1).setValue(false);
    sheet_(SHEETS.LOG).appendRow([now, email, tpl.name, status, err]);
  });
  return res;
}

// {{열이름}} / {{이름|웹사이트}} 치환 — 비어있지 않은 첫 값 사용
function render_(text, data) {
  return String(text).replace(/\{\{([^}]+)\}\}/g, function (_, keys) {
    var list = keys.split('|');
    for (var i = 0; i < list.length; i++) {
      var v = data[list[i].trim()];
      if (v !== undefined && v !== '') return v;
    }
    return '';
  });
}

function unknownKeys_(text, data) {
  var out = [];
  (String(text).match(/\{\{([^}]+)\}\}/g) || []).forEach(function (m) {
    m.slice(2, -2).split('|').forEach(function (k) {
      k = k.trim();
      if (!(k in data) && out.indexOf(k) < 0) out.push(k);
    });
  });
  return out;
}

function attachment_(link) {
  var id = String(link).match(/[-\w]{25,}/);
  if (!id) throw new Error('[설정]의 첨부파일_링크가 올바르지 않습니다.');
  return DriveApp.getFileById(id[0]).getBlob();
}

function isEmail_(s) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
}

// ===== 시트 읽기 =================================================
// 체크된 고객 행 목록: [{row: 시트행번호, data: {열이름: 값}}]
function checkedRows_() {
  var values = sheet_(SHEETS.CUSTOMER).getDataRange().getValues();
  var head = values[0];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var data = {};
    head.forEach(function (h, c) { data[String(h).trim()] = values[i][c]; });
    if (data[COL.CHECK] === true) out.push({ row: i + 1, data: data });
  }
  return out;
}

function headers_(sh) {
  var map = {};
  sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].forEach(function (h, i) { map[String(h).trim()] = i; });
  return map;
}

function templates_() {
  return sheet_(SHEETS.TEMPLATE).getDataRange().getValues().slice(1)
    .filter(function (r) { return String(r[0]).trim(); })
    .map(function (r) { return { name: String(r[0]).trim(), subject: r[1], body: r[2] }; });
}

function findTemplate_(name) {
  name = String(name || '').trim();
  return templates_().filter(function (t) { return t.name === name; })[0];
}

// 템플릿 이름을 입력받는다. 템플릿이 1개면 묻지 않고 그것을 쓴다.
function askTemplate_() {
  var list = templates_();
  if (!list.length) { alert_('[템플릿] 시트에 템플릿을 먼저 작성하세요.'); return null; }
  if (list.length === 1) return list[0];
  var ui = SpreadsheetApp.getUi();
  var names = list.map(function (t) { return t.name; });
  var res = ui.prompt('템플릿 선택', '사용할 템플릿 이름을 입력하세요:\n' + names.join('\n'), ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return null;
  var tpl = findTemplate_(res.getResponseText());
  if (!tpl) alert_('그런 이름의 템플릿이 없습니다: ' + res.getResponseText());
  return tpl || null;
}

function getConfig_() {
  var out = {};
  sheet_(SHEETS.CONFIG).getDataRange().getValues().slice(1).forEach(function (r) { out[String(r[0]).trim()] = r[1]; });
  return out;
}

// ===== 유틸 =====================================================
function deleteScheduleTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runScheduled') ScriptApp.deleteTrigger(t);
  });
}

function sheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('[' + name + '] 시트가 없습니다. 메뉴 ⚙ 설치를 먼저 실행하세요.');
  return sh;
}

function fmt_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd HH:mm');
}

function alert_(msg) {
  SpreadsheetApp.getUi().alert(msg);
}
