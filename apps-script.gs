/**
 * 글로벌 현장학습 예산 장부 · 수신용 Google Apps Script
 * ------------------------------------------------------------
 * 앱에서 지출을 저장하면
 *   1) 영수증 사진을 드라이브 폴더 "글로벌 현장학습 영수증 / 날짜" 에 저장
 *        (파일 이름: 번호_날짜_지출처_금액통화_1.jpg — 번호는 장부·엑셀의 영수증 번호)
 *   2) 스프레드시트 "글로벌 현장학습 지출 장부" 의 [지출] 시트에 한 줄 기록
 *   3) [요약] 시트에서 구분·항목별 합계·잔액이 자동 계산
 *   4) 앱의 [영수증 붙임 용지] → 지출 한 건에 한 장(품의 제목·사용일·기안자·영수증 번호 + 사진)인 PDF 를
 *        "글로벌 현장학습 영수증 / 붙임 용지" 폴더에 저장
 *
 * [처음 배포]  script.google.com → 새 프로젝트 → 이 코드 붙여넣기 → 저장
 *   → 배포 → 새 배포 → 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → 나온 URL 을 앱 ⚙ 설정에 붙여넣기
 * [코드 고친 뒤]  배포 → 배포 관리 → (기존 배포) 연필 → 버전 "새 버전" → 배포  (URL 유지)
 * [확인]  URL 을 브라우저로 열어 build 값이 "budget-v5" 이면 이 버전이 라이브
 */
const TOKEN = 'gbudget-2026';
const BUILD = 'budget-v5';
const RECEIPT_FOLDER = '글로벌 현장학습 영수증';
const ATTACH_FOLDER = '붙임 용지';
const BOOK_NAME = '글로벌 현장학습 지출 장부';
const SHEET_LOG = '지출';
const SHEET_SUM = '요약';
const HEAD = ['번호', '날짜', '지출처', '구분', '항목', '통화', '금액', '환율', '원화 환산',
  '카드 수수료 여유분', '합계(원)', '메모', '영수증 링크', '상태', '기록시각', '계산식', 'id'];
const COL_ID = 17, COL_LINK = 13, COL_STATE = 14;

function doGet() {
  return json_({ ok: true, build: BUILD });
}

function doPost(e) {
  // 두 요청이 겹쳐도 같은 줄이 두 번 들어가지 않게 한 번에 하나씩 처리
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var data = JSON.parse(e.postData.contents);
    if (data.token !== TOKEN) return json_({ ok: false, error: '토큰 불일치' });

    if (data.action === 'ping') return json_({ ok: true, build: BUILD, book: book_().getUrl(), folder: folder_().getUrl() });
    if (data.action === 'add') return json_(add_(data));
    if (data.action === 'cancel') return json_(cancel_(data));
    if (data.action === 'list') return json_(list_());
    if (data.action === 'budget') return json_(budget_(data));
    if (data.action === 'attach') return json_(attach_(data));
    return json_({ ok: false, error: '알 수 없는 요청: ' + data.action });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// 지출 1건 저장: 사진 → 드라이브, 내용 → 시트 한 줄
function add_(data) {
  var r = data.rec || {};
  var sheet = logSheet_();
  var row = findRow_(sheet, r.id);
  // 같은 id 가 이미 있으면(전송이 끊겨 다시 보낸 경우) 또 만들지 않고 그대로 알려 준다
  if (row) {
    var old = String(sheet.getRange(row, COL_LINK).getValue() || '');
    return { ok: true, dup: true, links: old ? old.split('\n') : [] };
  }

  var safe = function (s) { return String(s || '').replace(/[\\/:*?"<>|\n\r]/g, ' ').trim().slice(0, 30); };
  var day = dayFolder_(r.date);
  var base = ('00' + r.no).slice(-3) + '_' + safe(r.date) + '_' + safe(r.place) + '_' + r.amt + r.cur;
  var links = [];
  (data.photos || []).forEach(function (b64, i) {
    var blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', base + '_' + (i + 1) + '.jpg');
    links.push(day.createFile(blob).getUrl());
  });

  sheet.appendRow([r.no, r.date, r.place, r.cat, r.item, r.cur, r.amt, r.rate, r.krw,
    r.fee, r.total, r.memo || '', links.join('\n'), '정상', new Date(), "'" + (r.expr || ''), r.id]);
  return { ok: true, links: links };
}

// 지출 취소: 줄을 지우지 않고 상태만 '취소' 로 바꾼다(영수증 파일도 그대로 둔다)
function cancel_(data) {
  var sheet = logSheet_();
  var row = findRow_(sheet, data.id);
  if (!row) return { ok: true, missing: true };
  sheet.getRange(row, COL_STATE).setValue('취소');
  sheet.getRange(row, 1, 1, HEAD.length).setFontLine('line-through').setFontColor('#999999');
  return { ok: true };
}

// 시트에 있는 지출 전부(새 폰에서 불러오기 · 붙임 용지 만들기에 쓴다)
function list_() {
  var sheet = logSheet_();
  var n = sheet.getLastRow();
  if (n < 2) return { ok: true, recs: [] };
  var tz = Session.getScriptTimeZone();
  var recs = sheet.getRange(2, 1, n - 1, HEAD.length).getValues().map(function (v) {
    var d = v[1] instanceof Date ? Utilities.formatDate(v[1], tz, 'yyyy-MM-dd') : String(v[1]);
    var link = String(v[12] || '');
    return { no: Number(v[0]) || 0, date: d, place: String(v[2]), cat: String(v[3]), item: String(v[4]),
      cur: String(v[5]), amt: Number(v[6]), rate: Number(v[7]), krw: Number(v[8]), fee: Number(v[9]),
      total: Number(v[10]), memo: String(v[11]), links: link ? link.split('\n') : [],
      del: String(v[13]) === '취소', expr: String(v[15]), id: String(v[16]) };
  });
  return { ok: true, recs: recs };
}

// 예산표를 [요약] 시트에 다시 쓴다. 사용액은 [지출] 시트를 보는 수식이라 저절로 바뀐다.
function budget_(data) {
  var items = data.items || [];
  var ss = book_();
  logSheet_();
  var sh = ss.getSheetByName(SHEET_SUM) || ss.insertSheet(SHEET_SUM);
  sh.clear();
  var L = "'" + SHEET_LOG + "'!";
  var rows = [['구분', '항목', '예산(원)', '사용(원)', '잔액(원)', '사용률']];
  var bold = [1], subs = [];
  var tail = function (n) { return ['=C' + n + '-D' + n, '=IF(C' + n + '=0,0,D' + n + '/C' + n + ')']; };
  var start = 2;
  items.forEach(function (x, i) {
    var n = rows.length + 1;
    rows.push([x.cat, x.item, Number(x.budget) || 0,
      '=SUMIFS(' + L + 'K:K,' + L + 'D:D,A' + n + ',' + L + 'E:E,B' + n + ',' + L + 'N:N,"정상")'].concat(tail(n)));
    var next = items[i + 1];
    if (!next || next.cat !== x.cat) {       // 구분이 바뀌는 자리에 소계 한 줄
      var s = rows.length + 1;
      rows.push([x.cat, '소계', '=SUM(C' + start + ':C' + n + ')', '=SUM(D' + start + ':D' + n + ')'].concat(tail(s)));
      bold.push(s); subs.push(s); start = s + 1;
    }
  });
  var t = rows.length + 1;
  var plus = function (c) { return subs.length ? '=' + subs.map(function (s) { return c + s; }).join('+') : 0; };
  rows.push(['합계', '', plus('C'), plus('D')].concat(tail(t)));
  bold.push(t);
  sh.getRange(1, 1, rows.length, 6).setValues(rows);
  bold.forEach(function (r) { sh.getRange(r, 1, 1, 6).setFontWeight('bold').setBackground(r === 1 ? '#DCEFE6' : '#F1F5F3'); });
  sh.getRange(2, 3, rows.length - 1, 3).setNumberFormat('#,##0');
  sh.getRange(2, 6, rows.length - 1, 1).setNumberFormat('0.0%');
  sh.setColumnWidth(1, 170); sh.setColumnWidth(2, 230);
  sh.setFrozenRows(1);
  return { ok: true };
}

// 영수증 붙임 용지: 지출 한 건에 한 장. 맨 위에 품의 제목·사용일·기안자와 영수증 번호를 적고
// 그 아래 영수증 사진을 붙여 PDF 한 개로 만든다. 번호는 [지출] 시트·엑셀의 번호와 같다.
// 임시 구글 문서를 만들어 PDF 로 바꾼 뒤 문서는 휴지통으로 보낸다(PDF 만 남는다).
function attach_(data) {
  var from = String(data.from || ''), to = String(data.to || '9999-99-99');
  var titles = data.titles || {}, drafter = String(data.drafter || '');
  var order = data.cats || [];               // 예산표에 적힌 구분 순서
  var rank = function (c) { var i = order.indexOf(c); return i < 0 ? order.length : i; };
  var recs = list_().recs.filter(function (r) { return !r.del && r.date >= from && r.date <= to; });
  if (!recs.length) return { ok: false, error: '그 기간에 지출이 없습니다' };
  // 품의(구분)끼리 모으고, 그 안에서는 번호 순
  recs.sort(function (a, b) { return rank(a.cat) - rank(b.cat) || a.no - b.no; });

  var dates = recs.map(function (r) { return r.date; }).sort();
  var first = dates[0], last = dates[dates.length - 1];
  var name = '영수증 붙임 용지_' + first + (last !== first ? '~' + last : '') + '_' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMdd-HHmm');
  var GREEN = '#DCEFE6', W = 523;            // A4(595pt)에서 좌우 여백 36pt 를 뺀 너비
  var fmt = function (n) { return String(Math.round(Number(n) || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, ','); };
  var amount = function (r) { return r.cur === 'KRW' ? fmt(r.amt) + '원' : Number(r.amt).toFixed(2) + ' ' + r.cur; };
  // 머리표 한 줄: [이름칸, 값, 이름칸, 값 …]. 이름칸은 초록 바탕.
  // 빈 칸은 글자 서식을 받지 않을 수 있어 실패해도 넘어간다.
  var headRow = function (body, cells, widths) {
    var t = body.appendTable([cells]);
    t.setBorderColor('#555555');
    widths.forEach(function (w, c) { t.setColumnWidth(c, w); });
    cells.forEach(function (_, c) {
      var cell = t.getCell(0, c), label = c % 2 === 0;
      cell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(5).setPaddingRight(5);
      if (label) cell.setBackgroundColor(GREEN);
      try { cell.editAsText().setFontSize(label ? 10 : 11).setBold(label); } catch (x) {}
    });
    return t;
  };

  var doc = DocumentApp.create(name);
  var body = doc.getBody();
  body.setPageWidth(595).setPageHeight(842);
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(36).setMarginRight(36);

  var photos = 0, missing = 0;
  recs.forEach(function (r, pi) {
    if (pi > 0) body.appendPageBreak();
    // 새 문서의 첫 빈 문단은 첫 장 제목으로 쓴다
    var title = pi === 0 ? body.getChild(0).asParagraph() : body.appendParagraph('');
    title.setText('영수증 붙임 용지');
    title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    title.editAsText().setBold(true).setFontSize(18);
    title.setSpacingAfter(8);

    var h = headRow(body, ['품의 제목', titles[r.cat] || ''], [80, W - 80]);
    try { h.getCell(0, 1).editAsText().setBold(true); } catch (x) {}
    headRow(body, ['사용일', r.date.replace(/-/g, '. ') + '.', '기안자', drafter], [80, 181, 80, W - 341]);
    var n = headRow(body, ['영수증 번호', 'No. ' + r.no, '항목', r.item], [80, 181, 80, W - 341]);
    try { n.getCell(0, 1).editAsText().setBold(true).setFontSize(13); } catch (x) {}
    headRow(body, ['지출처 · 내역', r.place + (r.memo ? '  (' + r.memo + ')' : '')], [80, W - 80]);
    // 외화 금액은 그대로 남기고, 원화는 [환산]과 [수수료 여유분 포함]을 따로 적는다(실제 청구액과 견주어 볼 수 있게)
    var foreign = r.cur !== 'KRW';
    var m = headRow(body, ['결제 금액', amount(r), '적용 환율', foreign ? '1 ' + r.cur + ' = ' + r.rate + '원' : '-'], [80, 181, 80, W - 341]);
    try { m.getCell(0, 1).editAsText().setBold(true).setFontSize(13); } catch (x) {}
    headRow(body, ['원화 환산', fmt(r.krw) + '원', '여유분 포함', foreign ? fmt(r.total) + '원 (+' + fmt(r.fee) + ')' : fmt(r.total) + '원'], [80, 181, 80, W - 341]);

    if (!r.links.length) {
      var none = body.appendParagraph('(영수증 사진 없음)');
      none.setSpacingBefore(14); none.editAsText().setFontSize(10).setBold(false);
      return;
    }
    // 사진: 1장이면 크게, 여러 장이면 한 줄에 2장씩
    var one = r.links.length === 1, two = r.links.length === 2;
    var maxW = one ? 440 : (W / 2 - 14), maxH = one ? 540 : (two ? 540 : 265);
    var gap = body.appendParagraph(''); gap.setSpacingAfter(0); gap.editAsText().setFontSize(6);
    for (var k = 0; k < r.links.length; k += 2) {
      var pair = r.links.slice(k, k + 2);
      var pt = body.appendTable([pair.map(function () { return ''; })]);
      pt.setBorderWidth(0);
      pair.forEach(function (link, j) {
        var para = pt.getCell(0, j).getChild(0).asParagraph();
        para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        try {
          var id = /\/d\/([\w-]+)/.exec(link) || /[?&]id=([\w-]+)/.exec(link);
          var img = para.appendInlineImage(DriveApp.getFileById(id[1]).getBlob());
          var s = Math.min(maxW / img.getWidth(), maxH / img.getHeight());
          img.setWidth(Math.round(img.getWidth() * s)).setHeight(Math.round(img.getHeight() * s));
          photos++;
        } catch (err) {
          para.setText('(사진을 읽지 못했습니다)'); missing++;
        }
      });
    }
  });

  doc.saveAndClose();
  var src = DriveApp.getFileById(doc.getId());
  var root = folder_();
  var it = root.getFoldersByName(ATTACH_FOLDER);
  var out = it.hasNext() ? it.next() : root.createFolder(ATTACH_FOLDER);
  var pdf = out.createFile(src.getAs('application/pdf').setName(name + '.pdf'));
  src.setTrashed(true);
  return { ok: true, url: pdf.getUrl(), name: pdf.getName(), pages: recs.length, count: recs.length, photos: photos, missing: missing };
}

function book_() {
  var files = DriveApp.getFilesByName(BOOK_NAME);
  return files.hasNext() ? SpreadsheetApp.open(files.next()) : SpreadsheetApp.create(BOOK_NAME);
}

function logSheet_() {
  var ss = book_();
  var sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) {
    // 새로 만든 통합문서의 빈 첫 시트를 [지출] 로 쓴다
    sh = ss.getSheets()[0];
    if (sh.getLastRow() > 0) sh = ss.insertSheet(SHEET_LOG); else sh.setName(SHEET_LOG);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#DCEFE6');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('G:G').setNumberFormat('#,##0.00');
    sh.getRange('I:I').setNumberFormat('#,##0');
    sh.getRange('J:K').setNumberFormat('#,##0');
    sh.getRange('O:O').setNumberFormat('yyyy-mm-dd hh:mm');
  }
  return sh;
}

function findRow_(sheet, id) {
  var n = sheet.getLastRow();
  if (n < 2 || !id) return 0;
  var ids = sheet.getRange(2, COL_ID, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function folder_() {
  var it = DriveApp.getFoldersByName(RECEIPT_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(RECEIPT_FOLDER);
}

function dayFolder_(date) {
  var root = folder_();
  var name = String(date || '날짜없음');
  var it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
