/**
 * 猛健樂記錄 — Google 試算表後端 (Google Apps Script)
 *
 * 設定步驟：
 * 1. 開一個新的 Google 試算表
 * 2. 上方選單「擴充功能」→「Apps Script」
 * 3. 把編輯器裡原本的內容全部刪掉，貼上這整份程式碼，按存檔
 * 4. 右上角「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *    執行身分：我　　　誰可以存取：任何人
 * 5. 複製產生的「網頁應用程式網址」(結尾是 /exec)
 * 6. 回到記錄網頁，貼進「雲端試算表」欄位，按「測試連線」
 */

var SHEET_NAME = '猛健樂記錄';
var HEADERS = ['id', '日期', 'Day', '施打', '劑量(mg)', '第幾劑',
               '原始體重', '目標體重', '今日體重',
               '早餐', '午餐', '晚餐', '點心', '備註', '副作用', '更新時間'];

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'push') {
      writeAll(body.records || []);
      return out({ ok: true, action: 'push', count: (body.records || []).length });
    }
    return out({ ok: true, action: body.action || 'pull', records: readAll() });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function doGet() {
  try {
    return out({ ok: true, action: 'pull', records: readAll() });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }
  return sh;
}

/** 以網頁的資料為準，整份覆寫試算表 */
function writeAll(records) {
  var sh = sheet();
  sh.clear();
  var rows = [HEADERS];
  records.forEach(function (r) {
    rows.push([
      r.id || '',
      r.date || '',
      r.day || '',
      r.shot ? '是' : '',
      r.dose || '',
      r.shotNo || '',
      numOrBlank(r.start),
      numOrBlank(r.goal),
      numOrBlank(r.weight),
      r.b || '',
      r.l || '',
      r.d || '',
      r.s || '',
      r.note || '',
      (r.se || []).join('、'),
      r.updatedAt || ''
    ]);
  });
  sh.getRange(1, 1, rows.length, HEADERS.length).setValues(rows);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#f1f3f4');
  sh.setFrozenRows(1);
  sh.getRange(2, 2, Math.max(rows.length - 1, 1), 1).setNumberFormat('@');
  for (var c = 1; c <= HEADERS.length; c++) {
    sh.setColumnWidth(c, c >= 10 && c <= 14 ? 220 : 100);
  }
}

/** 讀回試算表內容（你也可以直接在試算表上編輯，再從網頁「從雲端下載」） */
function readAll() {
  var sh = sheet();
  var values = sh.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < values.length; i++) {
    var v = values[i];
    if (!v[1]) continue;
    list.push({
      id: String(v[0] || ('sheet-' + i)),
      date: asDate(v[1]),
      day: Number(v[2]) || 1,
      shot: String(v[3]).trim() !== '',
      dose: String(v[4] || ''),
      shotNo: Number(v[5]) || null,
      start: asNum(v[6]),
      goal: asNum(v[7]),
      weight: asNum(v[8]),
      b: String(v[9] || ''),
      l: String(v[10] || ''),
      d: String(v[11] || ''),
      s: String(v[12] || ''),
      note: String(v[13] || ''),
      se: String(v[14] || '').split(/[、,，]+/).filter(function (x) { return x.trim() !== ''; }),
      updatedAt: String(v[15] || '')
    });
  }
  return list;
}

function asDate(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v).trim();
}

function asNum(v) {
  var n = parseFloat(v);
  return isNaN(n) ? null : n;
}

function numOrBlank(v) {
  return (v === null || v === undefined || v === '') ? '' : v;
}
