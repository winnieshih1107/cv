/**
 * 猛健樂記錄 — Google 試算表後端 (Google Apps Script)
 *
 * 記錄會寫進試算表裡名為「猛健樂記錄」的分頁，記帳寫進「記帳」分頁，
 * 原本的工作表不會被動到。
 * 試算表本身「不需要」開放共用，腳本是用你自己的身分執行的。
 *
 * 設定步驟：
 * 1. 開啟你要存記錄的試算表
 * 2. 上方選單「擴充功能」→「Apps Script」
 * 3. 把編輯器裡原本的內容全部刪掉，貼上這整份程式碼，按存檔
 * 4. 右上角「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *    執行身分：我　　　誰可以存取：任何人
 * 5. 複製產生的「網頁應用程式網址」(結尾是 /exec)
 * 6. 第一次會跳出授權畫面（「未命名的專案 (Unverified)」）：
 *    按 REVIEW PERMISSIONS → 選你的帳號 → 左下「進階」→「前往⋯（不安全）」→「允許」
 *    這是你自己寫的腳本，Google 對所有未發佈的個人腳本都會這樣標示。
 * 7. 回到記錄網頁，貼進「雲端試算表」欄位，按「測試連線」
 */

/**
 * 記錄要寫到哪份試算表。
 * 留空＝寫進「開啟這份腳本的那份試算表」，也就是你從試算表選單
 * 「擴充功能 → Apps Script」打開的那一份，通常不用改。
 * 只有在用獨立的 Apps Script 專案時，才需要填試算表網址中間那段 ID。
 */
var SPREADSHEET_ID = '';
var SHEET_NAME = '猛健樂記錄';
var EXPENSE_SHEET = '記帳';
var HEADERS = ['id', '日期', 'Day', '施打', '劑量(mg)', '第幾劑',
               '原始體重', '目標體重', '今日體重',
               '早餐', '午餐', '晚餐', '點心', '熱量粗估',
               '運動', '運動分鐘', '消耗kcal', '運動內容',
               '備註', '副作用', '更新時間'];
var EXPENSE_HEADERS = ['id', '日期', '分類', '項目', '金額', '備註', '更新時間'];

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'push') {
      if (body.records) writeAll(body.records);
      if (body.expenses) writeExpenses(body.expenses);
      return out(info({
        action: 'push',
        count: (body.records || []).length,
        expenseCount: (body.expenses || []).length
      }));
    }
    return out(info({ action: body.action || 'pull', records: readAll(), expenses: readExpenses() }));
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

/**
 * GET 也可以用，並支援 JSONP (?callback=xxx)。
 * 手機瀏覽器擋掉跨網域回應時，網頁會改用這條路讀資料。
 */
function doGet(e) {
  var payload;
  try {
    payload = info({ action: 'pull', records: readAll(), expenses: readExpenses() });
  } catch (err) {
    payload = { ok: false, error: String(err) };
  }
  var cb = e && e.parameter && e.parameter.callback;
  if (cb && /^[A-Za-z0-9_$]+$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return out(payload);
}

/** 回應裡附上試算表名稱與網址，網頁就不需要把試算表 ID 寫死 */
function info(obj) {
  obj.ok = true;
  try {
    var ss = book();
    obj.title = ss.getName();
    obj.url = ss.getUrl();
  } catch (e) {}
  return obj;
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function book() {
  if (SPREADSHEET_ID) {
    try { return SpreadsheetApp.openById(SPREADSHEET_ID); } catch (e) {}
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheetNamed(name) {
  var ss = book();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  return sh;
}

function sheet() { return sheetNamed(SHEET_NAME); }

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
      numOrBlank(r.kcalIn),
      (r.exTypes || []).join('、'),
      numOrBlank(r.exMins),
      numOrBlank(r.exKcal),
      r.exNote || '',
      r.note || '',
      (r.se || []).join('、'),
      r.updatedAt || ''
    ]);
  });
  sh.getRange(1, 1, rows.length, HEADERS.length).setValues(rows);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#f1f3f4');
  sh.setFrozenRows(1);
  sh.getRange(2, 2, Math.max(rows.length - 1, 1), 1).setNumberFormat('@');
  var wide = { '早餐': 1, '午餐': 1, '晚餐': 1, '點心': 1, '運動內容': 1, '備註': 1, '副作用': 1 };
  for (var c = 1; c <= HEADERS.length; c++) {
    sh.setColumnWidth(c, wide[HEADERS[c - 1]] ? 220 : 100);
  }
}

/** 讀回試算表內容（你也可以直接在試算表上編輯，再從網頁「從雲端下載」） */
/** 依標題列對應欄位，這樣加欄位或欄位順序不同的舊試算表都讀得回來 */
function columnMap(headerRow) {
  var map = {};
  for (var c = 0; c < headerRow.length; c++) {
    map[String(headerRow[c]).trim()] = c;
  }
  return map;
}

function readAll() {
  var sh = sheet();
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var m = columnMap(values[0]);
  var get = function (row, name) {
    return m[name] === undefined ? '' : row[m[name]];
  };
  var list = [];
  for (var i = 1; i < values.length; i++) {
    var v = values[i];
    if (!get(v, '日期')) continue;
    list.push({
      id: String(get(v, 'id') || ('sheet-' + i)),
      date: asDate(get(v, '日期')),
      day: Number(get(v, 'Day')) || 1,
      shot: String(get(v, '施打')).trim() !== '',
      dose: String(get(v, '劑量(mg)') || ''),
      shotNo: Number(get(v, '第幾劑')) || null,
      start: asNum(get(v, '原始體重')),
      goal: asNum(get(v, '目標體重')),
      weight: asNum(get(v, '今日體重')),
      b: String(get(v, '早餐') || ''),
      l: String(get(v, '午餐') || ''),
      d: String(get(v, '晚餐') || ''),
      s: String(get(v, '點心') || ''),
      kcalIn: asNum(get(v, '熱量粗估')),
      exTypes: splitList(get(v, '運動')),
      exMins: asNum(get(v, '運動分鐘')),
      exKcal: asNum(get(v, '消耗kcal')),
      exNote: String(get(v, '運動內容') || ''),
      note: String(get(v, '備註') || ''),
      se: splitList(get(v, '副作用')),
      updatedAt: String(get(v, '更新時間') || '')
    });
  }
  return list;
}

function splitList(v) {
  return String(v || '').split(/[、,，]+/).filter(function (x) { return x.trim() !== ''; });
}

/** 記帳：以網頁的資料為準，整份覆寫「記帳」分頁 */
function writeExpenses(list) {
  var sh = sheetNamed(EXPENSE_SHEET);
  sh.clear();
  var rows = [EXPENSE_HEADERS];
  list.forEach(function (e) {
    rows.push([
      e.id || '',
      e.date || '',
      e.cat || '',
      e.item || '',
      numOrBlank(e.amount),
      e.note || '',
      e.updatedAt || ''
    ]);
  });
  sh.getRange(1, 1, rows.length, EXPENSE_HEADERS.length).setValues(rows);
  sh.getRange(1, 1, 1, EXPENSE_HEADERS.length).setFontWeight('bold').setBackground('#f1f3f4');
  sh.setFrozenRows(1);
  if (rows.length > 1) {
    sh.getRange(2, 5, rows.length - 1, 1).setNumberFormat('$#,##0');
  }
  sh.setColumnWidth(4, 260);
  sh.setColumnWidth(6, 180);
}

function readExpenses() {
  var sh = sheetNamed(EXPENSE_SHEET);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var m = columnMap(values[0]);
  var get = function (row, name) {
    return m[name] === undefined ? '' : row[m[name]];
  };
  var list = [];
  for (var i = 1; i < values.length; i++) {
    var v = values[i];
    if (!get(v, '日期')) continue;
    list.push({
      id: String(get(v, 'id') || ('sheet-exp-' + i)),
      date: asDate(get(v, '日期')),
      cat: String(get(v, '分類') || '其他'),
      item: String(get(v, '項目') || ''),
      amount: asNum(get(v, '金額')) || 0,
      note: String(get(v, '備註') || ''),
      updatedAt: String(get(v, '更新時間') || '')
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
