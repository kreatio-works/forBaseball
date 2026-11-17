/**
 * 自主練カウンター 記録受信用 GAS（スプレッドシートに紐づけて使う）
 *
 * - アプリから届いた記録を、記録シート（いちばん左のシート）に1行ずつ追加する
 * - 届くたびに「週間集計」シートを作り直す（選手ごと・日ごとの一覧。ひとこと付き）
 *
 * 記録シートの列（A〜M）
 *   受信日時 / 名前 / 日付 / 種目 / 入力方式 / 開始 / 終了 / 所要秒 /
 *   音声カウント / 手動カウント / 合計 / 1日の目標 / ひとこと
 *
 * コードを直したあとは「デプロイを管理 → 鉛筆 → 新バージョン → デプロイ」で更新する
 * （「新しいデプロイ」にするとURLが変わり、アプリ側の設定をやり直すことになる）
 */

var SUMMARY_NAME = '週間集計';
var HEADER = ['受信日時', '名前', '日付', '種目', '入力方式', '開始', '終了', '所要秒',
              '音声カウント', '手動カウント', '合計', '1日の目標', 'ひとこと'];
var DONE_COLOR = '#d9ead3';   // 目標に届いたセルの色（うすい緑）

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);   // 同時に届いても1件ずつ処理する
  try {
    var sheet = recordSheet_();
    ensureHeader_(sheet);

    var d = JSON.parse(e.postData.contents);
    sheet.appendRow([
      new Date(),   // サーバー側の受信日時（改ざん防止の基準）
      text_(d.name), d.date, text_(d.menu), d.type, d.start, d.end, d.seconds,
      d.voice, d.manual, d.total, d.goal, text_(d.memo)
    ]);

    rebuildSummary();
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

function doGet() {
  return ContentService.createTextOutput('自主練記録：受付OK');
}

// スプレッドシートを開いたときに、手動で集計し直すメニューを出す
function onOpen() {
  SpreadsheetApp.getUi().createMenu('自主練')
    .addItem('週間集計を更新', 'rebuildSummary')
    .addToUi();
}

// 記録シート＝「週間集計」以外で、いちばん左にあるシート
function recordSheet_() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() !== SUMMARY_NAME) return sheets[i];
  }
  return SpreadsheetApp.getActiveSpreadsheet().insertSheet('記録', 0);
}

// 見出し行がなければ作る。古い見出し（ひとこと列なし）なら足す
function ensureHeader_(sheet) {
  if (sheet.getLastRow() === 0) { sheet.appendRow(HEADER); return; }
  if (sheet.getRange(1, HEADER.length).getValue() === '') {
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  }
}

// 「週間集計」シートを作り直す
function rebuildSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rec = recordSheet_();
  ensureHeader_(rec);
  var rows = rec.getLastRow() > 1
    ? rec.getRange(2, 1, rec.getLastRow() - 1, HEADER.length).getValues() : [];
  var out = buildSummary_(rows, ss.getSpreadsheetTimeZone());

  var sh = ss.getSheetByName(SUMMARY_NAME) || ss.insertSheet(SUMMARY_NAME, ss.getSheets().length);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.clear();
  var nRows = out.values.length, nCols = out.values[0].length;
  var range = sh.getRange(1, 1, nRows, nCols);
  // 日付の2列（週・日付）は文字のまま置く（勝手に日付の書式に変わらないように）
  sh.getRange(1, 1, nRows, 1).setNumberFormat('@');
  sh.getRange(1, 3, nRows, 1).setNumberFormat('@');
  range.setValues(out.values);
  range.setBackgrounds(out.backgrounds);
  sh.getRange(1, 1, 1, nCols).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(4);
  range.createFilter();                // 名前・週でしぼりこめる
}

/**
 * 記録の行から集計表を作る（シートを触らない計算だけの部分）
 * 1行 = 選手1人の1日。列 = 週（月曜の日付）/ 名前 / 日付 / 曜日 / 種目… / ひとこと
 * - 種目の数字は、その日の合計（「0回に戻す」のマイナス行も足し込むので、戻したあとは0になる）
 * - ひとことは、その日に最後に届いたもの
 * - 並びは、新しい週が上 → 名前順 → 日付順
 */
function buildSummary_(rows, tz) {
  var WD = '日月火水木金土';
  var menus = [];      // 種目名（はじめて出てきた順）
  var days = {};       // 'yyyy-MM-dd|名前' → {week, name, date, items:{種目:{total, goal}}, memo}

  rows.forEach(function (r) {
    var date = dateStr_(r[2], tz);
    if (!date) return;
    var name = String(r[1] || '');
    var key = date + '|' + name;
    var day = days[key] || (days[key] = { week: mondayOf_(date), name: name, date: date, items: {}, memo: '' });
    if (r[4] === 'メモ') { day.memo = String(r[12] || ''); return; }

    var menu = String(r[3] || '');
    if (!menu) return;
    if (menus.indexOf(menu) < 0) menus.push(menu);
    var it = day.items[menu] || (day.items[menu] = { total: 0, goal: 0 });
    it.total += Number(r[10]) || 0;
    if (Number(r[11]) > 0) it.goal = Number(r[11]);
  });

  var list = Object.keys(days).map(function (k) { return days[k]; });
  list.sort(function (a, b) {
    if (a.week !== b.week) return a.week < b.week ? 1 : -1;
    if (a.name !== b.name) return a.name < b.name ? -1 : 1;
    return a.date < b.date ? -1 : a.date > b.date ? 1 : 0;
  });

  var header = ['週（月曜）', '名前', '日付', '曜日'].concat(menus.map(text_), ['ひとこと']);
  var values = [header];
  var backgrounds = [header.map(function () { return '#efefef'; })];

  list.forEach(function (day) {
    var p = day.date.split('-');
    var row = [day.week, text_(day.name), day.date, WD.charAt(new Date(+p[0], +p[1] - 1, +p[2]).getDay())];
    var bg = ['#ffffff', '#ffffff', '#ffffff', '#ffffff'];
    menus.forEach(function (m) {
      var it = day.items[m];
      row.push(it ? it.total : '');
      bg.push(it && it.goal > 0 && it.total >= it.goal ? DONE_COLOR : '#ffffff');
    });
    row.push(text_(day.memo));
    bg.push('#ffffff');
    values.push(row);
    backgrounds.push(bg);
  });

  return { values: values, backgrounds: backgrounds };
}

// 自由入力の文字（名前・種目・ひとこと）が「=」などで始まっても、数式として動かないようにする
function text_(v) {
  var s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// 日付のセル（文字でも日付型でも）を 'yyyy-MM-dd' にそろえる
function dateStr_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var m = String(v || '').match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : '';
}

// その日をふくむ週の月曜日（'yyyy-MM-dd'）
function mondayOf_(date) {
  var p = date.split('-');
  var d = new Date(+p[0], +p[1] - 1, +p[2]);
  d.setDate(d.getDate() - ((d.getDay() + 6) % 7));
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}
