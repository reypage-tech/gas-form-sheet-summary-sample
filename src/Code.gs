/**
 * フォーム回答取り込みの見本（架空データのみ）。
 * 送信時に受付番号・状態・担当・期限を付与し、一覧タブを組み直し、
 * 新着・期限切れをメールおよび Chat Webhook（任意）で通知する。
 */
var SHEETS = {
  RESPONSES: 'フォームの回答 1',
  SETTINGS: '設定',
  LIST: '一覧'
};

var COL = {
  TS: 'タイムスタンプ',
  TYPE: '種別',
  RECEIPT: '受付番号',
  STATUS: '状態',
  ASSIGNEE: '担当',
  DEADLINE: '期限'
};

var STATUS = {
  OPEN: '未対応',
  PROGRESS: '対応中',
  DONE: '完了'
};

var STATUS_ORDER = { '未対応': 0, '対応中': 1, '完了': 2 };

var PROP = {
  CHAT_WEBHOOK: 'CHAT_WEBHOOK_URL',
  NOTIFY_EMAIL: 'NOTIFY_EMAIL',
  DRY_RUN: 'DRY_RUN',
  RECEIPT_SEQ: 'RECEIPT_SEQ',
  RECEIPT_YEAR: 'RECEIPT_YEAR'
};

var MANAGED = [COL.RECEIPT, COL.STATUS, COL.ASSIGNEE, COL.DEADLINE];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('申込集計')
    .addItem('一覧を作り直す', 'rebuildListView')
    .addItem('期限切れの通知を送る', 'notifyOverdue')
    .addItem('未付与行を一括補完', 'enrichAllResponses')
    .addItem('設定シートのひな型', 'createSettingsTemplate')
    .addToUi();
}

function onFormSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = responseSheet_(ss);
  var row = e.range.getRow();
  enrichResponseRow_(ss, sh, row);
  var rec = rowRecord_(sh, row);
  notifyNew_(rec);
}

/** 受付番号が空の回答行すべてに、受付番号・状態・担当・期限を付与する（テスト・一括取り込み用）。 */
function enrichAllResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = responseSheet_(ss);
  var last = sh.getLastRow();
  if (last < 2) return { enriched: 0 };
  var header = headerMap_(sh);
  if (header[COL.RECEIPT] == null) ensureManagedColumns_(sh, header);
  header = headerMap_(sh);
  var idxReceipt = header[COL.RECEIPT];
  var n = 0;
  for (var r = 2; r <= last; r++) {
    var cur = sh.getRange(r, idxReceipt + 1).getValue();
    if (String(cur).trim()) continue;
    enrichResponseRow_(ss, sh, r);
    n++;
  }
  return { enriched: n };
}

function enrichResponseRow_(ss, sh, row) {
  var settings = readSettings_(ss);
  var header = headerMap_(sh);
  var vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  var type = String(vals[header[COL.TYPE]] || '').trim();
  var rule = settings[type] || settings['__default'] || { days: 5, assignee: '' };
  var ts = vals[header[COL.TS]];
  var received = toDateOnly_(ts) || today_();
  var deadline = addWeekdayDays_(received, rule.days);
  var receipt = nextReceiptNumber_();
  if (header[COL.RECEIPT] == null) ensureManagedColumns_(sh, header);
  header = headerMap_(sh);
  sh.getRange(row, header[COL.RECEIPT] + 1).setValue(receipt);
  sh.getRange(row, header[COL.STATUS] + 1).setValue(STATUS.OPEN);
  sh.getRange(row, header[COL.ASSIGNEE] + 1).setValue(rule.assignee);
  sh.getRange(row, header[COL.DEADLINE] + 1).setValue(deadline);
}

function rebuildListView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = responseSheet_(ss);
  var rows = readResponseRecords_(sh);
  rows.sort(compareListRows_);
  var outSh = ss.getSheetByName(SHEETS.LIST) || ss.insertSheet(SHEETS.LIST);
  outSh.clearContents();
  var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].slice();
  if (header.indexOf('期限切れ') < 0) header.push('期限切れ');
  var today = today_();
  var body = rows.map(function (rec) {
    var line = rec.values.slice();
    var overdue = isOverdue_(rec.deadline, rec.status, today);
    line.push(overdue ? '※' : '');
    return line;
  });
  var all = [header].concat(body);
  outSh.getRange(1, 1, all.length, all[0].length).setValues(all);
  outSh.getRange(1, 1, 1, all[0].length).setFontWeight('bold');
  outSh.setFrozenRows(1);
  return { listed: body.length, overdueMarked: body.filter(function (r) { return r[r.length - 1] === '※'; }).length };
}

function notifyOverdue() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rows = readResponseRecords_(responseSheet_(ss));
  var today = today_();
  var targets = rows.filter(function (r) { return isOverdue_(r.deadline, r.status, today); });
  if (!targets.length) return { sent: 0 };
  var lines = targets.map(formatNotifyLine_);
  var subject = '[申込集計] 期限切れ ' + targets.length + '件';
  var body = lines.join('\n');
  dispatchNotification_(subject, body, 'overdue');
  return { sent: targets.length };
}

function notifyNew_(rec) {
  if (!rec) return;
  var subject = '[申込集計] 新着 ' + rec.receipt;
  var body = formatNotifyLine_(rec);
  dispatchNotification_(subject, body, 'new');
}

function dispatchNotification_(subject, body, kind) {
  var props = PropertiesService.getScriptProperties();
  var dry = String(props.getProperty(PROP.DRY_RUN) || '').toLowerCase();
  var isDry = dry === '1' || dry === 'true' || dry === 'yes';
  var email = String(props.getProperty(PROP.NOTIFY_EMAIL) || '').trim();
  var webhook = String(props.getProperty(PROP.CHAT_WEBHOOK) || '').trim();
  var payload = subject + '\n' + body;
  if (isDry) {
    console.log('DRY_RUN notify', kind, payload);
    return;
  }
  if (email) {
    MailApp.sendEmail(email, subject, body);
  }
  if (webhook) {
    var chatResp = UrlFetchApp.fetch(webhook, {
      method: 'post',
      contentType: 'application/json; charset=UTF-8',
      payload: JSON.stringify({ text: payload }),
      muteHttpExceptions: true
    });
    console.log('Chat webhook HTTP', chatResp.getResponseCode());
  }
}

function readSettings_(ss) {
  var sh = ss.getSheetByName(SHEETS.SETTINGS);
  if (!sh) throw new Error('設定シートがありません。メニューからひな型を作成してください。');
  var rows = sh.getDataRange().getValues().slice(1);
  var map = { __default: { days: 5, assignee: '' } };
  rows.forEach(function (r) {
    var type = String(r[0]).trim();
    if (!type) return;
    map[type] = { days: parseInt(r[1], 10) || 5, assignee: String(r[2] || '').trim() };
  });
  return map;
}

function readResponseRecords_(sh) {
  var header = headerMap_(sh);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  return values.map(function (row, i) {
    return rowRecordFromValues_(row, header, i + 2);
  }).filter(function (r) { return r.receipt; });
}

function rowRecord_(sh, row) {
  var header = headerMap_(sh);
  var vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  return rowRecordFromValues_(vals, header, row);
}

function rowRecordFromValues_(vals, header, rowNum) {
  return {
    row: rowNum,
    values: vals,
    receipt: String(vals[header[COL.RECEIPT]] || '').trim(),
    status: String(vals[header[COL.STATUS]] || '').trim() || STATUS.OPEN,
    assignee: String(vals[header[COL.ASSIGNEE]] || '').trim(),
    deadline: vals[header[COL.DEADLINE]],
    type: String(vals[header[COL.TYPE]] || '').trim(),
    title: header['件名'] != null ? String(vals[header['件名']] || '').trim() : ''
  };
}

function compareListRows_(a, b) {
  var sa = STATUS_ORDER[a.status] != null ? STATUS_ORDER[a.status] : 9;
  var sb = STATUS_ORDER[b.status] != null ? STATUS_ORDER[b.status] : 9;
  if (sa !== sb) return sa - sb;
  if (a.assignee !== b.assignee) return a.assignee < b.assignee ? -1 : 1;
  var da = toDateOnly_(a.deadline);
  var db = toDateOnly_(b.deadline);
  if (!da && !db) return 0;
  if (!da) return 1;
  if (!db) return -1;
  return da - db;
}

function isOverdue_(deadline, status, today) {
  if (status === STATUS.DONE) return false;
  var d = toDateOnly_(deadline);
  var t = toDateOnly_(today);
  if (!d || !t) return false;
  return d < t;
}

function formatNotifyLine_(rec) {
  var dl = toDateOnly_(rec.deadline);
  var dlStr = dl ? Utilities.formatDate(dl, 'Asia/Tokyo', 'yyyy-MM-dd') : '';
  return rec.receipt + ' / ' + rec.type + ' / ' + rec.title + ' / 担当:' + rec.assignee + ' / 期限:' + dlStr + ' / 状態:' + rec.status;
}

/** 受付日を起点に、土日を除いて N 日分だけ日付を進める（受付日は 0 日目）。 */
function addWeekdayDays_(start, days) {
  var d = toDateOnly_(start);
  if (!d) d = today_();
  var n = Math.max(0, parseInt(days, 10) || 0);
  var added = 0;
  while (added < n) {
    d.setDate(d.getDate() + 1);
    var dow = d.getDay();
    if (dow !== 0 && dow !== 6) added++;
  }
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function nextReceiptNumber_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var props = PropertiesService.getScriptProperties();
    var y = today_().getFullYear();
    var storedYear = parseInt(props.getProperty(PROP.RECEIPT_YEAR) || '0', 10);
    var cur = parseInt(props.getProperty(PROP.RECEIPT_SEQ) || '0', 10);
    if (storedYear !== y) {
      cur = 0;
      props.setProperty(PROP.RECEIPT_YEAR, String(y));
    }
    var next = cur + 1;
    props.setProperty(PROP.RECEIPT_SEQ, String(next));
    return 'R-' + y + '-' + ('0000' + next).slice(-4);
  } finally {
    lock.releaseLock();
  }
}

function responseSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.RESPONSES);
  if (!sh) throw new Error('回答シート「' + SHEETS.RESPONSES + '」が見つかりません。');
  return sh;
}

function headerMap_(sh) {
  var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var map = {};
  header.forEach(function (h, i) { map[String(h).trim()] = i; });
  return map;
}

function ensureManagedColumns_(sh, header) {
  var lastCol = sh.getLastColumn();
  MANAGED.forEach(function (name) {
    if (header[name] == null) {
      lastCol++;
      sh.getRange(1, lastCol).setValue(name);
      header[name] = lastCol - 1;
    }
  });
}

function toDateOnly_(v) {
  if (!v && v !== 0) return null;
  if (v instanceof Date && !isNaN(v)) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  var s = String(v).trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  var d = new Date(s);
  return isNaN(d) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function today_() {
  if (typeof __TEST_TODAY !== 'undefined' && __TEST_TODAY) {
    return toDateOnly_(__TEST_TODAY);
  }
  return toDateOnly_(new Date());
}

function createSettingsTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.SETTINGS) || ss.insertSheet(SHEETS.SETTINGS);
  sh.clearContents();
  var rows = [['種別', '期限日数', '担当']].concat(SETTINGS_SAMPLE);
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold');
}

var SETTINGS_SAMPLE = [
  ['社内備品', 3, '佐藤 京子'],
  ['イベント申込', 7, '山田 恵'],
  ['アンケート回答', 5, '中村 翔'],
  ['研修受講', 10, '小林 美咲'],
  ['広報掲載依頼', 4, '加藤 直樹'],
  ['会議室利用', 2, '吉田 由美'],
  ['データ抽出', 6, '松本 健'],
  ['名刺発注', 5, '井上 彩'],
  ['社内報投稿', 4, '木村 拓也'],
  ['備品修理', 8, '林 真理']
];
