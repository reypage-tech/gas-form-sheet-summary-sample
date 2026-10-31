// Node で GAS サービスをスタブし src/Code.gs のロジックを検証する。実行: TZ=Asia/Tokyo node tests/local_check.js
const fs = require('fs');
const vm = require('vm');

function assert(cond, msg) {
  if (!cond) throw new Error(msg || 'assert failed');
}

function cellDateYmd(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2) + '-' + ('0' + v.getDate()).slice(-2);
  }
  var s = String(v).trim();
  var m = s.match(/^(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : s;
}

function parseCsv(path) {
  return fs.readFileSync(path, 'utf8').trim().split('\n').map(function (line) {
    return line.split(',');
  });
}

const tabs = {
  '設定': parseCsv('sample_data/settings.csv'),
  'フォームの回答 1': parseCsv('sample_data/responses.csv')
};

function makeRange(sheetData, r1, c1, r2, c2) {
  r2 = r2 == null ? r1 : r2;
  c2 = c2 == null ? c1 : c2;
  return {
    getRow: function () { return r1; },
    getValues: function () {
      const out = [];
      for (let r = r1; r <= r2; r++) {
        const row = sheetData[r - 1] || [];
        out.push([]);
        for (let c = c1; c <= c2; c++) {
          out[out.length - 1].push(row[c - 1] != null ? row[c - 1] : '');
        }
      }
      return out;
    },
    getValue: function () {
      const v = this.getValues();
      return v[0] && v[0][0];
    },
    setValue: function (val) {
      if (!sheetData[r1 - 1]) sheetData[r1 - 1] = [];
      sheetData[r1 - 1][c1 - 1] = val;
    },
    setValues: function (matrix) {
      matrix.forEach(function (rowVals, i) {
        const r = r1 + i;
        if (!sheetData[r - 1]) sheetData[r - 1] = [];
        rowVals.forEach(function (cell, j) {
          sheetData[r - 1][c1 - 1 + j] = cell;
        });
      });
    },
    setFontWeight: function () {}
  };
}

function sheet(name, tabStore) {
  const store = tabStore || tabs;
  const d = store[name];
  return {
    getDataRange: function () {
      return makeRange(d, 1, 1, d.length, d[0] ? d[0].length : 1);
    },
    getLastRow: function () { return d.length; },
    getLastColumn: function () { return d[0] ? d[0].length : 0; },
    getRange: function (r1, c1, numRows, numCols) {
      if (numRows === undefined && numCols === undefined) {
        return makeRange(d, r1, c1, r1, c1);
      }
      return makeRange(d, r1, c1, r1 + numRows - 1, c1 + numCols - 1);
    },
    clearContents: function () { d.length = 0; },
    appendRow: function (row) { d.push(row); },
    setFrozenRows: function () {}
  };
}

const props = { DRY_RUN: '1', RECEIPT_SEQ: '0' };
const mailLog = [];
const fetchLog = [];
const consoleLogs = [];
const lockLog = [];

function makeLockService(log) {
  return {
    getScriptLock: function () {
      return {
        waitLock: function () { log.push('waitLock'); },
        releaseLock: function () { log.push('releaseLock'); }
      };
    }
  };
}

function makeUrlFetchApp(fetchArr, responseCode) {
  var code = responseCode != null ? responseCode : 200;
  return {
    fetch: function (url, opts) {
      fetchArr.push({ url: url, opts: opts });
      return {
        getContentText: function () { return 'ok'; },
        getResponseCode: function () { return code; }
      };
    }
  };
}

const ss = {
  getSheetByName: function (n) { return tabs[n] ? sheet(n) : null; },
  insertSheet: function (n) {
    tabs[n] = [];
    return sheet(n);
  },
  toast: function (m) { console.log('TOAST', m); }
};

const ctx = {
  SpreadsheetApp: {
    getActiveSpreadsheet: function () { return ss; },
    getUi: function () {
      return {
        createMenu: function () {
          return {
            addItem: function () { return this; },
            addToUi: function () {}
          };
        }
      };
    }
  },
  PropertiesService: {
    getScriptProperties: function () {
      return {
        getProperty: function (k) { return props[k] != null ? String(props[k]) : null; },
        setProperty: function (k, v) { props[k] = String(v); }
      };
    }
  },
  MailApp: {
    sendEmail: function (to, subject, body) {
      mailLog.push({ to: to, subject: subject, body: body });
    }
  },
  UrlFetchApp: makeUrlFetchApp(fetchLog),
  LockService: makeLockService(lockLog),
  Utilities: {
    formatDate: function (d, tz, fmt) {
      if (fmt === 'yyyy-MM-dd') {
        return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
      }
      return String(d);
    }
  },
  console: {
    log: function () {
      consoleLogs.push(Array.prototype.slice.call(arguments).join(' '));
      console.log.apply(console, arguments);
    }
  },
  __TEST_TODAY: new Date(2026, 9, 20)
};

vm.createContext(ctx);
vm.runInContext(fs.readFileSync('src/Code.gs', 'utf8'), ctx);

function loadGas(customTabs, opts) {
  opts = opts || {};
  const t = customTabs;
  const p = Object.assign({ DRY_RUN: '1', RECEIPT_SEQ: '0', RECEIPT_YEAR: '2026' }, opts.props || {});
  const logs = [];
  const mail = [];
  const fetch = [];
  const locks = [];
  const ssLocal = {
    getSheetByName: function (n) { return t[n] ? sheet(n, t) : null; },
    insertSheet: function (n) {
      t[n] = [];
      return sheet(n, t);
    },
    toast: function (m) { console.log('TOAST', m); }
  };
  const c = {
    SpreadsheetApp: {
      getActiveSpreadsheet: function () { return ssLocal; },
      getUi: function () {
        return {
          createMenu: function () {
            return { addItem: function () { return this; }, addToUi: function () {} };
          }
        };
      }
    },
    PropertiesService: {
      getScriptProperties: function () {
        return {
          getProperty: function (k) { return p[k] != null ? String(p[k]) : null; },
          setProperty: function (k, v) { p[k] = String(v); }
        };
      }
    },
    MailApp: { sendEmail: function () { mail.push(Array.prototype.slice.call(arguments)); } },
    UrlFetchApp: makeUrlFetchApp(fetch, opts.httpCode),
    LockService: makeLockService(locks),
    Utilities: ctx.Utilities,
    console: {
      log: function () {
        logs.push(Array.prototype.slice.call(arguments).join(' '));
      }
    },
    __TEST_TODAY: new Date(2026, 9, 20)
  };
  vm.createContext(c);
  vm.runInContext(fs.readFileSync('src/Code.gs', 'utf8'), c);
  return { ctx: c, tabs: t, props: p, consoleLogs: logs, mailLog: mail, fetchLog: fetch, lockLog: locks };
}

const enrich = ctx.enrichAllResponses();
assert(enrich.enriched === 20, 'expected 20 enriched rows, got ' + enrich.enriched);

const resp = tabs['フォームの回答 1'];
const dataRows = resp.slice(1);
assert(dataRows.length === 20, 'expected 20 data rows');
dataRows.forEach(function (row, i) {
  assert(row[6] && String(row[6]).indexOf('R-2026-') === 0, 'row ' + (i + 2) + ' missing receipt');
  assert(row[7] === '未対応', 'row ' + (i + 2) + ' status');
  assert(row[8], 'row ' + (i + 2) + ' assignee');
  assert(row[9], 'row ' + (i + 2) + ' deadline');
});

const list = ctx.rebuildListView();
assert(list.listed === 20, 'listed count');
assert(list.overdueMarked === 15, 'expected 15 overdue marks at 2026-10-20, got ' + list.overdueMarked);

const listTab = tabs['一覧'];
assert(listTab && listTab.length === 21, 'list header + 20 rows');
const overdueCol = listTab[0].length - 1;
const marked = listTab.slice(1).filter(function (r) { return r[overdueCol] === '※'; }).length;
assert(marked === list.overdueMarked, 'overdue column consistent');

var receiptCol = listTab[0].indexOf('受付番号');
assert(receiptCol >= 0, 'list header must include 受付番号');
var listReceipts = listTab.slice(1).map(function (r) { return r[receiptCol]; });
var expectedListHead = ['R-2026-0003','R-2026-0013','R-2026-0008','R-2026-0017','R-2026-0001','R-2026-0011','R-2026-0005','R-2026-0014'];
assert(listReceipts.length >= expectedListHead.length, 'list row count');
for (var ei = 0; ei < expectedListHead.length; ei++) {
  assert(listReceipts[ei] === expectedListHead[ei],
    'list sort (状態・担当・期限) row ' + (ei + 2) + ': expected ' + expectedListHead[ei] + ', got ' + listReceipts[ei]);
}
assert(listReceipts[18] === 'R-2026-0010' && listReceipts[19] === 'R-2026-0020',
  'list tail 林 真理 by deadline: got ' + listReceipts.slice(18).join(', '));

assert(mailLog.length === 0, 'DRY_RUN must not send mail');

const overdue = ctx.notifyOverdue();
assert(overdue.sent === list.overdueMarked, 'overdue notify count');
assert(fetchLog.length === 0, 'empty webhook must not fetch');
assert(consoleLogs.some(function (l) { return l.indexOf('DRY_RUN notify') >= 0 && l.indexOf('overdue') >= 0; }), 'DRY_RUN overdue log');

ctx.notifyNew_(ctx.rowRecordFromValues_(resp[1], ctx.headerMap_(sheet('フォームの回答 1')), 2));
assert(consoleLogs.some(function (l) { return l.indexOf('DRY_RUN notify') >= 0 && l.indexOf('new') >= 0; }), 'DRY_RUN new log');

props.RECEIPT_SEQ = '450';
props.RECEIPT_YEAR = '2026';
ctx.__TEST_TODAY = new Date(2027, 0, 5);
var crossYear = ctx.nextReceiptNumber_();
assert(crossYear === 'R-2027-0001', 'year rollover receipt, got ' + crossYear);
assert(props.RECEIPT_SEQ === '1', 'seq after rollover');
assert(props.RECEIPT_YEAR === '2027', 'stored year after rollover');
ctx.__TEST_TODAY = new Date(2027, 0, 5);
var sameYear = ctx.nextReceiptNumber_();
assert(sameYear === 'R-2027-0002', 'same-year increment, got ' + sameYear);
assert(lockLog.length >= 4, 'nextReceiptNumber_ must wait and release lock, got ' + lockLog.length);
for (var li = 0; li < lockLog.length; li += 2) {
  assert(lockLog[li] === 'waitLock' && lockLog[li + 1] === 'releaseLock',
    'lock pairs must be waitLock then releaseLock at ' + li + ', got ' + lockLog.slice(li, li + 2));
}

(function testChatWebhookResponseLog() {
  const t = {
    '設定': parseCsv('sample_data/settings.csv'),
    'フォームの回答 1': parseCsv('sample_data/responses.csv').slice(0, 2)
  };
  const g = loadGas(t, {
    props: {
      DRY_RUN: '',
      CHAT_WEBHOOK_URL: 'https://chat.example/hook',
      NOTIFY_EMAIL: '',
      RECEIPT_SEQ: '0',
      RECEIPT_YEAR: '2026'
    },
    httpCode: 503
  });
  g.ctx.enrichAllResponses();
  g.ctx.dispatchNotification_('[申込集計] test', 'line', 'new');
  assert(g.fetchLog.length === 1, 'webhook must fetch once, got ' + g.fetchLog.length);
  assert(g.fetchLog[0].opts.muteHttpExceptions === true, 'Chat fetch must set muteHttpExceptions');
  assert(g.consoleLogs.some(function (l) { return l === 'Chat webhook HTTP 503'; }),
    'non-200 response must be logged, logs=' + JSON.stringify(g.consoleLogs));
})();

(function testRawSheetEnrichAll() {
  const t = {
    '設定': parseCsv('sample_data/settings.csv'),
    'フォームの回答 1': [
      ['タイムスタンプ', '氏名', '所属', '種別', '件名', '内容'],
      ['2026-10-01 09:15:00', '架空 太郎', '総務部', '社内備品', 'コピー用紙', '500枚'],
      ['2026-10-02 11:05:00', '架空 次郎', '開発部', '見本外種別', '件名', '本文']
    ]
  };
  const g = loadGas(t);
  const h0 = g.tabs['フォームの回答 1'][0];
  assert(h0.length === 6 && h0.indexOf('受付番号') < 0, 'raw header has no managed cols');
  const result = g.ctx.enrichAllResponses();
  assert(result.enriched === 2, 'raw sheet enrich count, got ' + result.enriched);
  const h1 = g.tabs['フォームの回答 1'][0];
  assert(h1.indexOf('受付番号') >= 0 && h1.indexOf('状態') >= 0 && h1.indexOf('担当') >= 0 && h1.indexOf('期限') >= 0,
    'managed headers appended');
  const rowKnown = g.tabs['フォームの回答 1'][1];
  const rowUnknown = g.tabs['フォームの回答 1'][2];
  assert(rowKnown[6] === 'R-2026-0001', 'known type receipt, got ' + rowKnown[6]);
  assert(rowKnown[7] === '未対応', 'known type status');
  assert(rowKnown[8] === '佐藤 京子', 'known type assignee');
  assert(cellDateYmd(rowKnown[9]) === '2026-10-06', '社内備品 3 weekdays from 2026-10-01, got ' + cellDateYmd(rowKnown[9]));
  assert(rowUnknown[6] === 'R-2026-0002', 'unknown type receipt, got ' + rowUnknown[6]);
  assert(rowUnknown[7] === '未対応', 'unknown type status');
  assert(rowUnknown[8] === '', 'unknown type assignee must be empty, got ' + JSON.stringify(rowUnknown[8]));
  assert(cellDateYmd(rowUnknown[9]) === '2026-10-09', '__default 5 weekdays from 2026-10-02, got ' + cellDateYmd(rowUnknown[9]));
})();

(function testOnFormSubmitRange() {
  const t = {
    '設定': parseCsv('sample_data/settings.csv'),
    'フォームの回答 1': [
      ['タイムスタンプ', '氏名', '所属', '種別', '件名', '内容', '受付番号', '状態', '担当', '期限'],
      ['2026-10-03 09:00:00', '架空 花子', '営業部', 'イベント申込', '展示会', '出展', '', '', '', '']
    ]
  };
  const g = loadGas(t);
  g.ctx.onFormSubmit({ range: { getRow: function () { return 2; } } });
  const row = g.tabs['フォームの回答 1'][1];
  assert(row[6] === 'R-2026-0001', 'onFormSubmit receipt, got ' + row[6]);
  assert(row[7] === '未対応', 'onFormSubmit status');
  assert(row[8] === '山田 恵', 'onFormSubmit assignee');
  assert(cellDateYmd(row[9]) === '2026-10-13', 'イベント申込 7 weekdays from 2026-10-03, got ' + cellDateYmd(row[9]));
  assert(g.consoleLogs.some(function (l) {
    return l.indexOf('DRY_RUN notify') >= 0 && l.indexOf('new') >= 0 && l.indexOf('R-2026-0001') >= 0;
  }), 'onFormSubmit notifyNew DRY_RUN log');
})();

(function testReadSettingsMissing() {
  const t = {
    'フォームの回答 1': [
      ['タイムスタンプ', '種別', '受付番号', '状態', '担当', '期限'],
      ['2026-10-01 09:00:00', '社内備品', '', '', '', '']
    ]
  };
  const g = loadGas(t);
  let msg = '';
  try {
    g.ctx.readSettings_(g.ctx.SpreadsheetApp.getActiveSpreadsheet());
  } catch (e) {
    msg = e.message;
  }
  assert(msg.indexOf('設定シートがありません') >= 0, 'readSettings_ must throw when 設定 missing, got ' + JSON.stringify(msg));
})();

console.log('OK enrich=' + enrich.enriched + ' listed=' + list.listed + ' overdue=' + list.overdueMarked);
