/**
 * 記帳 Ledger → Google Sheet 同步（Apps Script Web App）
 * ------------------------------------------------------------------
 * 安裝：
 *   1. 開啟你的試算表 → 擴充功能 → Apps Script
 *   2. 把這個檔案的內容整個貼進 Code.gs（覆蓋原本的）
 *   3. 部署 → 新增部署作業 → 類型「網頁應用程式」
 *        執行身分：我
 *        誰可以存取：任何人
 *   4. 複製 /exec 結尾的網址 → 貼到 App 的「設定 → Google Sheet 同步」
 *
 * 行為：每次記帳後 App 會把「全部資料」POST 過來，這支腳本把各工作表
 *       整個重建（full mirror）。所以 Sheet 上不要手改資料，會被蓋掉；
 *       要加公式／圖表請另開工作表用 QUERY／IMPORTRANGE 參照。
 */

var TABS = {
  items:   '支出明細',
  trips:   '旅遊花費紀錄',
  tpls:    '固定轉帳設定',
  bills:   '信用卡繳款',
  accounts:'帳戶設定',
  cards:   '卡片額度',
  loans:   '學貸進度',
  summary: '月結摘要',
  meta:    '同步狀態'
};

function doGet() {
  return ContentService.createTextOutput('ledger sync ok').setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    writeItems_(ss, d);
    writeTrips_(ss, d);
    writeTpls_(ss, d);
    writeBills_(ss, d);
    writeAccounts_(ss, d);
    writeCards_(ss, d);
    writeLoans_(ss, d);
    writeSummary_(ss, d);
    writeMeta_(ss, d);

    return ContentService.createTextOutput(JSON.stringify({ok: true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/* ───────────── 工具 ───────────── */

function sheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/** 清掉整張表後寫入 header + rows，並套用格式 */
function dump_(ss, name, header, rows, opts) {
  opts = opts || {};
  var sh = sheet_(ss, name);
  sh.clear();
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#f2efe8');
  if (rows.length) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
  sh.setFrozenRows(1);

  // 金額欄位千分位
  (opts.money || []).forEach(function (col) {
    if (rows.length) sh.getRange(2, col, rows.length, 1).setNumberFormat('#,##0');
  });
  // 百分比欄位
  (opts.pct || []).forEach(function (col) {
    if (rows.length) sh.getRange(2, col, rows.length, 1).setNumberFormat('0.0%');
  });
  // 日期欄位
  (opts.date || []).forEach(function (col) {
    if (rows.length) sh.getRange(2, col, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  });

  for (var i = 1; i <= header.length; i++) sh.autoResizeColumn(i);
  return sh;
}

function payName_(d, key) {
  var list = d.pays || [];
  for (var i = 0; i < list.length; i++) if (list[i].k === key) return list[i].n;
  return key || '—';
}
function accName_(d, key) {
  var list = d.accounts || [];
  for (var i = 0; i < list.length; i++) if (list[i].k === key) return list[i].n;
  return key || '';
}
function cardName_(d, key) {
  var list = d.cards || [];
  for (var i = 0; i < list.length; i++) if (list[i].k === key) return list[i].n;
  return key || '';
}

/* ───────────── 各工作表 ───────────── */

/** 支出明細：一筆支出一列。這是整個資料庫的事實表（fact table） */
function writeItems_(ss, d) {
  var rows = (d.items || []).slice().sort(function (a, b) {
    return String(a.date).localeCompare(String(b.date)) || (a.id - b.id);
  }).map(function (x) {
    return [
      x.id,
      x.date,
      String(x.date).slice(0, 7),
      x.cat,
      Number(x.amount) || 0,
      payName_(d, x.pay),
      x.pay || '',
      x.note || '',
      x.auto ? '固定轉帳' : '手動',
      x.auto || ''
    ];
  });
  dump_(ss, TABS.items,
    ['id', '日期', '月份', '分類', '金額', '支付方式', '支付代碼', '備註', '來源', '自動記入鍵'],
    rows, {money: [5], date: [2]});
}

/** 旅遊花費紀錄 */
function writeTrips_(ss, d) {
  var rows = (d.trips || []).slice().sort(function (a, b) {
    return String(a.date).localeCompare(String(b.date)) || (a.id - b.id);
  }).map(function (x) {
    return [x.id, x.date, x.trip || '', x.cat, x.cur || 'TWD',
            Number(x.amount) || 0, Number(x.twd) || 0, x.note || ''];
  });
  dump_(ss, TABS.trips,
    ['id', '日期', '旅程', '分類', '幣別', '原幣金額', '台幣金額', '備註'],
    rows, {money: [6, 7], date: [2]});
}

/** 固定轉帳設定：每月自動記入的樣板 */
function writeTpls_(ss, d) {
  var rows = (d.tpls || []).map(function (t) {
    return [t.id, t.n, Number(t.amt) || 0, t.cat, '永豐大戶',
            accName_(d, t.to), Number(t.day) || 1, t.on ? '啟用' : '停用'];
  });
  dump_(ss, TABS.tpls,
    ['樣板id', '項目', '每月金額', '分類', '轉出帳戶', '轉入帳戶', '執行日', '狀態'],
    rows, {money: [3]});
}

/** 信用卡繳款：每月 × 每卡一列，含手動勾選的繳款狀態 */
function writeBills_(ss, d) {
  var bills = d.bills || {};
  var cards = d.cards || [];
  var months = Object.keys(bills).sort();
  var rows = [];
  months.forEach(function (M) {
    cards.forEach(function (c) {
      var r = bills[M][c.k];
      if (!r) return;
      if (!Number(r.amt) && !r.paid) return;   // 完全沒用到的就不寫
      rows.push([
        M, c.n, accName_(d, c.payFrom),
        Number(r.amt) || 0,
        r.paid ? '✓ 已繳' : '未繳',
        r.ts ? new Date(r.ts) : '',
        c.by || '', c.debit || ''
      ]);
    });
  });
  dump_(ss, TABS.bills,
    ['月份', '卡別', '扣款帳戶', '帳單金額', '繳款狀態', '勾選時間', '最慢轉入日', '銀行扣款日'],
    rows, {money: [4]});
}

/** 帳戶設定（資產面下一步會擴充） */
function writeAccounts_(ss, d) {
  var roleName = {main: '主帳戶 / 薪轉', loan: '學貸扣款戶', cardpay: '信用卡代扣戶'};
  var rows = (d.accounts || []).map(function (a) {
    return [a.k, a.n, roleName[a.role] || a.role,
            a.inBy || '', a.debit || '', a.memo || ''];
  });
  dump_(ss, TABS.accounts,
    ['帳戶代碼', '帳戶名稱', '角色', '最慢轉入日', '銀行扣款日', '說明'],
    rows, {});
}

/** 卡片額度與本月使用 */
function writeCards_(ss, d) {
  var items = d.items || [];
  var months = {};
  items.forEach(function (x) { months[String(x.date).slice(0, 7)] = 1; });
  var monthList = Object.keys(months).sort();
  var lastM = monthList.length ? monthList[monthList.length - 1] : '';

  var payToCard = {};
  (d.pays || []).forEach(function (p) { if (p.card) payToCard[p.k] = p.card; });

  var rows = (d.cards || []).map(function (c) {
    var used = 0;
    items.forEach(function (x) {
      if (String(x.date).slice(0, 7) === lastM && payToCard[x.pay] === c.k) used += Number(x.amount) || 0;
    });
    var lim = Number(c.limit) || 0;
    var avail = Math.max(0, lim - used);
    return [c.k, c.n, accName_(d, c.payFrom), lim, used, avail,
            lim ? avail / lim : '', c.by || '', c.debit || '', lastM];
  });
  dump_(ss, TABS.cards,
    ['卡別代碼', '卡別', '扣款帳戶', '信用額度', '本月已刷', '可用餘額', '剩餘比例', '最慢轉入日', '銀行扣款日', '統計月份'],
    rows, {money: [4, 5, 6], pct: [7]});
}

/** 學貸進度 */
function writeLoans_(ss, d) {
  var loans = d.loans || {};
  var rows = Object.keys(loans).map(function (k) {
    var L = loans[k];
    return [k, L.n, accName_(d, L.acc), L.snapDate, Number(L.snapBal) || 0,
            Number(L.principal) || 0, Number(L.monthly) || 0, Number(L.debitDay) || '',
            Number(L.left) || 0, Number(L.paid) || 0,
            (Number(L.pct) || 0) / 100, Number(L.monthsLeft) || 0, L.estDone || ''];
  });
  dump_(ss, TABS.loans,
    ['代碼', '貸款', '扣款帳戶', '餘額快照日', '快照餘額', '原始本金', '每月還款',
     '扣款日', '目前剩餘', '已還金額', '已還比例', '剩餘期數', '預估清償'],
    rows, {money: [5, 6, 7, 9, 10], pct: [11], date: [4]});
}

/** 月結摘要：每月一列，分類小計 + 大戶結餘。拿來畫圖表最方便 */
function writeSummary_(ss, d) {
  var cats = (d.cats || []).map(function (c) { return c.k; });
  var header = ['月份', '總支出', '筆數'].concat(cats)
    .concat(['薪水', '大戶直接支出', '信用卡帳單', '已繳卡費', '大戶繳完後剩餘', '可花費比例']);
  var rows = (d.summary || []).map(function (s) {
    var r = [s.month, Number(s.total) || 0, Number(s.count) || 0];
    cats.forEach(function (k) { r.push(Number((s.cats || {})[k]) || 0); });
    var salary = Number(s.salary) || 0;
    r.push(salary, Number(s.dawhoOut) || 0, Number(s.billTotal) || 0,
           Number(s.billPaid) || 0, Number(s.dawhoLeft) || 0,
           salary ? (Number(s.dawhoLeft) || 0) / salary : '');
    return r;
  });
  var money = [2];
  for (var i = 0; i < cats.length; i++) money.push(4 + i);
  var base = 4 + cats.length;
  money = money.concat([base, base + 1, base + 2, base + 3, base + 4]);
  dump_(ss, TABS.summary, header, rows, {money: money, pct: [base + 5]});
}

/** 同步狀態 */
function writeMeta_(ss, d) {
  var sh = sheet_(ss, TABS.meta);
  sh.clear();
  sh.getRange(1, 1, 6, 2).setValues([
    ['最後同步時間', new Date()],
    ['App 產生時間', d.generatedAt || ''],
    ['payload 版本', d.v || ''],
    ['支出筆數', (d.items || []).length],
    ['旅遊筆數', (d.trips || []).length],
    ['固定轉帳項目', (d.tpls || []).length]
  ]);
  sh.getRange(1, 1, 6, 1).setFontWeight('bold');
  sh.getRange(1, 2).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.autoResizeColumn(1);
  sh.autoResizeColumn(2);
}
