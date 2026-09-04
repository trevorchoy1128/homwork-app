/**
 * 交功課紀錄 — Google Sheet 同步端點 (Google Apps Script)
 *
 * 安裝步驟：
 *  1. 開一個新的 Google Sheet
 *  2. 選單「擴充功能」→「Apps Script」
 *  3. 刪除預設內容，貼上本檔案全部內容，按儲存
 *  4. 右上「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *     - 執行身份：我
 *     - 誰可以存取：任何人
 *  5. 按「部署」，授權後複製「網頁應用程式網址」(以 /exec 結尾)
 *  6. 在 iPhone App 的「設定」頁貼上該網址
 *
 * 之後每次修改本檔案，都要重新「部署 → 管理部署作業 → 編輯 → 新版本」才會生效。
 */

var SHEET_STUDENTS = '學生名單';
var SHEET_RECORDS  = '功課紀錄';
var MATRIX_PREFIX  = '總覽-';
var SHEET_STATE    = '_app_data';   // 隱藏工作表：儲存 App 全部資料，供多裝置同步

// ===== 定時摘要及自動備份設定 =====
// 在 Apps Script 編輯器上方選擇函數「setupTriggers」按「執行」一次即可啟用（會要求授權 Gmail 及 Drive）。
// 選擇「deleteTriggers」執行則全部停用。
var SUMMARY_EMAIL = '';                 // 每日未交摘要收件人，留空 = 部署本程式的 Google 帳戶
var WEEKLY_EMAILS = '';                 // 每週統計收件人，可多個以逗號分隔（例如班主任），留空 = 部署者
var DAILY_HOUR    = 8;                  // 每日摘要時間（0-23）
var WEEKLY_DAY    = 'FRIDAY';           // 每週統計星期（MONDAY … SUNDAY）
var WEEKLY_HOUR   = 16;
var BACKUP_FOLDER = '交功課紀錄備份';    // Google Drive 資料夾名稱
var BACKUP_KEEP   = 30;                 // 保留最近幾天的備份

// App 讀取：GET ?action=students（學生名單）或 ?action=state（App 全部資料）
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'ping';
  try {
    if (action === 'students') return json({ ok: true, students: readStudents() });
    if (action === 'state') return json({ ok: true, state: readState(SpreadsheetApp.getActiveSpreadsheet()) });
    return json({ ok: true, msg: '交功課紀錄同步端點運作中' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// App 同步所有資料：POST JSON
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ver = 0;
    if (data.state) {
      // 多裝置衝突檢查：雲端版本(ver)比這部裝置上次同步時看到的版本(base)新 → 交回 App 讓老師選擇
      var cur = readState(ss);
      var curVer = (cur && cur.ver) || 0;
      if (!data.force && curVer > (data.base || 0)) {
        return json({ ok: false, conflict: true, state: cur });
      }
      ver = Math.max(Date.now(), curVer + 1);   // 由伺服器分配，單調遞增
      data.state.ver = ver;
      data.state.savedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd HH:mm');
      writeState(ss, data.state);
    }
    writeRecords(ss, data.records || []);
    writeMatrices(ss, data.matrices || []);
    writeStudents(ss, data.students || []);
    var sheet = ss.getSheetByName(SHEET_RECORDS);
    sheet.getRange('N1').setValue('最後同步：' + (data.exportedAt || new Date()) + (data.teacher ? '（' + data.teacher + '老師）' : ''));
    return json({ ok: true, rows: (data.records || []).length, ver: ver });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getOrCreate(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// 把 App 資料以 JSON 文字分段存入隱藏工作表（每格最多 50,000 字，這裡每段 40,000）
// 每段前加一個 'x'，避免內容以 = 或 + 開頭被當成公式
function writeState(ss, state) {
  var sh = getOrCreate(ss, SHEET_STATE);
  sh.clearContents();
  var s = JSON.stringify(state), CH = 40000, rows = [];
  for (var i = 0; i < s.length; i += CH) rows.push(['x' + s.substr(i, CH)]);
  if (rows.length) sh.getRange(1, 1, rows.length, 1).setValues(rows);
  sh.getRange(1, 2).setValue(state.ver || 0);
  sh.getRange(1, 3).setValue('請勿修改此工作表：App 同步資料');
  try { sh.hideSheet(); } catch (err) {}
}

function readState(ss) {
  var sh = ss.getSheetByName(SHEET_STATE);
  if (!sh || sh.getLastRow() < 1) return null;
  var vals = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  var s = vals.map(function (r) { return String(r[0] || '').substr(1); }).join('');
  if (!s) return null;
  try { return JSON.parse(s); } catch (err) { return null; }
}

function fill(sheet, header, rows) {
  sheet.clearContents();
  var all = [header].concat(rows);
  var width = header.length;
  if (!width) return;
  all = all.map(function (r) { r = r.slice(); while (r.length < width) r.push(''); return r.slice(0, width); });
  sheet.getRange(1, 1, all.length, width).setValues(all);
  sheet.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setFrozenRows(1);
}

function writeRecords(ss, records) {
  var header = ['日期', '班別', '科目', '類別', '項目', '學號', '姓名', '電話', '狀態', '提醒次數', '提醒時間', '更新時間'];
  var rows = records.map(function (r) {
    return [r.date, r.cls, r.subject || '', r.tag || '功課', r.title, r.no, r.name, r.phone, r.status, r.reminds || '', r.remindedAt, r.updatedAt];
  });
  var sheet = getOrCreate(ss, SHEET_RECORDS);
  fill(sheet, header, rows);
  // 未完成紅色、補交橙色（各類別字眼）
  var col = sheet.getRange('I:I'), rules = [];
  ['未交', '欠帶', '未測', '未默', '未完成'].forEach(function (w) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(w).setBackground('#fde2e1').setFontColor('#c5221f').setRanges([col]).build());
  });
  ['補交', '遲交', '補帶', '補測', '補默', '補做'].forEach(function (w) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(w).setBackground('#fff0d6').setFontColor('#b06000').setRanges([col]).build());
  });
  sheet.setConditionalFormatRules(rules);
}

function writeMatrices(ss, matrices) {
  var keep = {};
  matrices.forEach(function (m) {
    var name = (MATRIX_PREFIX + m.cls).slice(0, 100);
    keep[name] = true;
    var sheet = getOrCreate(ss, name);
    fill(sheet, m.header, m.rows);
    if (m.rows.length && m.header.length > 3) {
      var range = sheet.getRange(2, 4, m.rows.length, m.header.length - 3);
      sheet.setConditionalFormatRules([
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('✗').setBackground('#fde2e1').setFontColor('#c5221f').setRanges([range]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('遲').setBackground('#fff0d6').setFontColor('#b06000').setRanges([range]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('✓').setFontColor('#188038').setRanges([range]).build()
      ]);
      sheet.setFrozenColumns(3);
    }
  });
  // 刪除已不存在班別的總覽表
  ss.getSheets().forEach(function (sh) {
    var n = sh.getName();
    if (n.indexOf(MATRIX_PREFIX) === 0 && !keep[n]) ss.deleteSheet(sh);
  });
}

function writeStudents(ss, students) {
  var header = ['班別', '學號', '姓名', '電話', '科目', '其他電話'];
  var sheet = getOrCreate(ss, SHEET_STUDENTS);
  var rows = students.map(function (s) { return [s.cls, s.no, s.name, s.phone, s.subjects || '', s.others || '']; });
  fill(sheet, header, rows);
  // 電話及學號以純文字儲存，避免前置 0 被刪去
  if (rows.length) sheet.getRange(2, 2, rows.length, 5).setNumberFormat('@');
}

// ===================== 定時摘要 / 自動備份 =====================
function setupTriggers() {
  deleteTriggers();
  ScriptApp.newTrigger('dailySummary').timeBased().everyDays(1).atHour(DAILY_HOUR).create();
  ScriptApp.newTrigger('weeklySummary').timeBased().onWeekDay(ScriptApp.WeekDay[WEEKLY_DAY]).atHour(WEEKLY_HOUR).create();
  ScriptApp.newTrigger('dailyBackup').timeBased().everyDays(1).atHour(2).create();
  Logger.log('已建立三個定時任務：每日摘要 ' + DAILY_HOUR + ' 時、每週統計 ' + WEEKLY_DAY + ' ' + WEEKLY_HOUR + ' 時、每日備份 2 時');
}
function deleteTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
}
function loadState() { return readState(SpreadsheetApp.getActiveSpreadsheet()); }
function todayStr() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function ownerEmail() { return Session.getEffectiveUser().getEmail(); }
function byNo(a, b) { return ((+a.no || 999) - (+b.no || 999)) || String(a.name).localeCompare(String(b.name)); }
function isPendingA(a) { return !!a.due && a.due > todayStr(); }
function rosterOf(st, a) {
  return (st.students || []).filter(function (s) { return s.cls === a.cls && (!a.subject || (s.subjects || []).indexOf(a.subject) >= 0); }).sort(byNo);
}
function recOfA(st, a, s) { return ((st.records || {})[a.id] || {})[s.id] || {}; }
function statusOfA(st, a, s) { return recOfA(st, a, s).status || 'missing'; }
function label(a) { return a.cls + (a.subject ? ' ' + a.subject : '') + ' ' + (a.tag || '功課') + '「' + a.title + '」（' + a.date + (a.due ? '，收簿日 ' + a.due : '') + '）'; }

// 所有未交（不含未到收簿日的功課）
function outstanding(st) {
  var items = [];
  (st.assignments || []).filter(function (a) { return !isPendingA(a); })
    .sort(function (a, b) { return a.cls.localeCompare(b.cls) || b.date.localeCompare(a.date); })
    .forEach(function (a) {
      var miss = rosterOf(st, a).filter(function (s) { return statusOfA(st, a, s) === 'missing'; });
      if (miss.length) items.push({ a: a, miss: miss });
    });
  return items;
}

function dailySummary() {
  var st = loadState(); if (!st) return;
  var items = outstanding(st), total = 0;
  var html = '<div style="font-family:sans-serif">';
  items.forEach(function (it) {
    total += it.miss.length;
    html += '<h3 style="margin:16px 0 4px">' + label(it.a) + ' — ' + it.miss.length + ' 人未交</h3><ul style="margin:0">';
    it.miss.forEach(function (s) {
      var r = recOfA(st, it.a, s), n = (r.reminds || []).length;
      html += '<li>' + (s.no ? s.no + '. ' : '') + s.name + (n ? ' <span style="color:#888">（已提醒 ' + n + ' 次）</span>' : '') + '</li>';
    });
    html += '</ul>';
  });
  if (!items.length) html += '<p>🎉 今日沒有未交功課。</p>';
  html += '<p style="color:#888;font-size:12px;margin-top:24px">交功課紀錄 App 自動發送 · ' + todayStr() + '</p></div>';
  MailApp.sendEmail({ to: SUMMARY_EMAIL || ownerEmail(), subject: '【交功課】' + todayStr() + ' 未交名單（' + total + ' 人次）', htmlBody: html });
}

function weeklySummary() {
  var st = loadState(); if (!st) return;
  var now = new Date(), day = (now.getDay() + 6) % 7;
  var mon = new Date(now); mon.setDate(now.getDate() - day);
  var fmt = function (d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); };
  var from = fmt(mon), to = todayStr();
  var classes = {};
  (st.students || []).forEach(function (s) { (classes[s.cls] = classes[s.cls] || []).push(s); });
  var html = '<div style="font-family:sans-serif"><p>統計期間：' + from + ' 至 ' + to + '（另附累計）</p>';
  Object.keys(classes).sort().forEach(function (cls) {
    var rows = classes[cls].sort(byNo).map(function (s) {
      var wm = 0, wl = 0, tm = 0, tl = 0;
      (st.assignments || []).forEach(function (a) {
        if (a.cls !== s.cls || isPendingA(a) || (a.subject && (s.subjects || []).indexOf(a.subject) < 0)) return;
        var stt = statusOfA(st, a, s), inWeek = a.date >= from && a.date <= to;
        if (stt === 'missing') { tm++; if (inWeek) wm++; }
        if (stt === 'late') { tl++; if (inWeek) wl++; }
      });
      return { s: s, wm: wm, wl: wl, tm: tm, tl: tl };
    }).filter(function (r) { return r.wm || r.wl || r.tm || r.tl; });
    if (!rows.length) return;
    html += '<h3>' + cls + '</h3><table border="1" cellpadding="4" style="border-collapse:collapse;font-size:13px"><tr style="background:#eee"><th>學號</th><th>姓名</th><th>本週未交</th><th>本週遲交</th><th>累計未交</th><th>累計遲交</th></tr>';
    rows.sort(function (a, b) { return (b.tm - a.tm) || (b.wm - a.wm); }).forEach(function (r) {
      html += '<tr' + (r.tm >= 3 ? ' style="color:#c00"' : '') + '><td>' + (r.s.no || '') + '</td><td>' + r.s.name + '</td><td>' + r.wm + '</td><td>' + r.wl + '</td><td>' + r.tm + '</td><td>' + r.tl + '</td></tr>';
    });
    html += '</table>';
  });
  html += '<p style="color:#888;font-size:12px;margin-top:24px">交功課紀錄 App 自動發送</p></div>';
  MailApp.sendEmail({ to: WEEKLY_EMAILS || ownerEmail(), subject: '【交功課】本週欠交統計 ' + from + ' 至 ' + to, htmlBody: html });
}

function dailyBackup() {
  var st = loadState(); if (!st) return;
  var it = DriveApp.getFoldersByName(BACKUP_FOLDER);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER);
  var name = 'backup_' + todayStr() + '.json';
  var old = folder.getFilesByName(name); while (old.hasNext()) old.next().setTrashed(true);
  folder.createFile(name, JSON.stringify(st), 'application/json');
  // 只保留最近 BACKUP_KEEP 份
  var files = [], fs = folder.getFiles();
  while (fs.hasNext()) { var f = fs.next(); if (/^backup_\d{4}-\d{2}-\d{2}\.json$/.test(f.getName())) files.push(f); }
  files.sort(function (a, b) { return b.getName().localeCompare(a.getName()); });
  files.slice(BACKUP_KEEP).forEach(function (f) { f.setTrashed(true); });
}

// 讀取「學生名單」工作表（老師可在電腦上先在此表輸入整班名單，再從 App 匯入）
function readStudents() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STUDENTS);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var head = values[0].map(function (h) { return String(h).trim(); });
  var idx = function (names) {
    for (var i = 0; i < names.length; i++) { var k = head.indexOf(names[i]); if (k >= 0) return k; }
    return -1;
  };
  var iCls = idx(['班別', '班', 'Class']), iNo = idx(['學號', '號', 'No', 'No.']),
      iName = idx(['姓名', '名字', 'Name']), iPhone = idx(['電話', '家長電話', 'WhatsApp', 'Phone']),
      iSubj = idx(['科目', '選修', 'Subject', 'Subjects']), iOthers = idx(['其他電話', '其他聯絡', 'Other']);
  if (iCls < 0) iCls = 0; if (iNo < 0) iNo = 1; if (iName < 0) iName = 2; if (iPhone < 0) iPhone = 3; if (iSubj < 0) iSubj = 4; if (iOthers < 0) iOthers = 5;
  var clean = function (v) { return String(v == null ? '' : v).replace(/\.0$/, '').trim(); };
  return values.slice(1).filter(function (r) { return clean(r[iName]); }).map(function (r) {
    return { cls: clean(r[iCls]), no: clean(r[iNo]), name: clean(r[iName]), phone: clean(r[iPhone]), subjects: clean(r[iSubj]), others: clean(r[iOthers]) };
  });
}
