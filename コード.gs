/*****************************************************************
 * 自習室 業務報告・承認システム  ―  GAS バックエンド (V3)
 *****************************************************************/

const PHOTO_FOLDER_ID = '1Cx55qafluiNMSKZLZnrIDSLM5qpKKw3s';
const SHEET_ID = '1hnZpUIDD5onANEgHoBY2Lm7zC1BpNHiX1tsg__XVbdM'; 
const MANUAL_DOC_ID = '1kRsrEEVzjQqzIYaEUJE34w4zal_CzPHINWbdofGG9sk';
const SHIFT_START = { '朝': '09:00', '夕': '17:30' };
const ENV_RANGE = { tempMin: 23, tempMax: 27, humidMin: 60, humidMax: 80 };

function prop_(key) { return PropertiesService.getScriptProperties().getProperty(key) || ''; }

const SH_STAFF = 'スタッフ';
const SH_TASK = '業務マスタ';
const SH_REPORT = '日報';
const SH_RESULT = '業務実績';
const SH_ATTEND = '勤怠';
const SH_NOTICE = 'お知らせ'; 
const SH_QUEUE = '質問キュー';

function book_() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) { return book_().getSheetByName(name); }
function tz_() { return book_().getSpreadsheetTimeZone() || 'Asia/Tokyo'; }

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'form';
  const webappUrl = prop_('WEBAPP_URL');

  if (page === 'review') {
    const t = HtmlService.createTemplateFromFile('review');
    t.reportId = (e.parameter.id || '');
    t.webappUrl = webappUrl;
    return t.evaluate().setTitle('業務報告の承認').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'manual') {
    const t = HtmlService.createTemplateFromFile('manual');
    t.docId = MANUAL_DOC_ID; t.docUrl = 'https://docs.google.com/document/d/' + MANUAL_DOC_ID + '/edit'; t.webappUrl = webappUrl;
    return t.evaluate().setTitle('自習室マニュアル').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'dashboard') {
    const t = HtmlService.createTemplateFromFile('dashboard'); t.webappUrl = webappUrl;
    return t.evaluate().setTitle('管理ダッシュボード').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'visits') {
    const t = HtmlService.createTemplateFromFile('生徒来室状況'); t.webappUrl = webappUrl;
    return t.evaluate().setTitle('生徒来室状況').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'notice') {
    const t = HtmlService.createTemplateFromFile('notice'); t.webappUrl = webappUrl;
    return t.evaluate().setTitle('お知らせ・引き継ぎ').addMetaTag('viewport', 'width=device-width, initial-scale=1');
    
  }
  if (page === 'quiz') {
    const t = HtmlService.createTemplateFromFile('quiz');
    t.webappUrl = webappUrl;
    return t.evaluate().setTitle('小テスト管理').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'help') {
    const t = HtmlService.createTemplateFromFile('help');
    t.webappUrl = webappUrl;
    return t.evaluate().setTitle('質問受付フォーム').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  const t = HtmlService.createTemplateFromFile('index'); t.webappUrl = webappUrl;
  return t.evaluate().setTitle('業務報告').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  
}

function getInitialData() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const staff = findStaffByEmail_(email);
  const hour = Number(Utilities.formatDate(new Date(), tz_(), 'H'));
  const shift = hour < 15 ? '朝' : '夕';
  return { email: email, staff: staff, staffList: getStaffList_(), shift: shift, today: Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd (E)'), tasks: getTasks(shift) };
}

function getStaffList_() {
  const sh = sheet_(SH_STAFF); if (!sh) return [];
  const rows = sh.getDataRange().getDisplayValues(); const list = [];
  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][0]).trim(), name = String(rows[i][1]).trim();
    if (name) list.push({ staffId: id, name: name });
  }
  return list;
}

function findStaffByEmail_(email) {
  const sh = sheet_(SH_STAFF); if (!sh) return null;
  const rows = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).toLowerCase().trim() === email && email) return { staffId: String(rows[i][0]).trim(), name: String(rows[i][1]).trim() };
  }
  return null;
}

function getTasks(shift) {
  const sh = sheet_(SH_TASK);
  const rows = sh.getDataRange().getDisplayValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[1]).trim() !== shift) continue;
    list.push({
      taskId: String(r[0]).trim(), category: String(r[2]).trim(), name: String(r[3]).trim(),
      photoReq: r[4] === true || String(r[4]).toUpperCase() === 'TRUE',
      envInput: r[5] === true || String(r[5]).toUpperCase() === 'TRUE',
      order: Number(r[6]) || 0,
      manualText: String(r[7] || '').trim(), // マニュアル追加
      type: String(r[8] || '').trim() === '任意' ? 'opt' : 'req' // 任意ボーナス追加
    });
  }
  list.sort(function (a, b) { return a.order - b.order; });
  return list;
}

function addNotice(payload) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = sheet_(SH_NOTICE);
    if (!sh) return { ok: false, error: 'お知らせシートが見つかりません' };
    const now = new Date();
    const noticeId = 'N' + Utilities.formatDate(now, tz_(), 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 100);
    const dateStr = Utilities.formatDate(now, tz_(), 'yyyy/MM/dd HH:mm');
    
    let imageUrl = '';
    if (payload.photo && payload.photo.data) {
      imageUrl = savePhoto_(payload.photo, noticeId, "Notice");
    }
    sh.appendRow([noticeId, dateStr, payload.author || 'システム', payload.title || '無題', payload.content || '', '', imageUrl]);
    
    if (prop_('CHAT_WEBHOOK_URL')) {
      let msg = '📢【全体共有】新しい共有事項\n👤 ' + (payload.author) + '\n📌 ' + (payload.title) + '\n📄 ' + (payload.content);
      if (imageUrl) msg += '\n📷 画像あり';
      postChat_({ text: msg });
    }
    return { ok: true, noticeId: noticeId };
  } catch (err) { return { ok: false, error: String(err) };
  } finally { lock.releaseLock(); }
}

function getNotices(staffId) {
  const sh = sheet_(SH_NOTICE); if (!sh) return [];
  const rows = sh.getDataRange().getDisplayValues(); const list = [];
  for (let i = rows.length - 1; i >= 1; i--) {
    const id = rows[i][0]; if (!id) continue;
    const readers = (rows[i][5] || '').split(',').map(function(s) { return s.trim(); });
    list.push({
      id: id, date: rows[i][1], author: rows[i][2], title: rows[i][3], content: rows[i][4],
      isRead: staffId ? (readers.indexOf(staffId) >= 0) : false,
      imageUrl: rows[i][6] || ''
    });
  }
  return list;
}

// -----------------------------------------------------
// 外部フロントエンド用 APIエンドポイント (doPost)
// -----------------------------------------------------
function doPost(e) {
  try {
    // CORS回避のために text/plain で送られてくる想定でパース
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const payload = data.payload || {};
    let result = {};

    // アクション名に応じて既存の関数をルーティング
    switch (action) {
      case 'getInitialData': result = getInitialData(); break;
      case 'getTasks': result = getTasks(payload.shift); break;
      case 'getAttendanceStatus': result = getAttendanceStatus(payload.shift, payload.staffId, payload.staffName); break;
      case 'clockIn': result = clockIn(payload); break;
      case 'submitReport': result = submitReport(payload); break;
      case 'getNotices': result = getNotices(payload.staffId); break;
      case 'addNotice': result = addNotice(payload); break;
      case 'markNoticeRead': result = markNoticeRead(payload.noticeId, payload.staffId); break;
      case 'getDashboardData': result = getDashboardData(); break;
      case 'getStudentsToday': result = getStudentsToday(); break;
      case 'getQuizData': result = getQuizData(); break;
      case 'addHelpRequest': result = addHelpRequest(payload); break;
      case 'getActiveQueue': result = getActiveQueue(); break;
      case 'updateQueueStatus': result = updateQueueStatus(payload.reqId, payload.status, payload.staffName); break;
      case 'getStudentList': result = getStudentList(); break;
      default: result = { ok: false, error: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function markNoticeRead(noticeId, staffId) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = sheet_(SH_NOTICE); const rows = sh.getDataRange().getDisplayValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === noticeId) {
        let readers = (rows[i][5] || '').split(',').map(function(s){ return s.trim(); }).filter(String);
        if (readers.indexOf(staffId) < 0) {
          readers.push(staffId); sh.getRange(i + 1, 6).setValue(readers.join(', '));
        }
        break;
      }
    }
    return { ok: true };
  } catch(e) { return { ok: false, error: e.toString() }; } finally { lock.releaseLock(); }
}

function submitReport(payload) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const now = new Date();
    const reportId = 'R' + Utilities.formatDate(now, tz_(), 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 1000);
    const dateStr = Utilities.formatDate(now, tz_(), 'yyyy/MM/dd');
    const timeStr = Utilities.formatDate(now, tz_(), 'HH:mm');

    const tasks = payload.tasks || [];
    let doneCount = 0; let totalReq = 0; let bonusPoints = 0;
    const resultSheet = sheet_(SH_RESULT);
    const photoLinks = [];

    tasks.forEach(function (t) {
      if (t.type === 'req') totalReq++;
      if (t.done && t.type === 'req') doneCount++;
      if (t.done && t.type === 'opt') bonusPoints++; // ボーナス加算

      let url = '';
      if (t.photo && t.photo.data) {
        url = savePhoto_(t.photo, reportId, t.name);
        if (url) photoLinks.push({ name: t.name, url: url });
      }
      resultSheet.appendRow([reportId, dateStr, payload.shift, t.name, t.category || '', t.done ? '完了' : '未完了', url, t.memo || '']);
    });

    sheet_(SH_REPORT).appendRow([
      reportId, dateStr, payload.shift, payload.staffId || '', payload.staffName || '',
      timeStr, doneCount, totalReq, payload.temperature || '', payload.humidity || '',
      payload.studentStatus || '', payload.notes || '', '承認待ち', '', '', '', bonusPoints
    ]);

    notifyNewReport_({
      reportId: reportId, date: dateStr, shift: payload.shift, staffName: payload.staffName, staffId: payload.staffId,
      time: timeStr, doneCount: doneCount, total: totalReq, bonusPoints: bonusPoints,
      temperature: payload.temperature, humidity: payload.humidity,
      studentStatus: payload.studentStatus, notes: payload.notes, tasks: tasks, photoLinks: photoLinks
    });
    return { ok: true, reportId: reportId };
  } catch (err) { return { ok: false, error: String(err) }; } finally { lock.releaseLock(); }
}

function savePhoto_(photo, reportId, taskName) {
  try {
    const folder = PHOTO_FOLDER_ID ? DriveApp.getFolderById(PHOTO_FOLDER_ID) : DriveApp.getRootFolder();
    const bytes = Utilities.base64Decode(photo.data);
    const safeName = String(taskName).replace(/[\\/:*?"<>|]/g, '_');
    const blob = Utilities.newBlob(bytes, photo.mime || 'image/jpeg', reportId + '_' + safeName + '.jpg');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) { return ''; }
}

function notifyNewReport_(r) {
  const webappUrl = prop_('WEBAPP_URL');
  const reviewUrl = webappUrl ? webappUrl + '?page=review&id=' + encodeURIComponent(r.reportId) : '';
  const envBad = (r.temperature || r.humidity) && !envStatus_(r.temperature, r.humidity).ok;

  const taskHtml = (r.tasks || []).map(function (t) {
    const mark = t.done ? '✅' : (t.type==='opt' ? '➖' : '⛔️ 未完了');
    const optBadge = t.type==='opt' ? ' [任意]' : '';
    const cam = t.photo && t.photo.data ? ' 📷' : (t.photoReq ? '（写真なし）' : '');
    return '<li>' + mark + ' ' + escHtml_(t.name) + optBadge + cam + (t.memo ? '（' + escHtml_(t.memo) + '）' : '') + '</li>';
  }).join('');

  let html = '<div style="font-family:sans-serif;max-width:560px;line-height:1.6">'
    + '<h2 style="margin:0 0 6px">業務報告の承認依頼</h2>'
    + '<p style="margin:0 0 4px">提出 ' + r.time + '　完了 ' + r.doneCount + ' / ' + r.total + ' (獲得ボーナス: ' + r.bonusPoints + 'PT)</p>'
    + '<ul style="margin:8px 0;padding-left:20px">' + taskHtml + '</ul></div>';

  sendMail_(approverEmails_(), '【承認依頼】' + r.date + ' ' + r.shift + 'の部 / ' + (r.staffName || ''), html);

  if (prop_('CHAT_WEBHOOK_URL')) {
    const lines = (r.tasks || []).map(function (t) {
      const mark = t.done ? '✅' : (t.type==='opt' ? '➖' : '⛔️');
      return mark + ' ' + t.name + (t.type==='opt' ? '[任意]' : '') + (t.memo ? '（' + t.memo + '）' : '');
    }).join('\n');
    postChat_({
      text: '【業務報告】' + r.date + ' ' + r.shift + 'の部\n👤 ' + (r.staffName || '') + '\n⏰ 提出 ' + r.time + '　✔ ' + r.doneCount + '/' + r.total + ' 完了 (🌟ボーナス: ' + r.bonusPoints + 'PT)\n──────────\n' + lines,
      cardsV2: [{ cardId: r.reportId, card: { header: { title: '業務報告', subtitle: r.staffName }, sections: [{ widgets: [{ buttonList: { buttons: buttons_(reviewUrl, r.photoLinks) } }] }] } }]
    });
  }
}

function buttons_(reviewUrl, photoLinks) {
  const btns = [];
  if (reviewUrl) btns.push({ text: '承認画面を開く', onClick: { openLink: { url: reviewUrl } } });
  (photoLinks || []).slice(0, 4).forEach(function (p) { btns.push({ text: '📷 ' + p.name, onClick: { openLink: { url: p.url } } }); });
  return btns;
}

function postChat_(payload) {
  const webhook = prop_('CHAT_WEBHOOK_URL'); if (!webhook) return null;
  return UrlFetchApp.fetch(webhook, { method: 'post', contentType: 'application/json; charset=UTF-8', payload: JSON.stringify(payload), muteHttpExceptions: true });
}

function sendMail_(recipients, subject, htmlBody) {
  if (!recipients) return false;
  try { MailApp.sendEmail({ to: recipients, subject: subject, htmlBody: htmlBody, name: '自習室 業務報告システム' }); return true; } catch (e) { return false; }
}

function approverEmails_() { return prop_('APPROVER_EMAILS') || Session.getEffectiveUser().getEmail() || ''; }

function getStaffEmail_(staffId, staffName) {
  const sh = sheet_(SH_STAFF); if (!sh) return ''; const rows = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(staffId).trim() || (staffName && String(rows[i][1]).trim() === String(staffName).trim())) return String(rows[i][2]).trim();
  }
  return '';
}

function btnLink_(url, label, color) { return '<a href="' + url + '" style="display:inline-block;background:' + color + ';color:#fff;text-decoration:none;padding:11px 18px;border-radius:6px;font-weight:bold;margin:4px 8px 4px 0">' + label + '</a>'; }
function escHtml_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

function getReport(reportId) {
  const rep = sheet_(SH_REPORT).getDataRange().getDisplayValues(); let header = null; const want = String(reportId).trim();
  for (let i = 1; i < rep.length; i++) {
    if (String(rep[i][0]).trim() === want) {
      header = {
        reportId: rep[i][0], date: rep[i][1], shift: rep[i][2], staffId: rep[i][3], staffName: rep[i][4], time: rep[i][5],
        doneCount: rep[i][6], total: rep[i][7], temperature: rep[i][8], humidity: rep[i][9], studentStatus: rep[i][10], notes: rep[i][11],
        status: rep[i][12], approver: rep[i][13], approvedAt: rep[i][14], rejectReason: rep[i][15], bonusPoints: rep[i][16] || 0
      }; break;
    }
  }
  if (!header) return { ok: false, error: '報告が見つかりません' };
  const res = sheet_(SH_RESULT).getDataRange().getDisplayValues(); const tasks = [];
  for (let i = 1; i < res.length; i++) {
    if (String(res[i][0]) === reportId) tasks.push({ name: res[i][3], category: res[i][4], done: String(res[i][5]) === '完了', photoUrl: res[i][6], memo: res[i][7] });
  }
  const me = (Session.getActiveUser().getEmail() || '').toLowerCase();
  header.canApprove = isApprover_(me); header.viewer = me; header.tasks = tasks;
  header.attendance = attendanceLookup_(String(header.date), String(header.shift), String(header.staffId), String(header.staffName));
  header.env = envStatus_(header.temperature, header.humidity);
  return { ok: true, report: header };
}

function isApprover_(email) {
  const list = prop_('APPROVER_EMAILS').split(',').map(function (s) { return s.toLowerCase().trim(); }).filter(String);
  if (list.length === 0) return true; return list.indexOf(email) >= 0;
}

function decideReport(reportId, decision, reason) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const me = (Session.getActiveUser().getEmail() || '').toLowerCase();
    if (!isApprover_(me)) return { ok: false, error: '承認権限がありません' };
    const sh = sheet_(SH_REPORT); const rows = sh.getDataRange().getDisplayValues(); let rowIndex = -1;
    for (let i = 1; i < rows.length; i++) { if (String(rows[i][0]).trim() === reportId) { rowIndex = i + 1; break; } }
    if (rowIndex < 0) return { ok: false, error: '報告が見つかりません' };

    const statusText = decision === 'approve' ? '承認済み' : '差し戻し';
    const nowStr = Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd HH:mm');
    sh.getRange(rowIndex, 13).setValue(statusText); sh.getRange(rowIndex, 14).setValue(me);
    sh.getRange(rowIndex, 15).setValue(nowStr); sh.getRange(rowIndex, 16).setValue(decision === 'approve' ? '' : (reason || ''));

    const staffEmail = getStaffEmail_(rows[rowIndex-1][3], rows[rowIndex-1][4]);
    if (staffEmail) sendMail_(staffEmail, '【業務報告】' + (decision === 'approve' ? '承認' : '差し戻し'), '<p>承認者：' + me + '</p>');
    return { ok: true, status: statusText };
  } catch (err) { return { ok: false, error: String(err) }; } finally { lock.releaseLock(); }
}

function setup() {
  const ss = book_();
  ensureSheet_(ss, SH_STAFF, ['スタッフID', '氏名', 'メールアドレス']);
  ensureSheet_(ss, SH_TASK, ['タスクID', '部', 'カテゴリ', '業務名', '写真必須', '温湿度入力', '表示順', 'マニュアル説明', '種別']);
  ensureSheet_(ss, SH_REPORT, ['報告ID', '日付', '部', 'スタッフID', '氏名', '提出時刻', '完了数', '総数', '温度', '湿度', '生徒の状況', '特記事項', '承認状態', '承認者', '承認時刻', '却下理由', 'ボーナスPT']);
  ensureSheet_(ss, SH_RESULT, ['報告ID', '日付', '部', '業務名', 'カテゴリ', '完了', '写真URL', '備考']);
  ensureSheet_(ss, SH_ATTEND, ['日付', '部', 'スタッフID', '氏名', '出勤時刻', '予定時刻', '状態', '遅刻（分）', '遅刻理由', '通知済']);
  ensureSheet_(ss, SH_NOTICE, ['お知らせID', '日時', '投稿者', 'タイトル', '内容', '既読スタッフID', '画像URL']);
  ensureSheet_(ss, SH_QUEUE, ['リクエストID', '日時', '生徒ID', '生徒名', '質問科目・内容', 'ステータス', '対応スタッフ']); // ←これを追加
}

function ensureSheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}

function nowHHMM_() { return Utilities.formatDate(new Date(), tz_(), 'HH:mm'); }
function hhmmToMin_(s) { const m = String(s || '').match(/^(\d{1,2}):(\d{2})$/); return m ? Number(m[1]) * 60 + Number(m[2]) : null; }
function todayStr_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd'); }

function envStatus_(temp, humid) {
  const t = (temp === '' || temp == null) ? null : Number(temp); const h = (humid === '' || humid == null) ? null : Number(humid);
  const tempOk = (t == null) || (t >= ENV_RANGE.tempMin && t <= ENV_RANGE.tempMax);
  const humidOk = (h == null) || (h >= ENV_RANGE.humidMin && h <= ENV_RANGE.humidMax);
  return { ok: tempOk && humidOk, tempOk: tempOk, humidOk: humidOk, msg: '' };
}

function clockIn(payload) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const shift = payload.shift, staffId = payload.staffId || '', staffName = payload.staffName || '', reason = String(payload.reason||'').trim();
    if (!staffName) return { ok: false, error: '担当者が選択されていません' };
    const sh = sheet_(SH_ATTEND), dateStr = todayStr_(), timeStr = nowHHMM_(), sched = SHIFT_START[shift] || '09:00';
    const diff = hhmmToMin_(timeStr) - hhmmToMin_(sched); const late = diff > Number(prop_('LATE_GRACE_MIN') || '1');
    const rows = sh.getDataRange().getDisplayValues(); let rowIndex = -1;
    for (let i = 1; i < rows.length; i++) { if (String(rows[i][0]) === dateStr && String(rows[i][1]) === shift && String(rows[i][2]) === staffId) { rowIndex = i + 1; break; } }
    let firstTime = timeStr, state = late ? '遅刻' : '定刻', lateMin = late ? diff : '', notified = '';
    if (rowIndex < 0) { sh.appendRow([dateStr, shift, staffId, staffName, timeStr, sched, state, lateMin, reason, '']); rowIndex = sh.getLastRow(); }
    else { firstTime = rows[rowIndex - 1][4]; state = rows[rowIndex - 1][6]; lateMin = rows[rowIndex - 1][7]; notified = rows[rowIndex - 1][9]; if (reason) sh.getRange(rowIndex, 9).setValue(reason); }
    const isLate = (state === '遅刻');
    if (!isLate && notified !== '済') { sh.getRange(rowIndex, 10).setValue('済'); notified = '済'; }
    if (isLate && reason && notified !== '済') { sh.getRange(rowIndex, 10).setValue('済'); }
    return { ok: true, late: isLate, minutesLate: lateMin === '' ? 0 : Number(lateMin), scheduled: sched, actual: firstTime, needReason: isLate && !reason && notified !== '済' };
  } catch (err) { return { ok: false, error: String(err) }; } finally { lock.releaseLock(); }
}

function getAttendanceStatus(shift, staffId, staffName) {
  const sh = sheet_(SH_ATTEND); if (!sh) return null; const dateStr = todayStr_(); const rows = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === dateStr && String(rows[i][1]) === shift && String(rows[i][2]) === String(staffId)) {
      return { actual: rows[i][4], scheduled: rows[i][5], state: rows[i][6], minutesLate: rows[i][7] === '' ? 0 : Number(rows[i][7]), reason: rows[i][8] };
    }
  } return null;
}

function getDashboardData() {
  const me = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!isApprover_(me)) return { ok: false, error: 'このページは代表（承認者）専用です。' };
  const dateStr = todayStr_(), rep = sheet_(SH_REPORT).getDataRange().getDisplayValues(), reportsToday = [], pending = [];
  for (let i = 1; i < rep.length; i++) {
    const item = { reportId: rep[i][0], date: rep[i][1], shift: rep[i][2], staffName: rep[i][4], time: rep[i][5], doneCount: rep[i][6], total: rep[i][7], temperature: rep[i][8], humidity: rep[i][9], status: rep[i][12] };
    if (String(rep[i][1]) === dateStr) reportsToday.push(item);
    if (String(rep[i][12]) === '承認待ち') pending.push(item);
  }
  pending.sort(function (a, b) { return String(a.reportId) < String(b.reportId) ? 1 : -1; });
  const att = sheet_(SH_ATTEND) ? sheet_(SH_ATTEND).getDataRange().getDisplayValues() : [], attendToday = [];
  for (let i = 1; i < att.length; i++) {
    if (String(att[i][0]) === dateStr) attendToday.push({ shift: att[i][1], staffName: att[i][3], actual: att[i][4], scheduled: att[i][5], state: att[i][6], minutesLate: att[i][7] === '' ? 0 : Number(att[i][7]), reason: att[i][8] });
  }
  return { ok: true, viewer: me, today: Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd (E)'), webappUrl: prop_('WEBAPP_URL'), reportsToday: reportsToday, pending: pending.slice(0, 30), attendToday: attendToday };
}

// -----------------------------------------------------
// 生徒来室機能用
// -----------------------------------------------------
const STUDENT_SHEET_ID = '1QcruSLwoyPCQvCuaPK9m5Q3mFK2F2pacZPeu6VHEvps';
function getStudentsToday() {
  try {
    const ss = SpreadsheetApp.openById(STUDENT_SHEET_ID);
    const visitSh = ss.getSheetByName('管理シート'); if (!visitSh) return { ok: false, error: '管理シートが見つかりません' };
    const visit = visitSh.getDataRange().getDisplayValues();
    const strategyMap = {}; const profSh = ss.getSheetByName('個人情報マスタ');
    if (profSh) { const prof = profSh.getDataRange().getDisplayValues(); for (let i = 1; i < prof.length; i++) { if (String(prof[i][0] || '').trim()) strategyMap[String(prof[i][0]).trim()] = String(prof[i][7] || '').trim(); } }
    const today = Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd'); const stillIn = [], leftAlready = [];
    for (let i = 1; i < visit.length; i++) {
      const inTime = String(visit[i][3] || '').trim(); if (!inTime || inTime.indexOf(today) !== 0) continue;
      const id = String(visit[i][1] || '').trim(), outTime = String(visit[i][4] || '').trim();
      const entry = { id: id, name: String(visit[i][2] || '').trim(), inTime: inTime.replace(today, '').trim(), outTime: outTime ? outTime.replace(today, '').trim() : '', duration: String(visit[i][5] || '').trim(), strategy: strategyMap[id] || '' };
      if (outTime) leftAlready.push(entry); else stillIn.push(entry);
    }
    const byIn = function (a, b) { return a.inTime < b.inTime ? -1 : 1; };
    stillIn.sort(byIn); leftAlready.sort(byIn);
    return { ok: true, today: Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd (E)'), now: Utilities.formatDate(new Date(), tz_(), 'HH:mm'), stillIn: stillIn, leftAlready: leftAlready };
  } catch (e) { return { ok: false, error: '取得エラー：' + e }; }
}


// -----------------------------------------------------
// 小テスト管理用（ハイパーリンク完全対応版）
// -----------------------------------------------------
function getQuizData() {
  try {
    const ss = SpreadsheetApp.openById('15C3TN3oMgx8tEWinCUx_TMg_ZzfgRBnhIOEt3tC82aA');
    const sh = ss.getSheetByName('小テスト印刷用リンク');
    if (!sh) return { ok: false, error: '「小テスト印刷用リンク」シートが見つかりません' };
    
    const range = sh.getDataRange();
    const data = range.getValues();
    const richTexts = range.getRichTextValues(); // リッチテキスト（挿入したリンク）を取得
    const formulas = range.getFormulas(); // HYPERLINK関数を取得
    
    if (data.length < 2) return { ok: false, error: 'データがありません' };
    
    // B2〜AE2 (index 1の 1列目〜) に番号が割り振られていると想定
    const numbersRow = data[1] || [];
    const books = [];
    
    // A3 (index 2) 以降のデータをループ処理
    for (let i = 2; i < data.length; i++) {
      const bookName = data[i][0];
      if (!bookName) continue; // 参考書名が空欄の場合はスキップ
      
      const links = [];
      // B列 (index 1) 〜 AE列のリンクをチェック
      for (let j = 1; j < data[i].length; j++) {
        const cellValue = String(data[i][j]).trim();
        const richTextUrl = richTexts[i][j].getLinkUrl();
        const formula = formulas[i][j];
        
        let url = '';
        
        // パターン1: セルに「挿入 > リンク」で設定された場合
        if (richTextUrl && richTextUrl.indexOf('http') === 0) {
          url = richTextUrl;
        } 
        // パターン2: HYPERLINK関数が使われている場合
        else if (formula && formula.match(/hyperlink\("([^"]+)"/i)) {
          url = formula.match(/hyperlink\("([^"]+)"/i)[1];
        } 
        // パターン3: URLがそのままベタ打ちされている場合
        else if (cellValue.indexOf('http') === 0) {
          url = cellValue;
        }
        
        // リンクが見つかった場合のみリストに追加
        if (url) {
          const num = numbersRow[j] || j; // ヘッダーの番号がなければ列番号を使用
          links.push({ number: num, url: url });
        }
      }
      
      // リンクが1つ以上ある参考書だけをリストに追加
      if (links.length > 0) {
        books.push({ name: bookName, links: links });
      }
    }
    
    return { ok: true, books: books };
  } catch (e) {
    return { ok: false, error: '外部シートへのアクセス権限がないか、エラーが発生しました: ' + String(e) };
  }
}


// -----------------------------------------------------
// 質問待ちキュー（Help Request System）関連
// -----------------------------------------------------
function addHelpRequest(payload) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = sheet_(SH_QUEUE);
    const now = new Date();
    const reqId = 'Q' + Utilities.formatDate(now, tz_(), 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 100);
    const dateStr = Utilities.formatDate(now, tz_(), 'HH:mm');
    
    sh.appendRow([
      reqId, dateStr, 
      payload.studentId || '', 
      payload.studentName || '名無し', 
      payload.topic || '', 
      '待ち', ''
    ]);
    
    // Chat連携があれば通知
    if (prop_('CHAT_WEBHOOK_URL')) {
      postChat_({ text: '🚨【質問リクエスト】生徒から質問待ちが入りました！\n👤 ' + (payload.studentName) + '\n📝 内容: ' + (payload.topic) });
    }
    return { ok: true };
  } catch (e) {
    return { ok: false, error: String(e) };
  } finally {
    lock.releaseLock();
  }
}

function getActiveQueue() {
  const sh = sheet_(SH_QUEUE); if (!sh) return [];
  const rows = sh.getDataRange().getDisplayValues();
  const queue = [];
  // 古いものから順に表示したいので上からループ
  for (let i = 1; i < rows.length; i++) {
    const status = rows[i][5];
    if (status === '待ち' || status === '対応中') {
      queue.push({
        id: rows[i][0], time: rows[i][1], studentId: rows[i][2],
        studentName: rows[i][3], topic: rows[i][4], status: status, staff: rows[i][6]
      });
    }
  }
  return queue;
}

function updateQueueStatus(reqId, newStatus, staffName) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = sheet_(SH_QUEUE);
    const rows = sh.getDataRange().getDisplayValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === reqId) {
        sh.getRange(i + 1, 6).setValue(newStatus);
        if (staffName) sh.getRange(i + 1, 7).setValue(staffName);
        break;
      }
    }
    return { ok: true };
  } catch (e) {
    return { ok: false, error: String(e) };
  } finally {
    lock.releaseLock();
  }
}

// -----------------------------------------------------
// 生徒リスト取得（質問受付フォーム用）
// -----------------------------------------------------
function getStudentList() {
  try {
    const ss = SpreadsheetApp.openById('1QcruSLwoyPCQvCuaPK9m5Q3mFK2F2pacZPeu6VHEvps');
    let sh = null;
    
    // gid=1030179692 のシートを特定
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === 1030179692) {
        sh = sheets[i];
        break;
      }
    }
    
    if (!sh) return { ok: false, error: '指定された生徒名簿シートが見つかりません' };
    
    const data = sh.getDataRange().getDisplayValues();
    const students = [];
    
    // 1行目はヘッダーとみなし、2行目からループ処理
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][0]).trim();   // A列: 学習者ID
      const name = String(data[i][3]).trim(); // B列: 氏名 (※氏名がC列の場合は data[i][2] に変更してください)
      
      if (id) {
        students.push({ id: id, name: name || '名前未登録' });
      }
    }
    
    return { ok: true, students: students };
  } catch (e) {
    return { ok: false, error: '生徒リストの取得に失敗しました: ' + String(e) };
  }
}