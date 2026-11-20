/**
 * 看護師国試 徹底マスター：学習状況を受け取り、管理者のGmailへ自動で知らせる
 *
 * 届くメール（送り先はこのスクリプトを作ったGoogleアカウントのGmail。設定不要）
 *   ・新しい利用者が登録したとき … すぐに「新着」メール
 *   ・毎日20時 … 全員の学習レポート（課題の進捗・平均点・合格予測率）
 *
 * 使い方
 *   1. この中身を script.google.com の新しいプロジェクトに貼り付けて保存
 *   2. 上の関数選択で「はじめに実行」を選んで実行 → 許可（確認メールが届きます）
 *   3. デプロイ → 新しいデプロイ → ウェブアプリ（実行：自分／アクセス：全員）→ URLをアプリの config.js へ
 */

const REPORT_HOUR = 20;          // 毎日のレポート送信時刻（時）
const NOTIFY_NEW_USER = true;    // 新しい利用者の登録をすぐ知らせる
const TZ = "Asia/Tokyo";
const SHEET_NAME = "進捗";
const CAT_NAMES = {body:"人体",dis:"疾病",soc:"社会保障",fund:"基礎看護",adult:"成人",ger:"老年",ped:"小児",mat:"母性",psy:"精神",home:"地域・在宅",int:"統合と実践"};
const LV_NAMES = {1:"基礎",2:"応用",3:"難題"};

/** 最初に1回だけ実行：シート作成・毎日のトリガー設定・確認メール送信 */
function はじめに実行() {
  const email = Session.getEffectiveUser().getEmail();
  PropertiesService.getScriptProperties().setProperty("OWNER_EMAIL", email);
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === "sendDailyReport") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("sendDailyReport").timeBased().everyDays(1).atHour(REPORT_HOUR).inTimezone(TZ).create();
  const sh = sheet_();
  MailApp.sendEmail(email, "【国試マスター】通知の設定が完了しました",
    "このアドレスに、次のメールが自動で届きます。\n\n・新しい利用者が登録したとき（すぐ）\n・毎日" + REPORT_HOUR + "時：全員の学習レポート\n\n" +
    "学習状況の一覧：" + sh.getParent().getUrl() + "\n\n次は「デプロイ → 新しいデプロイ → ウェブアプリ」でURLを発行し、アプリの config.js に貼ってください。");
  console.log("設定完了：" + email + " に確認メールを送りました。");
}

/* ---------- 受信 ---------- */
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_("bad request"); }
  if (body.type === "progress" && body.data && body.data.uid) { saveProgress_(body.data); return out_("ok"); }
  return out_("ignored");
}
function doGet() { return out_("国試マスター 受信サーバー稼働中"); }
function out_(t) { return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.TEXT); }

/* ---------- 保存 ---------- */
function sheet_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss) {
    const id = props.getProperty("SHEET_ID");
    if (id) ss = SpreadsheetApp.openById(id);
    else { ss = SpreadsheetApp.create("国試マスター 学習状況"); props.setProperty("SHEET_ID", ss.getId()); }
  }
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["最終受信", "ID", "名前", "所属", "今日の解答", "平均点", "合格予測率", "マスター", "データ(JSON)"]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function saveProgress_(d) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  let isNew = false;
  try {
    const sh = sheet_();
    const row = [new Date(), d.uid, String(d.name || "").slice(0, 30), String(d.group || "").slice(0, 30),
      d.today || 0, d.avg == null ? "" : d.avg, d.pred ? d.pred.pass : "", (d.mastered || 0) + "/" + (d.totalQ || 0), JSON.stringify(d)];
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
    const i = ids.indexOf(d.uid);
    if (i >= 0) sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
    else { sh.appendRow(row); isNew = true; }
  } finally { lock.releaseLock(); }
  if (isNew && NOTIFY_NEW_USER) {
    mail_("【国試マスター】新しい利用者：" + d.name, "新しい利用者が登録しました。\n\n名前：" + d.name + (d.group ? "\n所属：" + d.group : "") +
      "\n登録日時：" + Utilities.formatDate(new Date(), TZ, "M/d HH:mm") + "\n\n毎日" + REPORT_HOUR + "時のレポートに含まれます。");
  }
}

/* ---------- レポート ---------- */
function loadAll_() {
  const sh = sheet_(); if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().map(function (r) {
    try { return JSON.parse(r[8]); } catch (e) { return null; }
  }).filter(Boolean);
}
function ymd_(d) { return Utilities.formatDate(d, TZ, "yyyy-M-d"); }
function judge_(p) { return p >= 80 ? "◎合格圏" : p >= 60 ? "○合格ライン付近" : p >= 40 ? "△ボーダー" : "✕要強化"; }
function studiedToday_(d) { const t = ymd_(new Date()); return d.lastDay === t && d.today > 0 && ymd_(new Date(d.t)) === t; }

function buildReport_() {
  const list = loadAll_();
  if (!list.length) return { text: "まだ利用者からのデータが届いていません。", html: "<p>まだ利用者からのデータが届いていません。</p>", count: 0, studied: 0 };
  // 合格予測率の低い順（要フォローの人が上）
  list.sort(function (a, b) { return (a.pred ? a.pred.pass : -1) - (b.pred ? b.pred.pass : -1); });
  const esc = function (s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]; }); };
  const texts = [], cards = []; let studied = 0;
  list.forEach(function (d) {
    const st = studiedToday_(d); if (st) studied++;
    const lv = d.lv || [];
    const prog = "ステップ「" + (LV_NAMES[d.level] || "-") + "」・マスター " + (d.mastered || 0) + "/" + (d.totalQ || 0) + "問" +
      (lv.length === 3 ? "（基礎" + lv[0][0] + "/" + lv[0][1] + "・応用" + lv[1][0] + "/" + lv[1][1] + "・難題" + lv[2][0] + "/" + lv[2][1] + "）" : "");
    const prog2 = "過去問の記録 " + (d.krDone || 0) + "/" + (d.krTotal || 0) + "分野・模試 " + (d.mockCount || 0) + "回・要復習 " + (d.wrong || 0) + "問";
    const avg = (d.avg == null ? "—" : d.avg + "点") + "（必修 " + (d.pred ? d.pred.ph + "%" : "—") + "・一般/状況 " + (d.pred ? d.pred.pg + "%" : "—") + "）" +
      (d.mockAvg ? "／模試平均 必修" + d.mockAvg.hp + "%・一般+状況" + d.mockAvg.gp + "%" : "");
    const pred = d.pred ? d.pred.pass + "%　" + judge_(d.pred.pass) : "—（30問以上解くと表示）";
    const today = st ? "今日 " + d.today + "問／直近7日 " + (d.last7 || 0) + "問" : "⚠ 今日は未学習（直近7日 " + (d.last7 || 0) + "問）";
    const weak = d.weak && d.weak.length ? d.weak.map(function (k) { return CAT_NAMES[k] || k; }).join("・") : "";
    const last = Utilities.formatDate(new Date(d.t), TZ, "M/d HH:mm");
    texts.push("■ " + d.name + (d.group ? "（" + d.group + "）" : "") + "\n" + today + "\n【課題の進捗】" + prog + "\n　" + prog2 +
      "\n【平均点】" + avg + "\n【合格予測率】" + pred + (weak ? "\n苦手：" + weak : "") + "\n最終受信：" + last);
    const color = !d.pred ? "#5d6b70" : d.pred.pass >= 80 ? "#15803d" : d.pred.pass >= 60 ? "#0f766e" : d.pred.pass >= 40 ? "#b45309" : "#c2410c";
    cards.push('<div style="border:1px solid #e2e8e6;border-radius:12px;padding:12px 14px;margin:0 0 12px">' +
      '<div style="display:flex;justify-content:space-between;align-items:baseline"><b style="font-size:16px">' + esc(d.name) + '</b><span style="font-size:12px;color:#5d6b70">' + esc(d.group) + '</span></div>' +
      '<div style="font-size:26px;font-weight:900;color:' + color + ';margin:4px 0">合格予測 ' + (d.pred ? d.pred.pass + '%' : '—') + ' <span style="font-size:13px">' + (d.pred ? judge_(d.pred.pass) : '30問以上で表示') + '</span></div>' +
      '<div style="font-size:14px;' + (st ? '' : 'color:#c2410c;font-weight:700') + '">' + esc(today) + '</div>' +
      '<table style="font-size:13px;margin-top:6px;border-collapse:collapse">' +
      '<tr><td style="color:#5d6b70;padding:2px 8px 2px 0;white-space:nowrap;vertical-align:top">課題の進捗</td><td>' + esc(prog) + '<br>' + esc(prog2) + '</td></tr>' +
      '<tr><td style="color:#5d6b70;padding:2px 8px 2px 0;vertical-align:top">平均点</td><td>' + esc(avg) + '</td></tr>' +
      (weak ? '<tr><td style="color:#5d6b70;padding:2px 8px 2px 0">苦手</td><td>' + esc(weak) + '</td></tr>' : '') +
      '<tr><td style="color:#5d6b70;padding:2px 8px 2px 0">最終受信</td><td>' + esc(last) + '</td></tr></table></div>');
  });
  const summary = "利用者 " + list.length + "人／今日学習した人 " + studied + "人";
  return {
    text: summary + "\n━━━━━━━━━━\n" + texts.join("\n\n"),
    html: '<div style="font-family:sans-serif;max-width:640px;color:#1d2a2e"><h2 style="color:#0f766e;margin:0 0 4px">看護師国試 学習レポート</h2>' +
      '<p style="margin:0 0 14px;color:#5d6b70">' + esc(summary) + '（合格予測率の低い順）</p>' + cards.join("") +
      '<p style="font-size:12px;color:#5d6b70">一覧：<a href="' + sheet_().getParent().getUrl() + '">国試マスター 学習状況（スプレッドシート）</a></p></div>',
    count: list.length, studied: studied
  };
}

/* ---------- 送信 ---------- */
function owner_() { return PropertiesService.getScriptProperties().getProperty("OWNER_EMAIL") || Session.getEffectiveUser().getEmail(); }
function mail_(subject, text, html) {
  const opt = { name: "国試マスター" }; if (html) opt.htmlBody = html;
  MailApp.sendEmail(owner_(), subject, text, opt);
}
/** 毎日のレポート（トリガーから自動実行） */
function sendDailyReport() {
  const r = buildReport_();
  const day = Utilities.formatDate(new Date(), TZ, "M/d");
  mail_("【国試マスター】" + day + " 学習レポート（学習 " + r.studied + "/" + r.count + "人）", r.text, r.html);
}
/** 動作確認用：今すぐレポートを送る */
function 今すぐレポート送信() { sendDailyReport(); }
