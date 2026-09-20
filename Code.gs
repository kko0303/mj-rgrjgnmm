/**
 * 雀卓スコア — Googleスプレッドシート連携
 *
 * 置き場所: スプレッドシートを開く → 拡張機能 → Apps Script → このコードを貼り付け
 * デプロイ: 右上「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *   ・次のユーザーとして実行: 自分
 *   ・アクセスできるユーザー: 全員
 *   → 発行されたURLを config.js の endpoint に貼り付けます。
 *
 * KEY は好きな合言葉に変えて、config.js の key と同じにしてください。
 */

var KEY = "change-me";

var SH_RESULT = "対局結果";
var SH_RAW    = "_raw";
var SH_TOTAL  = "通算成績";
var SH_STATE  = "_state";

var HEAD = ["日時","対局ID","ルール","プレイヤー","席","順位","持ち点","素点","ウマオカ","ポイント",
            "局数","和了","放銃","リーチ","ツモ","和了点合計","放銃点合計","平均和了点","平均放銃点",
            "親和了","連荘","飛び","所要分"];

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.key !== KEY) return out({ ok: false, error: "key" });
    if (body.action === "append") return out(append_(body.record));
    if (body.action === "list")   return out({ ok: true, records: list_() });
    if (body.action === "getState") return out({ ok: true, state: {
      settings: state_("settings"), game: state_("game") } });
    if (body.action === "putState") { state_(body.key, body.json); return out({ ok: true }); }
    return out({ ok: false, error: "unknown action" });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  if (!e || !e.parameter || e.parameter.key !== KEY) return out({ ok: false, error: "key" });
  return out({ ok: true, records: list_() });
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name, head) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (head) {
      sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function append_(rec) {
  if (!rec || !rec.gid) return { ok: false, error: "no record" };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var raw = sheet_(SH_RAW, ["対局ID", "JSON"]);
    var ids = raw.getLastRow() > 1
      ? raw.getRange(2, 1, raw.getLastRow() - 1, 1).getValues().map(function (r) { return r[0] })
      : [];
    if (ids.indexOf(rec.gid) >= 0) return { ok: true, duplicated: true };

    var sh = sheet_(SH_RESULT, HEAD);
    var when = new Date(rec.end || rec.ts);
    var rows = (rec.rows || []).map(function (x) {
      return [
        Utilities.formatDate(when, "Asia/Tokyo", "yyyy/MM/dd HH:mm"),
        rec.gid, rec.rule || "", x.name, ["起家","南","西","北"][x.seat], x.rank,
        x.pts, x.score, x.uma, x.total,
        x.hands, x.agari, x.houju, x.riichi, x.tsumo,
        x.agariPts, x.houjuPts,
        x.agari ? Math.round(x.agariPts / x.agari) : "",
        x.houju ? Math.round(x.houjuPts / x.houju) : "",
        x.oyaAgari, x.renchan, x.tobi ? "○" : "", x.mins
      ];
    });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEAD.length).setValues(rows);
    raw.appendRow([rec.gid, JSON.stringify(rec)]);
    total_();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function list_() {
  var raw = sheet_(SH_RAW, ["対局ID", "JSON"]);
  if (raw.getLastRow() < 2) return [];
  var vals = raw.getRange(2, 2, raw.getLastRow() - 1, 1).getValues();
  var outArr = [];
  for (var i = Math.max(0, vals.length - 300); i < vals.length; i++) {
    try { outArr.push(JSON.parse(vals[i][0])) } catch (e) {}
  }
  return outArr;
}

/** 端末どうしで共有する設定・進行中の対局（keyごとに1行） */
function state_(key, json) {
  var sh = sheet_(SH_STATE, ["key", "json", "updatedAt"]);
  var last = sh.getLastRow();
  var keys = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  var row = -1;
  for (var i = 0; i < keys.length; i++) if (keys[i][0] === key) { row = i + 2; break }
  if (json === undefined) {
    if (row < 0) return null;
    try { return JSON.parse(sh.getRange(row, 2).getValue()) } catch (e) { return null }
  }
  var v = [key, JSON.stringify(json), new Date()];
  if (row < 0) sh.appendRow(v); else sh.getRange(row, 1, 1, 3).setValues([v]);
  return true;
}

/** 通算成績シートを作り直す（スプレッドシートを直接見る人向け） */
function total_() {
  var sh = sheet_(SH_RESULT, HEAD);
  if (sh.getLastRow() < 2) return;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getValues();
  var per = {};
  v.forEach(function (r) {
    var name = r[3]; if (!name) return;
    var o = per[name] || (per[name] = { n:0, pt:0, rk:0, r:[0,0,0,0], hands:0, ag:0, hj:0, ri:0, tm:0, agP:0, hjP:0, tobi:0 });
    o.n++; o.pt += Number(r[9]) || 0; o.rk += Number(r[5]) || 0;
    var rank = Number(r[5]); if (rank >= 1 && rank <= 4) o.r[rank - 1]++;
    o.hands += Number(r[10]) || 0; o.ag += Number(r[11]) || 0; o.hj += Number(r[12]) || 0;
    o.ri += Number(r[13]) || 0; o.tm += Number(r[14]) || 0;
    o.agP += Number(r[15]) || 0; o.hjP += Number(r[16]) || 0;
    if (r[21] === "○") o.tobi++;
  });
  var head = ["プレイヤー","半荘数","合計pt","平均pt","平均順位","1着","2着","3着","4着",
              "和了率","放銃率","リーチ率","ツモ率","平均和了点","平均放銃点","飛び"];
  var rows = Object.keys(per).map(function (k) {
    var o = per[k];
    var pc = function (x) { return o.hands ? Math.round(x / o.hands * 1000) / 10 + "%" : "" };
    return [k, o.n, Math.round(o.pt * 10) / 10, Math.round(o.pt / o.n * 10) / 10,
            Math.round(o.rk / o.n * 100) / 100, o.r[0], o.r[1], o.r[2], o.r[3],
            pc(o.ag), pc(o.hj), pc(o.ri),
            o.ag ? Math.round(o.tm / o.ag * 1000) / 10 + "%" : "",
            o.ag ? Math.round(o.agP / o.ag) : "", o.hj ? Math.round(o.hjP / o.hj) : "", o.tobi];
  }).sort(function (a, b) { return b[2] - a[2] });

  var t = sheet_(SH_TOTAL, head);
  t.clear();
  t.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");
  t.setFrozenRows(1);
  if (rows.length) t.getRange(2, 1, rows.length, head.length).setValues(rows);
}
