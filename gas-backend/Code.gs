/**
 * KOFIRICE 金庫管理アプリ — バックエンド (Google Apps Script)
 *
 * スプレッドシートを追記専用の台帳（レジャー）として使い、
 * 複数端末から同じ在庫を見られるようにする。
 *
 * セットアップ手順は SAFE-README.md を参照。
 *
 * 設計方針:
 *  - movements シートは「追記のみ」。修正は打ち消し行(VOID)を足して行う。
 *    → 過去の記録が書き換わらないので、金庫の監査ができる。
 *  - 1行は必ず1つの区画(compartment)だけを増減させる。
 *    例:「売上袋からおつりを補充」は sales(−) と change(+) の2行に分けて記録する。
 *  - 同期は seq(=行番号) ベース。時計のズレに影響されない。
 */

// ===== 設定 =====================================================
// アプリ側の「接続キー」と同じ文字列にすること。推測されにくい値へ必ず変更する。
const TOKEN = 'CHANGE_ME_kofirice_safe_2026';

// 通常は空のままでよい（スプレッドシートに紐づけて使う場合）。
// 別のスプレッドシートを使いたい場合だけ、そのIDを入れる。
const SPREADSHEET_ID = '';

const DENOMS = [200, 100, 50, 20, 10, 5, 1];
const MOV_HEADER = ['seq', 'id', 'ts', 'store', 'compartment', 'type']
  .concat(DENOMS.map(function (d) { return 'n' + d; }))
  .concat(['amount', 'ref', 'actor', 'note', 'source', 'createdAt']);

// ===== エントリポイント ==========================================

function doGet() {
  // 接続テスト用。トークンなしでも動くが、データは一切返さない。
  return json({ ok: true, service: 'kofirice-safe', serverTime: new Date().toISOString() });
}

function doPost(e) {
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.token !== TOKEN) return json({ ok: false, error: 'bad_token' });

    var lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      if (req.action === 'sync') return json(handleSync(req));
      if (req.action === 'ping') return json({ ok: true, serverTime: new Date().toISOString() });
      return json({ ok: false, error: 'unknown_action' });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// ===== 同期本体 ==================================================
/**
 * 1往復で「送信(push)」と「受信(pull)」の両方を行う。
 *  req.push       : 端末側でまだサーバに送っていない movement の配列
 *  req.since      : 端末が持っている最後の seq
 *  req.stores     : 店舗設定を更新したいときだけ入れる
 *  req.storesRev  : 端末が把握している店舗設定のリビジョン
 */
function handleSync(req) {
  var sheet = getMovSheet();
  var existing = getExistingIds(sheet);
  var accepted = [];
  var rows = [];
  var now = new Date().toISOString();

  (req.push || []).forEach(function (m) {
    if (!m || !m.id) return;
    // 二重送信されても1回しか入らないようにする（通信リトライ対策）
    if (existing[m.id]) { accepted.push(m.id); return; }
    existing[m.id] = true;
    accepted.push(m.id);
    rows.push(movementToRow(m, now));
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, MOV_HEADER.length).setValues(rows);
    // seq 列は行番号をそのまま使う（追記専用なので単調増加が保証される）
    var start = sheet.getLastRow() - rows.length + 1;
    var seqs = [];
    for (var i = 0; i < rows.length; i++) seqs.push([start + i]);
    sheet.getRange(start, 1, rows.length, 1).setValues(seqs);
  }

  var storeResult = handleStores(req);

  return {
    ok: true,
    serverTime: now,
    accepted: accepted,
    movements: readMovementsSince(sheet, Number(req.since) || 0),
    stores: storeResult.stores,
    storesRev: storeResult.rev,
    storesConflict: storeResult.conflict
  };
}

/**
 * 店舗設定（店舗一覧・おつりボックスの目標在庫）の読み書き。
 * 端末が古いリビジョンを持ったまま上書きしようとした場合は拒否し、
 * サーバ側の内容をそのまま返す（勝手に上書きさせない）。
 */
function handleStores(req) {
  var props = PropertiesService.getScriptProperties();
  var rev = Number(props.getProperty('storesRev') || 0);
  var stores = JSON.parse(props.getProperty('stores') || 'null');

  if (stores === null) {
    stores = defaultStores();
    rev = 1;
    props.setProperty('stores', JSON.stringify(stores));
    props.setProperty('storesRev', String(rev));
  }

  if (req.stores) {
    if (Number(req.storesRev) !== rev) {
      return { stores: stores, rev: rev, conflict: true };
    }
    stores = req.stores;
    rev = rev + 1;
    props.setProperty('stores', JSON.stringify(stores));
    props.setProperty('storesRev', String(rev));
    writeStoresSheet(stores);
  }
  return { stores: stores, rev: rev, conflict: false };
}

function defaultStores() {
  // till-close.html の初期フロート設定に合わせた初期値。設定画面から変更できる。
  return [{
    code: 'S2',
    name: '2号店',
    target: { 200: 0, 100: 0, 50: 10, 20: 20, 10: 30, 5: 60, 1: 150 }
  }];
}

// ===== シート操作 ================================================

function getSS() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function getMovSheet() {
  var ss = getSS();
  var sh = ss.getSheetByName('movements');
  if (!sh) {
    sh = ss.insertSheet('movements');
    sh.getRange(1, 1, 1, MOV_HEADER.length).setValues([MOV_HEADER]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getExistingIds(sheet) {
  var last = sheet.getLastRow();
  var map = {};
  if (last < 2) return map;
  var ids = sheet.getRange(2, 2, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (ids[i][0]) map[String(ids[i][0])] = true;
  return map;
}

function readMovementsSince(sheet, since) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var startRow = Math.max(2, since + 1);
  if (startRow > last) return [];
  var values = sheet.getRange(startRow, 1, last - startRow + 1, MOV_HEADER.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!r[1]) continue;
    var notes = {};
    for (var j = 0; j < DENOMS.length; j++) notes[DENOMS[j]] = Number(r[6 + j]) || 0;
    out.push({
      seq: Number(r[0]) || (startRow + i),
      id: String(r[1]),
      ts: toIso(r[2]),
      store: String(r[3]),
      compartment: String(r[4]),
      type: String(r[5]),
      notes: notes,
      amount: Number(r[6 + DENOMS.length]) || 0,
      ref: r[7 + DENOMS.length] ? String(r[7 + DENOMS.length]) : '',
      actor: r[8 + DENOMS.length] ? String(r[8 + DENOMS.length]) : '',
      note: r[9 + DENOMS.length] ? String(r[9 + DENOMS.length]) : '',
      source: r[10 + DENOMS.length] ? String(r[10 + DENOMS.length]) : ''
    });
  }
  return out;
}

function movementToRow(m, now) {
  var notes = m.notes || {};
  var row = [0, String(m.id), toIso(m.ts) || now, String(m.store || ''), String(m.compartment || ''), String(m.type || '')];
  var amount = 0;
  for (var i = 0; i < DENOMS.length; i++) {
    var n = Number(notes[DENOMS[i]]) || 0;
    row.push(n);
    amount += DENOMS[i] * n;
  }
  row.push(amount);
  row.push(String(m.ref || ''));
  row.push(String(m.actor || ''));
  row.push(String(m.note || ''));
  row.push(String(m.source || ''));
  row.push(now);
  return row;
}

/** 人が見て分かるように、店舗設定も別シートに書き出しておく（読み取り専用の参考用）。 */
function writeStoresSheet(stores) {
  var ss = getSS();
  var sh = ss.getSheetByName('stores');
  if (!sh) sh = ss.insertSheet('stores');
  sh.clear();
  var header = ['code', 'name'].concat(DENOMS.map(function (d) { return 'target' + d; }));
  var rows = [header];
  stores.forEach(function (s) {
    var r = [s.code, s.name];
    DENOMS.forEach(function (d) { r.push(Number((s.target || {})[d]) || 0); });
    rows.push(r);
  });
  sh.getRange(1, 1, rows.length, header.length).setValues(rows);
  sh.setFrozenRows(1);
}

function toIso(v) {
  if (!v) return '';
  if (v instanceof Date) return v.toISOString();
  return String(v);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 手動実行用ユーティリティ ==================================

/** スプレッドシートを初期化する。デプロイ前に一度だけエディタから実行する。 */
function setup() {
  getMovSheet();
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('stores')) {
    props.setProperty('stores', JSON.stringify(defaultStores()));
    props.setProperty('storesRev', '1');
  }
  writeStoresSheet(JSON.parse(props.getProperty('stores')));
  Logger.log('セットアップ完了。ウェブアプリとしてデプロイしてください。');
}

/** 店舗設定をやり直したいときに実行する（movements の記録は消えない）。 */
function resetStores() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('stores');
  props.deleteProperty('storesRev');
  Logger.log('店舗設定をリセットしました。');
}
