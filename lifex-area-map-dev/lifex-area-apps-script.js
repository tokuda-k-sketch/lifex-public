/**
 * LIFEX 加盟エリアマップ用 GAS v2（読み取り＋地図からの書き込み）
 *
 * v1からの変更点: doPost追加（地図の入力モードから「スプシに保存」できる）
 * 更新手順: エディタでコード差し替え → 保存 → デプロイ → デプロイを管理 →
 *           鉛筆アイコン → バージョン「新バージョン」→ デプロイ（URLは変わらない）
 *
 * シート「施工範囲マスタ」の列（1行目はヘッダー）:
 *   A: 加盟店名 / B: 都道府県 / C: 市区町村 / D: 区分（エリア購入済 or 施工可能）/ E: 備考
 */
const SHEET_NAME = '施工範囲マスタ';
const PIN = 'lifex2026'; // 地図から保存するときの合言葉。変えたら地図入力時のPINも変わる

function doGet(e) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) return json_({ error: 'シート「' + SHEET_NAME + '」が見つかりません', rows: [] });

  const values = sh.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const member = String(values[i][0] || '').trim();
    const pref = String(values[i][1] || '').trim();
    const muni = String(values[i][2] || '').trim();
    const kind = String(values[i][3] || '').trim();
    if (!member && !pref && !muni) continue; // 空行スキップ
    rows.push({ member: member, pref: pref, muni: muni, kind: kind, row: i + 1 });
  }
  return json_({ updated: new Date().toISOString(), count: rows.length, rows: rows });
}

/**
 * 地図の入力モードからの保存。
 * body: { pin, member, rows: [{pref, muni, kind}...] }
 * 該当加盟店の既存行を全部消して、送られてきた行に差し替える（rows空＝その加盟店を削除）。
 * 他の加盟店の行・備考欄は触らない。
 */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ error: '送信データを読めませんでした' }); }

  if (String(body.pin || '') !== PIN) return json_({ error: 'PINが違います' });
  const member = String(body.member || '').trim();
  if (!member) return json_({ error: '加盟店名がありません' });
  if (!Array.isArray(body.rows)) return json_({ error: 'rowsがありません' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // 同時保存の衝突防止
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sh) return json_({ error: 'シート「' + SHEET_NAME + '」が見つかりません' });

    const values = sh.getDataRange().getValues();
    const header = values.length ? values[0] : ['加盟店名', '都道府県', '市区町村', '区分', '備考'];
    const kept = values.slice(1).filter(function (r) {
      const m = String(r[0] || '').trim();
      return m && m !== member; // この加盟店の行だけ入れ替える
    });
    const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ' 地図から入力';
    const newRows = body.rows.map(function (r) {
      return [member, String(r.pref || '').trim(), String(r.muni || '').trim(), String(r.kind || '').trim(), stamp];
    });

    const all = [header].concat(kept, newRows);
    sh.clearContents();
    sh.getRange(1, 1, all.length, header.length).setValues(all.map(function (r) {
      while (r.length < header.length) r.push('');
      return r.slice(0, header.length);
    }));
    return json_({ ok: true, member: member, saved: newRows.length });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
