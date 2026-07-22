/**
 * LIFEX KPI + 広告実績 データ配信スクリプト v2.1
 * ─────────────────────────────────────────────
 * 【設置場所】KPIスプレッドシート → 拡張機能 → Apps Script
 *            既存のコードを全て消して、このファイルの内容に置き換える
 *
 * 【v1からの変更点】
 *   ・従来のKPIデータ配信はそのまま（レスポンス構造も互換）
 *   ・広告スプシ（別ファイル）を閲覧参照して以下を追加配信
 *       - 月別_反響集計（LP1）: 月別×媒体別の登録数・配信費・CPA・CVR
 *       - 予約集計（LP1）    : 加盟店ごとの予約CV数・累計配信費
 *
 * 【v2 → v2.1の変更点】週次スナップショット機能（2026-07-13追加）
 *   ・URL末尾に ?snapshot=1 を付けて呼ぶと、その時点の全データを
 *     KPIスプシ内の「週次スナップショット」シートに1行保存する
 *     （シートは無ければ自動作成。同じ日の二重保存はスキップ。30件超は古い順に削除）
 *   ・URL末尾に ?history=9 を付けると過去のスナップショットを最大9件返す
 *   ・パラメータなしの呼び出しは今まで通り → ダッシュボードへの影響なし
 *   ・水曜の議題まとめエージェントが ?snapshot=1&history=9 で呼ぶ想定
 *
 * 【v2.1のデプロイ手順】※v2の開発版URLを変えずに更新する
 *   1. コードを置き換えて保存（Ctrl+S）
 *   2. 右上「デプロイ」→「デプロイを管理」→ v2（開発版）のデプロイの鉛筆マーク
 *      → バージョン:「新バージョン」→ デプロイ
 *      ※これでURLはそのまま中身だけ更新される（開発版ダッシュボード・
 *        クラウドエージェントのURL変更が不要）
 *      ※v1（本番ダッシュボードが使用中のデプロイ）は絶対に編集しないこと
 */

// 広告実績スプレッドシート（閲覧できればOK・編集権限は不要）
const AD_SPREADSHEET_ID = '1qsCsnB5SZtSA8XUE5OWxptKjX8eRokJkA6D8Bs-QK9k';

function doGet(e) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('KPIデータ');

  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'シート「KPIデータ」が見つかりません' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const rows = sheet.getDataRange().getValues();

  // 1行目はヘッダーなのでスキップ。店舗名が空の行もスキップ
  const stores = rows.slice(1)
    .filter(row => row[0] && String(row[0]).trim() !== '')
    .map(row => ({
      name:  String(row[0]  || ''),
      sub:   String(row[1]  || ''),
      s: {
        2:  Number(row[2]  || 0),
        3:  Number(row[3]  || 0),
        4:  Number(row[4]  || 0),
        5:  Number(row[5]  || 0),
        6:  Number(row[6]  || 0),
        7:  Number(row[7]  || 0),
        8:  Number(row[8]  || 0),
        9:  Number(row[9]  || 0),
        10: Number(row[10] || 0),
        11: Number(row[11] || 0),
        12: Number(row[12] || 0),
        13: Number(row[13] || 0)
      },
      notes: String(row[14] || '')
    }));

  // ── 広告実績（別スプシを閲覧参照）──
  // 読めなくてもKPI配信は止めない：エラー時は ads.error に理由を入れて返す
  let ads;
  try {
    ads = readAdData();
  } catch (err) {
    ads = { error: String(err) };
  }

  const out = {
    stores:    stores,
    ads:       ads,
    updatedAt: new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' })
  };

  // ── 週次スナップショット ──
  // 失敗してもKPI配信は止めない：エラーはレスポンスに載せて返すだけ
  const p = (e && e.parameter) || {};
  if (p.snapshot === '1') {
    try {
      out.snapshotSaved = saveSnapshot_({ stores: stores, ads: ads, updatedAt: out.updatedAt });
    } catch (err) {
      out.snapshotError = String(err);
    }
  }
  if (p.history) {
    try {
      out.history = readSnapshots_(Math.min(Number(p.history) || 9, SNAPSHOT_KEEP));
    } catch (err) {
      out.historyError = String(err);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ══════════ 週次スナップショット ══════════ */

const SNAPSHOT_SHEET = '週次スナップショット';
const SNAPSHOT_KEEP  = 30;     // 保持件数（週1保存で約半年分）
const SNAPSHOT_CHUNK = 45000;  // セルの5万文字制限を避けてJSONを分割する幅

/**
 * 現在の全データを「週次スナップショット」シートに1行追加する。
 * 行の形: A=日付(yyyy-MM-dd) / B=保存時刻 / C以降=JSONを45,000字ずつ分割
 * 同じ日付の行が既にあれば保存せず 'already-saved' を返す（二重実行対策）。
 */
function saveSnapshot_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SNAPSHOT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SNAPSHOT_SHEET);
    sheet.appendRow(['日付', '保存時刻', 'データ(JSON・分割)']);
  }

  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const last = sheet.getLastRow();
  if (last >= 2 && cellDate_(sheet.getRange(last, 1).getValue()) === today) {
    return 'already-saved';
  }

  const json = JSON.stringify(data);
  const row = [today, Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')];
  for (let i = 0; i < json.length; i += SNAPSHOT_CHUNK) {
    row.push(json.substring(i, i + SNAPSHOT_CHUNK));
  }
  sheet.appendRow(row);

  // 保持件数を超えた分は古い順に削除（ヘッダー行は残す）
  const count = sheet.getLastRow() - 1;
  if (count > SNAPSHOT_KEEP) sheet.deleteRows(2, count - SNAPSHOT_KEEP);

  return today;
}

/**
 * 直近n件のスナップショットを [{date, data}, ...] で返す（古い→新しいの順）。
 * ?snapshot=1&history=9 のように同時指定した場合、保存が先に走るので
 * 配列の最後の要素は「今日保存した分」になる。
 */
function readSnapshots_(n) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SNAPSHOT_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const vals = sheet.getDataRange().getValues();
  return vals.slice(1).slice(-n).map(function(row) {
    let data;
    try {
      data = JSON.parse(row.slice(2).map(function(c) {
        return (c === null || c === undefined) ? '' : String(c);
      }).join(''));
    } catch (err) {
      data = { parseError: String(err) };
    }
    return { date: cellDate_(row[0]), data: data };
  });
}

// セルの値が日付型でも文字列でも yyyy-MM-dd に揃える
function cellDate_(v) {
  return (v instanceof Date)
    ? Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd')
    : String(v).trim();
}

/* ══════════ 広告スプシの読み取り ══════════ */

function readAdData() {
  const ss = SpreadsheetApp.openById(AD_SPREADSHEET_ID);
  const monthly = findSheetByName(ss, '月別_反響集計（LP1）');
  const reserve = findSheetByName(ss, '予約集計（LP1）');
  return {
    months:       monthly ? parseMonthly(monthly)      : [],
    reservations: reserve ? parseReservations(reserve) : []
  };
}

// シート名の前後の空白ゆれを無視して探す
function findSheetByName(ss, target) {
  const t = target.replace(/[\s　]/g, '');
  const hit = ss.getSheets().filter(function(sh) {
    return sh.getName().replace(/[\s　]/g, '') === t;
  });
  return hit.length ? hit[0] : null;
}

/**
 * 月別_反響集計（LP1）のパース
 * レイアウト: 「META広告_YYYY年M月」のタイトル行 → ヘッダー行 → 店舗行…（空行まで）
 *             同じ行のL列に「Google広告_YYYY年M月」ブロックが並ぶ
 * META : A加盟店 Bコード C登録 D配信費 E CPA F表示 Gクリック H率 I CVR J LP流入
 * Google: L加盟店 Mコード N登録 O配信費 P CPA Q表示 Rクリック S率 T CVR U残予算
 */
function parseMonthly(sheet) {
  const vals = sheet.getDataRange().getDisplayValues();
  const months = [];

  for (let r = 0; r < vals.length; r++) {
    const m = String(vals[r][0]).match(/^META広告_(\d{4})年(\d{1,2})月/);
    if (!m) continue;

    const block = {
      ym:     m[1] + '-' + ('0' + m[2]).slice(-2),
      label:  m[1] + '年' + m[2] + '月',
      stores: []
    };

    // タイトル行の次がヘッダー、その次から店舗行
    for (let i = r + 2; i < vals.length; i++) {
      const name = String(vals[i][0]).trim();
      if (!name || /^META広告_/.test(name)) break; // 空行 or 次ブロックで終了
      if (name.charAt(0) === '※') continue;        // 注記行はスキップ
      block.stores.push({
        name:   name,
        code:   String(vals[i][1]).trim(),
        meta:   pickAd(vals[i], 2),
        google: pickAd(vals[i], 13)
      });
    }
    months.push(block);
  }
  return months;
}

// base=登録列のインデックス。{登録, 配信費, CPA, CVR} を数値化して返す
function pickAd(row, base) {
  return {
    reg:  num(row[base]),
    cost: num(row[base + 1]),
    cpa:  num(row[base + 2]),
    cvr:  num(row[base + 6])   // "0.21%" → 0.21
  };
}

/**
 * 予約集計（LP1）のパース
 * 「加盟店」「コード」が横に並ぶヘッダーを探し、
 * その下の表（加盟店 / コード / 予約数 / 累計配信費）を読む
 */
function parseReservations(sheet) {
  const vals = sheet.getDataRange().getDisplayValues();

  for (let r = 0; r < vals.length; r++) {
    for (let c = 0; c < vals[r].length - 1; c++) {
      if (String(vals[r][c]).trim() !== '加盟店') continue;
      if (String(vals[r][c + 1]).trim() !== 'コード') continue;

      const list = [];
      for (let i = r + 1; i < vals.length; i++) {
        const name = String(vals[i][c]).trim();
        if (!name) break;
        if (/合計|小計/.test(name)) continue;
        const count = num(vals[i][c + 2]);
        const cost  = num(vals[i][c + 3]);
        // 数値が両方無い行はデータ行ではない（別セクションの見出し等）→ 表の終わり
        if (count === null && cost === null) break;
        list.push({
          name:  name,
          code:  String(vals[i][c + 1]).trim(),
          count: count,
          cost:  cost
        });
      }
      return list;
    }
  }
  return [];
}

// "¥835,900" "1.61%" "3" → 数値、 "#DIV/0!" "" "-" → null
function num(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).replace(/[¥,％%\s　]/g, '');
  if (s === '' || s === '-' || s.indexOf('#') === 0) return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}
