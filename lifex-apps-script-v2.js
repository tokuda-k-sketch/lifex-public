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
 *
 * 【v2.1 → v2.2の変更点】商談ふりかえりログ機能（2026-08-09追加）
 *   ・doGet(?furikaeri=1[&s=TOKEN]) で「ふりかえりログ」シートを返す
 *     （既存のKPI配信ロジックより前で早期returnする完全に独立したルート）
 *   ・doPost() を新規追加。商談ふりかえりシートの「提出する」「コメント保存」
 *     から呼ばれる（action: furikaeri_submit / furikaeri_comment）
 *   ・トークンは既存の「アクセス管理」シート（KPIダッシュボードと共通）を流用
 *   ・「ふりかえりログ」シートは初回アクセス時に自動作成。既存シートには一切影響なし
 *
 * 【v2.2 → v2.3の変更点】MTGアーカイブ視聴ログ機能（2026-09-14追加）
 *   ・doGet(?archiveLog=1&s=TOKEN&session=回&video=動画ID) で視聴クリックを1件記録
 *   ・doGet(?archiveData=1[&s=TOKEN]) で視聴ログ一覧を返す（本部の可視化ページ用）
 *   ・トークンは既存の「アクセス管理」シートを流用。「視聴ログ」シートは初回アクセス時に自動作成
 */

// 広告実績スプレッドシート（閲覧できればOK・編集権限は不要）
const AD_SPREADSHEET_ID = '1qsCsnB5SZtSA8XUE5OWxptKjX8eRokJkA6D8Bs-QK9k';

function doGet(e) {
  const p0 = (e && e.parameter) || {};

  // ── ふりかえりログの閲覧（既存KPI配信とは完全に別ルート。ここで早期return）──
  if (p0.furikaeri === '1') {
    return handleFurikaeriGet_(p0);
  }

  // ── MTGアーカイブ視聴ログの記録・閲覧（既存KPI配信とは完全に別ルート。ここで早期return）──
  if (p0.archiveLog === '1') {
    return handleArchiveLogRecord_(p0);
  }
  if (p0.archiveData === '1') {
    return handleArchiveLogList_(p0);
  }

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
    .filter(function(row) { return row[0] && String(row[0]).trim() !== ''; })
    .map(function(row) {
      return {
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
      };
    });

  // ── 広告実績（別スプシを閲覧参照）──
  let ads;
  try {
    ads = readAdData();
  } catch (err) {
    ads = { error: String(err) };
  }

  // ── 目標 ──
  let targets;
  try {
    targets = readTargets_(stores);
  } catch (err) {
    targets = [];
  }

  // ── アクセス管理（加盟店別ビュー用） ──
  // 未登録店舗を自動追加。読み書きエラーでも通常配信は止めない
  let accessRows = [];
  try {
    accessRows = ensureAccessControl_(stores);
  } catch (err) {
    accessRows = [];
  }

  const p = (e && e.parameter) || {};
  const storeToken = p.s || null;

  // トークンが付いていれば加盟店ビュー → 自店舗のデータのみ返す
  let displayStores = stores;
  let storeGroup = null;
  var isStoreView = false;

  if (storeToken) {
    const matched = accessRows.filter(function(ac) { return ac.token === storeToken; });
    if (!matched.length) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: 'invalid_token' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    isStoreView = true;
    storeGroup = matched[0].groupName;
    const matchKeys = new Set(
      matched.map(function(m) { return normKey_(m.name) + '|' + normKey_(m.sub); })
    );
    // 自店舗のみ・内部メモ(notes)は除去
    displayStores = stores
      .filter(function(st) {
        return matchKeys.has(normKey_(st.name) + '|' + normKey_(st.sub));
      })
      .map(function(st) {
        return { name: st.name, sub: st.sub, s: st.s };
      });
  }

  // 全体集計は常に全店舗から算出（加盟店ビューのKPIカード・ファネル用）
  const globalAggregate = computeGlobalAggregate_(stores);

  const out = {
    stores:          displayStores,
    globalAggregate: globalAggregate,
    storeGroup:      storeGroup,
    isStoreView:     isStoreView,
    ads:             isStoreView ? null : ads,
    targets:         targets,
    updatedAt:       new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' })
  };

  // ── 週次スナップショット（管理者ビューのみ保存可） ──
  if (p.snapshot === '1' && !isStoreView) {
    try {
      out.snapshotSaved = saveSnapshot_({ stores: stores, ads: ads, updatedAt: out.updatedAt });
    } catch (err) {
      out.snapshotError = String(err);
    }
  }
  if (p.history) {
    try {
      const raw = readSnapshots_(Math.min(Number(p.history) || 9, SNAPSHOT_KEEP));
      if (isStoreView) {
        // 加盟店ビューでは各スナップショットも自店舗のデータのみに絞り込む
        const filterKeys = new Set(
          accessRows
            .filter(function(ac) { return ac.token === storeToken; })
            .map(function(m) { return normKey_(m.name) + '|' + normKey_(m.sub); })
        );
        out.history = raw.map(function(h) {
          const filtered = (h.data.stores || []).filter(function(st) {
            return filterKeys.has(normKey_(st.name) + '|' + normKey_(st.sub));
          });
          return { date: h.date, data: { stores: filtered } };
        });
      } else {
        out.history = raw;
      }
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

/* ══════════ 目標シートの読み取り ══════════ */

const TARGET_SHEET = '目標';

/**
 * 「目標」シートを読み、各店の自己設定目標を配列で返す。
 * シートが無ければ ensureTargetSheet_ が全店舗名入りで自動作成する。
 * レイアウト（1行目ヘッダー）:
 *   A=店舗名 / B=支店 / C=目標アポ率 / D=目標接続率 / E=目標契約率
 *   ※数値は 25 でも 25% でも可。空欄はその指標の目標なしとして扱う。
 * 返り値: [{ name, sub, kpi2(=アポ率), kpi1(=接続率), kpi3(=契約率) }, ...]
 *   ダッシュボードのKPIキー名(kpi1/kpi2/kpi3)に合わせてある。
 */
function readTargets_(stores) {
  const sheet = ensureTargetSheet_(stores);
  if (!sheet) return [];

  const rows = sheet.getDataRange().getValues();
  return rows.slice(1)
    .filter(function(row) { return row[0] && String(row[0]).trim() !== ''; })
    .map(function(row) {
      return {
        name: String(row[0] || ''),
        sub:  String(row[1] || ''),
        kpi2: numOrNull_(row[2]),  // 目標アポ率
        kpi1: numOrNull_(row[3]),  // 目標接続率
        kpi3: numOrNull_(row[4])   // 目標契約率
      };
    });
}

/**
 * 「目標」シートが無ければ、ヘッダー＋全店舗名を入れて自動作成する。
 * 既にあれば一切さわらない（徳田さんが入力した目標値を絶対に上書きしない）。
 * → デプロイ後、ダッシュボードを一度開けば「目標」シートが全店名入りで出現。
 *   あとはC列（目標アポ率）に数字を入れるだけで、その店の目標ゲージが点灯する。
 */
function ensureTargetSheet_(stores) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = findSheetByName(ss, TARGET_SHEET);
  if (sheet) return sheet;

  sheet = ss.insertSheet(TARGET_SHEET);
  sheet.appendRow(['店舗名', '支店', '目標アポ率(%)', '目標接続率(%)', '目標契約率(%)']);
  (stores || []).forEach(function(st) {
    sheet.appendRow([st.name || '', st.sub || '', '', '', '']);
  });
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  sheet.setColumnWidths(1, 5, 130);
  return sheet;
}

// "25" "25%" "25.0" → 数値、 "" null "-" → null
function numOrNull_(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).replace(/[％%\s　]/g, '');
  if (s === '' || s === '-') return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
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

/* ══════════ アクセス管理（加盟店別ビュー） ══════════
 *
 * 「アクセス管理」シートで加盟店ごとのアクセストークンを管理する。
 *
 * シート構成（1行目ヘッダー）:
 *   A=店舗名 / B=支店 / C=トークン / D=グループ表示名 / E=通知方法（将来用） / F=送信先（将来用）
 *
 * ■ 同じトークン値を複数行に書くと、そのすべての店舗が1つのURLで見られる
 *   （例: 中央建設①〜④に同じトークンを設定 → 1つのURLで4支店分を表示）
 *
 * ■ 新しい加盟店を追加するには「KPIデータ」シートに行を追加するだけでOK。
 *   次回GASが呼ばれたとき「アクセス管理」シートに自動で行が追加され、
 *   ランダムなトークンが発行される。
 *
 * ■ 将来の定期通知（LINE/メール等）は E列・F列を使って実装予定。
 */

const ACCESS_SHEET = 'アクセス管理';

function ensureAccessControl_(stores) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSheetByName(ss, ACCESS_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(ACCESS_SHEET);
    sheet.appendRow(['店舗名', '支店', 'トークン', 'グループ表示名', '通知方法（将来用）', '送信先（将来用）']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    sheet.setColumnWidths(1, 6, 140);
  }

  // 既存データを読む
  const vals = sheet.getDataRange().getValues();
  const existing = vals.slice(1).filter(function(r) { return String(r[0]).trim() !== ''; });

  // KPIデータに存在するが未登録の店舗を自動追加（個別トークンを発行）
  const existingKeys = new Set(
    existing.map(function(r) { return normKey_(String(r[0])) + '|' + normKey_(String(r[1])); })
  );
  const toAdd = (stores || []).filter(function(st) {
    return !existingKeys.has(normKey_(st.name) + '|' + normKey_(st.sub));
  });
  if (toAdd.length) {
    const newRows = toAdd.map(function(st) {
      return [st.name, st.sub, generateToken_(), st.name, '', ''];
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 6).setValues(newRows);
  }

  // 最新データを返す（トークンが空の行はスキップ）
  return sheet.getDataRange().getValues().slice(1)
    .filter(function(r) { return String(r[0]).trim() && String(r[2]).trim(); })
    .map(function(r) {
      return {
        name:         String(r[0]).trim(),
        sub:          String(r[1]).trim(),
        token:        String(r[2]).trim(),
        groupName:    String(r[3]).trim() || String(r[0]).trim(),
        notifyMethod: String(r[4]).trim(),
        notifyTarget: String(r[5]).trim()
      };
    });
}

// 店舗名の正規化（半角全角スペース除去・小文字化）。マッチング用
function normKey_(s) {
  return String(s || '').replace(/[\s　]/g, '').toLowerCase();
}

// 10文字のランダムトークンを生成（UUID短縮）
function generateToken_() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 10);
}

/**
 * 全店舗の合計KPIを計算して返す。
 * 加盟店ビューの「全体集計」欄（KPIカード・ファネル）に使用。
 * HTMLのcalcKPIs/sumStoresと同等のロジック。
 */
function computeGlobalAggregate_(stores) {
  var total = 0, connected = 0, appt = 0, interview = 0, contract = 0;
  stores.forEach(function(st) {
    var g = function(k) { return Number(st.s[k] || 0); };
    total     += [2,3,4,5,6,7,8,9,10,11,12,13].reduce(function(a,k){ return a+g(k); }, 0);
    connected += [4,5,6,7,8,9,10,11].reduce(function(a,k){ return a+g(k); }, 0);
    appt      += [5,6,7,8,9,10,11].reduce(function(a,k){ return a+g(k); }, 0);
    interview += [6,7,8,9,10,11].reduce(function(a,k){ return a+g(k); }, 0);
    contract  += [10,11].reduce(function(a,k){ return a+g(k); }, 0);
  });
  var pct = function(n, d) { return d > 0 ? (n / d * 100).toFixed(1) : null; };
  return {
    total: total, connected: connected, appt: appt, interview: interview, contract: contract,
    kpi1:    pct(connected, total),
    kpi2:    pct(appt, connected),
    kpi3:    pct(contract, interview),
    oa:      pct(appt, total),
    oi:      pct(interview, total),
    oc:      pct(contract, total),
    apptToIv: appt > 0 ? (interview / appt * 100).toFixed(1) : null
  };
}

/* ══════════ 商談ふりかえりログ（2026-08-09追加） ══════════
 *
 * 商談ふりかえりシート（shodan-furikaeri.html）の「提出する」ボタンから
 * doPostで送られた内容を「ふりかえりログ」シートに1行ずつ蓄積する。
 * 閲覧はdoGet(?furikaeri=1)で行う。KPIダッシュボードと同じ「アクセス管理」
 * シートのトークンを流用し、店舗トークン付き=自店の提出のみ、
 * トークンなし=全店舗横断（本部・徳田さん用）で返す。
 *
 * シート「ふりかえりログ」（無ければ初回アクセス時に自動作成）
 *   A=ID B=提出日時 C=店舗名 D=支店 E=商談日 F=対象のお客様 G=結果
 *   H=人_つかめたこと I=人_次に聞くこと J=物_つかめたこと K=物_次に聞くこと
 *   L=金_つかめたこと M=金_次に聞くこと N=時_つかめたこと O=時_次に聞くこと
 *   P=敵_つかめたこと Q=敵_次に聞くこと R=分かれ目の場面 S=本音チェック
 *   T=次に変える1つ U=いつ誰で試すか V=店長コメント
 */

const FURIKAERI_SHEET = 'ふりかえりログ';
const FURIKAERI_HEADERS = [
  'ID', '提出日時', '店舗名', '支店', '商談日', '対象のお客様', '結果',
  '人_つかめたこと', '人_次に聞くこと', '物_つかめたこと', '物_次に聞くこと',
  '金_つかめたこと', '金_次に聞くこと', '時_つかめたこと', '時_次に聞くこと',
  '敵_つかめたこと', '敵_次に聞くこと', '分かれ目の場面', '本音チェック',
  '次に変える1つ', 'いつ誰で試すか', '店長コメント',
  // ── AI採点（2026-08-18追加）──
  '録音ファイルURL', '文字起こしテキスト', 'AI採点ステータス', 'AI採点結果JSON', 'AI採点日時'
];

function ensureFurikaeriSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = findSheetByName(ss, FURIKAERI_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(FURIKAERI_SHEET);
    sheet.appendRow(FURIKAERI_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, FURIKAERI_HEADERS.length).setFontWeight('bold');
    sheet.setColumnWidths(1, FURIKAERI_HEADERS.length, 160);
    return sheet;
  }
  // 既存シート（本番稼働中）に新しい列が定義された場合、既存データ列は一切動かさず
  // 末尾に不足分だけ追加する（AI採点機能の追加時など、後方互換のためのマイグレーション）
  const existingHeaderCount = sheet.getLastColumn();
  if (existingHeaderCount < FURIKAERI_HEADERS.length) {
    const missing = FURIKAERI_HEADERS.slice(existingHeaderCount);
    sheet.getRange(1, existingHeaderCount + 1, 1, missing.length).setValues([missing]);
    sheet.getRange(1, 1, 1, FURIKAERI_HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

// 「アクセス管理」シートをそのまま読むだけ（KPIデータ側の自動登録は行わない。
// ふりかえりのトークン検証は、既にKPIダッシュボードで発行済みのトークンに乗る想定）
function readAccessControlRaw_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = findSheetByName(ss, ACCESS_SHEET);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter(function(r) { return String(r[0]).trim() && String(r[2]).trim(); })
    .map(function(r) {
      return { name: String(r[0]).trim(), sub: String(r[1]).trim(), token: String(r[2]).trim(), group: String(r[3] || '').trim() };
    });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function handleFurikaeriGet_(p) {
  const sheet = ensureFurikaeriSheet_();
  const vals = sheet.getDataRange().getValues();
  const headers = vals[0];
  const dataRows = vals.slice(1).filter(function(r) { return String(r[0]).trim() !== ''; });

  const storeToken = p.s || null;
  let allowedKeys = null;
  if (storeToken) {
    const accessRows = readAccessControlRaw_();
    const matched = accessRows.filter(function(ac) { return ac.token === storeToken; });
    if (!matched.length) return jsonOut_({ error: 'invalid_token' });
    allowedKeys = new Set(
      matched.map(function(m) { return normKey_(m.name) + '|' + normKey_(m.sub); })
    );
  }

  let entries = dataRows.map(function(row) {
    const obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });

  if (allowedKeys) {
    entries = entries.filter(function(en) {
      return allowedKeys.has(normKey_(en['店舗名']) + '|' + normKey_(en['支店']));
    });
  }

  entries.reverse(); // 新しい提出が先頭に来るように

  return jsonOut_({ entries: entries, isStoreView: !!storeToken });
}

function doPost(e) {
  let p;
  try {
    p = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ error: 'bad_request' });
  }

  if (p.action === 'furikaeri_submit') return handleFurikaeriSubmit_(p);
  if (p.action === 'furikaeri_comment') return handleFurikaeriComment_(p);
  return jsonOut_({ error: 'unknown_action' });
}

function handleFurikaeriSubmit_(p) {
  const sheet = ensureFurikaeriSheet_();
  const id = Utilities.getUuid();
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  const recordingUrl = p.recordingUrl || '';
  const transcriptText = p.transcriptText || '';
  sheet.appendRow([
    id, now,
    p.storeName || '', p.storeSub || '', p.dealDate || '', p.customerName || '', p.result || '',
    p.jin1 || '', p.jin2 || '', p.mono1 || '', p.mono2 || '',
    p.kane1 || '', p.kane2 || '', p.toki1 || '', p.toki2 || '',
    p.teki1 || '', p.teki2 || '',
    p.turningPoint || '', (p.honneChecks || []).join('、'),
    p.nextAction || '', p.nextWhen || '',
    '', // 店長コメント（提出時は空）
    recordingUrl, transcriptText,
    (recordingUrl || transcriptText) ? '未処理' : '', // AI採点ステータス
    '', // AI採点結果JSON
    ''  // AI採点日時
  ]);
  return jsonOut_({ ok: true, id: id });
}

function handleFurikaeriComment_(p) {
  const sheet = ensureFurikaeriSheet_();
  const vals = sheet.getDataRange().getValues();
  const idCol = 0, storeCol = 2, subCol = 3, commentCol = vals[0].indexOf('店長コメント');

  // トークンが無ければ本部（全件コメント可）。あれば自店の行にのみ許可
  let allowedKeys = null;
  if (p.token) {
    const accessRows = readAccessControlRaw_();
    const matched = accessRows.filter(function(ac) { return ac.token === p.token; });
    if (!matched.length) return jsonOut_({ error: 'invalid_token' });
    allowedKeys = new Set(
      matched.map(function(m) { return normKey_(m.name) + '|' + normKey_(m.sub); })
    );
  }

  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][idCol]) === String(p.id)) {
      if (allowedKeys) {
        const key = normKey_(vals[i][storeCol]) + '|' + normKey_(vals[i][subCol]);
        if (!allowedKeys.has(key)) return jsonOut_({ error: 'forbidden' });
      }
      sheet.getRange(i + 1, commentCol + 1).setValue(p.comment || '');
      return jsonOut_({ ok: true });
    }
  }
  return jsonOut_({ error: 'not_found' });
}

/* ══════════ AI採点（Gemini連携・2026-08-18追加） ══════════
 *
 * 「ふりかえりログ」に録音URL or 文字起こしが入っている行を、時間主導トリガー
 * （runAiScoringBatch_）が定期的に見つけて採点し、結果をJSON文字列で書き戻す。
 * 判定ロジックは採点ロジック仕様_v4.md（人物金時敵5＋LIFEX4軸4＋型6＝15項目、
 * 自己評価との差異検出、研修事例候補フラグ）をそのままシステムプロンプト化。
 *
 * 【事前準備（徳田さん側・1回だけ）】
 *   1. GASエディタ左メニュー「プロジェクトの設定」→「スクリプト プロパティ」で
 *      キー: GEMINI_API_KEY / 値: 取得したGemini APIキー を追加
 *   2. 関数選択ドロップダウンで createAiScoringTrigger を選んで一度だけ実行
 *      （15分おきの自動実行トリガーが設定される。再実行しても重複しない）
 *   3. 動作確認だけしたい時は runAiScoringNow を実行すると即座に1回分処理される
 *
 * 【既知の制約（未検証・要テスト）】
 *   ・録音ファイルが非常に大きい場合（2時間超の高音質録音等）、GASのUrlFetchApp
 *     リクエストサイズ上限（実測ベースでおよそ50MB前後）に引っかかる可能性がある。
 *     引っかかった場合はGoogleドライブ側で圧縮した音声のみのファイルに変換する
 *     運用ルールを別途検討する。
 *   ・APIキー未取得のため、Gemini呼び出し部分は実機での動作確認がまだできていない。
 *     キー取得後、まず1件をrunAiScoringNowで試してから本番の21社に展開すること。
 */

const GEMINI_MODEL_ = 'gemini-2.5-flash';

function getGeminiApiKey_() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
}

// 徳田さんが手動で1回だけ実行する設定関数（末尾に_を付けず、実行ドロップダウンに出るようにしてある）
function createAiScoringTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runAiScoringBatch_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runAiScoringBatch_').timeBased().everyMinutes(15).create();
  Logger.log('AI採点の自動実行トリガーを15分間隔で設定しました');
}

// 動作確認用に今すぐ1バッチ実行したい時にこれを手動実行する
function runAiScoringNow() {
  runAiScoringBatch_();
}

function runAiScoringBatch_() {
  const apiKey = getGeminiApiKey_();
  if (!apiKey) {
    Logger.log('GEMINI_API_KEYが未設定のため、AI採点をスキップしました');
    return;
  }

  const sheet = ensureFurikaeriSheet_();
  const vals = sheet.getDataRange().getValues();
  const headers = vals[0];
  const col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    if (String(row[col['ID']]).trim() === '') continue;
    if (row[col['AI採点ステータス']] !== '未処理') continue;

    // 多重実行防止のため先に「処理中」に更新してから重い処理に入る
    sheet.getRange(r + 1, col['AI採点ステータス'] + 1).setValue('処理中');
    SpreadsheetApp.flush();

    try {
      const result = scoreOneDeal_(apiKey, row, col);
      sheet.getRange(r + 1, col['AI採点結果JSON'] + 1).setValue(JSON.stringify(result));
      sheet.getRange(r + 1, col['AI採点日時'] + 1).setValue(
        Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')
      );
      sheet.getRange(r + 1, col['AI採点ステータス'] + 1).setValue('完了');
    } catch (err) {
      sheet.getRange(r + 1, col['AI採点結果JSON'] + 1).setValue(JSON.stringify({ error: String(err) }));
      sheet.getRange(r + 1, col['AI採点ステータス'] + 1).setValue('要確認（エラー）');
      Logger.log('AI採点エラー（行' + (r + 1) + '）: ' + err);
    }
    SpreadsheetApp.flush();
  }
}

function scoreOneDeal_(apiKey, row, col) {
  const recordingUrl = String(row[col['録音ファイルURL']] || '').trim();
  const transcriptText = String(row[col['文字起こしテキスト']] || '').trim();
  if (!recordingUrl && !transcriptText) {
    throw new Error('録音ファイルURL・文字起こしテキストのどちらも空です');
  }

  let fileUri = null;
  let fileMimeType = null;
  if (recordingUrl) {
    const fileId = extractDriveFileId_(recordingUrl);
    if (!fileId) throw new Error('録音ファイルのURLからIDを取得できませんでした: ' + recordingUrl);
    const blob = DriveApp.getFileById(fileId).getBlob();
    const uploaded = uploadFileToGemini_(blob, apiKey);
    fileUri = uploaded.uri;
    fileMimeType = uploaded.mimeType;
  }

  const selfReport = buildSelfReportText_(row, col);
  const promptParts = [{ text: buildScoringUserPrompt_(selfReport, transcriptText, !!recordingUrl) }];
  if (fileUri) {
    promptParts.push({ fileData: { fileUri: fileUri, mimeType: fileMimeType } });
  }

  return callGeminiGenerateContent_(apiKey, promptParts);
}

function buildSelfReportText_(row, col) {
  const pairs = [
    ['人', '人_つかめたこと', '人_次に聞くこと'],
    ['物', '物_つかめたこと', '物_次に聞くこと'],
    ['金', '金_つかめたこと', '金_次に聞くこと'],
    ['時', '時_つかめたこと', '時_次に聞くこと'],
    ['敵', '敵_つかめたこと', '敵_次に聞くこと']
  ];
  return pairs.map(function(p) {
    return p[0] + '｜つかめたこと: ' + (row[col[p[1]]] || '（記入なし）') +
      ' ／ 次に聞くこと: ' + (row[col[p[2]]] || '（記入なし）');
  }).join('\n');
}

function buildScoringUserPrompt_(selfReport, transcriptText, hasRecording) {
  let prompt = V4_SYSTEM_PROMPT_;
  prompt += '\n\n## この商談の営業担当の自己評価（ふりかえりシート・PART Aのみ）\n' + selfReport;
  if (hasRecording) {
    prompt += '\n\n## 入力データ\n録音・録画ファイルを添付しています。これを直接解析して採点してください。';
  } else {
    prompt += '\n\n## 入力データ（文字起こしテキストのみ・音声データなし）\n' + transcriptText +
      '\n\n※音声データが無いため、「喋り方」の項目は判定不可（N/A）としてください。';
  }
  return prompt;
}

const V4_SYSTEM_PROMPT_ =
  'あなたはLIFEX（工務店向けFC事業）の商談を採点するコーチです。' +
  '「商談ふりかえりシート（人物金時敵・深掘り版）」とLIFEXの勝ち筋4軸（トータルコスト・デザイン・性能・標準装備）の融合基準（v4）で、' +
  '商談の録音・録画（または文字起こし）を読み、15項目で採点してください。\n\n' +
  '## 大原則\n' +
  '・減点方式にしない。どの項目もまず「できていた点」を探してから伸びしろを示す。\n' +
  '・本人が気づいていない良さこそ最大の価値。良い点は積極的に拾う。\n' +
  '・他店・他の営業担当との比較や優劣付けはしない。\n' +
  '・キーワードの一致ではなく、発言の意味・文脈を理解して判定する。\n' +
  '・発言者ラベルが無くても、話題の主導権や発言内容から営業／お客様を文脈で判定する。\n\n' +
  '## PART A ヒアリング力（人物金時敵・5項目）\n' +
  '見るポイント：人=決裁者・キーマン・他に相談する相手の有無／物=土地条件・要望・提案への反応／' +
  '金=総予算感・資金調達方法・具体的な数字／時=入居希望時期・決定期限／敵=検討中の他社・現状の不満\n' +
  'スコア基準：0=話題が出ていない／1=出たが深掘りされず表面的／2=情報はあるが重要な深掘りが1つ以上抜けている／' +
  '3=しっかり深掘りされ具体的な数字・決定事項まで確認できている／4=3に加え次のアクション（確認事項・期限）まで接続できている\n\n' +
  '## PART B 訴求力（LIFEXの勝ち筋4軸・4項目）※採点基準の核心\n' +
  '知識を一方的に説明できたかではなく、PART A（人物金時敵）のヒアリングで拾った顧客の真意・悩みに対して、この4軸がどう「解決」として機能したかを見る。\n' +
  '見るポイント：トータルコスト=PART A「金」で見えた予算の不安・競合比較への解決になっていたか／' +
  'デザイン=PART A「物」の要望への解決になっていたか／性能=PART A「敵」で見えた不満・競合の性能訴求への解決になっていたか／' +
  '標準装備=PART A「物」「金」で見えた要望・コスト意識への解決になっていたか\n' +
  'スコア基準：0=話題が出ていない／1=説明はあるがどの顧客ニーズへの解決か不明（カタログ的説明）／' +
  '2=顧客の要望への言及はあるがPART Aの悩みとの結びつきが浅い／3=PART Aの具体的な悩みへの解決として提示され好意的な反応まで確認できている／' +
  '4=3に加え顧客自身が自分の言葉で納得を示し具体的な選択・決定に至っている\n' +
  '※PART Aで該当する言及が無かった場合は一般的な訴求として通常基準（顧客の反応の有無）で判定し、その旨を根拠に明記する。\n\n' +
  '## PART C 商談の「型」の実践度（6項目）\n' +
  '見るポイント：会話法=共感・復唱・関連質問／間の取り方=考える時間を待てているか・話を遮っていないか／' +
  'テストクロージング=途中の温度感確認／次の目標設定=終わりに次アクション・期限を具体的に提示できたか／' +
  '伝え方のわかりやすさ=専門用語をお客様目線の言葉に翻訳できているか／' +
  '喋り方=声のトーン・話す速さ・熱意が反応に応じて調整できているか（音声データが無い場合は必ずN/A＝判定不可とし、推測で点数を付けない）\n' +
  'スコア基準：0=場面が見当たらない／1=使おうとしているが不自然・タイミングが悪い／2=使えているが効果的に機能していない場面がある／' +
  '3=効果的に使えている／4=効果的に使えていて商談の流れが良い方向に変わった場面がある\n\n' +
  '## 自己評価との差異検出（対象はPART Aの5項目のみ）\n' +
  '営業担当の自己評価（ふりかえりシートの記述）と、あなたが商談から独立に読み取った内容を比較する。\n' +
  '・一致：自己記述とほぼ一致→自己認識が的確という加点コメント\n' +
  '・気づいていない強み：実際は深く聞けているのに本人の記述は簡潔・言及なし→最優先で拾い「実は〇〇がとても良かった」という本人に伝える言い回しで出力\n' +
  '・自己評価とのギャップ：本人は「聞けた」と書いているが実際の深掘りは浅い→個別フィードバック限定（一覧には出さない）扱いとし、'
  + '「間違っている」ではなく「実はもっとできていた」という成長方向の言い回しにする\n\n' +
  '## 研修事例候補\n' +
  '高得点かつ「気づいていない強み」が目立つ回は研修の模範事例候補としてフラグする。\n\n' +
  '出力は必ず指定されたJSON形式（スキーマ）に従い、各スコア項目には根拠となる具体的なやり取りを良かった点・伸びしろの両方に含めること。';

const SCORING_ITEM_SCHEMA_ = {
  type: 'OBJECT',
  properties: {
    key: { type: 'STRING' },
    score: { type: 'INTEGER' },
    good: { type: 'STRING' },
    grow: { type: 'STRING' }
  },
  required: ['key', 'score', 'good', 'grow']
};

const SCORING_ITEM_SCHEMA_NULLABLE_ = {
  type: 'OBJECT',
  properties: {
    key: { type: 'STRING' },
    na: { type: 'BOOLEAN' },
    naReason: { type: 'STRING' },
    score: { type: 'INTEGER' },
    good: { type: 'STRING' },
    grow: { type: 'STRING' }
  },
  required: ['key', 'na', 'score', 'good', 'grow']
};

const SCORING_RESPONSE_SCHEMA_ = {
  type: 'OBJECT',
  properties: {
    summary: { type: 'STRING' },
    partA: { type: 'ARRAY', items: SCORING_ITEM_SCHEMA_ },
    partB: { type: 'ARRAY', items: SCORING_ITEM_SCHEMA_ },
    partC: { type: 'ARRAY', items: SCORING_ITEM_SCHEMA_NULLABLE_ },
    selfDiff: {
      type: 'OBJECT',
      properties: {
        noticedStrengths: { type: 'STRING' },
        gapsPrivate: { type: 'STRING' }
      },
      required: ['noticedStrengths', 'gapsPrivate']
    },
    overallComment: { type: 'STRING' },
    nextOneChange: { type: 'STRING' },
    trainingCandidate: {
      type: 'OBJECT',
      properties: {
        flag: { type: 'BOOLEAN' },
        reason: { type: 'STRING' }
      },
      required: ['flag', 'reason']
    }
  },
  required: ['summary', 'partA', 'partB', 'partC', 'selfDiff', 'overallComment', 'nextOneChange', 'trainingCandidate']
};

function callGeminiGenerateContent_(apiKey, userParts) {
  const body = {
    contents: [{ role: 'user', parts: userParts }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: SCORING_RESPONSE_SCHEMA_
    }
  };
  const resp = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL_ + ':generateContent?key=' + apiKey,
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    }
  );
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  let json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    throw new Error('Gemini応答がJSONとして解釈できませんでした: ' + text.slice(0, 500));
  }
  if (code !== 200) {
    throw new Error('Gemini APIエラー（' + code + '）: ' + (json.error && json.error.message || text.slice(0, 500)));
  }
  const part = json.candidates && json.candidates[0] && json.candidates[0].content &&
    json.candidates[0].content.parts && json.candidates[0].content.parts[0];
  if (!part || !part.text) {
    throw new Error('Gemini応答の形式が想定外です: ' + text.slice(0, 500));
  }
  return JSON.parse(part.text);
}

// Googleドライブの共有リンクからファイルIDを取り出す
// 対応形式: https://drive.google.com/file/d/FILE_ID/view?... / https://drive.google.com/open?id=FILE_ID
function extractDriveFileId_(url) {
  const s = String(url || '');
  const m = s.match(/\/d\/([a-zA-Z0-9_-]{15,})/) || s.match(/[?&]id=([a-zA-Z0-9_-]{15,})/);
  return m ? m[1] : null;
}

// Gemini Files APIへ音声・録画ファイルをアップロードし、file_uriを取得する
// （2時間超の録音等、非常に大きいファイルではUrlFetchAppのリクエストサイズ上限に
//   引っかかる可能性があり未検証。APIキー取得後に実データで要確認）
function uploadFileToGemini_(blob, apiKey) {
  const boundary = 'lifexshodan' + Utilities.getUuid().replace(/-/g, '');
  const mimeType = blob.getContentType() || 'application/octet-stream';
  const metadata = JSON.stringify({ file: { display_name: (blob.getName() || 'shodan-recording') } });
  const CRLF = '\r\n';

  const preambleBytes = Utilities.newBlob(
    '--' + boundary + CRLF +
    'Content-Type: application/json; charset=UTF-8' + CRLF + CRLF +
    metadata + CRLF +
    '--' + boundary + CRLF +
    'Content-Type: ' + mimeType + CRLF + CRLF
  ).getBytes();
  const closingBytes = Utilities.newBlob(CRLF + '--' + boundary + '--').getBytes();
  const bodyBytes = preambleBytes.concat(blob.getBytes()).concat(closingBytes);

  const resp = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/upload/v1beta/files?key=' + apiKey,
    {
      method: 'post',
      contentType: 'multipart/related; boundary=' + boundary,
      payload: bodyBytes,
      muteHttpExceptions: true
    }
  );
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  let json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    throw new Error('Geminiファイルアップロードの応答がJSONとして解釈できませんでした: ' + text.slice(0, 500));
  }
  if (code !== 200 || !json.file || !json.file.uri) {
    throw new Error('Geminiファイルアップロードに失敗しました（' + code + '）: ' + text.slice(0, 500));
  }

  // 動画・長尺音声はACTIVEになるまで数秒かかることがあるため、状態を確認して待つ
  let state = json.file.state;
  let tries = 0;
  while (state === 'PROCESSING' && tries < 10) {
    Utilities.sleep(3000);
    const check = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/' + json.file.name + '?key=' + apiKey,
      { muteHttpExceptions: true }
    );
    const checkJson = JSON.parse(check.getContentText());
    state = checkJson.state;
    if (state === 'ACTIVE') { json.file = checkJson; break; }
    tries++;
  }
  if (state === 'FAILED') {
    throw new Error('Geminiファイル処理が失敗しました: ' + JSON.stringify(json.file));
  }

  return { uri: json.file.uri, mimeType: json.file.mimeType || mimeType };
}

/* ══════════ MTGアーカイブ視聴ログ（2026-09-14追加） ══════════
 *
 * 目標達成MTGアーカイブ（lifex-mtg-archive.html）で、加盟店様が各回のタイトル／
 * サムネイルをクリックした（＝YouTubeへ視聴に行った）タイミングを記録する。
 * トークンは既存の「アクセス管理」シート（KPIダッシュボード・ふりかえりと共通）を流用。
 * 店舗識別は ?s=トークン 付きの個別URLをアーカイブページに配布する前提。
 *
 *   doGet(?archiveLog=1&s=TOKEN&session=回&video=動画ID) … 1件記録（画像ビーコンで呼ぶ想定）
 *   doGet(?archiveData=1[&s=TOKEN])                      … ログ一覧取得（本部用可視化ページから呼ぶ。
 *                                                            s省略時は全件、指定時はその店舗分のみ）
 *
 * 「視聴ログ」シートは初回アクセス時に自動作成。既存シートには一切影響しない。
 */
const ARCHIVE_LOG_SHEET = '視聴ログ';
const ARCHIVE_LOG_HEADERS = ['記録日時', '店舗（グループ表示名）', '回', '動画ID', 'トークン'];

function ensureArchiveLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = findSheetByName(ss, ARCHIVE_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ARCHIVE_LOG_SHEET);
    sheet.appendRow(ARCHIVE_LOG_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, ARCHIVE_LOG_HEADERS.length).setFontWeight('bold');
    sheet.setColumnWidths(1, ARCHIVE_LOG_HEADERS.length, 160);
  }
  return sheet;
}

function handleArchiveLogRecord_(p) {
  const token = p.s || '';
  const accessRows = readAccessControlRaw_();
  const matched = accessRows.filter(function(ac) { return ac.token === token; });
  if (!matched.length) return jsonOut_({ ok: false, error: 'invalid_token' });

  const sheet = ensureArchiveLogSheet_();
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  const groupName = matched[0].group || matched[0].name;
  sheet.appendRow([now, groupName, p.session || '', p.video || '', token]);
  return jsonOut_({ ok: true });
}

function handleArchiveLogList_(p) {
  const sheet = ensureArchiveLogSheet_();
  const vals = sheet.getDataRange().getValues();
  const headers = vals[0];
  let rows = vals.slice(1).filter(function(r) { return String(r[0]).trim() !== ''; });

  const token = p.s || null;
  if (token) {
    rows = rows.filter(function(r) { return String(r[4]) === token; });
  }

  const entries = rows.map(function(r) {
    const obj = {};
    headers.forEach(function(h, i) { obj[h] = r[i]; });
    return obj;
  });

  // 全店舗一覧（グループ表示名でユニーク化）も併せて返す。「まだ1回もログがない店」の可視化に使う
  const accessRows = readAccessControlRaw_();
  const groupSet = {};
  accessRows.forEach(function(ac) { groupSet[ac.group || ac.name] = true; });

  return jsonOut_({ entries: entries, stores: Object.keys(groupSet) });
}
