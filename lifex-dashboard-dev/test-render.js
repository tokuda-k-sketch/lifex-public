/**
 * 事前検証スクリプト（デプロイ前のシミュレーション）
 * 1. 広告スプシをCSVで取得し、GAS v2と同じパース処理に通す
 * 2. 本番GAS(v1)からKPIデータを取得
 * 3. 2つを合成した「GAS v2相当のレスポンス」でダッシュボードJSを実行
 * 4. 店舗カードに広告実績が正しく出るか検証
 */
const fs = require('fs');
const path = require('path');

const AD_ID = '1qsCsnB5SZtSA8XUE5OWxptKjX8eRokJkA6D8Bs-QK9k';
const GAS_V1 = 'https://script.google.com/macros/s/AKfycbwaF-NIhkqRmnk3gDsOhWmpQNfzBBCBi6O_efu3zuwRQpfi_0_hWEdNFW4qxpMMqinJ/exec';

// ── 簡易CSVパーサ（クォート・カンマ・改行対応）──
function parseCSV(text) {
  const rows = [[]]; let field = '', inQ = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQ) {
      if (ch === '"') { if (text[i+1] === '"') { field += '"'; i++; } else inQ = false; }
      else field += ch;
    } else if (ch === '"') inQ = true;
    else if (ch === ',') { rows[rows.length-1].push(field); field = ''; }
    else if (ch === '\n') { rows[rows.length-1].push(field); field = ''; rows.push([]); }
    else if (ch !== '\r') field += ch;
  }
  rows[rows.length-1].push(field);
  return rows;
}

// ── GAS v2 のパース関数を読み込む（SpreadsheetApp等は呼ばれない範囲だけ使う）──
const gasSrc = fs.readFileSync(path.join(__dirname, '..', 'lifex-apps-script-v2.js'), 'utf8');
(0, eval)(gasSrc); // parseMonthly / parseReservations / pickAd / num がグローバルに定義される

const fakeSheet = vals => ({ getDataRange: () => ({ getDisplayValues: () => vals }) });

async function main() {
  // 1. 広告スプシ取得
  const monthlyCsv = await (await fetch(`https://docs.google.com/spreadsheets/d/${AD_ID}/export?format=csv&gid=213210306`)).text();
  const reserveCsv = await (await fetch(`https://docs.google.com/spreadsheets/d/${AD_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent('予約集計（LP1）  ')}`)).text();

  const months = parseMonthly(fakeSheet(parseCSV(monthlyCsv)));
  const reservations = parseReservations(fakeSheet(parseCSV(reserveCsv)));

  console.log('== パース結果 ==');
  console.log('月ブロック:', months.map(m => `${m.label}(${m.stores.length}店)`).join(' / '));
  console.log('予約集計:', reservations.map(r => `${r.name}:${r.count}件/¥${r.cost}`).join(', '));

  // 2. 本番GASからKPIデータ
  const v1 = await (await fetch(GAS_V1)).json();
  console.log('KPI店舗数:', v1.stores.length);

  // 3. GAS v2相当のレスポンスを合成
  const payload = { stores: v1.stores, ads: { months, reservations }, updatedAt: 'テスト' };

  // 4. ダッシュボードJSを疑似DOMで実行
  const html = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
  const script = html.match(/<script>([\s\S]*)<\/script>/)[1];

  const elements = {};
  const makeEl = id => ({ id, innerHTML: '', textContent: '', className: '', style: {}, setAttribute(){}, });
  global.document = {
    getElementById: id => elements[id] || (elements[id] = makeEl(id)),
    addEventListener(){}, hidden: false,
  };
  global.fetch = async () => ({ ok: true, json: async () => payload });

  (0, eval)(script);
  await global.init();

  // 5. 検証
  const sg = elements['js-sg'].innerHTML;
  const errors = [];
  const expect = (cond, msg) => { console.log((cond ? 'OK  ' : 'NG  ') + msg); if (!cond) errors.push(msg); };

  expect(sg.includes('広告実績'), '広告実績セクションが描画されている');
  expect(sg.includes('予約単価'), '予約単価が描画されている');
  expect(sg.includes('来場単価'), '来場単価が描画されている');
  expect((sg.match(/ad-sec/g)||[]).length >= v1.stores.length - 2, `店舗カードの大半に広告セクションあり (${(sg.match(/ad-sec/g)||[]).length}/${v1.stores.length})`);

  // いえプロ: 予約1件・累計¥835,900 → 予約単価 ¥835,900
  const iepro = reservations.find(r => r.name.includes('いえプロ'));
  if (iepro && iepro.count > 0) {
    const tanka = '¥' + Math.round(iepro.cost / iepro.count).toLocaleString('ja-JP');
    expect(sg.includes(tanka), `いえプロの予約単価 ${tanka} が表示されている`);
  }
  // 中央建設(出雲店) → 中央建設① にマッチしているか
  expect(!sg.includes('中央建設①が見つかりません'), '中央建設のエイリアス解決');
  // 最新月METAの竹村工務店の配信費が表に出ている（値はスプシから動的に取得）
  const lastMo = months[months.length - 1];
  const takemura = lastMo.stores.find(s => s.name === '竹村工務店');
  if (takemura && takemura.meta.cost != null) {
    const costStr = '¥' + Math.round(takemura.meta.cost).toLocaleString('ja-JP');
    expect(sg.includes(costStr), `竹村工務店 ${lastMo.label}META配信費 ${costStr} が表示されている`);
  }
  // 予約集計に見出し行・別セクションが混入していない
  expect(!reservations.some(r => /加盟店|Google広告|META/.test(r.name)), '予約集計に見出し行が混入していない');
  expect(reservations.length <= 25, `予約集計の件数が妥当 (${reservations.length}件)`);
  // WD → WITHDOM Group のマッチ
  const wdCard = sg.split('sc-name">').find(c => c.startsWith('WD'));
  expect(wdCard && wdCard.includes('広告実績') && !wdCard.includes('広告データなし'), 'WD → WITHDOM Group マッチ');
  // フラワーホーム（広告スプシに無い店）は「広告データなし」
  const flower = sg.split('sc-name">').find(c => c.startsWith('フラワーホーム'));
  if (flower) expect(flower.includes('広告データなし'), 'フラワーホーム → 広告データなし表示');

  console.log('\nステータス表示:', elements['js-status'].textContent);

  // 6. 実データを埋め込んだプレビューHTMLを生成（ブラウザで見た目確認用）
  const previewHtml = html
    .replace('<script>', '<script>\n/* プレビュー用: 取得済みデータを埋め込み・通信なし */\nconst __PRELOADED = ' + JSON.stringify(payload) + ';\nwindow.fetch = async () => ({ ok: true, json: async () => __PRELOADED });\n')
    .replace('<span class="dev-pill">開発版</span>', '<span class="dev-pill">プレビュー（データ固定）</span>');
  fs.writeFileSync(path.join(__dirname, 'preview.html'), previewHtml);
  console.log('プレビュー生成: preview.html');

  console.log(errors.length ? `\n★ 失敗 ${errors.length}件` : '\n★ 全チェック通過');
  process.exit(errors.length ? 1 : 0);
}
main().catch(e => { console.error('テスト実行エラー:', e); process.exit(1); });
