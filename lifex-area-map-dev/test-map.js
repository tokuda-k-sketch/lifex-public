// index.html のサンプルデータと map-data.js の整合性チェック
// 実行: node test-map.js
'use strict';
const fs = require('fs');

eval(fs.readFileSync('map-data.js', 'utf8').replace('window.MAP_DATA', 'global.MAP_DATA'));
const M = global.MAP_DATA;

const html = fs.readFileSync('index.html', 'utf8');
const m = html.match(/\/\*SAMPLE_START\*\/([\s\S]*?)\/\*SAMPLE_END\*\//);
if (!m) { console.error('NG: SAMPLEマーカーが見つからない'); process.exit(1); }
const SAMPLE_MEMBERS = (0, eval)(m[1].replace('const SAMPLE_MEMBERS =', '') + '');

let fail = 0;
const ok = (cond, label) => { console.log((cond ? 'OK ' : 'NG ') + label); if (!cond) fail++; };

// 1. 基本データ
ok(M.munis.length > 1800, `市区町村数 ${M.munis.length}（1800超）`);
ok(M.prefs.length === 47, `都道府県図形 ${M.prefs.length}`);
ok(Object.keys(M.prefBBox).length === 47, 'prefBBox 47県分');
ok(M.munis.every(x => x.d.startsWith('M') && !x.d.includes('NaN')), '全パスが M 開始・NaNなし');
ok(!M.munis.some(x => x.c === '13421'), '小笠原村は除外済み');

// 2. サンプルの地名が全て一意に解決するか
const byKey = new Map();
M.munis.forEach(x => byKey.set(x.p + '/' + x.full, x));
const unresolved = [];
let total = 0;
for (const mem of SAMPLE_MEMBERS) {
  for (const [p, n] of [...mem.purchased, ...mem.buildable]) {
    total++;
    if (!byKey.has(p + '/' + n)) unresolved.push(p + n);
  }
}
ok(unresolved.length === 0, `サンプル地名 ${total}件すべて解決` + (unresolved.length ? ' → 未解決: ' + unresolved.join(', ') : ''));

// 3. 政令市の区が引けるか（2分割運用の前提）
['大阪府/堺市堺区', '宮城県/仙台市泉区', '埼玉県/さいたま市大宮区', '北海道/札幌市中央区'].forEach(k => {
  ok(byKey.has(k), `政令市の区: ${k}`);
});

// 4. 全国の政令市カバレッジ（区を持つ市の数=20のはず）
const seirei = new Set(M.munis.filter(x => x.g).map(x => x.g));
ok(seirei.size === 20, `政令指定都市 ${seirei.size}市（区単位データあり）: ` + [...seirei].join('、'));

// 5. full名の県内重複がないか（検索・照合の前提）
const dup = [];
const seen = new Set();
M.munis.forEach(x => {
  const k = x.p + '/' + x.full;
  if (seen.has(k)) dup.push(k); else seen.add(k);
});
ok(dup.length === 0, '県内での市区町村名の重複なし' + (dup.length ? ' → ' + dup.slice(0, 5).join(', ') : ''));

console.log(fail ? `\n${fail}件失敗` : '\n全チェック通過');
process.exit(fail ? 1 : 0);
