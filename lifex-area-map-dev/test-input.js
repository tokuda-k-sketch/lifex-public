// index.html の rowsToMembers（スプシ行→加盟店データ変換）の単体テスト
// 実行: node test-input.js
// 注意: 'use strict' を付けると直接evalの関数宣言がこのスコープに出てこないので付けない
const fs = require('fs');

eval(fs.readFileSync('map-data.js', 'utf8').replace('window.MAP_DATA', 'global.MAP_DATA'));
const M = global.MAP_DATA;

// index.html と同じインデックスを構築
const byKey = new Map(), prefMunis = new Map(), cityWards = new Map();
M.munis.forEach(m => {
  byKey.set(m.p + '/' + m.full, m);
  if (!prefMunis.has(m.p)) prefMunis.set(m.p, []);
  prefMunis.get(m.p).push(m.full);
  if (m.g) {
    const k = m.p + '/' + m.g;
    if (!cityWards.has(k)) cityWards.set(k, []);
    cityWards.get(k).push(m.full);
  }
});

// index.html から rowsToMembers 本体を抽出して評価（ロジックの二重管理を避ける）
const html = fs.readFileSync('index.html', 'utf8');
const m = html.match(/\/\*EXPAND_START\*\/([\s\S]*?)\/\*EXPAND_END\*\//);
if (!m) { console.error('NG: EXPANDマーカーが見つからない'); process.exit(1); }
eval(m[1]);

let fail = 0;
const ok = (cond, label) => { console.log((cond ? 'OK ' : 'NG ') + label); if (!cond) fail++; };

// 1. 正式名そのまま
let r = rowsToMembers([{ member: 'A社', pref: '埼玉県', muni: '川口市', kind: 'エリア購入済' }]);
ok(r.members.length === 1 && r.members[0].purchased.length === 1 && r.warnings.length === 0, '正式名1件 → 購入済1件');

// 2. 政令市名 → 全区展開
r = rowsToMembers([{ member: 'A社', pref: '大阪府', muni: '堺市', kind: '施工可能' }]);
ok(r.members[0].buildable.length === 7, '堺市 → 7区に展開（実際: ' + r.members[0].buildable.length + '）');

// 3. 全域 → 県内全市区町村
const kochiCount = prefMunis.get('高知県').length;
r = rowsToMembers([{ member: 'B社', pref: '高知県', muni: '全域', kind: 'エリア購入済' }]);
ok(r.members[0].purchased.length === kochiCount, `高知県全域 → ${kochiCount}市区町村に展開`);

// 4. 認識できない地名 → 警告に積んで行はスキップ
r = rowsToMembers([{ member: 'C社', pref: '大阪府', muni: '存在しない市', kind: '施工可能' }]);
ok(r.members.length === 0 && r.warnings.length === 1, '不明な地名 → 警告1件・データ0件');

// 5. 区分の表記ゆれ（「購入」「施工」を含めばOK）
r = rowsToMembers([
  { member: 'D社', pref: '宮城県', muni: '名取市', kind: '購入済' },
  { member: 'D社', pref: '宮城県', muni: '岩沼市', kind: '施工可' },
  { member: 'D社', pref: '宮城県', muni: '富谷市', kind: 'その他' },
]);
ok(r.members[0].purchased.length === 1 && r.members[0].buildable.length === 1 && r.warnings.length === 1,
  '区分ゆれ吸収＋不明区分は警告');

// 6. 購入エリアの重複 → 警告
r = rowsToMembers([
  { member: 'E社', pref: '埼玉県', muni: '蕨市', kind: 'エリア購入済' },
  { member: 'F社', pref: '埼玉県', muni: '蕨市', kind: 'エリア購入済' },
]);
ok(r.warnings.length === 1 && r.warnings[0].includes('重複'), '購入エリア重複 → 警告: ' + (r.warnings[0] || 'なし'));

// 7. 施工可能の被りは正常データ（警告なし・両社に入る）
r = rowsToMembers([
  { member: 'G社', pref: '埼玉県', muni: '戸田市', kind: '施工可能' },
  { member: 'H社', pref: '埼玉県', muni: '戸田市', kind: '施工可能' },
]);
ok(r.members.length === 2 && r.warnings.length === 0, '施工可能の被り → 警告なしで両社に登録');

// 8. 空行・空白入りは無視/トリム
r = rowsToMembers([
  { member: '', pref: '', muni: '', kind: '' },
  { member: ' I社 ', pref: ' 高知県 ', muni: ' 高知市 ', kind: ' エリア購入済 ' },
]);
ok(r.members.length === 1 && r.members[0].name === 'I社' && r.members[0].purchased.length === 1, '空行スキップ＋トリム');

// ---- collapseToRows（塗った内容→スプシ行の自動圧縮）----
const mc = html.match(/\/\*COLLAPSE_START\*\/([\s\S]*?)\/\*COLLAPSE_END\*\//);
if (!mc) { console.error('NG: COLLAPSEマーカーが見つからない'); process.exit(1); }
eval(mc[1]);

const key = (p, f) => p + '/' + f;

// 9. 県内全部塗り → 「全域」1行に圧縮
let d = { purchased: new Set(prefMunis.get('高知県').map(f => key('高知県', f))), buildable: new Set() };
let rows = collapseToRows(d);
ok(rows.length === 1 && rows[0].muni === '全域' && rows[0].kind === 'エリア購入済', '高知県全部 → 「全域」1行');

// 10. 政令市の全区 → 市名1行に圧縮
d = { purchased: new Set(), buildable: new Set(cityWards.get('大阪府/堺市').map(f => key('大阪府', f))) };
rows = collapseToRows(d);
ok(rows.length === 1 && rows[0].muni === '堺市' && rows[0].kind === '施工可能', '堺市7区 → 「堺市」1行');

// 11. 全区に1つ足りない → 個別の区で出力
const wards6 = cityWards.get('大阪府/堺市').filter(f => f !== '堺市堺区');
d = { purchased: new Set(), buildable: new Set(wards6.map(f => key('大阪府', f))) };
rows = collapseToRows(d);
ok(rows.length === 6 && rows.every(r => r.muni.startsWith('堺市')), '堺区抜きの6区 → 個別6行');

// 12. 混在（全域＋別県の個別＋区分違い）
d = {
  purchased: new Set(prefMunis.get('高知県').map(f => key('高知県', f))),
  buildable: new Set([key('埼玉県', '川口市'), key('埼玉県', '蕨市')]),
};
rows = collapseToRows(d);
ok(rows.length === 3 &&
   rows.filter(r => r.kind === 'エリア購入済').length === 1 &&
   rows.filter(r => r.kind === '施工可能').length === 2, '全域＋個別2市の混在 → 3行');

// 13. 圧縮結果を rowsToMembers に戻すと元の市区町村数に一致（往復テスト）
const back = rowsToMembers(rows.map(r => ({ member: 'X社', pref: r.pref, muni: r.muni, kind: r.kind })));
ok(back.members[0].purchased.length === prefMunis.get('高知県').length &&
   back.members[0].buildable.length === 2 && back.warnings.length === 0, '圧縮→展開の往復で元に戻る');

console.log(fail ? `\n${fail}件失敗` : '\n全チェック通過');
process.exit(fail ? 1 : 0);
