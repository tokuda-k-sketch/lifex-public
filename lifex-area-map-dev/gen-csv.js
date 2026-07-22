// スプシ取り込み用CSVを生成（UTF-8 BOM付き・Googleスプレッドシートにそのままインポート可）
// 実行: node gen-csv.js
//  → 市区町村マスタ.csv（入力規則のプルダウン用・全1,896件）
//  → 施工範囲マスタ_雛形.csv（入力タブの雛形＋記入例）
'use strict';
const fs = require('fs');

eval(fs.readFileSync('map-data.js', 'utf8').replace('window.MAP_DATA', 'global.MAP_DATA'));
const M = global.MAP_DATA;
const BOM = '﻿';

// 市区町村マスタ: 都道府県 / 市区町村 / 検索用（県名+市区町村名）
const master = ['都道府県,市区町村,検索用'];
M.munis.forEach(m => master.push(`${m.p},${m.full},${m.p}${m.full}`));
fs.writeFileSync('市区町村マスタ.csv', BOM + master.join('\r\n'), 'utf8');

// 施工範囲マスタ雛形（記入例入り。実運用時は例の行を消して使う）
const template = [
  '加盟店名,都道府県,市区町村,区分,備考',
  '（記入例）〇〇工務店,高知県,全域,エリア購入済,「全域」で県内全市区町村',
  '（記入例）〇〇ホーム,大阪府,堺市,施工可能,政令市名だけ書くと全区に展開',
  '（記入例）〇〇ホーム,大阪府,堺市堺区,エリア購入済,区単位ならこの書き方',
  '（記入例）〇〇建設,埼玉県,川口市,施工可能,',
];
fs.writeFileSync('施工範囲マスタ_雛形.csv', BOM + template.join('\r\n'), 'utf8');

console.log(`市区町村マスタ.csv: ${M.munis.length}件 / 施工範囲マスタ_雛形.csv: 記入例${template.length - 1}行`);
