// muni_s0001.json / pref_s0001.json (TopoJSON) → map-data.js (SVGパス集)
// 実行: node convert.js
// - 政令指定都市は行政区単位のまま保持（後からグループ化で2分割に対応）
// - 沖縄県は左上にインセット移動（経度+7 / 緯度+14）
// - 小笠原村(13421)は表示対象外（南鳥島まで含むと地図が間延びするため）
'use strict';
const fs = require('fs');

const OKI_SHIFT = { lon: 5.5, lat: 15 }; // 佐渡・能登と重ならない位置に調整済み
const EXCLUDE_CODES = new Set(['13421']);
const COS = Math.cos((36 * Math.PI) / 180);
const WIDTH = 840;

function decodeArcs(topo) {
  const [sx, sy] = topo.transform.scale;
  const [tx, ty] = topo.transform.translate;
  return topo.arcs.map((arc) => {
    let x = 0, y = 0;
    return arc.map(([dx, dy]) => {
      x += dx; y += dy;
      return [x * sx + tx, y * sy + ty];
    });
  });
}

// ring = arcインデックス列 → [lon,lat]列（負値 ~i は逆順）
function assembleRing(ringArcs, arcs) {
  const pts = [];
  for (const idx of ringArcs) {
    let seg = idx >= 0 ? arcs[idx] : arcs[~idx].slice().reverse();
    if (pts.length) seg = seg.slice(1); // 接続点の重複を除去
    pts.push(...seg);
  }
  return pts;
}

function geomRings(geom, arcs) {
  // Polygon: [ring...] / MultiPolygon: [[ring...]...] → フラットなring配列
  const out = [];
  if (geom.type === 'Polygon') {
    for (const r of geom.arcs) out.push(assembleRing(r, arcs));
  } else if (geom.type === 'MultiPolygon') {
    for (const poly of geom.arcs) for (const r of poly) out.push(assembleRing(r, arcs));
  }
  return out;
}

function loadFeatures(file, isPref) {
  const topo = JSON.parse(fs.readFileSync(file, 'utf8'));
  const objName = Object.keys(topo.objects)[0];
  const arcs = decodeArcs(topo);
  const feats = [];
  for (const g of topo.objects[objName].geometries) {
    const p = g.properties || {};
    const pref = p.N03_001 || p.name || '';
    const code = p.N03_007 || '';
    if (!isPref && EXCLUDE_CODES.has(code)) continue;
    feats.push({
      pref,
      gun: p.N03_003 || null,   // 郡名 or 政令市名
      name: p.N03_004 || pref,  // 市区町村名（都道府県ファイルでは県名）
      code,
      rings: geomRings(g, arcs),
    });
  }
  return feats;
}

// 沖縄シフト → 単純図法（x=lon*cos36, y=-lat）
function project([lon, lat], pref) {
  if (pref === '沖縄県') { lon += OKI_SHIFT.lon; lat += OKI_SHIFT.lat; }
  return [lon * COS, -lat];
}

console.log('読み込み中…');
const munis = loadFeatures('muni_s0001.json', false);
let prefs = [];
try { prefs = loadFeatures('pref_s0001.json', true); }
catch (e) { console.log('都道府県ファイルはスキップ:', e.message); }
console.log(`市区町村: ${munis.length} / 都道府県図形: ${prefs.length}`);

// 全体バウンディングボックス
let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
for (const f of munis) for (const ring of f.rings) for (const pt of ring) {
  const [x, y] = project(pt, f.pref);
  if (x < minX) minX = x; if (x > maxX) maxX = x;
  if (y < minY) minY = y; if (y > maxY) maxY = y;
}
const PAD = 12; // 外周余白（沖縄インセット枠の見切れ防止）
const k = (WIDTH - PAD * 2) / (maxX - minX);
const HEIGHT = Math.round((maxY - minY) * k + PAD * 2);
const toXY = (pt, pref) => {
  const [x, y] = project(pt, pref);
  return [((x - minX) * k + PAD), ((y - minY) * k + PAD)];
};

function ringsToPath(rings, pref, bbox) {
  let d = '';
  for (const ring of rings) {
    ring.forEach((pt, i) => {
      const [x, y] = toXY(pt, pref);
      if (bbox) {
        if (x < bbox[0]) bbox[0] = x; if (x > bbox[2]) bbox[2] = x;
        if (y < bbox[1]) bbox[1] = y; if (y > bbox[3]) bbox[3] = y;
      }
      d += (i === 0 ? 'M' : 'L') + x.toFixed(1) + ' ' + y.toFixed(1);
    });
    d += 'Z';
  }
  return d;
}

// 県ズーム枠の計算から遠隔離島を除外（表示はする。枠が間延びするのを防ぐ）
// 東京都: 島嶼部（コード13360番台以降）／鹿児島県: 奄美群島（大島郡・奄美市）
const excludeFromBBox = (f) =>
  (f.pref === '東京都' && Number(f.code) >= 13360) ||
  (f.pref === '鹿児島県' && (f.gun === '大島郡' || f.name === '奄美市'));

const prefBBox = {}; // 県ごとのbbox（沖縄はインセット後の座標）
const outMunis = munis.map((f) => {
  if (!prefBBox[f.pref]) prefBBox[f.pref] = [Infinity, Infinity, -Infinity, -Infinity];
  const d = ringsToPath(f.rings, f.pref, excludeFromBBox(f) ? null : prefBBox[f.pref]);
  const isSeirei = f.gun && f.gun.endsWith('市');
  return {
    c: f.code,
    p: f.pref,
    g: isSeirei ? f.gun : null,                    // 政令市名（区の親）
    n: f.name,
    full: isSeirei ? f.gun + f.name : f.name,      // 「堺市堺区」「河南町」
    d,
  };
});

// 県内で同名の市区町村（例: 北海道の泊村×2）は郡名付きで区別
{
  const count = {};
  for (const m of outMunis) count[m.p + '/' + m.full] = (count[m.p + '/' + m.full] || 0) + 1;
  const rawByCode = Object.fromEntries(munis.map((f) => [f.code, f]));
  for (const m of outMunis) {
    if (count[m.p + '/' + m.full] > 1) {
      const gun = rawByCode[m.c].gun;
      if (gun) m.full = gun + m.n;
    }
  }
}

const outPrefs = prefs.map((f) => ({ p: f.pref, d: ringsToPath(f.rings, f.pref, null) }));

// 沖縄インセット枠
const ob = prefBBox['沖縄県'];
const okinawaBox = ob ? [ob[0] - 8, ob[1] - 8, ob[2] - ob[0] + 16, ob[3] - ob[1] + 16].map((v) => +v.toFixed(1)) : null;

const data = {
  viewBox: [0, 0, WIDTH, HEIGHT],
  okinawaBox,
  prefBBox: Object.fromEntries(Object.entries(prefBBox).map(([k2, b]) => [k2, b.map((v) => +v.toFixed(1))])),
  prefs: outPrefs,
  munis: outMunis,
};
const js = 'window.MAP_DATA = ' + JSON.stringify(data) + ';\n';
fs.writeFileSync('map-data.js', js);
console.log(`map-data.js 出力完了: ${(js.length / 1024 / 1024).toFixed(2)} MB / viewBox ${WIDTH}x${HEIGHT}`);
console.log('都道府県数:', Object.keys(prefBBox).length);
