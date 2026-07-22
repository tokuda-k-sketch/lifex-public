/**
 * LIFEX KPI データ配信スクリプト
 * ─────────────────────────────────────────────
 * 【設置場所】Googleスプレッドシート → 拡張機能 → Apps Script
 * 【シート名】「KPIデータ」という名前のシートを作成してください
 *
 * 【スプレッドシートの列構成】
 *   A列: 店舗名（例: 久留米建設）
 *   B列: 担当者・支店名（例: 髙口様）
 *   C列: ステータス2  ／ 会員登録
 *   D列: ステータス3  ／ 連絡が繋がらない
 *   E列: ステータス4  ／ 通話済み（追客中）
 *   F列: ステータス5  ／ 来場予約（アポ取得）
 *   G列: ステータス6  ／ 面談済み
 *   H列: ステータス7  ／ 土地探し中
 *   I列: ステータス8  ／ ローン事前審査待ち
 *   J列: ステータス9  ／ 中長期計画
 *   K列: ステータス10 ／ 内定
 *   L列: ステータス11 ／ 契約済み
 *   M列: ステータス12 ／ 失注
 *   N列: ステータス13 ／ 顧客対象外
 *   O列: 所見・備考（自由テキスト）
 *
 * 【デプロイ手順】
 *   1. 上メニュー「デプロイ」→「新しいデプロイ」
 *   2. 種類:「ウェブアプリ」を選択
 *   3. 実行ユーザー:「自分」
 *   4. アクセスできるユーザー:「全員」
 *   5. デプロイ → 表示されるURLをコピーしてダッシュボードに貼る
 */

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

  const payload = JSON.stringify({
    stores:    stores,
    updatedAt: new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' })
  });

  return ContentService
    .createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}
