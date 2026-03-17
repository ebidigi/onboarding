/**
 * 学習速度テスト - 回答記録用 Google Apps Script
 *
 * 【セットアップ手順】
 * 1. 新しいスプレッドシートを作成
 * 2. シート名を「学習速度テスト」に変更
 * 3. 1行目にヘッダーを入力（下記参照）
 * 4.「拡張機能」→「Apps Script」→ このコードを貼り付け
 * 5.「デプロイ」→「新しいデプロイ」→ ウェブアプリ → アクセスできるユーザー「全員」→ デプロイ
 * 6. 生成されたURLをtest.htmlのGAS_URLに設定
 *
 * 【ヘッダー（1行目に設定）】
 * A: 受験日時
 * B: テスト名
 * C: 氏名
 * D: 正解数
 * E: 総問題数
 * F: 正答率(%)
 * G: Level1(記憶)
 * H: Level2(理解)
 * I: Level3(応用)
 * J: 企業・組織形態
 * K: M&A・経営戦略
 * L: 財務・会計
 * M: 営業・マーケティング
 * N: 契約・法務
 * O: IT基礎知識
 * P: ベンチャー・スタートアップ
 * Q: 平均回答時間(秒)
 * R: 未回答数
 * S: タブ離脱回数
 * T: 総所要時間(秒)
 * U: 開始時刻
 * V: 終了時刻
 */

const SHEET_NAME = '学習速度テスト';

// GET: 認証確認画面（既存テストと同じ仕組み）
function doGet(e) {
  const user = Session.getActiveUser().getEmail();
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>認証完了</title>
      <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Hiragino Sans', sans-serif; text-align: center; padding: 40px 20px; background: #f7fafc; }
        .card { background: #fff; border-radius: 12px; padding: 2rem; max-width: 400px; margin: 0 auto; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
        .icon { font-size: 3rem; margin-bottom: 0.5rem; }
        .email { background: #ebf8ff; padding: 0.5rem 1rem; border-radius: 8px; font-weight: 700; color: #2b6cb0; margin: 1rem 0; }
      </style>
    </head>
    <body>
      <div class="card">
        <div class="icon">&#9989;</div>
        <h2>認証完了</h2>
        <div class="email">${user}</div>
        <p>このタブを閉じてテスト画面に戻ってください。</p>
      </div>
    </body>
    </html>
  `;
  return HtmlService.createHtmlOutput(html).setTitle('認証完了');
}

// POST: テスト結果を記録
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error(`シート "${SHEET_NAME}" が見つかりません。シート名を確認してください。`);
    }

    const categories = data.categories || {};
    const levels = data.levels || {};

    const row = [
      new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' }),
      data.testName || '学習速度テスト',
      data.name || '',
      data.correct,
      data.total,
      data.percentage,
      levels['Level1'] || '0/0',
      levels['Level2'] || '0/0',
      levels['Level3'] || '0/0',
      categories['企業・組織形態'] || '0/0',
      categories['M&A・経営戦略'] || '0/0',
      categories['財務・会計'] || '0/0',
      categories['営業・マーケティング'] || '0/0',
      categories['契約・法務'] || '0/0',
      categories['IT基礎知識'] || '0/0',
      categories['ベンチャー・スタートアップ'] || '0/0',
      data.avgTime || '',
      data.unanswered || 0,
      data.tabSwitchCount || 0,
      data.totalTimeSec || '',
      data.startTime || '',
      data.endTime || '',
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
