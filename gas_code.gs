/**
 * toBビジネス用語テスト - 回答記録用 Google Apps Script
 *
 * 【セットアップ手順】
 * 1. スプレッドシートを開く → 「拡張機能」→「Apps Script」
 * 2. このコードを貼り付け
 * 3.「デプロイ」→「デプロイを管理」→ 鉛筆アイコン → バージョン「新しいバージョン」→「デプロイ」
 */

const SHEET_NAME = 'シート1';

// GET: Google認証確認画面を表示
function doGet(e) {
  const user = Session.getActiveUser().getEmail();

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Google認証完了</title>
      <style>
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Hiragino Sans', sans-serif;
          text-align: center;
          padding: 40px 20px;
          background: #f7fafc;
          color: #2d3748;
        }
        .card {
          background: #fff;
          border-radius: 12px;
          padding: 2rem;
          max-width: 400px;
          margin: 0 auto;
          box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }
        .icon { font-size: 3rem; margin-bottom: 0.5rem; }
        h2 { font-size: 1.3rem; margin-bottom: 0.5rem; }
        .email {
          background: #ebf8ff;
          padding: 0.5rem 1rem;
          border-radius: 8px;
          font-weight: 700;
          color: #2b6cb0;
          margin: 1rem 0;
          font-size: 1rem;
        }
        .close-msg {
          display: inline-block;
          background: #fff5f5;
          color: #e53e3e;
          border: 2px solid #e53e3e;
          padding: 0.8rem 1.5rem;
          border-radius: 8px;
          font-size: 1rem;
          font-weight: 700;
          margin-top: 1rem;
        }
        .note { font-size: 0.85rem; color: #718096; margin-top: 0.8rem; }
      </style>
    </head>
    <body>
      <div class="card">
        <div class="icon">&#9989;</div>
        <h2>認証完了</h2>
        <div class="email">${user}</div>
        <p>上記のGoogleアカウントで認証されました。</p>
        <div class="close-msg">このタブを閉じて、テスト画面に戻ってください</div>
        <p class="note">※ テスト画面の「テストを開始する」ボタンをクリックしてください</p>
      </div>
    </body>
    </html>
  `;

  return HtmlService.createHtmlOutput(html)
    .setTitle('Google認証完了');
}

// POST: テスト結果を記録（メールアドレスはGAS側で自動取得）
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    // Googleアカウントのメールアドレスを自動取得
    const email = Session.getActiveUser().getEmail() || data.email || '不明';

    const row = [
      new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' }),
      data.testName || '第1弾',
      email,
      data.name || '',
      data.correct,
      data.total,
      data.percentage,
      data.categories['企業・組織形態'] || '0/0',
      data.categories['M&A・経営戦略'] || '0/0',
      data.categories['財務・会計'] || '0/0',
      data.categories['営業・マーケティング'] || '0/0',
      data.categories['契約・法務'] || '0/0',
      data.categories['IT・DX関連'] || '0/0',
      data.categories['ベンチャー・スタートアップ'] || '0/0',
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, email: email }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
