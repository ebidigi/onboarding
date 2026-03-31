/**
 * 研修テスト共通 - 回答記録用 Google Apps Script
 *
 * スプレッドシート: https://docs.google.com/spreadsheets/d/1A6uYGKB7lWJz498WvQPCoL8V9QNnpK0F3vAsW3koK9U/edit
 *
 * 対応テスト:
 *   - スタンス研修 理解度     → シート「スタンス研修理解度テスト」
 *   - ビジネスマナー・ルール  → シート「マナー研修理解度テスト」
 *   - DigiMan・ビジネスモデル理解 → シート「DigiMan・ビジネスモデル理解テスト」
 *   - その他すべてのテスト    → シート「研修テスト」（マスター）
 *
 * testName ごとに専用シートへ記録し、全テストを「研修テスト」シートにも集約します。
 */

const MASTER_SHEET_NAME = '研修テスト';
const TARGET_SS_ID = '1A6uYGKB7lWJz498WvQPCoL8V9QNnpK0F3vAsW3koK9U';

// testName → 専用シート名のマッピング
const TEST_SHEET_MAP = {
  'スタンス研修 理解度':        'スタンス研修理解度テスト',
  'ビジネスマナー・ルール':     'マナー研修理解度テスト',
  'DigiMan・ビジネスモデル理解': 'DigiMan・ビジネスモデル理解テスト',
};

const HEADER_ROW = ['受験日時', 'テスト名', 'メールアドレス', '氏名', '正答数', '問題数', '正答率(%)', 'カテゴリ別詳細'];

/**
 * 指定名のシートを取得する。存在しない場合はヘッダー付きで新規作成する。
 */
function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(HEADER_ROW);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

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

// POST: テスト結果を記録
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(TARGET_SS_ID);

    const email = Session.getActiveUser().getEmail() || data.email || '不明';
    const testName = data.testName || '';

    // カテゴリ別結果を動的に組み立て
    const categoryDetails = [];
    if (data.categories) {
      for (const [categoryName, score] of Object.entries(data.categories)) {
        categoryDetails.push(categoryName + ': ' + score);
      }
    }
    const categoryString = categoryDetails.join(' / ');

    const row = [
      new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' }),
      testName,
      email,
      data.name || '',
      data.correct,
      data.total,
      data.percentage,
      categoryString,
    ];

    // 1. マスターシート（全テスト集約）に記録
    const masterSheet = getOrCreateSheet(ss, MASTER_SHEET_NAME);
    masterSheet.appendRow(row);

    // 2. テスト名に対応する専用シートに記録
    const dedicatedSheetName = TEST_SHEET_MAP[testName];
    if (dedicatedSheetName) {
      const dedicatedSheet = getOrCreateSheet(ss, dedicatedSheetName);
      dedicatedSheet.appendRow(row);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, email: email }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
