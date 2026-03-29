/**
 * ロジカルシンキング研修クイズ - 回答記録用 Google Apps Script
 */

const SHEET_NAME = 'ロジカルシンキング';

// GET: 認証確認画面
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

    const sections = data.sections || {};
    const levels = data.levels || {};

    const row = [
      new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' }),
      data.testName || 'ロジカルシンキング研修クイズ',
      data.name || '',
      data.correct,
      data.total,
      data.percentage,
      sections['分類する'] || '0/0',
      sections['分解する'] || '0/0',
      sections['水準を整理する'] || '0/0',
      levels['Level1'] || '0/0',
      levels['Level2'] || '0/0',
      levels['Level3'] || '0/0',
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
