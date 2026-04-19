// Google Apps Script - Code.gs
// スプレッドシートのデータをJSON形式で返すWeb API
// スプレッドシート列構成: Unit(例:1-1) | 英語 | 日本語

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {};

  const gradeMap = {
    "中1": "j1",
    "中2": "j2",
    "中3": "j3"
  };

  Object.entries(gradeMap).forEach(([sheetName, gradeKey]) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const units = {};

    // 1行目はヘッダー（Unit, 英語, 日本語）なのでスキップ
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const unit = String(row[0]).trim(); // 例: "1-1", "1-2", "2-1"
      const en   = String(row[1]).trim();
      const ja   = String(row[2]).trim();
      if (!unit || !en || !ja) continue;

      if (!units[unit]) units[unit] = [];
      units[unit].push({ en, ja });
    }

    result[gradeKey] = units;
  });

  const output = ContentService.createTextOutput(JSON.stringify(result));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
