/**
 * Google スプレッドシートのシートからデータを取得し、
 * 各行をJSONオブジェクトに変換し、最終的に
 * HTMLダイアログを表示して、JSONファイルをダウンロードできるようにします。
 */
function run() {
  // アクティブなスプレッドシートのアクティブなシートからデータを取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const jsonArray = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const jsonObject = {};
    for (let j = 0; j < headers.length; j++) {
      jsonObject[headers[j]] = row[j];
    }
    jsonArray.push(jsonObject);
  }
  const jsonText = JSON.stringify(JSON.stringify(jsonArray, null, 2));

  // スプレッドシート名、シート名、日時を取得
  const spreadsheetName = spreadsheet.getName().replace(/\s+/g, '_');
  const sheetName = sheet.getName().replace(/\s+/g, '_');
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');

  const fileName = `${spreadsheetName}_${sheetName}_${timestamp}.json`;

  const html = HtmlService.createHtmlOutput(`
    <html>
      <body>
        <h3>✅ JSONファイルをダウンロード</h3>
        <button onclick="download()">📥 ダウンロード</button>
        <script>
          function download() {
            const blob = new Blob([${jsonText}], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = '${fileName}';
            a.click();
            URL.revokeObjectURL(url);
          }
        </script>
      </body>
    </html>
  `).setWidth(400).setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(html, 'ダウンロード');
}


/**
 * Google スプレッドシートのメニューに「📦 ConvertSheetToJson」というカスタムメニューを追加する関数
 * @returns {void}
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 ConvertSheetToJson')
    .addItem('JSONに変換', 'run')
    .addToUi();
}
