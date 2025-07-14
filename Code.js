function doGet() {
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

  return ContentService.createTextOutput(JSON.stringify(jsonArray))
    .setMimeType(ContentService.MimeType.JSON);
}
