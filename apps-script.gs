function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var p = e.parameter;
  sheet.appendRow([
    new Date(),
    p.name || '',
    p.phone || '',
    p.email || '',
    p.preferred_format || '',
    p.message || ''
  ]);
  return ContentService.createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}
