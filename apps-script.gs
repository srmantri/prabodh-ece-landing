var NOTIFY_EMAIL = "prabodh.inst@gmail.com";

function ordinal(n) {
  var s = ["th", "st", "nd", "rd"];
  var v = n % 100;
  return n + (s[(v - 20) % 10] || s[v] || s[0]);
}

function formatTimestamp(date) {
  var tz = Session.getScriptTimeZone();
  var day = ordinal(date.getDate());
  var monthYear = Utilities.formatDate(date, tz, "MMMM yyyy");
  var time = Utilities.formatDate(date, tz, "h:mm a");
  return day + " " + monthYear + ", " + time;
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var p = e.parameter;
  sheet.appendRow([
    formatTimestamp(new Date()),
    p.name || '',
    p.phone || '',
    p.email || '',
    p.preferred_format || '',
    p.preferred_date || '',
    p.preferred_time_slot || '',
    p.message || ''
  ]);

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: "New ECE Course Enquiry: " + (p.name || "Unknown"),
    body:
      "New enquiry from the ECE landing page:\n\n" +
      "Name: " + (p.name || '') + "\n" +
      "Phone: " + (p.phone || '') + "\n" +
      "Email: " + (p.email || '') + "\n" +
      "Preferred Format: " + (p.preferred_format || '') + "\n" +
      "Preferred Date: " + (p.preferred_date || '') + "\n" +
      "Preferred Time Slot: " + (p.preferred_time_slot || '') + "\n" +
      "Message: " + (p.message || '')
  });

  return ContentService.createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}
