// Xcout — Google Sheet sign-up receiver (Google Apps Script Web App)
// Paste this into a Sheet-bound Apps Script, deploy as Web App, and use the /exec URL.
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Signups') || ss.insertSheet('Signups');

    var data = {};
    if (e && e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch (err) { data = e.parameter || {}; }
    } else if (e) { data = e.parameter || {}; }

    var headers = ['Timestamp','Role','Name','Phone','Email','Location',
      'Position','Age','Height','Team_or_School','Video_Link',
      'Coach_Team','Coach_Level','Coach_Years',
      'Scout_Org','Scout_Title','Scout_Regions',
      'Org_Name','Org_Type','Org_Link','Message'];

    if (sheet.getLastRow() === 0) sheet.appendRow(headers);

    var row = headers.map(function (h) {
      return h === 'Timestamp' ? new Date() : (data[h] || '');
    });
    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'error', error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
