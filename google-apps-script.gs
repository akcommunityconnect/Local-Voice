/*
  SETUP
  1. Create a private Google Sheet and open Extensions > Apps Script.
  2. Replace the default code with this file's contents and save.
  3. Click Deploy > New deployment > Web app.
  4. Set Execute as: Me; Who has access: Anyone.
  5. Authorize, deploy, and copy the Web app URL.
  6. Paste that URL into index.html where indicated.

  The Sheet itself remains private in your Google account. Website visitors
  can submit a report but cannot view the Sheet or other submissions.
*/
function saveReport(messageValue) {
  const message = String(messageValue || '').trim();
  if (!message || message.length > 1800) {
    return { success: false, error: 'Please enter a report of up to 1,800 characters.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Sr. No.', 'Submitted at', 'Message']);
    } else if (sheet.getRange(1, 1).getDisplayValue() !== 'Sr. No.') {
      sheet.insertColumnBefore(1);
      sheet.getRange(1, 1).setValue('Sr. No.');
      const existingReports = sheet.getLastRow() - 1;
      if (existingReports > 0) {
        sheet.getRange(2, 1, existingReports, 1)
          .setValues(Array.from({ length: existingReports }, (_, index) => [index + 1]));
      }
    }

    const reportId = sheet.getLastRow();
    sheet.appendRow([reportId, new Date(), message]);
    return { success: true, reportId: reportId };
  } finally {
    lock.releaseLock();
  }
}

function doPost(event) {
  try {
    const data = JSON.parse((event && event.postData && event.postData.contents) || '{}');
    const result = saveReport(data.message);
    return ContentService.createTextOutput(result.success ? 'Saved: ' + result.reportId : result.error);
  } catch (error) {
    return ContentService.createTextOutput('The report could not be saved.');
  }
}

function doGet(event) {
  const parameters = (event && event.parameter) || {};
  if (parameters.action === 'submit' && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(parameters.callback || '')) {
    let result;
    try {
      result = saveReport(parameters.message);
    } catch (error) {
      result = { success: false, error: 'The report could not be saved. Please try again.' };
    }
    return ContentService.createTextOutput(parameters.callback + '(' + JSON.stringify(result) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput('Local Voice report endpoint is running.');
}
