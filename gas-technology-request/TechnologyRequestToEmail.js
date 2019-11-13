
function sendEmailOfLastEditedRow() {

  var targetEmail = ''; // Enter the email address to receive message

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  var lastRow = sheet.getLastRow();
  var numRows = 1;   // Number of rows to process; we only need the most recently added row
  var cols = sheet.getLastColumn();

  // Fetch column header row
  var firstRow = sheet.getRange(1, 1, 1, cols);

  // Fetch the last row
  var lastRow = sheet.getRange(lastRow, 1, numRows, cols)

  // Fetch values for each row in the Range.
  var firstRowData = firstRow.getValues();
  var lastRowData = lastRow.getValues();

  // Fetch your row as an array
  var firstRowArray = firstRowData[0];
  var lastRowArray = lastRowData[0];

  var emailBody = '';

  for (i = 0; i < firstRowArray.length; i++) {

    // if the field has a value, add to emailBody
    if (lastRowArray[i]) {
      emailBody += firstRowArray[i] + ': ' + lastRowArray[i] + '\n\n';
    }

    Logger.log(firstRowArray[i] + ': ' + lastRowArray[i] + ' (' + i + ')');
  }
  var emailSubject = 'Technology Request: ';

  if (lastRowArray[2]) {
    emailSubject += lastRowArray[2];
  }
  if (lastRowArray[3]) {
    emailSubject += ' ' + lastRowArray[3] + '.';
  }
  if (lastRowArray[4]) {
    emailSubject += ' ' + lastRowArray[4]
  }
  if (lastRowArray[6]) {
    emailSubject += ' - ' + lastRowArray[6]
  }
  if (lastRowArray[8]) {
    emailSubject += ' - ' + lastRowArray[8]
  }

  // Send an email (change this to your email)
  GmailApp.sendEmail(targetEmail, emailSubject, emailBody);

  // Log contents for debugging
  Logger.log(emailBody);

}
