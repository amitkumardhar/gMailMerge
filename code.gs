function sendEmails(startRow,startCol,numRows,numCols) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0]; // First column
    var message = 'Dear '+row[1]+' '+row[2]+',\n\nPlease find the attachment.\n\nWarm Regards,\nJohn Doe\nHR Manager'; // Second column
    var files = DriveApp.getFilesByName(row[4]);
    var file;
    if (files.hasNext()) {
      file = files.next();
      if (files.hasNext()) {
        sheet.getRange(startRow + i, 6,1,1).setValue('Multiple_Files');
        return 0;
      }
    }
    else {
      sheet.getRange(startRow + i, 6,1,1).setValue('No_File');
      return 0;
    }
    var emailSent = row[5]; // Third column
    if (emailSent != 'EMAIL_SENT') { // Prevents sending duplicates
      var subject = 'Regarding your attachement';
      
      MailApp.sendEmail(emailAddress, subject, message, {
        name: 'JOHN DOE', 
        attachments: [file.getAs(MimeType.PDF)]
      });
      sheet.getRange(startRow+i, 6,1,1).setValue('EMAIL_SENT');
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}
function myFunction() {
  // Fetch the range of cells A2:B5
  sendEmails(2,1,1,5);
}
