function saveXlsxDataToSheets() {

  var documentProperties = PropertiesService.getDocumentProperties();
  const label = 'label:' + documentProperties.getProperty(ConfigVars.GMAIL_LABEL);
  var threads = GmailApp.search(label);
  var salesDataSheetName = documentProperties.getProperty(ConfigVars.SHEET_NAME);
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();

    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var attachments = message.getAttachments();

      for (var k = 0; k < attachments.length; k++) {
        var attachment = attachments[k];
        
        // Check if the attachment is an XLSX file
        if (attachment.getContentType() === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
          var file = DriveApp.createFile(attachment);
          var spreadsheet = SpreadsheetApp.openById('<Your Sheet ID goes here>');
          var sheet = spreadsheet.getSheetByName(salesDataSheetName);
          
          // Get the XLSX file as an Excel blob
          var xlsxBlob = file.getBlob();
          
          // Import XLSX data to Google Sheets
          var importedData = sheet.insertSheet().importXlsx(xlsxBlob);
          
          // Delete the temporary XLSX file from Google Drive
          DriveApp.getFileById(file.getId()).setTrashed(true);
          
          // Mark the email as read or perform any other desired actions
          message.markRead();
          
          // Break the loop to process one XLSX file per email
          break;
        }
      }
    }
  }
}
