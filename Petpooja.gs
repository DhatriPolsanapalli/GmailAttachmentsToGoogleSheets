function saveXlsxDataToSheets() {

  var documentProperties = PropertiesService.getDocumentProperties();
  const label = 'label:' + documentProperties.getProperty(ConfigVars.GMAIL_LABEL);
  var threads = GmailApp.search(label);
  Logger.log(threads)
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
          var file = Drive.Files.insert({ title : 'temp_converted'}, attachment, {
        convert: true
      });
          var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
          var sheet = spreadsheet.getSheetByName(salesDataSheetName);
          Logger.log(sheet)
          const sourceValues = getExcelData(file.id)
          // Get the XLSX file as an Excel blob
          
          sheet.getRange(sheet.getLastRow()+1, 1, sourceValues.length, sourceValues[0].length).setValues(sourceValues);
          
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


function getExcelData(tempID) { 
  const source = SpreadsheetApp.openById(tempID);
  //The sheetname of the excel where you want the data from
  const sourceSheet = source.getSheets()[0];
  //The range you want the data from

  var lastRow = sourceSheet.getLastRow();
  var lastColumn = sourceSheet.getLastColumn();
  return sourceSheet.getRange(lastRow, lastColumn).getValues();
}