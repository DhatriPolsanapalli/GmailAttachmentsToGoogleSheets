function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Gmail to sheets')
      .addItem('Config', 'updateLabelSheetName')
      .addItem('Fetch Data', 'saveXlsxDataToSheets')
      .addToUi();
}

function updateLabelSheetName() {
 var ui = SpreadsheetApp.getUi(); // Same variations.
  var documentProperties = PropertiesService.getDocumentProperties();

  openFirstPrompt(ui, documentProperties);
  
}

function openFirstPrompt(ui, documentProperties){
  var firstPrompt = ui.prompt(
      'Let\'s get Label from you GMail to scan for XLS attachment:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = firstPrompt.getSelectedButton();
  var text = firstPrompt.getResponseText();
  if (button == ui.Button.OK) {
    documentProperties.setProperty(ConfigVars.GMAIL_LABEL, text);
    openSecondPrompt(ui, documentProperties);
  } 
}

function openSecondPrompt(ui, documentProperties){
  var secondPrompt = ui.prompt(
      'Let\'s get Sheet Name where the data needs to be saved to:',
      
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = secondPrompt.getSelectedButton();
  var text = secondPrompt.getResponseText();
  if (button == ui.Button.OK) {
    documentProperties.setProperty(ConfigVars.SHEET_NAME, text);
  } 
}

