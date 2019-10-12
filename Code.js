function onInstall(e) {
  onOpen(e);
}

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen(e) {  
  SpreadsheetApp.getUi().createAddonMenu()
  .addItem('Convert sheet->LaTeX', 'sheetToLatex')
  .addItem('Convert LaTeX->sheet', 'latexToSheet')
  .addSeparator()
  .addItem('Support', 'support')
  .addToUi()
}

function support() {
  var html = HtmlService.createHtmlOutputFromFile("Support").setHeight(88);
  SpreadsheetApp.getUi().showModalDialog(html, "Support")
}






