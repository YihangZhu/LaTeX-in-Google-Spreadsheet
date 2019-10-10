function onInstall(e) {
  onOpen(e);
}

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {  
  SpreadsheetApp.getUi().createAddonMenu()
  .addItem('Convert sheet->LaTeX', 'sheetToLatex')
  .addItem('Convert LaTeX->sheet', 'latexToSheet')
  .addSeparator()
  .addItem('Help', 'help')
  .addToUi()
}

function help() {
  var html = HtmlService.createHtmlOutputFromFile("Help").setHeight(88);
  SpreadsheetApp.getUi().showModalDialog(html, "help")
}






