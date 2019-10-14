function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {  
  SpreadsheetApp.getUi().createAddonMenu()
  .addItem('Convert sheet->LaTeX', 'sheetToLatex')
  .addItem('Convert LaTeX->sheet', 'latexToSheet')
  .addToUi()
}
