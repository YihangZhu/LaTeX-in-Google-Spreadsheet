function onInstall(e) {
    onOpen(e);
}

function onOpen(e) {
    SpreadsheetApp.getUi().createAddonMenu()
        .addItem('Convert sheet->LaTeX', 'sheetToLatex')
        .addItem('Convert LaTeX->sheet', 'latexToSheet')
        .addItem("Donate", 'donate')
        .addToUi()
}

function donate() {
    var ui = HtmlService.createHtmlOutputFromFile("donation")
    SpreadsheetApp.getUi().showModelessDialog(ui, "Donation for the domain cost")
}