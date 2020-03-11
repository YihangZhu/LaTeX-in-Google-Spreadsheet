function onInstall(e) {
    onOpen(e);
}

function onOpen(e) {
    SpreadsheetApp.getUi().createAddonMenu()
        .addItem('Convert sheet->LaTeX', 'sheetToLatex')
        .addItem('Convert LaTeX->sheet', 'latexToSheet')
        .addItem("666", 'donate')
        .addToUi()
}

function donate(){
  var ui = HtmlService.createHtmlOutputFromFile("donation") 
  SpreadsheetApp.getUi().showModelessDialog(ui,"Donate 0.66 for the domain cost");
}

/**
 * trim " ", "#" and "," in str
 * @returns the trimmed str
 */
function trim(str) {
    do {
        var original = str;
        str = str.replace(/(,+)|(\s+)|(#+)/g, '');
    } while (original !== str);
    return str;
}

/**
 * repeat str for num times.
 * num larger than 1.
 **/
function repeat(str, num) {
    var newStr = "";
    for (var i = 0; i < num; i++) {
        newStr += str;
    }
    return newStr;
}

/**
 *
 * @param str original string
 * @param substr
 * @returns true if substr is trimmed from string, false if str does not include substr.
 */
function trimStart(str, substr) {
    var ind = str.indexOf(substr);
    if (ind !== -1) {
        str = str.substring(ind + substr.length);
    }
    return str;
}