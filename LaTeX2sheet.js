function latexToSheet() {
    var result = SpreadsheetApp.getUi().alert("Only work for the LaTeX code generated via the spread-latex. Please provide me all the LaTeX code at least from \\begin{tabular} to \\end{tabular}. Click OK to continue.", SpreadsheetApp.getUi().ButtonSet.OK_CANCEL)
    if (result != 'OK') {
        return 0;
    }
    ui = SpreadsheetApp.getUi();
    // var spreadsheet = SpreadsheetApp.openById('1gBEkOtHDZoUsF4RwV_mgrbAJ7rUlCzEPs3bRH51rK-o');
    // var range = spreadsheet.getSheetByName("table maker").getDataRange()

    var range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange();
    var table = range.getValues();
    var latexCode = []
    // console.log(table)
    var start_r = -1
    var start_c = -1
    var end_c = -1
    for (var r = 0; r < table.length; r++) {
        for (var c = 0; c < table[r].length; c++) {
            if (String(table[r][c]).includes('\\begin{tabular}')) {
                start_r = r
                start_c = c
            }
            if (String(table[r][c]).includes('\\end{tabular}')) {
                end_c = c
            }
        }
    }
    if (start_r === -1 || start_c === -1) {
        SpreadsheetApp.getUi().alert("Please provide me all the LaTeX code from \\begin{tabular} to \\end{tabular}", SpreadsheetApp.getUi().ButtonSet.OK)
        return 0;
    }

    if (end_c === -1) {
        SpreadsheetApp.getUi().alert("\\end{tabular} is not found", SpreadsheetApp.getUi().ButtonSet.OK)
        return 0;
    }

    while (true) {
        start_r += 1;
        cell_value = table[start_r][start_c]
        if (cell_value.includes('\\bottomrule') || cell_value.includes('\\end{tabular}')) {
            break;
        }
        cell_value = getContent(cell_value)
        latexCode.push(cell_value);
    }


    // convert latex code to the table.
    var sheet;
    var result = SpreadsheetApp.getUi().alert("Clear the current sheet for the new table? Click \"No\" if needs a new sheet.", SpreadsheetApp.getUi().ButtonSet.YES_NO_CANCEL)
    if (result == "YES") {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        sheet.clear();
    } else if (result == "NO") {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    } else {
        return 0;
    }
    // console.log(latexCode)
    var rowIndex = 1;
    var columnIndex = 1;
    var str;
    for (let count = 0; count < latexCode.length; count++) {
        var row = latexCode[count]
        // if (count == 16){
        // console.log(count)
        // }
        row = row.split("&");
        for (var j = 0; j < row.length; j++) {
            //remove all the white space at the beginning and end of the string.
            var cell = trimWhiteSpace(row[j]);
            //remove all the automatically generated ",,"
            var cs = 1;
            var rs = 1;

            str = trimStart(cell, "\\multicolumn");
            if (str !== cell) {
                cell = str;
                idx_start = cell.indexOf("{")
                idx_end = cell.indexOf("}")
                cs = Number(cell.substring(idx_start + 1, idx_end));
                cell = cell.substring(idx_end + 5, cell.lastIndexOf("}"));
            }

            str = trimStart(cell, "\\multirow");
            if (str !== cell) {
                cell = str;
                idx_start = cell.indexOf("{")
                idx_end = cell.indexOf("}")
                rs = Number(cell.substring(idx_start + 1, idx_end));
                cell = cell.substring(idx_end + 5, cell.lastIndexOf("}"));
            }

            if (rs > 1 || cs > 1) {
                sheet.getRange(rowIndex, columnIndex, rs, cs).merge()
            }

            str = trimStart(cell, "\\cellcolor");
            if (str !== cell) {
                cell = str;
                var color = "#" + cell.substring(cell.indexOf("{") + 1, cell.indexOf("}"));
                sheet.getRange(rowIndex, columnIndex).setBackground(color);
                cell = trimStart(cell, "}");
            }

            str = trimStart(cell, "\\ul");
            if (str !== cell) {
                cell = str;
                sheet.getRange(rowIndex, columnIndex).setFontLine('underline');
                cell = cell.substring(cell.indexOf("{") + 1, cell.lastIndexOf("}"));
            }
            str = trimStart(cell, "\\textbf");
            if (str !== cell) {
                cell = str;
                sheet.getRange(rowIndex, columnIndex).setFontWeight("bold");
                cell = cell.substring(cell.indexOf("{") + 1, cell.lastIndexOf("}"));
            }
            if (cell !== "") {
                var percentage = ""
                var currency = ""
                // if cell object is a number
                if (cell.indexOf("%") !== -1) {
                    var a = cell.substring(cell.length - 2, cell.length)
                    var b = cell.substring(0, cell.length - 2)
                    if (a === "\\%" && !isNaN(b)) {
                        percentage = "%"
                        cell = b
                    } else {
                        cell = cell.replace(/\\%/g, "%");
                    }
                }
                if (cell.indexOf("#") !== -1) {
                    cell = cell.replace(/\\#/g, "#")
                }
                if (cell.indexOf("_") !== -1 & cell.indexOf("$") === -1) {
                    cell = cell.replace(/\\_/g, "_")
                }
                if (cell.indexOf("\\$") !== -1) {
                    var a = cell.substring(0, 2)
                    var b = cell.substring(2, cell.length)
                    if (a == "\\$" && !isNaN(b)) {
                        cell = b
                        currency = "$"
                    } else {
                        cell = cell.replace(/\\$/g, "$");
                    }
                }
                if (!isNaN(cell)) {
                    var form = "0";
                    if (cell.indexOf(".") !== -1) {
                        var decimals = cell.split(".")[1].split("e")[0]
                        form += "." + repeat("0", decimals.length);
                    }
                    if (cell.indexOf('e') !== -1) {
                        form += "E+00"
                    }
                    form = currency + form + percentage
                    sheet.getRange(rowIndex, columnIndex).setNumberFormat(form);
                    sheet.getRange(rowIndex, columnIndex).setHorizontalAlignment("right");
                } else {
                    sheet.getRange(rowIndex, columnIndex).setShowHyperlink(false);
                    //                sheet.getRange(rowIndex, columnIndex).setValue(cell);
                }
                sheet.getRange(rowIndex, columnIndex).setValue(cell);
            }
            columnIndex += cs;
        }
        rowIndex++;
        columnIndex = 1;
    }
    sheet.autoResizeColumns(1, sheet.getDataRange().getNumColumns());

    Browser.msgBox('The table is loaded successfully!', Browser.Buttons.OK)
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

function getContent(str) {
    var ind = str.lastIndexOf('\\\\');
    if (ind != -1) {
        str = str.substring(0, ind - 1)
    }
    str = trimWhiteSpace(str)
    return str
}

function trimWhiteSpace(str) {
    do {
        var original = str;
        str = str.replace(/(^\s)|(\s$)/g, "");
    } while (original !== str);
    return str;
}
