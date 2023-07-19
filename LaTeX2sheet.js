function latexToSheet() {
    ui = SpreadsheetApp.getUi();
    // var spreadsheet = SpreadsheetApp.openById('1gBEkOtHDZoUsF4RwV_mgrbAJ7rUlCzEPs3bRH51rK-o');
    // var range = spreadsheet.getSheetByName("table maker").getDataRange()

    var range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange();
    var table = range.getValues();
    var latexCode = []
    var record = true
    // console.log(table)
    for (var r = 0; r < table.length && record; r++) {
        for (var c = 0; c < table[r].length && record; c++) {
            if (table[r][c].includes('\\begin{tabular}')) {
                while (true){
                    r += 1;
                    cell_value = table[r][c]
                    if (cell_value.includes('\\bottomrule')) {
                        record = false
                        break;
                    }
                    cell_value = getContent(cell_value)
                    latexCode.push(cell_value);
                }
            }
        }
    }
    if (latexCode.length === 0){
        SpreadsheetApp.getUi().alert("V_Jul_19_2023: please provide me all the code from \\begin{tabular} to \\end{tabular}", SpreadsheetApp.getUi().ButtonSet.OK)
        return 0;
    }

    // convert latex code to the table.
    var sheet;
    var result = SpreadsheetApp.getUi().alert("V_Jul_19_2023: Clear the current sheet for the new table? Click \"No\" if needs a new sheet.", SpreadsheetApp.getUi().ButtonSet.YES_NO_CANCEL)
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
                // if cell object is a number
                if (cell.indexOf("%") !== -1) {
                    cell = cell.replace(/\\%/g, "%");
                }
                if (cell.indexOf("_") !== -1 & cell.indexOf("$") === -1) {
                    cell = cell.replace(/\\_/g, "_")
                }
                if (!isNaN(cell)) {

                    var form = "0";
                    if (cell.indexOf(".") !== -1) {
                        var decimals = cell.split(".")[1]
                        form += "." + repeat("0", decimals.length);
                    }
                    if (cell.indexOf("%") !== -1) {
                        form += "%";
                    }
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
