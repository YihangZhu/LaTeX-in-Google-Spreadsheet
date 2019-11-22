function latexToSheet() {
    ui = SpreadsheetApp.getUi();
//  var spreadsheet = SpreadsheetApp.openById('11VL3bqvCkUJb-v_zRbscAiI--Y3b4YYdyGopZUv05k0');
//  var range = spreadsheet.getSheetByName("table maker").getDataRange()

    var range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange();
    var latexCode = range.getValues().join(" ");

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

    var rowIndex = 1;
    var columnIndex = 1;
    var str;
    while (true) {
        var ind = latexCode.indexOf("\\\\");
        if (ind === -1) {
            break;
        }
        var row = latexCode.substr(0, ind);
        latexCode = latexCode.substring(ind + 2);

        // trim the code for the borders
        while (true) {
            str = trimStart(row, "\\cmidrule");
            if (str !== row) {
                row = str;
                row = trimStart(row, "}");
            } else {
                break;
            }
        }

        row = trimStart(row, "\\midrule");
        row = trimStart(row, "\\toprule");

        if (row.indexOf("\\bottomrule") !== -1) {
            continue;
        }
        row = row.split("&");
        for (var j = 0; j < row.length; j++) {
            var cell = trim(row[j]);
            var cs = 1;
            var rs = 1;

            str = trimStart(cell, "\\multicolumn");
            if (str !== cell) {
                cell = str;
                cs = Number(cell.substring(cell.indexOf("{") + 1, cell.indexOf("}")));
                cell = cell.substring(cell.indexOf("{") + 7, cell.lastIndexOf("}"));
            }

            str = trimStart(cell, "\\multirow");
            if (str !== cell) {
                cell = str;
                rs = Number(cell.substring(cell.indexOf("{") + 1, cell.indexOf("}")));
                cell = cell.substring(cell.indexOf("{") + 7, cell.lastIndexOf("}"));
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
                    cell = cell.replace("\\%", "%");
                }
                if (!isNaN(cell)) {

                    var form = "0";
                    if (cell.indexOf(".") !== -1) {
                        form += "." + repeat("0", cell.split(".")[1].length);
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

