var ui = null;
var html = null;
var gray = false;
var check = false;

function sheetToLatex() {
    ui = SpreadsheetApp.getUi();
    var lineBreak = "<br>";
    var title = "Generated by Spread-LaTeX";
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var table = sheet.getActiveRange();

    //get number of rows for column names for determining midrule border
    var result = ui.prompt("How many rows for the table header:",
        ui.ButtonSet.OK_CANCEL);
    if (result.getSelectedButton() != 'OK') {
        return 0;
    }

    result = result.getResponseText();
    var columnNameRows = -1;
    if (result !== "" && result !== "0") {
        columnNameRows = Number(result);
    }

    html = HtmlService.createHtmlOutput("Generating ...... ").setWidth(1600);
    ui.showModalDialog(html, "LaTeX code:");


    //start loading the selected table

    var objects = table.getValues();
    var formats = table.getNumberFormats();
    var fontWeights = table.getFontWeights();
    var underlines = table.getFontLines();
    var backgrounds = table.getBackgrounds();
    var rows = table.getNumRows();
    var cols = table.getNumColumns();

    //alignment for each column
    var alignment = "l";
    for (var c = 0; c < cols; c++) {
        alignment += "r"
    }

    var content = "";

    var cmidrules = [];
    for (var r = 0; r < rows; r++) {
        for (c = 0; c < cols; c++) {
            var cell = table.getCell(r + 1, c + 1);
            var isPartOfMerge = cell.isPartOfMerge();
            var mergedRange;
            var cellCol, cellRow;
            var rangeLastCol, rangeFirstRow;
            var rangeNumRows, rangeNumCols;
            if (isPartOfMerge) {
                cellCol = cell.getColumn();
                cellRow = cell.getRow();
                mergedRange = cell.getMergedRanges()[0];

                rangeLastCol = mergedRange.getLastColumn();
                rangeFirstRow = mergedRange.getRow();
                rangeNumRows = mergedRange.getNumRows();
                rangeNumCols = mergedRange.getNumColumns();

                // if the cell is the top left of the merged range, read the cell
                if (cellRow === rangeFirstRow && cellCol === mergedRange.getColumn()) {
                    var cellValue = readCell(objects[r][c], formats[r][c], fontWeights[r][c], underlines[r][c], backgrounds[r][c],
                        true, rangeNumRows, rangeNumCols);
                    if (cellValue == "exit(cancel)") {
                        return 0;
                    }
                    content += cellValue;
                }
            } else {
                var cellValue = readCell(objects[r][c], formats[r][c], fontWeights[r][c], underlines[r][c], backgrounds[r][c],
                    false, 1, 1);
                if (cellValue == "exit(cancel)") {
                    return 0;
                }
                content += cellValue;
            }

            // record a cmidrule border, if the cell or the merged range is not blank
            if (r + 1 < columnNameRows) {
                if (isPartOfMerge) {
                    if (cellRow == mergedRange.getLastRow()) {
                        //the cell is not in the first row of a multirow merged range
                        cmidrules.push(c + 1)
                    }
                } else if (not_blank(cell)) {
                    cmidrules.push(c + 1);
                }
            }

            // add & if the cell is the last column or not the first row of the merged range, or just a single cell
            if (c < cols - 1) {
                if (!isPartOfMerge || (cellCol === rangeLastCol || cellRow !== rangeFirstRow)) {
                    content += "\t&";
                }
            }
        }
        content += "\t\\\\";

        //add cmidrule borders for the row
        var begin = cmidrules[0];
        var end = begin;
        var len = cmidrules.length;
        for (var i = 1; i < len; i++) {
            if (cmidrules[i] > end + 1 || i === len - 1) {
                if (i === len - 1) {
                    end = cmidrules[i];
                }
                content += "\\cmidrule{" + begin + "-" + end + "}";
                begin = cmidrules[i];
                end = begin;
            } else {
                end = cmidrules[i];
            }
        }

        cmidrules = [];

        if (r + 1 === columnNameRows) {
            content += "\\midrule";
        }
        content += lineBreak;
    }

    var str = "%Please add the following packages if necessary:" + lineBreak
        + "%\\usepackage{booktabs, multirow} % for borders and merged ranges" + lineBreak
        + "%\\usepackage{soul}% for underlines" + lineBreak
        + "%\\usepackage[table]{xcolor} % for cell colors" + lineBreak
        + "%\\usepackage{changepage,threeparttable} % for wide tables" + lineBreak
        + "%If the table is too wide, replace \\begin{table}[!htp]...\\end{table} with" + lineBreak
        + "%\\begin{adjustwidth}{-2.5 cm}{-2.5 cm}\\centering\\begin{threeparttable}[!htb]...\\end{threeparttable}\\end{adjustwidth}" + lineBreak
        + "\\begin{table}[!htp]\\centering" + lineBreak
        + "\\caption{" + title + "}\\label{tab:  }" + lineBreak
        + "\\scriptsize" + lineBreak
        + "\\begin{tabular}{" + alignment + "}\\toprule" + lineBreak
        + content
        + "\\bottomrule" + lineBreak
        + "\\end{tabular}" + lineBreak
        + "\\end{table}";

    html.setContent(str);
    ui.showModalDialog(html, "LaTeX code: (select all + copy)");
}

function readCell(object, format, fontWeight, underline, backgraound, isMergedRange, rows, cols) {
    var type = typeof (object);
    if (type === "number" && format !== "") {
        if (format.indexOf("E") !== -1) {
            var saveN = format.split("E")[0].split(".")
            var n = 0;
            if (saveN.length > 1) {
                n = saveN[1].length;
            }
            object = object.toExponential(n);
        } else {
            var trimedFormat = format.replace("%", "");
            // javascript regnex: replace all "#" remove all "," in the string. modifier g is to replace all the matches, + is find all the match
            trimedFormat = trimedFormat.replace(/#/g, "0")
            trimedFormat = trimedFormat.replace(/,+/, "");
            var decimalPlaces = 0;
            var decimals = trimedFormat.split(".")
            if (decimals.length > 1) {
                decimalPlaces = decimals[1].length;
            }

            // if the data is a percentage value
            var percentage = ""
            if (format.indexOf("%") !== -1) {
                object = object * 100;
                percentage = "\\%"
            }

            // 2.00   2.02
            if (decimalPlaces < 15) {
                object += Number.EPSILON
                object = object.toFixed(decimalPlaces)
            }

            object = String(object) + percentage

            if (format.indexOf("#,##") !== -1) {
                var str = object.split(".");
                object = str[0].replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
                if (str.length > 1) {
                    object = object + "." + str[1];
                }
            }
        }
    } else {
        object = String(object);
        var ind1 = object.indexOf("_")
        var ind2 = object.indexOf("$")
        if ((ind1 !== -1) && (ind2 === -1)) {
            object = object.replace(/_/g, "\\_")
        }
    }
//    if (object.indexOf('&') === -1) {
//      object = object.replace("%", "\\%");
//    }

    if (fontWeight === 'bold') {
        object = "\\textbf{" + object + "}";
    }

    if (underline === 'underline') {
        object = "\\ul{" + object + "}";
    }
    // if the cell is highlighted
    if (backgraound !== '#ffffff') {
        if (!check) {
            var response = ui.alert("Replace all the highlights with gray color?", ui.ButtonSet.YES_NO_CANCEL);
            if (response == "CANCEL") {
                return "exit(cancel)";
            } else {
                gray = response == "YES";
            }
            check = true;
            ui.showModalDialog(html, "LaTeX code:");
        }
        if (gray) {
            object = "\\cellcolor[HTML]{A8A8A8}" + object;
        } else {
            var color = String(backgraound).replace("#", "");
            object = "\\cellcolor[HTML]{" + color + "}" + object;
        }
    }

    if (isMergedRange) {
        if (rows > 1) {
            object = "\\multirow{" + rows + "}{*}{" + object + "}";
        }
        if (cols > 1) {
            object = "\\multicolumn{" + cols + "}{c}{" + object + "}";
        }
    }
    return object;
}

function not_blank(cell) {
    var str = String(cell.getValue())
    str = str.replace(/\s+/, "")
    return str.length > 0;
}