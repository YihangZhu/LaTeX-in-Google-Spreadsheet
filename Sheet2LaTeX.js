
function sheetToLatex(){ 
  var ui = SpreadsheetApp.getUi();
  var lineBreak = "<br>";
  var title = "Generated by Spread-LaTeX" ;
  
  //get number of rows for column names for determining midrule border
  var result = ui.prompt("Number of the rows for the column names:",
      ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() != 'OK'){
    return 0;
  }
  
  result = result.getResponseText();
  var columnNameRows = -1;
  if (result !== "" && result !== "0"){
    columnNameRows = Number(result);
  }
  
  var html = HtmlService.createHtmlOutput("Generating ...... ").setWidth(1600);
  ui.showModalDialog(html,"LaTeX code:");
  
  //start loading the selected table
  var table = SpreadsheetApp.getActiveSheet().getActiveRange();
  var objects = table.getValues();
  var formats = table.getNumberFormats();
  var fontWeights = table.getFontWeights();
  var underlines = table.getFontLines();
  var backgrounds = table.getBackgrounds();
  var rows = table.getNumRows();
  var cols = table.getNumColumns();

  //alignment for each column
  var alignment = "l";
  for (var c=0; c<cols; c++){
     alignment += "r"
  }
  
  var content = "";

  var cmidrules = [];
  for (var r=0; r<rows; r++){  
    for (c=0; c<cols; c++){
      var cell = table.getCell(r+1, c+1);
      var isPartOfMerge = cell.isPartOfMerge();
      var mergedRange;
      var cellCol, cellRow;
      var rangeLastCol, rangeFirstRow;
      var rangeNumRows, rangeNumCols;
      if (isPartOfMerge){
        cellCol = cell.getColumn();
        cellRow = cell.getRow();
        mergedRange = cell.getMergedRanges()[0];
 
        rangeLastCol = mergedRange.getLastColumn();
        rangeFirstRow = mergedRange.getRow();
        rangeNumRows = mergedRange.getNumRows();
        rangeNumCols = mergedRange.getNumColumns();
        
        // if the cell is the top left of the merged range, read the cell
        if (cellRow === rangeFirstRow && cellCol === mergedRange.getColumn()){
          content += readCell(objects[r][c], formats[r][c], fontWeights[r][c], underlines[r][c], backgrounds[r][c],
              true, rangeNumRows, rangeNumCols);
        }
      }else {
        content += readCell(objects[r][c], formats[r][c], fontWeights[r][c], underlines[r][c], backgrounds[r][c],
            false, 1, 1);
      }
     
      // record a cmidrule border, if the cell or the merged range is not blank
      if (r+1 < columnNameRows && ((!cell.isBlank())||(isPartOfMerge && !mergedRange.isBlank()))){
        if (isPartOfMerge && rangeNumRows > 1){
          //the cell is not in the first row of a multirow merged range
          if (cellRow > rangeFirstRow){
            cmidrules.push(c+1);
          }
        }else{
          cmidrules.push(c+1);
        }
      }
      
      // add & if the cell is the last column or not the first row of the merged range, or just a single cell
      if (c < cols-1) {
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
    for (var i=1; i<len; i++){
      if (cmidrules[i] > end+1||i === len-1){
        if(i===len-1){
          end = cmidrules[i];
        }
        content += "\\cmidrule{"+begin+"-"+end+"}";
        begin = cmidrules[i];
        end = begin;
      }else{
        end = cmidrules[i];
      }
    }
    
    cmidrules = [];

    if (r+1 === columnNameRows){
      content += "\\midrule";
    }
    content += lineBreak;
  }
  
  var str = "%Please add the following required packages to your document preamble:" + lineBreak
      + "%\\usepackage{booktabs, multirow}" + lineBreak
      + "%\\usepackage{soul}% for underlines" + lineBreak
      + "%\\usepackage[table]{xcolor} % for cell colors" + lineBreak
      + "%If the table is too wide, replace \\begin{table}[!htp]...\\end{table} with" + lineBreak
      + "%\\usepackage{changepage,threeparttable}" + lineBreak
      + "%\\begin{adjustwidth}{-2.5 cm}{-2.5 cm}\\centering\\begin{threeparttable}[!htb]...\\end{threeparttable}\\end{adjustwidth}" + lineBreak
      + "\\begin{table}[!htp]\\centering" + lineBreak
      + "\\caption{" + title + "}\\label{tab:  }" + lineBreak
      + "\\scriptsize" + lineBreak
      + "\\begin{tabular}{"+alignment+"}\\toprule" + lineBreak
      + content
      + "\\bottomrule" + lineBreak
      + "\\end{tabular}" + lineBreak
      + "\\end{table}";
  
  html.setContent(str);
  ui.showModalDialog(html,"LaTeX code:");
}

function readCell(object, format, fontWeight, underline, backgraound, isMergedRange, rows, cols){
  var type = typeof(object);
  if (type === "number" && format !== ""){
//    object = Utilities.formatString(format, object)
    var trimedFormat = format.replace("%","");
    var decimalPlaces;
    if (trimedFormat === "0"){
      decimalPlaces = 0;
    }else{
      decimalPlaces = trimedFormat.split(".")[1].length;
    }
    // if the data is a percentage value
    if (format.indexOf("%") !== -1){
      object = object*100;
      if(decimalPlaces <= 6){
        object = object.toFixed(decimalPlaces)
      }
      object = String(object) + "%"
    }else{
      if(decimalPlaces <= 6){
        object = object.toFixed(decimalPlaces)
      }
    }
  } 
  object = String(object);

  object = object.replace("%","\\%");

  if (fontWeight === 'bold'){
    object = "\\textbf{"+object+"}";
  }

  if (underline === 'underline'){
    object = "\\ul{" + object + "}";
  }
  // if the cell is highlighted
  if (backgraound !== '#ffffff'){
    var color = String(backgraound).replace("#","");
    object = "\\cellcolor[HTML]{"+color+"}" + object;
  }
  
  if (isMergedRange) {
    if (rows > 1) {
      object = "\\multirow{"+rows+"}{*}{" + object + "}";
    } 
    if (cols >1){
      object = "\\multicolumn{"+cols+"}{c}{"+ object +"}";
    }
  } 
  return object;
}
