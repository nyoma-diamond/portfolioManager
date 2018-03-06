function newPortBar() {
  var html = HtmlService.createHtmlOutputFromFile("newPortBar")
      .setTitle("Portfolio Management")
      .setWidth(300);
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

function newPort(newPortName, inCash, creDate, intRate, compFreq) {
  if (!ss.getSheetByName(newPortName)) {
    insertPortBase(newPortName, inCash, creDate);
    insertHistory(newPortName);
    insertUtil(newPortName, intRate, compFreq);
    SpreadsheetApp.setActiveSheet(ss.getSheetByName(newPortName));
  }
  else {
    SpreadsheetApp.getUi().alert(newPortName+" already exists")
    return;
  }
  
}

function insertPortBase(newPortName, inCash, creDate) {
  ss.insertSheet(newPortName);
  var newSheet = ss.getSheetByName(newPortName);
  var rowCount = newSheet.getMaxRows();
  var columnCount = newSheet.getMaxColumns();
  newSheet.deleteRows(6, rowCount-5);
  newSheet.deleteColumns(20, columnCount-19);
  var wholeSheet = newSheet.getRange("A1:S6");
  var legendRow = newSheet.getRange("A1:S1");
  var bottom = newSheet.getRange("A5:S6");
  var portSumm = newSheet.getRange("A5:S5");
  var indexRow = newSheet.getRange("A6:S6");
  
  var legend = [
    "Ticker", 
    "Company Name", 
    "Date Obtained", 
    "Quantity", 
    "Price Paid (per share)", 
    "Total Paid", 
    "Current Share Price", 
    "Market Value", 
    "Lifetime Return", 
    "P&L", 
    "Percent of Portfolio", 
    "Day Change", 
    "Day P&L", 
    "52 Week High", 
    "52 Week Low", 
    "52 Week Sparkline", 
    "Earnings Per Share", 
    "P/E Ratio", 
    "Sector"
  ];
    
  var portSummVal = [
    "Total", 
    newPortName, 
    creDate, 
    "#N/A", 
    "#N/A", 
    inCash,
    "#N/A", 
    "=SUM(H1:H2)", 
    "=H5/F5-1", 
    "=H5-F5", 
    "=H5/H$5", 
    "day change", //UTILITY || HISTORY CALC GOES HERE
    "=SUM(M2:M3)", 
    "=MAX('"+newPortName+" History'!B2:B)",
    "=MIN('"+newPortName+" History'!B2:B)",
    "sparkline", //HISTORY CALC GOES HERE
    "#N/A", 
    "#N/A", 
    "portfolio"
  ];
  
  var inx = [
    ".INX", 
    "=GOOGLEFINANCE(A6, \"name\")", 
    creDate,
    "=F5/E6",
    "=INDEX(GOOGLEFINANCE(A6,\"price\",DATE(RIGHT(C6,4),LEFT(C6,2),MID(C6,4,2))),2,2)",
    "=D6*E6", 
    "=GOOGLEFINANCE(A6, \"price\")", 
    "=G6*D6", 
    "=H6/F6-1", 
    "=H6-F6", 
    "#N/A", 
    "=GOOGLEFINANCE(A6, \"changepct\")", 
    "=GOOGLEFINANCE(A6, \"closeyest\")*L6*D6/100", 
    "=GOOGLEFINANCE(A6, \"high52\")", 
    "=GOOGLEFINANCE(A6, \"low52\")", 
    "=SPARKLINE(GOOGLEFINANCE(A6, \"price\", TODAY()-365, TODAY(), \"WEEKLY\"))", 
    "=GOOGLEFINANCE(A6, \"eps\")", 
    "=GOOGLEFINANCE(A6, \"pe\")", 
    "Index"
  ];
  
  var horAligns = [
    "left",
    "left",
    "right",
    "right",
    "right",
    "right",
    "right",
    "right",
    "right",
    "right",
    "right",
    "right",
    "right",
    "right",
    "right",
    "right",
    "right",
    "right",
    "left"
  ];
  
  legendRow.setValues([legend]);
  portSumm.setValues([portSummVal]);
  indexRow.setValues([inx]);

  wholeSheet.setVerticalAlignment("middle");
  wholeSheet.setFontFamily("Times New Roman");
  wholeSheet.setFontSize(11);
  legendRow.setNumberFormat("\"text\"");
  legendRow.setBackground("#bdbdbd");
  legendRow.setFontWeight("bold");
  legendRow.setHorizontalAlignment("left");
  legendRow.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  bottom.setFontSize(13);
  bottom.setFontWeight("bold");
  bottom.setBorder(true, false, true, false, false, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  portSumm.setBackground("#ffe599");
  indexRow.setBackground("#9fc5e8");
  
  //CONDITIONAL FORMATTING GOES HERE  
  
  for (var i = 1; i <= 19; i++) {
    newSheet.autoResizeColumn(i);
  }
  
    for (var i = 2; i <= 6; i++) {
    newSheet.getRange(i,1,1,19).setNumberFormats([formats]);
    newSheet.getRange(i,1,1,19).setHorizontalAlignments([horAligns]);
  }
  
  for (var i = 1; i <= 4; i++) {
    newSheet.setRowHeight(i, 25);
  }
  
  for (var i = 5; i <= 6; i++) {
    newSheet.setRowHeight(i, 50);
  }
}

function insertHistory(newPortName) {
  ss.insertSheet(newPortName+" History");
  var newHist = ss.getSheetByName(newPortName+" History");
  
  var rowCount = newHist.getMaxRows();
  var columnCount = newHist.getMaxColumns();
  newHist.deleteRows(3, rowCount-2);
  newHist.deleteColumns(4, columnCount-3);
  
  var wholeHist = newHist.getRange("A1:C3");
  var topRow = newHist.getRange("A1:C1"); 
  var curRow = newHist.getRange("A2:C2");
  
  topRow.setValues([["Date (mm/dd/yyyy)", "Portfolio Value", "Portfolio Value (Fridays only)"]])
  curRow.setValues([["=\"Current (\"&TEXT(NOW(), \"MM/dd/yyyy hh:mm\")&\")\"", "='"+newPortName+"'!H5", ""]])
  
  wholeHist.setVerticalAlignment("middle");
  topRow.setHorizontalAlignment("left");
  topRow.setFontWeight("bold");
  topRow.setNumberFormat("\"text\"");
  curRow.setNumberFormats([["\"text\"", "\"$\"#,##0.00",  "\"$\"#,##0.00"]]);
  curRow.setHorizontalAlignment("right");

  for (var i = 1; i <= 3; i++) {
    newHist.autoResizeColumn(i);
  }
  
  for (var i = 1; i <= 2; i++) {
    newHist.setRowHeight(i, 21);
  }
}

function insertUtil(newPortName, intRate, compFreq) {
  ss.insertSheet(newPortName+" Utility");
  var newUtil = ss.getSheetByName(newPortName+" Utility");
  
}