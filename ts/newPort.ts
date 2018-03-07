/// <reference path ="./gNamespaces.ts" />
function newPortBar(): void {
	const html: GHtml.HtmlOutput = HtmlService.createHtmlOutputFromFile("html/newPortBar")
			.setTitle("Portfolio Management")
			.setWidth(300);
	SpreadsheetApp.getUi()
			.showSidebar(html);
}

function newPort(newPortName: string, inCash: string, creDate: string, intRate: string, compFreq: string): void {
	if (!ss.getSheetByName(newPortName)) {
		insertPortBase(newPortName, inCash, creDate);
		insertHistory(newPortName);
		insertUtil(newPortName, intRate, compFreq);
		SpreadsheetApp.setActiveSheet(ss.getSheetByName(newPortName));
	}
	else {
		SpreadsheetApp.getUi().alert(`${newPortName} already exists`)
		return;
	}
}

function insertPortBase(newPortName: string, inCash: string, creDate: string): void {
	const legend: string[] = [
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

	const finalRowCount: number = 6;
	const finalColumnCount: number = legend.length;

	ss.insertSheet(newPortName);
	const newSheet: GSheets.Sheet = ss.getSheetByName(newPortName);
	const rowCount: number = newSheet.getMaxRows();
	const columnCount: number = newSheet.getMaxColumns();
	newSheet.deleteRows(finalRowCount, 1+rowCount-finalRowCount);
	newSheet.deleteColumns(finalColumnCount+1, columnCount-finalColumnCount);
	const wholeSheet: GSheets.Range = newSheet.getRange("A1:S6");
	const legendRow: GSheets.Range = newSheet.getRange("A1:S1");
	const bottom: GSheets.Range = newSheet.getRange("A5:S6");
	const portSumm: GSheets.Range = newSheet.getRange("A5:S5");
	const indexRow: GSheets.Range = newSheet.getRange("A6:S6");

	const portSummVal: string[] = [
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
		`=MAX('${newPortName} History'!B2:B`,
		`=MIN('${newPortName} History'!B2:B`,
		"sparkline", //HISTORY CALC GOES HERE
		"#N/A",
		"#N/A",
		"portfolio"
	];

	const inx: string[] = [
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

	const horAligns: string[] = [
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

	for (let column = 1; column <= finalColumnCount; column++) {
		newSheet.autoResizeColumn(column);
	}

	for (let row = 2; row <= finalRowCount; row++) {
		newSheet.getRange(row, 1, 1, finalColumnCount).setNumberFormats([formats]);
		newSheet.getRange(row, 1, 1, finalColumnCount).setHorizontalAlignments([horAligns]);
	}

	for (let row = 1; row <= finalRowCount-2; row++) {
		newSheet.setRowHeight(row, 25);
	}

	for (let row = finalRowCount-1; row <= finalRowCount; row++) {
		newSheet.setRowHeight(row, 50);
	}
}

function insertHistory(newPortName: string): void {
	ss.insertSheet(`${newPortName} History`);
	const newHist: GSheets.Sheet = ss.getSheetByName(`${newPortName} History`);

	const rowCount: number = newHist.getMaxRows();
	const columnCount: number = newHist.getMaxColumns();
	newHist.deleteRows(3, rowCount-2);
	newHist.deleteColumns(4, columnCount-3);

	const wholeHist: GSheets.Range = newHist.getRange("A1:C3");
	const topRow: GSheets.Range = newHist.getRange("A1:C1");
	const curRow: GSheets.Range = newHist.getRange("A2:C2");

	topRow.setValues([["Date (mm/dd/yyyy)", "Portfolio Value", "Portfolio Value (Fridays only)"]])
	curRow.setValues([["=\"Current (\"&TEXT(NOW(), \"MM/dd/yyyy hh:mm\")&\")\"", `='${newPortName}'~H5`, ""]])

	wholeHist.setVerticalAlignment("middle");
	topRow.setHorizontalAlignment("left");
	topRow.setFontWeight("bold");
	topRow.setNumberFormat("\"text\"");
	curRow.setNumberFormats([["\"text\"", "\"$\"#,##0.00",  "\"$\"#,##0.00"]]);
	curRow.setHorizontalAlignment("right");

	for (let i = 1; i <= 3; i++) {
		newHist.autoResizeColumn(i);
	}

	for (let i = 1; i <= 2; i++) {
		newHist.setRowHeight(i, 21);
	}
}

function insertUtil(newPortName: string, intRate: string, compFreq: string): void {
	ss.insertSheet(`${newPortName} Utility`);
	const newUtil: GSheets.Sheet = ss.getSheetByName(`${newPortName} Utility`);
}
