/// <reference path ="./gNamespaces.ts" />
function newPortBar(): void {
	const html: GHtml.HtmlOutput = HtmlService.createHtmlOutputFromFile("html/newPortBar")
			.setTitle("Portfolio Management")
			.setWidth(300);
	SpreadsheetApp.getUi()
			.showSidebar(html);
}

function newPort(newPortName: string, inCash: string, creDate: string, intRate: string, compFreq: string): void {
	const finalPortRowCount: number = 6;

	if (!ss.getSheetByName(newPortName)) {
		const portSumm52: string[] = [
			`=MAX('${newPortName} History'!B2:B)`,
			`=MIN('${newPortName} History'!B2:B)`,
			"sparkline" //HISTORY CALC GOES HERE
		];

		insertPortBase(newPortName, inCash, creDate, finalPortRowCount);
		insertHistory(newPortName, finalPortRowCount);
		insertUtil(newPortName, intRate, compFreq);
		ss.getSheetByName(newPortName).getRange(`N${finalPortRowCount-1}:P${finalPortRowCount-1}`).setValues([portSumm52]); //this is to refresh the 52week calculations
		SpreadsheetApp.setActiveSheet(ss.getSheetByName(newPortName));
	}
	else {
		SpreadsheetApp.getUi().alert(`${newPortName} already exists`)
		return;
	}
}

function insertPortBase(newPortName: string, inCash: string, creDate: string, finalRowCount: number): void {
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

	const finalColumnCount: number = legend.length;

	ss.insertSheet(newPortName);
	const newSheet: GSheets.Sheet = ss.getSheetByName(newPortName);
	const rowCount: number = newSheet.getMaxRows();
	const columnCount: number = newSheet.getMaxColumns();
	newSheet.deleteRows(finalRowCount, 1+rowCount-finalRowCount);
	newSheet.deleteColumns(finalColumnCount+1, columnCount-finalColumnCount);
	const wholeSheet: GSheets.Range = newSheet.getRange(`A1:S${finalRowCount}`);
	const legendRow: GSheets.Range = newSheet.getRange("A1:S1");
	const bottom: GSheets.Range = newSheet.getRange(`A${finalRowCount-1}:S${finalRowCount}`);
	const portSumm: GSheets.Range = newSheet.getRange(`A${finalRowCount-1}:S${finalRowCount-1}`);
	const indexRow: GSheets.Range = newSheet.getRange(`A${finalRowCount}:S${finalRowCount}`);

	const portSummVal: string[] = [
		"Total",
		newPortName,
		creDate,
		"#N/A",
		"#N/A",
		inCash,
		"#N/A",
		"=SUM(H1:H2)",
		`=H${finalRowCount-1}/F${finalRowCount-1}-1`,
		`=H${finalRowCount-1}-F${finalRowCount-1}`,
		`=H${finalRowCount-1}/H$${finalRowCount-1}`,
		"day change", //UTILITY || HISTORY CALC GOES HERE
		"=SUM(M1:M2)",
		`=MAX('${newPortName} History'!B2:B)`,
		`=MIN('${newPortName} History'!B2:B)`,
		"sparkline", //HISTORY CALC GOES HERE
		"#N/A",
		"#N/A",
		"Portfolio"
	];

	const inx: string[] = [
		".INX",
		`=GOOGLEFINANCE(A${finalRowCount}, "name")`,
		creDate,
		`=F${finalRowCount-1}/E${finalRowCount}`,
		`=INDEX(GOOGLEFINANCE(A${finalRowCount},"price",DATE(RIGHT(C${finalRowCount},4),LEFT(C${finalRowCount},2),MID(C${finalRowCount},4,2))),2,2)`,
		`=D${finalRowCount}*E${finalRowCount}`,
		`=GOOGLEFINANCE(A${finalRowCount}, "price")`,
		`=G${finalRowCount}*D${finalRowCount}`,
		`=H${finalRowCount}/F${finalRowCount}-1`,
		`=H${finalRowCount}-F${finalRowCount}`,
		"#N/A",
		`=GOOGLEFINANCE(A${finalRowCount}, "changepct")`,
		`=GOOGLEFINANCE(${finalRowCount}, "closeyest")*L${finalRowCount}*D${finalRowCount}/100`,
		`=GOOGLEFINANCE(A${finalRowCount}, "high52")`,
		`=GOOGLEFINANCE(A${finalRowCount}, "low52")`,
		`=SPARKLINE(GOOGLEFINANCE(A${finalRowCount}, "price", TODAY()-365, TODAY(), \"WEEKLY\"))`,
		`=GOOGLEFINANCE(A${finalRowCount}, "eps")`,
		`=GOOGLEFINANCE(A${finalRowCount}, "pe")`,
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

function insertHistory(newPortName: string, finalPortRowCount: number): void {
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
	curRow.setValues([["=\"Current (\"&TEXT(NOW(), \"MM/dd/yyyy hh:mm\")&\")\"", `='${newPortName}'!H${finalPortRowCount}`, ""]])

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
