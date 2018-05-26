/// <reference path ="./gNamespaces.ts" />
function newPortBar(): void {
	const html: GHtml.HtmlOutput = HtmlService.createHtmlOutputFromFile("html/newPortBar").setTitle("Portfolio Management").setWidth(300);
	
	SpreadsheetApp.getUi().showSidebar(html);
}

function insertPortBase(newPortName: string, creDate: string, initCash: string): void {
	const port: Portfolio = new Portfolio(newPortName);
	const sheetName: string = port.sheetNameMap[SheetType.Main];
	const newSheet: GSheets.Sheet = ss.insertSheet(sheetName); //Sheet insertion
	const rowCount: number = newSheet.getMaxRows();
	const columnCount: number = newSheet.getMaxColumns();

	//Sheet prep START
	newSheet.deleteRows(finalPortRowCount, 1+rowCount-finalPortRowCount);
	newSheet.deleteColumns(finalPortColumnCount+1, columnCount-finalPortColumnCount);
	//Sheet prep END

	const wholeSheet: GSheets.Range = newSheet.getRange(1, 1, finalPortRowCount, finalPortColumnCount);
	const legendRow: GSheets.Range = newSheet.getRange(1, 1, 1, finalPortColumnCount);
	const bottom: GSheets.Range = newSheet.getRange(finalPortRowCount-1, 1, 2, finalPortColumnCount);
	const portSumm: GSheets.Range = newSheet.getRange(finalPortRowCount-1, 1, 1, finalPortColumnCount);
	const indexRow: GSheets.Range = newSheet.getRange(finalPortRowCount, 1, 1, finalPortColumnCount);
	const portSummVal: string[] = [
		"Total", 
		sheetName, 
		creDate, 
		"#N/A", 
		"#N/A", 
		initCash, 
		"#N/A", 
		"=SUM(H1:H2)", 
		`=H${finalPortRowCount-1}/F${finalPortRowCount-1}-1`, 
		`=H${finalPortRowCount-1}-F${finalPortRowCount-1}`, 
		`=H${finalPortRowCount-1}/H$${finalPortRowCount-1}`, 
		"day change", //UTILITY || HISTORY CALC GOES HERE
		"=SUM(M1:M2)", 
		"high52", 
		"low52", 
		"sparkline", 
		"#N/A", 
		"#N/A", 
		"Portfolio"
	];
	const inx: string[] = [
		".INX", 
		`=GOOGLEFINANCE(A${finalPortRowCount}, "name")`, 
		creDate, 
		`=F${finalPortRowCount-1}/E${finalPortRowCount}`, 
		`=INDEX(GOOGLEFINANCE(A${finalPortRowCount}, "price", DATE(RIGHT(C${finalPortRowCount}, 4), LEFT(C${finalPortRowCount}, 2), MID(C${finalPortRowCount}, 4, 2))), 2, 2)`, 
		`=D${finalPortRowCount}*E${finalPortRowCount}`, 
		`=GOOGLEFINANCE(A${finalPortRowCount}, "price")`, 
		`=G${finalPortRowCount}*D${finalPortRowCount}`, 
		`=H${finalPortRowCount}/F${finalPortRowCount}-1`, 
		`=H${finalPortRowCount}-F${finalPortRowCount}`, 
		"#N/A", 
		`=GOOGLEFINANCE(A${finalPortRowCount}, "changepct")`, 
		`=GOOGLEFINANCE(A${finalPortRowCount}, "closeyest")*L${finalPortRowCount}*D${finalPortRowCount}/100`, 
		`=GOOGLEFINANCE(A${finalPortRowCount}, "high52")`, 
		`=GOOGLEFINANCE(A${finalPortRowCount}, "low52")`, 
		`=SPARKLINE(GOOGLEFINANCE(A${finalPortRowCount}, "price", TODAY()-365, TODAY(), \"WEEKLY\"))`, 
		`=GOOGLEFINANCE(A${finalPortRowCount}, "eps")`, 
		`=GOOGLEFINANCE(A${finalPortRowCount}, "pe")`, 
		"Index"
	];
	const cashRow: string[] = [
		"Cash", 
		`${sheetName} Cash`, 
		creDate, 
		"#N/A", 
		"#N/A", 
		initCash, 
		"#N/A", 
		initCash, //UTILITY || HISTORY CALC GOES HERE FOR INTEREST (initCash is placeholder)
		"#N/A", 
		"#N/A", 
		`=H2/H$${finalPortRowCount-1}`, 
		"changepct", //UTILITY || HISTORY CALC GOES HERE FOR INTEREST
		"dayp&l", //UTILITY || HISTORY CALC GOES HERE FOR INTEREST
		"#N/A", 
		"#N/A", 
		"#N/A", 
		"#N/A", 
		"#N/A", 
		"Cash", 
	];

	//Value pasting START
	legendRow.setValues([legend]);
	portSumm.setValues([portSummVal]);
	indexRow.setValues([inx]);
	newSheet.getRange("A2:S2").setValues([cashRow]);
	//Value pasting END

	//Formatting START
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
	
	for (let row = 2; row <= finalPortRowCount; row++) {
		newSheet.getRange(row, 1, 1, finalPortColumnCount).setNumberFormats([formats]);
		newSheet.getRange(row, 1, 1, finalPortColumnCount).setHorizontalAlignments([horAligns]);
	}
	
	for (let column = 1; column <= finalPortColumnCount; column++) {
		newSheet.autoResizeColumn(column);
	}
	
	for (let row = 1; row <= finalPortRowCount-2; row++) {
		newSheet.setRowHeight(row, 25);
	}
	
	for (let row = finalPortRowCount-1; row <= finalPortRowCount; row++) {
		newSheet.setRowHeight(row, 50);
	}
	//Formatting END
}

function insertHistory(newPortName: string): void {
	const port: Portfolio = new Portfolio(newPortName);
	const newHist: GSheets.Sheet = ss.insertSheet(port.sheetNameMap[SheetType.History]);

	const rowCount: number = newHist.getMaxRows();
	const columnCount: number = newHist.getMaxColumns();
	newHist.deleteRows(3, rowCount-2);
	newHist.deleteColumns(4, columnCount-3);

	const wholeHist: GSheets.Range = newHist.getRange("A1:C3");
	const topRow: GSheets.Range = newHist.getRange("A1:C1");
	const curRow: GSheets.Range = newHist.getRange("A2:C2");

	topRow.setValues([["Date (mm/dd/yyyy)", "Portfolio Value", "Portfolio Value (Fridays only)"]])
	curRow.setValues([["=\"Current (\"&TEXT(NOW(), \"MM/dd/yyyy hh:mm\")&\")\"", `='${newPortName}'!H${finalPortRowCount-1}`, ""]])

	wholeHist.setVerticalAlignment("middle");
	topRow.setHorizontalAlignment("left");
	topRow.setFontWeight("bold");
	topRow.setNumberFormat("\"text\"");
	curRow.setNumberFormats([["\"text\"", "\"$\"#,##0.00", "\"$\"#,##0.00"]]);
	curRow.setHorizontalAlignment("right");

	for (let column = 1; column <= 3; column++) {
		newHist.autoResizeColumn(column);
	}

	for (let row = 1; row <= 2; row++) {
		newHist.setRowHeight(row, 21);
	}
}

function insertUtil(newPortName: string, intRate: string, compFreq: string): void {
	const port: Portfolio = new Portfolio(newPortName);
	const newUtil: GSheets.Sheet = ss.insertSheet(port.sheetNameMap[SheetType.Utility]);
}

function newPort(newPortName: string, creDate: string, initCash: string, intRate: string, compFreq: string): void {
	const newPort: Portfolio = new Portfolio(newPortName);

	if (!newPort.anyExist()) {
		const histSheetName: string = newPort.sheetNameMap[SheetType.History];
		const portName: string = newPort.name;
		const portSumm52: string[] = [
			`=MAX('${histSheetName}'!B2:B)`, 
			`=MIN('${histSheetName}'!B2:B)`, 
			"sparkline" //HISTORY CALC GOES HERE
		];

		insertPortBase(portName, creDate, initCash);
		insertHistory(portName);
		insertUtil(portName, intRate, compFreq);
		ss.getSheetByName(portName).getRange(`N${finalPortRowCount-1}:P${finalPortRowCount-1}`).setValues([portSumm52]); //this is to refresh the 52week calculations
		SpreadsheetApp.setActiveSheet(ss.getSheetByName(newPort.sheetNameMap[SheetType.Main]));
	}
	else {
		SpreadsheetApp.getUi().alert(`${newPortName} already exists`)
		return;
	}
}

function portSubmitCheck(newPortName: string, creDateStr: string, initCashStr: string, intRateStr: string, compFreqStr: string): void | string {
	const ui: GBase.Ui = SpreadsheetApp.getUi();
	const port: Portfolio = new Portfolio(newPortName);
	const initCash: number = Number(initCashStr);
	const date: number = Date.parse(creDateStr);
	const intRate: number = Number(intRateStr);
	const compFreq: number = Number(compFreqStr);
	const curDate: number = Date.now();
	const validInputMap: object = { };
	const badIn: string[] = [];

	validInputMap["Portfolio Name"] = [newPortName, !port.anyExist()];
	validInputMap["Creation Date"] = [creDateStr, (!isNaN(date) && creDateStr != "" && creDateStr.length != 4) && date < curDate && date > firstMarket];
	validInputMap["Initial Cash"] = [initCashStr, (initCashStr != "" && initCash >= 0)];
	validInputMap["Interest Rate"] = [intRateStr, (intRateStr != "" && intRate >= 0)];
	validInputMap["Compounding Frequency"] = [compFreqStr, (compFreqStr != "" && compFreq >= 0)];

	for (let key in validInputMap) {
		if (!validInputMap[key][1]) badIn.push(key);
	}

	if (badIn.length == 0) {
		let inputsAsString = "| ";
		
		for (let key in validInputMap) {
			inputsAsString += key + ": " + validInputMap[key][0] + " | "; // figure out a better way to do this (\n somehow?)
		}
		
		const button: GBase.Button = ui.alert("Please Confirm", inputsAsString, ui.ButtonSet.YES_NO);
		
		if (button == ui.Button.YES) newPort(newPortName, creDateStr, initCashStr, intRateStr, compFreqStr);
	}
	else if (badIn.length == 1 && badIn[0] == "Portfolio Name") {
		const button: GBase.Button = ui.alert("Error", `"${port.name}" already exists.`, ui.ButtonSet.OK_CANCEL);

		return button.toString();
	}
	else badInput(badIn);
}

