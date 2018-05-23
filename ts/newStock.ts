/// <reference path ="./gNamespaces.ts" />
function newStockBar(): void {
	const html: GHtml.HtmlOutput = HtmlService.createHtmlOutputFromFile("html/newStockBar").setTitle("Portfolio Management").setWidth(300);
	
	SpreadsheetApp.getUi().showSidebar(html);
}

function newStockOutput(portName: string, ticker: string, date: string, quantity: string, price: string): void {
	const port: Portfolio = new Portfolio(portName);
	const sheet: GSheets.Sheet = port.getSheetMap()[SheetType.Main];
	const sheetRows: number = sheet.getMaxRows();
	const initCashAddress: GSheets.Range = sheet.getRange(sheetRows-3, 8);
	const initCash: number = Number(initCashAddress.getValue().toString());
	const company = JSON.parse(UrlFetchApp.fetch(`https://api.iextrading.com/1.0/stock/${ticker}/company`, {"muteHttpExceptions": true}).getContentText());
	const compSector: string = (company.sector != "") ? company.sector : "#N/A";
	const priceOut: string = (price != "0") ? price : "=INDEX(GOOGLEFINANCE(A2, \"price\", DATE(RIGHT(C2, 4), LEFT(C2, 2), MID(C2, 4, 2))), 2, 2)";
	const newData: string[] = [
		ticker, 
		"=GOOGLEFINANCE(A2, \"name\")", 
		date, 
		quantity, 
		priceOut, 
		"=D2*E2", 
		"=GOOGLEFINANCE(A2, \"price\")", 
		"=G2*D2", 
		"=H2/F2-1", 
		"=H2-F2", 
		`=H2/H$${sheetRows}`, 
		"=GOOGLEFINANCE(A2, \"changepct\")", 
		"=GOOGLEFINANCE(A2, \"closeyest\")*L2*D2/100", 
		"=GOOGLEFINANCE(A2, \"high52\")", 
		"=GOOGLEFINANCE(A2, \"low52\")", 
		"=SPARKLINE(GOOGLEFINANCE(A2, \"price\", TODAY()-365, TODAY(), \"WEEKLY\"))", 
		"=GOOGLEFINANCE(A2, \"eps\")", 
		"=GOOGLEFINANCE(A2, \"pe\")", 
		compSector
	];

	initCashAddress.setValue(initCash-parseInt(price)*parseInt(quantity));
	sheet.insertRowBefore(2);
	sheet.getRange(2, 1, 1, finalPortColumnCount).setValues([newData]);
	sheet.getRange(2, 1, 1, finalPortColumnCount).setNumberFormats([formats]);
	ss.setActiveSheet(sheet);
}

function stockSubmitCheck(portName: string, ticker: string, dateStr: string, quantityStr: string, priceStr: string): void | string {
	const ui: GBase.Ui = SpreadsheetApp.getUi();
	const port: Portfolio = new Portfolio(portName);
	const date: number = Date.parse(dateStr);
	const quantity: number = Number(quantityStr);
	const price: number = Number(priceStr);
	const curDate: number = Date.now();
	const validInputMap: object = { };
	const badIn: string[] = [];
	const tickLookup: GUrl.HTTPResponse = UrlFetchApp.fetch(`https://api.iextrading.com/1.0/stock/${ticker}/company`, {"muteHttpExceptions": true});
	
	validInputMap["Portfolio Name"] = [portName, port.importantExist()];
	validInputMap["Ticker"] = [ticker, (tickLookup.getResponseCode() != 404)];
	validInputMap["Date Obtained"] = [dateStr, (!isNaN(date) && dateStr != "" && dateStr.length != 4 && date < curDate && date > firstMarket)];
	validInputMap["Quantity"] = [quantityStr, (quantity > 0 && quantityStr != "")];
	validInputMap["Price"] = [priceStr, (price >= 0 && priceStr != "")];

	for (let key in validInputMap) {
		if (!validInputMap[key][1]) badIn.push(key);
	}

	if (badIn.length == 0) {
			let inputsAsString = "| ";
		
		for (let key in validInputMap) {
			inputsAsString += key + ": " + validInputMap[key][0] + " | "; // figure out a better way to do this (\n somehow?)
		}
		
		const button: GBase.Button = ui.alert("Please Confirm", inputsAsString, ui.ButtonSet.YES_NO);
		
		if (button == ui.Button.YES) newStockOutput(portName, ticker, dateStr, quantityStr, priceStr);
	}
	else if (badIn.length == 1 && badIn[0] == "Portfolio Name") {
		const newPortButton: GBase.Button = ui.alert("Alert", `The portfolio "${port.name}" does not exist. Would you like to create a new one?`, ui.ButtonSet.YES_NO_CANCEL);
		return newPortButton.toString();
	}
	else badInput(badIn);
}
