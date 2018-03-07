/// <reference path ="./gNameSpaces.ts" />
function newStock(): void {
	const html: GHtml.HtmlOutput = HtmlService.createHtmlOutputFromFile("html/newStockBar")
			.setTitle("Portfolio Management")
			.setWidth(300);
	SpreadsheetApp.getUi()
			.showSidebar(html);
}

function newStockOutput(portName: string, ticker: string, date: string, quantity: string, price: string): void {
	const sheet: GSheets.Sheet = ss.getSheetByName(portName);
	SpreadsheetApp.setActiveSheet(sheet);

	const newData: string[] = [
		ticker,
		"=GOOGLEFINANCE($A2, \"name\")",
		date,
		quantity,
		price,
		"=D2*E2",
		"=GOOGLEFINANCE(A2, \"price\")",
		"=G2*D2",
		"=H2/F2-1",
		"=H2-F2",
		"=H2/H$"+sheet.getMaxRows(),
		"=GOOGLEFINANCE(A2, \"changepct\")",
		"=GOOGLEFINANCE(A2, \"closeyest\")*L2*D2/100",
		"=GOOGLEFINANCE(A2, \"high52\")",
		"=GOOGLEFINANCE(A2, \"low52\")",
		"=SPARKLINE(GOOGLEFINANCE(A2, \"price\", TODAY()-365, TODAY(), \"WEEKLY\"))",
		"=GOOGLEFINANCE(A2, \"eps\")",
		"=GOOGLEFINANCE(A2, \"pe\")",
		"Sector" //SECTOR LOOKUP GOES HERE
	];
	sheet.insertRowBefore(2);
	sheet.getRange("A2:S2").setValues([newData]);
	sheet.getRange(2,1,1,19).setNumberFormats([formats]);
}

function submitCheck(portName: string, ticker: string, dateStr: string, quantityStr: string, priceStr: string): void {
	const date: number = Date.parse(dateStr);
	const quantity: number = Number(quantityStr);
	const price: number = Number(priceStr);

	const portTitle: string = " Portfolio Name";
	const tickerTitle: string = " Ticker";
	const dateTitle: string = " Date Obtained";
	const quantityTitle: string = " Quantity";
	const priceTitle: string = " Price per Share";

	const validMap: object = { };

	validMap[portTitle] = checkSheetExist(portName) ? true : portTitle;
	validMap[tickerTitle] = (ticker != "") ? true : tickerTitle;
	validMap[dateTitle] = (!isNaN(date) && dateStr != "") ? true : dateTitle;
	validMap[quantityTitle] = (quantity > 0 && quantityStr != "") ? true : quantityTitle;
	validMap[priceTitle] = (price >= 0 && priceStr != "") ? true : priceTitle;

	const badIn: string[] = [];
	for (let key in validMap) {
		if (validMap[key] !== true) {
			badIn.push(key);
		}
	}

	if (badIn.length == 0) {
		newStockOutput(portName, ticker, dateStr, quantityStr, priceStr);
	}
	else {
		badInput(badIn);
	}
}
