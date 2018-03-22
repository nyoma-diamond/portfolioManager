/// <reference path ="./gNamespaces.ts" />
function newStockBar(): void {
	const html: GHtml.HtmlOutput = HtmlService.createHtmlOutputFromFile("html/newStockBar").setTitle("Portfolio Management").setWidth(300);
	SpreadsheetApp.getUi().showSidebar(html);
}

function newStockOutput(portName: string, ticker: string, date: string, quantity: string, price: string): void {
	const port: Portfolio = new Portfolio(portName);
	const sheet: GSheets.Sheet = port.getSheetMap()[SheetType.Main];
	const priceOut: string = (price != "0") ? price : "=INDEX(GOOGLEFINANCE(A2,\"price\",DATE(RIGHT(C2,4),LEFT(C2,2),MID(C2,4,2))),2,2)";
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
		`=H2/H$${sheet.getMaxRows()-1}`,
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
	sheet.getRange(2,1,1,finalPortColumnCount).setNumberFormats([formats]);
	ss.setActiveSheet(sheet);
}

function stockSubmitCheck(portName: string, ticker: string, dateStr: string, quantityStr: string, priceStr: string): void | string {
	const port: Portfolio = new Portfolio(portName);
	const date: number = Date.parse(dateStr);
	const quantity: number = Number(quantityStr);
	const price: number = Number(priceStr);
	const validInputMap: object = { };
	const badIn: string[] = [];

	validInputMap[" Portfolio Name"] = port.importantExist();
	validInputMap[" Ticker"] = (ticker != "");
	validInputMap[" Date Obtained"] = (!isNaN(date) && dateStr != "");
	validInputMap[" Quantity"] = (quantity > 0 && quantityStr != "");
	validInputMap[" Price"] = (price >= 0 && priceStr != "");

	for (let key in validInputMap) {
		if (!validInputMap[key]) badIn.push(key);
	}

	if (badIn.length == 0) newStockOutput(portName, ticker, dateStr, quantityStr, priceStr);
	else if (badIn.length == 1 && badIn[0] == " Portfolio Name") {
		const ui: GBase.Ui = SpreadsheetApp.getUi();
		const button: GBase.Button = ui.alert("Alert", `The portfolio "${port.name}" does not exist. Would you like to create a new one?`, ui.ButtonSet.YES_NO_CANCEL)

		if (button === ui.Button.YES) {
			newPortBar();
		}
		else return button.toString();
	}
	else badInput(badIn);
}
