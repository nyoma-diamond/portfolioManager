function newStock() {
	let html = HtmlService.createHtmlOutputFromFile("html/newStockBar")
			.setTitle("Portfolio Management")
			.setWidth(300);
	SpreadsheetApp.getUi()
			.showSidebar(html);
}

function submitCheck(portName, ticker, dateStr, quantityStr, priceStr) {
	let date = Date.parse(dateStr);
	let quantity = Number(quantityStr);
	let price = Number(priceStr);

	let portTitle = " Portfolio Name";
	let tickerTitle = " Ticker";
	let dateTitle = " Date Obtained";
	let quantityTitle = " Quantity";
	let priceTitle = " Price per Share";

	let validMap = { };

	validMap[portTitle] = checkSheetExist(portName) ? true : portTitle;
	validMap[tickerTitle] = (ticker != "") ? true : tickerTitle;
	validMap[dateTitle] = (!isNaN(date) && dateStr != "") ? true : dateTitle;
	validMap[quantityTitle] = (quantity > 0 && quantityStr != "") ? true : quantityTitle;
	validMap[priceTitle] = (price >= 0 && priceStr != "") ? true : priceTitle;

	let badIn = [];
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

function newStockOutput(portName, ticker, date, quantity, price) {
	let sheet = ss.getSheetByName(portName);
	SpreadsheetApp.setActiveSheet(sheet);

	let newData = [
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
