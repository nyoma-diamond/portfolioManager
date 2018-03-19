/// <reference path ="./gNamespaces.ts" />
const ss: GSheets.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function onOpen(): void {
	const portManMen: GBase.Menu = SpreadsheetApp.getUi().createMenu("Portfolio Management");
	portManMen.addItem("New Portfolio", "newPortBar").addToUi();
	portManMen.addItem("New Stock", "newStock").addToUi();
	portManMen.addItem("Delete Portfolio", "deletePort").addToUi();
}

function badInput(badIn: string[]) {
	const ui: GBase.Ui = SpreadsheetApp.getUi();
	if (badIn.length == 1) {
		ui.alert("Error", `The following input was invalid:\n${badIn.toString()}`, ui.ButtonSet.OK);
	}
	else {
		ui.alert("Error", `The following inputs were invalid:\n${badIn.toString()}`, ui.ButtonSet.OK);
	}
}

function testAlert(input: any): void {
	SpreadsheetApp.getUi().alert(`Input is "${input}"`);
}

function checkSheetExist(nameIn: string): boolean {
	if (!ss.getSheetByName(nameIn)) {
		return false;
	}
	else {
		return true;
	}
}

function deletePortConfirm(portName: string): void {
	const ui: GBase.Ui = SpreadsheetApp.getUi();
	const port: Portfolio = new Portfolio(portName);

	if (port.anyExist()) {		
		let confResponse: GBase.Button = ui.alert("WARNING", "You are about to permanently delete a portfolio. Are you sure?", ui.ButtonSet.OK_CANCEL);
		if (confResponse == ui.Button.OK) {
			for (let victim of port.getSheetArray()) {
				ss.deleteSheet(victim);
			}
		}
	}
	else {
		const errorResponse: GBase.Button = ui.alert("Error", `${portName} doesn't exist`, ui.ButtonSet.OK_CANCEL)
		if (errorResponse == ui.Button.OK) {
			deletePort();
		}
	}
}

function deletePort(): void {
	const ui: GBase.Ui = SpreadsheetApp.getUi();
	let response: GBase.PromptResponse = ui.prompt("Delete Portfolio", "Enter which portfolio you wish to delete", ui.ButtonSet.OK_CANCEL);
	let portName: string = response.getResponseText();

	if (response.getSelectedButton() == ui.Button.OK) {
		deletePortConfirm(portName);
	}
}

const formats: string[] = [
	"\"text\"",
	"\"text\"",
	"mm/dd/yyyy",
	"#,##0.00",
	"\"$\"#,##0.00",
	"\"$\"#,##0.00",
	"\"$\"#,##0.00",
	"\"$\"#,##0.00",
	"#,##0.00%",
	"\"$\"#,##0.00",
	"#,##0.00%",
	"#,0.00\"%\"",
	"\"$\"#,##0.00",
	"\"$\"#,##0.00",
	"\"$\"#,##0.00",
	"\"$\"#,##0.00",
	"#,##0.00",
	"#,##0.00",
	"\"text\""
];

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

const finalNewPortColumnCount: number = legend.length;