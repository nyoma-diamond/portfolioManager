/// <reference path ="./gNamespaces.ts" />
const ss: GSheets.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function onOpen(): void { 
	const portManMen: GBase.Menu = SpreadsheetApp.getUi().createMenu("Portfolio Management");
	portManMen.addItem("New Portfolio", "newPortBar").addToUi();
	portManMen.addItem("New Stock", "newStockBar").addToUi();
	portManMen.addItem("Add Cash", "addCash").addToUi();
	portManMen.addItem("Subtract Cash", "subtractCash").addToUi();
	portManMen.addItem("Delete Portfolio", "deletePort").addToUi();
}

function badInput(badIn: string[]) {
	const ui: GBase.Ui = SpreadsheetApp.getUi();

	if (badIn.length == 1) ui.alert("Error", `The following input was invalid:\n${badIn.toString()}`, ui.ButtonSet.OK);
	else ui.alert("Error", `The following inputs were invalid:\n${badIn.join(", ")}`, ui.ButtonSet.OK);
}

function testAlert(input: any): void {
	SpreadsheetApp.getUi().alert(`Input was "${input}"`);
}

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
const formats: string[] = [
	"\"text\"", 
	"\"text\"", 
	"mm/dd/yyyy", 
	"#,##0.00", 
	"\"$\"#,##0.00", 
	"\"$\"#,##0.00", 
	"\"$\"#,##0.00", 
	"\"$\"#,##0.00", 
	"#, ##0.00%", 
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
const finalPortColumnCount: number = legend.length; //Number of columns a new portfolio will have
const finalPortRowCount: number = 5; //Number of rows a new portfolio will have
const firstMarket = -4822502400000; //Unix timestamp for when NYSE first opened
