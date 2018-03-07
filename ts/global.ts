/// <reference path ="./gNameSpaces.ts" />
const ss: GSheets.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function onOpen(): void {
	const portManMen: GBase.Menu = SpreadsheetApp.getUi().createMenu("Portfolio Management");
	portManMen.addItem("New Stock", "newStock").addToUi();
	portManMen.addItem("New Portfolio", "newPortBar").addToUi();
	portManMen.addItem("Base Test", "insertPortBase").addToUi();
	portManMen.addItem("History Test", "recordAllHistory").addToUi();
	portManMen.addItem("History Base Test", "insertHistory").addToUi();
}

function testAlert(input: any): void {
	SpreadsheetApp.getUi().alert("input is "+input);
}

function badInput(badIn: string[]) {
	const ui: GBase.Ui = SpreadsheetApp.getUi();
	if (badIn.length == 1) {
		ui.alert("Error", "The following input was invalid:\n"+badIn.toString(), ui.ButtonSet.OK);
	}
	else {
		ui.alert("Error", "The following inputs were invalid:\n"+badIn.toString(), ui.ButtonSet.OK);
	}
}

function checkSheetExist(nameIn: string): boolean {
	if (!ss.getSheetByName(nameIn)) {
		return false;
	}
	else {
		return true;
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
