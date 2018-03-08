/// <reference path ="./gNamespaces.ts" />
const ss: GSheets.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function onOpen(): void {
	const portManMen: GBase.Menu = SpreadsheetApp.getUi().createMenu("Portfolio Management");
	portManMen.addItem("New Stock", "newStock").addToUi();
	portManMen.addItem("New Portfolio", "newPortBar").addToUi();
	portManMen.addItem("Delete Portfolio", "deletePort").addToUi();
	portManMen.addItem("Base Test", "insertPortBase").addToUi();
	portManMen.addItem("History Test", "recordAllHistory").addToUi();
	portManMen.addItem("History Base Test", "insertHistory").addToUi();
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

function deletePort(): void {
	const ui: GBase.Ui = SpreadsheetApp.getUi();
	let response: GBase.PromptResponse = ui.prompt("Delete Portfolio", "Enter which portfolio you wish to delete", ui.ButtonSet.OK_CANCEL);
	let portName: string = response.getResponseText();
	let histName: string = `${portName} History`;
	let utilName: string = `${portName} Utility`;

	if (response.getSelectedButton() == ui.Button.OK) {
		deletePortConfirm(portName, histName, utilName);
	}
	else return;
}

function deletePortConfirm(portName: string, histName: string, utilName: string) {
	const ui: GBase.Ui = SpreadsheetApp.getUi();
	if (checkSheetExist(portName) || checkSheetExist(histName) || checkSheetExist(utilName)) {		
		let confResponse: GBase.Button = ui.alert("WARNING", "You are about to permanently delete a portfolio, this cannot be undone. Are you sure?", ui.ButtonSet.OK_CANCEL);
		if (confResponse == ui.Button.OK) {
			const existsMap: object = { };

			existsMap[`${portName}`] = checkSheetExist(portName);
			existsMap[`${portName} History`] = checkSheetExist(histName);
			existsMap[`${portName} Utility`] = checkSheetExist(utilName);

			const toDelete: GSheets.Sheet[] = [];
			for (let key in existsMap) {
				if (existsMap[key]) {
					toDelete.push(ss.getSheetByName(key));
				}
			}
			for (let victim of toDelete) {
				ss.deleteSheet(victim);
			}
		}
	}
	else {
		const errorResponse: GBase.Button = ui.alert("Error", `${portName} doesn't exist`, ui.ButtonSet.OK_CANCEL)
		if (errorResponse == ui.Button.OK) {
			deletePort();
		}
		else return;
	}
}

function testAlert(input: any): void {
	SpreadsheetApp.getUi().alert(`Input is ${input}`);
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
