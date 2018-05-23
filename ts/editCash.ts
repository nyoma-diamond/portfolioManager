/// <reference path ="./gNamespaces.ts" />
function addCash(): void {
	editCash(true);
}

function subtractCash(): void {
	editCash(false);
}

function editCash(addBoo: boolean): void {
	const ui: GBase.Ui = SpreadsheetApp.getUi();
	const operatorUp: string = (addBoo) ? "Add" : "Subtract";
	const operator: string = (addBoo) ? "add" : "subtract";
	const operatorGrammar: String = (addBoo) ? "to" : "from";
	const response: GBase.PromptResponse = ui.prompt(`${operatorUp} Cash`, `Enter which portfolio you wish to ${operator} cash ${operatorGrammar}:`, ui.ButtonSet.OK_CANCEL);

	if (response.getSelectedButton() == ui.Button.OK) editCashConfirm(response.getResponseText(), addBoo);
}

function editCashConfirm(portName: string, addBoo: boolean): void {
	const ui: GBase.Ui = SpreadsheetApp.getUi();
	const port: Portfolio = new Portfolio(portName);
	let ready: boolean = false;
	let failed: boolean = false;
	const operatorUp: string = (addBoo) ? "Add" : "Subtract";
	const operator: string = (addBoo) ? "add" : "subtract";
	const operatorGrammar: String = (addBoo) ? "to" : "from";

	if (port.anyExist()) {
		while (!ready) {
			const title: string = (!failed) ? `${operatorUp} Cash` : "Error";
			const request: string = (!failed) ? `How much cash would you like to ${operator} ${operatorGrammar} "${port.name}"?` : `Please enter a valid amount to ${operator}`;
			const confResponse: GBase.PromptResponse = ui.prompt(title, request, ui.ButtonSet.OK_CANCEL);
			const respText: string = confResponse.getResponseText();
			const respButton: GBase.Button = confResponse.getSelectedButton();

			if (respButton == ui.Button.OK && /^\d+$/.test(respText)) {	
				const sheet: GSheets.Sheet = port.getSheetMap()[SheetType.Main];
				const initCashAddress: GSheets.Range = sheet.getRange(sheet.getMaxRows()-3, 8);
				const initCash: number = Number(initCashAddress.getValue().toString());
				const cashToAdd: number = Number(respText);

				if (addBoo) initCashAddress.setValue(initCash+cashToAdd);
				else initCashAddress.setValue(initCash-cashToAdd);

				ready = true;
			}
			else if (respButton != ui.Button.OK) return;
			else failed = true;
		}
	}
	else {
		const errResponse: GBase.Button = ui.alert("Error", `"${port.name}" doesn't exist`, ui.ButtonSet.OK_CANCEL);

		if (errResponse == ui.Button.OK) editCash(addBoo);
	}
}
