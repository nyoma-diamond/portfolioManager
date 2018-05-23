/// <reference path ="./gNamespaces.ts" />
function addCashConfirm(portName: string): void {
	const ui: GBase.Ui = SpreadsheetApp.getUi();
	const port: Portfolio = new Portfolio(portName);
	let ready: boolean = false;
	let failed: boolean = false;

	if (port.anyExist()) {
		while (!ready) {
			const title: string = (!failed) ? "Add Cash" : "Error";
			const request: string = (!failed) ? `How much cash would you like to add to "${port.name}"?` : `Please enter a valid amount to add`;
			const confResponse: GBase.PromptResponse = ui.prompt(title, request, ui.ButtonSet.OK_CANCEL);
			const respText: string = confResponse.getResponseText();
			const respButton: GBase.Button = confResponse.getSelectedButton();

			if (respButton == ui.Button.OK && /^\d+$/.test(respText)) {	
				const sheet: GSheets.Sheet = port.getSheetMap()[SheetType.Main];
				const initCashAddress: GSheets.Range = sheet.getRange(sheet.getMaxRows()-3, 8);
				const initCash: number = Number(initCashAddress.getValue().toString());
				const cashToAdd: number = Number(respText);

				initCashAddress.setValue(initCash+cashToAdd);
				ready = true;
			}
			else if (respButton != ui.Button.OK) return;
			else failed = true;
		}
	}
	else {
		const errResponse: GBase.Button = ui.alert("Error", `"${port.name}" doesn't exist`, ui.ButtonSet.OK_CANCEL);

		if (errResponse == ui.Button.OK) addCash();
	}
}

function addCash(): void {
	const ui: GBase.Ui = SpreadsheetApp.getUi();
	const response: GBase.PromptResponse = ui.prompt("Add Cash", "Enter which portfolio you wish to add cash to:", ui.ButtonSet.OK_CANCEL);

	if (response.getSelectedButton() == ui.Button.OK) addCashConfirm(response.getResponseText());
}

function subtractCashConfirm(portName: string): void {
	const ui: GBase.Ui = SpreadsheetApp.getUi();
	const port: Portfolio = new Portfolio(portName);
	let ready: boolean = false;
	let failed: boolean = false;

	if (port.anyExist()) {
		while (!ready) {
			const title: string = (!failed) ? "Subtract Cash" : "Error";
			const request: string = (!failed) ? `How much cash would you like to subtract from "${port.name}"?` : `Please enter a valid amount to subtract`;
			const confResponse: GBase.PromptResponse = ui.prompt(title, request, ui.ButtonSet.OK_CANCEL);
			const respText: string = confResponse.getResponseText();
			const respButton: GBase.Button = confResponse.getSelectedButton();

			if (respButton == ui.Button.OK && /^\d+$/.test(respText)) {	
				const sheet: GSheets.Sheet = port.getSheetMap()[SheetType.Main];
				const initCashAddress: GSheets.Range = sheet.getRange(sheet.getMaxRows()-3, 8);
				const initCash: number = Number(initCashAddress.getValue().toString());
				const cashToSubtract: number = Number(respText);

				initCashAddress.setValue(initCash-cashToSubtract);
				ready = true;
			}
			else if (respButton != ui.Button.OK) return;
			else failed = true;
		}
	}
	else {
		const errResponse: GBase.Button = ui.alert("Error", `"${port.name}" doesn't exist`, ui.ButtonSet.OK_CANCEL);

		if (errResponse == ui.Button.OK) subtractCash();
	}
}

function subtractCash(): void {
	const ui: GBase.Ui = SpreadsheetApp.getUi();
	const response: GBase.PromptResponse = ui.prompt("Subtract Cash", "Enter which portfolio you wish to subtract cash from:", ui.ButtonSet.OK_CANCEL);

	if (response.getSelectedButton() == ui.Button.OK) subtractCashConfirm(response.getResponseText());
}