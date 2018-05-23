/// <reference path ="./gNamespaces.ts" />
function deletePortConfirm(portName: string): void {
	const ui: GBase.Ui = SpreadsheetApp.getUi();
	const port: Portfolio = new Portfolio(portName);

	if (port.anyExist()) {		
		const confResponse: GBase.Button = ui.alert("WARNING", `You are about to permanently delete "${port.name}". Are you sure?`, ui.ButtonSet.OK_CANCEL);

		if (confResponse == ui.Button.OK) {	
			for (let victim of port.getSheetArray()) {
				ss.deleteSheet(victim);
			}
		}
		else return;
	}
	else {
		const errResponse: GBase.Button = ui.alert("Error", `"${port.name}" doesn't exist`, ui.ButtonSet.OK_CANCEL);

		if (errResponse == ui.Button.OK) {
			deletePort();
		}
	}

}

function deletePort(): void {
	const ui: GBase.Ui = SpreadsheetApp.getUi();
	const response: GBase.PromptResponse = ui.prompt("Delete Portfolio", "Enter which portfolio you wish to delete:", ui.ButtonSet.OK_CANCEL);

	if (response.getSelectedButton() == ui.Button.OK) deletePortConfirm(response.getResponseText());
}
