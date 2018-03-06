let ss = SpreadsheetApp.getActiveSpreadsheet();

function onOpen() {
	let portManMen = SpreadsheetApp.getUi().createMenu("Portfolio Management");
	portManMen.addItem("New Stock", "newStock").addToUi();
	portManMen.addItem("New Portfolio", "newPortBar").addToUi();
	portManMen.addItem("Base Test", "insertPortBase").addToUi();
	portManMen.addItem("History Test", "recordAllHistory").addToUi();
	portManMen.addItem("History Base Test", "insertHistory").addToUi();
}

function testAlert(input) {
	SpreadsheetApp.getUi().alert("input is "+input);
}

function badInput(badIn) {
	let ui = SpreadsheetApp.getUi();
	if (badIn.length == 1) {
		ui.alert("Error", "The following input was invalid:\n"+badIn.toString(), ui.ButtonSet.OK);
	}
	else {
		ui.alert("Error", "The following inputs were invalid:\n"+badIn.toString(), ui.ButtonSet.OK);
	}
}

function checkSheetExist(nameIn) {
	if (!ss.getSheetByName(nameIn)) {
		return false;
	}
	else {
		return true;
	}
}

let formats = [
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
