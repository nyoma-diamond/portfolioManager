function recordHistory(histSheet) {
	const values = histSheet.getRange("A2:B2").getValues();
	const date = new Date();
	const month = date.getMonth()+1;
	const day = date.getDate();
	const year = date.getFullYear();
	const isFriday = date.getDay() == 5;
	values[0][0] = month + "/" + day + "/" + year;

	histSheet.insertRowBefore(3);
	const newRow = histSheet.getRange("A3:B3");
	newRow.setValues(values);
	newRow.setNumberFormats([["mm/dd/yyyy", "\"$\"#,##0.00"]]);

	// these both record only on fridays (two different methods). The second method is currently commented out
	histSheet.getRange("C3").setValue("=IF(WEEKDAY(A3)=6,B3,\"\")");
	/*if (isFriday === true) {
		histSheet.getRange("B3").copyTo(histSheet.getRange("C3"));
	}*/

	const rowCount = histSheet.getMaxRows();
	if (rowCount > 366) {
		histSheet.deleteRows(367,rowCount-366);
	}
}

function recordAllHistory() {
	const allSheets = ss.getSheets();

	const allHistSheets = allSheets.filter(
		function(entry) {
			return entry.getName().indexOf(" History") > -1
		}
	);

	for (let i = 0; i < allHistSheets.length; i++) {
		recordHistory(allHistSheets[i]);
	}
}
