function recordHistory(histSheet) {
	var values = histSheet.getRange("A2:B2").getValues();
	var date = new Date();
	var month = date.getMonth()+1;
	var day = date.getDate();
	var year = date.getFullYear();
	var isFriday = date.getDay() == 5;
	values[0][0] = month + "/" + day + "/" + year;

	histSheet.insertRowBefore(3);
	var newRow = histSheet.getRange("A3:B3");
	newRow.setValues(values);
	newRow.setNumberFormats([["mm/dd/yyyy", "\"$\"#,##0.00"]]);

	// these both record only on fridays (two different methods). The second method is currently commented out
	histSheet.getRange("C3").setValue("=IF(WEEKDAY(A3)=6,B3,\"\")");
	/*if (isFriday === true) {
		histSheet.getRange("B3").copyTo(histSheet.getRange("C3"));
	}*/

	var rowCount = histSheet.getMaxRows();
	if (rowCount > 366) {
		histSheet.deleteRows(367,rowCount-366);
	}
}

function recordAllHistory() {
	var allSheets = ss.getSheets();

	var allHistSheets = allSheets.filter(
		function(entry) {
			return entry.getName().indexOf(" History") > -1
		}
	);

	for (var i = 0; i < allHistSheets.length; i++) {
		recordHistory(allHistSheets[i]);
	}
}
