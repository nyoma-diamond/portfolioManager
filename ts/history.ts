/// <reference path ="./gNamespaces.ts" />
function recordHistory(histSheet: GSheets.Sheet): void {
	const values: Object[][] = histSheet.getRange("A2:B2").getValues();
	const date: Date = new Date();
	const month: number = date.getMonth()+1;
	const day: number = date.getDate();
	const year: number = date.getFullYear();
	values[0][0] = month + "/" + day + "/" + year;
	const isFriday: boolean = date.getDay() == 5;

	histSheet.insertRowBefore(3);

	const newRow: GSheets.Range = histSheet.getRange("A3:B3");
	
	newRow.setValues(values);
	newRow.setNumberFormats([["mm/dd/yyyy", "\"$\"#,##0.00"]]);

	if (isFriday) histSheet.getRange("B3").copyTo(histSheet.getRange("C3"));

	const rowCount: number = histSheet.getMaxRows();

	if (rowCount > 366) histSheet.deleteRows(367, rowCount-366);
}

function recordAllHistory(): void {
	const allSheets: GSheets.Sheet[] = ss.getSheets();

	const allHistSheets: GSheets.Sheet[] = allSheets.filter(
		function(entry) {
			return entry.getName().indexOf(sheetExtensionMap[SheetType.History]) > -1
		}
	);

	for (let histSheet of allHistSheets) {
		recordHistory(histSheet);
	}
}
