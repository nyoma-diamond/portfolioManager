/// <reference path ="./gNamespaces.ts" />
type SheetMap = {[type: string]: GSheets.Sheet;};
type SheetNameMap = {[type: string]: string;};

enum SheetType {
	Main = "Main", 
	History = "History", 
	Utility = "Utility", 
	Sector = "Sector", 
	Stock = "Stock", 
	SP500 = "SP500"
}

const sheetExtensionMap: SheetNameMap = { };

sheetExtensionMap[SheetType.Main] = "";
sheetExtensionMap[SheetType.History] = " History";
sheetExtensionMap[SheetType.Utility] = " Utility";
sheetExtensionMap[SheetType.Sector] = " Sector Breakdown";
sheetExtensionMap[SheetType.Stock] = " Stock Breakdown";
sheetExtensionMap[SheetType.SP500] = " vs. S&P 500";

class Portfolio {
	public readonly name: string;
	public readonly sheetNameMap: SheetNameMap;
	public readonly sheetNames: string[];

	public constructor(name: string) {
		this.name = name;
		this.sheetNameMap = { };
		this.sheetNames = [ ];

		for (let type in sheetExtensionMap) {
			const sheetExtension: string = sheetExtensionMap[type];
			const sheetName: string = name + sheetExtension;
			this.sheetNameMap[type] = sheetName;
			this.sheetNames.push(sheetName);
		}
	}

	public getSheetMap(): SheetMap {
		let sheets: SheetMap = { };

		for (let type in sheetExtensionMap) {
			sheets[type] = ss.getSheetByName(this.sheetNameMap[type]);
		}
		return sheets;
	}

	public getSheetArray(): GSheets.Sheet[] {
		let sheetArr: GSheets.Sheet[] = [ ];

		for (let sheetName of this.sheetNames) {
			const sheet: GSheets.Sheet = ss.getSheetByName(sheetName);
			if (sheet) sheetArr.push(sheet);
		}
		return sheetArr;
	}
	
	public allExist(): boolean {
		const sheets: SheetMap = this.getSheetMap();

		for (let type in sheetExtensionMap) {
			if (!sheets[type]) return false;
		}

		return true;
	}

	public importantExist(): boolean {
		const sheets: SheetMap = this.getSheetMap();
		if (sheets[SheetType.Main] && sheets[SheetType.History] && sheets[SheetType.Utility]) return true;
		return false;
	}

	public anyExist(): boolean {
		const sheets: SheetMap = this.getSheetMap();

		for (let type in sheetExtensionMap) {
			if (sheets[type]) return true;
		}
		return false;
	}
}
