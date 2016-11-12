

function returnMainSheet() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var mainSheetName = "成本分析總表";
	var sheet = ss.getSheetByName(mainSheetName);
	ss.setActiveSheet(sheet).setActiveSelection("A1");
}
