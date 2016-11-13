
function returnMainSheet() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var mainSheetName = "成本分析總表";
	var sheet = ss.getSheetByName(mainSheetName);
	ss.setActiveSheet(sheet).setActiveSelection("A1");
}

function jumpToCompSheet() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var selRange = ss.getActiveRange();  
	var selRangeData = selRange.getValue();  
	var sheet = ss.getSheetByName(selRangeData);
	ss.setActiveSheet(sheet).setActiveSelection("A1");
}
