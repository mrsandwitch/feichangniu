


function jumpToDetail() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var selRange = ss.getActiveRange();  
	var selRangeData = selRange.getValue().trim();  
	var currentSheet = ss.getActiveSheet();
	var currentSheetName = currentSheet.getName();

	if (currentSheetName == "2-各公司收入") {
		var jumpToSheet = ss.getSheetByName("3-收入單公司");
		jumpToSheet.getRange(2,1).setValues([[selRangeData]]);
		ss.setActiveSheet(jumpToSheet).setActiveSelection("A1");
	} else if (currentSheetName == "2-各公司費用") {
		var jumpToSheet = ss.getSheetByName("3-費用單公司");
		ss.setActiveSheet(jumpToSheet).setActiveSelection("A1");
		jumpToSheet.getRange(2,1).setValues([[selRangeData]]);
	} else if (currentSheetName == "3-收入單公司") {
		var companyName = currentSheet.getRange(2,1).getValues();
		var jumpToSheet = ss.getSheetByName("4-收入單合約");
		ss.setActiveSheet(jumpToSheet).setActiveSelection("A1");
		jumpToSheet.getRange(2,1).setValues([[companyName]]);
		jumpToSheet.getRange(4,1).setValues([[selRangeData]]);
	} else if (currentSheetName == "3-費用單公司") {
		var companyName = currentSheet.getRange(2,1).getValues();
		var jumpToSheet = ss.getSheetByName("4-費用單合約");
		ss.setActiveSheet(jumpToSheet).setActiveSelection("A1");
		jumpToSheet.getRange(2,1).setValues([[companyName]]);
		jumpToSheet.getRange(4,1).setValues([[selRangeData]]);
	}

}

function returnUpperSheet() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var currentSheet = ss.getActiveSheet();
	var currentSheetName = currentSheet.getName();
	var prevSheetName;

	switch (currentSheetName) {
		case "3-收入單公司":
			prevSheetName = "2-各公司收入";
			break;
		case "3-費用單公司":
			prevSheetName = "2-各公司費用";
			break;
		case "4-收入單合約":
			prevSheetName = "3-收入單公司";
			break;
		case "4-費用單合約":
			prevSheetName = "3-費用單公司";
			break;
		default:
			return;
			break;
	}

	var jumpToSheet = ss.getSheetByName(prevSheetName);
	ss.setActiveSheet(jumpToSheet).setActiveSelection("A1");
}
