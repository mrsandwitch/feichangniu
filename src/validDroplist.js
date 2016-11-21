
function genValidDroplist() {
	// Load data
	var paramSheetName = "param";
	var costSheetLevel3 = "3-費用單公司";
	var incomeSheetLevel3 = "3-收入單公司";
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	paramSheet = ss.getSheetByName(paramSheetName);
	paramData = paramSheet.getDataRange().getValues();
	var costCompRowIx = 3;
	var incomeCompRowIx = 6;
	var costCompNameList = [];
	var incomeCompNameList = [];

	for (var i = 1; i < paramData[0].length; i++) {
		if (paramData[costCompRowIx-1][i] != "") {
			costCompNameList.push(paramData[costCompRowIx-1][i]);
		}
	}

	for (var i = 1; i < paramData[0].length; i++) {
		if (paramData[incomeCompRowIx-1][i] != "") {
			incomeCompNameList.push(paramData[incomeCompRowIx-1][i]);
		}
	}

	var costDroplistCell = ss.getSheetByName(costSheetLevel3).getRange('A2');
	var rule = SpreadsheetApp.newDataValidation().requireValueInList(costCompNameList).build();
	costDroplistCell.setDataValidation(rule);

	var incomeDroplistCell = ss.getSheetByName(incomeSheetLevel3).getRange('A2');
	var rule = SpreadsheetApp.newDataValidation().requireValueInList(incomeCompNameList).build();
	incomeDroplistCell.setDataValidation(rule);

}

