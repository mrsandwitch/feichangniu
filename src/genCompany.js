
function genCompTab() {
	// Load data
	var sheetName = "成本分析總表";
	var templateName = "範本";
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	dataSheet = ss.getSheetByName(sheetName);
	data = dataSheet.getDataRange().getValues();
	var costCompRowIx = 4;
	var costCompColIx = 8;
	var saleCompRowIx = 4;
	var saleCompColIx = 1;
	var costCompNameIx = [];
	var costCompName = [];
	var saleCompNameIx = [];
	var saleCompName = [];

	// Read cost company name
	for(var j=costCompRowIx;j<data[0].length;j++){
		var temp=data[j][costCompColIx];
		if (temp != "" && temp != "總計") {
			costCompName.push(temp);
			costCompNameIx.push([j+1, costCompColIx+1]);
		}
	}

	// Read sale company name
	for(var j=saleCompRowIx;j<data[0].length;j++){
		var temp=data[j][saleCompColIx];
		if (temp != "" && temp != "總計") {
			saleCompName.push(temp);
			saleCompNameIx.push([j+1, saleCompColIx+1]);
		}
	}

	// Generate Cost sheet tab
	for (var i = 0; i < costCompName.length; i++) {
		var templateSheet = ss.getSheetByName(templateName);
		var temp2 = ss.getSheetByName(costCompName[i]);
		var clonedSheet;
		if (typeof temp2 == 'undefined' || temp2 == null) {
			clonedSheet = templateSheet.copyTo(ss);
		} else {
			//continue;
			clonedSheet = ss.getSheetByName(costCompName[i]);
		}

		clonedSheet.setName(costCompName[i]);
		var clonedSheetId = clonedSheet.getSheetId();
		clonedSheet.getRange(2,1).setValue(costCompName[i]);

		var value = '=HYPERLINK(' + '"' + ss.getUrl() + "#gid=" + clonedSheetId + '",' + '"' + costCompName[i] + '")';
		dataSheet.getRange(costCompNameIx[0][0],costCompNameIx[0][1]).setFormula(value);
	}

	// Generate Sale sheet tab
	for (var i = 0; i < saleCompName.length; i++) {
		var templateSheet = ss.getSheetByName(templateName);
		var temp2 = ss.getSheetByName(saleCompName[i]);
		var clonedSheet;
		if (typeof temp2 == 'undefined' || temp2 == null) {
			clonedSheet = templateSheet.copyTo(ss);
		} else {
			//continue;
			clonedSheet = ss.getSheetByName(saleCompName[i]);
		}

		clonedSheet.setName(saleCompName[i]);
		var clonedSheetId = clonedSheet.getSheetId();
		clonedSheet.getRange(2,1).setValue(costCompName[i]);

		var value = '=HYPERLINK(' + '"' + ss.getUrl() + "#gid=" + clonedSheetId + '",' + '"' + saleCompName[i] + '")';
		dataSheet.getRange(saleCompNameIx[0][0],saleCompNameIx[0][1]).setFormula(value);
	}


	Logger.log("hello");


}
