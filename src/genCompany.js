
var costCompRowIx = 4;
var costCompColIx = 8;
var summarySheetName = "成本分析總表";
var costWholeMoneyColIx = 9;

function genCompTab() {
	// Load data
	var sheetName = "成本分析總表";
	var costTemplateName = "成本範本";
	var saleTemplateName = "銷貨範本";
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
	var wholeSheetName = [];
	var compSheetStartLoc = 5;

	// Read cost company name
	for(var j=costCompRowIx;j<data[0].length;j++){
		var temp=data[j][costCompColIx];
		if (temp != "" && temp != "總計") {
			costCompName.push(temp);
			wholeSheetName.push(temp);
			costCompNameIx.push([j+1, costCompColIx+1]);
		}
	}

	// Read sale company name
	for(var j=saleCompRowIx;j<data[0].length;j++){
		var temp=data[j][saleCompColIx];
		if (temp != "" && temp != "總計") {
			saleCompName.push(temp);
			wholeSheetName.push(temp);
			saleCompNameIx.push([j+1, saleCompColIx+1]);
		}
	}

	// Generate Sale sheet tab
	for (var i = 0; i < saleCompName.length; i++) {
		var templateSheet = ss.getSheetByName(saleTemplateName);
		var temp2 = ss.getSheetByName(saleCompName[i]);
		var clonedSheet;
		if (typeof temp2 == 'undefined' || temp2 == null) {
			clonedSheet = templateSheet.copyTo(ss);
		} else {
			continue;
			//clonedSheet = ss.getSheetByName(saleCompName[i]);
		}

		clonedSheet.setName(saleCompName[i]);
		var clonedSheetId = clonedSheet.getSheetId();
		clonedSheet.getRange(2,1).setValue(saleCompName[i]);
		clonedSheet.setTabColor("ffc0cb");

		var value = '=HYPERLINK(' + '"' + ss.getUrl() + "#gid=" + clonedSheetId + '",' + '"' + saleCompName[i] + '")';
		dataSheet.getRange(saleCompNameIx[i][0],saleCompNameIx[i][1]).setFormula(value);
	}

	// Generate Cost sheet tab
	for (var i = 0; i < costCompName.length; i++) {
		var templateSheet = ss.getSheetByName(costTemplateName);
		var temp2 = ss.getSheetByName(costCompName[i]);
		var clonedSheet;
		if (typeof temp2 == 'undefined' || temp2 == null) {
			clonedSheet = templateSheet.copyTo(ss);
		} else {
			continue;
			//clonedSheet = ss.getSheetByName(costCompName[i]);
		}

		clonedSheet.setName(costCompName[i]);
		var clonedSheetId = clonedSheet.getSheetId();
		clonedSheet.getRange(2,1).setValue(costCompName[i]);
		clonedSheet.setTabColor("ffff00");

		var value = '=HYPERLINK(' + '"' + ss.getUrl() + "#gid=" + clonedSheetId + '",' + '"' + costCompName[i] + '")';
		dataSheet.getRange(costCompNameIx[i][0],costCompNameIx[i][1]).setFormula(value);
	}

	// Rearrange the sheet order cost at the front
	var sheets = ss.getSheets();
	for (var i = compSheetStartLoc; i < sheets.length; i++) {
		if (sheets[i].getName() != wholeSheetName[i-compSheetStartLoc]) {
			var tempSheet = ss.getSheetByName(wholeSheetName[i-compSheetStartLoc]);
			ss.setActiveSheet(tempSheet);
			ss.moveActiveSheet(i+1);
		}
	}

	// Save cost company list for articulation
	var artiSheetName = "勾稽暫存";
	artiSheet = ss.getSheetByName(artiSheetName);
	var costCompNameTrans = [];
	for (var i = 0; i < costCompName.length; i++) {
		costCompNameTrans.push([costCompName[i]]);
	}
	artiSheet.getRange(1,1,costCompName.length,1).setValues(costCompNameTrans);
	
}

