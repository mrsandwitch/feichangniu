
function compareAndSave(companyId, invoiceId, invoiceMoney) {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var artiSheetName = "勾稽暫存";
	artiSheet = ss.getSheetByName(artiSheetName);
	savedData = artiSheet.getDataRange().getValues();

	for (var i = 0; i < companyId.length; i++) {
		var entry = String(invoiceId[i]) + ":" + String(invoiceMoney[i]);
		if (companyId[i] == "" || companyId[i] < 1) {
			continue;
		}
		if (savedData[companyId[i]-1].indexOf(entry) === -1) {
			var saveIx = 1;
			for (var j = 1; j < savedData[companyId[i]-1].length; j++) {
				if (savedData[companyId[i]-1][j] == "") {
					saveIx = j;
					break;
				} else {
					saveIx += 1;
				}
			}
			savedData[companyId[i]-1][saveIx] = entry;
			artiSheet.getRange(companyId[i], saveIx + 1).setValues([[entry]]);
		} else {
			Logger.log(entry+" exist");
		}
	}
}
	 
function onEdit(e){
	// Set a comment on the edited cell to indicate when it was changed.
	var range = e.range;
	var changeUpperBound = 44;
	var changeLeftBound = 3;
	var changeRightBound = 9;
	var companyIdIx = 4;
	var invoiceColIx = 5;
	var invoiceMoneyIx = 8;
	var invoiceRowIx = 45;
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var currentSheet;
	var currentSheetName;
		
	if (range.getRow() > changeUpperBound &&
			range.getColumn() > changeLeftBound &&
			range.getColumn() < changeRightBound) {

		currentSheet = ss.getActiveSheet();
		currentSheetName = currentSheet.getName();
		var maxRow = currentSheet.getDataRange().getNumRows();
		//ss.toast(maxRow);
		var companyId = currentSheet
			.getRange(invoiceRowIx, companyIdIx, maxRow - invoiceRowIx + 1, 1)
			.getValues();
		var invoiceId = currentSheet
			.getRange(invoiceRowIx, invoiceColIx, maxRow - invoiceRowIx + 1, 1)
			.getValues();
		var invoiceMoney = currentSheet
			.getRange(invoiceRowIx, invoiceMoneyIx, maxRow - invoiceRowIx + 1, 1)
			.getValues();
		compareAndSave(companyId, invoiceId, invoiceMoney);

	} else {
		//ss.toast("out of the bound");
	}
}

function allArticulaton() {
	var changeUpperBound = 44;
	var changeLeftBound = 3;
	var changeRightBound = 9;
	var companyIdIx = 4;
	var invoiceColIx = 5;
	var invoiceMoneyIx = 8;
	var invoiceRowIx = 45;
	var compSheetStartLoc = 5;
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var currentSheet;
	var currentSheetName;

	var wholeInvoiceMoney = {};

	var sheets = ss.getSheets();
	for (var i = compSheetStartLoc + 1; i < sheets.length; i++) {
		var currentSheet = sheets[i];
		currentSheetName = currentSheet.getName();
		var maxRow = currentSheet.getDataRange().getNumRows();

		if (maxRow < invoiceRowIx) {
			continue;
		}

		var companyId = currentSheet
			.getRange(invoiceRowIx, companyIdIx, maxRow - invoiceRowIx + 1, 1)
			.getValues();
		var invoiceId = currentSheet
			.getRange(invoiceRowIx, invoiceColIx, maxRow - invoiceRowIx + 1, 1)
			.getValues();
		var invoiceMoney = currentSheet
			.getRange(invoiceRowIx, invoiceMoneyIx, maxRow - invoiceRowIx + 1, 1)
			.getValues();

		

		for (var j = 0; j < invoiceMoney.length; j++) {
			if (companyId[j] == "") {
				continue;
			}
			if (typeof wholeInvoiceMoney[String(companyId[j])] == 'undefined') {
				wholeInvoiceMoney[String(companyId[j])] = [invoiceMoney[j]];
			} else {
				wholeInvoiceMoney[String(companyId[j])].push(invoiceMoney[j]);
			}
		}
	}

	// Set color
	var summarySheet = ss.getSheetByName(summarySheetName);
	maxRow = summarySheet.getDataRange().getNumRows();
	costWholeMoney = summarySheet.getRange(costCompRowIx + 1, costWholeMoneyColIx + 1, maxRow - costWholeMoneyColIx, 1).getValues();
	Logger.log(costWholeMoney);

	for (var i = 0; i < costWholeMoney.length; i++) {
		if (costWholeMoney[i] == "") {
			continue;
		}
		var sum = 0;
		for(var j = 0; j < wholeInvoiceMoney[String(i+1)].length; j++) {
			sum += parseInt(wholeInvoiceMoney[String(i+1)][j]);
		}
		Logger.log(sum);
		Logger.log(Math.abs(parseInt(costWholeMoney[i]) - sum));
		if (Math.abs(parseInt(costWholeMoney[i]) - sum) > 0.5) {
			summarySheet.getRange(costCompRowIx + 1 + i, costWholeMoneyColIx + 1)
				.setBackground("#ffc0cb");
		} else {
			summarySheet.getRange(costCompRowIx + 1 + i, costWholeMoneyColIx + 1)
				.setBackground("#ffffff");
		}

	}

}



