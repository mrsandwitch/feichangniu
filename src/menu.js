
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('非常牛指令')
	  .addItem('返回總表', 'returnMainSheet')
      .addSeparator()
      .addItem('新增公司', 'addCompany')
      .addSeparator()
      .addItem('test', 'testFunc')
      .addToUi();
}
//-----------------------------------------------/

function returnMainSheet() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var mainSheetName = "成本分析總表";
	var sheet = ss.getSheetByName(mainSheetName);
	ss.setActiveSheet(sheet).setActiveSelection("A1");
}
