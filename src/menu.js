
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

function onOpen() {
	  var ui = SpreadsheetApp.getUi();
	  // Or DocumentApp or FormApp.
	  ui.createMenu('非常牛指令')
		  .addItem('返回總表', 'returnMainSheet')
		  .addSeparator()
		  .addItem('進入公司細項', 'jumpToCompSheet')
		  .addSeparator()
		  .addItem('勾稽', 'allArticulaton')
		  .addSeparator()
		  .addItem('產生公司分頁', 'genCompTab')
		  .addToUi();
}
