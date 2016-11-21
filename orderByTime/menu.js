

function createMenu() {
	  var ui = SpreadsheetApp.getUi();
	  // Or DocumentApp or FormApp.
	  ui.createMenu('非常牛指令')
		  .addItem('返回上層', 'returnUpperSheet')
		  .addSeparator()
		  .addItem('進入細項', 'jumpToDetail')
		  .addSeparator()
		  .addToUi();
}
