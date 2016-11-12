
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('非常牛指令')
      .addItem('新增公司', 'addCompany')
      .addSeparator()
      .addItem('test', 'testFunc')
      .addToUi();
}
//-----------------------------------------------/

function testFunc() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("hello");
  var ssNew = ss.create("Finances", 50, 5);
  Logger.log(ssNew.getUrl());
}

function testFunc() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var companyName = "testComp";
	var companyParentFolder = "";
	//Logger.log("hello");
	//var ssNew = SpreadsheetApp.create("Finances", 50, 5);
	//Logger.log(ssNew.getUrl());
	//var companyFolder = DriveApp.createFolder(companyName);
	//var folders = DriveApp.getFolders();
	var id = ss.getId();
	var file = DriveApp.getFileById(id);
	var parentFolder = file.getParents()[0].getFoldersByName(companyParentFolder);
	Logger.log(parentFolder.getName());
	
//	while (folders.hasNext()) {
//		var folder = folders.next();
//		Logger.log(folder.getName());
//	}
}

function testFunc() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var companyName = "testComp";
	var companyParentFolder = "各廠商";
	var companyFolder;
	var id = ss.getId();
	var file = DriveApp.getFileById(id);
	var parentFolderIter = file.getParents().next()
		.getFoldersByName(companyParentFolder);
	var parentFolder = parentFolderIter.next();

	// Prevent duplicate folder with same name
	var companyFldIt = parentFolder.getFoldersByName(companyName);
	if (!companyFldIt.hasNext()) {
		companyFolder = parentFolder.createFolder(companyName);
	} else {
		companyFolder = companyFldIt.next();
	}

}

function testFunc2() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	ss.copy("Copy of " + ss.getName());

}

function create_ss_in_folder(folder, name) {
	var ss = SpreadsheetApp.create(name);
	var id = ss.getId();
	var file = DriveApp.getFileById(id);
	var folder = get_folder_by_name(folder);
	folder.addFile(file);
	DriveApp.getRootFolder().removeFile(file);
	return id;
}


