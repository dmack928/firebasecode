//opening function of the embedded fweb app when pages loads
function doGet(e) {
  Logger.log(Utilities.jsonStringify(e));
  var myUser = Session.getActiveUser().getEmail();
  var myPage = e.parameter.id;
  Logger.log("myPage = " + myPage);
  Logger.log("myUser = " + myUser);
  //var userProperties = PropertiesService.getUserProperties();
  //userProperties.setProperty('MY_LOC', myPage);
  //userProperties.setProperty('USER', myUser);
  if(e.parameter.docid != undefined || e.parameter.docid != null){
    var docNum = e.parameter.docid;}
  else {var docNum = "";}

  var userInfo = '<script> \
(function() { \
    window.varUserInfo = { \
        myUser: ["' + myUser + '", "0"], \
        myLoc:"' + myPage + '",\
        myDoc:"' + docNum + '",\
    }; \
})();\
</script>';

  var initProjectFx = '<script> \
projectFunction.initAll(); </script>';
  var initMyFx = '<script> \
myFunction.initAll(); </script>';


	switch (myPage) {
        case "EP":
		case "NC":
        case "LT":
			var html = HtmlService.createTemplateFromFile('projects').evaluate().getContent();
			return HtmlService.createTemplate(html + userInfo + initProjectFx).evaluate();
			break;
//		case "PC":
//			var html = HtmlService.createTemplateFromFile('purchasing_page').evaluate().getContent();
//			return HtmlService.createTemplate(html + userInfo + initMyFx).evaluate();
//			break;
		case "RQ":
			var html = HtmlService.createTemplateFromFile('labrequest_page').evaluate().getContent();
			return HtmlService.createTemplate(html + userInfo + initMyFx).evaluate();
			break;
		default:
			Logger.log("error loading HTML template doGet");
			break;
	}
}

//function to recall specific document links to access
function globalVariables(){ var variables = {EP_NC_ss: '16Klv-BEMofgTTFsxcEXxDScEMuKZ_tXFHA4P5_xsOt8', //spreadsheet id for lab req and report database
                                             LT_ss: "1gE-CpIFKBW2GPFy5LAQbeiTSbNEW5JXElGfE1E7nntI",  //spreadsheet id for long term project database
                                             PC_ss: "1WP-rQNFGert8KX2gfZPYBTx4m3saGWqqbP397l5hpak", //spreadsheet id for purchasing database
//                                             CUSTDB_ss: "1H0yfMxQnINeCxvW-AKBas7s9aW4r2RmAv3TustSQDBw", //spreadsheet id for customer database
                                             CUSTDB_ss: "17OVsJF6MEzSE5lXwmF1_iY44P0T2lbFDJnl3uQKmVwg", //spreadsheet id for customer database
//                                             ROOTFLDR: "0BxlkPxuAhnwRYlp0TG5pbFJCekE", //location of the root folder in Google Drive
                                             ROOTFLDR: "0BxlkPxuAhnwRaV90Tk9WQ1JrWk0", //location of the root folder in Google Drive
                                             RQSTtemplate: '1An2fAV9mkdOSda7RsVa1EuhZWa6hRb4QLdYooOgYiDA', //google doc template for lab request document
                                             RPRTtemplate: '1g5XS_PKhGAHz3bNA6gw0k3uRnMSr64WbAVsqLzvO2lw', //google doc template for lab report document
                                             commonDB: "1xsG_0E8i_hmVvvGWtZNsOtyYFwKcao1xJZj61zOD4RA",//the coomon databse of names and market lists
                                             }; return variables; }
//Upload folder: https://drive.google.com/drive/u/0/folders/0BxlkPxuAhnwRaV90Tk9WQ1JrWk0
//LabReqDB: https://drive.google.com/open?id=16Klv-BEMofgTTFsxcEXxDScEMuKZ_tXFHA4P5_xsOt8
//Root folder: https://drive.google.com/open?id=0BxlkPxuAhnwRYlp0TG5pbFJCekE
//Lab Req Template: https://drive.google.com/open?id=1An2fAV9mkdOSda7RsVa1EuhZWa6hRb4QLdYooOgYiDA
//Lab Reprt Template: https://drive.google.com/open?id=1g5XS_PKhGAHz3bNA6gw0k3uRnMSr64WbAVsqLzvO2lw
//customer db: https://drive.google.com/open?id=1H0yfMxQnINeCxvW-AKBas7s9aW4r2RmAv3TustSQDBw
//purch db:https://drive.google.com/open?id=1WP-rQNFGert8KX2gfZPYBTx4m3saGWqqbP397l5hpak
//longterm: https://drive.google.com/open?id=1gE-CpIFKBW2GPFy5LAQbeiTSbNEW5JXElGfE1E7nntI

//function that opens and read the common database to pass data object to client
function collectDB(args) {
  var ss = SpreadsheetApp.openById(globalVariables().commonDB);
  var sheet = ss.getSheetByName(args[0]);
  var list = [];

  if(args[1]==null){
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var myRange = sheet.getRange(2, 1, lastRow, lastCol);
  var _list = myRange.getValues();
  Logger.log("_list: " + _list);
  for(var i=0; i<_list.length; i++){
  list.push([String(_list[i][0]),_list[i][1]])}}

  if(args[1]=="sales"){
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var myRange = sheet.getRange(2, 1, lastRow, lastCol);
  var _list = myRange.getValues();
  Logger.log("_list: " + _list);
  for(var i=0; i<_list.length; i++){
  if(_list[i][5]!=0){
  list.push([String(_list[i][0]),_list[i][1],_list[i][2]])}

  }}
  if(args[1]=="lab"){
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var myRange = sheet.getRange(2, 1, lastRow, lastCol);
  var _list = myRange.getValues();
  Logger.log("_list: " + _list);
  for(var i=0; i<_list.length; i++){
  if(_list[i][6]!=0){
  list.push([String(_list[i][0]),_list[i][1],_list[i][2]])}

  }}
  if(args[1]=="contacts"){
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var myRange = sheet.getRange(2, 1, lastRow, lastCol);
  var _list = myRange.getValues();
  Logger.log("_list: " + _list);
  for(var i=0; i<_list.length; i++){
  if(_list[i][8]!=0){
  list.push(["0",_list[i][1],_list[i][2]])}

  }}
  if(args[1]=="levels"){
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var myRange = sheet.getRange(2, 1, lastRow, lastCol);
  var _list = myRange.getValues();
  Logger.log("_list: " + _list);
  for(var i=0; i<_list.length; i++){
  if(_list[i][7]!=0){
  list.push([_list[i][2],String(_list[i][7])])}

  }}
  Logger.log("list" + list);
  return list;
}

function namedata1(args){
  var mkt  = collectDB(["market",null]);
  var sales  = collectDB(["users","sales"]);
  var _lab = collectDB(["users","lab"]);
  var _contacts = collectDB(["users","contacts"]);
  var _levels = collectDB(["users","levels"]);
  var orcoData = {
  market_segment: mkt,
  salesgroup: sales,
  labgroup: _lab,
  contacts: _contacts,
  userLevels: _levels,
};
  var client = JSON.parse(args);
  Logger.log("client = " + client[0]);
  if(client[0] == true){Logger.log("TRUE in namedata!");return orcoData;}
  else if(client[0] == false){Logger.log("FALSE in namedata!"); return JSON.stringify([orcoData, client[1], client[2]])};
};

//function namedata(args){ var orcoData = {//'args' must be an array [bool, "",""]
//  market_segment: [
//["10","Fabric Apparel"],
//["20","Fabric Non Apparel"],
//["21","Agriculture"],
//["22","Anodizing"],
//["23","Coatings"],
//["24","Detergents & Cleaning Products"],
//["25","Inks"],
//["28","Oils, Waxes, Lubricants & Plastics"],
//["30","Dealers"],
//["38","Construction"],
//["99","Miscellaneous"],
//["666","Other: Distinct from Misc."]],
//
//  salesgroup: [
//["1","Barry Brady","bbrady"],
//["2","Randy Yorston","ryorston"],
//["3","Mike Sylvia","msylvia"],
//["4","Bob Rossi","rrossi"],
//["6","Matt Doyle","mdoyle"],
//["7","Greg Gormley","ggormley"],
//["8","Vince Hankins","vhankins"],
//["9","Anne McClean","bbrady"],
//["10","William Huckaby","whuckaby"],
//["11","Benji Bagwell", "bbagwell"],
//["999","ODP House","mdoyle"],
//  ],
//
//  labgroup: [
//["1","Bob Richardson","brichardson"],
//["2","Mark Axile","maxile"],
//["3","Chris Gustafson","cgustafson"],
//["6","Christine Leal","colorLab"],
//["7","Peter Guilbault","pguilbault"],
//["8","Desmek Hall","sampledept"],
//["4","Derek Williams","dwilliams"],
//["5","John Neves","jneves"],
//["9","Mike Sylvia","msylvia"],
//["11","Carroll Dickerson","cdickerson"],
//["12","Rick Little","rlittle"],
////["13","Nirmala Chidurala","concordLab"],
//["14","Nancy Mitchem","concordLab"],
//["15","Rafael Segura","concordLab"],
//["16", "Kaylin Sutton", "ksutton"],
//["17", "Donnie Plyler", "dplyler"],
//  ],
//
//  contacts: [
//["0","Bob Richardson","brichardson"],
//["0","Mark Axile","maxile"],
//["0","Chris Gustafson","cgustafson"],
//["0","Derek Williams","dwilliams"],
//["0","John Neves","jneves"],
//["0","Carroll Dickerson","cdickerson"],
//["0","Rick Little","rlittle"],
//["0","Kaylin Sutton","ksutton"],
//["0","Barry Brady","bbrady"],
//["0","Matt Doyle","mdoyle"],
//["0","Donnie Plyler","dplyler"],
//    ],
//  //userLevels: 10: super user that can edit, view, control and transfer projects
//  //             9: users that can view and edit all fields (cannot transfer)
//  //             8: users that can view and edit formulas, results, and actions only
//  //             7: users that can view only formula, results, actions
//  //             0: lowly users that only view results, actions NO formula view, NO edit - you know...salespeople
//  userLevels: [
//  ["dwilliams", "10"],
//  ["brichardson", "9"],
//  ["maxile", "9"],
//  ["cgustafson", "9"],
//  ["cdickerson", "10"],
//  ["jneves", "9"],
//  ["colorlab", "8"],
//  ["concordlab", "8"],
//  ["cleal", "8"],
//  ["lbouchard", "8"],
//  ["sampledept", "7"],
//  ["msylvia", "8"],
//  ["cturner", "7"],
//  ["mdoyle", "7"],
//  ["bbrady", "7"],
//  ["jdamelio", "7"],
//  ["ryorston", "7"],
//  ["dewey", "0"],
//  ["rlittle","9"],
//  ["ksutton", "9"],
//  ['smarkantonakis', "7"],
//  ['dplyler', "9"],
//  ]
//};
//
//  var client = JSON.parse(args);
//  Logger.log("client = " + client[0]);
//
//  //TRUE if function called withing server script within "Code.gs", FALSE when called from client script
//  if(client[0] == true){Logger.log("TRUE in namedata!");return orcoData;}
//  else if(client[0] == false){Logger.log("FALSE in namedata!"); return JSON.stringify([orcoData, client[1], client[2]])};
//  }

function include(filename) {
	return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function moveProject(infoForMove){
//  var userProperties = PropertiesService.getUserProperties();
//  var myPlace = userProperties.getProperty('MY_LOC');
  var objInfo = JSON.parse(infoForMove);
  Logger.log('moveProject() ID = ' + objInfo.project);
  Logger.log('moveProject to = ' + objInfo.moveTo);

  var SHEET = globalVariables();
  switch(objInfo.moveFrom){
    case 'PC':
      Logger.log("objInfo.moveFrom " + objInfo.moveFrom)
      var sourceId = SpreadsheetApp.openById(SHEET.PC_ss);
      break;
    case 'LT':
      Logger.log("objInfo.moveFrom " + objInfo.moveFrom)
      var sourceId = SpreadsheetApp.openById(SHEET.LT_ss);
      break;
    default:
      Logger.log("objInfo.moveFrom " + objInfo.moveFrom)
      var sourceId = SpreadsheetApp.openById(SHEET.EP_NC_ss);
      break;
  }
  switch(objInfo.moveTo){
    case 'PC':
      Logger.log("objInfo.moveTo " + objInfo.moveTo)
      var targetId = SpreadsheetApp.openById(SHEET.PC_ss);
      break;
    case 'LT':
      Logger.log("objInfo.moveTo " + objInfo.moveTo)
      Logger.log("SHEET.LT_ss " + SHEET.LT_ss)
      var targetId = SpreadsheetApp.openById(SHEET.LT_ss);
      break;
    default:
      Logger.log("objInfo.moveTo " + objInfo.moveTo)
      var targetId = SpreadsheetApp.openById(SHEET.EP_NC_ss);
      break;
  }


	var lock = LockService.getScriptLock();
	lock.waitLock(30000);
    var source = sourceId.getSheetByName(objInfo.moveFrom);
    var target = targetId.getSheetByName(objInfo.moveTo);
	source.activate();
    source.showSheet();
	target.activate();
    var rowFrom = 0;
	var lrObjrng = source.getRange(2, 1, source.getLastRow() - 1, 1).getValues();
	for (var i = 0; i < lrObjrng.length; i++) {
		if (lrObjrng[i] == objInfo.project) {
            rowFrom = i + 2;
			var sourceValues = source.getRange(rowFrom, 1, 1, 3).getValues();
			Logger.log("Range selected");
		};
	}

  target.showSheet();
  var row = target.getLastRow() + 1;
  var targetRange = target.getRange(row,1,1,3);
//  sourceRange.copyTo(targetRange);
  targetRange.setValues(sourceValues);
  source.deleteRow(rowFrom)


  SpreadsheetApp.flush();
	lock.releaseLock();


  var myResponse = JSON.stringify("Success!");
  return myResponse;
}


function getprojects(fromClient) {
  var lock = LockService.getScriptLock();
//  var userProperties = PropertiesService.getUserProperties();
//  var myPlace = userProperties.getProperty('MY_LOC');
//  var myUser = userProperties.getProperty('USER');
  var myPlace = JSON.parse(fromClient);
  Logger.log("myPlace is " + myPlace);
  var SHEET = globalVariables();
  Logger.log(SHEET);
  switch(myPlace){
    case 'PC':
      var ss = SpreadsheetApp.openById(SHEET.PC_ss);
      break;
    case 'LT':
      var ss = SpreadsheetApp.openById(SHEET.LT_ss);
      break;
    default:
      var ss = SpreadsheetApp.openById(SHEET.EP_NC_ss);
      break;
  }
	lock.waitLock(30000);
    Logger.log("where am i? " + myPlace);
	var sheet = ss.getSheetByName(myPlace);
	sheet.activate();
	var lrObjrng = sheet.getRange(2, 1, ss.getLastRow() - 1, 3).getValues();
	lock.releaseLock();
//    var infoToClient = {user: myUser, loc: myPlace, sheetRng: lrObjrng,};
    var infoToClient = {sheetRng: lrObjrng,};
	return JSON.stringify(infoToClient);
}

//function that saves any new data into google sheet database
function update_projects(myObj) {
myObj = JSON.parse(myObj);

//var userProperties = PropertiesService.getUserProperties();
//  var myPlace = userProperties.getProperty('MY_LOC');
  var myPlace = myObj.loc;
  var SHEET = globalVariables();
  switch(myPlace){
    case 'PC':
      var ss = SpreadsheetApp.openById(SHEET.PC_ss);
      break;
    case 'LT':
      var ss = SpreadsheetApp.openById(SHEET.LT_ss);
      break;
    default:
      var ss = SpreadsheetApp.openById(SHEET.EP_NC_ss);
      break;
  }
	var lock = LockService.getScriptLock();
//	myObj = JSON.parse(myObj);
	lock.waitLock(30000);
	var sheet = ss.getSheetByName(myPlace);
	sheet.activate();
	var lrObjrng = sheet.getRange(2, 1, ss.getLastRow() - 1, 1).getValues();
	Logger.log("myObj: " + myObj.objPJLIST);
	for (var i = 0; i < lrObjrng.length; i++) {
		if (lrObjrng[i] == myObj.objPJLIST[0]) {
			var myRng = sheet.getRange(i + 2, 3);
			Logger.log("Range set");
			myRng.setValue(JSON.stringify(myObj.objPJLIST[2]));
		};
	}
	SpreadsheetApp.flush();
	lock.releaseLock();
	return JSON.stringify(sheet.getRange(2, 1, ss.getLastRow() - 1, 3).getValues());
}

//following script uploadFileToDrive from below sourced - modified to fit this application
//https://stackoverflow.com/questions/31126181/uploading-multiple-files-to-google-drive-with-google-app-script/34777747#34777747
function uploadFileToDrive(base64Data, fileName, myFolder) {
	//Logger.log(base64Data);
	Logger.log(fileName);
	Logger.log(myFolder);
	var lock = LockService.getScriptLock();
	lock.waitLock(30000);
	try {
		var splitBase = base64Data.split(','),
			type = splitBase[0].split(';')[0].replace('data:', '');
		var byteCharacters = Utilities.base64Decode(splitBase[1]);
		var ss = Utilities.newBlob(byteCharacters, type);
		ss.setName(fileName);
		var fileinfo = {
			filename: "",
			fileurl: "",
			folder: "",
			folderurl: "",
			folderid: "",
		}

		var rootfolder = DriveApp.getFolderById(globalVariables().ROOTFLDR);
		var folder;
		var folders = rootfolder.getFoldersByName(myFolder);
		Logger.log(folders);
		if (folders.hasNext()) {
			folder = folders.next();
		} else {
			folder = rootfolder.createFolder(myFolder);
		}
		var file = folder.createFile(ss);
		fileinfo.fileurl = file.getUrl();
		fileinfo.filename = file.getName();
		fileinfo.folder = folder.getName();
		fileinfo.folderurl = folder.getUrl();
		fileinfo.folderid = folder.getId();
		Logger.log(JSON.stringify(fileinfo));
		return JSON.stringify(fileinfo);
	}
	//    return true;}
	catch (e) {
		Logger.log(e);
		Logger.log('Error: ' + e.toString());
		return 999;
	}
	lock.releaseLock();
}

//function that parses time and date
function formatdate(datestr) {
	var retDate = {
		cal: datestr.split("T")[0],
		time: datestr.split(".")[0].split("T")[1],
	}
	return retDate;
}

//function that opens and read the customer database to pass data object to client
function compdbacc() {
  //var userProperties = PropertiesService.getUserProperties();
  //var myUser = userProperties.getProperty('USERID');
  var ss = SpreadsheetApp.openById(globalVariables().CUSTDB_ss);
  var lock = LockService.getScriptLock();
//  var sheet = ss.getSheetByName("CUSTDB");
  var sheet = ss.getSheetByName("custdb");
  var lastRow = sheet.getLastRow();
  var myRange = sheet.getRange(1, 1, lastRow);
  var complist = myRange.getValues();
  SpreadsheetApp.flush();
  lock.releaseLock();
  Logger.log(complist);
//  return complist;
//  var infoToClient = {user: myUser, compRng: complist,};
  var infoToClient = {compRng: complist,};
  return JSON.stringify(infoToClient);
}

//https://drive.google.com/open?id=16Klv-BEMofgTTFsxcEXxDScEMuKZ_tXFHA4P5_xsOt8 is LabProjectDB is the Google Sheets doc
//function that creates a new lab requeistion number
function getNewReqNo(locale) {
  var SHEET = globalVariables();
  Logger.log(locale);
  var ss = SpreadsheetApp.openById(SHEET.EP_NC_ss);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var range, count;
  var sheet = ss.getSheetByName("count");
  sheet.activate();
  if (locale == "EP") {
    range = sheet.getRange(1, 1);
    count = range.getValue();
    range.setValue(count + 1);
  }
  if (locale == "NC") {
    range = sheet.getRange(2, 1);
    count = range.getValue();
    range.setValue(count + 1);
  }
  var lrno = locale + "-" + count;
  Logger.log(lrno);
  var folderinfo = createProjectFolder(lrno);
  var info = {
    folderid: folderinfo.folderid,
    folderurl: folderinfo.folderurl,
    lrno: lrno,
  }
  lock.releaseLock();
  return JSON.stringify(info);
}

//function to pass a client object to the server function to print a lab report
function processReport(dataObject){
  printDOC(dataObject, false);
  Logger.log("processReport");
  return true;
  }

//function to pass a client object to the server function and to save lab request data creating a new record for lab request
function processLabReq(inOBJ) {
  var LROBJ = [];
  LROBJ[0] = 0;
  LROBJ[1] = JSON.parse(inOBJ);
  LROBJ[0] = LROBJ[1].lrno;
  Logger.log(inOBJ);
  var ss = SpreadsheetApp.openById(globalVariables().EP_NC_ss);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var range, count;
  var sheet; // = ss.getSheetByName("count");
  switch(LROBJ[1].location){
    case "EP":
    sheet = ss.getSheetByName("EP");
    sheet.activate();
      break;
    case "NC":
    sheet = ss.getSheetByName("NC");
    sheet.activate();
      break;
    default:
      Logger.log("ERROR in processLabReq location switch-case");
      break;
  }
  sheet.showSheet();
  var row = sheet.getLastRow() + 1;
  range = sheet.getRange(row, 1);
  range.setValue(LROBJ[1].lrno);
  range = sheet.getRange(row, 2);
  range.setValue(JSON.stringify(LROBJ[1]));
  Logger.log("LROBJ = " + LROBJ);
 LROBJ[2] = {
    lrno: LROBJ[1].lrno,
    status: "1",
   owner: {
     index:"0",
     name:"Pending",
     email:"",
   },
    reviewed: {
      by: "",
      rev_date: "",
    },
    lastupdate: {
      d: "",
      t: "",
      byUser: "",
    },
    results: "",
    action: "",
    updates: {
    last: "",
    logged: "",
    },
    docs: "",
    key: "",
    stamp: {
      opened: "",
      closed: ""
    },
    summary: "",
   purchasing: {
     itemcount:"0",
     data:""
   }
  }
  range = sheet.getRange(row, 3);
  range.setValue(JSON.stringify(LROBJ[2]));
  SpreadsheetApp.flush();
  Logger.log("Finished processing the form....");
  lock.releaseLock();
  Logger.log("process form");
  var custdata = [LROBJ[1].company_name, LROBJ[1].company_addr1, LROBJ[1].company_addr2, LROBJ[1].customer_city, LROBJ[1].customer_state, LROBJ[1].customer_zip_code, LROBJ[1].customer_country, LROBJ[1].customer_name, LROBJ[1].customer_email, LROBJ[1].customer_phone];
  updatecustdb(LROBJ[1].database_index, LROBJ[1].mod_cust_info, custdata);
  Logger.log("updatecustomerdb");
  LROBJ = JSON.stringify(LROBJ);
  var prnInfo = {user:LROBJ[1].user , loc:LROBJ[1].location , object:LROBJ};
  printDOC(prnInfo, true);
  Logger.log("printcopy");
  return 200;
}

//function that collects selcted customer record
function gather_comp_info(compRow) {
  var ss = SpreadsheetApp.openById(globalVariables().CUSTDB_ss);
  var lock = LockService.getScriptLock();
//  var sheet = ss.getSheetByName("CUSTDB");
  var sheet = ss.getSheetByName("custdb");
  var myRange = sheet.getRange(compRow, 1);
  var comp_info = myRange.getValue();
  SpreadsheetApp.flush();
  lock.releaseLock();
  return comp_info;
}

//following script uploadFileToDrive from below sourced - modified to fit this application
//https://stackoverflow.com/questions/31126181/uploading-multiple-files-to-google-drive-with-google-app-script/34777747#34777747
function createProjectFolder(myfoldername) {
  var rootfolder = DriveApp.getFolderById(globalVariables().ROOTFLDR);
  var folder;
  var folders = rootfolder.getFoldersByName(myfoldername);
  Logger.log(folders);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = rootfolder.createFolder(myfoldername);
  }
  var projectloc = {
    folderid: folder.getId(),
    folderurl: folder.getUrl(),
  }
  return projectloc;
}

//function that updates the customer database with a new customer record or modify a custoemr record
function updatecustdb(isnew, ismod, custdata) {
  var ss = SpreadsheetApp.openById(globalVariables().CUSTDB_ss);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
//  var sheet = ss.getSheetByName("CUSTDB");
  var sheet = ss.getSheetByName("custdb");
  var myRange, str;
  isnew = parseInt(isnew, 10);
  str = "Nothing Updated in CUSTDB ismod=" + ismod + " isnew= " + parseInt(isnew) + " custdata=" + custdata;
  var newdata = {
    customer: custdata[0],
    addr1: custdata[1],
    addr2: custdata[2],
    city: custdata[3],
    state: custdata[4],
    zipcode: custdata[5],
    country: custdata[6],
    attn: custdata[7],
    email: custdata[8],
    phone: custdata[9],
    code: "",
    fob: "",
  }
  if (isnew == 0 && !ismod) {
    var newRow = sheet.getLastRow() + 1;
    myRange = sheet.getRange(newRow, 1);
    myRange.setValue(JSON.stringify(newdata));
    myRange = sheet.getRange(1, 1, newRow);
    myRange.sort(1);
    str = "new entry made....";
  }
  if (ismod && isnew > 0) {
    myRange = sheet.getRange(isnew, 1);
    myRange.setValue(JSON.stringify(newdata));
    str = "updated an customer record....";
  }
  SpreadsheetApp.flush();
  lock.releaseLock();
  return 400;
}

//https://stackoverflow.com/questions/33205269/google-doc-script-search-and-replace-text-string-and-change-font-e-g-boldfac

function boldfaceText(findMe, theBody) {

  // put to boldface the argument
  var foundElement = theBody.findText(findMe);

  while (foundElement != null) {
    // Get the text object from the element
    var foundText = foundElement.getElement().asText();

    // Where in the Element is the found text?
    var start = foundElement.getStartOffset();
    var end = foundElement.getEndOffsetInclusive();

    // Change the background color to yellow
    foundText.setBold(start, end, true);

    // Find the next match
    foundElement = theBody.findText(findMe, foundElement);
  }

}

//server side function to print a lab request (isLR==TRUE) or a lab report (isLR==FALSE)
function printDOC(inOBJ, isLR) {
  //Logger.log(inOBJ);
  //Logger.log('Printing....'+JSON.parse(inOBJ));
  if(!isLR){
    var myObj = JSON.parse(inOBJ);
    Logger.log("myObj !isLR = " + myObj);
     var data = myObj.object;}
  else if(isLR){
    var myObj = inOBJ;
    Logger.log("myObj.object isLR = " + myObj.object);
    Logger.log("myObj.object isLR = " + myObj.object.length);
    var data = JSON.parse(myObj.object);
  }

    // var data = myObj.object;
    Logger.log("data = " + data);
    Logger.log("data[0] = " + data[0]);
    Logger.log("data[1] = " + data[1]);
    Logger.log("data[2] = " + data[2]);
    var myPlace = myObj.loc;
    var myUser = myObj.user;
//  Logger.log("myUser in printDOC: " + myUser + " and myPlace: " + myPlace);
  switch (data[1].location){
    case "EP"://https://script.google.com/a/organicdye.com/macros/s/AKfycbzqPtSgEc_KLnpMSX15IqywkdA3HOwkQA9DCXBYS2QNg7bZfRA/exec
      var link_location = "https://script.google.com/a/organicdye.com/macros/s/AKfycbxV6ZjbbVbqLZf5nrmGoiy6_jpS_VgeY_qpln2Ey4qE4XvKtU0/exec?id=EP";
      break;
    case "NC":
      var link_location = "https://script.google.com/a/organicdye.com/macros/s/AKfycbxV6ZjbbVbqLZf5nrmGoiy6_jpS_VgeY_qpln2Ey4qE4XvKtU0/exec?id=NC";
      break;
    default:
      Logger.log("ERR in printDOC() data[1].location switch-case: " + data[1].location);
      break;
  }

//  switch(isLR){//fales is lab report
//    case false:
if(!isLR){var docid = DriveApp.getFileById(globalVariables().RPRTtemplate).makeCopy().getId();
//      var data = JSON.parse(inOBJ);
      var emailTo = [myUser];
      var subject = 'Report for Lab Request ' + data[0] + '     ' + data[1].company_name;
      var message = 'Please find attached the report for lab request ' + data[0] + '. \
\nYou can always access the project at ' + link_location + "&docid="+data[0]+ '. \
\nThank you!';
      var htmlMsg = '<p>Please find attached the report for lab request ' + data[0] + '. \
<br> You can always access the project <a href=' + link_location + "&docid="+data[0] + '>using this link here.</a> \
<br>Thank you!</p>';
       var myFilename  = 'Lab Report ' + data[0] + '.pdf';}
//      break;

//    case true:
 else if(isLR){      var docid = DriveApp.getFileById(globalVariables().RQSTtemplate).makeCopy().getId();
      var user2 = "";
      myUser = data[1].user;
      if (((data[1].salesperson != "House Accounts" && data[1].salesperson != "House Accounts") && data[1].salesperson != "other") && data[1].salesperson != 'ODP House'){
        Logger.log("I should not be here for ODP House");
        user2 = data[1].salesperson.split("", 1) + data[1].salesperson.split(" ")[1] + "@organicdye.com";}
      switch (data[1].location) {
        case "EP":
            var emailTo = [data[1].user, user2, 'colorlab@organicdye.com','cleal@organicdye.com','jneves@organicdye.com', 'cgustafson@organicdye.com','mdoyle@organicdye.com'];
          //var emailTo = [data[1].user, 'dwilliams@organicdye.com'];
          break;
        case "NC":
          var emailTo = [data[1].user, user2, 'mneale@organicdye.com','ksutton@organicdye.com', 'bbrady@organicdye.com', 'mdoyle@organicdye.com'];
//          var emailTo = [data[1].user, 'dwilliams@organicdye.com'];
          break;
        default:
          Logger.log('Error in data location switch');
          return 666;
      }
var emailStr = (function (){
   var str = "";
   for(var i = 0; i < emailTo.length; i++){
     str = str + String.fromCharCode(44) + emailTo[i];}
    return str;})();

        var folder = DriveApp.getFolderById(data[1].pjId);
     Logger.log("emailTo: "+emailStr);
     var attach = {
       fileName,
       content: .pdf,
       mimeType:'application/pdf'
     };
      var myFilename  = 'Lab Request ' + data[0] + '.pdf';
      var subject = data[1].company_name + ' Lab Request ' + data[1].lrno;
      var message = user2 + '\nPlease find attached your recent lab request.\
\nYou can follow the status of the project at ' + link_location  + "&docid="+data[0]+  '. \
\nThank you!';
      var htmlMsg = '<p>Please find attached your recent lab request. \
<br> You can follow the status of the project <a href=' + link_location  + "&docid="+data[0]+  '>using this link here.</a> \
<br>Thank you!</p>';
         }
//      break;
//    default:
  else{      Logger.log("Error in PrintDOC switch-case determining of Lab Req page or Project page");}
//      break;
//  }
    var doc = DocumentApp.openById(docid);
	var lock = LockService.getScriptLock();
	lock.waitLock(30000);

    var body = doc.getActiveSection();
//    var text = body.editAsText();

	var mydate = formatdate(data[1].datein);
    body.replaceText("%STREET1%",data[1].company_addr1);
    body.replaceText("%STREET2%",data[1].company_addr2);
    body.replaceText("%CITY%",data[1].customer_city);
    body.replaceText("%STATE%",data[1].customer_state);
    body.replaceText("%ZIP%",data[1].customer_zip_code);
    body.replaceText("%COUNTRY%",data[1].customer_country);
	body.replaceText("%ATTN%",data[1].customer_name);
	body.replaceText("%CUSTPHONE%",data[1].customer_phone);
	body.replaceText("%CUSTEMAIL%",data[1].customer_email);

    body.replaceText("%BRANCH%",data[1].location);

  body.replaceText("%MARKET%", data[1].market_end_use.name);
  if(data[1].detailed_instructions != null){
    body.replaceText("%INSTRUCTIONS%",data[1].detailed_instructions);}//if detailed inst not empty

  else if(data[1].detailed_instructions == null && data[1].afteraction.sample != null){
    body.replaceText("%INSTRUCTIONS%","%SAMPLESR% %SYORN%\n%SAMPLESIZE% %SSIZE%\n%PACKTYPE% %PTYPE%\n%BOOKLET% %BYORN%\n%SPECTROREP% %SPYORN%\n%MAILTO% %TOWHERE%\n");
    //samples
    boldfaceText("%SAMPLESR%", body);
    body.replaceText("%SAMPLESR%","Samples required?");
    body.replaceText("%SYORN%",data[1].afteraction.sample);

    if(data[1].afteraction.sample == "Yes, samples to be requested"){
    //size
      boldfaceText("%SAMPLESIZE%", body);
      body.replaceText("%SAMPLESIZE%","Sample size requested:");
      if(data[1].afteraction.size != "other"){
        body.replaceText("%SSIZE%",data[1].afteraction.size);}
      else if (data[1].afteraction.size == "other") {
        body.replaceText("%SSIZE%",data[1].afteraction.othersize);}
    boldfaceText("%PACKTYPE%", body);
    body.replaceText("%PACKTYPE%","Packaging requested:");
    body.replaceText("%PTYPE%",data[1].afteraction.packType);
    }//if sample=yes
    else if (data[1].afteraction.sample == "No"){
      body.replaceText("%SAMPLESIZE% %SSIZE%","");
      body.replaceText("%PACKTYPE% %PTYPE%","");
    }//if sample no

    boldfaceText("%BOOKLET%", body);
    body.replaceText("%BOOKLET%","Booklet/Display requested:");
    body.replaceText("%BYORN%",data[1].afteraction.booklet);

    boldfaceText("%SPECTROREP%", body);
    if(data[1].afteraction.spectro != "-1"){
    body.replaceText("%SPECTROREP%","Copy of Spectrophotometer print-out requested:");
    body.replaceText("%SPYORN%",data[1].afteraction.spectro);
    }//spectro report yes
    else if (data[1].afteraction.spectro == "-1"){
      body.replaceText("%SPECTROREP% %SPYORN%","");}
    boldfaceText("%MAILTO%", body);
      body.replaceText("%MAILTO%","Disposition of booklet and/or samples:\n");
    body.replaceText("%TOWHERE%",data[1].afteraction.mailto);

  }//if sample not null and detail inst empty

  switch (parseInt(data[1].priority)) {
    case 1:
      //range.setValue("Normal: In its turn");
	  body.replaceText("%PRIORITY%","Normal: In its turn");
	  body.replaceText("%JUSTIFY%","");
      break;
    case 2:
      //range.setValue("High \nTakes priority over Normal");
      body.replaceText("%PRIORITY%","High: Takes priority over Normal");
	  body.replaceText("%JUSTIFY%","");
      break;
    case 3:
      //range.setValue("RUSH \nSee justification below");
      body.replaceText("%PRIORITY%","RUSH: Justification below");
	  body.replaceText("%JUSTIFY%",data[1].justification);
      break;
    default:
      //range.setValue("DEFAULT - switch error data.priority");
	body.replaceText("%PRIORITY%","error in switch-case");
	body.replaceText("%JUSTIFY%","");
      break;
  }

    body.replaceText("%DUEDATE%",data[1].due_date);
	body.replaceText("%POTENTIAL%",data[1].potential_volume + " " + data[1].unit);
  	body.replaceText("%BASIS%",data[1].potBasis);

	body.replaceText("%DATEIN%",data[1].datein);

  switch (parseInt(data[1].samples)) {
    case 0:
      body.replaceText("%SAMPLES%","N/A");
      body.replaceText("%LIST%","N/A");
      break;
    case 1:
      body.replaceText("%SAMPLES%","Yes: Chemicals and/or Dyes/Pigments");
      body.replaceText("%LIST%",data[1].samplelist);
      break;
    case 2:
      body.replaceText("%SAMPLES%","Yes: Articles or Substrate ONLY");
      body.replaceText("%LIST%",data[1].samplelist);
      break;
    case 3:
      body.replaceText("%SAMPLES%","Yes: Both Articles & Chemicals/Dyes/Pigments");
      body.replaceText("%LIST%",data[1].samplelist);
      break;
    case 4:
      body.replaceText("%SAMPLES%","None: No chemicals/dyes/goods");
      body.replaceText("%LIST%","None");
      break;
    default:
      body.replaceText("%SAMPLES%","DEFAULT - switch error data.samples");
      body.replaceText("%LIST%","ERROR" + data[1].samples);

      break;
  }

  switch (parseInt(data[1].SDS)) {
    case 0:
      body.replaceText("%SDS%","");
      break;
    case 1:
      body.replaceText("%SDS%","SDS included \nwith sample");
      break;
    case 2:
      body.replaceText("%SDS%","SDS attached \nto this submission");
      break;
    case 3:
      body.replaceText("%SDS%","Salesperson promises \nto acquire an SDS");
      break;
    case 4:
      body.replaceText("%SDS%","SDS not required \nsample is an article");
      break;
    default:
      body.replaceText("%SDS%","DEFAULT - switch error data.SDS = " + data[1].SDS);
      break;
  }

      body.replaceText("%METHOD%", data[1].spectro + ' using ' + data[1].medium);
      body.replaceText("%LIGHT%", "First: " + data[1].light[0] + ", Second: " + data[1].light[1] + ", Third: " + data[1].light[2]);

	body.replaceText("%SUBMIT%",formatdate(data[1].datein).cal);
	body.replaceText("%OPEN%",data[2].stamp.opened);
	body.replaceText("%CLOSED%",data[2].stamp.closed);
	body.replaceText("%SALESPERSON%",data[1].salesperson);
	if (data.salesperson == 'other') {
		body.replaceText("%SALESPERSON%",data[1].sales_other);
	}
	body.replaceText("%COMPANY%",data[1].company_name);
	body.replaceText("%OWNER%",data[2].owner.name);
	body.replaceText("%SCOPE%",data[1].scope_of_work);
	body.replaceText("%PROBLEMSCOPE%",data[1].problem_scope);
	body.replaceText("%RESULTS%",data[2].results);
	body.replaceText("%ACTION%",data[2].action);
	body.replaceText("%UPDATES%",data[2].updates.logged);
    var myData = namedata1(JSON.stringify([true,"",""])).userLevels;

    function myLevel(userId){
      for(var i = 0; i < myData.length; i++){
        if(myData[i][0] == userId.split('@')[0]){return myData[i][1];}
        }
    };

  var userLevel = myLevel(myUser);

  Logger.log("myUser level is at: " + userLevel);

	switch (userLevel) {
        //the list below are users who are able to print the formula on teh report
		case '10':
		case '9':
		case '8':
		case '7':
            body.replaceText("%F%","Formula/Key");
			body.replaceText("%KEY%",data[2].key);
        if(data[2].purchasing.itemcount>0){
	      body.replaceText("%PURCHASING%","Purchasing Activity Notes");
          var pnotes = "";
          for (var k = 0; k < data[2].purchasing.itemcount; k++){
           pnotes = pnotes + "Request for " +  data[2].purchasing.data[4*k] + " supplied as a " + data[2].purchasing.data[4*k+1] + " product\
" + " with special attributes requiring " + data[2].purchasing.data[4*k+2] + ". \n\n Notes:\n" + data[2].purchasing.data[4*k+3] + "\n\n---\n\n";
            }
	      body.replaceText("%PNOTES%",pnotes);
        }
        else if(data[2].purchasing.itemcount == 0){
          body.replaceText("%PURCHASING%","");
          body.replaceText("%PNOTES%","");
        }
			break;
        case '0':
		default:
        //others not in the list will see a report but without formula
            body.replaceText("%F%","");
            body.replaceText("%KEY%","");
            body.replaceText("%PURCHASING%","");
            body.replaceText("%PNOTES%","");
			break;
	}
	var rootfolder = DriveApp.getFolderById("0BxlkPxuAhnwRaV90Tk9WQ1JrWk0");
	var folders = rootfolder.getFoldersByName(data[0]);
	var folder;
	if (folders.hasNext()) {
		folder = folders.next();
	} else {
		Logger.log("Error finding folder");
	}
	var files = folder.getFiles();
	var listOfFiles = '';
	while (files.hasNext()) {
		var file = files.next();
		file = file + '; ';
		listOfFiles = listOfFiles + file;
	}

  body.replaceText("%UPLOADED%",listOfFiles);

  var fol = body.getParent();
  for (i = 0; i<fol.getNumChildren(); i++){
    var child = fol.getChild(i);//.asFooterSection();
    child.replaceText("%TIMESTAMP%", 'Printed on ' + mydate.cal + " at " + mydate.time + ' by ' + myUser);//timestamp printed
    child.replaceText("%REQNO%",data[0]);
  }


  var footer = doc.getFooter();
  footer.editAsText().replaceText("%TIMESTAMP%", 'Printed on ' + mydate.cal + " at " + mydate.time + ' by ' + myUser);
  footer.editAsText().replaceText("%REQNO%", data[0]);

  doc.saveAndClose();

	var pdf = DriveApp.getFileById(docid).getAs('application/pdf').getBytes();

        var folder = DriveApp.getFolderById(data[1].pjId);
      if(isLR){
      var pdftodrive = DriveApp.getFileById(docid);
      var theBlob = pdftodrive.getBlob().getAs('application/pdf').setName(data[1].company_name + ' ' + myFilename);
      var newfile = folder.createFile(theBlob);}

      var attach = {
		fileName: myFilename,
		content: pdf,
		mimeType: 'application/pdf'
	};

  MailApp.sendEmail(emailTo, subject, message,{htmlBody: htmlMsg,
	attachments: [attach]
  });
	DriveApp.getFileById(docid).setTrashed(true);
	lock.releaseLock();
	return true;
}

function SendUpdate(infofromclient){
  if (JSON.parse(infofromclient)==0){
      return JSON.stringify('SendUpdate completed execution as null arg!');
  }

  else if (infofromclient){
      var package = JSON.parse(infofromclient);
      var subject = 'New update posted for Lab Request ' + package.reqno + ' & ' + package.cust;
      var message = package.user + ' wrote:\n\n' + package.post;
      MailApp.sendEmail(package.email, subject, message);
      return JSON.stringify('SendUpdate completed execution sending email!');
  }
  else {
      return JSON.stringify('SendUpdate FAILED execution!');
  }  }

//convertMS function taken from https://gist.github.com/remino/1563878
function convertMS(ms) {
  var d, h, m, s;
  s = Math.floor(ms / 1000);
  m = Math.floor(s / 60);
  s = s % 60;
  h = Math.floor(m / 60);
  m = m % 60;
  d = Math.floor(h / 24);
  h = h % 24;
  return { d: d, h: h, m: m, s: s };
};
/*********trigger function*************************/
function _makeLabReports() {
  var monthlist = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var lock = LockService.getScriptLock();
  var SHEET = globalVariables();
  var ss = SpreadsheetApp.openById(SHEET.EP_NC_ss);
	lock.waitLock(30000);
	var sheet = ss.getSheetByName("EP");
	sheet.activate();
	var ep_lrObjrng = sheet.getRange(2, 2, ss.getLastRow() - 1, 2).getValues();
//	lock.releaseLock();
	sheet = ss.getSheetByName("NC");
	sheet.activate();
	var nc_lrObjrng = sheet.getRange(2, 2, ss.getLastRow() - 1, 2).getValues();
	lock.releaseLock();
    function isOdd(num) {return num % 2;}
    var lrObjrng = [];

    for (var i = 0; i < ep_lrObjrng.length; i++){
      lrObjrng.push(ep_lrObjrng[i]);
      }
    for (var i = 0; i < nc_lrObjrng.length; i++){
      lrObjrng.push(nc_lrObjrng[i]);
      }


  var indate, opendate, closedate;
  var mydata = [];
  var numDays = 0;
  var timeNow = Date.now();

  for (var i = 0; i < lrObjrng.length; i++){
    //Logger.log(JSON.parse(lrObjrng[i][0]).lrno);
    indate = new Date(JSON.parse(lrObjrng[i][0]).datein);
    opendate = new Date(JSON.parse(lrObjrng[i][1]).stamp.opened);
    closedate = new Date(JSON.parse(lrObjrng[i][1]).stamp.closed)
  if (JSON.parse(lrObjrng[i][1]).stamp.closed == ""){closedate = new Date("12/31/2100")};
    numDays = convertMS(timeNow - closedate.getTime());
//    Logger.log(" age < 10 = " + (numDays.d<10));
    if(numDays.d < 10){
    mydata.push([
    JSON.parse(lrObjrng[i][0]).lrno,    //0
    JSON.parse(lrObjrng[i][1]).status,  //1
    JSON.parse(lrObjrng[i][0]).salesperson,//2
    JSON.parse(lrObjrng[i][0]).company_name,//3
    indate,//4
    opendate,//5
    closedate,//6
    JSON.parse(lrObjrng[i][1]).summary,//7
    JSON.parse(lrObjrng[i][1]).updates.last,//8
    JSON.parse(lrObjrng[i][1]).action]);//9
    }
}

  //1BPuaEjb5c2tXVae6AYpe0bLaZv-PdOrUL4OP6GKkfWQ
  var docid = DriveApp.getFileById('1BPuaEjb5c2tXVae6AYpe0bLaZv-PdOrUL4OP6GKkfWQ').makeCopy().getId();
  Logger.log("docid = " + docid);
  var doc = DocumentApp.openById(docid);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var d =  new Date();
  var body = doc.getBody();
  var text = body.editAsText();
  var header = doc.getHeader();

  var numChild = body.getParent().getNumChildren();
  for (var i = 0; i < numChild; i ++){
    Logger.log("child["+i+"] = "+body.getParent().getChild(i).getType());}

  var newdate = monthlist[d.getMonth()]+' '+d.getDate()+', '+d.getFullYear();
  Logger.log("newdate = " + newdate);
  header.getParent().getChild(2).asHeaderSection().editAsText().replaceText("%DATE%", newdate);

  var inqueue = [];
  var inprogress = [];
  var closedTenDays = [];

  for (var k=0; k < mydata.length; k++){
    if(mydata[k][1] == 1){inqueue.push(mydata[k])};
    if(mydata[k][1] == 0){inprogress.push(mydata[k])};
    if(mydata[k][1] == 2){closedTenDays.push(mydata[k])};
        }
 var section1 = body.appendParagraph("Projects In Queue");
 section1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
 var textblock = "\n";
 text.appendText(textblock);
  var docColor = '#ffffff';//white
//  var texttoline = body.getText();
  var table1 = body.appendTable();
  table1.setBorderColor('#ffffff')
  var tr;// = table.appendTableRow();
  var td;// = tr.appendTableCell(‘My Text’);

  for(var i = 0; i < inqueue.length; i++){
//    texttoline = body.getText();
    Logger.log("inqueue["+i+"] = " + inqueue[i][1] + "  " + inqueue[i][0]);
    textblock = inqueue[i][0] + ' for ' + inqueue[i][3] + ' Salesperson: ' + inqueue[i][2] + '\n Project Scope: ' + inqueue[i][7] + '\n' + 'Last Update: ' + inqueue[i][8] + '\n\n';
//    body.appendParagraph(textblock);
    tr = table1.appendTableRow();
    td = tr.appendTableCell(textblock)
    if(isOdd(i)==0){docColor = '#bcbcbc';} else {docColor = '#ffffff';}
    td.setBackgroundColor(docColor);
    //Logger.log("texttoline = " + texttoline.length + "  textblock =" + textblock.length)
    //body.editAsText().setBackgroundColor(texttoline.length, texttoline.length+ textblock.length, docColor);
  };

 body.appendParagraph('\n').appendPageBreak();
 var section2 = body.appendParagraph("Projects In Progress");
 section2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  var table2 = body.appendTable();
  table2.setBorderColor('#ffffff')

  for(var i = 0; i < inprogress.length; i++){
    Logger.log("inprogress["+i+"] = " + inprogress[i][1] + "  " + inprogress[i][0]);
    textblock = inprogress[i][0] + ' for ' + inprogress[i][3] + ' Salesperson: ' + inprogress[i][2] + '\n Project Scope: ' + inprogress[i][7] + '\n' + 'Last Update: ' + inprogress[i][8] + '\n\n';
    tr = table2.appendTableRow();
    td = tr.appendTableCell(textblock);
    Logger.log("i = " + td.isAtDocumentEnd());
    if(td.isAtDocumentEnd()){body.appendPageBreak();}
    if(isOdd(i)==0){docColor = '#bcbcbc';} else {docColor = '#ffffff';}
    td.setBackgroundColor(docColor);
//    body.appendParagraph(textblock)
  };

 body.appendParagraph("\n").appendPageBreak();
 var section3 = body.appendParagraph("Projects Closed in the Past 10 days");
 section3.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  var table3 = body.appendTable();
  table3.setBorderColor('#ffffff')
  for(var i = 0; i < closedTenDays.length; i++){
    Logger.log("closedTenDays["+i+"] = " + closedTenDays[i][1] + "  " + closedTenDays[i][0]);
    textblock = closedTenDays[i][0] + ' for ' + closedTenDays[i][3] + ' Salesperson: ' + closedTenDays[i][2] + '\n Project Scope: ' + closedTenDays[i][7] + '\n' + 'Last Update: ' + closedTenDays[i][8] + '\n' + 'Action Resulting: ' + closedTenDays[i][9] + '\n\n';
    tr = table3.appendTableRow();
    td = tr.appendTableCell(textblock)
    if(isOdd(i)==0){docColor = '#bcbcbc';} else {docColor = '#ffffff';}
    td.setBackgroundColor(docColor);
//    body.appendParagraph(textblock);
  };
  body.appendParagraph("\n").appendPageBreak();
  doc.saveAndClose();
  var emailTo = [];

  Logger.log(emailTo);
  var emailTo = ['dwilliams@organicdye.com', 'cdickerson@organicdye.com', 'amclean@organicdye.com'];
  //var emailTo = ['jneves@organicdye.com'];
  var subject = 'Weekly Status of Lab Requests for ' + monthlist[d.getMonth()]+'_'+d.getDate()+'_'+d.getFullYear()+'.pdf';
  var message = 'Please find attached the status for lab requests for ' + newdate + '. \
\nYou can always access the projects at https://sites.google.com/organicdye.com/orcoprojects/welcome  \
\nThank you!';
  var htmlMsg = '<p>Please find attached the status for lab requests for ' + newdate + '. \
<br> You can always access the projects <a href=https://sites.google.com/organicdye.com/orcoprojects/welcome>using this link here.</a> \
<br>Thank you!';
  var myFilename  = 'Lab_Request_Status_'+monthlist[d.getMonth()]+'_'+d.getDate()+'_'+d.getFullYear()+'.pdf';

  var pdf = DriveApp.getFileById(docid).getAs('application/pdf').getBytes();
  var pdftodrive = DriveApp.getFileById(docid);

      var attach = {
		fileName: myFilename,
		content: pdf,
		mimeType: 'application/pdf'
	};
  MailApp.sendEmail(emailTo, subject, message,{htmlBody: htmlMsg,
		attachments: [attach]
                    });

	lock.releaseLock();

DriveApp.getFileById(docid).setTrashed(true);

return true;
}

/*********trigger function*************************/
function _InternalLabReports() {
  var monthlist = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var lock = LockService.getScriptLock();
  var SHEET = globalVariables();
  var ss = SpreadsheetApp.openById(SHEET.EP_NC_ss);
	lock.waitLock(30000);
	var sheet = ss.getSheetByName("EP");
	sheet.activate();
	var ep_lrObjrng = sheet.getRange(2, 2, ss.getLastRow() - 1, 2).getValues();
//	lock.releaseLock();
	sheet = ss.getSheetByName("NC");
	sheet.activate();
	var nc_lrObjrng = sheet.getRange(2, 2, ss.getLastRow() - 1, 2).getValues();
	lock.releaseLock();
    function isOdd(num) {return num % 2;}
    var lrObjrng = [];

    for (var i = 0; i < ep_lrObjrng.length; i++){
      lrObjrng.push(ep_lrObjrng[i]);
      }
    for (var i = 0; i < nc_lrObjrng.length; i++){
      lrObjrng.push(nc_lrObjrng[i]);
      }


  var indate, opendate, closedate;
  var mydata = [];
  var numDays = 0;
  var timeNow = Date.now();

  for (var i = 0; i < lrObjrng.length; i++){
    Logger.log(JSON.parse(lrObjrng[i][0]).datein);
    indate = new Date(JSON.parse(lrObjrng[i][0]).datein);
    opendate = new Date(JSON.parse(lrObjrng[i][1]).stamp.opened);
    closedate = new Date(JSON.parse(lrObjrng[i][1]).stamp.closed)
  if (JSON.parse(lrObjrng[i][1]).stamp.closed == ""){closedate = new Date("12/31/2100")};
    numDays = convertMS(timeNow - closedate.getTime());
//    Logger.log(" age < 10 = " + (numDays.d<10));
    if(numDays.d < 10){
    mydata.push([
    JSON.parse(lrObjrng[i][0]).lrno,    //0
    JSON.parse(lrObjrng[i][1]).status,  //1
    JSON.parse(lrObjrng[i][0]).salesperson,//2
    JSON.parse(lrObjrng[i][0]).company_name,//3
    indate,//4
    opendate,//5
    closedate,//6
    JSON.parse(lrObjrng[i][0]).scope_of_work,//7
    JSON.parse(lrObjrng[i][1]).updates.last,//8
    JSON.parse(lrObjrng[i][1]).action]);//9
    }
  }
  //1BPuaEjb5c2tXVae6AYpe0bLaZv-PdOrUL4OP6GKkfWQ
  var docid = DriveApp.getFileById('1BPuaEjb5c2tXVae6AYpe0bLaZv-PdOrUL4OP6GKkfWQ').makeCopy().getId();
  Logger.log("docid = " + docid);
  var doc = DocumentApp.openById(docid);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var d =  new Date();
  var body = doc.getBody();
  var text = body.editAsText();
  var header = doc.getHeader();

  var numChild = body.getParent().getNumChildren();
  for (var i = 0; i < numChild; i ++){
    Logger.log("child["+i+"] = "+body.getParent().getChild(i).getType());}

  var newdate = monthlist[d.getMonth()]+' '+d.getDate()+', '+d.getFullYear();
  Logger.log("newdate = " + newdate);
  header.getParent().getChild(2).asHeaderSection().editAsText().replaceText("%DATE%", newdate);

  var inqueue = [];
  var inprogress = [];
//  var closedTenDays = [];

  for (var k=0; k < mydata.length; k++){
    if(mydata[k][1] == 1){inqueue.push(mydata[k])};
    if(mydata[k][1] == 0){inprogress.push(mydata[k])};
        }
 var section1 = body.appendParagraph("Projects In Queue");
 section1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
 var textblock = "\n";
 text.appendText(textblock);
  var docColor = '#ffffff';//white
  var table1 = body.appendTable();
  table1.setBorderColor('#ffffff')
  var tr;// = table.appendTableRow();
  var td;// = tr.appendTableCell(‘My Text’);

  for(var i = 0; i < inqueue.length; i++){
    Logger.log("inqueue["+i+"] = " + inqueue[i][1] + "  " + inqueue[i][0]);
    textblock = 'Date: '+inqueue[i][4] + '\n'+inqueue[i][0] + ' for ' + inqueue[i][3] + ' Salesperson: ' + inqueue[i][2] + '\n Project Scope: ' + inqueue[i][7] + '\n';// + 'Last Update: ' + inqueue[i][8] + '\n\n';
    tr = table1.appendTableRow();
    td = tr.appendTableCell(textblock)
    if(isOdd(i)==0){docColor = '#bcbcbc';} else {docColor = '#ffffff';}
    td.setBackgroundColor(docColor);
  };

 body.appendParagraph('\n').appendPageBreak();
 var section2 = body.appendParagraph("Projects In Progress");
 section2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  var table2 = body.appendTable();
  table2.setBorderColor('#ffffff')

  for(var i = 0; i < inprogress.length; i++){
    Logger.log("inprogress["+i+"] = " + inprogress[i][1] + "  " + inprogress[i][0]);
    textblock = 'Date: '+inprogress[i][4] + '\n'+inprogress[i][0] + ' for ' + inprogress[i][3] + ' Salesperson: ' + inprogress[i][2] + '\n Project Scope: ' + inprogress[i][7] + '\n';// + 'Last Update: ' + inprogress[i][8] + '\n\n';
    tr = table2.appendTableRow();
    td = tr.appendTableCell(textblock);
    Logger.log("i = " + td.isAtDocumentEnd());
    if(td.isAtDocumentEnd()){body.appendPageBreak();}
    if(isOdd(i)==0){docColor = '#bcbcbc';} else {docColor = '#ffffff';}
    td.setBackgroundColor(docColor);
  };

 body.appendParagraph("\n").appendPageBreak();
  doc.saveAndClose();

Logger.log(emailTo);
 var emailTo = ['dwilliams@organicdye.com'];
 var subject = 'Weekly Status of Lab Requests for ' + monthlist[d.getMonth()]+'_'+d.getDate()+'_'+d.getFullYear()+'.pdf';
 var message = 'Please find attached the status for lab requests for ' + newdate + '. \
 \\nYou can always access the projects at https://sites.google.com/organicdye.com/orcoprojects/welcome  \
 \nThank you!';
 var htmlMsg = '<p>Please find attached the status for lab requests for ' + newdate + '. \
 <br> You can always access the projects <a href=https://sites.google.com/organicdye.com/orcoprojects/welcome>using this link here.</a> \
 <br>Thank you!';
 var myFilename  = 'Lab_Request_Status_'+monthlist[d.getMonth()]+'_'+d.getDate()+'_'+d.getFullYear()+'.pdf';

  var pdf = DriveApp.getFileById(docid).getAs('application/pdf').getBytes();
  var pdftodrive = DriveApp.getFileById(docid);

      var attach = {
		fileName: myFilename,
		content: pdf,
		mimeType: 'application/pdf'
	};
  MailApp.sendEmail(emailTo, subject, message,{htmlBody: htmlMsg,
		attachments: [attach]
                    });

	lock.releaseLock();

 DriveApp.getFileById(docid).setTrashed(true);

 return true;
}
