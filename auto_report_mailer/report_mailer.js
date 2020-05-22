function runtest() {
  //  Run this to send to test_recipients.
  //  Also uses egalleymaker_report_stg folder by default.
  //    can pass second parameter as boolean 'false' (no quotes) to use default prod folder instead
  var testing = true;
  var egalleymaker_stg_folder_id = '18bUCterLZvCnn0tDa-YoCJk2W2ggKwyZ';
  Logger.log("testing!")
  main(testing, egalleymaker_stg_folder_id);
}

function sendmail(linkUrl,reportDate,driveid,testing) {
  testing = testing || false;
  linkUrl = linkUrl || "<dummylink-url undefined>";
  reportDate = reportDate || "<date unavailable>";
  driveid = driveid || "1Q44nYVNNZCPw7IF3LuGHQ9jD5NPrMCg0hScJ_DJ07cg";  //from testing file
 var mail_ss = SpreadsheetApp.openById("1-oeosXbcBkfewADNxnHt-EBXGUO2eF6DreIKA-AFBHs");
  if (testing === true) {
    var addressSheet = mail_ss.getSheetByName("test_recipients");
  } else {
    var addressSheet = mail_ss.getSheetByName("recipients");
  }
 var Tolength = addressSheet.getRange("A1:A").getValues().filter(String).length;
 var CClength = addressSheet.getRange("B1:B").getValues().filter(String).length;
 var Editorslength = addressSheet.getRange("C1:C").getValues().filter(String).length;
 var recipientsTOarray = addressSheet.getSheetValues(2, 1, Tolength-1, 1);
 var recipientsTO = recipientsTOarray.join(', ').toString();
 var recipientsCCarray = addressSheet.getSheetValues(2, 2, CClength-1, 1);
 var recipientsCC = recipientsCCarray.join(', ').toString();
 var subjectTxt = mail_ss.getSheetByName("email_txt").getSheetValues(1,2,1,1).toString();
 var subjectTxt = subjectTxt.replace(/DATE/,reportDate);
 var bodyTxt = mail_ss.getSheetByName("email_txt").getSheetValues(2,2,1,1).toString();
 var bodyTxt = bodyTxt.replace(/URLHERE/, linkUrl).replace(/DATE/,reportDate);

  ///// set permissions
  var recipientsALLarray = recipientsCCarray.concat(recipientsTOarray);
  var sheetEditors = addressSheet.getSheetValues(2, 3, Editorslength-1, 1);
  var targetFile = SpreadsheetApp.openById(driveid);

  for(i=0; i<recipientsALLarray.length; i++) {
    try{
    targetFile.addViewers(recipientsALLarray[i]);
      Logger.log('added '+recipientsALLarray[i]+ ' as viewer');
    }
    catch(e){
      Logger.log('An error has occurred: '+e.message)
    }
  }
  for(i=0; i<sheetEditors.length; i++) {
    try{
    targetFile.addEditors(sheetEditors[i]);
      Logger.log('added '+sheetEditors[i] + ' as editor');
    }
    catch(e){
      Logger.log('An error has occurred: '+e.message)
    }
  }
  writelogtoDoc(recipientsTO)
  Logger.log("all recipients: " + recipientsALLarray);
  MailApp.sendEmail({
    to: recipientsTO,
    cc: recipientsCC,
    subject: subjectTxt,
    htmlBody: bodyTxt
  });
  return "mail sent";
}

function main(testing, egalleymaker_folder_id) {
  testing = testing || false;
  egalleymaker_folder_id = egalleymaker_folder_id || '0B-vhqV0CBZhDR1RjMVVxWkMzb3c';
  ///// Find our spreadsheet
  refreshDriveFolderList(egalleymaker_folder_id)
  var status = 'file not found';
  var testwriter = 'matthew.retzer@macmillan.com' // use this so we don't pick up already shared/mailed items
  var namesample = 'egalleymaker_report_'
  var file = false;

  // get modified by date
  var FiveDaysBeforeNow = new Date().getTime()-3600*1000*24*6;
  var cutOffDate = new Date(FiveDaysBeforeNow);
  var cutOffDateAsString = Utilities.formatDate(cutOffDate, "GMT", "yyyy-MM-dd");
  Logger.log(cutOffDateAsString);

  //  var files = DriveApp.searchFiles('title contains ".tsv"');
  var files = DriveApp.searchFiles('title contains "'+namesample+'" and "'+ egalleymaker_folder_id + '" in parents and modifiedDate > "' + cutOffDateAsString + '" and not "'+testwriter+'" in writers');
  while (files.hasNext()) {
    var status = 'file found';
    if (file && files.next().getLastUpdated() > file.getLastUpdated()) {
      var file = files.next();
    } else if (file == false) {
      var file = files.next();
    }
  }
  if (file != false) {
    var driveid = file.getId();
    var driveurl = file.getUrl();
    ///// Rename the spreadsheet
    //var todayDate = Utilities.formatDate(new Date(), "EDT", "MM-dd-yyyy");
    var name = file.getName();
    var newname = name.split('.').shift(1); //+ '_' + todayDate;
    if (name.match(/.tsv/)) {
      file.setName(newname);
      Logger.log(name+","+newname);
    } else {
      Logger.log(name);
    }
    reportDate = formatSheet(file, newname)
    var status = sendmail(driveurl,reportDate,driveid,testing);
  }
  Logger.log(status);
  writelogtoDoc(status);
}

///// Format our sheet
function formatSheet(file, newname) {
  var ss = SpreadsheetApp.open(file);
  var reportDate = newname.split('_')[2]
  var sheet = ss.getSheets()[0];
  var borderRange = sheet.getDataRange();
  var headerRange = sheet.getRange("A1:H1");
  borderRange.setBorder(true, true,true, true, true, true, '#808080', null).setHorizontalAlignment("left");
  headerRange.setFontWeight("bold").setBackground('#d0e0e3').setHorizontalAlignment("center");
  //adding conditional to prevent error if we have only a header row
  if (sheet.getLastRow() > 1) {
    var stepsRange = sheet.getDataRange().offset(1, 0, sheet.getLastRow() - 1);
    setAlternatingRowBackgroundColors_(stepsRange, '#ffffff', '#eeeeee');
    for (var column = 1; column<=stepsRange.getNumColumns(); column++) {
      sheet.autoResizeColumn(column);
    }
  }
  return reportDate
}

function setAlternatingRowBackgroundColors_(range, oddColor, evenColor) {
  var backgrounds = [];
  for (var row = 1; row <= range.getNumRows(); row++) {
    var rowBackgrounds = [];
    for (var column = 1; column <= range.getNumColumns(); column++) {
      if (row % 2 == 0) {
        rowBackgrounds.push(evenColor);
      } else {
        rowBackgrounds.push(oddColor);
      }
    }
    backgrounds.push(rowBackgrounds);
  }
  range.setBackgrounds(backgrounds);
}

function writelogtoDoc(scriptStatus) {
  scriptStatus = scriptStatus || "no scriptStatus available";
  var doc = DocumentApp.openById('1HTWhWYrVSFaK9aqPkeaLMPvWzvAi6AuICEOnBM4f60U');
  var body = doc.getBody();
  var text = body.editAsText();
  var todayTime = new Date(); //Utilities.formatDate(new Date(), "EDT", "MM-dd-yyyy'T'HH:mm:ss'Z'");
  var logText = '[' + todayTime + ']  ' + scriptStatus
  Logger.log(logText);
  text.insertText(0, logText + '\n');
}

function refreshDriveFolderList(drivefolder_id) {
  // Log the name of every file in the user's selected Drive folder.
  var drivefolder = DriveApp.getFolderById(drivefolder_id);
  var files = drivefolder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    //    Logger.log(file.getName());
  }
}
