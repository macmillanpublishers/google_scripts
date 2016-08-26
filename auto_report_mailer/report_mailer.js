function sendmail(linkUrl,reportDate,driveid) {
  linkUrl = linkUrl || "<dummylink-url undefined>";
  reportDate = reportDate || "<date unavailable>";
  driveid = driveid || "1Q44nYVNNZCPw7IF3LuGHQ9jD5NPrMCg0hScJ_DJ07cg";  //from testing file
 var mail_ss = SpreadsheetApp.openById("1-oeosXbcBkfewADNxnHt-EBXGUO2eF6DreIKA-AFBHs"); 
 var addressSheet = mail_ss.getSheets()[0];
 var Tolength = addressSheet.getRange("A1:A").getValues().filter(String).length;
 var CClength = addressSheet.getRange("B1:B").getValues().filter(String).length;
 var Editorslength = addressSheet.getRange("C1:C").getValues().filter(String).length;
 var recipientsTOarray = addressSheet.getSheetValues(2, 1, Tolength-1, 1);
 var recipientsTO = recipientsTOarray.join(', ').toString();
 var recipientsCCarray = addressSheet.getSheetValues(2, 2, CClength-1, 1);
 var recipientsCC = recipientsCCarray.join(', ').toString();
 var subjectTxt = mail_ss.getSheets()[1].getSheetValues(1,2,1,1).toString();
 var subjectTxt = subjectTxt.replace(/DATE/,reportDate);
 var bodyTxt = mail_ss.getSheets()[1].getSheetValues(2,2,1,1).toString();
 var bodyTxt = bodyTxt.replace(/URLHERE/, linkUrl).replace(/DATE/,reportDate);

  ///// set permissions
  var recipientsALLarray = recipientsCCarray.concat(recipientsTOarray);
  var sheetEditors = addressSheet.getSheetValues(2, 3, Editorslength-1, 1);
  var targetFile = SpreadsheetApp.openById(driveid);
  targetFile.addViewers(recipientsALLarray);
  targetFile.addEditors(sheetEditors);
  
  MailApp.sendEmail({
    to: recipientsTO,
    cc: recipientsCC,
    subject: subjectTxt,
    htmlBody: bodyTxt
  }); 
  return "mail sent";
}

function main() {
///// Find our spreadsheet  
  var status = 'file not found';
  var files = DriveApp.searchFiles('title contains ".tsv"');
  while (files.hasNext()) {
    var status = 'file found';   
    var file = files.next();
    var ss = SpreadsheetApp.open(file); 
    var driveid = file.getId();
    var driveurl = file.getUrl();
///// Rename the spreadsheet, and edit permissions
    var name = file.getName();
    var newname = name.split('.').shift(1);
      file.setName(newname);
///// Format our sheet
    var reportDate = newname.split('_')[2]
    var sheet = ss.getSheets()[0];
    var borderRange = sheet.getDataRange();
    var headerRange = sheet.getRange("A1:H1");
    borderRange.setBorder(true, true,true, true, true, true, '#808080', null).setHorizontalAlignment("left");
    headerRange.setFontWeight("bold").setBackground('#d0e0e3').setHorizontalAlignment("center");
    //sheet.getRange('a1').setValue('test');//
    var stepsRange = sheet.getDataRange().offset(1, 0, sheet.getLastRow() - 1);
    setAlternatingRowBackgroundColors_(stepsRange, '#ffffff', '#eeeeee');
    for (var column = 1; column<=stepsRange.getNumColumns(); column++) {
      sheet.autoResizeColumn(column);
    }  
    var status = sendmail(driveurl,reportDate,driveid);
  }
  writelogtoDoc(status);
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
  var todayTime = new Date();
  var logText = '[' + todayTime + ']  ' + scriptStatus
  Logger.log(logText);
  text.insertText(0, logText + '\n');
}
