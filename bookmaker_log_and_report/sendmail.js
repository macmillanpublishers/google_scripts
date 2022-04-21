function sendmail(resource_sheet_id,filename,linkUrl,startdate,enddate,driveid,testing,submitter) {
  // default values for tests
  resource_sheet_id = resource_sheet_id || recurring_rprt_resource_sheet_id;
  testing = typeof testing !== 'undefined' ? testing : true; // revert to default if val is unset, not just when falsey
  linkUrl = linkUrl || "google.com";
  filename = filename || "reportfile"
  startdate = startdate || "date unavailable";
  enddate = enddate || "date unavailable";
  driveid = driveid || recurring_rprt_resource_sheet_id;  // just to use an id we already have

  // get sheets, set vars
  var mail_ss = SpreadsheetApp.openById(resource_sheet_id);
  if (testing === true) {
    var addressSheet = mail_ss.getSheetByName("test_recipients");
  } else {
    var addressSheet = mail_ss.getSheetByName("recipients");
  }

  var daterange_str = getShortDateStr(startdate)+" to "+getShortDateStr(enddate)
  var subjectTxt = mail_ss.getSheetByName("email_txt").getSheetValues(1,2,1,1).toString();
  var subjectTxt = subjectTxt.replace(/DATERANGE/,daterange_str);
  var bodyTxt = mail_ss.getSheetByName("email_txt").getSheetValues(2,2,1,1).toString();
  var bodyTxt = bodyTxt.replace(/URLHERE/, linkUrl).replace(/DATERANGE/,daterange_str).replace(/FILENAME/,filename);

  // blank submitter indicates scheduled report, permissions should also be granted
  if (submitter == '') {
    var Tolength = addressSheet.getRange("A1:A").getValues().filter(String).length;
    var CClength = addressSheet.getRange("B1:B").getValues().filter(String).length;
    var Editorslength = addressSheet.getRange("C1:C").getValues().filter(String).length;
    var recipientsTOarray = addressSheet.getSheetValues(2, 1, Tolength-1, 1);
    var recipientsTO = recipientsTOarray.join(', ').toString();
    if (CClength > 1) {
      var recipientsCCarray = addressSheet.getSheetValues(2, 2, CClength-1, 1);
      var recipientsCC = recipientsCCarray.join(', ').toString();
    } else {
      var recipientsCCarray = [];
      var recipientsCC = '';
    }

    ///// set permissions
    var recipientsALLarray = recipientsCCarray.concat(recipientsTOarray);
    if (Editorslength > 1) {
      var sheetEditors = addressSheet.getSheetValues(2, 3, Editorslength-1, 1);
    } else {
      var sheetEditors = [];
    }
    addSheetPermissions(driveid, sheetEditors, [], recipientsALLarray)
  } else {
    // non-blank submitter means form submission, just mail to them, theyre already made owners of sheet
    var recipientsTO = submitter
    var recipientsCC = ''
    var recipientsALLarray = submitter
  }

  console.info("sending mail: '"+subjectTxt+"' to recipients: " + recipientsALLarray);
  try {
    MailApp.sendEmail({
      to: recipientsTO,
      cc: recipientsCC,
      subject: subjectTxt,
      htmlBody: bodyTxt
    });
    console.log("mail sent")
  } catch (e) {
    console.error("ERROR sending mail: '"+subjectTxt+"' to recipients: " + recipientsALLarray+":: "+e);
  }
}
