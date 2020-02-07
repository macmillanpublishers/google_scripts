function onFormSubmit(e) {

console.info(e.namedValues) // log to stackdriver

// skip weird blank submissions / resubmissions
  if (e.namedValues['ISBN:'].toString() === "") {
    console.info("ISBN is blank! : '" + e.namedValues['ISBN:'].toString() + "'exiting function early.");
    return;
  } else {
    console.info("ISBN found: '" + e.namedValues['ISBN:'].toString() + "', continuing");
  }

  // Lock script with lockservice so we don't get concurrent runs of script
  var lock = LockService.getScriptLock();
  lock.waitLock(180000);
  console.info('<--- post-lock timestamp');

  // VARS - Sheets
  var drive_folder = DriveApp.getFolderById('1tY4orif5YOHN68fbiEvyldhWP8J7UAI0');
  var ss_main = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_main = ss_main.getActiveSheet();
  var ss_settings = SpreadsheetApp.openById("1gwKAtJFrxsbgoA8hRSbQ0h9eoAkQQPqgGDB39GFgbmU");
  var sheet_imprints = ss_settings.getSheetByName("Imprint Names");
  var imprints_first_row = 6;
  // vars - for header row
  var cols_to_hide = [21,26,31,36,41,46,51,56,61,65,69];
  var string_replace_dict = {
    "If stockhouse is not listed in drop-down menu, please indicate the correct stockhouse here":"Unlisted Stockhouse"
  };

  // new submitted row
  var lastrow = sheet_main.getLastRow();
  var src_sheet_lastcol = sheet_main.getLastColumn();
  var new_submitted_row = sheet_main.getRange(lastrow,1,1,src_sheet_lastcol).getValues();
  var new_submitted_row_formats = sheet_main.getRange(lastrow,1,1,src_sheet_lastcol).getNumberFormats();
  var submitter = getValueByColNameRowNum(sheet_main, 'Email Address', lastrow);
  var timestamp = getValueByColNameRowNum(sheet_main, 'Timestamp', lastrow);
  var imprint = getValueByColNameRowNum(sheet_main, 'Company:', lastrow).replace(/'/g, '');
  var unlisted_stockhouse = checkforUnlistedStockhouse(sheet_main, lastrow);
  //Logger.log('unlisted_stockhouse: '+ unlisted_stockhouse);


  // retrieve imprint sheet list from reference sheet
  var imprintValues = sheet_imprints.getRange(imprints_first_row, 1, ss_settings.getLastRow()-imprints_first_row+1,4).getValues();

  // find names of destination ss & sheet & recipient
  var ss_name_new = '';
  var sheet_name_new = '';
  var imprint_rownum = '';
  var email_recipients = '';
  for(var i = 0; i < imprintValues.length; i++) {
    if(sanitizeString(imprintValues[i][0]) == sanitizeString(imprint)) {
      ss_name_new = imprintValues[i][1];
      sheet_name_new = imprintValues[i][2];
      email_recipients = imprintValues[i][3];
      // for assigning theme later if needed
      imprint_rownum = i;
    }
  }

  // get ss id and spreadsheet, create spreadsheet as needed
  if (SearchFiles(ss_name_new)) {
    var ss_id_new = SearchFiles(ss_name_new);
  } else {
    var ss_id_new = createSS(ss_name_new, sheet_name_new, drive_folder);
  }

  // create sheet as needed
  var ss_new = SpreadsheetApp.openById(ss_id_new);
  var sheet_new = createSheetAsNeeded(ss_new, sheet_name_new);

  // add header row as needed
  var src_sheet_lastcol = sheet_main.getLastColumn();
  var header_row_values = sheet_main.getRange(2,1,1,src_sheet_lastcol).getValues();
  //  var header_row = sheet_src.getRange(2,1,1,src_sheet_lastcol);
  var header_inserted = pasteHeaderRow(header_row_values, sheet_new, cols_to_hide, string_replace_dict, imprint_rownum, src_sheet_lastcol);

  // prepend to target sheet!
  var newrow = sheet_new.insertRowAfter(1);
  var dest_sheet_lastcol = sheet_new.getLastColumn();

  // copy cell by cell to match header
  for (var i = 0; i < new_submitted_row[0].length; i++) {
    var dest_column = getColWithHeaderText(sheet_new,header_row_values[0][i])
    Logger.log(dest_column)
    if (dest_column >= 0) {
      if (header_row_values[0][i] == "Timestamp") {
        sheet_new.getRange(2,dest_column,1,1).setNumberFormat("m/d/yyyy h:mm:ss");
      } else {
        sheet_new.getRange(2,dest_column,1,1).setNumberFormat("@");
        // var newvalue = sheet_new.getRange(2,dest_column).setValue(String(new_submitted_row[0][i]));
      }
      var newvalue = sheet_new.getRange(2,dest_column).setValue(new_submitted_row[0][i]);
    }
  }


  // resize columnsbased on first row of responses (if new):
  if (header_inserted == 'y') {
    sheet_new.autoResizeColumns(1, src_sheet_lastcol);
  }

  // release the lock!
  lock.releaseLock();

  // send notification email to any email_recipients
  if (email_recipients) {
    // calculate sheet url
    var linkUrl = '';
    linkUrl += ss_new.getUrl() || "<dummylink-url undefined>";
    linkUrl += '#gid=';
    linkUrl += sheet_new.getSheetId();
    // get subject and body from settings sheet
    // (set alt text for unlisted stockhouse alert)
    if (unlisted_stockhouse == true) {
      var subjectTxt = ss_settings.getSheetByName("email_notification-txt").getSheetValues(4,2,1,1).toString();
      var bodyTxt = ss_settings.getSheetByName("email_notification-txt").getSheetValues(5,2,1,1).toString();
      var guest_recipients = ss_settings.getSheetByName("email_notification-txt").getSheetValues(6,2,1,1).toString();
      email_recipients = email_recipients + ", " + guest_recipients
    } else {
      var subjectTxt = ss_settings.getSheetByName("email_notification-txt").getSheetValues(1,2,1,1).toString();
      var bodyTxt = ss_settings.getSheetByName("email_notification-txt").getSheetValues(2,2,1,1).toString();
    }
    var subjectTxt = subjectTxt.replace(/IMPRINT/, sanitizeString(imprint));
    var bodyTxt = bodyTxt.replace(/URLHERE/, linkUrl).replace(/TIMESTAMP/,timestamp).replace(/SUBMITTER/,submitter).replace(/IMPRINT/, imprint);
    // Send mail!
    var mailstatus = sendMail(email_recipients, subjectTxt, bodyTxt);
    console.info(mailstatus);
  }
  console.info("script completed");
}

function getColWithHeaderText(dest_sheet,header_text){
  var range = dest_sheet.getDataRange()//.getValues();
  var width = range.getWidth();
  for(var i = 1; i <= width; i++) {
    var data = range.getCell(1,i).getValues();
    if (data == header_text) {
      return(i); // return the column number if we find it
    }
  }
  return(-1);
}


function getValueByColNameRowNum(sheet, colName, row) {
//  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var col = data[0].indexOf(colName);
  if (col != -1) {
    return data[row-1][col];
  }
}

function checkforUnlistedStockhouse(sheet, row) {
  colName = 'STOCKHOUSE: Name'
  unlistedstring = "Stockhouse Not Listed"
  unlisted_stockhouse = false;
  var data = sheet.getDataRange().getValues();
  // Logger.log("data0 "+ data[0])
  for(i = 0; i < data[0].length; i++) {
    //if (data[0][i] == colName) {
    //  Logger.log("values: " + data[row-1][i]);
    //}
    if (data[0][i] == colName && data[row-1][i].indexOf(unlistedstring) > -1){
      unlisted_stockhouse = true;
    }
  }
  return unlisted_stockhouse;
}

function pasteHeaderRow(header_row_values, sheet_dest, cols_to_hide, string_replace_dict, imprint_rownum, src_sheet_lastcol) {
  // CHECK to see if we need new header row (empty sheet)
  lastrow = sheet_dest.getLastRow();
  if (lastrow == 0) {

    // VARS
    var header_row_colors = [ "plum", "darkseagreen", "Moccasin", "salmon", "lightsteelblue", "palegreen", "lightpink", "lemonchiffon", "royal blue", "grey" ];
    //    colornum = imprint_rownum % 10;

    // paste
    target_range = sheet_dest.getRange(1,1,1,src_sheet_lastcol);
    target_range.setValues(header_row_values);

    // CLEANUP: hide cols
    for (var i = 0; i < cols_to_hide.length; i++) {
      sheet_dest.hideColumns(cols_to_hide[i]);
    }
    // replace header text
    for (var k = 0; k < Object.keys(string_replace_dict).length; k++) {
      var fstring = Object.keys(string_replace_dict)[k];
      var rstring = string_replace_dict[Object.keys(string_replace_dict)[k]];
      fandrHeader(sheet_dest,fstring,rstring);
    }
    // apply banding
    var banding_target = sheet_dest.getRange(1, 1, 5, src_sheet_lastcol);
    banding_target.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY).setHeaderRowColor(header_row_colors[imprint_rownum]);

    return "y";
  } else {
    Logger.log("sheet already has contents, skipping insert header row");
    return "n";
  }
}

function fandrHeader(sheet,find,repl) {
  var r=sheet.getDataRange();
//  var rws=r.getNumRows();
  var rws = 1;
  var cls=r.getNumColumns();
  var i,j,a; //,find,repl;
//  find="If stockhouse is not listed in drop-down menu, please indicate the correct stockhouse here";
//  repl="Unlisted Stockhouse";
  for (i=1;i<=rws;i++) {
    for (j=1;j<=cls;j++) {
      a=r.getCell(i, j).getValue();
      if (r.getCell(i,j).getFormula()) {continue;}
      try {
        a=a.replace(find,repl);
        r.getCell(i, j).setValue(a);
      }
      catch (err) {continue;}
    }
  }
}


function sanitizeString(s) {
  var newstring = s.replace(/[^a-z0-9 ]/gi, '');
//  Logger.log(newstring);
  return newstring;
}

function createSS(ss_name, sheet_name, drive_folder) {
  var file=SpreadsheetApp.create(ss_name);
  var copyFile=DriveApp.getFileById(file.getId());
  //  Logger.log(file.getId())
  drive_folder.addFile(copyFile);
  DriveApp.getRootFolder().removeFile(copyFile);
  // create sheet
  var sheets = file.getSheets();
  sheets[0].setName(sheet_name);
//  Logger.log(file.getId())
//  Logger.log(copyFile.getId())
  return file.getId();
}

function createSheetAsNeeded(ss, sheetname) {
 var sheetcheck = ss.getSheetByName(sheetname);
 if (!sheetcheck) {
   ss.insertSheet(sheetname);
 }
 return ss.getSheetByName(sheetname);
}

function SearchFiles(filename) {
  //Please enter your search term in the place of Letter
  var searchFor ='title = "' + filename + '"';
  var names =[];
  var fileIds=[];
  var files = DriveApp.searchFiles(searchFor);
  while (files.hasNext()) {
    var file = files.next();
    var fileId = file.getId();// To get FileId of the file
    fileIds.push(fileId);
    var name = file.getName();
    names.push(name);
  }
  return fileIds[0];
}

function sendMail(email_recipients, subjectTxt, bodyTxt) {
  MailApp.sendEmail({
    to: email_recipients,
    subject: subjectTxt,
    htmlBody: bodyTxt
  });
  return "mail sent";
}

function onFormSubmit_Leading0s(e){   // To to the resources menu and set this to run onFormSubmit();
  var source = e.source;
  var range = e.range;
  var sheet = range.getSheet();
  var row = range.getRow();

  var formURL = source.getFormUrl();
  var form = FormApp.openByUrl(formURL);
  var responses = form.getResponses();
  var responcesL = responses.length;
  var lastResponce = responses[responcesL-1];

  for (var k = 0; k < Object.keys(itemIds).length; k++) {
    var itemId = Object.keys(itemIds)[k];
    var column = itemIds[Object.keys(itemIds)[k]];
    var item = form.getItemById(itemId);
    if(lastResponce.getResponseForItem(item)) {
      var value = lastResponce.getResponseForItem(item).getResponse();

      var range2edit = sheet.getRange(row, column);
      range2edit.setNumberFormat("@STRING@");
      range2edit.setValue(value);
    }
  }
}
