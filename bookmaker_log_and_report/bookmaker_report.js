// form constants
const form_report_folder_id = 'REDACTED'
const form_title_daterange = 'Date Range'
const form_title_datestart = 'Report START date:'
const form_title_dateend = 'Report END date:'
const form_range_week = 'the past week'
const form_range_month = 'the past month'
const form_range_all = 'All Available'

// recurring_report constants
const recurring_report_folder_id = 'REDACTED'
const recurring_rprt_resource_sheet_id = 'REDACTED'

// SET TESTING/STAGING VALUES HERE
const form_staging = true  // <-- for form: determines whether we generate report from staging log or prod log, adjusts name of report
const recurring_staging = true  // <-- for recurring_report: determines whether we generate report from staging log or prod log, adjusts name of report & recipients
const test_submitter = 'REDACTED'  // <-- default form submitter value when running report function directly for debug

////////////////////////// GAPPS & Drive functions
function setupFormTrigger() {
  ScriptApp.newTrigger('runFromForm')
  .forForm('REDACTED')
  .onFormSubmit()
  .create()
}

function addSheetPermissions(file_id, editors, owners, viewers) {
  for (var v of viewers) {
    Drive.Permissions.insert(
      { 'value': v,
        'type': 'user',
        'role': 'reader'},
      file_id,
      { 'sendNotificationEmails': 'false'}
    )
  }
  for (var e of editors) {
    Drive.Permissions.insert(
      { 'value': e,
        'type': 'user',
        'role': 'writer'},
      file_id,
      { 'sendNotificationEmails': 'false'}
    )
  }
  for (var o of owners) {
    Drive.Permissions.insert(
      { 'value': o,
        'type': 'user',
        'role': 'owner'},
      file_id,
      { 'sendNotificationEmails': 'false'}
    )
  }
}

////////////////////////// Sheet functions
function newSheet(ssname, parent_folder_id) {
  // create file
  var resource = {
    title: ssname,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: parent_folder_id }]
  }
  var file = Drive.Files.insert(resource)
  return file
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

function formatSheet(sheet, lastcolumn) {
  var borderRange = sheet.getDataRange();
  var headerRange = sheet.getRange(1,1,1,lastcolumn);
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
}

//////////////////////////////////////////////////// MAIN REPORT functions
function report(ssname, parent_folder_id, submitter, startdate, enddate, staging){
  // testing values set as defaults
  ssname = ssname || 'test';
  parent_folder_id = parent_folder_id || form_report_folder_id;
  submitter = typeof submitter !== 'undefined' ? submitter : test_submitter; // <-- set default value only if param is undefined, not just falsey
  // submitter = submitter || test_submitter;
  startdate = startdate || new Date('2022-04-01');
  enddate = enddate || new Date('2022-05-31');
  staging = typeof staging !== 'undefined' ? staging : true; 

  // fixing to use shared resources, locking
  var lock = LockService.getScriptLock();
  try {
      lock.waitLock(30000); // wait 30 seconds for others' use of the code section and lock to stop and then proceed
      console.log("script lock acquired for 'report' function")
  } catch (e) {
      console.error('Could not obtain lock after 30 seconds, exiting');
      return "bookmaker_log script or sheet is busy, could not obtain lock within 30 seconds"
  }

  // get new sheet
  var new_ss = newSheet(ssname, parent_folder_id);
  var new_ss_id = new_ss.id
  var ss_open = SpreadsheetApp.openById(new_ss_id);
  var reportsheet = ss_open.getSheets()[0]
  // get old sheet
  var logsheet = getLogSpreadsheet({}, staging);
  // get/set headers
  var headers = []
  var lc = logsheet.getLastColumn()
  if (lc > 0) {
    headers = logsheet.getRange(1,1,1,lc).getValues()[0];
  }
  reportsheet.getRange(1,1,1,lc).setValues([headers])

  // filter rows from old sheet by date
  const date_index = headers.indexOf('date')
  var non_header_rows = logsheet.getDataRange().getValues();
  non_header_rows.shift(); // get rid of header row
  var filtered_rows = non_header_rows.filter(function(row) {
    var rowdate = row[date_index]
    if (rowdate && rowdate.valueOf() >= startdate.valueOf() && rowdate.valueOf() <= enddate.valueOf()) {
      return row
    }
  })
  // paste valid rows into new sheet
  reportsheet.getRange(2, 1, filtered_rows.length, filtered_rows[0].length).setValues(filtered_rows)

  // format sheet
  formatSheet(reportsheet, lc);
  // change owner of sheet to submitter if this is form entry
  if (submitter != '') { addSheetPermissions(new_ss_id, [], [submitter], []); }
  // send mail
  var rs_url = ss_open.getUrl()
  sendmail(recurring_rprt_resource_sheet_id,ssname,rs_url,startdate,enddate,new_ss_id,staging,submitter)

  // we're done with shared resources, release lock
  lock.releaseLock()
}

/////////////////////////////////////////// Calling main function from a form
function getShortDateStr(date_obj) {
  var short_date_str = (date_obj.getMonth()+1)  + "-" + date_obj.getDate() + "-" + date_obj.getFullYear().toString().substr(-2)
  return short_date_str
}

function getDates(daterange) {
  var sdate = new Date()
  var edate = new Date()
  switch (daterange) {
    case form_range_week:
      sdate = sdate.setDate(sdate.getDate() - 7);
      break;
    case form_range_month:
      sdate = sdate.setMonth(sdate.getMonth() - 1);
      break;
    case form_range_all:
      sdate = new Date('2020-09-09') // oldest date avail from Drive bookmaker logs
      break;
  }
  return {'datestart':new Date(sdate), 'dateend':edate}
}

function runFromForm(e) {
  // get submission info
  var f_submitted = e.response.getItemResponses()
  var f_submitter = e.response.getRespondentEmail()
  var daterange = ''
  var datestart = ''
  var dateend = ''
  for (var i=0; i<f_submitted.length; i++) {
    // console.log("Title: ", f_submitted[i].getItem().getTitle(), ", Item: ", f_submitted[i].getResponse())
    switch (f_submitted[i].getItem().getTitle()) {
      case form_title_daterange:
        daterange = f_submitted[i].getResponse();  // returns a string
        break;
      case form_title_datestart:
        datestart = new Date(f_submitted[i].getResponse());
        break;
      case form_title_dateend:
        dateend = new Date(f_submitted[i].getResponse());
        break;
    }
  }
  if (datestart == '' || dateend == '') {
    var dates_dict = getDates(daterange)
    datestart = dates_dict['datestart']
    dateend = dates_dict['dateend']
  }
  console.info("bmr form submitted by: "+f_submitter+", daterange: "+daterange+", datestart: "+datestart+", dateend: "+dateend+", staging: "+form_staging)
  // RUN REPORT
  var datestart_str = getShortDateStr(datestart)
  var dateend_str = getShortDateStr(dateend)
  var servername = 'bookmaker'
  if (form_staging == true) { servername = 'bookmaker_stg'}
  report(servername+'_report_'+datestart_str+'_to_'+dateend_str, form_report_folder_id, f_submitter, datestart, dateend, form_staging)
}

/////////////////////////////////////////// calling main function from scheduled trigger
function runFromTrigger() {
  // set dates
  var datestart = new Date()
  var dateend = new Date()
  // // 1 week prior:
  // var datestart = new Date(datestart.setDate(datestart.getDate() - 7));
  // // 1 month prior
  var datestart = new Date(datestart.setMonth(datestart.getMonth() - 1));

  console.info("bmr report generated via trigger, datestart: "+datestart+", dateend: "+dateend+", staging: "+recurring_staging)
  // RUN REPORT
  var datestart_str = getShortDateStr(datestart)
  var dateend_str = getShortDateStr(dateend)
  var servername = 'bookmaker'
  if (recurring_staging == true) { servername = 'bookmaker_stg'}
  report(servername+'_report_'+datestart_str+'_to_'+dateend_str, recurring_report_folder_id, '', datestart, dateend, recurring_staging)
}
