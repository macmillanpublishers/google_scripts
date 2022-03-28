const sheet_bm_log_id = 'REDACTED'
const sheet_bm_log_stg_id = 'REDACTED'
const logsheet_name = 'log'

function getLogSpreadsheet(items, staging=false) {
    // set sheet
    var ss = null
    if (staging == true || ('staging' in items && items['staging'] == "True")) {
      ss = SpreadsheetApp.openById(sheet_bm_log_stg_id)
    } else {
      ss = SpreadsheetApp.openById(sheet_bm_log_id)
    }
    var logsheet = ss.getSheetByName(logsheet_name)
  return logsheet
}

function writeEntry(items={}) {
  var rtrn_strng = ''
  //  log/alert if reqrd params missins
  if(Object.keys(items).length) {
    // get sheet
    var logsheet = getLogSpreadsheet(items);
    // get/set headers
    var headers = []
    var lc = logsheet.getLastColumn()
    if (lc > 0) {
      headers = logsheet.getRange(1,1,1,lc).getValues()[0];
    }
    for (var i in items) {
      // console.log('item: ', i, ". headers: ", headers, ". lc:", lc)
      if (i != 'staging' && !headers.includes(i)) {
        console.warn("parameter: "+i+" was not in existing list of keys for this report. A new column was added with name: "+i)
        logsheet.getRange(1, lc + 1).setValue(i);
        headers.push(i)
        lc+=1;
      }
    } 
    // add values to match headers on new row
    var newrow_index = logsheet.getLastRow()+1
    for (var j in items) {
      if (j != 'staging') {
        var param_val = items[j]
        // get index of header 'j'
        var thisheader_index = headers.indexOf(j)+1
        // set value
        logsheet.getRange(newrow_index, thisheader_index).setValue(param_val);
      }    
    }
    rtrn_strng = "New entry added! Params: "+JSON.stringify(items)
    console.info(rtrn_strng);
  } else {
    rtrn_strng = "entry has no params or key item is missing; unable to write entry"
    console.error(rtrn_strng);
  }
  return rtrn_strng
}

function main(items={}) {
  // // for test/debug, run with a mocked items dict as needed
  // items = {'staging': "True", 'key1':'value12', 'key2':'value22'}
  var lock = LockService.getScriptLock();
  try {
      lock.waitLock(30000); // wait 30 seconds for others' use of the code section and lock to stop and then proceed
  } catch (e) {
      console.error('Could not obtain lock after 30 seconds, exiting');
      return "bookmaker_log script or sheet is busy, could not obtain lock within 30 seconds"
  }
  var output = writeEntry(items)
  lock.releaseLock();
  return output
}

function helloWorld(dict={}) {
  console.log("Hello, world!", sheet_bm_log_stg_id);
  
  return "hi" + " " + dict['bestpet']
}