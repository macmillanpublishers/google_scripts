/**
 * A special function that inserts a custom menu when the spreadsheet opens.
 */
function onOpen() {
  var menu = [{name: 'Set up events', functionName: 'setUpConference_'}];
  SpreadsheetApp.getActive().addMenu('Event Manager', menu);
}

/**
 * A set-up function that uses the conference data in the spreadsheet to create
 * Google Calendar events, a Google Form, and a trigger that allows the script
 * to react to form responses.
 */
function setUpConference_() {
  if (ScriptProperties.getProperty('calId')) {
    Browser.msgBox('Your event is already set up. Look in Google Drive!');
  }
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('DemoSetup');
  var range = sheet.getDataRange();
  var values = range.getValues();
  setUpCalendar_(values, range);
  setUpForm_(ss, values);
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit()
      .create();
  ss.removeMenu('Event Manager');
}

/**
 * Creates a Google Calendar with events for each conference session in the
 * spreadsheet, then writes the event IDs to the spreadsheet for future use.
 *
 * @param {String[][]} values Cell values for the spreadsheet range.
 * @param {Range} range A spreadsheet range that contains conference data.
 */
function setUpCalendar_(values, range) {
  var cal = CalendarApp.getDefaultCalendar();
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    var title = session[0];
    var start = joinDateAndTime_(session[1], session[2]);
    var end = joinDateAndTime_(session[1], session[3]);
    var options = {location: session[4], sendInvites: true};
    var description = "This demo will show you how to set up the Egalleymaker folder and submit a manuscript, and will review the types of notifications you may receive.";
    var event = cal.createEvent(title, start, end, options)
    .setGuestsCanSeeGuests(true) .setGuestsCanInviteOthers(false) .setDescription(description)
    .addPopupReminder(15);
    // how to add a video call?
    session[5] = event.getId();
  }
  range.setValues(values);

  // Store the ID for the Calendar, which is needed to retrieve events by ID.
  ScriptProperties.setProperty('calId', cal.getId());
}

/**
 * Creates a single Date object from separate date and time cells.
 *
 * @param {Date} date A Date object from which to extract the date.
 * @param {Date} time A Date object from which to extract the time.
 * @return {Date} A Date object representing the combined date and time.
 */
function joinDateAndTime_(date, time) {
  date = new Date(date);
  date.setHours(time.getHours());
  date.setMinutes(time.getMinutes());
  return date;
}

/**
 * Creates a Google Form that allows respondents to select which conference
 * sessions they would like to attend, grouped by date and start time.
 *
 * @param {Spreadsheet} ss The spreadsheet that contains the conference data.
 * @param {String[][]} values Cell values for the spreadsheet range.
 */
function setUpForm_(ss, values) {
  // Create the form and add a multiple-choice question for each timeslot.
  var form = FormApp.create('Egalley Demo Signup');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  form.addTextItem().setTitle('Name').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true);
  var item = form.addMultipleChoiceItem().setTitle('Choose one available session');

  updateFormOptions(item, values);
}

// Create and update the option buttons on the form
// Limits sign ups to 24 people
function updateFormOptions(item, values) {
  // create text string for each option button, add to an array of choices
  var allChoices = [];
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    var optionText = createOptionString(session);
    var attendees = session[6];
    var openSeats = 24 - attendees;
    if (openSeats < 1) {
      optionText = "SESSION FULL! " + optionText;
    };

    // create 'choice' object and add to array
    allChoices.push(optionText);
  }

  // update option buttons on form
  item.setChoiceValues(allChoices);
}

// Create the full string for the option button from the original data
function createOptionString(session) {
  var day = session[1].toLocaleDateString();
  var time = session[2].toLocaleTimeString();
  var optionText = day + ' ' + time;

  return optionText;
}

/**
 * A trigger-driven function that sends out calendar invitations and a
 * personalized Google Docs itinerary after a user responds to the form.
 *
 * @param {Object} e The event parameter for form submission to a spreadsheet;
 *     see https://developers.google.com/apps-script/understanding_events
 */
function onFormSubmit(e) {
  var user = {name: e.namedValues['Name'][0], email: e.namedValues['Email'][0]};

  // Grab the session data again so that we can match it to the user's choices.
  var origSheet = e.range.getSheet().getParent().getSheetByName('DemoSetup');
  var response = [];
  var values = origSheet.getDataRange().getValues();
  var submitOK = true;
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    var title = 'Choose one available session';
    var timeslot = createOptionString(session);

    // For every selection in the response, find the matching timeslot and title
    // in the spreadsheet and add the session data to the response array.
    if (e.namedValues[title] && e.namedValues[title] == timeslot) {
      Logger.log('Registered for: ' + timeslot);
      // check attendee total again
      var currentTotal = session[6];
      if (currentTotal < 24) {
        // add session to response
        response.push(session);
        // add 1 to attendee total
        var rowNum = i + 1;
        var cell = origSheet.getRange(rowNum, 7, 1, 1);
        cell.setValue(currentTotal + 1);
      } else {
        submitOK = false;
      };
    } else {
      submitOK = false;
    };

    if (submitOK = false) {
      //alert user that session is full
      // set close message
      var msg = "This session is full, please reload the form and try again.";
      form.setAcceptingResponses(false).setCustomClosedFormMessage(msg);
    };
  };

    if (response.length < 1) {
      MailApp.sendEmail({
        to: 'erica.warren@macmillan.com',
        subject: 'google app error',
        body: 'Something did not work: ' + user.email + ' sheet: ' + origSheet.getName() + 'last timeslot: ' + timeslot,
      });
    }
  sendInvites_(user, response);
  sendDoc_(user, response);
}

/**
 * Add the user as a guest for every session he or she selected.
 *
 * @param {Object} user An object that contains the user's name and email.
 * @param {String[][]} response An array of data for the user's session choices.
 */
 function sendInvites_(user, response) {
   var id = ScriptProperties.getProperty('calId');
   var cal = CalendarApp.getCalendarById(id);
   for (var i = 0; i < response.length; i++) {
     cal.getEventSeriesById(response[i][5]).addGuest(user.email);
   }
 }

/**
 * .addGuest() can't send email notification, so we'll send our own.
 *
 * @param {Object} user An object that contains the user's name and email.
 * @param {String[][]} response An array of data for the user's session choices.
 */
function sendDoc_(user, response) {
  // var doc = DocumentApp.create('Registration successful!');
  //     // .addEditor(user.email);
  // var body = doc.getBody();
  // var table = [['Session', 'Date', 'Time']];
  // for (var i = 0; i < response.length; i++) {
  //   table.push([response[i][0], response[i][1].toLocaleDateString(),
  //       response[i][2].toLocaleTimeString()]);
  // }
  // body.insertParagraph(0, doc.getName())
  //     .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  // table = body.appendTable(table);
  // table.getRow(0).editAsText().setBold(true);
  // doc.saveAndClose();
  var sessions = 'REGISTERED SESSIONS:\n';
  for (var i = 0; i < response.length; i++) {
    var timeslot = createOptionString(response[i]);
    sessions = sessions + response[i][0] + ': ' + timeslot + '\n';
  }

  // Email a link to the Doc as well as a PDF copy.
  MailApp.sendEmail({
    to: user.email,
    subject: 'Registration successful!',
    body: 'Hi ' + user.name + ',\nThanks for signing up! The following sessions have been added to your calendar: ' + '\n\n' + sessions + '\nWhen the event starts, simply open it in Google Calendar and click the link next to \"Video call\" to join the Hangout.',
    // attachments: doc.getAs(MimeType.PDF),
  });
}
