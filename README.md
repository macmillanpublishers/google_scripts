# google_scripts
Scripts for Google Apps

# To Use
Create a new Google Document / Sheet / Etc. Click on **Tools > Script editor** to open the Script editor. Paste the code in the code window and save.

You can [import/export files](https://developers.google.com/apps-script/import-export) via the Google Drive API but we'll have to figure that out in time.

# auto_report_mailer
This script is owned by the Macmillan workflows account.
It is executed via a timed trigger, set for Monday mornings.
It takes an uploaded .tsv file in Drive (containing info on the prior week's egalleymaker runs), formats and renames it, and shares the updated Drive sheet via custom email.
For more details see internal Confluence documentation [here](https://confluence.macmillan.com/pages/viewpage.action?pageId=42868497).

# bookmaker_log_and_report
This set of script modules serves two purposes:
1. enable bookmaker to log all runs to a google sheet (log_bookmaker_run.js module). The 'main' function is invoked via google-appscript api by the product, per run.
2. enable 'pretty' reports for specified date ranges, against the bookmaker log google-sheet (bookmaker_report.js module). Can be invoked via a form (onSubmit trigger > runFromForm function) or via a scheduled trigger (pointing to runFromTrigger function).
### dev-notes
- Note: id's to specific google resources and email addresses are marked here as REDACTED.
- Bookmaker logging api is pointed at a script deployment (the only active one); currently all report triggers are pointed at the 'head' (live script in Script Editor).
- Test reports can be run directly via the 'report' function.
- Staging/testing resources can be set by constants at the top of the reporting module for reports... for logging the 'staging' value comes from input param of the same name passed to 'main'.

# image_rights_form-onsubmit
This script is owned by the Macmillan workflows account.
It is tied to a 'Master' Spreadsheet, which is in turn linked to the Macmillan Image-Rights Google Form: every time a submission is made to the form, the onSubmit function is run.
It cleans up the new data and recreates it in a new sheet by imprint, and sends a notification to the designated watcher for that imprint.
For more details see internal Confluence documentation [here](https://confluence.macmillan.com/display/PWG/Image+Rights+form).

# event-registrationer
An updated version of [this](https://developers.google.com/apps-script/quickstart/forms).
## Purpose
* Creates a series of calendar events based on info in a Google Sheet
* Creates a sign-up form for those same events
* Adds people as guests to the event when they submit the form
* Sends a confirmation email when someone signs up
* Limits events to 24 people (because Hangouts are limited to 25 people, including the host)

## Setup
1. Create a new Google Sheet file
1. Rename the primary sheet `DemoSetup`
1. Add the following column heads and information:

| Column | Heading         | Contains                                   |
|--------|-----------------|--------------------------------------------|
| A      | 'Session Title' | Title that will be added to calendar event |
| B      | 'Date'          | Day of event, format: MM/DD/YYYY           |
| C      | 'Start Time'    | Time event starts, format: HH:MM AM        |
| D      | 'End Time'      | Time event ends, format: HH:MM AM          |
| E      | 'Location'      | Text to appear in Location event field.    |
| F      | 'Event'         | Leave blank                                |
| G      | 'Attendees'     | Enter `0` for each event                   |

1. Go to *Tools > Script editor*.
1. Delete any code that is there and add the code in the `event-registrationer.js` file.
1. Edit the calendar event description in the `setUpCalendar_` function.
1. Edit the form title in the `setUpForm_` function.
1. If you want to have a different cap on sign ups for each session, edit the `updateFormOptions` function.
1. Save the code pane.
1. Return to the main sheet and reload the page.
1. Click the **Event Manager** menu item that appears and select **Set up events**.
1. You may be prompted to authorize the app. Follow the prompts.

A message will pop up to let you know the code is running. When it's done, you should have new Calendar Events for each of your rows, a new sheet tab called 'Form Responses 1', and (if you select that tab and then go to **Form > Go to live form**) a form for people to fill out.

Before distributing the form, open each new calendar event and add a Video Call if you would like to do a Google Hangout for that meeting.

You may also want to go to the script editor again and click on . There is a new trigger called "onFormSubmit". By default it will send you an email once a day with a summary of any errors from the last 24 hours. If you would like to change this notification default, click on **Resources > Current project triggers** and select the **notifications** link. Then select your preferences and click Save.

Now you can distribute the form, and people who sign up will automatically be added to the calendar events and get an email confirmation.

Note that the creator of the form will also receive a copy of the email sent to each person who registers. If you don't want these filling up your inbox, create a Gmail filter before distributing the form.

## To do
Still kinda rough around the edges. Some key improvements:

- [ ] Add event description as a field in the setup sheet
- [ ] Add form title as a field in the setup sheet
- [ ] Change `Attendees` column to be *max* number of people allowed, change code to get current number of guests in the event and only allow someone to register if it's less than max in sheet (this should also take care of people removing themselves from the event)
- [ ] Automatically add Hangouts link
- [ ] Better handling in form for when a session is full
- [ ] Possible to book actual conference rooms? Would also have to check availability during setup
- [ ] If can book conference rooms, see if we can automatically add max number of seats
- [ ] Validate user-input email addresses (or pull from user info, not a text field)
- [ ] Add meeting duration or end time to the listing on the Form.
- [ ] Some way to not get a bunch of emails every time someone signs up for a session
- [ ] Another setup function to format the initial sheet and get other info that is currently manual
