// Global settings
const SETTINGS = {
  reminderDaysAfter: 8,
  followUpEmails: 3,
  statusUpdateColumn: 'F', // Column for warning messages
  emailSentColumn: 'H', // Column for the number of inquiries
  applicationDateColumn: 'E', // Column for the application date
  companyNameColumn: 'A', // Column for the company name
  positionNameColumn: 'B', // Column for the position name
  contactNameColumn: 'I', // Column for the contact name
  contactEmailColumn: 'J', // Column for the contact email
  languageColumn: 'D', // Column for the language
  followUpMessage: 'The position has likely been filled',
  followUpBackground: '#ea4335', // Red background
  responseReceivedBackground: '#00ff00', // Green background
  applicationMethodColumn:  'C',
};

// Function to create triggers
function createTriggers() {
  // First, delete all existing daily triggers to avoid duplicates
  let existingTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < existingTriggers.length; i++) {
    if (existingTriggers[i].getHandlerFunction() === 'checkApplications') {
      ScriptApp.deleteTrigger(existingTriggers[i]);
    }
  }

  // Create the midnight trigger
  ScriptApp.newTrigger('checkApplications')
    .timeBased()
    .everyDays(1)
    .atHour(0) // Midnight
    .create();

  // Create the noon trigger
  ScriptApp.newTrigger('checkApplications')
    .timeBased()
    .everyDays(1)
    .atHour(12) // Noon
    .create();
}

// Function to run when the spreadsheet is opened, which adds trigger creation to the menu
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Check Applications', 'checkApplications')
    .addItem('Generate Reports', 'generateReports')
    .addItem('Create Triggers', 'createTriggers') // Add to the menu
    .addToUi();
}

// Function to check applications and create reminders
function checkApplications() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let lastRow = sheet.getLastRow();
  let dataRange = sheet.getRange('A2:J' + lastRow);
  let data = dataRange.getValues();
  let backgrounds = dataRange.getBackgrounds();
  let currentDate = new Date();
  currentDate.setHours(0, 0, 0, 0); // Set the time to midnight to consider only the date
  let today = new Date(currentDate); // Define the 'today' variable for use

  let sixMonthsAgo = new Date();
  sixMonthsAgo.setMonth(today.getMonth() - 6);

  for (let i = 0; i < data.length; i++) {
    let row = data[i];
    if (!row || row.length === 0 || row[SETTINGS.applicationDateColumn.charCodeAt(0) - 'A'.charCodeAt(0)] === undefined) {
      Logger.log('The row is empty or undefined, or the application date is missing: ' + (i + 2));
      continue;
    }
    let applicationMethod = row[SETTINGS.applicationMethodColumn.charCodeAt(0) - 'A'.charCodeAt(0)];
    let isEmailApplication = applicationMethod === 'e-mail'; // Check if the application method is E-mail
    let cellDate = new Date(row[SETTINGS.applicationDateColumn.charCodeAt(0) - 'A'.charCodeAt(0)]);
    cellDate.setHours(0, 0, 0, 0);
    let daysDiff = Math.floor((currentDate - cellDate) / (1000 * 3600 * 24));
    let cellColor = backgrounds[i][SETTINGS.applicationDateColumn.charCodeAt(0) - 'A'.charCodeAt(0)];
    let requiredColor = '#fbbc04'; // The color indicating a follow-up is needed

    let eventName = 'Follow-up for ' + row[SETTINGS.companyNameColumn.charCodeAt(0) - 'A'.charCodeAt(0)];
    let events = CalendarApp.getDefaultCalendar().getEvents(sixMonthsAgo, currentDate);
    let eventExists = events.some(function(event) {
      return event.getTitle() === eventName && 
             event.getStartTime().getTime() === cellDate.getTime();
    });

    let emailSent = backgrounds[i][SETTINGS.emailSentColumn.charCodeAt(0) - 'A'.charCodeAt(0)] === SETTINGS.followUpBackground;
    let responseReceived = backgrounds[i][SETTINGS.statusUpdateColumn.charCodeAt(0) - 'A'.charCodeAt(0)] === SETTINGS.responseReceivedBackground;

    if (!eventExists && !emailSent && !responseReceived && daysDiff >= SETTINGS.reminderDaysAfter && cellColor === requiredColor && isEmailApplication) {
      createCalendarEvent(row, i + 2);
      createEmailDraft(row, i + 2);
    }
  }
}

function createCalendarEvent(row, index) {
  let calendarId = 'ASD@gmail.com'; // The calendar ID of your Google Calendar
  let currentDate = new Date();
  currentDate.setHours(0, 0, 0, 0); // Set the time to midnight to only take the date
  let today = new Date(currentDate);
  let calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    throw new Error('No calendar found with the given id: ' + calendarId);
  }
  
  let applicationDate = new Date(row[SETTINGS.applicationDateColumn.charCodeAt(0) - 'A'.charCodeAt(0)]);
  let reminderDate = new Date(applicationDate);
  reminderDate.setDate(reminderDate.getDate() + SETTINGS.reminderDaysAfter);
  
  let sixMonthsAgo = new Date();
  sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);

  // Set the search range from the last date to the end of the next business day
  let searchEndDate = getNextWorkday(reminderDate);
  searchEndDate.setHours(23, 59, 59, 999); // A nap végéig keressünk

  let events = calendar.getEvents(sixMonthsAgo, searchEndDate);

//Logging the found events - just in case. Not necessary.
// events.forEach(function(event) {
//  Logger.log('Event title: ' + event.getTitle());
//  Logger.log('Event start: ' + event.getStartTime());
//});

  let eventName = 'Follow-up for ' + row[SETTINGS.companyNameColumn.charCodeAt(0) - 'A'.charCodeAt(0)];
  
// Check if the event already exists
let existingEvent = events.find(function(event) {
  let eventStart = event.getStartTime().getTime();
  let reminderStart = getNextWorkday(reminderDate).getTime();
  return event.getTitle() === eventName && eventStart === reminderStart;
});

if (existingEvent) {
  Logger.log('An event already exists for: ' + eventName + ' on ' + reminderDate.toDateString());
} else {
 // Create the new event here if it does not exist
  let nextWorkday = getNextWorkday(reminderDate);
  let event = calendar.createEvent(eventName, nextWorkday, new Date(nextWorkday.getTime() + (30 * 60 * 1000)), {
    description: 'Check the application status for ' + (row[SETTINGS.positionNameColumn.charCodeAt(0) - 'A'.charCodeAt(0)] || 'Unknown Position')
  });
  Logger.log('Event created: ' + event.getTitle() + ' at: ' + nextWorkday);
  }
}

function getNextWorkday(date) {
  let nextWorkday = new Date(date);
  nextWorkday.setDate(nextWorkday.getDate() + 1); // Next day
  if (nextWorkday.getDay() === 6) { // If it is Saturday
    nextWorkday.setDate(nextWorkday.getDate() + 2); // Change to Monday - we don't work in weekends :)
  } else if (nextWorkday.getDay() === 0) { // Or it is Sunday
    nextWorkday.setDate(nextWorkday.getDate() + 1); // Change to Monday - we don't work in weekends :)
  }
  nextWorkday.setHours(8, 0, 0, 0); // Set to 8am
  return nextWorkday;
}

// Create e-mail draft -->Can you send the emails automatically, but it is a risk. I decided to send it myself after double check.
//There is 2 language that I use for e-mails: English and Hungarian.
function createEmailDraft(row, index) {
  let emailColumnIndex = SETTINGS.contactEmailColumn.charCodeAt(0) - 'A'.charCodeAt(0); 
  let emailAddress = row[emailColumnIndex] || 'iNeedAnEmailAddress'; 
  let contactName = row[SETTINGS.contactNameColumn.charCodeAt(0) - 65]; 
  let language = row[SETTINGS.languageColumn.charCodeAt(0) - 65].toLowerCase();
  let positionName = row[SETTINGS.positionNameColumn.charCodeAt(0) - 65]; 
  let applicationDate = new Date(row[SETTINGS.applicationDateColumn.charCodeAt(0) - 'A'.charCodeAt(0)]);
  let formattedDate = applicationDate.getFullYear() + '-' + 
                      ('0' + (applicationDate.getMonth() + 1)).slice(-2) + '-' + 
                      ('0' + applicationDate.getDate()).slice(-2); // YYYY-MM-DD date format

  // If the position name is not specified, the application message is made generic - just for Hungarian language
  let applicationMessage = positionName
    ? 'E-mailben jelentkeztem Önöknél a ' + positionName + ' '+ 'pozícióra ' + applicationDate + '-án/-én.'
    : 'E-mailben jelentkeztem Önöknél ' + applicationDate + '-án/-én.';

  // If positionName is empty, set the default value depending on the language
  if (positionName === '') {
    positionName = language === 'hu' ? 'megpályázott' : 'the position';
  }

// Setting the subject and text of the letter based on language
  let subject, message;
  if (language === 'hu') {
    subject = positionName ? 'Érdeklődés a ' + positionName + ' pozíció kapcsán' : 'Érdeklődés';
    message = 'Kedves ' + (contactName || 'Hölgyem/Uram') + '!\n\n' +
              applicationMessage + ' Szeretnék érdeklődni, hogy áll a kiválasztási folyamat.\n\n' +
              'Amennyiben további információra vagy adatokra van szükségük részemről, kérem, jelezze felém. Remélem, hogy hamarosan beszélünk.\n\n' +
              'Köszönöm szíves válaszát!\n\n' +
              '[YOUR NAME]\n' +
              '[YOUR PHONE NUMBER]';
  } else {
    subject = positionName ? 'Inquiry about my application for the ' + positionName : 'Inquiry about my application';
    message = 'Dear ' + (contactName || 'Sir/Madame') + ',\n\n' +
              'I am writing to inquire about the status of my application for the ' + 
              (positionName ? positionName + ' position, submitted on ' : 'position submitted on ') + 
              formattedDate + '.\n\n' +
              'If you need any further information, please let me know. I look forward to your response.\n\n' +
              'Thank you for your time.\n\n' +
              '[YOUR NAME]\n' +
              '[YOUR PHONE NUMBER]';
  }
  let drafts = GmailApp.getDrafts();
  let draftExists = drafts.some(function(draft) {
  let draftInfo = draft.getMessage();
    return draftInfo.getTo() === emailAddress && draftInfo.getSubject() === subject;
  });

  if (!draftExists) {
    GmailApp.createDraft(row[SETTINGS.contactEmailColumn.charCodeAt(0) - 65], subject, message);
    Logger.log('Email draft created for: ' + emailAddress);
  } else {
    Logger.log('Draft already exists for: ' + emailAddress);
  }
}

// Generate reports
function generateReports() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let dataRange = sheet.getDataRange(); // Az összes adatot tartalmazó tartomány
  let lastRow = sheet.getLastRow();
  let chartSheetName = 'Reports';
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let chartSheet = spreadsheet.getSheetByName(chartSheetName);

// If there is no 'Reports' tab, create one
  if (!chartSheet) {
    chartSheet = spreadsheet.insertSheet(chartSheetName);
  } else {
  // Delete existing charts so we can update them
    let charts = chartSheet.getCharts();
    for (let i = 0; i < charts.length; i++) {
      chartSheet.removeChart(charts[i]);
    }
  }

// For example, if we want a chart that shows the number of interested messages
  let emailSentColumn = SETTINGS.emailSentColumn + '2:' + SETTINGS.emailSentColumn + lastRow;
  let emailSentRange = sheet.getRange(emailSentColumn);
  let emailSentValues = emailSentRange.getValues();

// Count the messages of inquiry 
  let emailSentCounts = emailSentValues.reduce(function (countMap, row) {
  let count = row[0] || 0;
  countMap[count] = (countMap[count] || 0) + 1;
    return countMap;
  }, {});

  // Convert the numbers to a range for the chart
  let chartData = [['Number of the messages of inquiry', 'Number of messages']];
  for (let key in emailSentCounts) {
    chartData.push([key, emailSentCounts[key]]);
  }

// Create the range for the chart data
  let chartRange = chartSheet.getRange(1, 1, chartData.length, 2);
  chartRange.setValues(chartData);

// Create the chart
  let pieChart = chartSheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(chartRange)
    .setPosition(2, 3, 0, 0)
    .setOption('title', 'Number of the messages of inquiry')
    .build();

  chartSheet.insertChart(pieChart);
}

function updateStatus(row, index, sheet) {
// For example, if the 'Got any Feedback?' column is filled
  if (row[15]) {
    sheet.getRange('A' + (index + 2) + ':Q' + (index + 2)).setBackground('#00ff00'); // Green background
    sheet.getRange('F' + (index + 2)).setValue('Feedback has been given');
  }
}
