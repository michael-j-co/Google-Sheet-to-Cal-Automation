// Constants for configuration
const SPREADSHEET_ID = '1NMWNUQILbWR1NONJzqxGuZeQ1ERezixnNNcI5skl-5w';
const USER_NAME = 'Michael C.';
const CALENDAR_NAME = 'Work Shifts';
const SECONDARY_LOCATION_ROWS = 2; // Last two rows for special shifts.

function checkForNewSheetAndUpdateCalendar() {
  const sheet = getSecondToLastSheet();
  if (!sheet) {
    Logger.log("No available sheet to process.");
    return;
  }

  const calendar = getOrCreateCalendar(CALENDAR_NAME);
  const data = sheet.getDataRange().getValues();
  const { days, dates, shiftTimes } = parseHeaderData(data);

  for (let i = 0; i < shiftTimes.length; i++) {
    for (let j = 0; j < days.length; j++) {
      const shiftDetails = data[i + 3][j + 1];
      if (shiftDetails && shiftDetails.includes(USER_NAME)) {
        const eventDetails = createEventDetails(dates[j], days[j], shiftTimes[i], i, shiftTimes.length);
        const uniqueId = `${sheet.getSheetName()} - ${eventDetails.startDateTime.toISOString()} - ${eventDetails.endDateTime.toISOString()}`;
        
        // Check for existing event before creating
        if (isEventAlreadyCreated(calendar, eventDetails.startDateTime, eventDetails.endDateTime, uniqueId)) {
          Logger.log(`Event already exists: ${uniqueId}`);
          continue;
        }

        addEventToCalendar(calendar, eventDetails, sheet.getSheetName(), uniqueId);
      }
    }
  }
}

// Helper function to check if an event with the unique identifier already exists within a date range
function isEventAlreadyCreated(calendar, startDateTime, endDateTime, uniqueId) {
  const events = calendar.getEvents(startDateTime, endDateTime);
  for (const event of events) {
    if (event.getDescription().includes(uniqueId)) {
      return true;
    }
  }
  return false;
}

function getSecondToLastSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = spreadsheet.getSheets();
  return sheets.length >= 2 ? sheets[sheets.length - 2] : null;
}

function getOrCreateCalendar(calendarName) {
  let calendar = CalendarApp.getCalendarsByName(calendarName)[0];
  return calendar || CalendarApp.createCalendar(calendarName);
}

function parseHeaderData(data) {
  const days = data[0].slice(1, 8);
  const dates = data[1].slice(1, 8);
  const shiftTimes = data.slice(3).map(row => row[0]);
  return { days, dates, shiftTimes };
}

function createEventDetails(date, day, shiftTime, shiftIndex, shiftTimesLength) {
  const [startTime, endTime] = shiftTime.split(' - ');
  const startDateTime = parseDateTime(date, startTime);
  const endDateTime = parseDateTime(date, endTime, startDateTime);

  const isSpecialShift = (shiftIndex >= shiftTimesLength - SECONDARY_LOCATION_ROWS);
  const location = isSpecialShift ? 'Tipuana MPR' : 'Gayley Heights';

  return { day, startDateTime, endDateTime, location };
}

function parseDateTime(date, time, startDateTime = null) {
  const dateTime = new Date(date);
  const [hour, minute] = time.match(/\d+/g);
  const isPM = time.includes('PM');

  dateTime.setHours(parseInt(hour) + (isPM && hour != 12 ? 12 : 0));
  dateTime.setMinutes(parseInt(minute));

  if (startDateTime && dateTime <= startDateTime) {
    dateTime.setDate(dateTime.getDate() + 1);
  }
  return dateTime;
}

function addEventToCalendar(calendar, { day, startDateTime, endDateTime, location }, sheetName, uniqueId) {
  calendar.createEvent(`${day} Shift - ${sheetName}`, startDateTime, endDateTime, {
    location,
    description: `Work shift for ${USER_NAME} - ${sheetName}\nUnique ID: ${uniqueId}`,
  });
  Logger.log(`Event created for ${day}: ${startDateTime} - ${endDateTime} at ${location}`);
}
