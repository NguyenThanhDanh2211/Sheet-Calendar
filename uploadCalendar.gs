/**
 * Upload TKB từ sheet lên calendar
 */

function uploadScheduleToCalendar() {
  var calendarId = Session.getActiveUser().getEmail();
  var calendar = CalendarApp.getCalendarById(calendarId);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule');
  var data = sheet
    .getRange(7, 1, sheet.getLastRow() - 5, sheet.getLastColumn())
    .getValues();

  // Clear existing events within the date range to avoid duplicates
  var startDate = new Date('2024-05-20');
  var endDate = new Date('2024-06-30');
  var existingEvents = calendar.getEvents(startDate, endDate);
  existingEvents.forEach(function (event) {
    event.deleteEvent();
  });

  data.forEach(function (row) {
    var weekNumber = row[0];
    var startDateStr = row[1];
    var endDateStr = row[2];
    var timeRange = row[3];

    if (
      startDateStr &&
      endDateStr &&
      timeRange &&
      typeof startDateStr === 'string' &&
      typeof endDateStr === 'string' &&
      typeof timeRange === 'string'
    ) {
      var startDate = parseDate(startDateStr);
      var endDate = parseDate(endDateStr);
      var timeParts = timeRange.split(' - ');
      if (timeParts.length === 2) {
        var startTime = parseTime(timeParts[0]);
        var endTime = parseTime(timeParts[1]);

        for (var i = 4; i <= 9; i++) {
          if (row[i]) {
            var eventTitle = row[i];
            var eventDate = new Date(startDate);
            eventDate.setDate(startDate.getDate() + (i - 4));

            var startDateTime = new Date(eventDate);
            startDateTime.setHours(startTime.hours, startTime.minutes);

            var endDateTime = new Date(eventDate);
            endDateTime.setHours(endTime.hours, endTime.minutes);

            calendar.createEvent(eventTitle, startDateTime, endDateTime);
          }
        }
      }
    }
  });
}

// Hàm để phân tích ngày từ chuỗi định dạng 'dd/MM/yyyy'
function parseDate(dateStr) {
  if (typeof dateStr !== 'string') {
    throw new TypeError(
      'Expected a string in parseDate, but got ' + typeof dateStr
    );
  }
  var parts = dateStr.split('/');
  return new Date(parts[2], parts[1] - 1, parts[0]);
}

// Hàm để phân tích thời gian từ chuỗi định dạng 'hh:mm'
function parseTime(timeStr) {
  if (typeof timeStr !== 'string') {
    throw new TypeError(
      'Expected a string in parseTime, but got ' + typeof timeStr
    );
  }
  var parts = timeStr.split(':');
  return { hours: parseInt(parts[0], 10), minutes: parseInt(parts[1], 10) };
}
