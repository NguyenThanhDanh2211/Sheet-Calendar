/**
 * Tạo thời khóa biểu mẫu từ CTDT, đưa TKB lên calendar
 */

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();

  var mainMenu = ui.createMenu('Công cụ quản lý sheet-calendar');

  // Quản lý TKB trên Calendar
  var calendarMenu = ui
    .createMenu('Quản lý TKB trên Calendar')
    .addItem(
      'Tạo TKB mẫu theo Học kỳ (->Calendar)',
      'Show_Sidebar_createCalendar'
    );

  // Quản lý TKB trên Sheet
  var sheetMenu = ui
    .createMenu('Quản lý TKB trên Sheet')
    .addItem('Download thời khóa biểu (Calendar->Sheet)', 'downloadCalendar');

  // Sub-menu cho uploadToCalendar
  var uploadMenu = ui
    .createMenu('Đưa thời khóa biểu có sẵn lên Calendar (Sheet->Calendar)')
    .addItem('5buổi/week 2h/buổi', 'uploadToCalendar')
    .addItem('2-4-6 3h/buổi', 'uploadWeeklyToCalendar')
    .addItem('3-5-7 3h/buổi', 'uploadDailyToCalendar');

  // Thêm sub-menu vào menu "Quản lý TKB trên Sheet"
  sheetMenu.addSubMenu(uploadMenu);

  // Thêm các menu con vào menu chính
  mainMenu.addSubMenu(calendarMenu).addSubMenu(sheetMenu).addToUi();
}

// Dummy functions for the new menu items
function uploadWeeklyToCalendar() {
  SpreadsheetApp.getUi().alert('Upload TKB cho từng tuần');
}

function uploadDailyToCalendar() {
  SpreadsheetApp.getUi().alert('Upload TKB cho từng ngày');
}

function createEvent(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('course');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet "course" không tồn tại.');
    return;
  }

  var data = sheet.getDataRange().getValues();

  var calendarId = Session.getActiveUser().getEmail();
  var calendar = CalendarApp.getCalendarById(calendarId);

  var startDate = new Date(form.startDate + 'T' + form.startTime + ':00');

  var hoursPerDay =
    (new Date('1970-01-01T' + form.endTime + ':00') -
      new Date('1970-01-01T' + form.startTime + ':00')) /
    (1000 * 60 * 60);

  // Define the days of the week for classes based on hours per day
  var classDays;
  if (hoursPerDay == 2) {
    classDays = [1, 2, 3, 4, 5]; // Monday to Friday (T2 to T6)
  } else if (hoursPerDay == 3) {
    classDays = form.classDays === 'MWF' ? [1, 3, 5] : [2, 4, 6]; // T2, T4, T6 or T3, T5, T7
  } else {
    throw new Error(
      'Invalid hours per day. Only 2 or 3 hour classes are supported.'
    );
  }

  // Iterate over each course
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == form.semester) {
      var subject = data[i][2];
      var totalHours = data[i][3];
      var sessionCount = 1;

      // Hours already scheduled
      var hoursScheduled = 0;

      // Repeat until the total hours for the course are scheduled
      while (hoursScheduled < totalHours) {
        // Check if the current day is one of the specified class days
        if (classDays.includes(startDate.getDay())) {
          // Create a calendar event for the class
          var eventEnd = new Date(startDate);
          eventEnd.setHours(
            eventEnd.getHours() +
              (parseInt(form.endTime.split(':')[0]) -
                parseInt(form.startTime.split(':')[0]))
          );
          eventEnd.setMinutes(
            eventEnd.getMinutes() +
              (parseInt(form.endTime.split(':')[1]) -
                parseInt(form.startTime.split(':')[1]))
          );

          var eventTitle = subject + '_TL' + ('0' + sessionCount).slice(-2);

          calendar.createEvent(eventTitle, startDate, eventEnd);

          // Update the hours scheduled
          hoursScheduled += hoursPerDay;
          sessionCount++;
        }

        // Move to the next day
        startDate.setDate(startDate.getDate() + 1);
      }
    }
  }
}

function Show_Sidebar_createCalendar() {
  var template = HtmlService.createTemplateFromFile('Sidebar_createCalendar');
  var html = template.evaluate();
  SpreadsheetApp.getUi().showSidebar(html);
}
