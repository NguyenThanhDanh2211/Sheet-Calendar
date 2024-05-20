function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Công cụ quản lý sheet-calendar')
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu('Quản lý lịch Calendar')
        .addItem('Xếp thời khóa biểu tự động', 'Show_Sidebar_createCalendar')
        .addItem('Download thời khóa biểu', 'downloadCalendar')
    )
    .addToUi();
}

function createEvent(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('course');
  var data = sheet.getDataRange().getValues();

  var calendarId = Session.getActiveUser().getEmail();
  var calendar = CalendarApp.getCalendarById(calendarId);

  var startDate = new Date(form.startDate + 'T' + form.startTime + ':00'); // Bắt đầu từ giờ bắt đầu

  var hoursPerDay =
    (new Date('1970-01-01T' + form.endTime + ':00') -
      new Date('1970-01-01T' + form.startTime + ':00')) /
    (1000 * 60 * 60);

  // Duyệt qua từng môn học
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == form.semester) {
      var subject = data[i][2];
      var totalHours = data[i][3];
      var sessionCount = 1; // Khởi tạo bộ đếm số buổi học

      // Số giờ đã học
      var hoursThisWeek = 0;

      // Lặp lại cho đến khi đã đăng ký đủ số giờ cho môn đó
      while (hoursThisWeek < totalHours) {
        // Kiểm tra xem ngày hiện tại là thứ 7 hoặc chủ nhật hay không
        if (startDate.getDay() !== 0 && startDate.getDay() !== 6) {
          // Tạo sự kiện trên Calendar từ giờ bắt đầu đến giờ kết thúc
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

          var eventTitle = subject + '-TL' + ('0' + sessionCount).slice(-2); // Tạo tiêu đề sự kiện với ký tự -TL0x

          calendar.createEvent(eventTitle, startDate, eventEnd);

          // Cập nhật số giờ đã học
          hoursThisWeek += hoursPerDay;
          sessionCount++; // Tăng số buổi học
        }

        // Tăng ngày bắt đầu lên cho ngày tiếp theo
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
