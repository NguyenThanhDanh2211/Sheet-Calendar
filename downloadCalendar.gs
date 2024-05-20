function downloadCalendar() {
  var calendarId = Session.getActiveUser().getEmail();

  var startDate = new Date('2024-05-20');
  var endDate = new Date('2024-06-30');

  var events = CalendarApp.getCalendarById(calendarId).getEvents(
    startDate,
    endDate
  );

  // Tạo một mảng để lưu trữ thời khóa biểu
  var schedule = [];

  for (var i = 0; i < events.length; i++) {
    var event = events[i];

    var eventStartTime = event.getStartTime();
    // (0: Chủ Nhật, 1: Thứ Hai,..., 6: Thứ Bảy)
    var dayOfWeek = eventStartTime.getDay();
    var daysOfWeek = [
      'Chủ Nhật',
      'Thứ Hai',
      'Thứ Ba',
      'Thứ Tư',
      'Thứ Năm',
      'Thứ Sáu',
      'Thứ Bảy',
    ];

    // Xem ngày của sự kiện có từ t2 đến t6 không
    if (dayOfWeek >= 1 && dayOfWeek <= 5) {
      // Thêm thông tin của sự kiện vào mảng thời khóa biểu
      var startDate = formatDate(eventStartTime);
      var endDate = formatDate(event.getEndTime());
      var startTime = formatTime(eventStartTime);
      var endTime = formatTime(event.getEndTime());
      var dayName = daysOfWeek[dayOfWeek];

      schedule.push([
        dayName,
        startDate,
        endDate,
        startTime,
        endTime,
        event.getTitle(),
      ]);
    }
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule');

  sheet.clear();

  // Ghi dữ liệu mới vào sheet
  if (schedule.length > 0) {
    sheet
      .getRange(3, 1, schedule.length, schedule[0].length)
      .setValues(schedule);
  }
}

// Hàm để định dạng ngày
function formatDate(date) {
  var dd = String(date.getDate()).padStart(2, '0');
  var mm = String(date.getMonth() + 1).padStart(2, '0'); // Tháng bắt đầu từ 0
  var yyyy = date.getFullYear();
  return dd + '/' + mm + '/' + yyyy;
}

// Hàm để định dạng giờ
function formatTime(date) {
  var hh = String(date.getHours()).padStart(2, '0');
  var mm = String(date.getMinutes()).padStart(2, '0');
  return hh + ':' + mm;
}
