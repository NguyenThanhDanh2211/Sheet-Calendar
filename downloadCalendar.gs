function downloadCalendar() {
  var calendarId = Session.getActiveUser().getEmail();

  var startDate = new Date('2024-05-20');
  var endDate = new Date('2024-06-30');

  var events = CalendarApp.getCalendarById(calendarId).getEvents(
    startDate,
    endDate
  );

  // Tạo một mảng để lưu trữ thời khóa biểu
  var schedule = [
    [
      'Tuần',
      'Từ ngày',
      'Đến ngày',
      'Giờ bắt đầu',
      'Giờ kết thúc',
      'Thứ 2',
      'Thứ 3',
      'Thứ 4',
      'Thứ 5',
      'Thứ 6',
      'Thứ 7',
    ],
  ];
  // Duyệt qua các tuần trong khoảng thời gian từ startDate đến endDate
  var currentDate = new Date(startDate);
  var weekNumber = 1;

  while (currentDate <= endDate) {
    var weekDates = getWeekDatesFromDate(currentDate);
    var weekStartDate = weekDates[0];
    var weekEndDate = weekDates[1];

    // Duyệt qua các sự kiện trong tuần hiện tại
    var weekEvents = events.filter((event) => {
      var eventStartTime = event.getStartTime();
      return eventStartTime >= weekStartDate && eventStartTime <= weekEndDate;
    });

    if (weekEvents.length > 0) {
      // Chỉ thêm tuần nếu có sự kiện trong tuần đó
      var eventDetails = ['', '', '', '', '', '']; // Tạo một mảng rỗng cho các ngày trong tuần từ Thứ 2 đến Thứ 7
      for (var i = 0; i < weekEvents.length; i++) {
        var event = weekEvents[i];
        var eventDay = event.getStartTime().getDay();
        var startTime = formatTime(event.getStartTime());
        var endTime = formatTime(event.getEndTime());
        var eventInfo = event.getTitle();
        if (eventDay >= 1 && eventDay <= 6) {
          eventDetails[eventDay - 1] += eventInfo;
        }
      }

      // Tạo một mảng con để lưu thông tin về tuần
      var eventInfo = [
        weekNumber,
        formatDate(weekStartDate),
        formatDate(weekEndDate),
        startTime,
        endTime,
      ].concat(eventDetails);

      // Thêm mảng con vào mảng chính
      schedule.push(eventInfo);
    }

    // Tăng số tuần và ngày hiện tại lên một tuần
    weekNumber++;
    currentDate = new Date(weekEndDate);
    currentDate.setDate(currentDate.getDate() + 1);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Calendar');

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

// Hàm để lấy ngày bắt đầu và kết thúc của tuần chứa ngày đã cho
function getWeekDatesFromDate(date) {
  var startDate = new Date(date);
  startDate.setDate(startDate.getDate() - startDate.getDay() + 1); // Set ngày bắt đầu của tuần (Thứ Hai)

  var endDate = new Date(startDate);
  endDate.setDate(endDate.getDate() + 5); // Set ngày kết thúc của tuần (Thứ Bảy)

  return [startDate, endDate];
}
