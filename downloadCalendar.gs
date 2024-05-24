/**
 * Download TKB từ calendar về sheet
 */

function downloadCalendar() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Nhập ngày bắt đầu và ngày kết thúc (định dạng YYYY-MM-DD)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    var dates = response.getResponseText().split(' ');
    if (dates.length !== 2) {
      ui.alert('Vui lòng nhập đúng định dạng: YYYY-MM-DD YYYY-MM-DD');
      return;
    }

    var startDate = new Date(dates[0]);
    var endDate = new Date(dates[1]);

    if (isNaN(startDate) || isNaN(endDate)) {
      ui.alert('Ngày nhập không hợp lệ.');
      return;
    }

    var calendarId = Session.getActiveUser().getEmail();
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
        'Giờ học',
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
        var daysEvents = {}; // Lưu các sự kiện theo từng ngày trong tuần
        for (var i = 0; i < weekEvents.length; i++) {
          var event = weekEvents[i];
          var eventDay = event.getStartTime().getDay();
          var startTime = formatTime(event.getStartTime());
          var endTime = formatTime(event.getEndTime());
          var time = startTime + ' - ' + endTime;
          var eventInfo = event.getTitle();

          if (!daysEvents[eventDay]) {
            daysEvents[eventDay] = [];
          }
          daysEvents[eventDay].push({ info: eventInfo, time: time });
        }

        var maxEventsInADay = Math.max(
          ...Object.values(daysEvents).map((events) => events.length)
        );
        for (var j = 0; j < maxEventsInADay; j++) {
          var eventDetails = ['', '', '', '', '', ''];
          var timeDetail = '';
          for (var day in daysEvents) {
            if (daysEvents[day][j]) {
              eventDetails[day - 1] = daysEvents[day][j].info;
              timeDetail = daysEvents[day][j].time; // Lấy thời gian cho cột "Giờ học"
            }
          }
          var eventInfo = [
            weekNumber,
            formatDate(weekStartDate),
            formatDate(weekEndDate),
            timeDetail, // Gán giờ học vào cột "Giờ học"
          ].concat(eventDetails);
          schedule.push(eventInfo);
        }
      }

      // Tăng số tuần và ngày hiện tại lên một tuần
      weekNumber++;
      currentDate = new Date(weekEndDate);
      currentDate.setDate(currentDate.getDate() + 1);
    }

    var sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('downloadSchedule');

    if (!sheet) {
      sheet =
        SpreadsheetApp.getActiveSpreadsheet().insertSheet('downloadSchedule');
    }

    sheet.clear();

    for (var i = 1; i <= 10; i++) {
      sheet
        .getRange(6, i)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBackgroundRGB(123, 222, 123);
    }

    // Ghi dữ liệu mới vào sheet
    if (schedule.length > 0) {
      sheet
        .getRange(6, 1, schedule.length, schedule[0].length)
        .setValues(schedule);

      // Merge cells cho các cột tuần, ngày bắt đầu, ngày kết thúc nếu cần thiết
      var currentWeek = schedule[1][0];
      var startRow = 7;
      for (var row = 8; row < schedule.length + 6; row++) {
        if (schedule[row - 6][0] != currentWeek) {
          if (row - startRow > 1) {
            sheet
              .getRange(startRow, 1, row - startRow, 1)
              .merge()
              .setHorizontalAlignment('center')
              .setVerticalAlignment('middle');
            sheet
              .getRange(startRow, 2, row - startRow, 1)
              .merge()
              .setHorizontalAlignment('center')
              .setVerticalAlignment('middle');
            sheet
              .getRange(startRow, 3, row - startRow, 1)
              .merge()
              .setHorizontalAlignment('center')
              .setVerticalAlignment('middle');
          }
          startRow = row;
          currentWeek = schedule[row - 6][0];
        }
      }
      // Merge the last batch if necessary
      if (schedule.length + 6 - startRow > 1) {
        sheet
          .getRange(startRow, 1, schedule.length + 6 - startRow, 1)
          .merge()
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
        sheet
          .getRange(startRow, 2, schedule.length + 6 - startRow, 1)
          .merge()
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
        sheet
          .getRange(startRow, 3, schedule.length + 6 - startRow, 1)
          .merge()
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
      }
    }

    // Merge cells and set values for the title
    var title1 = sheet
      .getRange(1, 1, 1, 10)
      .merge()
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    var title2 = sheet
      .getRange(2, 1, 1, 10)
      .merge()
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    var title3 = sheet
      .getRange(3, 1, 1, 10)
      .merge()
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    var title4 = sheet
      .getRange(4, 1, 1, 10)
      .merge()
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Set values
    title1.setValue('TRUNG TÂM CÔNG NGHỆ PHẦN MỀM ĐẠI HỌC CẦN THƠ');
    title2.setValue('CANTHO UNIVERSITY SOFTWARE CENTER');
    title3.setValue(
      'Khu III, Đại học Cần Thơ - 01 Lý Tự Trọng, TP. Cần Thơ - Tel: 0292.3731072 & Fax: 0292.3731071 - Email: cusc@ctu.edu.vn'
    );
    title4.setValue('THỜI KHÓA BIỂU');

    // Apply formatting
    var fontItalic = SpreadsheetApp.newTextStyle().setItalic(true).build();

    title1.setFontWeight('bold').setFontSize(11);
    title2.setFontWeight('bold').setFontSize(17);
    title4.setFontWeight('bold').setFontSize(18);
    title3.setTextStyle(fontItalic);
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
  endDate.setDate(endDate.getDate() + 6); // Set ngày kết thúc của tuần (Thứ Bảy)

  return [startDate, endDate];
}
