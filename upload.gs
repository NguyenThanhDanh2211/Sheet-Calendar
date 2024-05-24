/**
 * Upload schedule from sheet to Google Calendar
 */
function uploadToCalendar() {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TamplateSchedule');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet "Schedule" không tồn tại!');
    return;
  }

  Logger.log('Sheet "Schedule" đã được tìm thấy');

  var data = sheet.getDataRange().getValues();
  var calendarId = Session.getActiveUser().getEmail();
  var calendar = CalendarApp.getCalendarById(calendarId);

  Logger.log('Bắt đầu xử lý dữ liệu từ hàng 10');

  // Bắt đầu đọc từ hàng thứ 11 để bỏ qua tiêu đề và hàng trống
  for (var i = 9; i < data.length; i += 1) {
    // i += 2 để cách một dòng
    var dateStr = getMergedValue(sheet, i, 0); // Lấy giá trị của ô ngày
    var timeDetail = data[i][4];

    Logger.log(
      'Đang xử lý dòng ' +
        (i + 1) +
        ': dateStr = ' +
        dateStr +
        ', timeDetail = ' +
        timeDetail
    );

    if (
      !dateStr ||
      !timeDetail ||
      typeof dateStr !== 'string' ||
      typeof timeDetail !== 'string'
    ) {
      Logger.log(
        'Bỏ qua dòng ' + (i + 1) + ' do thiếu dữ liệu hoặc không hợp lệ'
      );
      continue; // Bỏ qua nếu không có ngày, giờ học hoặc không phải chuỗi
    }

    var startDate = parseDate(dateStr);
    if (!startDate) {
      Logger.log(
        'Bỏ qua dòng ' + (i + 1) + ' do định dạng ngày không hợp lệ: ' + dateStr
      );
      continue; // Bỏ qua nếu không phân tích được ngày
    }

    var timeParts = timeDetail.split('-');
    if (timeParts.length !== 2) {
      Logger.log(
        'Bỏ qua dòng ' +
          (i + 1) +
          ' do định dạng giờ không hợp lệ: ' +
          timeDetail
      );
      continue; // Bỏ qua nếu không có đúng định dạng giờ (bắt đầu-kết thúc)
    }

    var startTimeParts = timeParts[0].trim().split(':');
    var endTimeParts = timeParts[1].trim().split(':');
    if (startTimeParts.length !== 2 || endTimeParts.length !== 2) {
      Logger.log(
        'Bỏ qua dòng ' +
          (i + 1) +
          ' do định dạng giờ bắt đầu/kết thúc không hợp lệ: ' +
          timeDetail
      );
      continue; // Bỏ qua nếu giờ không đúng định dạng
    }

    // Thêm sự kiện vào Calendar
    var hasEvent = false; // Biến để kiểm tra nếu có sự kiện trong ngày đó
    for (var day = 5; day <= 9; day++) {
      var eventTitle = getMergedValue(sheet, i, day);
      if (eventTitle && typeof eventTitle === 'string') {
        hasEvent = true; // Đã tìm thấy sự kiện trong ngày đó
        var eventDate = new Date(startDate);
        eventDate.setDate(startDate.getDate() + (day - 5));

        var eventStartTime = new Date(eventDate);
        eventStartTime.setHours(startTimeParts[0], startTimeParts[1]);

        var eventEndTime = new Date(eventDate);
        eventEndTime.setHours(endTimeParts[0], endTimeParts[1]);

        calendar.createEvent(eventTitle, eventStartTime, eventEndTime);
        Logger.log('Đã tạo sự kiện: ' + eventTitle + ' vào ngày ' + eventDate);
      }
    }

    // Nếu không có sự kiện trong ngày đó, bỏ qua
    if (!hasEvent) {
      Logger.log('Không có sự kiện trong ngày ' + dateStr + ', bỏ qua...');
    }
  }

  SpreadsheetApp.getUi().alert('Đã tải lịch lên Google Calendar!');
}

function getMergedValue(sheet, row, column) {
  var range = sheet.getRange(row + 1, column + 1); // Ranges in Google Sheets are 1-indexed
  if (range.isPartOfMerge()) {
    var mergedRanges = range.getMergedRanges();
    var firstRange = mergedRanges[0];
    return firstRange.getValue();
  } else {
    return sheet.getRange(row + 1, column + 1).getValue();
  }
}
// Hàm để phân tích chuỗi ngày tháng năm (dd/mm/yyyy)
function parseDate(dateStr) {
  var parts = dateStr.split('/');
  if (parts.length !== 3) return null; // Bỏ qua nếu không có đúng định dạng ngày

  var day = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1; // Tháng bắt đầu từ 0
  var year = parseInt(parts[2], 10);

  return new Date(year, month, day);
}
