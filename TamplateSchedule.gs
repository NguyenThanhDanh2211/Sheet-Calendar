/**
 * Tạo KTB mẫu (->sheets)
 */

function TamplateSchedule() {
  TamplateSchedule_Header();

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var courseSheet = spreadsheet.getSheetByName('course');
  if (!courseSheet) {
    SpreadsheetApp.getUi().alert('Sheet "course" không tồn tại.');
    return;
  }

  // Lấy dữ liệu từ sheet "course_test"
  var data = courseSheet.getDataRange().getValues();
  var courses = data;

  var timetableSheet = spreadsheet.getSheetByName('TamplateSchedule');
  if (!timetableSheet) {
    timetableSheet = spreadsheet.insertSheet('TamplateSchedule');
  }

  // Tạo một đối tượng để lưu trữ lịch học
  var schedule = [
    [
      'Ngày bắt đầu',
      '-',
      'Ngày kết thúc',
      'TUẦN',
      'Giờ học',
      'THỨ HAI',
      'THỨ BA',
      'THỨ TƯ',
      'THỨ NĂM',
      'THỨ SÁU',
    ],
  ];

  var hoursPerDay = 2;
  var dayIndex = 1; // Bắt đầu từ thứ Hai
  var currentHour = 9; // Bắt đầu từ 09:00
  var currentWeek = 1;
  var startDate = new Date('2024-05-20T00:00:00Z'); // Ngày bắt đầu cứng
  var currentDate = new Date(startDate); // Biến để tính toán ngày thay đổi

  // Duyệt qua các khóa học và phân bổ giờ học
  for (var i = 0; i < courses.length; i++) {
    if (courses[i][0] !== 'HK1') {
      // Chỉ xử lý các khóa học thuộc học kỳ 1
      continue;
    }

    var courseName = courses[i][2];
    var totalHours = courses[i][3];
    var hoursScheduled = 0;

    while (hoursScheduled < totalHours) {
      if (dayIndex > 5) {
        // Chuyển sang tuần tiếp theo
        dayIndex = 1;
        currentHour = 9; // Reset giờ bắt đầu
        currentWeek++;
        currentDate.setDate(currentDate.getDate() + 7); // Chuyển sang ngày đầu tiên của tuần tiếp theo
        schedule.push([
          formatDate(currentDate),
          '-',
          formatDate(
            new Date(
              currentDate.getFullYear(),
              currentDate.getMonth(),
              currentDate.getDate() + 4
            )
          ),
          currentWeek,
          '',
          '',
          '',
          '',
          '',
          '',
        ]); // Thêm hàng mới cho tuần tiếp theo
      }

      var currentRow = schedule.length - 1;
      var dayOffset = (currentWeek - 1) * 7 + dayIndex - 1;
      var dayDate = new Date(startDate);
      dayDate.setDate(dayDate.getDate() + dayOffset);

      var startTime = ('0' + currentHour).slice(-2) + ':00';
      var endTime = ('0' + (currentHour + hoursPerDay)).slice(-2) + ':00';

      schedule[currentRow][0] = formatDate(currentDate);
      schedule[currentRow][2] = formatDate(
        new Date(
          currentDate.getFullYear(),
          currentDate.getMonth(),
          currentDate.getDate() + 4
        )
      );
      schedule[currentRow][3] = currentWeek;
      schedule[currentRow][4] = startTime + '-' + endTime;

      schedule[currentRow][dayIndex + 4] = courseName; // +4 để đúng cột từ thứ Hai đến thứ Sáu
      hoursScheduled += hoursPerDay;
      dayIndex++;
    }
  }

  // Ghi dữ liệu vào sheet thời khóa biểu
  timetableSheet
    .getRange(10, 1, schedule.length, schedule[0].length)
    .setValues(schedule);
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function TamplateSchedule_Header() {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TamplateSchedule');
  if (!sheet) {
    sheet =
      SpreadsheetApp.getActiveSpreadsheet().insertSheet('TamplateSchedule');
  }

  sheet.clear();

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
    .getRange(5, 1, 1, 10)
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

  sheet.getRange(7, 3).setValue('Mã lớp:').setFontWeight('bold');

  sheet.getRange(6, 9).setValue('Bắt đầu học từ ngày:').setFontWeight('bold');
  sheet
    .getRange(7, 9)
    .setValue('Học Lý thuyết tại phòng:')
    .setFontWeight('bold');
  sheet
    .getRange(8, 9)
    .setValue('Học Thực hành tại phòng:')
    .setFontWeight('bold');

  // Tạo tiêu đề cột cho thời khóa biểu
  var headers = [
    'TUẦN',
    'Giờ học',
    'THỨ HAI',
    'THỨ BA',
    'THỨ TƯ',
    'THỨ NĂM',
    'THỨ SÁU',
  ];
  sheet
    .getRange(9, 1, 1, 3)
    .merge()
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setValue('NGÀY')
    .setBackground('#cfe2f3');
  sheet
    .getRange(9, 4, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#cfe2f3');
}
