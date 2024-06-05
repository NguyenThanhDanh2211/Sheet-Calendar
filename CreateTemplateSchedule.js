/**
 * Main: CreateTemplateSchedule
 */

function CreateTemplateSchedule(formData) {
  var name = formData.name;
  var semester = formData.semester;
  var startDate = formData.startDate;
  var startHour = formData.startHour;
  var selectedDay = formData.ngayhoc;
  var theoryRoom = formData.theoryRoom;
  var practiceRoom = formData.practiceRoom;
  var restDay = formData.restDay;
  var startHourMore = formData.startHourMore;
  var weekMore = formData.startHourMoreW;

  var startweekMore, endweekMore;
  if (weekMore) {
    startweekMore = parseFloat(weekMore.split('-')[0]);
    endweekMore = parseFloat(weekMore.split('-')[1]);
  }

  if (selectedDay == 'days357') {
    var result = banambay(semester, startDate, startHour, restDay);
  } else if (selectedDay == 'days246') {
    var result = haitusau(semester, startDate, startHour, restDay);
  } else {
    result = haiTOsau(semester, startDate, startHour, restDay);
  }

  if (!result || result.length === 0) {
    // Xử lý khi hàm TemplateSchedule trả về giá trị không hợp lệ
    SpreadsheetApp.getUi().alert(
      'Không thể tạo lịch trình mẫu. Dữ liệu thời khóa biểu không khả dụng.'
    );
    return;
  }

  var schedule = result.schedule;
  var examDay = result.examDay;

  // Tạo một mảng 2 chiều để lưu trữ dữ liệu chuyển đổi
  var transposedValues = [];
  var currentRow = [];
  var weekCount = 1; // Biến đếm số tuần

  for (var i = 0; i < schedule.length; i += 5) {
    // Reset currentRow for every new line in the transposed sheet
    currentRow = [];

    // Extract start and end time
    var startTime, endTime;
    if (
      startHourMore &&
      weekMore &&
      weekCount >= startweekMore &&
      weekCount <= endweekMore
    ) {
      startTime = startHourMore;
      if (selectedDay == 'days246' || selectedDay == 'days357') {
        endTime = addHours(startHourMore, 3);
      } else {
        endTime = addHours(startHourMore, 2);
      }
    } else {
      startTime = schedule[i][2].split(' - ')[0];
      endTime = schedule[i][2].split(' - ')[1];
    }

    // Add the start date of the first row
    if (i < schedule.length) {
      currentRow.push(schedule[i][0]); // Start date of the first row
    } else {
      currentRow.push('');
    }

    currentRow.push('-');

    // Add the end date of the third row
    if (i + 4 < schedule.length) {
      currentRow.push(schedule[i + 4][1]); // End date of the third row
    } else {
      currentRow.push(schedule[Math.min(i + 4, schedule.length - 1)][1]); // End date of the last available row
    }

    // Add the week number
    currentRow.push(weekCount);
    // Add the start and end time
    currentRow.push(startTime + ' - ' + endTime);

    // Add the course names
    for (var j = 0; j < 5; j++) {
      if (i + j < schedule.length) {
        currentRow.push(schedule[i + j][3]);
      } else {
        currentRow.push('');
      }
    }
    // Push the currentRow to transposedValues
    transposedValues.push(currentRow);
    weekCount++;
  }

  // Tạo một sheet mới để lưu trữ dữ liệu chuyển đổi
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var newSheet = spreadsheet.getSheetByName(name);
  if (!newSheet) {
    newSheet = spreadsheet.insertSheet(name);
  } else {
    newSheet.clear();
  }

  TemplateSchedule_Header(
    newSheet,
    name,
    startDate,
    selectedDay,
    theoryRoom,
    practiceRoom
  );

  // Ghi dữ liệu chuyển đổi vào sheet mới
  newSheet
    .getRange(10, 1, transposedValues.length, transposedValues[0].length)
    .setValues(transposedValues);

  var examHeaders = ['Course Name', 'Exam Date'];
  newSheet
    .getRange(transposedValues.length + 15, 5, 1, examHeaders.length)
    .setValues([examHeaders]);
  newSheet
    .getRange(
      transposedValues.length + 16,
      5,
      examDay.length,
      examDay[0].length
    )
    .setValues(examDay);

  formatSheet(newSheet);
}

function addHours(time, hours) {
  var parts = time.split(':');
  var hoursPart = parseInt(parts[0], 10);
  var minutesPart = parseInt(parts[1], 10);
  hoursPart += hours;
  return (
    (hoursPart < 10 ? '0' : '') +
    hoursPart +
    ':' +
    (minutesPart < 10 ? '0' : '') +
    minutesPart
  );
}

function TemplateSchedule_Header(
  sheet,
  name,
  startDate,
  selectedDay,
  theoryRoom,
  practiceRoom
) {
  if (!sheet) {
    Logger.log('Bảng không tồn tại.');
    return;
  }

  // sheet.clear();

  sheet.getRange(7, 3).setValue('Mã lớp:').setFontWeight('bold');
  sheet
    .getRange(7, 4, 1, 2)
    .merge()
    .setValue(name)
    .setFontWeight('bold')
    .setFontColor('red');
  sheet.getRange(6, 9).setValue('Bắt đầu học từ ngày:').setFontWeight('bold');
  sheet
    .getRange(6, 10)
    .setValue(formatDate(new Date(startDate)))
    .setFontWeight('bold')
    .setFontColor('red');
  sheet
    .getRange(7, 9)
    .setValue('Học Lý thuyết tại phòng:')
    .setFontWeight('bold');
  sheet
    .getRange(7, 10)
    .setValue(theoryRoom)
    .setFontWeight('bold')
    .setFontColor('red');
  sheet
    .getRange(8, 9)
    .setValue('Học Thực hành tại phòng:')
    .setFontWeight('bold');
  sheet
    .getRange(8, 10)
    .setValue(practiceRoom)
    .setFontWeight('bold')
    .setFontColor('red');

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

  // Tạo tiêu đề cột cho thời khóa biểu
  var headers = [];

  if (selectedDay === 'days357') {
    headers = [
      'TUẦN',
      'Giờ học',
      'THỨ BA',
      'THỨ TƯ',
      'THỨ NĂM',
      'THỨ SÁU',
      'THỨ BẢY',
    ];
  } else {
    headers = [
      'TUẦN',
      'Giờ học',
      'THỨ HAI',
      'THỨ BA',
      'THỨ TƯ',
      'THỨ NĂM',
      'THỨ SÁU',
    ];
  }
  sheet
    .getRange(9, 4, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#6096B4');
  sheet
    .getRange(9, 1, 1, 3)
    .merge()
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setValue('NGÀY')
    .setBackground('#6096B4');

  sheet
    .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setFontFamily('Times New Roman');
  sheet
    .getRange(6, 1, sheet.getMaxRows() - 5, sheet.getMaxColumns())
    .setFontSize(12);
}

function formatSheet(sheet) {
  if (!sheet) {
    Logger.log('Bảng không tồn tại.');
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(1, 1, lastRow, lastColumn);
  var values = range.getValues();
  var backgrounds = range.getBackgrounds();
  var fonts = range.getFontWeights();
  var fontStyles = range.getFontStyles();
  var fontColors = range.getFontColors();

  for (var i = 9; i < lastRow; i++) {
    for (var j = 0; j < lastColumn; j++) {
      var cellValue = values[i][j].toString().toLowerCase();
      if (cellValue.indexOf('thi') !== -1) {
        backgrounds[i][j] = '#93BFCF'; // Màu xanh
      }
      if (cellValue.indexOf('nghỉ') !== -1) {
        backgrounds[i][j] = '#ffff00'; // Màu vàng
        fonts[i][j] = 'bold'; // Chữ đậm
        fontColors[i][j] = '#ff0000'; // Màu đỏ
        range.getCell(i + 1, j + 1).setHorizontalAlignment('center'); // Căn giữa chữ
      }
      if (cellValue.indexOf('self study') !== -1) {
        fontStyles[i][j] = 'italic'; // In nghiêng chữ
        range.getCell(i + 1, j + 1).setHorizontalAlignment('center'); // Căn giữa chữ
      }
    }
  }
  range.setBackgrounds(backgrounds);
  range.setFontWeights(fonts);
  range.setFontColors(fontColors);
  range.setFontStyles(fontStyles);
}
