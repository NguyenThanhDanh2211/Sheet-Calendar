function haitusau(semester, startDate, startHour, ngaynghi) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var courseSheet = spreadsheet.getSheetByName('course');
  if (!courseSheet) {
    Logger.log('Không tìm thấy bảng');
    return;
  }

  var courses = courseSheet.getDataRange().getValues();
  var schedule = [];
  var courseEndDates = {};
  var examDay = [];
  var assignedExamDates = {}; // Sử dụng để đánh dấu rằng ngày thi đã được gán cho môn học này

  startDate = new Date(startDate + 'T' + startHour);
  var startTime = parseFloat(startHour.split(':')[0]);
  var startMinutes = parseFloat(startHour.split(':')[1]);
  var hoursPerDay = 3;
  var lastCourseEndDate = null;
  var lastCourseExamDate = null;

  var holidays = ngaynghi ? parseHolidays(ngaynghi) : [];

  for (var i = 1; i < courses.length; i++) {
    if (courses[i][0] !== semester) {
      continue;
    }

    var courseName = courses[i][2];
    var courseTotalHour = courses[i][3];
    if (courseTotalHour == 0) {
      continue;
    }
    var type = courses[i][4].split('/')[0];

    var currentDate = new Date(startDate);
    var lessonCount = 1;

    while (courseTotalHour > 0) {
      while (
        currentDate.getDay() === 0 ||
        currentDate.getDay() === 6 ||
        isHoliday(currentDate, holidays) ||
        assignedExamDates[formatDate(currentDate)]
      ) {
        if (
          isHoliday(currentDate, holidays) &&
          !(currentDate.getDay() === 0 || currentDate.getDay() === 6)
        ) {
          schedule.push(
            createScheduleEntry(
              currentDate,
              startTime,
              startMinutes,
              hoursPerDay,
              'NGHỈ'
            )
          );
        }
        currentDate.setDate(currentDate.getDate() + 1);
      }

      var endDate = new Date(currentDate);
      endDate.setHours(startTime + hoursPerDay, startMinutes);

      if (
        currentDate.getDay() === 1 ||
        currentDate.getDay() === 3 ||
        currentDate.getDay() === 5
      ) {
        schedule.push(
          createScheduleEntry(
            currentDate,
            startTime,
            startMinutes,
            hoursPerDay,
            courseName +
              ' - ' +
              type +
              (lessonCount < 10 ? '0' + lessonCount : lessonCount)
          )
        );
        lessonCount++;
        courseTotalHour -= hoursPerDay;
      } else {
        schedule.push(
          createScheduleEntry(
            currentDate,
            startTime,
            startMinutes,
            hoursPerDay,
            'self study'
          )
        );
      }

      currentDate.setDate(currentDate.getDate() + 1);
      startDate = new Date(currentDate);
    }

    // Store the end date of the course outside the loop
    courseEndDates[courseName] = formatDate(endDate);

    if (courses[i][4].includes('BC')) {
      var reportDate = parseDate(courseEndDates[courseName]);
      reportDate.setDate(reportDate.getDate() + 7);
      // Dời ngày báo cáo nếu trùng ngày nghỉ
      while (
        isHoliday(reportDate, holidays) ||
        reportDate.getDay() === 0 ||
        reportDate.getDay() === 6
      ) {
        reportDate.setDate(reportDate.getDate() + 1);
      }
      examDay.push(['Báo cáo ' + courseName, formatDate(reportDate)]);
    } else {
      var examDate = parseDate(courseEndDates[courseName]);
      examDate.setDate(examDate.getDate() + 7);
      // Dời ngày thi nếu trùng ngày nghỉ hoặc thứ 7 và Chủ nhật
      while (
        isHoliday(examDate, holidays) ||
        examDate.getDay() === 0 ||
        examDate.getDay() === 6
      ) {
        examDate.setDate(examDate.getDate() + 1);
      }
      examDay.push([courseName, formatDate(examDate)]);
      assignedExamDates[formatDate(examDate)] = true; // Đánh dấu rằng đã gán ngày thi cho môn học này

      // Thêm ngày thi vào mảng
      schedule.push(
        createScheduleEntry(
          examDate,
          startTime,
          startMinutes,
          hoursPerDay,
          'Thi - ' + courseName
        )
      );
    }

    // Lưu trữ ngày kết thúc và ngày thi của môn học cuối cùng
    lastCourseEndDate = endDate;
    lastCourseExamDate = examDate;
  }

  if (lastCourseEndDate && lastCourseExamDate) {
    var selfStudyDate = new Date(
      lastCourseEndDate.setDate(lastCourseEndDate.getDate() + 1)
    );
    while (selfStudyDate < lastCourseExamDate) {
      if (
        selfStudyDate.getDay() !== 0 &&
        selfStudyDate.getDay() !== 6 &&
        !isHoliday(selfStudyDate, holidays)
      ) {
        schedule.push(
          createScheduleEntry(
            selfStudyDate,
            startTime,
            startMinutes,
            hoursPerDay,
            'self study'
          )
        );
      }
      selfStudyDate.setDate(selfStudyDate.getDate() + 1);
    }
  }

  // Sort the schedule by the second column (End date)
  schedule.sort(
    (a, b) => new Date(parseDate(a[1])) - new Date(parseDate(b[1]))
  );

  // var scheduleSheet = spreadsheet.getSheetByName('schedule');
  // if (!scheduleSheet) {
  //   scheduleSheet = spreadsheet.insertSheet('schedule');
  // }

  // var headers = ['Start', 'End', 'Time', 'Course Detail'];
  // scheduleSheet.clear(); // Clear existing data
  // scheduleSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  // scheduleSheet.getRange(2, 1, schedule.length, schedule[0].length).setValues(schedule);

  // // Create a new sheet for exam dates
  // var examSheet = spreadsheet.getSheetByName('exam_schedule');
  // if (!examSheet) {
  //   examSheet = spreadsheet.insertSheet('exam_schedule');
  // }

  // var examHeaders = ['Course Name', 'Exam Date'];
  // examSheet.clear(); // Clear existing data
  // examSheet.getRange(1, 1, 1, examHeaders.length).setValues([examHeaders]);
  // examSheet.getRange(2, 1, examDay.length, examDay[0].length).setValues(examDay);

  var result = {
    schedule: schedule,
    examDay: examDay,
  };

  return result;
}

function parseHolidays(ngaynghi) {
  var holidays = [];
  var holidayRanges = ngaynghi.split(',');

  for (var i = 0; i < holidayRanges.length; i++) {
    var range = holidayRanges[i].trim();
    if (range.includes('-')) {
      var dates = range.split('-');
      var start = parseDate(dates[0].trim());
      var end = parseDate(dates[1].trim());
      while (start <= end) {
        holidays.push(new Date(start));
        start.setDate(start.getDate() + 1);
      }
    } else {
      holidays.push(parseDate(range));
    }
  }

  return holidays;
}

function isHoliday(date, holidays) {
  for (var i = 0; i < holidays.length; i++) {
    if (formatDate(date) === formatDate(holidays[i])) {
      return true;
    }
  }
  return false;
}

function startTimeFormatted(hours, minutes) {
  var formattedHours = hours < 10 ? '0' + hours : hours;
  var formattedMinutes = minutes < 10 ? '0' + minutes : minutes;
  return formattedHours + ':' + formattedMinutes;
}

function formatDate(date) {
  var dd = String(date.getDate()).padStart(2, '0');
  var mm = String(date.getMonth() + 1).padStart(2, '0'); // January is 0!
  var yyyy = date.getFullYear();
  return dd + '/' + mm + '/' + yyyy;
}

function parseDate(dateStr) {
  var parts = dateStr.split('/');
  var day = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1; // Month is zero-based in JavaScript
  var year = parseInt(parts[2], 10);
  return new Date(year, month, day);
}

function getDayOfWeekNumber(date) {
  var days = {
    0: 'CN',
    1: 'T2',
    2: 'T3',
    3: 'T4',
    4: 'T5',
    5: 'T6',
    6: 'T7',
  };
  return days[date.getDay()];
}
