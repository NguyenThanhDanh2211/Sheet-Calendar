function haiTOsau(semester, startDate, startHour, ngaynghi) {
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
  var restDay = {};

  startDate = new Date(startDate + 'T' + startHour);
  var startTime = parseFloat(startHour.split(':')[0]);
  var startMinutes = parseFloat(startHour.split(':')[1]);
  var hoursPerDay = 2;
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
        assignedExamDates[formatDate(currentDate)] ||
        restDay[formatDate(currentDate)]
      ) {
        if (
          isHoliday(currentDate, holidays) &&
          !(currentDate.getDay() === 0 || currentDate.getDay() === 6)
        ) {
          // Thêm ngày nghỉ vào mảng schedule nếu chưa thêm
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

      if (currentDate.getDay() >= 1 && currentDate.getDay() <= 5) {
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

      // Kiểm tra nếu ngày trước ngày thi không phải là ngày nghỉ thì thêm "Self study"
      var dayBeforeExam = new Date(examDate);
      dayBeforeExam.setDate(dayBeforeExam.getDate() - 1);

      if (
        !isHoliday(dayBeforeExam, holidays) &&
        dayBeforeExam.getDay() !== 0 &&
        dayBeforeExam.getDay() !== 6
      ) {
        schedule.push(
          createScheduleEntry(
            dayBeforeExam,
            startTime,
            startMinutes,
            hoursPerDay,
            'Self study'
          )
        );
        restDay[formatDate(dayBeforeExam)] = true; // Đánh dấu rằng đã gán ngày self study
      }
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
        !restDay[formatDate(selfStudyDate)] &&
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

  var result = {
    schedule: schedule,
    examDay: examDay,
  };

  return result;
}
