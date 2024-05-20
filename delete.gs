function deleteEventsFromDateRange() {
  var calendar = CalendarApp.getDefaultCalendar();
  var startDate = new Date('2024-05-20T00:00:00Z'); // Ngày bắt đầu
  var endDate = new Date('2024-12-31T23:59:59Z'); // Ngày kết thúc

  var events = calendar.getEvents(startDate, endDate);

  for (var i = 0; i < events.length; i++) {
    events[i].deleteEvent();
  }

  Logger.log(
    'All events from May 20, 2024 to the end of 2024 have been deleted.'
  );
}
