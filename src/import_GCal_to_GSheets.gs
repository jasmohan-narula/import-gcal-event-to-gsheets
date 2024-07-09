function importGoogleCalendar() { 
  var sheet = SpreadsheetApp.getActiveSheet();
  var calendarId = sheet.getRange('B1').getValue().toString(); 
  var calendar = CalendarApp.getCalendarById(calendarId);
  
  if (!calendar) {
    Logger.log('Calendar not found: ' + calendarId);
    return;
  }

  // Set filters
  var startDate = sheet.getRange('B2').getValue();
  var endDate = sheet.getRange('B3').getValue();
  var searchText = sheet.getRange('B4').getValue();
  
  // Set header in Google Sheet
  var header = [["Title", "Description", "Location", "Start DateTime", "End DateTime", "Duration", "Start Time", "End Time", "Text - Intermediate", "Call (Edit)", "Call Via (Edit)", "GuestList (Edit)", "Final Timesheet Text"]];
  var range = sheet.getRange("A6:M6");
  range.setValues(header);
  range.setFontWeight("bold");

  // Get events based on filters from Google Calendar
  var events = (searchText == '') ? calendar.getEvents(startDate, endDate) : calendar.getEvents(startDate, endDate, {search: searchText});
  
  // Prepare data for bulk update
  var data = [];
  for (var i = 0; i < events.length; i++) {
    // Skip configuration and header rows
    var row = i + 7;

    var event = events[i];

    // Calculate Time
    var startTime = event.getStartTime();
    var endTime = event.getEndTime();
    var duration = (endTime - startTime) / (1000 * 60 * 60); // Duration in hours
    var startTimeFormatted = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'hh:mm a');
    var endTimeFormatted = Utilities.formatDate(endTime, Session.getScriptTimeZone(), 'hh:mm a');

    var callEdit = false;

    var timeSheetFinalText = `{(${startTimeFormatted}-${endTimeFormatted}) "${event.getTitle()}"}`;
    // Print "Final Timesheet Text"
    Logger.log(timeSheetFinalText);

    data.push([event.getTitle(), event.getDescription(), event.getLocation(), startTime, endTime, duration, startTimeFormatted, endTimeFormatted, timeSheetFinalText, callEdit, '', '', '']);
  }

  // Update sheet in bulk
  if (data.length > 0) {
    var dataRange = sheet.getRange(7, 1, data.length, 13);
    dataRange.setValues(data);

    // Apply formatting
    var startRange = sheet.getRange(7, 4, data.length, 1);
    var endRange = sheet.getRange(7, 5, data.length, 1);
    startRange.setNumberFormat('MM/dd/yyyy hh:mm AM/PM');
    endRange.setNumberFormat('MM/dd/yyyy hh:mm AM/PM');

    var durationRange = sheet.getRange(7, 6, data.length, 1);
    durationRange.setNumberFormat('0.00');

    var startTimeRange = sheet.getRange(7, 7, data.length, 1);
    var endTimeRange = sheet.getRange(7, 8, data.length, 1);
    startTimeRange.setNumberFormat('hh:mm AM/PM');
    endTimeRange.setNumberFormat('hh:mm AM/PM');

    // Insert checkboxes for "Call (Edit)"
    var callEditRange = sheet.getRange(7, 10, data.length, 1);
    callEditRange.insertCheckboxes();

    // Set final timesheet text formula dynamically
    for (var j = 0; j < data.length; j++) {
      var row = j + 7;
      var finalCellForCall = sheet.getRange(row, 13);
      var cellNameForText = "I" + row;
      var cellNameForCallCondition = "J" + row;
      var cellNameForCallViaCondition = "K" + row;

      var finalFormula = '=CONCATENATE(' + cellNameForText + ', IF(' + cellNameForCallCondition + '=TRUE, " Call", ""), IF(' + cellNameForCallViaCondition + '="", "", CONCATENATE(" [via ",' + cellNameForCallViaCondition + ',"]")))';
      Logger.log(finalFormula);

      finalCellForCall.setFormula(finalFormula);
    }
  }
}
