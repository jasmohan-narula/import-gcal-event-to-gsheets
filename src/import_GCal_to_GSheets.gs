function importGoogleCalendar() { 
  var sheet = SpreadsheetApp.getActiveSheet();
  var calendarId = sheet.getRange('B1').getValue().toString(); 
  var calendar = CalendarApp.getCalendarById(calendarId);
 
  // Set filters
  var startDate = sheet.getRange('B2').getValue();
  var endDate = sheet.getRange('B3').getValue();
  var searchText = sheet.getRange('B4').getValue();
 
  // Print header
  var header = [["Title", "Description", "Location", "Start", "End", "Duration", "Start Time", "End Time","Text - Intermediate", "Call (Edit)", "Call Via (Edit)", "GuestList (Edit)", "Final Timesheet Text"]];
	
  var range = sheet.getRange("A6:M6");
  range.setValues(header);
  range.setFontWeight("bold")
 
  // Get events based on filters from Google Calendar
  var events = (searchText == '') ? calendar.getEvents(startDate, endDate) : calendar.getEvents(startDate, endDate, {search: searchText});
 
  // Display events 
  for (var i=0; i<events.length; i++) {
    // Skip configuration and header rows
    var row = i+7;
    
    var details = [[events[i].getTitle(), events[i].getDescription(), events[i].getLocation(), events[i].getStartTime(), events[i].getEndTime(), 0, events[i].getStartTime(), events[i].getEndTime()]];
    
    range = sheet.getRange(row,1,1,8);
    range.setValues(details);
 

    // Format the Start and End columns
    var cell = sheet.getRange(row, 4);
    cell.setNumberFormat('mm/dd/yyyy hh:mm');

    cell = sheet.getRange(row, 5);
    cell.setNumberFormat('mm/dd/yyyy hh:mm');

    // Time Set
    startDate = sheet.getRange(row, 7).setNumberFormat('hh:mm AM/PM').getDisplayValue();
    //console.log(startDate);
    endDate = sheet.getRange(row, 8).setNumberFormat('hh:mm AM/PM').getDisplayValue();
    //console.log(endDate);

    // Calculate the Duration column
    cell = sheet.getRange(row, 6);
    cell.setFormula('=(HOUR(E' + row + ')+(MINUTE(E' +row+ ')/60))-(HOUR(D' +row+ ')+(MINUTE(D' +row+ ')/60))');
    cell.setNumberFormat('0.00');


    var location =  events[i].getLocation();
    if(location != ""){
      sheet.getRange(row, 10).insertCheckboxes("Y");
    } else {
      sheet.getRange(row, 10).insertCheckboxes("N");
    }


    //Clear Call Via
    sheet.getRange(row, 11).setValue('');


    // Timesheet Final Text
    var timeSheetFinalText = '';
    timeSheetFinalText = "{(" + startDate + "-" + endDate + ")" + " \"" + events[i].getTitle() + "\"";
    timeSheetFinalText = timeSheetFinalText + " Call ";
    timeSheetFinalText = timeSheetFinalText + "[via ]";
    timeSheetFinalText = timeSheetFinalText + "}";
    console.log(timeSheetFinalText);
    sheet.getRange(row, 9).setValue(timeSheetFinalText);



    

    var finalCellForCall = sheet.getRange(row, 13);
    var cellNameForText = "I" + row ;
    var cellNameForCall = "K" + row;
    var formulaSubstituteValue = '"' + "[via " + '"&' + cellNameForCall + '&"' + "]" + '"';

    finalCellForCall.setFormula('=SUBSTITUTE(' + cellNameForText + ', "[via ]", ' + formulaSubstituteValue + ')');    
  }
}