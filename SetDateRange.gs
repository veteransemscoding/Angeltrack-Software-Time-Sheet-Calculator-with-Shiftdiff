//sets the payroll week for the previous week in a range from sun to sun

function setPayrollWeek(sheet) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Sheet1");

  var startDate = getStartOfLastWeek();
  var endDate = getPreviousSunday();

  adjustClockTimes(sheet, startDate, endDate);
}
//creates the start of the payroll week by setting it to sunday
function getStartOfLastWeek() {
  var currentDate = new Date();
  var startOfLastWeek = new Date(
    currentDate.getFullYear(),
    currentDate.getMonth(),
    currentDate.getDate() - 7
  );

  var startOfLastWeekSunday = startOfLastWeek.getDate() - startOfLastWeek.getDay();
  var payrollWeekStart = new Date(startOfLastWeek.getFullYear(), startOfLastWeek.getMonth(), startOfLastWeekSunday);
  payrollWeekStart.setHours(8, 0, 0); // Set time to 8:00 AM

  return payrollWeekStart;
}
//creates the end of the payroll week by setting it to the most recent sunday
function getPreviousSunday() {
  var currentDate = new Date();
  var previousSunday = currentDate.getDate() - currentDate.getDay();
  var payrollWeekEnd = new Date(currentDate.getFullYear(), currentDate.getMonth(), previousSunday);
  payrollWeekEnd.setHours(8, 0, 0); // Set time to 8:00 AM

  return payrollWeekEnd;
}
//searches for times outside of the range and sets it to the the cutoff
function adjustClockTimes(sheet, startDate, endDate) {
  var clockedInColumn = getColumnIndexByHeader(sheet, "ClockedIn");
  var clockedOutColumn = getColumnIndexByHeader(sheet, "ClockedOut");

  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  for (var i = 1; i < data.length; i++) {
    var clockedIn = new Date(data[i][clockedInColumn]);
    var clockedOut = new Date(data[i][clockedOutColumn]);

    if (clockedIn < startDate) {
      data[i][clockedInColumn] = startDate;
    }

    if (clockedOut > endDate) {
      data[i][clockedOutColumn] = endDate;
    }
  }

  dataRange.setValues(data);
}
