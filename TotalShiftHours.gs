//calculates the total hours
function calculateTotalHours(sheet) {
  var clockedInColumn = getColumnIndexByHeader(sheet, "ClockedIn");
  var clockedOutColumn = getColumnIndexByHeader(sheet, "ClockedOut");
  var totalHoursColumn = getColumnIndexByHeader(sheet, "Total Hours");

  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  // Iterate through each row of data, starting from the second row
  for (var i = 1; i < data.length; i++) {
    var clockedIn = new Date(data[i][clockedInColumn]);
    var clockedOut = new Date(data[i][clockedOutColumn]);

    if (clockedIn !== "" && clockedOut !== "") {
      var totalHours = calculateDecimalTime(clockedIn, clockedOut);
      sheet.getRange(i + 1, totalHoursColumn + 1).setValue(totalHours);
    }
  }
}
//convert the time to decimal form
function calculateDecimalTime(clockedIn, clockedOut) {
  var diffMillis = clockedOut - clockedIn;
  var diffHours = diffMillis / (1000 * 60 * 60); // Convert milliseconds to hours

  return diffHours.toFixed(1);
}
