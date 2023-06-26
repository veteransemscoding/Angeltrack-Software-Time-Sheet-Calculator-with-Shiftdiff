function isSunday(date) {
  return date.getDay() === 0; // 0 represents Sunday (Monday is 1, Tuesday is 2, and so on)
}


function calculateSundayDaySD(sheet) {
  var basedOutOfStationColumn = getColumnIndexByHeader(sheet, 'BasedOutOfStation');
  var sundayDaySDColumn = getColumnIndexByHeader(sheet, 'Sunday Day SD');
  var clockedInColumn = getColumnIndexByHeader(sheet, 'ClockedIn');
  var clockedOutColumn = getColumnIndexByHeader(sheet, 'ClockedOut');

  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  for (var i = 1; i < data.length; i++) {
    var basedOutOfStation = data[i][basedOutOfStationColumn];

    if (basedOutOfStation.includes('EMT')) {
      var clockedIn = new Date(data[i][clockedInColumn]);
      var clockedOut = new Date(data[i][clockedOutColumn]);

      if (clockedIn !== '' && clockedOut !== '' && isSunday(clockedIn)) {
        var sundayDaySD = calculateSundayDaySDHours(clockedIn, clockedOut);
        if (sundayDaySD >= 0) {
          sheet.getRange(i + 1, sundayDaySDColumn + 1).setValue(sundayDaySD.toFixed(1));
        } else {
          sheet.getRange(i + 1, sundayDaySDColumn + 1).setValue(0);
        }
      }
    }
  }
}

function calculateSundayDaySDHours(clockedIn, clockedOut) {
  var sundayDayStart = new Date(clockedIn);
  sundayDayStart.setHours(8, 0, 0); // Set Sunday day start time to 8:00 AM

  var sundayDayEnd = new Date(clockedIn);
  sundayDayEnd.setHours(20, 0, 0); // Set Sunday day end time to 8:00 PM

  var sundayDayHours = 0;

  if (clockedIn < sundayDayStart) {
    clockedIn = sundayDayStart;
  }
  if (clockedOut > sundayDayEnd) {
    clockedOut = sundayDayEnd;
  }

  if (clockedOut > clockedIn) {
    sundayDayHours = (clockedOut - clockedIn) / (1000 * 60 * 60);
  }

  return sundayDayHours;
}

