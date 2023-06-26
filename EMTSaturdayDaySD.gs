function isSaturday(date) {
  return date.getDay() === 6; // 6 represents Saturday (Sunday is 0, Monday is 1, and so on)
}

function calculateSaturdayDaySD(sheet) {
  var basedOutOfStationColumn = getColumnIndexByHeader(sheet, 'BasedOutOfStation');
  var saturdayDaySDColumn = getColumnIndexByHeader(sheet, 'Saturday Day SD');
  var clockedInColumn = getColumnIndexByHeader(sheet, 'ClockedIn');
  var clockedOutColumn = getColumnIndexByHeader(sheet, 'ClockedOut');

  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  for (var i = 1; i < data.length; i++) {
    var basedOutOfStation = data[i][basedOutOfStationColumn];

    if (basedOutOfStation.includes('EMT')) {
      var clockedIn = new Date(data[i][clockedInColumn]);
      var clockedOut = new Date(data[i][clockedOutColumn]);

      if (clockedIn !== '' && clockedOut !== '' && isSaturday(clockedIn)) {
        var saturdayDaySD = calculateSaturdayDaySDHours(clockedIn, clockedOut);
        if (saturdayDaySD >= 0) {
          sheet.getRange(i + 1, saturdayDaySDColumn + 1).setValue(saturdayDaySD.toFixed(1));
        } else {
          sheet.getRange(i + 1, saturdayDaySDColumn + 1).setValue(0);
        }
      }
    }
  }
}

function calculateSaturdayDaySDHours(clockedIn, clockedOut) {
  var saturdayDayStart = new Date(clockedIn);
  saturdayDayStart.setHours(8, 0, 0); // Set Saturday day start time to 8:00 AM

  var saturdayDayEnd = new Date(clockedIn);
  saturdayDayEnd.setHours(20, 0, 0); // Set Saturday day end time to 8:00 PM

  var saturdayDayHours = 0;

  if (clockedIn < saturdayDayStart) {
    clockedIn = saturdayDayStart;
  }
  if (clockedOut > saturdayDayEnd) {
    clockedOut = saturdayDayEnd;
  }

  if (clockedOut > clockedIn) {
    saturdayDayHours = (clockedOut - clockedIn) / (1000 * 60 * 60);
  }

  return saturdayDayHours;
}
