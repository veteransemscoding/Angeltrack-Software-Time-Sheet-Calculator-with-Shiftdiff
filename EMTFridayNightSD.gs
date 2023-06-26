function isFriday(date) {
  return date.getDay() === 5; // 5 represents Friday (Saturday is 6, Sunday is 0, and so on)
}


function calculateFridayNightSD(sheet) {
  var basedOutOfStationColumn = getColumnIndexByHeader(sheet, 'BasedOutOfStation');
  var fridayNightSDColumn = getColumnIndexByHeader(sheet, 'Friday Night SD');
  var clockedInColumn = getColumnIndexByHeader(sheet, 'ClockedIn');
  var clockedOutColumn = getColumnIndexByHeader(sheet, 'ClockedOut');

  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  for (var i = 1; i < data.length; i++) {
    var basedOutOfStation = data[i][basedOutOfStationColumn];

    if (basedOutOfStation.includes('EMT')) {
      var clockedIn = new Date(data[i][clockedInColumn]);
      var clockedOut = new Date(data[i][clockedOutColumn]);

      if (clockedIn !== '' && clockedOut !== '' && isFriday(clockedIn)) {
        var fridayNightSD = calculateFridayNightSDHours(clockedIn, clockedOut);
        if (fridayNightSD >= 0) {
          sheet.getRange(i + 1, fridayNightSDColumn + 1).setValue(fridayNightSD.toFixed(1));
        } else {
          sheet.getRange(i + 1, fridayNightSDColumn + 1).setValue(0);
        }
      }
    }
  }
}

function calculateFridayNightSDHours(clockedIn, clockedOut) {
  var fridayNightStart = new Date(clockedIn);
  fridayNightStart.setHours(20, 0, 0); // Set Friday night start time to 8:00 PM

  var fridayNightEnd = new Date(clockedIn);
  fridayNightEnd.setHours(23, 59, 59); // Set Friday night end time to 11:59 PM

  var fridayNightHours = 0;

  if (clockedIn < fridayNightEnd && clockedOut > fridayNightStart) {
    if (clockedOut > fridayNightEnd) {
      clockedOut = fridayNightEnd;
    }
    fridayNightHours = (clockedOut - fridayNightStart) / (1000 * 60 * 60);
  }

  return fridayNightHours;
}

function isFriday(date) {
  return date.getDay() === 5; // 5 represents Friday (Saturday is 6, Sunday is 0, and so on)
}
