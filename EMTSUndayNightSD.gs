function isSunday(date) {
  if (date instanceof Date) {
    return date.getDay() === 0; // 0 represents Sunday
  }
  return false;
}


function calculateSundayNightSD(sheet) {
  var basedOutOfStationColumn = getColumnIndexByHeader(sheet, 'BasedOutOfStation');
  var sundayNightSDColumn = getColumnIndexByHeader(sheet, 'Sunday Night SD');
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
        var sundayNightSD = calculateSundayNightSDHours(clockedIn, clockedOut);
        if (sundayNightSD >= 0) {
          sheet.getRange(i + 1, sundayNightSDColumn + 1).setValue(sundayNightSD.toFixed(1));
        } else {
          sheet.getRange(i + 1, sundayNightSDColumn + 1).setValue(0);
        }
      }
    }
  }
}

function calculateSundayNightSDHours(clockedIn, clockedOut) {
  var sundayMorningStart = new Date(clockedIn);
  sundayMorningStart.setHours(0, 0, 0); // Set Sunday morning start time to 0:00 AM

  var sundayMorningEnd = new Date(clockedIn);
  sundayMorningEnd.setHours(8, 0, 0); // Set Sunday morning end time to 8:00 AM

  var sundayNightStart = new Date(clockedIn);
  sundayNightStart.setHours(20, 0, 0); // Set Sunday night start time to 8:00 PM

  var sundayNightEnd = new Date(clockedIn);
  sundayNightEnd.setHours(23, 59, 59); // Set Sunday night end time to 11:59 PM

  var sundayNightHours = 0;

  if (clockedIn < sundayMorningEnd && clockedOut > sundayMorningStart) {
    if (clockedIn < sundayMorningStart) {
      clockedIn = sundayMorningStart;
    }
    if (clockedOut > sundayMorningEnd) {
      clockedOut = sundayMorningEnd;
    }
    sundayNightHours += (clockedOut - clockedIn) / (1000 * 60 * 60);
  }

  if (clockedIn < sundayNightEnd && clockedOut > sundayNightStart) {
    if (clockedOut > sundayNightEnd) {
      clockedOut = sundayNightEnd;
    }
    sundayNightHours += (clockedOut - sundayNightStart) / (1000 * 60 * 60);
  }

  return sundayNightHours;
}

