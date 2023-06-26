function isSaturday(date) {
  if (date instanceof Date) {
    return date.getDay() === 6; // 6 represents Saturday
  }
  return false;
}


function calculateSaturdayNightSD(sheet) {
  var basedOutOfStationColumn = getColumnIndexByHeader(sheet, 'BasedOutOfStation');
  var saturdayNightSDColumn = getColumnIndexByHeader(sheet, 'Saturday Night SD');
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
        var saturdayNightSD = calculateSaturdayNightSDHours(clockedIn, clockedOut);
        if (saturdayNightSD >= 0) {
          sheet.getRange(i + 1, saturdayNightSDColumn + 1).setValue(saturdayNightSD.toFixed(1));
        } else {
          sheet.getRange(i + 1, saturdayNightSDColumn + 1).setValue(0);
        }
      }
    }
  }
}

function calculateSaturdayNightSDHours(clockedIn, clockedOut) {
  var saturdayMorningStart = new Date(clockedIn);
  saturdayMorningStart.setHours(0, 0, 0); // Set Saturday morning start time to 0:00 AM

  var saturdayMorningEnd = new Date(clockedIn);
  saturdayMorningEnd.setHours(8, 0, 0); // Set Saturday morning end time to 8:00 AM

  var saturdayNightStart = new Date(clockedIn);
  saturdayNightStart.setHours(20, 0, 0); // Set Saturday night start time to 8:00 PM

  var saturdayNightEnd = new Date(clockedIn);
  saturdayNightEnd.setHours(23, 59, 59); // Set Saturday night end time to 11:59 PM

  var saturdayNightHours = 0;

  // Calculate hours for Saturday morning
  if (clockedIn < saturdayMorningEnd && clockedOut > saturdayMorningStart) {
    if (clockedIn < saturdayMorningStart) {
      clockedIn = saturdayMorningStart;
    }
    if (clockedOut > saturdayMorningEnd) {
      clockedOut = saturdayMorningEnd;
    }
    saturdayNightHours += (clockedOut - clockedIn) / (1000 * 60 * 60);
  }

  // Calculate hours for Saturday night
  if (clockedIn < saturdayNightEnd && clockedOut > saturdayNightStart) {
    if (clockedOut > saturdayNightEnd) {
      clockedOut = saturdayNightEnd;
    }
    saturdayNightHours += (clockedOut - saturdayNightStart) / (1000 * 60 * 60);
  }

  return saturdayNightHours;
}

