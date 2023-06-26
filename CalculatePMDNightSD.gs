function calculateParamedicShiftDiff(sheet) {
  var basedOutOfStationColumn = getColumnIndexByHeader(sheet, 'BasedOutOfStation');
  var paramedicSDColumn = getColumnIndexByHeader(sheet, 'Paramedic SD');
  var clockedInColumn = getColumnIndexByHeader(sheet, 'ClockedIn');
  var clockedOutColumn = getColumnIndexByHeader(sheet, 'ClockedOut');

  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  for (var i = 1; i < data.length; i++) {
    var basedOutOfStation = data[i][basedOutOfStationColumn];

    if (basedOutOfStation.includes('Paramedic')) {
      var clockedIn = new Date(data[i][clockedInColumn]);
      var clockedOut = new Date(data[i][clockedOutColumn]);

      if (clockedIn !== '' && clockedOut !== '') {
        var paramedicShiftDiff = calculateParamedicShiftDiffHours(clockedIn, clockedOut);
        if (paramedicShiftDiff >= 0) {
          sheet.getRange(i + 1, paramedicSDColumn + 1).setValue(paramedicShiftDiff.toFixed(1));
        } else {
          sheet.getRange(i + 1, paramedicSDColumn + 1).setValue(0);
        }
      }
    }
  }
}

function calculateParamedicShiftDiffHours(clockedIn, clockedOut) {
  var nightShiftStart = new Date(clockedIn);
  nightShiftStart.setHours(20, 0, 0); // Set night shift start time to 8:00 PM

  var nightShiftEnd = new Date(clockedIn);
  nightShiftEnd.setDate(nightShiftEnd.getDate() + 1); // Add one day to the clockedIn date
  nightShiftEnd.setHours(8, 0, 0); // Set night shift end time to 8:00 AM

  var nightShiftHours = 0;

  if (clockedOut <= nightShiftStart || clockedIn >= nightShiftEnd) {
    // No night shift hours within the shift range
    nightShiftHours = 0;
  } else if (clockedIn <= nightShiftStart && clockedOut >= nightShiftEnd) {
    // The entire shift falls within the night shift range
    nightShiftHours = 12;
  } else if (clockedIn <= nightShiftStart && clockedOut > nightShiftStart && clockedOut < nightShiftEnd) {
    // Shift starts before night shift start and ends within the night shift range
    nightShiftHours = (clockedOut - nightShiftStart) / (1000 * 60 * 60);
  } else if (clockedIn >= nightShiftStart && clockedIn < nightShiftEnd && clockedOut >= nightShiftEnd) {
    // Shift starts within the night shift range and ends after night shift end
    nightShiftHours = (nightShiftEnd - clockedIn) / (1000 * 60 * 60);
  } else if (clockedIn >= nightShiftStart && clockedOut <= nightShiftEnd) {
    // Shift falls within the night shift range but starts and ends outside the range
    nightShiftHours = (clockedOut - clockedIn) / (1000 * 60 * 60);
  }

  return nightShiftHours;
}
