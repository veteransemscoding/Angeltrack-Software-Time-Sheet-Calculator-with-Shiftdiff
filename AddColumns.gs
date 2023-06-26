//creates new columns to calculate the shift diffs
function addNewColumns(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var newHeaders = [
    'Total Hours',
    'Paramedic SD',
    'Saturday Day SD',
    'Sunday Day SD',
    'Friday Night SD',
    'Saturday Night SD',
    'Sunday Night SD',
    'Monday Morning SD'
  ];

  var columnCount = headers.length;
  var newColumnStartIndex = columnCount + 1;

  for (var i = 0; i < newHeaders.length; i++) {
    var header = newHeaders[i];

    if (!headers.includes(header)) {
      sheet.getRange(1, newColumnStartIndex + i).setValue(header);
    }
  }
}
