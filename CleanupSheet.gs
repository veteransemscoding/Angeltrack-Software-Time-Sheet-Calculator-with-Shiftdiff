//removes the columns and keeps name, location, clock in and out

function removeUnusedColumns(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var columnsToDelete = [];

  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];

    if (
      header !== 'Employee Name' &&
      header !== 'BasedOutOfStation' &&
      header !== 'ClockedIn' &&
      header !== 'ClockedOut'
    ) {
      columnsToDelete.push(i + 1);
    }
  }

  for (var j = columnsToDelete.length - 1; j >= 0; j--) {
    sheet.deleteColumn(columnsToDelete[j]);
  }
}

// sorts the columns by name
function sortColumnA(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(2, 1, lastRow - 1, lastColumn); // Assuming the data starts from row 2, adjust as needed

  range.sort([{ column: 1, ascending: true }]); // Sort the range based on column A (Employee Name) in ascending order
}


