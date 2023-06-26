function cleanUpFinalPayrollSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var finalPayrollSheet = spreadsheet.getSheetByName('Final Payroll'); // Replace 'Final Payroll' with the name of your Final Payroll sheet

  // Set font style and alignment
  var range = finalPayrollSheet.getRange(1, 1, finalPayrollSheet.getLastRow(), finalPayrollSheet.getLastColumn());
  range.setFontFamily('Verdana').setFontSize(12).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // Bold and center header
  var headersRange = finalPayrollSheet.getRange(1, 1, 1, finalPayrollSheet.getLastColumn());
  headersRange.setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // Resize columns
  finalPayrollSheet.setColumnWidth(1, 250); // Employee Name
  finalPayrollSheet.setColumnWidth(2, 110); // Position
  finalPayrollSheet.setColumnWidth(3, 110); // Regular Hours
  finalPayrollSheet.setColumnWidth(4, 110); // Overtime Hours
  finalPayrollSheet.setColumnWidth(5, 110); // Weekend Day SD
  finalPayrollSheet.setColumnWidth(6, 110); // Weekend Night SD
  finalPayrollSheet.setColumnWidth(7, 110); // Paramedic SD

  // Delete columns 8 and above
  var numColumnsToDelete = finalPayrollSheet.getMaxColumns() - 7; // Calculate the number of columns to delete
  if (numColumnsToDelete > 0) {
    finalPayrollSheet.deleteColumns(8, numColumnsToDelete);
  }

  // Delete empty rows
  var dataRange = finalPayrollSheet.getRange("A:A");
  var values = dataRange.getValues();
  var numRows = values.length;

  var lastRowIndex = numRows;
  for (var i = numRows - 1; i >= 0; i--) {
    if (values[i][0] === "") {
      lastRowIndex = i;
    } else {
      break;
    }
  }

  if (lastRowIndex < numRows) {
    var numEmptyRows = numRows - lastRowIndex - 1;
    if (numEmptyRows > 0) {
      finalPayrollSheet.deleteRows(lastRowIndex + 2, numEmptyRows);
    }
  }
}
