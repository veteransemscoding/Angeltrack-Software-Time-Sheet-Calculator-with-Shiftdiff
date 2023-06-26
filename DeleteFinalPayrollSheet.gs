function deleteFinalPayrollSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var finalPayrollSheet = spreadsheet.getSheetByName('Final Payroll');

  if (finalPayrollSheet) {
    spreadsheet.deleteSheet(finalPayrollSheet);
    Logger.log('Final Payroll sheet deleted.');
  } else {
    Logger.log('Final Payroll sheet does not exist.');
  }
}
