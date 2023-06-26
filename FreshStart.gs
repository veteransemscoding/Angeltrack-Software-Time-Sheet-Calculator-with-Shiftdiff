function deleteAndRecreateSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Delete the existing "Sheet1" sheet
  var sheet1 = spreadsheet.getSheetByName("Sheet1");
  spreadsheet.deleteSheet(sheet1);
  
  // Create a new "Sheet1" sheet
  spreadsheet.insertSheet("Sheet1");
  
  // Delete the existing "Final Payroll" sheet
  var finalPayrollSheet = spreadsheet.getSheetByName("Final Payroll");
  spreadsheet.deleteSheet(finalPayrollSheet);
}
