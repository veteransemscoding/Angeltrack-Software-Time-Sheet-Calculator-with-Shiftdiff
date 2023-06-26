function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Calculate Payroll')
    .addItem('Start Fresh', 'startFresh')
    .addItem('Calculate Now Megan', 'runCode')
    .addToUi();
}


function startFresh() {

  deleteAndRecreateSheets();
 
}


function runCode() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Sheet1");

  removeUnusedColumns(sheet);
  sortColumnA(sheet);
  setPayrollWeek(sheet);
  addNewColumns(sheet);
  calculateTotalHours(sheet);
  calculateParamedicShiftDiff(sheet);
  calculateSaturdayDaySD(sheet);
  calculateSundayDaySD(sheet);
  calculateFridayNightSD(sheet);
  calculateSaturdayNightSD(sheet);
  calculateSundayNightSD(sheet);
  calculateMondayMorningSD(sheet);
  deleteFinalPayrollSheet();
  createFinalPayrollSheet();
  cleanUpFinalPayrollSheet('Final Payroll');
  createAndCleanupPayrollSpreadsheet();
}

function getColumnIndexByHeader(sheet, header) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === header) {
      return i;
    }
  }

  return -1; // Column not found
}
