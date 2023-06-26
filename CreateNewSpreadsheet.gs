function createAndCleanupPayrollSpreadsheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentFinalPayrollSheet = spreadsheet.getSheetByName('Final Payroll'); // Replace 'Final Payroll' with the name of your "Final Payroll" sheet
 
 
  var companyName = Browser.inputBox('Company Name', 'Please enter the company name:', Browser.Buttons.OK_CANCEL);
 
  if (companyName == 'cancel') {
    // User canceled the input
    return;
  }
 
  // Get the current date
  var currentDate = new Date();

  // Calculate the start and end dates for the payroll week
  var endDayOfWeek = currentDate.getDay() === 0 ? 7 : currentDate.getDay(); // Sunday is 0, Monday is 1, and so on
  var previousSunday = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() - endDayOfWeek);
  var twoSundaysAgo = new Date(previousSunday.getFullYear(), previousSunday.getMonth(), previousSunday.getDate() - 7);
  var startOfWeek = new Date(twoSundaysAgo.getFullYear(), twoSundaysAgo.getMonth(), twoSundaysAgo.getDate() + 1);
  var endOfWeek = new Date(previousSunday.getFullYear(), previousSunday.getMonth(), previousSunday.getDate());

  // Format the payroll week range for the spreadsheet name
  var payrollWeekName = formatDate(startOfWeek) + '-' + formatDate(endOfWeek) + ' ' + companyName + ' Payroll';

  // Create a new spreadsheet
  var newSpreadsheet = SpreadsheetApp.create(payrollWeekName);

  // Get the ID and name of the new spreadsheet
  var newSpreadsheetId = newSpreadsheet.getId();
  var newSpreadsheetName = newSpreadsheet.getName();

  // Get the data from the current "Final Payroll" sheet
  var finalPayrollData = currentFinalPayrollSheet.getDataRange().getValues();

  // Copy the data to the new spreadsheet
  var newSheet = newSpreadsheet.getActiveSheet();
  newSheet.getRange(1, 1, finalPayrollData.length, finalPayrollData[0].length).setValues(finalPayrollData);

  // Rename the sheet to 'Final Payroll'
  newSheet.setName('Final Payroll');

  // Log the URL of the new spreadsheet
  Logger.log('New Payroll Spreadsheet URL: ' + newSpreadsheet.getUrl());

  // Call the cleanupPayrollSpreadsheet function
  cleanupPayrollSpreadsheet(newSpreadsheet);

  // Return the name of the new spreadsheet
  return newSpreadsheetName;
}


function createAndCleanupPayrollSpreadsheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentFinalPayrollSheet = spreadsheet.getSheetByName('Final Payroll'); // Replace 'Final Payroll' with the name of your "Final Payroll" sheet
  // Prompt the user for the company name
  var companyName = Browser.inputBox('Company Name', 'Please enter the company name:', Browser.Buttons.OK_CANCEL);
 
  if (companyName == 'cancel') {
    // User canceled the input
    return;
  }
 
  // Get the current date
  var currentDate = new Date();




  // Calculate the start and end dates for the payroll week
  var endDayOfWeek = currentDate.getDay() === 0 ? 7 : currentDate.getDay(); // Sunday is 0, Monday is 1, and so on
  var previousSunday = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() - endDayOfWeek);
  var twoSundaysAgo = new Date(previousSunday.getFullYear(), previousSunday.getMonth(), previousSunday.getDate() - 7);
  var startOfWeek = new Date(twoSundaysAgo.getFullYear(), twoSundaysAgo.getMonth(), twoSundaysAgo.getDate() + 1);
  var endOfWeek = new Date(previousSunday.getFullYear(), previousSunday.getMonth(), previousSunday.getDate());




  // Format the payroll week range for the spreadsheet name
  var payrollWeekName = formatDate(startOfWeek) + '-' + formatDate(endOfWeek) + ' ' + companyName + ' Payroll';




  // Create a new spreadsheet
  var newSpreadsheet = SpreadsheetApp.create(payrollWeekName);




  // Get the ID and name of the new spreadsheet
  var newSpreadsheetId = newSpreadsheet.getId();
  var newSpreadsheetName = newSpreadsheet.getName();




  // Get the data from the current "Final Payroll" sheet
  var finalPayrollData = currentFinalPayrollSheet.getDataRange().getValues();




  // Copy the data to the new spreadsheet
  var newSheet = newSpreadsheet.getActiveSheet();
  newSheet.getRange(1, 1, finalPayrollData.length, finalPayrollData[0].length).setValues(finalPayrollData);




  // Rename the sheet to 'Final Payroll'
  newSheet.setName('Final Payroll');




  // Log the URL of the new spreadsheet
  Logger.log('New Payroll Spreadsheet URL: ' + newSpreadsheet.getUrl());




  // Call the cleanupPayrollSpreadsheet function
  cleanupPayrollSpreadsheet(newSpreadsheet);




  // Return the name of the new spreadsheet
  return newSpreadsheetName;
}


function cleanupPayrollSpreadsheet(spreadsheet) {
  var finalPayrollSpreadsheet = SpreadsheetApp.openById(spreadsheet.getId());
  var finalPayrollSheet = finalPayrollSpreadsheet.getSheetByName('Final Payroll'); // Replace 'Final Payroll' with the name of your Final Payroll sheet




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




  // Create additional sheets
  var emtBillerDispatcherSheet = finalPayrollSpreadsheet.insertSheet('EMT Biller Dispatcher');
  var netDriversSheet = finalPayrollSpreadsheet.insertSheet('NET Drivers');




  // Copy header to the new sheets
  finalPayrollSheet.getRange(1, 1, 1, finalPayrollSheet.getLastColumn()).copyTo(emtBillerDispatcherSheet.getRange(1, 1));
  finalPayrollSheet.getRange(1, 1, 1, finalPayrollSheet.getLastColumn()).copyTo(netDriversSheet.getRange(1, 1));




  // Copy rows to the new sheets based on the 'Position' column
  var positionColumnIndex = 2; // Assuming 'Position' is in the second column (column B)




  for (var row = 2; row <= finalPayrollSheet.getLastRow(); row++) {
    var positionValue = finalPayrollSheet.getRange(row, positionColumnIndex).getValue();
   
    if (positionValue === 'CV Van') {
      finalPayrollSheet.getRange(row, 1, 1, finalPayrollSheet.getLastColumn()).copyTo(netDriversSheet.getRange(netDriversSheet.getLastRow() + 1, 1));
    } else {
      finalPayrollSheet.getRange(row, 1, 1, finalPayrollSheet.getLastColumn()).copyTo(emtBillerDispatcherSheet.getRange(emtBillerDispatcherSheet.getLastRow() + 1, 1));
    }
  }




  // Perform cleanup on the new sheets
  cleanupSheet(emtBillerDispatcherSheet);
  cleanupSheet(netDriversSheet);
}

function cleanupSheet(sheet) {
  // Set font style and alignment
  var range = sheet.getDataRange();
  range.setFontFamily('Verdana').setFontSize(12).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);


  // Bold and center header
  var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headersRange.setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);


  // Resize columns
  sheet.setColumnWidth(1, 130); // Employee Name
  sheet.setColumnWidth(2, 130); // Position
  sheet.setColumnWidth(3, 110); // Regular Hours
  sheet.setColumnWidth(4, 110); // Overtime Hours
  sheet.setColumnWidth(5, 110); // Weekend Day SD
  sheet.setColumnWidth(6, 110); // Weekend Night SD
  sheet.setColumnWidth(7, 110); // Paramedic SD


  // Delete columns 8 and above
  var numColumnsToDelete = sheet.getMaxColumns() - 7; // Calculate the number of columns to delete
  if (numColumnsToDelete > 0) {
    sheet.deleteColumns(8, numColumnsToDelete);
  }


  // Delete empty rows
  var dataRange = sheet.getRange("A:A");
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
      sheet.deleteRows(lastRowIndex + 2, numEmptyRows);
    }
  }

  var lastRow = sheet.getLastRow();
  var lastRowRange = sheet.getRange(lastRow + 1, 1, 1, sheet.getLastColumn());


  // Set font style and properties for the last row
  lastRowRange.setFontFamily('Verdana').setFontSize(12).setFontWeight('bold');


  // Set "Totals" in column B
  var totalsRange = sheet.getRange(lastRow + 1, 2);
  totalsRange.setValue("Totals");


  // Calculate totals for each column
  var columnsToTotal = [3, 4, 5, 6, 7]; // Columns C, D, E, F, G
  for (var i = 0; i < columnsToTotal.length; i++) {
    var columnToTotal = columnsToTotal[i];
    var totalFormula = "=SUM(" + sheet.getRange(2, columnToTotal).getA1Notation() + ":" + sheet.getRange(lastRow, columnToTotal).getA1Notation() + ")";
    var totalRange = sheet.getRange(lastRow + 1, columnToTotal);
    totalRange.setFormula(totalFormula);
  }


var newLastNameColumn = sheet.insertColumnAfter(1); // Insert a column after the first column
var newFirstNameColumn = sheet.insertColumnAfter(1); // Insert a column after the first column
sheet.deleteColumn(4); // Delete the fourth column
var newParamedicsdColumn = sheet.insertColumnAfter(8);
var newWeekendNightColumn = sheet.insertColumnAfter(7);
var newWeekendDayColumn = sheet.insertColumnAfter(6);

// Set the headers for the new columns
sheet.getRange(1, 1).setValue("Last Name"); // Change the header of the first column to "Last Name"
sheet.getRange(1, 2).setValue("First Name"); // Change the header of the second column to "First Name"
sheet.getRange(1, 3).setValue("Employee ID"); // Change the header of the third column to "Employee ID"
sheet.getRange(1, 11).setValue("Paramedic Night Rate");
sheet.getRange(1, 9).setValue("Weekend Night Rate");
sheet.getRange(1, 7).setValue("Weekend Day Rate");

// Get the range of the original names column
var namesRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
var names = namesRange.getValues();


var middleNameColumn = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1);
middleNameColumn.setNumberFormat("@"); // Set the format of the middle name column to plain text


for (var i = 0; i < names.length; i++) {
  var fullName = names[i][0];
  fullName = fullName.replace(",", ""); // Remove commas from the full name
  var nameParts = fullName.split(" ");
  var lastName = nameParts[0];
  var firstName = nameParts[1];
  var middleName = nameParts[2] || ""; // Use an empty string if middle name doesn't exist
  sheet.getRange(i + 2, 1).setValue(lastName);
  sheet.getRange(i + 2, 2).setValue(firstName);
  sheet.getRange(i + 2, 3).setValue(middleName); // Update the values in the third column
}

//Make the Paramedic Rate Column = 5
var paramedicNightRateColumn = sheet.getRange(2, 11, sheet.getLastRow() - 1, 1);
var paramedicNightRateValues = paramedicNightRateColumn.getValues();

for (var i = 0; i < paramedicNightRateValues.length; i++) {
  paramedicNightRateValues[i][0] = 5;
}

paramedicNightRateColumn.setValues(paramedicNightRateValues);

// Make the weekend Night Rate = 2
var weekendNightRateColumn = sheet.getRange(2, 9, sheet.getLastRow() - 1, 1);
var weekendNightRateValues = weekendNightRateColumn.getValues();

for (var i = 0; i < weekendNightRateValues.length; i++) {
  weekendNightRateValues[i][0] = 2;
}

weekendNightRateColumn.setValues(weekendNightRateValues);


//Make the weekend date rate = 1
var weekendDayRateColumn = sheet.getRange(2, 7, sheet.getLastRow() - 1, 1);
var weekendDayRateValues = weekendDayRateColumn.getValues();

for (var i = 0; i < weekendDayRateValues.length; i++) {
  weekendDayRateValues[i][0] = 1;
}

weekendDayRateColumn.setValues(weekendDayRateValues);



var lastRowend = sheet.getLastRow();
sheet.getRange(lastRowend, 1, 1, sheet.getLastColumn()).clearContent();


}







function formatDate(date) {
  var day = date.getDate();
  var month = date.getMonth() + 1;
  var year = date.getFullYear();




  // Add leading zeros if necessary
  day = (day < 10) ? '0' + day : day;
  month = (month < 10) ? '0' + month : month;




  return month + '/' + day + '/' + year;
}





