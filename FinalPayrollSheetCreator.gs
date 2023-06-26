function createFinalPayrollSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var originalSheet = spreadsheet.getSheetByName('Sheet1'); // Replace 'Original Sheet' with the name of your original sheet
  var finalPayrollSheet = spreadsheet.insertSheet('Final Payroll'); // Create a new sheet named 'Final Payroll'

  // Set the column headers in the new sheet
  var columnHeaders = ['Employee Name', 'Position', 'Regular Hours', 'Overtime Hours', 'Weekend Day SD', 'Weekend Night SD', 'Paramedic SD'];
  finalPayrollSheet.getRange(1, 1, 1, columnHeaders.length).setValues([columnHeaders]);

  // Get the data from the original sheet
  var originalDataRange = originalSheet.getDataRange();
  var originalData = originalDataRange.getValues();

  // Create an object to store combined data for each employee
  var employeeData = {};

  // Loop through the rows of original data
  for (var i = 1; i < originalData.length; i++) {
    var employeeName = originalData[i][0];
    var basedOutOfStation = originalData[i][1];
    var totalHours = parseFloat(originalData[i][4]);
    var saturdayDaySD = parseNumericValue(originalData[i][6]);
    var sundayDaySD = parseNumericValue(originalData[i][7]);
    var fridayNightSD = parseNumericValue(originalData[i][8]);
    var saturdayNightSD = parseNumericValue(originalData[i][9]);
    var sundayNightSD = parseNumericValue(originalData[i][10]);
    var mondayMorningSD = parseNumericValue(originalData[i][11]);
    var paramedicSD = parseNumericValue(originalData[i][5]);
   

    // Check if the employee data already exists in the employeeData object
    if (employeeData.hasOwnProperty(employeeName)) {
      // Employee data exists, update the values
      employeeData[employeeName].totalHours += totalHours;
      employeeData[employeeName].saturdayDaySD += saturdayDaySD;
      employeeData[employeeName].sundayDaySD += sundayDaySD;
      employeeData[employeeName].fridayNightSD += fridayNightSD;
      employeeData[employeeName].saturdayNightSD += saturdayNightSD;
      employeeData[employeeName].sundayNightSD += sundayNightSD;
      employeeData[employeeName].mondayMorningSD += mondayMorningSD;
      employeeData[employeeName].paramedicSD += paramedicSD;
    } else {
      // Employee data doesn't exist, create a new entry
      employeeData[employeeName] = {
        basedOutOfStation: basedOutOfStation,
        totalHours: totalHours,
        saturdayDaySD: saturdayDaySD,
        sundayDaySD: sundayDaySD,
        fridayNightSD: fridayNightSD,
        saturdayNightSD: saturdayNightSD,
        sundayNightSD: sundayNightSD,
        mondayMorningSD: mondayMorningSD,
        paramedicSD: paramedicSD,
      };
    }
  }

  // Prepare the data for the final payroll sheet
  var finalPayrollData = [];

  for (var employee in employeeData) {
    if (employeeData.hasOwnProperty(employee)) {
      var regularHours = Math.min(employeeData[employee].totalHours, 40);
      var overtimeHours = Math.max(employeeData[employee].totalHours - 40, 0);
      var weekendDaySD = employeeData[employee].saturdayDaySD + employeeData[employee].sundayDaySD;
      var weekendNightSD = employeeData[employee].fridayNightSD + employeeData[employee].saturdayNightSD + employeeData[employee].sundayNightSD + employeeData[employee].mondayMorningSD;
      var paramedicSD = employeeData[employee].paramedicSD;

      var rowData = [
        employee,
        employeeData[employee].basedOutOfStation,
        regularHours > 0 ? regularHours : '',
        overtimeHours > 0 ? overtimeHours : '',
        weekendDaySD > 0 ? weekendDaySD : '',
        weekendNightSD > 0 ? weekendNightSD : '',
        paramedicSD > 0 ? paramedicSD : '',
      ];
      finalPayrollData.push(rowData);
    }
  }

  // Write the data to the final payroll sheet
  finalPayrollSheet.getRange(2, 1, finalPayrollData.length, finalPayrollData[0].length).setValues(finalPayrollData);

  // Adjust column widths
  finalPayrollSheet.autoResizeColumns(1, finalPayrollData[0].length);
}


function parseNumericValue(value) {
  if (typeof value === 'number') {
    return value;
  } else if (typeof value === 'string') {
    var parsedValue = parseFloat(value.replace(/[^\d.-]/g, ''));
    return isNaN(parsedValue) ? 0 : parsedValue;
  } else {
    return 0;
  }
}
