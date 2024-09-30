function processPayroll() {
  Logger.clear();  // Clear previous logs
  
  // URL of the source Google Sheet (Sheet 1)
  var sourceSheetUrl = 'https://docs.google.com/spreadsheets/d/1lPuhJ9IO3U_yH0R7ohKpH4WzEfllpaRPD8ys3e8-2Iw/edit?usp=sharing';
  
  Logger.log("Starting payroll process...");

  // Open the source spreadsheet
  var sourceSpreadsheet;
  try {
    sourceSpreadsheet = SpreadsheetApp.openByUrl(sourceSheetUrl);
    Logger.log("Source spreadsheet opened successfully.");
  } catch (e) {
    Logger.log("Error: Unable to open the source spreadsheet. " + e.toString());
    SpreadsheetApp.getUi().alert('Error: Unable to open the source spreadsheet.');
    return;
  }
  
  // Get the first sheet of the source spreadsheet
  var sheet1 = sourceSpreadsheet.getSheets()[0];
  
  // Get data from Sheet 1
  var sheet1Data = sheet1.getDataRange().getValues();
  Logger.log("Sheet 1 data retrieved. Total rows: " + sheet1Data.length);
  
  // Get the active (destination) spreadsheet
  var destSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get Sheet28 from the destination spreadsheet
  var sheet28 = destSpreadsheet.getSheetByName("Sheet28");
  if (!sheet28) {
    Logger.log("Error: Sheet28 not found in the destination spreadsheet.");
    SpreadsheetApp.getUi().alert('Error: Sheet28 not found in the destination spreadsheet.');
    return;
  }
  
  // Get data from Sheet28
  var sheet28Data = sheet28.getDataRange().getValues();
  Logger.log("Sheet28 data retrieved. Total rows: " + sheet28Data.length);
  
  // Get the current month
  var currentMonth = new Date().getMonth() + 1;
  Logger.log("Current month: " + currentMonth);
  
  // Store updates in memory for batch processing
  var updates = [];
  
  // Process each employee in Sheet 1
  var employeeData = {};
  for (var i = 1; i < sheet1Data.length; i++) {
    var employeeName = sheet1Data[i][0];  // Column A: Employee Name
    if (!employeeName) continue;  // Skip empty rows
    
    if (!employeeData[employeeName]) {
      employeeData[employeeName] = {
        grossIncome: 0,
        zwWgWHK_WGA: 0,
        wkoSurcharge: 0
      };
    }
    
    var burdenName = sheet1Data[i][1];  // Column B: Employer Burden Name in Sheet 1
    var burdenAmount = parseFloat(sheet1Data[i][3]) || 0;  // Column D: Amount in Sheet 1
    
    if (burdenName === "Gross Income") {
      employeeData[employeeName].grossIncome += burdenAmount;
    } else if (burdenName === "Gediff. WGA wg  WHK (Return to work fund) together with ZW") {
      employeeData[employeeName].zwWgWHK_WGA += burdenAmount;
    } else if (burdenName === "WKOSurcharge Childcare Act") {
      employeeData[employeeName].wkoSurcharge += burdenAmount;
    }
  }
  
  // Process each unique employee
  for (var employeeName in employeeData) {
    Logger.log("Processing employee: " + employeeName);
    
    // Find the rows for the employee in Sheet28
    var employeeRows = findEmployeeRows(sheet28Data, employeeName, currentMonth);
    if (employeeRows.length === 0) {
      Logger.log("No matching rows found for employee: " + employeeName);
      continue;  // Skip if no matching rows found
    }
    
    // Prepare updates for this employee
    for (var i = 0; i < employeeRows.length; i++) {
      var row = employeeRows[i];
      var updatesForEmployee = [
        { row: row, column: 25, value: employeeData[employeeName].grossIncome },  // Column Y
        { row: row, column: 41, value: employeeData[employeeName].zwWgWHK_WGA },  // Column AO
        { row: row, column: 40, value: employeeData[employeeName].wkoSurcharge }  // Column AN
      ];
      
      Logger.log("Prepared updates for employee: " + employeeName + ", Row: " + row + ", Total updates: " + updatesForEmployee.length);
      updates.push(...updatesForEmployee);
    }
  }
  
  // Apply all updates in batch
  var successCount = 0;
  var failureCount = 0;
  for (var i = 0; i < updates.length; i++) {
    var update = updates[i];
    try {
      var range = sheet28.getRange(update.row, update.column);
      var oldValue = range.getValue();
      range.setValue(update.value);
      var newValue = range.getValue();
      if (newValue === update.value) {
        Logger.log("Update successful - Row: " + update.row + ", Column: " + update.column + ", Old Value: " + oldValue + ", New Value: " + newValue);
        successCount++;
      } else {
        Logger.log("Update failed - Row: " + update.row + ", Column: " + update.column + ", Attempted Value: " + update.value + ", Actual New Value: " + newValue);
        failureCount++;
      }
    } catch (e) {
      Logger.log("Error applying update - Row: " + update.row + ", Column: " + update.column + ", Value: " + update.value + ". Error: " + e.toString());
      failureCount++;
    }
  }
  
  Logger.log("Payroll process complete. Successful updates: " + successCount + ", Failed updates: " + failureCount);
  SpreadsheetApp.getUi().alert('Process Complete', 'Payroll data has been updated for all employees in Sheet28. Successful updates: ' + successCount + ', Failed updates: ' + failureCount, SpreadsheetApp.getUi().ButtonSet.OK);
}

function findEmployeeRows(sheetData, employeeName, currentMonth) {
  var rows = [];
  for (var i = 1; i < sheetData.length; i++) {  // Start from 1 to skip header row
    var rowEmployeeName = sheetData[i][1];  // Column B: Employee Name
    var payDate = sheetData[i][7];  // Column H: Pay Date
    
    if (rowEmployeeName === employeeName) {
      if (payDate instanceof Date && payDate.getMonth() + 1 === currentMonth) {
        rows.push(i + 1);  // Adding 1 because array is 0-indexed but sheets are 1-indexed
      } else if (typeof payDate === 'string') {
        var parsedDate = new Date(payDate);
        if (!isNaN(parsedDate) && parsedDate.getMonth() + 1 === currentMonth) {
          rows.push(i + 1);
        }
      }
    }
  }
  Logger.log("Found " + rows.length + " rows for employee: " + employeeName);
  return rows;
}
