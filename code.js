/**
 * Google Apps Script to automate copying payroll data from source sheet to master sheet.
 */

// Constants
const SOURCE_SHEET_ID = '1Mlxf2Rv5G1FDXOi4J5LYwP7C0rNGQmnG18hBUr9uuPg'; // Replace with your actual source sheet ID
const SOURCE_SHEET_NAME = 'Sheet1'; // Replace with your actual source sheet name if different
const START_COLUMN = 23; // Column W (A=1, B=2, ..., W=23)

/**
 * Adds a custom menu to the master sheet upon opening.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Horizons Payroll')
    .addItem('Copy Horizons Payroll Data', 'runPayrollUpdate')
    .addToUi();
}

/**
 * Main function to run the payroll update process.
 */
function runPayrollUpdate() {
  const ui = SpreadsheetApp.getUi();
  const masterSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const tabs = masterSpreadsheet.getSheets();

  // Step 1: Get list of country-specific tabs (excluding the master sheet itself)
  const tabNames = tabs.map(sheet => sheet.getName()).filter(name => name.toLowerCase() !== 'master'); // Adjust if your master sheet has a different name

  if (tabNames.length === 0) {
    ui.alert('No country-specific tabs found.');
    return;
  }

  // Step 2: Prompt user to input tabs to process
  const tabNamesString = tabNames.join(', ');
  const tabPrompt = ui.prompt('Select Tabs to Process', 
    `Enter the names of the sheets you want to process, separated by commas.\nAvailable tabs: ${tabNamesString}`, 
    ui.ButtonSet.OK_CANCEL);

  if (tabPrompt.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Operation cancelled.');
    return;
  }

  const selectedTabsInput = tabPrompt.getResponseText();
  if (!selectedTabsInput.trim()) {
    ui.alert('No tabs entered. Operation cancelled.');
    return;
  }

  // Parse the input into an array, trimming whitespace
  const selectedTabs = selectedTabsInput.split(',').map(name => name.trim());

  // Validate selected tabs
  const invalidTabs = selectedTabs.filter(name => !tabNames.includes(name));
  if (invalidTabs.length > 0) {
    ui.alert(`The following tabs are invalid or do not exist: ${invalidTabs.join(', ')}. Please check and try again.`);
    return;
  }
  

  // Step 3: Collect employees from selected tabs
  const currentDate = new Date();
  const currentMonth = currentDate.getMonth(); // 0-based index (January is 0)
  const currentYear = currentDate.getFullYear();
  let employees = [];

  selectedTabs.forEach(tabName => {
    const sheet = masterSpreadsheet.getSheetByName(tabName);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      // Assuming headers are in the first row
      for (let i = 1; i < data.length; i++) {
        const employeeName = data[i][1]; // Column B (index 1)
        const code = data[i][2]; // Column C (index 2)
        const date = data[i][6]; // Column G (index 6)

        if (
          typeof code === 'string' &&
          code.substring(0, 3) === 'PSM' &&
          date instanceof Date &&
          date.getMonth() === currentMonth &&
          date.getFullYear() === currentYear
        ) {
          employees.push({
            name: employeeName,
            sheetName: tabName,
            rowIndex: i + 1 // For 1-based indexing in Sheets
          });
        }
      }
    }
  });

  if (employees.length === 0) {
    ui.alert('No employees found matching the criteria.');
    return;
  }


  // Step 4: Fetch payroll data from source sheet for relevant employees
  const sourceData = getSourceData(employees);
  if (!sourceData) {
    ui.alert('Failed to retrieve data from the source sheet.');
    return;
  }

  // Step 5: Store employees and sourceData in PropertiesService for later use
  PropertiesService.getUserProperties().setProperty('employees', JSON.stringify(employees));
  PropertiesService.getUserProperties().setProperty('sourceData', JSON.stringify(sourceData.data));

  // Step 6: Prompt user to map burdens
  showMappingDialog(sourceData.burdens, getMasterHeaders());
}

/**
 * Retrieves payroll data from the source sheet.
 * @param {Array} employees - List of employees to retrieve data for.
 * @returns {Object} - Contains employee data and unique burdens.
 */
function getSourceData(employees) {
  try {
    const sourceSpreadsheet = SpreadsheetApp.openById(SOURCE_SHEET_ID);
    const sourceSheet = sourceSpreadsheet.getSheetByName(SOURCE_SHEET_NAME);
    const data = sourceSheet.getDataRange().getValues();

    // Filter data for relevant employees and non-zero amounts
    const relevantData = data.slice(1).filter(row => {
      const employeeName = row[0]; // Assuming employee name is in the first column
      const burden = row[1]; // Assuming burden is in the second column
      const amount = row[3]; // Assuming amount is in the fourth column
      return employees.some(emp => emp.name === employeeName) && burden && amount !== 0;
    });

    // Extract unique burdens from the filtered data
    const burdens = Array.from(new Set(relevantData.map(row => row[1])));

    return {
      data: relevantData,
      burdens: burdens
    };
  } catch (error) {
    Logger.log('Error accessing source sheet: ' + error);
    return null;
  }
}


/**
 * Retrieves master sheet headers from Column W onwards.
 * @returns {Array} - Array of objects containing header names and their actual column indices.
 */
function getMasterHeaders() {
  const masterSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = masterSpreadsheet.getActiveSheet();
  const lastColumn = activeSheet.getLastColumn();
  
  // Calculate the number of columns to fetch from Column W onwards
  const numberOfColumns = lastColumn - (START_COLUMN - 1);
  
  // Fetch headers starting from Column W
  const headers = activeSheet.getRange(1, START_COLUMN, 1, numberOfColumns).getValues()[0];
  
  // Create an array of header objects with name and actual column index
  const headerObjects = headers.map((header, index) => ({
    name: header.trim(),
    columnIndex: START_COLUMN + index // Actual column index in the sheet
  })).filter(headerObj => headerObj.name !== ''); // Remove empty headers
  
  return headerObjects;
}


/**
 * Displays the mapping dialog.
 * @param {Array} burdens - List of burdens from the source sheet.
 * @param {Array} masterHeaders - Array of master sheet headers with names and column indices.
 */
function showMappingDialog(burdens, masterHeaders) {
  const htmlTemplate = HtmlService.createTemplateFromFile('Mapping');
  
  // Extract only the header names for the dropdown
  htmlTemplate.burdens = burdens;
  htmlTemplate.masterHeaders = masterHeaders.map(header => header.name);
  htmlTemplate.promptTitle = 'Map Burdens to Master Sheet Columns';
  htmlTemplate.promptDescription = 'For each burden from the source sheet, select the corresponding column in the master sheet where the amount should be copied. Select "Skip" to ignore a burden.';
  
  const dialog = htmlTemplate.evaluate()
    .setWidth(600)
    .setHeight(800);
  
  SpreadsheetApp.getUi().showModalDialog(dialog, 'Map Burdens to Master Sheet Columns');
}

/**
 * Processes the mapping submitted from the dialog.
 * @param {Object} mapping - Mapping of burdens to master sheet columns.
 */
function processMapping(mapping) {
  const ui = SpreadsheetApp.getUi();
  const masterSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Retrieve employees and sourceData from PropertiesService
  const employeesJSON = PropertiesService.getUserProperties().getProperty('employees');
  const sourceDataJSON = PropertiesService.getUserProperties().getProperty('sourceData');

  if (!employeesJSON || !sourceDataJSON) {
    ui.alert('Necessary data is missing. Please run the payroll update process again.');
    return;
  }

  const employees = JSON.parse(employeesJSON);
  const sourceData = JSON.parse(sourceDataJSON);

  // FX Rate will always be copied to Column M (13)
  const fxRateColumnIndex = 13; // Column M

  // Retrieve master headers with actual column indices
  const masterHeaders = getMasterHeaders();

  // Create a mapping of burden to master column index
  const burdenToColumnIndex = {};

  for (const [burden, columnName] of Object.entries(mapping)) {
    if (columnName === 'Skip') {
      // User chose to skip mapping this burden
      continue;
    }
    const headerObj = masterHeaders.find(header => header.name === columnName);
    if (headerObj) {
      burdenToColumnIndex[burden] = headerObj.columnIndex; // Use actual column index
    } else {
      ui.alert(`Mapped column "${columnName}" for burden "${burden}" not found in master sheet.`);
    }
  }

  // Log the mapping for debugging
  Logger.log('Burden to Column Index Mapping:');
  for (const [burden, colIndex] of Object.entries(burdenToColumnIndex)) {
    Logger.log(`${burden} -> Column ${colIndex} (${columnNumberToLetter(colIndex)})`);
  }

  // Proceed to copy payroll data
  copyPayrollData(masterSpreadsheet, employees, sourceData, burdenToColumnIndex, fxRateColumnIndex);

  // Clear the stored properties after processing
  PropertiesService.getUserProperties().deleteProperty('employees');
  PropertiesService.getUserProperties().deleteProperty('sourceData');

  ui.alert('Payroll data has been successfully updated.');
}

/**
 * Copies payroll data to the master sheet based on the mapping.
 * @param {Spreadsheet} masterSpreadsheet - The master spreadsheet.
 * @param {Array} employees - List of employees to process.
 * @param {Array} sourceData - Payroll data from the source sheet.
 * @param {Object} burdenToColumnIndex - Mapping of burdens to master sheet column indices.
 * @param {number} fxRateColumnIndex - Column index for FX Rate in master sheet.
 */
function copyPayrollData(masterSpreadsheet, employees, sourceData, burdenToColumnIndex, fxRateColumnIndex) {
  // Create a map of employee name to their data
  const employeeDataMap = {};
  employees.forEach(employee => {
    employeeDataMap[employee.name] = employee;
  });

  // Group source data by employee
  const groupedSourceData = {};
  sourceData.forEach(row => {
    const employeeName = row[0];
    const burden = row[1];
    const amount = row[3]; // Column D (index 3)
    const fxRate = row[7]; // Column H (index 7)

    if (!groupedSourceData[employeeName]) {
      groupedSourceData[employeeName] = {
        fxRate: fxRate,
        burdens: []
      };
    }

    groupedSourceData[employeeName].burdens.push({
      burden: burden,
      amount: amount
    });
  });

  // Iterate through each employee and update the master sheet
  for (const [employeeName, data] of Object.entries(groupedSourceData)) {
    const employeeInfo = employeeDataMap[employeeName];
    if (!employeeInfo) {
      Logger.log(`Employee "${employeeName}" not found in selected master sheets.`);
      continue;
    }

    const masterSheet = masterSpreadsheet.getSheetByName(employeeInfo.sheetName);
    if (!masterSheet) {
      Logger.log(`Sheet "${employeeInfo.sheetName}" not found in master spreadsheet.`);
      continue;
    }

    const dataRange = masterSheet.getDataRange().getValues();
    let targetRow = null;

    // Find the row with the current month's date and matching employee
    const currentDate = new Date();
    const currentMonth = currentDate.getMonth();
    const currentYear = currentDate.getFullYear();

    for (let i = 1; i < dataRange.length; i++) {
      const date = dataRange[i][6]; // Column G (index 6)
      const name = dataRange[i][1]; // Column B (index 1)

      if (
        date instanceof Date &&
        date.getMonth() === currentMonth &&
        date.getFullYear() === currentYear &&
        name === employeeName
      ) {
        targetRow = i + 1; // 1-based index
        break;
      }
    }

    if (!targetRow) {
      Logger.log(`No matching row found for employee "${employeeName}" in sheet "${employeeInfo.sheetName}".`);
      continue;
    }

    Logger.log(`Copying data for employee "${employeeName}" to sheet "${employeeInfo.sheetName}", Row ${targetRow}`);

    // Set FX Rate in Column M (always overwrite)
    masterSheet.getRange(targetRow, fxRateColumnIndex).setValue(data.fxRate);
    Logger.log(`Setting FX Rate in Column M (${fxRateColumnIndex}): ${data.fxRate}`);

    // Iterate through the burdens and set values
    data.burdens.forEach(burdenData => {
      const { burden, amount } = burdenData;

      if (amount === 0) {
        // Skip burdens with amount 0
        return;
      }

      const targetColIndex = burdenToColumnIndex[burden];
      if (!targetColIndex) {
        // Skip burdens that are not mapped
        return;
      }

      // Set the amount in the mapped column (overwrite existing value)
      masterSheet.getRange(targetRow, targetColIndex).setValue(amount);
      Logger.log(`Setting burden "${burden}" with amount ${amount} to Column ${targetColIndex} (${columnNumberToLetter(targetColIndex)})`);
    });
  }
}

/**
 * Utility function to convert column number to letter (e.g., 1 -> A).
 * @param {number} column - The column number.
 * @returns {string} - The corresponding column letter.
 */
function columnNumberToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
