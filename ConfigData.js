var configData = {
  sourceSheetsConfig: [
    {
      url: "https://docs.google.com/spreadsheets/d/172a99fheqj4HLIPsPrpPwWZsVJJsg8mzBpH9P4ZISRY/edit#gid=1688091276",
      sheetName: "MARCOS",
      targetColumn: "C",
      weekRows: [3103, 3118],
      startRow: 3350
    },
    {
      url: "https://docs.google.com/spreadsheets/d/1zIYlgBOZstb7xsMCHSlDXK-wEqenT5iNXvkMj_htCRo/edit#gid=1688091276",
      sheetName: "WILLIAM",
      targetColumn: "D",
      weekRows: [1791, 1806],
      startRow: 2030
    },
    {
      url: "https://docs.google.com/spreadsheets/d/1jZ_uHeVUW128OI30DqI_tJOz6GWWHu98M6xo7jHLtSQ/edit#gid=1688091276",
      sheetName: "GABRIEL",
      targetColumn: "E",
      weekRows: [3102, 3117],
      startRow: 3340
    },
    {
      url: "https://docs.google.com/spreadsheets/d/1VzYfowTgjKbvI2cz5WThvvtuqYokoAZWMA0GL0nYSZ0/edit#gid=1688091276",
      sheetName: "FÁBIO",
      targetColumn: "F",
      weekRows: [3126, 3141],
      startRow: 3370
    },
    {
      url: "https://docs.google.com/spreadsheets/d/1q4GYq9L3ouItGFDte-e75U4cTogBYQ_JpgWHxEg-2CU/edit#gid=1688091276",
      sheetName: "ILDO",
      targetColumn: "G",
      weekRows: [3099, 3114],
      startRow: 3340
    },
    {
      url: "https://docs.google.com/spreadsheets/d/1A263AEsXeBZNAdsPLBC3bCE2GSsMY3Ux0nBvlSfrpuE/edit#gid=83327853",
      sheetName: "GUILHERME",
      targetColumn: "H",
      weekRows: [1975, 1990],
      startRow: 2220
    },
    {
      url: "https://docs.google.com/spreadsheets/d/1pc6f6kaplpL9GY65LKQB22xtaLSK8T6ZLMCXjaXo6_4/edit#gid=83327853",
      sheetName: "LUANA",
      targetColumn: "I",
      weekRows: [1974, 1989],
      startRow: 2220
    }
  ],
};

function initializeScriptProperties() {
  const scriptProperties = PropertiesService.getScriptProperties();
  if (!scriptProperties.getProperty('sourceSheetsConfig')) {
    scriptProperties.setProperty('sourceSheetsConfig', JSON.stringify(configData.sourceSheetsConfig));
  }
  if (!scriptProperties.getProperty('currentWeek')) {
    scriptProperties.setProperty('currentWeek', configData.currentWeek.toString());
  }
  if (!scriptProperties.getProperty('summarySheetStartRows')) {
    scriptProperties.setProperty('summarySheetStartRows', JSON.stringify(configData.summarySheetStartRows));
  }
  Logger.log('Script properties initialized');
}

function getSourceSheetsConfigFromScript() {
  const storedConfig = JSON.parse(PropertiesService.getScriptProperties().getProperty('sourceSheetsConfig')) || [];
  return configData.sourceSheetsConfig.map(sheet => {
    const dynamicSheet = storedConfig.find(dSheet => dSheet.sheetName === sheet.sheetName);
    if (dynamicSheet) {
      dynamicSheet.weekRows.forEach((row, index) => {
        sheet.weekRows[index] = row;
      });
    }
    return sheet;
  });
}

function findCurrentWeekRow(sheetConfig) {
  const sourceSpreadsheet = SpreadsheetApp.openByUrl(sheetConfig.url);
  const sourceSheet = sourceSpreadsheet.getSheetByName(sheetConfig.sheetName);
  
  const today = new Date();
  const todayDateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yyyy");

  const range = sourceSheet.getDataRange();
  const values = range.getValues();

  Logger.log(`Searching for today's date (${todayDateString}) in sheet ${sheetConfig.sheetName} starting from row ${sheetConfig.startRow}`);

  for (let row = sheetConfig.startRow - 1; row < values.length; row++) { // Adjust index to 0-based
    for (let col = 0; col < values[row].length; col++) {
      const cellValue = values[row][col];
      if (cellValue instanceof Date) {
        const cellDateString = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "dd/MM/yyyy");
        if (cellDateString === todayDateString) {
          Logger.log(`Found today's date (${todayDateString}) in cell [${row + 1}, ${col + 1}]`);
          return row + 1; // Return the row number (1-based index)
        }
      }
    }
  }

  throw new Error(`Today's date (${todayDateString}) not found in sheet ${sheetConfig.sheetName}`);
}

function findStartingRowInSummarySheet(summarySheet, todayDateString) {
  const startRow = 220;
  const range = summarySheet.getRange(startRow, 1, summarySheet.getMaxRows() - startRow + 1, summarySheet.getMaxColumns());
  const values = range.getValues();
  let firstTodayRow = -1;

  Logger.log(`Searching for today's date (${todayDateString}) in the summary sheet starting from row ${startRow}`);

  // Find the first occurrence of today's date
  for (let row = 0; row < values.length; row++) {
    for (let col = 0; col < values[row].length; col++) {
      if (values[row][col] instanceof Date) {
        const cellDateString = Utilities.formatDate(values[row][col], Session.getScriptTimeZone(), "dd/MM/yyyy");
        if (cellDateString === todayDateString) {
          if (firstTodayRow === -1) {
            firstTodayRow = row + startRow;
            break; // Stop searching after finding the first occurrence
          }
        }
      }
    }
    if (firstTodayRow !== -1) break;
  }

  if (firstTodayRow === -1) throw new Error(`Today's date (${todayDateString}) not found in summary sheet`);

  // Check for 'FIRME' above the first occurrence
  const searchRange = 5; // Number of rows to search upwards for 'FIRME'
  for (let row = firstTodayRow - startRow; row >= 0 && row >= firstTodayRow - startRow - searchRange; row--) {
    const cellValue = values[row][1];
    if (cellValue && cellValue.toString().trim().toUpperCase() === 'FIRME') {
      Logger.log(`Found 'FIRME' in cell [${row + startRow}, 2]`);
      return row + startRow + 1; // The row immediately below "FIRME"
    }
  }

  // Check for a second occurrence within the next 6 rows if 'FIRME' was not found above the first occurrence
  let secondTodayRow = -1;
  for (let row = firstTodayRow - startRow + 1; row <= firstTodayRow - startRow + 6 && row < values.length; row++) {
    for (let col = 0; col < values[row].length; col++) {
      if (values[row][col] instanceof Date) {
        const cellDateString = Utilities.formatDate(values[row][col], Session.getScriptTimeZone(), "dd/MM/yyyy");
        if (cellDateString === todayDateString) {
          secondTodayRow = row + startRow;
          break;
        }
      }
    }
    if (secondTodayRow !== -1) break;
  }

  if (secondTodayRow !== -1) {
    for (let row = secondTodayRow - startRow; row >= 0 && row >= secondTodayRow - startRow - searchRange; row--) {
      const cellValue = values[row][1];
      Logger.log(`Checking cell [${row + startRow}, 2] for 'FIRME': ${cellValue}`);
      if (cellValue && cellValue.toString().trim().toUpperCase() === 'FIRME') {
        Logger.log(`Found 'FIRME' in cell [${row + startRow}, 2]`);
        return row + startRow + 1; // The row immediately below "FIRME"
      }
    }
  }

  throw new Error(`'FIRME' not found above today's date in summary sheet`);
}



function getCurrentWeek() {
  const storedWeek = PropertiesService.getScriptProperties().getProperty('currentWeek');
  return storedWeek ? parseInt(storedWeek, 10) : configData.currentWeek;
}


function setCurrentWeek(newWeek) {
  configData.currentWeek = newWeek;
  PropertiesService.getScriptProperties().setProperty('currentWeek', newWeek);
}


function getSummarySheetStartRow(weekIndex) {
  const storedRows = JSON.parse(PropertiesService.getScriptProperties().getProperty('summarySheetStartRows'));
  const mergedRows = [...configData.summarySheetStartRows, ...(storedRows ? storedRows.slice(configData.summarySheetStartRows.length) : [])];
  return mergedRows[weekIndex] !== undefined ? mergedRows[weekIndex] : 158; // Default to the first row if index is out of bounds
}


function addNewWeekConfig(config) {
  try {
    const { newWeekIndex, weekRows, summarySheetStartRow } = config;
    const sourceSheetsConfig = getSourceSheetsConfigFromScript();
    const storedRows = JSON.parse(PropertiesService.getScriptProperties().getProperty('summarySheetStartRows')) || [];

    sourceSheetsConfig.forEach(sheet => {
      const weekRow = weekRows[sheet.sheetName];
      if (weekRow) {
        if (!sheet.weekRows) {
          sheet.weekRows = [];
        }
        // Ensure that the weekRows array can accommodate the new week index
        if (newWeekIndex - 22 >= sheet.weekRows.length) {
          sheet.weekRows.length = newWeekIndex - 22 + 1;
        }
        sheet.weekRows[newWeekIndex - 22] = parseInt(weekRow, 10);
      }
    });

    // Ensure the storedRows array can accommodate the new week index
    if (newWeekIndex - 22 >= storedRows.length) {
      storedRows.length = newWeekIndex - 22 + 1;
    }
    
    storedRows[newWeekIndex - 22] = parseInt(summarySheetStartRow, 10);

    PropertiesService.getScriptProperties().setProperty('summarySheetStartRows', JSON.stringify(storedRows));
    PropertiesService.getScriptProperties().setProperty('sourceSheetsConfig', JSON.stringify(sourceSheetsConfig));

    Logger.log(`Added new week: ${newWeekIndex}`);
    Logger.log(`Week rows: ${JSON.stringify(weekRows)}`);
    Logger.log(`Summary sheet start row: ${summarySheetStartRow}`);
  } catch (e) {
    Logger.log(`Error in addNewWeekConfig: ${e.message}`);
  }
}
