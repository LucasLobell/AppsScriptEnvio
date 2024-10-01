function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Atualizar')
    .addItem('Atualizar Semana', 'openConfigSidebar')
    .addToUi();

    // Automatically open the sidebar when the spreadsheet is opened
    openConfigSidebar();
}

function openConfigSidebar() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('ConfigSidebar')
    .setTitle('Atualização da Planilha');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}



// Other existing functions (updateSheetConfig, addNewWeekConfig, logStoredProperties) go here...


function setCurrentWeek(currentWeek) {
  // Assuming setCurrentWeek is a valid function that updates a property or performs some action
  PropertiesService.getScriptProperties().setProperty('currentWeek', currentWeek);
}


function updateSheetConfig(currentWeek) {
  setCurrentWeek(currentWeek);
  const currentWeekIndex = Math.max(currentWeek - 22, 0); // Ensure the index is never less than 0


  // Log the update for verification
  Logger.log(`Updated current week to: ${currentWeek}`);
  Logger.log(`Updated current week index to: ${currentWeekIndex}`);
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
        sheet.weekRows[newWeekIndex - 22] = parseInt(weekRow, 10);
      }
    });


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


function logStoredProperties() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    Logger.log('Source Sheets Config: ' + (scriptProperties.getProperty('sourceSheetsConfig') || 'null'));
    Logger.log('Current Week: ' + (scriptProperties.getProperty('currentWeek') || 'null'));
    Logger.log('Summary Sheet Start Rows: ' + (scriptProperties.getProperty('summarySheetStartRows') || 'null'));
  } catch (e) {
    Logger.log(`Error in logStoredProperties: ${e.message}`);
  }
}


// Add this function to retrieve the source sheets config
function getSourceSheetsConfigFromScript() {
  const storedConfig = PropertiesService.getScriptProperties().getProperty('sourceSheetsConfig');
  const dynamicConfig = storedConfig ? JSON.parse(storedConfig) : [];
 
  // Merge static and dynamic configurations
  return configData.sourceSheetsConfig.map(sheet => {
    const dynamicSheet = dynamicConfig.find(dSheet => dSheet.sheetName === sheet.sheetName);
    if (dynamicSheet) {
      sheet.weekRows = [...sheet.weekRows, ...dynamicSheet.weekRows.slice(sheet.weekRows.length)];
    }
    return sheet;
  });
}

