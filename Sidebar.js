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

const scriptProperties = PropertiesService.getScriptProperties();

function getSourceSheetsConfigFromScript() {
  const storedConfig = JSON.parse(scriptProperties.getProperty('sourceSheetsConfig')) || [];
  return configData.sourceSheetsConfig.map(sheet => ({
    ...sheet,
    weekRows: [...sheet.weekRows, ...(storedConfig.find(dSheet => dSheet.sheetName === sheet.sheetName)?.weekRows.slice(sheet.weekRows.length) || [])]
  }));
}

