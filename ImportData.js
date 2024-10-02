function importDataFromSourceSheets(selectedSheet = 'all', selectedDay = 'all') {
  logStoredProperties(); // Log properties to ensure correct data is being retrieved

  const sourceSheetsConfig = getSourceSheetsConfigFromScript();
  const possibleWords = getPossibleWords();
  const allPossibleWords = getAllPossibleWords(possibleWords);
  const possiblePrefixes = getPossiblePrefixes();

  const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = targetSpreadsheet.getSheetByName('2024');

  const today = new Date();
  const todayDateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yyyy");
  const summarySheetStartRow = findStartingRowInSummarySheet(summarySheet, todayDateString);

  Logger.log(`Selected Sheet: ${selectedSheet}`);
  Logger.log(`Selected Day: ${selectedDay}`);
  Logger.log(`Summary sheet starting row determined to be: ${summarySheetStartRow}`);

  sourceSheetsConfig.forEach(sheetConfig => {
    if (selectedSheet !== 'all' && sheetConfig.sheetName !== selectedSheet) return;

    try {
      const currentWeekRow = findCurrentWeekRow(sheetConfig);
      Logger.log(`Found current week row: ${currentWeekRow} for sheet ${sheetConfig.sheetName}`);

      processSheet(sheetConfig, summarySheet, allPossibleWords, possiblePrefixes, currentWeekRow, selectedDay, summarySheetStartRow, selectedSheet);
    } catch (e) {
      Logger.log(`Error processing sheet ${sheetConfig.sheetName}: ${e.toString()}`);
    }
  });
}

function processSheet(sheetConfig, summarySheet, allPossibleWords, possiblePrefixes, weekRowNumber, selectedDay, summarySheetStartRow, selectedSheet) {
  const sourceSpreadsheet = SpreadsheetApp.openByUrl(sheetConfig.url);
  const sourceSheet = sourceSpreadsheet.getSheetByName(sheetConfig.sheetName);

  const columnsPerDay = 6; // Number of columns for each day
  const days = ['seg.', 'ter.', 'qua.', 'qui.', 'sex.'];

  days.forEach((day, index) => {
    if (selectedDay !== 'all' && day !== selectedDay) return;

    const startColumn = (index * columnsPerDay) + 2; // Starting column index (adjust as per structure)
    const endColumn = startColumn + columnsPerDay - 1;
    
    // Batch retrieve the entire range in one call for both morning and afternoon
    const fullRange = sourceSheet.getRange(`${getColumnLetter(startColumn)}${weekRowNumber + 2}:${getColumnLetter(endColumn)}${weekRowNumber + 12}`);
    const fullValues = fullRange.getValues();
    const fullFontColors = fullRange.getFontColors();
    const fullBackgroundColors = fullRange.getBackgrounds();

    // Morning and afternoon slices from the full batch data
    const morningValues = fullValues.slice(0, 5).flat();
    const afternoonValues = fullValues.slice(6).flat();
    const morningFontColors = fullFontColors.slice(0, 5).flat();
    const afternoonFontColors = fullFontColors.slice(6).flat();
    const morningBackgroundColors = fullBackgroundColors.slice(0, 5).flat();
    const afternoonBackgroundColors = fullBackgroundColors.slice(6).flat();

    const activities = [...morningValues, ...afternoonValues];
    const fontColors = [...morningFontColors, ...afternoonFontColors];
    const backgroundColors = [...morningBackgroundColors, ...afternoonBackgroundColors];

    const foundWords = processActivities(activities, fontColors, backgroundColors, allPossibleWords, possiblePrefixes);
    const combinedWordsString = Array.from(foundWords).join('/');

    Logger.log(`Found words for ${day}: ${combinedWordsString}`);

    // Calculate the target cell in the summary sheet
    const targetCell = `${sheetConfig.targetColumn}${summarySheetStartRow + index}`;

    updateSummarySheet(summarySheet, targetCell, combinedWordsString);
  });
}

function updateSummarySheet(summarySheet, targetCell, combinedWordsString) {
  const targetRange = summarySheet.getRange(targetCell);
  targetRange.setValues([[combinedWordsString]]); // Use setValues even for single cells
  targetRange.setBackground(combinedWordsString ? '#ffc000' : '#00b050');
  Logger.log(`Set value for ${targetCell}: ${combinedWordsString}`);
}

function processActivities(activities, fontColors, backgroundColors, allPossibleWords, possiblePrefixes) {
  const foundWords = new Set();

  activities.forEach((activity, index) => {
    if (activity && fontColors[index] !== '#ff0000' && backgroundColors[index] !== '#ff0000') {
      const words = activity.trim().split(/[\s\-./]+/);
      let prefixFound = false;

      words.forEach(word => {
        if (possiblePrefixes.includes(word)) {
          prefixFound = true;
        } else if (prefixFound && allPossibleWords.some(possibleWord => word.toLowerCase().includes(possibleWord.toLowerCase()))) {
          foundWords.add(word.toUpperCase());
          prefixFound = false;
        }
      });
    }
  });

  return foundWords;
}

// Helper function to convert column index to letter (A = 1, B = 2, etc.)
function getColumnLetter(columnIndex) {
  let letter = '';
  let temp = columnIndex;
  while (temp > 0) {
    const mod = (temp - 1) % 26;
    letter = String.fromCharCode(mod + 65) + letter;
    temp = Math.floor((temp - mod) / 26);
  }
  return letter;
}