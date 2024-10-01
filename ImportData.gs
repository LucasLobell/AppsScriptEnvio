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

  const days = ['seg.', 'ter.', 'qua.', 'qui.', 'sex.'];
  const dayColumnRanges = ['B:G', 'H:M', 'N:S', 'T:Y', 'Z:AF'];

  days.forEach((day, index) => {
    if (selectedDay !== 'all' && day !== selectedDay) return;

    const dayColumnRange = dayColumnRanges[index];
    const morningRange = `${dayColumnRange.split(':')[0]}${weekRowNumber + 2}:${dayColumnRange.split(':')[1]}${weekRowNumber + 6}`;
    const afternoonRange = `${dayColumnRange.split(':')[0]}${weekRowNumber + 8}:${dayColumnRange.split(':')[1]}${weekRowNumber + 12}`;
    const targetCell = `${sheetConfig.targetColumn}${summarySheetStartRow + index}`; // Adjust row index as necessary

    Logger.log(`Processing ${day} for sheet ${sheetConfig.sheetName}...`);
    Logger.log(`Morning Range: ${morningRange}`);
    Logger.log(`Afternoon Range: ${afternoonRange}`);

    const morningActivities = sourceSheet.getRange(morningRange).getValues().flat();
    const afternoonActivities = sourceSheet.getRange(afternoonRange).getValues().flat();
    const morningFontColors = sourceSheet.getRange(morningRange).getFontColors().flat();
    const afternoonFontColors = sourceSheet.getRange(afternoonRange).getFontColors().flat();
    const morningBackgroundColors = sourceSheet.getRange(morningRange).getBackgrounds().flat();
    const afternoonBackgroundColors = sourceSheet.getRange(afternoonRange).getBackgrounds().flat();

    const activities = morningActivities.concat(afternoonActivities);
    const fontColors = morningFontColors.concat(afternoonFontColors);
    const backgroundColors = morningBackgroundColors.concat(afternoonBackgroundColors);

    const foundWords = processActivities(activities, fontColors, backgroundColors, allPossibleWords, possiblePrefixes);
    const combinedWordsString = Array.from(foundWords).join('/');

    Logger.log(`Found words for ${day}: ${combinedWordsString}`);

    updateSummarySheet(summarySheet, targetCell, combinedWordsString);
  });
}

function updateSummarySheet(summarySheet, targetCell, combinedWordsString) {
  const targetRange = summarySheet.getRange(targetCell);
  Logger.log(`Updating target cell ${targetCell} with value: ${combinedWordsString}`);

  targetRange.setValue(combinedWordsString);
  if (combinedWordsString) {
    targetRange.setBackground('#ffc000'); // Set background to yellow if words are found
  } else {
    targetRange.setBackground('#00b050'); // Reset background color if no words found
  }

  Logger.log(`Set value for ${targetCell}: ${combinedWordsString}`);
}


function processActivities(activities, fontColors, backgroundColors, allPossibleWords, possiblePrefixes) {
  const foundWords = new Set();
  activities.forEach((activity, index) => {
    if (activity && typeof activity === 'string' && activity.split(/[\s\-./]+/).length > 1 &&
        fontColors[index] !== '#ff0000' && backgroundColors[index] !== '#ff0000') {
      const words = activity.trim().split(/[\s\-./]+/);
      let prefixFound = false;
      for (let idx = 0; idx < words.length; idx++) {
        const word = words[idx];
        if (possiblePrefixes.includes(word)) {
          prefixFound = true;
        } else if (prefixFound) {
          const nextWords = words.slice(idx).join(' ');
          const matchedWord = allPossibleWords.find(possibleWord => nextWords.toLowerCase().includes(possibleWord.toLowerCase()));
          if (matchedWord) {
            foundWords.add(matchedWord.toUpperCase());
            prefixFound = false;
          }
        }
      }
    }
  });
  return foundWords;
}
