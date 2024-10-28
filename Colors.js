/** @OnlyCurrentDoc */


function Green() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().activate();
  spreadsheet.getActiveRangeList().setBackground('#00b050');
};

function Yellow() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().activate();
  spreadsheet.getActiveRangeList().setBackground('#ffc000');
};

function Red() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().activate();
  spreadsheet.getActiveRangeList().setBackground('#ff0000');
};

function White() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().activate();
  spreadsheet.getActiveRangeList().setBackground('#ffffff');
};

function Clean() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().activate();
  spreadsheet.getActiveRangeList().setValue('');
  spreadsheet.getActiveRangeList().setBackground('#00b050');
};

function updateSummarySheetCellColors() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName('2024'); // Adjust the sheet name if necessary

  // Set dataStartRow to the first row of data (after headers)
  var dataStartRow = 2; // Adjust if your data starts at a different row

  // Get the last row with data
  var lastRow = summarySheet.getLastRow();

  // Determine the columns for dates and implantadores
  var dateColumn = 2; // Column B
  var implantadorStartColumn = 3; // Column C
  var implantadorEndColumn = 9; // Column I
  var numImplantadorColumns = implantadorEndColumn - implantadorStartColumn + 1;

  // Get the range of dates and implantadores
  var numRows = lastRow - dataStartRow + 1;
  var dateRange = summarySheet.getRange(dataStartRow, dateColumn, numRows, 1);
  var implantadorRange = summarySheet.getRange(dataStartRow, implantadorStartColumn, numRows, numImplantadorColumns);

  var dateValues = dateRange.getValues();
  var implantadorValues = implantadorRange.getValues();

  var currentDate = new Date();
  currentDate.setHours(0, 0, 0, 0); // Normalize to midnight

  // Map to store mentions of each company
  var companyData = {};

  // First, process all dates and implantadores to build the data map
  for (var i = 0; i < dateValues.length; i++) {
    var dateCell = dateValues[i][0];
    var date;

    if (dateCell instanceof Date) {
      date = new Date(dateCell.getFullYear(), dateCell.getMonth(), dateCell.getDate());
    } else {
      // Try parsing the date if it's not a Date object
      date = new Date(dateCell);
    }

    if (date instanceof Date && !isNaN(date)) {
      date.setHours(0, 0, 0, 0); // Normalize date

      for (var j = 0; j < numImplantadorColumns; j++) {
        var companyCell = implantadorValues[i][j];
        var companies = companyCell ? companyCell.split('/') : [];

        companies.forEach(function(company) {
          company = company.trim().toUpperCase();

          if (company) {
            if (!companyData[company]) {
              companyData[company] = {
                mentions: [], // array of {date, rowIndex, weekNumber}
                lastMentionDateBeforeToday: null,
                lastMentionWeekNumber: null
              };
            }

            // Store the mention
            var rowIndex = dataStartRow + i;
            var columnIndex = implantadorStartColumn + j;
            var weekNumber = getWeekNumber(date);

            companyData[company].mentions.push({
              date: date,
              rowIndex: rowIndex,
              columnIndex: columnIndex,
              weekNumber: weekNumber
            });

            if (date.getTime() < currentDate.getTime()) {
              // Date is before today
              if (!companyData[company].lastMentionDateBeforeToday || date.getTime() > companyData[company].lastMentionDateBeforeToday.getTime()) {
                companyData[company].lastMentionDateBeforeToday = date;
                companyData[company].lastMentionWeekNumber = weekNumber;
              }
            }
          }
        });
      }
    } else {
      Logger.log('Invalid date at row ' + (dataStartRow + i) + ': ' + dateCell);
    }
  }

  // Now, go through the company data and update cell colors
  for (var company in companyData) {
    var data = companyData[company];

    if (data.lastMentionDateBeforeToday) {
      var daysSinceLastMention = calculateBusinessDaysBetween(data.lastMentionDateBeforeToday, currentDate);

      Logger.log('Company: ' + company);
      Logger.log('Last Mention Date Before Today: ' + data.lastMentionDateBeforeToday);
      Logger.log('Business Days Since Last Mention: ' + daysSinceLastMention);

      if (daysSinceLastMention >= 2) {
        // Color all mention cells during the week of the last occurrence
        data.mentions.forEach(function(mention) {
          if (mention.weekNumber === data.lastMentionWeekNumber) {
            var cell = summarySheet.getRange(mention.rowIndex, mention.columnIndex);
            cell.setBackground('red');
          }
        });
      }
    }
  }
}

// Helper function to calculate business days between two dates (excluding weekends)
function calculateBusinessDaysBetween(startDate, endDate) {
  var count = 0;
  var currentDate = new Date(startDate.getTime());
  currentDate.setHours(0, 0, 0, 0);

  while (currentDate < endDate) {
    var dayOfWeek = currentDate.getDay();
    if (dayOfWeek !== 0 && dayOfWeek !== 6) { // Exclude Sundays (0) and Saturdays (6)
      count++;
    }
    currentDate.setDate(currentDate.getDate() + 1);
  }

  return count;
}

// Helper function to get the week number of a date
function getWeekNumber(date) {
  var oneJan = new Date(date.getFullYear(), 0, 1);
  var dayOfYear = Math.floor((date - oneJan) / (24 * 60 * 60 * 1000)) + 1;
  return Math.ceil(dayOfYear / 7);
}
