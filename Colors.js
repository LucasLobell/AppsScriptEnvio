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
  var summarySheet = ss.getSheetByName('2024'); // Correct sheet name

  // Use existing variable for data starting row
  if (!dataStartRow) {
    // If summarySheetStartRow is not defined globally, you might need to define it here
    // For example, find the starting row based on today's date
    var todayDateString = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var dataStartRow = findStartingRowInSummarySheet(summarySheet, todayDateString);
  }

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

  // Map to store the last mention and future mentions of each company
  var companyData = {};

  // First, process all dates and implantadores to build the data map
  for (var i = 0; i < dateValues.length; i++) {
    var dateCell = dateValues[i][0];
    if (dateCell instanceof Date) {
      var date = new Date(dateCell);
      date.setHours(0, 0, 0, 0); // Normalize date

      var dateIsPast = date.getTime() < currentDate.getTime();
      var dateIsFutureOrToday = date.getTime() >= currentDate.getTime();

      for (var j = 0; j < numImplantadorColumns; j++) {
        var companyCell = implantadorValues[i][j];
        var companies = companyCell ? companyCell.split('/') : [];

        companies.forEach(function(company) {
          company = company.trim().toUpperCase(); // Standardize company name

          if (company) { // Ensure company is not empty after trimming
            if (!companyData[company]) {
              companyData[company] = {
                lastMentionDate: null,
                lastMentionRow: null,
                lastMentionColumn: null,
                futureMentions: []
              };
            }

            if (dateIsPast) {
              // Past mention
              if (!companyData[company].lastMentionDate || date.getTime() > companyData[company].lastMentionDate.getTime()) {
                companyData[company].lastMentionDate = date;
                companyData[company].lastMentionRow = dataStartRow + i;
                companyData[company].lastMentionColumn = implantadorStartColumn + j;
              }
            } else if (dateIsFutureOrToday) {
              // Today or future mention
              companyData[company].futureMentions.push(date);
            }
          }
        });
      }
    }
  }

  // Now, go through the company data and update cell colors
  for (var company in companyData) {
    var data = companyData[company];

    if (data.lastMentionDate) {
      var daysSinceLastMention = (currentDate.getTime() - data.lastMentionDate.getTime()) / (1000 * 60 * 60 * 24);

      if (daysSinceLastMention >= 2 && data.futureMentions.length === 0) {
        // Color the cell red
        var cell = summarySheet.getRange(data.lastMentionRow, data.lastMentionColumn);
        cell.setBackground('red');
      }
      // Else, do not modify the cell's background color
    }
  }
}
