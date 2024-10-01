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