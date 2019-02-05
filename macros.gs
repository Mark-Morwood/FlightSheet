function Clean() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getCurrentCell().setValue('GRAHAM_Alan');
  spreadsheet.getCurrentCell().offset(1, 0).activate();
};