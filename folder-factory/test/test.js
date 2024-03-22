/* test functions to use from the Apps Script editor */

function insert() {
  let c = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSelection().getCurrentCell();
  Logger.log(c.getA1Notation());
  FolderApp.insertHeaderAt(c);
}
