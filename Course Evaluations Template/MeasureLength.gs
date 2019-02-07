/**
* Measure length of the template sheet to determine how many line breaks are necessary for page break
* @return {number} rowsLength The number of rows with data
*/
function measureLengthOfTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template");
  var rowsLength = ss.getLastRow();
  //Logger.log("The template is " + rowsLength + " rows long.");
  return rowsLength;
}
