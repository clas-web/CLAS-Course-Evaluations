//************************************************************************************************************************************************************************
//Create selector
function showPicker() {
  var html = HtmlService.createHtmlOutputFromFile('Picker.html')
  .setWidth(600)
  .setHeight(425)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Folder');
}
//**********************************************************************************************************************************************************************************************
//Authorize script
function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}
//**********************************************************************************************************************************************************************************************
//import datasheet(s) using the Picker
/**
* @NotOnlyCurrentDoc
*/
function uploadSheet(id){  
  var currentDate = new Date();
  var file = DriveApp.getFileById(id);
  Logger.log(file.getMimeType());
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Is the attachment a CSV file?  
  if (file.getMimeType() == "text/csv") {    
    Logger.log("CSV sheet");
    var htmlOutput = HtmlService.createHtmlOutput('<title>Please wait...</title> Importing CSV...');
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Saving...');    
    try{
      var newDataSheet = sheet.insertSheet('datasheet-'+currentDate);
    }catch(e){
      var newDataSheet = sheet.insertSheet('datasheet-');
    }
    var csvData = Utilities.parseCsv(file.getBlob().getDataAsString(), ",");
    newDataSheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);    
    
    // Is the attachment a Google Sheet?
  } else if(file.getMimeType() == "application/vnd.google-apps.spreadsheet"){      
    Logger.log("Google Sheet");
    //import all sheets of importing spreadsheet
    var htmlOutput = HtmlService.createHtmlOutput('<title>Please wait...</title> Importing Google Sheet...');
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Saving...');
    var SSSheets = SpreadsheetApp.openByUrl(file.getUrl()).getSheets();
    for (var n=0; n < SSSheets.length; n++){
      // Get full range of data
      var sheetRange = SSSheets[n].getDataRange();
      // get the data values in range
      var sheetData = sheetRange.getValues();
      try{
        var newDataSheet = sheet.insertSheet(file.getName()+"_"+SSSheets[n].getName()+"_"+currentDate);
      }catch(e){
        var newDataSheet = sheet.insertSheet(file.getName()+"_"+SSSheets[n].getName()+"_");
      }
      newDataSheet.getRange(1, 1, SSSheets[n].getLastRow(), SSSheets[n].getLastColumn()).setValues(sheetData);                        
    }
    
    //Is the attachment an Excel file?
  } else if (file.getMimeType() == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"){
    Logger.log("Excel sheet");
    var htmlOutput = HtmlService.createHtmlOutput('<title>Please wait...</title> Importing Excel Sheet...');
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Saving...');
    //convert to Google Sheet
    var convertedFile = {
      title: file.getName()+"_"+currentDate
    };
    convertedFile = Drive.Files.insert(convertedFile,file, {
      convert:true
    });
    //import Google Sheet
    var SSSheets = SpreadsheetApp.openById(convertedFile.id).getSheets();
    for (var n=0; n < SSSheets.length; n++){
      // Get full range of data
      var sheetRange = SSSheets[n].getDataRange();
      // get the data values in range
      var sheetData = sheetRange.getValues();
      var newDataSheet = sheet.insertSheet(convertedFile.title+"_"+SSSheets[n].getName()+"");
      newDataSheet.getRange(1, 1, SSSheets[n].getLastRow(), SSSheets[n].getLastColumn()).setValues(sheetData);                        
    }      
    
    //Is the file not a CSV or Google Sheet? Then don't import
    //Excel? May need to enable Google Drive API first: https://stackoverflow.com/questions/11681873/converting-xls-to-google-spreadsheet-in-google-apps-script
    //https://developers.google.com/apps-script/guides/services/advanced
  }else {
    Logger.log('Not a CSV or Google Sheet');
    SpreadsheetApp.getActiveSpreadsheet().toast("No datasheets imported", "Not a Spreadsheet");
  }
  runScript();
}
//************************************************************************************************************************************************************************

