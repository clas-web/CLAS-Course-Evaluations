//https://stackoverflow.com/questions/32527682/auto-close-modal-dialog-after-server-code-is-done-close-dialog-in-google-spre
//https://stackoverflow.com/questions/11681873/converting-xls-to-google-spreadsheet-in-google-apps-script
//https://developers.google.com/apps-script/guides/html/
//https://developers.google.com/apps-script/guides/html/communication
//https://ctrlq.org/code/20239-copy-google-spreadsheets
//https://github.com/odeke-em/drive/wiki/List-of-MIME-type-short-keys
//https://stackoverflow.com/questions/33701881/using-an-html-drop-down-menu-with-google-apps-script-on-google-sheets
//************************************************************************************************************************************************************************
//Add functions to spreadsheet menu
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Run Script', functionName: 'runScript'},
    //{name: 'Print Output as a PDF', functionName: 'printPDF'}, 
    //{name: 'Force Quit', functionName: 'breakOperation'},   
  ];
    spreadsheet.addMenu('Generate Sheets', menuItems);
    }
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
    //Run the script
    function runScript(){  
    
    //Declare variables
    var ss = SpreadsheetApp.getActiveSpreadsheet(); //Course Evaluations spreadsheet
    var allSheets = ss.getSheets(); //collect sheet objects 
    var indexOfSheets = [];   //array of sheet names
  
  //create semester list
  var semesters = [/*"Fall-2018",*/"Spring-2018","Fall-2017","Spring-2017","Fall-2016"/*,"Spring-2016"*/];
  var semesterString = "\""+semesters.join(" ")+"\"";
  var semesterCheckArray = "<strong> Select the semester:</strong><br> <form> <select id='sem-dropdown'>";  
  
  //create semester drop down menu
  for (var a=0;a<semesters.length;a++){
    semesterCheckArray += '<option value="'+semesters[a]+'">'+semesters[a]+'</option>';        
  }  
  semesterCheckArray += '</select> </form> <p id="semester"></p>';
  
  //get dept codes from my CLAS OAT sheet
  var dropdownMenu = "<h1>Student Evaluations Report</h1> <strong>Select your department:</strong> <br><form> <select id='dpt-dropdown'>";
  var deptCodes = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1zNrJRaS1eMSB6vTCq_FB78Zrphq8mfzxSEbGaBsa3Wk/edit#gid=997728017')
  .getSheetByName('Departments').getRange('B2:B').getValues();
  //create dropdown menu for depts
  for (var m=0; m < deptCodes.length; m++){
    dropdownMenu += "<option value="+deptCodes[m]+">"+deptCodes[m]+"</option>";    
  }
  dropdownMenu += "</select> </form> <p id='dept'></p>";
  
  //Create HTML output for script
  var message = "<br> <br> <br> <br> <input type='button' value='Import a datasheet...' onClick='google.script.run.showPicker();' style='float: right;' />";
  var sheetChoices = "<h1>Select your datasheet</h1> Select the sheet below that you want to be your source datasheet. If your datasheet has more than " +
    "250 entries, the operation may time out.<br> <body> <script> google.script.run.withSuccessHandler(google.script.host.close();) </script>";      
  var currentSheet = "";
  
  // create array of sheet names
  allSheets.forEach(function(sheet){
    indexOfSheets.push([sheet.getSheetName()]);
  });
  
  Logger.log('Length of indexOfSheets:' + indexOfSheets.length);
  
  //add sheet names as choices
  for (var n=0; n < indexOfSheets.length; n++){           
    currentSheet = "\""+indexOfSheets[n]+"\"";
    Logger.log("currentSheet: "+currentSheet);    
    sheetChoices += "<br> <input type='button' value='"+indexOfSheets[n]+"' onClick='cloneSheets("+currentSheet+")' />" + "<br>";    
  }
  
  
  //HTML functions for passing dropdown menu choices to Google Script
  var script = "<script> function cloneSheets(currentSheet, semester, semesterString) { var deptChoice = document.getElementById('dpt-dropdown').value;" +
    "document.getElementById('dept').innerHTML = deptChoice; var semesterChoice = document.getElementById('sem-dropdown').value; " +
      "document.getElementById('semester').innerHTML = semesterChoice; var semesterString = "+semesterString+";" +
        "google.script.run.cloneGoogleSheet(currentSheet,deptChoice, semesterChoice, semesterString); google.script.host.close(); } </script>";
  
  //Display HTML output
  var htmlApp = HtmlService
  .createHtmlOutput(dropdownMenu + semesterCheckArray + sheetChoices + message + script + " </body>")
  .setWidth(400)
  .setHeight(indexOfSheets.length*125);
  
  SpreadsheetApp.getUi().showModalDialog(htmlApp, "Run Script");     
  var output = HtmlService.createTemplate(htmlApp);
  //Receive HTML output as a formatted HTML file for troubleshooting
  //Logger.log(output.getCode());
}
//************************************************************************************************************************************************************************    
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

function getByName(colName, row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("datasheet");
  var data = sheet.getDataRange().getValues();
  var col = data[0].indexOf(colName);
  if (col != -1) {
    return data[row-1][col];
  }
}

function getByName2(data, colName, row) {
  var col = data[0].indexOf(colName);
  if (col != -1) {
    return data[row-1][col];
  }
}


//************************************************************************************************************************************************************************
//Update the output by the entry row on datasheet
function updateDataTemplate(datasheet, target, current, semester, semesters){

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("datasheet");
  var data = sheet.getDataRange().getValues();
  
  semesters = semesters.split(" ");
  target.getRange('C1').setValue(getByName2(data,"Instructor", current)) //Instructor
  target.getRange('C2').setValue(getByName2(data,"COURSE",current)); //Class  
  target.getRange('C3').setValue(getByName2(data,"Term_Year",current));  //Term
  target.getRange('F2').setValue(getByName2(data,"title",current));   //Course Name
  target.getRange('F3').setValue(getByName2(data,"XLST",current));   //Course Name
  
  /*CLASS */
  target.getRange('B33').setValue(getByName2(data,"CGPA",current));
  target.getRange('C33').setValue(getByName2(data,"DFW",current));
  target.getRange('D33').setValue(getByName2(data,"PGPA",current));
  /*COURSE */
  target.getRange('B34').setValue(getByName2(data,"Common_CGPA",current));
  target.getRange('C34').setValue(getByName2(data,"Common_DFW",current));
  target.getRange('D34').setValue(getByName2(data,"Common_PGPA",current));
  /* GROUP */ 
  target.getRange('B35').setValue(getByName2(data,"Group_CGPA",current));
  target.getRange('C35').setValue(getByName2(data,"Group_DFW",current));
  target.getRange('D35').setValue(getByName2(data,"Group_CGPA",current));
  /* DEPT */
  target.getRange('B36').setValue(getByName2(data,"Dept_CGPA",current));
  target.getRange('C36').setValue(getByName2(data,"Dept_DFW",current));
  target.getRange('D36').setValue(getByName2(data,"Dept_CGPA",current));
  
  //Grade Distribution
  target.getRange('B40').setValue(getByName2(data,"percent_A",current));
  target.getRange('C40').setValue(getByName2(data,"percent_B",current));
  target.getRange('D40').setValue(getByName2(data,"percent_C",current));
  target.getRange('E40').setValue(getByName2(data,"percent_D",current));
  target.getRange('F40').setValue(getByName2(data,"percent_F",current));
  target.getRange('G40').setValue(getByName2(data,"percent_W",current));
  target.getRange('H40').setValue(getByName2(data,"percent_X",current));
  
  target.getRange('B41').setValue(getByName2(data,"Common_percent_A",current));
  target.getRange('C41').setValue(getByName2(data,"Common_percent_B",current));
  target.getRange('D41').setValue(getByName2(data,"Common_percent_C",current));
  target.getRange('E41').setValue(getByName2(data,"Common_percent_D",current));
  target.getRange('F41').setValue(getByName2(data,"Common_percent_F",current));
  target.getRange('G41').setValue(getByName2(data,"Common_percent_W",current));
  target.getRange('H41').setValue(getByName2(data,"Common_percent_X",current));
  
  target.getRange('B42').setValue(getByName2(data,"Group_percent_A",current));
  target.getRange('C42').setValue(getByName2(data,"Group_percent_B",current));
  target.getRange('D42').setValue(getByName2(data,"Group_percent_C",current));
  target.getRange('E42').setValue(getByName2(data,"Group_percent_D",current));
  target.getRange('F42').setValue(getByName2(data,"Group_percent_F",current));
  target.getRange('G42').setValue(getByName2(data,"Group_percent_W",current));
  target.getRange('H42').setValue(getByName2(data,"Group_percent_X",current));
  
  target.getRange('B43').setValue(getByName2(data,"Dept_percent_A",current));
  target.getRange('C43').setValue(getByName2(data,"Dept_percent_B",current));
  target.getRange('D43').setValue(getByName2(data,"Dept_percent_C",current));
  target.getRange('E43').setValue(getByName2(data,"Dept_percent_D",current));
  target.getRange('F43').setValue(getByName2(data,"Dept_percent_F",current));
  target.getRange('G43').setValue(getByName2(data,"Dept_percent_W",current));
  target.getRange('H43').setValue(getByName2(data,"Dept_percent_X",current));
  
  //Student Evaluation Information
  target.getRange('E7').setValue(getByName2(data,"responseInSection",current));
  target.getRange('E8').setValue(getByName2(data,"RespN",current));
  target.getRange('E9').setValue(getByName2(data,"Group_resp_sections",current));
  target.getRange('E10').setValue(getByName2(data,"Dept_resp_sections",current));
  target.getRange('E11').setValue("N/A");
  target.getRange('E12').setValue("N/A");
  target.getRange('E13').setValue(getByName2(data,"Group_sections",current));
  target.getRange('E14').setValue(getByName2(data,"Dept_sections",current));
  
  //Determine which semester is listed
//  for (var x = 0; x < semesters.length; x++){
//    if (semester == semesters[x]){
      
      //Evaluation Questions
//      target.getRange(18,8-x).setValue("='"+datasheet+"'!Q"+current).setFontWeight("bold").setHorizontalAlignment("center");
//      target.getRange(19,8-x).setValue("='"+datasheet+"'!R"+current).setFontWeight("bold").setHorizontalAlignment("center");
//      target.getRange(20,8-x).setValue("='"+datasheet+"'!S"+current).setFontWeight("bold").setHorizontalAlignment("center");
//      target.getRange(21,8-x).setValue("='"+datasheet+"'!T"+current).setFontWeight("bold").setHorizontalAlignment("center");
//      target.getRange(22,8-x).setValue("='"+datasheet+"'!U"+current).setFontWeight("bold").setHorizontalAlignment("center");
//      target.getRange(23,8-x).setValue("='"+datasheet+"'!V"+current).setFontWeight("bold").setHorizontalAlignment("center");
//      
//    }
//  }
  
  target.getRange('H18').setValue(getByName2(data,"Q1A",current));
  target.getRange('H19').setValue(getByName2(data,"Q2A",current));
  target.getRange('H20').setValue(getByName2(data,"Q3A",current));
  target.getRange('H21').setValue(getByName2(data,"Q4A",current));
  target.getRange('H22').setValue(getByName2(data,"Q5A",current));
  target.getRange('H23').setValue(getByName2(data,"Q6A",current));
  target.getRange('H24').setValue(getByName2(data,"Q7A",current));
  target.getRange('H25').setValue(getByName2(data,"Q8A",current));  
  target.getRange('H26').setValue(getByName2(data,"Q9A",current));
  target.getRange('H27').setValue(getByName2(data,"Q10A",current));
  target.getRange('H28').setValue(getByName2(data,"Q11A",current));  
  target.getRange('H29').setValue(getByName2(data,"Q12A",current));
  
  target.getRange('I18').setValue(getByName2(data,"Group_Q1A",current));
  target.getRange('I19').setValue(getByName2(data,"Group_Q2A",current));
  target.getRange('I20').setValue(getByName2(data,"Group_Q3A",current));
  target.getRange('I21').setValue(getByName2(data,"Group_Q4A",current));
  target.getRange('I22').setValue(getByName2(data,"Group_Q5A",current));
  target.getRange('I23').setValue(getByName2(data,"Group_Q6A",current));
  target.getRange('I24').setValue(getByName2(data,"Group_Q7A",current));
  target.getRange('I25').setValue(getByName2(data,"Group_Q8A",current)); 
  target.getRange('I26').setValue(getByName2(data,"Group_Q9A",current));
  target.getRange('I27').setValue(getByName2(data,"Group_Q10A",current));
  target.getRange('I28').setValue(getByName2(data,"Group_Q11A",current));
  target.getRange('I29').setValue(getByName2(data,"Group_Q12A",current));
  
  target.getRange('J18').setValue(getByName2(data,"Dept_Q1A",current));
  target.getRange('J19').setValue(getByName2(data,"Dept_Q2A",current));
  target.getRange('J20').setValue(getByName2(data,"Dept_Q3A",current));
  target.getRange('J21').setValue(getByName2(data,"Dept_Q4A",current));
  target.getRange('J22').setValue(getByName2(data,"Dept_Q5A",current));
  target.getRange('J23').setValue(getByName2(data,"Dept_Q6A",current)); 
  target.getRange('J24').setValue(getByName2(data,"Dept_Q7A",current)); 
  target.getRange('J25').setValue(getByName2(data,"Dept_Q8A",current)); 
  target.getRange('J26').setValue(getByName2(data,"Dept_Q9A",current)); 
  target.getRange('J27').setValue(getByName2(data,"Dept_Q10A",current)); 
  target.getRange('J28').setValue(getByName2(data,"Dept_Q11A",current)); 
  target.getRange('J29').setValue(getByName2(data,"Dept_Q12A",current)); 
  
  target.getRange('G18').setValue(getByName2(data,"F17_Q1A",current));
  target.getRange('G19').setValue(getByName2(data,"F17_Q2A",current));
  target.getRange('G20').setValue(getByName2(data,"F17_Q3A",current));
  target.getRange('G21').setValue(getByName2(data,"F17_Q4A",current));
  target.getRange('G22').setValue(getByName2(data,"F17_Q5A",current));
  target.getRange('G23').setValue(getByName2(data,"F17_Q6A",current));
  target.getRange('G24').setValue(getByName2(data,"F17_Q7A",current));
  target.getRange('G25').setValue(getByName2(data,"F17_Q8A",current));
  target.getRange('G26').setValue(getByName2(data,"F17_Q9A",current));
  target.getRange('G27').setValue(getByName2(data,"F17_Q10A",current));
  target.getRange('G28').setValue(getByName2(data,"F17_Q11A",current));
  target.getRange('G29').setValue(getByName2(data,"F17_Q12A",current));
  
}



//************************************************************************************************************************************************************************
function copyRowsToEnd(source, target, pageBreak)
{
  /* You must use the copyValuesToRange function so the data values are copied, not the formulas.  
  You then use copyFormatToRange to copy the range. lastEntryStart is the start of the new data, lastEntryEnd the end. You must set them because getLastRow changes between the function calls. */
  
  //Positions the active range so the new entry can be copied to the output  
  var rowsLength = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getLastRow();  
  var sourceVal = source.getRange(1,1,rowsLength+pageBreak,11);
  var lastEntryStart = target.getLastRow()+pageBreak; //The plus two adds a row between outputs for readability and then starts the next report on the next row  
  var lastEntryEnd = target.getLastRow()+rowsLength+pageBreak; //The 40 is the length of the template (38) plus 2 extra rows for readability, each entry report will be 40 rows total
  
  
  //Prevents function from getting stopped at 1000 rows by adding 50 new rows to the end when we're nearing the limit
  if(target.getLastRow()>900){    
    target.insertRowsAfter(target.getLastRow(),50);
    lastEntryStart = target.getLastRow()+pageBreak;
    lastEntryEnd = target.getLastRow()+rowsLength+pageBreak;
    //Finally copy values and formatting to the output sheet
    sourceVal.copyValuesToRange(target,1,11,lastEntryStart,lastEntryEnd);
    sourceVal.copyFormatToRange(target,1,11,lastEntryStart,lastEntryEnd);    
  } else {       
    sourceVal.copyValuesToRange(target,1,11,lastEntryStart,lastEntryEnd);
    sourceVal.copyFormatToRange(target,1,11,lastEntryStart,lastEntryEnd);
  }
}
//************************************************************************************************************************************************************************
function cloneGoogleSheet(datasheet, dept, semester, semesters) {
  
  Logger.log("datasheet: "+datasheet);
  Logger.log("dropdown menu: "+dept);
  Logger.log("semester: "+semester);
  Logger.log("semesters: "+semesters);
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var sheet = ss.getSheetByName('Template').copyTo(ss);
  //number of rows that will push the end of the report off the page for printing
  var rowsLength = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getLastRow();
  var fullPageLength = 55; //It takes 55 rows to make one page before data runs to the next page
  var pageBreak = fullPageLength - rowsLength;
  var current = 1;
  
  //The datasheet is whatever sheet has our data on it
  var numOfReports = ss.getSheetByName(datasheet).getRange('A2:A').getValues().filter(String).length; //number of reports to generate based on how many entries are on datasheet
  //This removes the input screen and adds a progress bar
  var quitOperation = "<br><br><br> <input type='button' value='Force Quit' onClick='google.script.run.breakOperation("+"\""+dept+"\""+","+"\""+semester+"\""+")' />";
  var htmlOutput = HtmlService.createHtmlOutput('<b>Please wait...</b> <progress value="1" max="'+numOfReports+'"></progress>' + quitOperation);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Saving...');
  Logger.log(numOfReports);
  
  /* Before cloning the Template sheet to the Output sheet, delete any previous copy of the Output sheet*/
  var old = ss.getSheetByName(dept + "-"+semester);
  if (old) ss.deleteSheet(old);
  var old2 = ss.getSheetByName("Running-Report-Function");
  if (old2) ss.deleteSheet(old2);
  
  //Create Output sheet
  target = ss.insertSheet();
  target.setName(dept + "-"+semester);  
  SpreadsheetApp.flush(); // Utilities.sleep(2000);
  
  //Creates a sheet that becomes the updated template, updates until all reports are run
  sheet.setName("Running-Report-Function");
  
  /* Make the new sheet active */
  ss.setActiveSheet(sheet);
  SpreadsheetApp.getActiveSheet();     
  
  //Iterate through all entries in datasheet
  for(current=2; current <= numOfReports+1; current++){
    
    Logger.log("We're on "+ (current-1) + " out of "+numOfReports);
    updateDataTemplate(datasheet, sheet, current, semester, semesters);
    copyRowsToEnd(sheet, target, pageBreak);
    Logger.log('Target last row: ' + target.getLastRow());
    htmlOutput = HtmlService.createHtmlOutput('<b>Please wait...</b> <progress value="'+current+'" max="'+(numOfReports+1)+'"></progress>' + quitOperation);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Saving...');
  }
  //****************************************************************************    
  //Delete blank rows caused by page break
  var blankRows = target.getRange(1, 1, pageBreak-1, 11).getDisplayValues();
  var lineCounter = 0;
  
  //First make sure the lines aren't blank
  for (var i = 0; i < blankRows.length; i++) {
    var blankLine = blankRows[i].join('');
    
    if(blankLine ==""){
      //Logger.log(blankLine);
      //Logger.log("Line "+i+" is blank");    
    } else {       
      Logger.log(blankLine);
      Logger.log("Line "+i+" is not blank");
      lineCounter++;
    }
  }
  //If the first couple of lines caused by pageBreak are empty, delete them
  if(lineCounter == 0){  
    Logger.log("First couple of rows are blank, delete first pageBreak since lines are empty");
    target.deleteRows(1, pageBreak-1);
  } else {   
    Logger.log(blankRows);
    Logger.log("Will not delete pageBreak since lines aren't empty, double check report page breaks...");
  }
  //****************************************************************************      
  //Delete function sheet, no longer needed since reports are saved to Output sheet
  var old2 = ss.getSheetByName("Running-Report-Function");
  if (old2) ss.deleteSheet(old2);
  
  ss.setActiveSheet(ss.getSheetByName(dept + "-"+semester));
  
  //Ask user if they want to print results
  output = HtmlService.createHtmlOutput('<script>google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(output, 'Complete');
  var results = Browser.msgBox(
    'Would you like these reports emailed to you? \\n' +
    '\\nYes - Email you & save to Google Drive' +
    '\\nNo - No email, just save to Google Drive' +
    '\\nCancel - Do not save or email report',
    Browser.Buttons.YES_NO_CANCEL);
  
  // Process the user's response.
  if (results == 'yes' || results == 'no') {
    printPDF(dept, results, semester, semesters);  
  } else {
    // User clicked "Cancel" or X in the title bar.        
    ss.toast('Report Complete', 'Finished', 10);
    Logger.log('Report complete, no print.');
  }
}
//************************************************************************************************************************************************************************
//https://stackoverflow.com/questions/45209619/google-apps-script-getasapplication-pdf-layout
//https://ctrlq.org/code/19869-email-google-spreadsheets-pdf
//Export the function as a PDF
function printPDF(dept, results, semester, semesters){
  //***********************************************************************
  dept = dept /*|| "AERO"*/;  
  semester = semester /*|| "Fall-2017"*/;      
  results = results || Browser.msgBox(
    'Would you like these reports emailed to you? \\n' +
    '\\nYes - Email you & save to Google Drive' +
    '\\nNo - No email, just save to Google Drive',    
    Browser.Buttons.YES_NO);
  //***********************************************************************
  //Get sheet info
  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = sourceSheet.getSheetByName(dept + "-"+semester);
  var parentFolder; //Folder to save PDF in
  var currentDate = new Date();
  
  //Checks if folder exists, if it doesn't, create it
  try {
    //Folder exists
    parentFolder = DriveApp.getFoldersByName('CLAS Course Evaluations').next();   
    Logger.log('folder exists');
  }
  catch(e) {
    //Folder doesn't exist, create folder
    parentFolder = DriveApp.createFolder('CLAS Course Evaluations');
    Logger.log('folder does not exist, creating folder');
  }
  
  //nane PDF
  var PDFName = dept +"-"+ semester + "-" + currentDate;
  SpreadsheetApp.getActiveSpreadsheet().toast('Creating PDF '+PDFName);
  
  // export url
  var PDFurl = 'https://docs.google.com/spreadsheets/d/'+sourceSheet.getId()+'/export?exportFormat=pdf&format=pdf' // export as pdf
  + '&size=letter'                           // paper size legal / letter / A4
  + '&portrait=true'                     // orientation, false for landscape
  + '&fitw=true'                        // fit to page width, false for actual size
  + '&sheetnames=true&printtitle=true' // hide optional headers and footers
  + '&pagenum=CENTER&gridlines=false' // hide page numbers and gridlines
  + '&fzr=false'                         // do not repeat row headers (frozen rows) on each page
  + '&top_margin=.75&bottom_margin=.75&left_margin=.25&right_margin=.25' //Narrow margins
  + '&gid='+outputSheet.getSheetId();    // the sheet's Id
  
  //authorize script
  var token = ScriptApp.getOAuthToken();
  
  // request export url
  var response = UrlFetchApp.fetch(PDFurl, {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });
  
  //name PDF blob
  var PDFBlob = response.getBlob().setName(PDFName+'.pdf');
  
  // delete pdf if it already exists
  var files = parentFolder.getFilesByName(PDFName);
  while (files.hasNext())
  {
    files.next().setTrashed(true);
  }
  
  // create pdf file from blob
  var createPDFFile = parentFolder.createFile(PDFBlob);  
  var folderURL = parentFolder.getUrl();
  
  //Display link to folder
  var htmlApp = HtmlService
  .createHtmlOutput("Your report has been saved to your Google Drive "+"<a href='"+folderURL+"'  target='_top'>here</a>.").setHeight(50);
  
  //crashes sheet, removing for now
  //htmlApp.append("<input type='button' value='Click here to delete the sheet' onClick='google.script.run.breakOperation("+"\""+dept+"\""+","+"\""+semester+"\""+")' />");
  
  SpreadsheetApp.getActiveSpreadsheet().show(htmlApp);
  
  //email PDF if the user wants
  if (results=='yes'){
    emailPDF(PDFBlob, folderURL);
  }
}
//************************************************************************************************************************************************************************
function emailPDF(PDFBlob, folderURL){
  
  // Send the PDF of the spreadsheet to this email address
  var email = Session.getActiveUser().getEmail(); 
  
  // Subject of email message
  var subject = "PDF generated from Student Evaluation"; 
  
  // Email Body
  var body = "This has also been saved to your Google Drive at "+folderURL;
  
  // If allowed to send emails, send the email with the PDF attachment
  if (MailApp.getRemainingDailyQuota() > 0) 
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      attachments:[PDFBlob]     
    });  
}
//************************************************************************************************************************************************************************
//Stop script
function breakOperation(dept, semester){
  //remove Output sheet
  try{
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dept + "-"+semester));
  } catch(e){
    Logger.log('Output not deleted');
    //return;
  }  
  //remove script sheet
  try{
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Running-Report-Function'));
  } catch(e){
    Logger.log('Running-Report-Function not deleted');
    //return;
  }
  
  //remove HTML overlay
  var output = HtmlService.createHtmlOutput('<script>google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(output, 'Force Quit');
  return true;
}
//************************************************************************************************************************************************************************
