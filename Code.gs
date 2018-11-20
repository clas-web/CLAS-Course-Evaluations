//https://productforums.google.com/forum/#!topic/docs/67lySF27l5c
/**
* Creates a trigger for when a spreadsheet opens.
*/

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Run Script', functionName: 'createSpreadsheetOpenTrigger'},
  ];
    spreadsheet.addMenu('Click this first', menuItems);
    
    var ss = SpreadsheetApp.getActive();
    ScriptApp.newTrigger('myOnEdit')    
    .forSpreadsheet(ss)
    .onEdit()
    .create();    
    }    
    
    //Create trigger for sheet to have permissions
    function createSpreadsheetOpenTrigger() {
    var ss = SpreadsheetApp.getActive();
    ScriptApp.newTrigger('myOnEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
    }
    
    
    //************************************************************************************************************************************************************************
    //Automatically add laptop to history sheet when "checked" back in
    function checkIn(e){
    var spss = SpreadsheetApp.getActiveSpreadsheet();
    var laptops = spss.getSheetByName("laptop info-current"); //laptop sheet
    var destinationSheet = spss.getSheetByName("checkout history"); //history sheet
    var destRow = destinationSheet.getLastRow(); //last row of the history sheet
    var destRange = destinationSheet.getRange(destRow+1, 4, 1, 20); //this is the history range we want our checkout info to be copied to
    var destInputValues = destinationSheet.getRange(destRow+1, 1, 1, 3); //this is the history range where we add our date and issue reports
    var currentDate = new Date(); //the current date and time
    //***************************************************************************************************
    //Check laptop in, report issues (optional)
    var ui = SpreadsheetApp.getUi();    
    var userIssues = ui.alert('OAT Laptop Check In', 'Did the user report any issues with the laptop?', ui.ButtonSet.YES_NO_CANCEL);
    var issueWas;
    var issueWasText;
    Logger.log("userIssues: " + userIssues);
    if (userIssues == ui.Button.YES){
    //report user's issue
    issueWas = ui.prompt('OAT Laptop Check In', 'What was the user\'s issue?', ui.ButtonSet.OK_CANCEL);
    issueWasText = issueWas.getResponseText();
    } else if (userIssues == ui.Button.NO){
    //report that there was no issue
    issueWasText = "n";
    } else {
    //quit
    Logger.log("return (quit)");
    //e.range.setValue("TRUE"); //check the box again since the operation was cancelled
    return;
    }  
    Logger.log("userIssues: "+issueWasText);
    //***************************************************************************************************
    var techIssues = ui.alert('OAT Laptop Check In', 'Did you find any issues with the laptop?', ui.ButtonSet.YES_NO_CANCEL);  
    var techIssueWas;
    var techIssueWasText;
    Logger.log("techIssues: " + techIssues);
    if (techIssues == ui.Button.YES){
    //report your issue
    techIssueWas = ui.prompt('OAT Laptop Check In', 'What was your issue?', ui.ButtonSet.OK_CANCEL);
    techIssueWasText = techIssueWas.getResponseText();
    } else if (techIssues == ui.Button.NO){
    //report that you found no issue
    techIssueWasText = "n";
    } else {
    //quit
    Logger.log("return (quit)");
    // e.range.setValue("TRUE"); //check the box again since the operation was cancelled
    return;
    }  
    Logger.log("techIssues: "+techIssueWasText);
    //***************************************************************************************************
    //Copy values to history sheet, then clear from laptop sheet if value is marked false
    if (e.range.getDisplayValue()=="FALSE"){
    var laptopRange = laptops.getRange(e.range.getRow(),4,1,20); //get laptop information
    var userRange = laptops.getRange(e.range.getRow(),14,1,20); //get user infomation
    var inputArray = [[currentDate,issueWasText,techIssueWasText]]; //create an array of our input information to copy to range
  destInputValues.setValues(inputArray); //copy date and issue reports to history sheet
  var copiedText = laptopRange.copyTo(destRange); //copy laptop + user information to history sheet
  userRange.clearContent(); //delete user information from laptop sheet for future users
  spss.toast("The laptop has been checked in.", "OAT Laptop Check In", 10); //Display a pop-up to indicate the laptop has been checked in
} else {    
  Logger.log("e isn't false");    
}  
}
//************************************************************************************************************************************************************************
//Check out laptop to a user
function checkOut(userDept, userLastName, userFirstName, userPhone, userNinernet, userReturnDate, userNotes, userOther){
  var spss = SpreadsheetApp.getActiveSpreadsheet();
  var laptops = spss.getSheetByName("laptop info-current"); //laptop sheet
  var thisRow = /*e.range.getRow() || */laptops.getActiveRange().getRow(); //row of the checked out laptop
  var userRange = laptops.getRange(thisRow, 14, 1, 9); //this is the laptop range where we want our checkout info to be place
  var currentDate = new Date(); //the current date and time
  var userDataArray = [];
  
  //run function for dialog and return array of choices
  userDataArray = [[userDept, userLastName, userFirstName, userPhone, userNinernet, currentDate, userReturnDate, userNotes, userOther]];
  
  /*//Run if no data was entered, uncheck value
  if (userDataArray.join()==""){
  laptops.getRange(thisRow,3).setValue("FALSE"); //uncheck the box again since the operation was cancelled   
  return;
  }*/
  
  userRange.setValues(userDataArray);
  Logger.log("userDataArray: "+userDataArray.join(""));
  Logger.log("userRange: "+userRange.getValues());
}
//************************************************************************************************************************************************************************
//open HTML dialog for checking out a laptop
/**
* @NotOnlyCurrentDoc
*/
function checkOutHTMLDialog(){
  
  //Declare variables
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //Course Evaluations spreadsheet
  
  //get dept codes from my CLAS OAT sheet
  var dropdownMenu = "<strong>Department</strong> <br><form> <select id='dpt-dropdown'>";
  var deptCodes = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1zNrJRaS1eMSB6vTCq_FB78Zrphq8mfzxSEbGaBsa3Wk/edit#gid=997728017')
  .getSheetByName('Departments').getRange('B2:B').getValues();
  //create dropdown menu for depts
  for (var m=0; m < deptCodes.length; m++){
    dropdownMenu += "<option value="+deptCodes[m]+">"+deptCodes[m]+"</option>";    
  }
  dropdownMenu += "</select> </form> <p id='dept'></p>";
  
  //Prepare HTML dialog
  var HTMLMessage = '<form> <strong>First Name</strong><input type="text" name="userFirstName" style="float: right;"/> <br><br>' +
    '<strong>Last Name</strong><input type="text" name="userLastName" style="float: right;"/> <br><br>' +
      '<strong>Phone (x7####) </strong><input type="tel" name="userPhone" style="float: right;"/> <br><br>' +
        '<strong>NinerNET</strong><input type="text" name="userNinernet" style="float: right;"/> <br><br>' +
          '<strong>Estimated Return Date</strong><input type="date" name="userReturnDate" style="float: right;"/> <br><br>' +
            '<strong>Notes</strong><input type="text" name="userNotes" style="float: right;"/> <br><br>' +
              '<strong>Other Equipment Lended</strong><input type="text" name="userOther" style="float: right;"/> <br><br>' +
                '<input type="submit" value="Submit" onClick="formSubmit()" /> </form> ';
  
  //Prepare HTML script in dialog that calls GAS function
  var HTMLScript = '<script type="text/javascript"> function formSubmit(userLastName, ' + 
    'userFirstName, userPhone, userNinernet, userReturnDate, userNotes, userOther) { var deptChoice = document.getElementById("dpt-dropdown").value;'+
      'var userLastName=document.getElementsByName("userLastName")[0].value;' +
        'var userFirstName=document.getElementsByName("userFirstName")[0].value;' +
          'var userPhone=document.getElementsByName("userPhone")[0].value;' +
            'var userNinernet=document.getElementsByName("userNinernet")[0].value;' +
              'var userReturnDate=document.getElementsByName("userReturnDate")[0].value;' +
                'var userNotes=document.getElementsByName("userNotes")[0].value;' +
                  'var userOther=document.getElementsByName("userOther")[0].value;' +
                    'document.getElementById("dept").innerHTML = deptChoice; google.script.run.checkOut(deptChoice, userLastName, ' + 
                      'userFirstName, userPhone, userNinernet, userReturnDate, userNotes, userOther); google.script.host.close();} ' + 
                        '</script> </body>';
  
  //Display HTML output  
  var htmlApp = HtmlService  
  .createHtmlOutput("<body>" + dropdownMenu + HTMLMessage + HTMLScript)
  .setWidth(400)
  .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(htmlApp, "User Information");     
  var output = HtmlService.createTemplate(htmlApp);
  //Receive HTML output as a formatted HTML file for troubleshooting
  //Logger.log(output.getCode());  
}
//************************************************************************************************************************************************************************
//run function whenver sheet is edited
function myOnEdit(e){
  //End if the active cell is not in the "Checked Out?" column
  if (e.range.columnStart != 3) {
    Logger.log("Wrong column, won't delete anything");
    return;
  }  
  
  //Check out laptop if value is checked
  if (e.range.getDisplayValue() == "TRUE"){        
    checkOutHTMLDialog();
  } else if (e.range.getDisplayValue() == "FALSE") {    
    //Check in laptop if value is unchecked
    checkIn(e);  //run script to check in laptop  
  } else {
    return;
  }
}
//************************************************************************************************************************************************************************
/* Uncheck active cell
function setFalse(){
var laptops = spss.getSheetByName("laptop info-current"); //laptop sheet
var thisRow = laptops.getActiveRange().getRow(); //row of the checked out laptop
laptops.getRange(thisRow,3).setValue("FALSE"); //uncheck the box again since the operation was cancelled 
}*/