/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen2() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Generate Output', functionName: 'cloneGoogleSheet_'},
    {name: 'Place Holder', functionName: 'createCopy_'}
  ];
  spreadsheet.addMenu('Generate Sheets', menuItems);
}

 
function updateDataTemplate2(target, current){

    

    target.getRange('C1').setValue('=datasheet!A'+current);
    target.getRange('C2').setValue('=datasheet!B'+current);

    target.getRange('B7').setValue('=datasheet!G'+current);
 
    target.getRange('C7').setValue('=datasheet!H'+current);
    target.getRange('D7').setValue('=datasheet!F'+current);
    /*COURSE */
    target.getRange('B8').setValue('=INDEX(DataSheet!G:G,MATCH(DataSheet!$C'+current+',DataSheet!$B:$B,0))');
    target.getRange('C8').setValue('=INDEX(DataSheet!H:H,MATCH(DataSheet!$C'+current+',DataSheet!$B:$B,0))');
    target.getRange('D8').setValue('=INDEX(DataSheet!F:F,MATCH(DataSheet!$C'+current+',DataSheet!$B:$B,0))');
    /* GROUP */ 
    target.getRange('B9').setValue('=INDEX(DataSheet!G:G,MATCH(DataSheet!$D'+current+',DataSheet!$C:$C,0))');
    target.getRange('C9').setValue('=INDEX(DataSheet!H:H,MATCH(DataSheet!$D'+current+',DataSheet!$C:$C,0))');
    target.getRange('D9').setValue('=INDEX(DataSheet!F:F,MATCH(DataSheet!$D'+current+',DataSheet!$C:$C,0))');
    /*   DEPT   */
    target.getRange('B10').setValue('=INDEX(DataSheet!G:G,MATCH(DataSheet!$W'+current+',DataSheet!$C:$C,0))');
    target.getRange('C10').setValue('=INDEX(DataSheet!H:H,MATCH(DataSheet!$W'+current+',DataSheet!$C:$C,0))');
    target.getRange('D10').setValue('=INDEX(DataSheet!F:F,MATCH(DataSheet!$W'+current+',DataSheet!$C:$C,0))');

    target.getRange('B14').setValue('=datasheet!X'+current);
    target.getRange('C14').setValue('=datasheet!Y'+current);
    target.getRange('D14').setValue('=datasheet!Z'+current);
    target.getRange('E14').setValue('=datasheet!AA'+current);
    target.getRange('F14').setValue('=datasheet!AB'+current);
    target.getRange('G14').setValue('=datasheet!AC'+current);
    target.getRange('H14').setValue('=datasheet!AD'+current);

    target.getRange('B15').setValue('=INDEX(datasheet!X:X,MATCH(datasheet!C'+current+',datasheet!$B:$B,0))');
    target.getRange('C15').setValue('=INDEX(datasheet!Y:Y,MATCH(datasheet!C'+current+',datasheet!$B:$B,0))');
    target.getRange('D15').setValue('=INDEX(datasheet!Z:Z,MATCH(datasheet!C'+current+',datasheet!$B:$B,0))');
    target.getRange('E15').setValue('=INDEX(datasheet!AA:AA,MATCH(datasheet!C'+current+',datasheet!$B:$B,0))');
    target.getRange('F15').setValue('=INDEX(datasheet!AB:AB,MATCH(datasheet!C'+current+',datasheet!$B:$B,0))');
    target.getRange('G15').setValue('=INDEX(datasheet!AC:AC,MATCH(datasheet!C'+current+',datasheet!$B:$B,0))');
    target.getRange('H15').setValue('=INDEX(datasheet!AD:AD,MATCH(datasheet!C'+current+',datasheet!$B:$B,0))');

    target.getRange('B16').setValue('=INDEX(datasheet!X:X,MATCH(datasheet!D'+current+',datasheet!$C:$C,0))');
    target.getRange('C16').setValue('=INDEX(datasheet!Y:Y,MATCH(datasheet!D'+current+',datasheet!$C:$C,0))');
    target.getRange('D16').setValue('=INDEX(datasheet!Z:Z,MATCH(datasheet!D'+current+',datasheet!$C:$C,0))');
    target.getRange('E16').setValue('=INDEX(datasheet!AA:AA,MATCH(datasheet!D'+current+',datasheet!$C:$C,0))');
    target.getRange('F16').setValue('=INDEX(datasheet!AB:AB,MATCH(datasheet!D'+current+',datasheet!$C:$C,0))');
    target.getRange('G16').setValue('=INDEX(datasheet!AC:AC,MATCH(datasheet!D'+current+',datasheet!$C:$C,0))');
    target.getRange('H16').setValue('=INDEX(datasheet!AD:AD,MATCH(datasheet!D'+current+',datasheet!$C:$C,0))');

    target.getRange('B17').setValue('=INDEX(datasheet!X:X,MATCH(datasheet!W'+current+',datasheet!$C:$C,0))');
    target.getRange('C17').setValue('=INDEX(datasheet!Y:Y,MATCH(datasheet!W'+current+',datasheet!$C:$C,0))');
    target.getRange('D17').setValue('=INDEX(datasheet!Z:Z,MATCH(datasheet!W'+current+',datasheet!$C:$C,0))');
    target.getRange('E17').setValue('=INDEX(datasheet!AA:AA,MATCH(datasheet!W'+current+',datasheet!$C:$C,0))');
    target.getRange('F17').setValue('=INDEX(datasheet!AB:AB,MATCH(datasheet!W'+current+',datasheet!$C:$C,0))');
    target.getRange('G17').setValue('=INDEX(datasheet!AC:AC,MATCH(datasheet!W'+current+',datasheet!$C:$C,0))');
    target.getRange('H17').setValue('=INDEX(datasheet!AD:AD,MATCH(datasheet!W'+current+',datasheet!$C:$C,0))');

    target.getRange('E20').setValue('=datasheet!P'+current);
    target.getRange('E21').setValue('=datasheet!E'+current);
    target.getRange('E22').setValue('=INDEX(datasheet!P:P,MATCH(datasheet!D'+current+',datasheet!$C:$C,0))');
    target.getRange('E23').setValue('=INDEX(datasheet!P:P,MATCH(datasheet!W'+current+',datasheet!$C:$C,0))');
    target.getRange('E24').setValue('=datasheet!AE'+current);
    target.getRange('E25').setValue('=INDEX(datasheet!AE:AE,MATCH(datasheet!C'+current+',datasheet!$B:$B,0))');
    target.getRange('E26').setValue('=INDEX(datasheet!D:D,MATCH(datasheet!D'+current+',datasheet!$C:$C,0))');
    target.getRange('E27').setValue('=INDEX(datasheet!D:D,MATCH(datasheet!W'+current+',datasheet!$C:$C,0))');

    target.getRange('G32').setValue('=datasheet!Q'+current);
    target.getRange('G33').setValue('=datasheet!R'+current);
    target.getRange('G34').setValue('=datasheet!S'+current);
    target.getRange('G35').setValue('=datasheet!T'+current);
    target.getRange('G36').setValue('=datasheet!U'+current);
    target.getRange('G37').setValue('=datasheet!V'+current);

    target.getRange('H32').setValue('=INDEX(datasheet!Q:Q,MATCH(datasheet!$D'+current+',datasheet!$C:$C,0))');
    target.getRange('H33').setValue('=INDEX(datasheet!R:R,MATCH(datasheet!$D'+current+',datasheet!$C:$C,0))');
    target.getRange('H34').setValue('=INDEX(datasheet!S:S,MATCH(datasheet!$D'+current+',datasheet!$C:$C,0))');
    target.getRange('H35').setValue('=INDEX(datasheet!T:T,MATCH(datasheet!$D'+current+',datasheet!$C:$C,0))');
    target.getRange('H36').setValue('=INDEX(datasheet!U:U,MATCH(datasheet!$D'+current+',datasheet!$C:$C,0))');
    target.getRange('H37').setValue('=INDEX(datasheet!V:V,MATCH(datasheet!$D'+current+',datasheet!$C:$C,0))');

    target.getRange('I32').setValue('=INDEX(datasheet!Q:Q,MATCH(datasheet!$W'+current+',datasheet!$C:$C,0))');
    target.getRange('I33').setValue('=INDEX(datasheet!R:R,MATCH(datasheet!$W'+current+',datasheet!$C:$C,0))');
    target.getRange('I34').setValue('=INDEX(datasheet!S:S,MATCH(datasheet!$W'+current+',datasheet!$C:$C,0))');
    target.getRange('I35').setValue('=INDEX(datasheet!T:T,MATCH(datasheet!$W'+current+',datasheet!$C:$C,0))');
    target.getRange('I36').setValue('=INDEX(datasheet!U:U,MATCH(datasheet!$W'+current+',datasheet!$C:$C,0))');
    target.getRange('I37').setValue('=INDEX(datasheet!V:V,MATCH(datasheet!$W'+current+',datasheet!$C:$C,0))');

}

function copyRowsToEnd_2(source, target)
{
/* You must use the copyValuesToRange function so the data values are copied, not the formulas.  
  You then use copyFormatToRange to copy the range.  
  
  lastEntryStart is the start of the new data, lastEntryEnd the end.  You must set them because getLastRow changes between the function calls.
 */
  //var destRange = target.getRange(target.getLastRow()+2,1);
  var sourceVal = source.getRange(1,1,38,11);
  //sourceVal.copyTo(destRange);
  var lastEntryStart = target.getLastRow()+2+12;
  var lastEntryEnd = target.getLastRow()+39;
  sourceVal.copyValuesToRange(target,1, 11,target.getLastRow()+2,target.getLastRow()+39);
  sourceVal.copyFormatToRange(target,1, 11,lastEntryStart,lastEntryEnd);
}



function cloneGoogleSheet_2() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Template').copyTo(ss);

  
  /* Before cloning the sheet, delete any previous copy */
  var old = ss.getSheetByName("Output");
  if (old) ss.deleteSheet(old); // or old.setName(new Name);
  var old2 = ss.getSheetByName("Test1");
  if (old2) ss.deleteSheet(old2); // or old.setName(new Name);
  
  target  = ss.insertSheet();
  target.setName("Output");
  
  SpreadsheetApp.flush(); // Utilities.sleep(2000);
  sheet.setName("Test1");
 
  /* Make the new sheet active */
  ss.setActiveSheet(sheet);
  SpreadsheetApp.getActiveSheet()
  for(current=2; current<26; current++){
    updateDataTemplate(sheet, current);
    copyRowsToEnd_(sheet, target);
  }
  
  var old2 = ss.getSheetByName("Test1");
  if (old2) ss.deleteSheet(old2); // or old.setName(new Name);
  
  ss.setActiveSheet(ss.getSheetByName("Output"));
  
}

