// Works
// Will copy over from test, one asset 12 times to the correct row's in sheet6
//

function version3(){
    Logger.clear();
  var test = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test");
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet6");
  
  var startRow = 2;
  var pasteRow = sheet6.getLastRow()+1;
  var endRow = pasteRow+12;
  for(i=1;i<=4;i++){
    // copy the cell from test
   var temp = test.getRange(startRow,i).getValue();
   
    // paste in sheet6 at next avialable row 12 times
    for(var next=0;next<12;next++){
    sheet6.getRange(pasteRow+next,i+2).setValue(temp);
    }
  }
}
