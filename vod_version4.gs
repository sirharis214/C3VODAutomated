// Works
// Paste each Asset 12 times along with status 1,2,3 Formulas from  H13:M13
// H13:M13 Must be protected at all cost.
                 
function version6(){
    Logger.clear();
  var test = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test");
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet6");
  var platforms = ["Android","AppleTV 3","AppleTV 4","Desktop","Directv","Dish","FireTV","iOS","Roku","Spectrum","X1","Xbox One"];
  
  var lastAssetRow = test.getLastRow();
  
  for(var startRow=2;startRow<=lastAssetRow;startRow++){

    var pasteRow = sheet6.getLastRow()+1;
    var endRow = pasteRow+12;
   
    for(i=1;i<=4;i++){
      // copy the cell from test
     var temp = test.getRange(startRow,i).getValue();
   
      // paste in sheet6 at next avialable row 12 times
      for(var next=0;next<12;next++){
        sheet6.getRange(pasteRow+next,i+2).setValue(temp);
        sheet6.getRange(pasteRow+next,7).setValue(platforms[next]);
        var focusRow = pasteRow+next;
        sheet6.getRange('H'+focusRow).activate();
        sheet6.getRange('H13:M13').copyTo(sheet6.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      } // End next
    } // End i
  
  } //End startRow
} // End Function


function PasteFormula() {
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet6");
  sheet6.getRange('H14').activate();
  sheet6.getRange('H13:M13').copyTo(sheet6.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
};
