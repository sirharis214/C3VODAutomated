// 
//           
function version7(){
    Logger.clear();
  var test = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test");
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet6");
  var platforms = ["Android","AppleTV 3","AppleTV 4","Desktop","Directv","Dish","FireTV","iOS","Roku","Spectrum","X1","Xbox One"];
  
  var lastAssetRow = test.getLastRow();
  //var date = airDate();
  /*
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
        sheet6.getRange('B'+focusRow).setValue(date);
        sheet6.getRange("A"+focusRow).setFormula('=TEXT(B'+focusRow+',"MMMM")');
      } // End next
    } // End i
  
  } //End startRow */
  var NumberOfRows = reSizeMe(sheet6,lastAssetRow);
  var firstRow = 14;
  var lastRow = 14+NumberOfRows;
  var all = sheet6.getRange("A"+firstRow+":F"+lastRow).activate();
  all.setHorizontalAlignment("center"); // Alligning text to center
  sheet6.getRange('A'+firstRow).activate(); 
  
} // End Function

function reSizeMe(Sheet6,rowNum){
  var sheet6 = Sheet6;
  var lastRow = rowNum;
  var TotalRowsCreated = (lastRow-1)*12;
  //Logger.log(TotalRowsCreated);
  return TotalRowsCreated;
  
}


function airDate(){
  //var date = Utilities.formatDate(new Date(), "GMT", "M-dd-YY");
  var ui = SpreadsheetApp.getUi();
  var date = ui.prompt("Please Enter the Air Date\nM/dd/YYYY").getResponseText();
  //Logger.log(date);
  return date;
  
}

// Testing.
function getMonth(){
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet6");
  var month = sheet6.getRange("A14").setFormula('=TEXT(B14,"MMMM")');
}
