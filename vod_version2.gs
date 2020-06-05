// Able to get asset as array 
// Copy into sheet6, 
//Bugs:
//     Copyiny correctly into appropriate Column and Rows
//     But it is copying only asset ID in all cells
//     Logs do show all 4 values accurately, (network,series,episodename,mcpID)



function version2(){
    Logger.clear();
  var test = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test");
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet6");
  test.getRange('A2').activate();
  sheet6.getRange('A2').activate();
  
  var currentCell = test.getCurrentCell();
  var lastRow = test.getLastRow();
 
  // on sheet test, start copying at row 
  for(var row=2; row<=lastRow; row++){
    //var Row = test.getRange(row,1,1).getValue(); 
    var tempArray = test.getRange(row,1,1,4).getValues();
   // Logger.log(tempArray);
    
    for(var i=0;i<tempArray.length;i++){
      var eachShowArray = tempArray[i];
      
      for(var ct=0; ct<eachShowArray.length; ct++){
        var pasteThis = eachShowArray[ct];
        //Logger.log(pasteThis);
        //Above this, pastThis is being Logged accuratley 
        //Below this, pasteThis value stays as ID
        //Pasting each index 48 times across all cells
        // ????? WHY !! ?????
        // Loop is Broken somewhere, it should...
        // start at index 0, complete a row, switch index, complete row, switch index ..complete row etc
        for(var row=14;row<26;row++){
          for(var col=3;col<=6;col++){
          sheet6.getRange(row,col,1).setValue(pasteThis);
            Logger.log(ct);
            Logger.log(pasteThis);
            
          }
        }
      }  
    }
    
    //on sheet test, start copying at col of row
    //getRange(row#,col#,Include#ofRow
   /* for(var col=1; col<=4;col++){
      var eachPart = test.getRange(row,col,1).getValue();
      // each cell of all assets
      //Logger.log(eachPart);
      
      var rowPaste = sheet6.getLastRow()+1;
      Logger.log(rowPaste);
  
    } */
  }
  
}

