function copyNetworkSeries(){
  // Works
  // Copies Network and Series from test and Paste to sheet 7 to appropriate columns
  var test = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test");
  var sheet7 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet7");
  var safePlace = "A2"; // Some safe cell
  
  test.getRange('A2').activate();
  var lastRow = test.getLastRow();
  //Logger.log(lastRow);
  
  var currentCell = test.getCurrentCell();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  
  var range1 = "A2:B"+lastRow;
  test.getRange(range1).setBackground('#00ff00');
  test.getRange(range1).activate();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  
  
  sheet7.getRange('A2').activate();
  var range2 = "test!A2:B"+lastRow;
  sheet7.getRange(range2).copyTo(sheet7.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);  
  test.setActiveSelection(safePlace);
  
  copyCmcMpx();
}

function copyCmcMpx(){
 // Works
  // Copies CMC-C and MPX from test and Paste to sheet 7 to appropriate columns
  var test = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test");
  var sheet7 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet7");
  var safePlace = "H2"; // Some safe cell
  
  test.getRange('H2').activate();
  var lastRow = test.getLastRow();
  Logger.log(lastRow);
  
  var currentCell = test.getCurrentCell();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  
  var range1 = "H2:I"+lastRow;
  test.getRange(range1).setBackground('#00ff00');
  test.getRange(range1).activate();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  
  sheet7.getRange('C2').activate();
  var range2 = "test!H2:I"+lastRow;
  sheet7.getRange(range2).copyTo(sheet7.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);  
  test.setActiveSelection(safePlace); 
  
  copyDishDirect();
}

function copyDishDirect(){
 // Works
  // Copies Dish,DirectTV,DirectTV-NBC  from test and Paste to sheet 7 to appropriate columns
  var test = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test");
  var sheet7 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet7");
  var safePlace = "K2"; // Some safe cell
  
  test.getRange('K2').activate();
  var lastRow = test.getLastRow();
  Logger.log(lastRow);
  
  var currentCell = test.getCurrentCell();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  
  var range1 = "K2:M"+lastRow;
  test.getRange(range1).setBackground('#00ff00');
  test.getRange(range1).activate();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  
  sheet7.getRange('E2').activate();
  var range2 = "test!K2:M"+lastRow;
  sheet7.getRange(range2).copyTo(sheet7.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);  
  test.setActiveSelection(safePlace); 
  
  // All columns to resize width
  autoResize(sheet7);
  // Fill in Syndication N/A's
  chkEmpty(sheet7,"E",lastRow);
  chkEmpty(sheet7,"F",lastRow);
  chkEmpty(sheet7,"G",lastRow);
}

function autoResize(sheet){
  var target = sheet;
  target.getRange(1, 1, target.getMaxRows(), target.getMaxColumns()).activate();
  target.autoResizeColumns(1, 26);
  
  var all = target.getRange(1, 1, target.getMaxRows(), target.getMaxColumns()).activate();
  all.setHorizontalAlignment("center");
  
  target.getRange('A2').activate();
}

function chkEmpty(sheet,col,lastRow){
  var target = sheet;
  var targetCol = col;
  var lastR = lastRow;
  
  var count = 2
  
  for (var count=2; count<=lastR; count++){
    var addr = col+count;
    target.getRange(addr).activate();
    var currentCell = target.getCurrentCell();
    currentCell.activateAsCurrentCell();
    if(currentCell.getValue() == 0){
      var location = currentCell.getA1Notation();
      Logger.log(location);
      Logger.log("Cell empty"); 
      target.getRange(location).setValue("N/A");
      target.getRange(location).setBackground('#b7b7b7');
    }
    
    /*else{
      Logger.log("Cell NOT Empty");
    } */
  }
  
  //Logger.clear();
}

