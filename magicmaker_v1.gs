// This Script Admin,Owner and Developer is Haris Nasir.
// Any and all changes must be approved by Haris Nasir due to the complex build of this script.
// Thursday 05/21/2020 14:43 ET

// There are 2 parts to creating the daily C3-VOD Doc with this script. 
// This Script begins on Google Sheet "test"
// Part 1 is to fill out the syndication timestamps.
// Do NOT worry about formating this sheet, color's and N/A's will be filled out by the script.
// Once you are done filling in the syndication timestamps, Script will prompt you to enter the Air Date.
// Allow the format to be M/dd/YYYY  ie; 5/21/2020
// The Sheet's Sheet7, and Sheet6 will automatically be generated after you enter the Air Date

function MagicMaker(){
  // Sheet where your focus should be
  // Fill out Syndications here
  var test = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test");
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet6");
  var sheet7 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet7");
  var date = airDate();
  var lastRow = test.getLastRow(); //Row Number of the Last asset in our list for today
  var firstEmptyRow = sheet6.getLastRow()+1;
  var TotalRowsCreated = (lastRow-1)*12;
  
  // Creating Sheet7 First
  // Copy Network & Series into Sheet7
     copyNetworkSeries(test,sheet7,lastRow);
 
  // Copy Syndication for CMC-C & MPX into Sheet7
     copyCmcMpx(test,sheet7,lastRow);
  
  // Copy Syndication for Dish, Direct-TV, DirecTV-NBC@VOD50-C into Sheet7
     copyDishDirect(test,sheet7,lastRow);
  
  // Autoresize All column width's for Sheet7
     autoResize(sheet7);
  
  // Fill-in N/A's for empty cell's with no Syndication Timestamps
  // Sending 3 parameters : Sheet to Check for empty cells, Column Letter to Check, Row Number to Check till.
  // lastRow of Sheet7 would be the same as the lastRow in 'test'.
     chkEmpty(sheet7,"E",lastRow);
     chkEmpty(sheet7,"F",lastRow);
     chkEmpty(sheet7,"G",lastRow);
  
  // Sheet7 is Complete, Now we create the VOD doc
  // We will still be using the data from 'test' sheet
     createVOD(test,sheet6,lastRow,date);
  
  // Allign text Center - The VOD doc
     centerMeDoc(sheet6,TotalRowsCreated,firstEmptyRow);
  
 
 }


function copyNetworkSeries(Test,Sheet7,rowNum){
  // Copies Formats and Copies Network and Series from test and Paste them to appropriate columns in sheet7.
  
  // Declaring sheets as variables
  var test = Test;
  var sheet7 = Sheet7;
  var lastRow = rowNum;
  var safePlace = "A2"; // Home Cell Location for Network & Series
  
  // Cell Poisiton where we wish to begin Copying data
  test.getRange('A2').activate();
  
  var currentCell = test.getCurrentCell();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate(); //Selecting all rows with data in Column
  currentCell.activateAsCurrentCell();
  
  var range1 = "A2:B"+lastRow;
  test.getRange(range1).setBackground('#00ff00');
  test.getRange(range1).activate();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate(); //Selecting All columns associated with range
  currentCell.activateAsCurrentCell();
  
  sheet7.getRange('A2').activate(); // Cell Position of where we wish to begin Pasting on Sheet7
  var range2 = "test!A2:B"+lastRow;
  sheet7.getRange(range2).copyTo(sheet7.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);  

}

function copyCmcMpx(Test,Sheet7,rowNum){
  // Copies Formats and Copies timestamps of CMC-C and MPX from test and Paste them to appropriate columns in sheet7.
  
  // Declaring sheets as variables
  var test = Test;
  var sheet7 = Sheet7;
  var lastRow = rowNum;
  var safePlace = "H2"; // Home Cell location for CMC-C & MPX
  
  // Cell Poisiton where we wish to begin Copying data
  test.getRange('H2').activate();

  var currentCell = test.getCurrentCell();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate(); //Selecting all rows with data in Column
  currentCell.activateAsCurrentCell();
  
  var range1 = "H2:I"+lastRow;
  test.getRange(range1).setBackground('#00ff00');
  test.getRange(range1).activate();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate(); //Selecting All columns associated with range
  currentCell.activateAsCurrentCell();
  
  sheet7.getRange('C2').activate(); // Cell Position of where we wish to begin Pasting on Sheet7
  var range2 = "test!H2:I"+lastRow;
  sheet7.getRange(range2).copyTo(sheet7.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);  
 
}

function copyDishDirect(Test,Sheet7,rowNum){
  // Copies Formats and Copies timestamps of Dish, Direct-TV, DirecTV-NBC@VOD50-C from test and Paste them to appropriate columns in sheet7.
  
  // Declaring sheets as variables
  var test = Test;
  var sheet7 = Sheet7;
  var lastRow = rowNum;
  var safePlace = "K2"; // Home Cell location for Dish, Direct-TV, DirecTV-NBC@VOD50-C
  
  test.getRange('K2').activate();

  var currentCell = test.getCurrentCell();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate(); //Selecting all rows with data in Column
  currentCell.activateAsCurrentCell();
  
  var range1 = "K2:M"+lastRow;
  test.getRange(range1).setBackground('#00ff00');
  test.getRange(range1).activate();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate(); //Selecting All columns associated with range
  currentCell.activateAsCurrentCell();
  
  sheet7.getRange('E2').activate(); // Cell Position of where we wish to begin Pasting on Sheet7
  var range2 = "test!K2:M"+lastRow;
  sheet7.getRange(range2).copyTo(sheet7.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);  
  test.setActiveSelection(safePlace); 
  
}

function autoResize(sheet){
  // Selecting all rows and columns and auto resizing the Column width
  // Allign text to center 
  
  var target = sheet;
  target.getRange(1, 1, target.getMaxRows(), target.getMaxColumns()).activate(); // Start at top-left corner, selecting all Rows and Columns with data in it.
  target.autoResizeColumns(1, 26); // Resizing Column Width
  
  var all = target.getRange(1, 1, target.getMaxRows(), target.getMaxColumns()).activate(); // Start at top-left corner, selecting all Rows and Columns with data in it.
  all.setHorizontalAlignment("center"); // Alligning text to center
  
  target.getRange('A2').activate(); // Unselect all columns and rows by selecting cell A2
}

function chkEmpty(sheet,col,lastRow){
  // Look for Empty Cell's and Add N/A's
  // Format Cell's to gray
  
  // Declaring sheets as variables
  var target = sheet;
  var targetCol = col;
  var lastR = lastRow;
  
  var count = 2
  
  // Looping through each row starting at row 2
  for (var count=2; count<=lastR; count++){
    var addr = col+count;  // Column Letter+Row Number  ie; E2
    target.getRange(addr).activate(); // Select starting cell
    var currentCell = target.getCurrentCell(); // Create variable of current cell to hold its value
    currentCell.activateAsCurrentCell();
    if(currentCell.getValue() == 0){ //If the cell is empty
      var location = currentCell.getA1Notation(); // Get the cell's location
      target.getRange(location).setValue("N/A");  // Add text "N/A"
      target.getRange(location).setBackground('#b7b7b7'); // Change color to Gray
    }  
  } 
  
}
        
function createVOD(Test,Sheet6,rowNum,Date){
  // Clear Logger
    Logger.clear();

  // Declaring variables
  var test = Test;
  var sheet6 = Sheet6;
  var lastRow = rowNum;
  var date = Date;
  
  // An Array to hold the 12 different Platform Devices
  var platforms = ["Android","AppleTV 3","AppleTV 4","Desktop","Directv","Dish","FireTV","iOS","Roku","Spectrum","X1","Xbox One"];
  
  //var lastAssetRow = test.getLastRow();
  //var date = airDate();
  
  for(var startRow=2;startRow<=lastRow;startRow++){ // Loop - Starting at Row 2 in test, As long as Row is less than The Last Row Number in test.

    var pasteRow = sheet6.getLastRow()+1; // Get the next Empty row in sheet6 to paste to
    var endRow = pasteRow+12; 
   
    for(i=1;i<=4;i++){  // For Each column starting at column 1. Note: 4 Columns to Paste into (Network,Series,Title,MCPiD)
      var temp = test.getRange(startRow,i).getValue(); // Variable to get the value we need to copy and paste into Sheet6
   
      for(var next=0;next<12;next++){ // paste in sheet6 at next available row, 12 times
        sheet6.getRange(pasteRow+next,i+2).setValue(temp); // Entering the Values Network,Series,Title,MCPiD
        sheet6.getRange(pasteRow+next,7).setValue(platforms[next]); // Entering the Platform Devices from Array
        var focusRow = pasteRow+next;
        sheet6.getRange('H'+focusRow).activate(); // Activating Starting Cell to Paste Status Formulas
        sheet6.getRange('H13:M13').copyTo(sheet6.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false); // Copy Pasting Formulas from H13:M13
        sheet6.getRange('B'+focusRow).setValue(date); // Adding Air Date to Column B
        sheet6.getRange("A"+focusRow).setFormula('=TEXT(B'+focusRow+',"MMMM")'); // Extracting Month from Column B and adding to Column A as Text format
      } // End next
    } // End i 
  } //End startRow
} // End Function

function centerMeDoc(Sheet6,TotalRows,FirstEmptyRow){
  // Resize All the Rows and Columns we created in VOD (sheet6 Doc)
  var sheet6 = Sheet6;
  var firstEmptyRow = FirstEmptyRow; // First Empty Row Number in sheet6
  var TotalRowsCreated = TotalRows;  // Total Rows Created
  var lastRow = firstEmptyRow+TotalRows;
  
  var all = sheet6.getRange("A"+firstEmptyRow+":F"+lastRow).activate();
  all.setHorizontalAlignment("center"); // Alligning text to center
  sheet6.getRange('A'+firstEmptyRow).activate(); 
  //return TotalRowsCreated;
}

function airDate(){
  // Prompt User to enter Air Date
  var ui = SpreadsheetApp.getUi();
  var date = ui.prompt("Please Enter the Air Date\nM/dd/YYYY").getResponseText(); // Store UserInput as Variable

  // Return Air Date Value
  return date;  
}




