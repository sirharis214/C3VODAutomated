// This Script Admin,Owner and Developer is Haris Nasir.
// Any and all changes must be approved by Haris Nasir due to the complex build of this script.
// Thursday 05/21/2020 14:43 ET

// * Update Version 14 June 2,2020 5:15 AM
// Building on Version 13, 
// magicmakerv5 in git
// Upon completion of Creating VOD doc, Hiding Rows with date 5/14/2020 ( formula rows) 
// Purpose: inorder to have VODdoc ready for L1's to begin Sweeps. 
// * Cleaned up Code from rough drafts.
// * Updated all Log updates and texts to keep track of Script progress and performance. 


function BuildSyndication(){
  // This function is for the button "Build Syndication",
  // which builds only the syndication sheet without creating the VODdoc afterwards.
  
  Logger.clear();
  Logger.log("Beginning to create Syndication Sheet...");
  var test = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test");
  var sheet7 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet7");
  var lastRow = test.getLastRow(); //Row Number of the Last asset in our list for today

  // Copy Network & Series into Sheet7
     Logger.log("Copying Network and Series...");
     copyNetworkSeries(test,sheet7,lastRow);
 
  // Copy Syndication for CMC-C & MPX into Sheet7
     Logger.log("Copying Syndication for CMC-C and MPX...");
     copyCmcMpx(test,sheet7,lastRow);
  
  // Copy Syndication for Dish, Direct-TV, DirecTV-NBC@VOD50-C into Sheet7
     Logger.log("Copying Sundication for Dish and DirectTV...");
     copyDishDirect(test,sheet7,lastRow);
  
  // Autoresize All column width's for Sheet7
     Logger.log("Resizing Columns to fit text's...");
     autoResize(sheet7);

  // Fill-in N/A's for empty cell's with no Syndication Timestamps
  // Sending 3 parameters : Sheet to Check for empty cells, Column Letter to Check, Row Number to Check till.
  // lastRow of Sheet7 would be the same as the lastRow in 'test'.
       Logger.log("Filling in N/A's...");
     chkEmpty(sheet7,"C",lastRow);
     chkEmpty(sheet7,"D",lastRow);
     chkEmpty(sheet7,"E",lastRow);
     chkEmpty(sheet7,"F",lastRow);
     chkEmpty(sheet7,"G",lastRow);
       Logger.log("N/A's have been Inserted.");
  
}

function BuildDoc(){
  // This function is for the button "Build Doc",
  // which builds the VOD doc without the step of creating Syndication Sheet
  
  Logger.clear();
  //Beginning to create VODdoc...
  var test = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test");
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet6");
  var brand6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Brand6");
  var date = airDate();
  var lastRow = test.getLastRow(); //Row Number of the Last asset in our list for today
  var firstEmptyRow = sheet6.getLastRow()+1;
  var TotalRowsCreated = (lastRow-2)*12;
  
   //Logger.log("FirstEmptyRow Before Creating Doc: "+firstEmptyRow);

  // Before creating Doc, we must clear all filters and reveal rows 2-13
  // They contain Data validation Formulas for Columns H:M 
     Logger.log("Filtering and Revealing VODdoc rows with date 5/14/2020.");
     var UniqueDates = showProtected(sheet6);
    
  // We will still be using the data from 'test' sheet
     Logger.log("Starting process to create VODdoc...");
     createVOD(test,sheet6,lastRow,date);
  
  // Fill Brand 6's
     Logger.log("Inserting Brand 6's...");
     Brand6(test,sheet6,brand6,firstEmptyRow);
  
  // Allign text Center and Resize Columns 1-7 - VODdoc
     Logger.log("Aligning text Center and Resizing Columns...");
     centerMeDoc(sheet6,TotalRowsCreated,firstEmptyRow);
  
  // Final step, Hide rows with date 5/14/2020
     Logger.log("Hiding rows with Date 5/14/2020");
     UniqueDates.push("5/14/2020");
     hideFormula(sheet6,UniqueDates);
  sheet6.getRange('A'+firstEmptyRow).activate();
  Logger.log("VODdoc Ready!");
}

function MagicMaker(){
  // Sheet where your focus should be
  // Fill out Syndications here
  
  Logger.clear();
  Logger.log("We Will Create Syndication Sheet &\nAutomatically begin to create VODdoc.");
  var test = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test");
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet6");
  var brand6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Brand6");
  var sheet7 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet7");
  var date = airDate();
  var lastRow = test.getLastRow(); //Row Number of the Last asset in our list for today
  var firstEmptyRow = sheet6.getLastRow()+1;
  var TotalRowsCreated = (lastRow-2)*12;

  // Creating Sheet7 First
     Logger.log("Beginning process to create Syndication Sheet...");
  // Copy Network & Series into Sheet7
     Logger.log("Copying Network and Series...");
     copyNetworkSeries(test,sheet7,lastRow);
 
  // Copy Syndication for CMC-C & MPX into Sheet7
     Logger.log("Copying Syndication for CMC-C and MPX...");
     copyCmcMpx(test,sheet7,lastRow);
  
  // Copy Syndication for Dish, Direct-TV, DirecTV-NBC@VOD50-C into Sheet7
     Logger.log("Copying Sundication for Dish and DirectTV...");
     copyDishDirect(test,sheet7,lastRow);
  
  // Autoresize All column width's for Sheet7
     Logger.log("Resizing Columns to fit text's...");
     autoResize(sheet7);

  // Fill-in N/A's for empty cell's with no Syndication Timestamps
  // Sending 3 parameters : Sheet to Check for empty cells, Column Letter to Check, Row Number to Check till.
  // lastRow of Sheet7 would be the same as the lastRow in 'test'.
       Logger.log("Filling in N/A's...");
     chkEmpty(sheet7,"C",lastRow);
     chkEmpty(sheet7,"D",lastRow);
     chkEmpty(sheet7,"E",lastRow);
     chkEmpty(sheet7,"F",lastRow);
     chkEmpty(sheet7,"G",lastRow);
       Logger.log("N/A's have been Inserted.");
  
  // Sheet7 is Complete, 
     Logger.log("Syndication Sheet Created.");
  
  // Before creating Doc, we must clear all filters and reveal rows 2-13
  // They contain Data validation Formulas for Columns H:M 
     Logger.log("Starting process of Filtering and Revealing VODdoc rows with date 5/14/2020...");
     var UniqueDates = showProtected(sheet6);
  
  // We will still be using the data from 'test' sheet
     Logger.log("Starting process to create VODdoc...");
     createVOD(test,sheet6,lastRow,date);
  
  // Fill Brand 6's
     Logger.log("Inserting Brand 6's...");
     Brand6(test,sheet6,brand6,firstEmptyRow);
  
  // Allign text Center and Resize Columns 1-7 - VODdoc
     Logger.log("Aligning text Center and Resizing Columns...");
     centerMeDoc(sheet6,TotalRowsCreated,firstEmptyRow);
  
  // Final step, Hide rows with date 5/14/2020
     Logger.log("Hiding rows with Date 5/14/2020");
     UniqueDates.push("5/14/2020");
     hideFormula(sheet6,UniqueDates);
  sheet6.getRange('A'+firstEmptyRow).activate();
  Logger.log("VODdoc Ready!");
   
 }

function showProtected(sheet){
  // showProtected() filters out any dates to show us a clean sheet that only displays rows 1-13
  // Remember: rows 2-13 must be unfiltered from the doc, we need to see these rows to get column H:M for status formulas.
  // We do this by showing rows that have the date 5/14/2020 (dates of Column 2-13)
     var targetSheet = sheet;

  // Clear any filters that may be applied to sheet
     Logger.log("Checking for any filters in VODdoc...");
     checkFilter(targetSheet);
  
  // Get an Array of all unique dates in doc
     var myDates = ridDouble(targetSheet);
     //Logger.log(myDates);
  
  // Create Filter to show data for 5/14/2020 which contain the filters we need.
     Logger.log("Creating filter to Hide list of Unique dates except 5/14/2020...");
     targetSheet.getRange('B:B').activate();
     targetSheet.getRange('B:B').createFilter();
  
     var criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(myDates)
     .build();
     targetSheet.getFilter().setColumnFilterCriteria(2, criteria);
      return myDates;
      Logger.log("Filtering and Revealing rows with Date 5/14/2020 Complete...");
}

function checkFilter(sheet) {
  var targetSheet = sheet;
  var filter = targetSheet.getFilter();
  
  targetSheet.getRange('B:B').activate();
  
  if (filter !== null) {
    Logger.log("Filter on Column B found, removing filters...");
    //targetSheet.getRange('B:B').activate();
    targetSheet.getFilter().remove();
    return;
  }
  else{
   Logger.log("No Filters found."); 
  }

};

function ridDouble(sheet){
  var targetSheet = sheet;
  
  // Make Dates Format as Plain Text
     targetSheet.getRange('B:B').activate();
     targetSheet.getActiveRangeList().setNumberFormat('@');
  
  var values = targetSheet.getDataRange().getValues(); // Gets all Values in sheet
  var allDates = []; // Holds dates which continue duplicates
  var cleanDates = []; // Holds dates after filtering out duplicates
  
  Logger.log("Getting all unique dates...");
  for(var i=14;i<values.length;++i){
    var cell = values[i][1] ; // i is the index of the row, the 1 is the index which represents column B
    allDates.push(cell)    
  }
  
  allDates.sort();
  for(var n=0;n<allDates.length;n++){
   var Val1 = allDates[n]; // assign first value in list
    var cnt = n+1;
   var Val2 = allDates[cnt]; // assign second value in list
   
    if(Val1 == Val2){ // compare values, duplicates get thrown out
      //Logger.log(Val1+" Same Dates");
    }
    else{
     //Logger.log(Val1+" Added to List");
      cleanDates.push(Val1); // Unique dates gets pushed to clean list
    }
  }
  Logger.log("Unique dates list complete...");
  return cleanDates; // return clean dates
}


function hideFormula(sheet,hideDates){
     var targetSheet = sheet;
     var hide = hideDates;
     //Logger.log("Dates to Hide: "+ hide);
  
     var criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(hide)
     .build();
     
     targetSheet.getRange('B:B').activate();
     targetSheet.getFilter().setColumnFilterCriteria(2, criteria);

     Logger.log("Hiding Date 5/14/2020 Complete...");
}


function copyNetworkSeries(Test,Sheet7,rowNum){
  // Copies Formats Values of column's Network and Series from test and Paste them to appropriate columns in sheet7.
  
  // Declaring sheets as variables
  var test = Test;
  var sheet7 = Sheet7;
  var lastRow = rowNum;
  var safePlace = "A3"; // Home Cell Location for Network & Series
  
  // Cell Poisiton where we wish to begin Copying data
  test.getRange('A3').activate();
  
  var currentCell = test.getCurrentCell();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate(); //Selecting all rows with data in Column
  currentCell.activateAsCurrentCell();
  
  var range1 = "A3:B"+lastRow;
  test.getRange(range1).setBackground('#00ff00');
  test.getRange(range1).activate();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate(); //Selecting All columns associated with range
  currentCell.activateAsCurrentCell();
  
  sheet7.getRange('A3').activate(); // Cell Position of where we wish to begin Pasting on Sheet7
  var range2 = "test!A3:B"+lastRow;
  sheet7.getRange(range2).copyTo(sheet7.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);  
   Logger.log("Network and Series Columns complete.");
}

function copyCmcMpx(Test,Sheet7,rowNum){
  // Copies Formats and Copies timestamps of CMC-C and MPX from test and Paste them to appropriate columns in sheet7.
  
  // Declaring sheets as variables
  var test = Test;
  var sheet7 = Sheet7;
  var lastRow = rowNum;
  var safePlace = "H3"; // Home Cell location for CMC-C & MPX
  
  // Cell Poisiton where we wish to begin Copying data
  test.getRange('H3').activate();

  var currentCell = test.getCurrentCell();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate(); //Selecting all rows with data in Column
  currentCell.activateAsCurrentCell();
  
  var range1 = "H3:I"+lastRow;
  test.getRange(range1).setBackground('#00ff00');
  test.getRange(range1).activate();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate(); //Selecting All columns associated with range
  currentCell.activateAsCurrentCell();
  
  sheet7.getRange('C3').activate(); // Cell Position of where we wish to begin Pasting on Sheet7
  var range2 = "test!H3:I"+lastRow;
  sheet7.getRange(range2).copyTo(sheet7.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);  
   Logger.log("CMC-C and MPX Syndication Columns complete.");
}

function copyDishDirect(Test,Sheet7,rowNum){
  // Copies Formats and Copies timestamps of Dish, Direct-TV, DirecTV-NBC@VOD50-C from test and Paste them to appropriate columns in sheet7.
  
  // Declaring sheets as variables
  var test = Test;
  var sheet7 = Sheet7;
  var lastRow = rowNum;
  var safePlace = "K3"; // Home Cell location for Dish, Direct-TV, DirecTV-NBC@VOD50-C
  
  test.getRange('K3').activate();

  var currentCell = test.getCurrentCell();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate(); //Selecting all rows with data in Column
  currentCell.activateAsCurrentCell();
  
  var range1 = "K3:M"+lastRow;
  test.getRange(range1).setBackground('#00ff00');
  test.getRange(range1).activate();
  test.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate(); //Selecting All columns associated with range
  currentCell.activateAsCurrentCell();
  
  sheet7.getRange('E3').activate(); // Cell Position of where we wish to begin Pasting on Sheet7
  var range2 = "test!K3:M"+lastRow;
  sheet7.getRange(range2).copyTo(sheet7.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);  
  test.setActiveSelection(safePlace); 
   Logger.log("Dish and DirectTV Syndication Columns complete.");
}

function chkEmpty(sheet,col,lastRow){
  // Look for Empty Cell's and Add N/A's
  // Format Cell's to gray
  
  // Declaring sheets as variables
  var target = sheet;
  var targetCol = col;
  var lastR = lastRow;
  
  var count = 3
  
  // Looping through each row starting at row 2
  for (var count=3; count<=lastR; count++){
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
// N/A's Inserted
}

function autoResize(sheet){
  // For sheet7
  // Selecting all rows and columns and auto resizing the Column width 
  // Allign text to center 
  
  var target = sheet;
  target.getRange(1, 1, target.getMaxRows(), target.getMaxColumns()).activate(); // Start at top-left corner, selecting all Rows and Columns with data in it.
  target.autoResizeColumns(1, 26); // Resizing Column Width
  
  var all = target.getRange(1, 1, target.getMaxRows(), target.getMaxColumns()).activate(); // Start at top-left corner, selecting all Rows and Columns with data in it.
  all.setHorizontalAlignment("center"); // Alligning text to center
  
  target.getRange('A3').activate(); // Unselect all columns and rows by selecting cell A2
    Logger.log("Text alignment and Column Resize complete.");
}


function createVOD(Test,Sheet6,rowNum,Date){

  // Declaring variables
  var test = Test;
  var sheet6 = Sheet6;
  var lastRow = rowNum;
  var date = Date;
  
  // An Array to hold the 12 different Platform Devices
  var platforms = ["Android","AppleTV 3","AppleTV 4","Desktop","Directv","Dish","FireTV","iOS","Roku","Spectrum","X1","Xbox One"];
  
  Logger.log("Copying each asset into VODdoc & Assigning platform tags.");
  for(var startRow=3;startRow<=lastRow;startRow++){ // Loop - Starting at Row 2 in test, As long as Row is less than The Last Row Number in test.

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
  // Resize Columns and text-align center in VODdoc (sheet6)
  var target = Sheet6;
  var firstEmptyRow = FirstEmptyRow; // First Empty Row Number in sheet6
  var TotalRowsCreated = TotalRows;  // Total Rows Created
  var lastRow = firstEmptyRow+TotalRows;
  
  // center text
  var all = target.getRange("A"+firstEmptyRow+":F"+lastRow).activate();
  all.setHorizontalAlignment("center"); // Alligning text to center
  target.getRange('A'+firstEmptyRow).activate();  
  // resize column
  target.getRange(1, 1, target.getMaxRows(), target.getMaxColumns()).activate(); // Start at top-left corner, selecting all Rows and Columns with data in it.
  target.autoResizeColumns(1,7); // Resizing Column 1-7
  
  target.getRange('A'+firstEmptyRow).activate(); // Unselect all columns and rows by selecting FirstEmptyRow
    Logger.log("Text alignment and Column Resize complete.");
}

function Brand6(assets,destination,source,EmptyRow){
  var test = assets;
  var sheet6 = destination;
  var brand6 = source;
  var firstEmptyRow = EmptyRow;
  
  for(var i=1;i<=3;i++){
    //Logger.log("Starting Value: "+ firstEmptyRow);
    var network = sheet6.getRange("C"+firstEmptyRow).getValue();
    
    switch(network){
    case "BRAVO":
      brand6.getRange('C1').activate();
      var currentCell = brand6.getCurrentCell();
      brand6.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      sheet6.getRange("H"+firstEmptyRow).activate();
      sheet6.getRange('Brand6!C1:C12').copyTo(sheet6.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        firstEmptyRow = firstEmptyRow+12;
        break;
    case "CNBC":
      brand6.getRange('C14').activate();
      var currentCell = brand6.getCurrentCell();
      brand6.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      sheet6.getRange("H"+firstEmptyRow).activate();
      sheet6.getRange('Brand6!C14:C25').copyTo(sheet6.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        firstEmptyRow = firstEmptyRow+12;
        break;
    case "MSNBC":
      brand6.getRange('C27').activate();
      var currentCell = brand6.getCurrentCell();
      brand6.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      sheet6.getRange("H"+firstEmptyRow).activate();
      sheet6.getRange('Brand6!C27:C38').copyTo(sheet6.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        firstEmptyRow = firstEmptyRow+12;
        break;
   
    case "NBC":
      brand6.getRange('C40').activate();
      var currentCell = brand6.getCurrentCell();
      brand6.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      sheet6.getRange("H"+firstEmptyRow).activate();
      sheet6.getRange('Brand6!C40:C51').copyTo(sheet6.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        firstEmptyRow = firstEmptyRow+12;
        break;
    case "NBC NEWS":
      brand6.getRange('C53').activate();
      var currentCell = brand6.getCurrentCell();
      brand6.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      sheet6.getRange("H"+firstEmptyRow).activate();
      sheet6.getRange('Brand6!C53:C64').copyTo(sheet6.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        firstEmptyRow = firstEmptyRow+12;
        break;     
    case "Oxygen":
      brand6.getRange('C66').activate();
      var currentCell = brand6.getCurrentCell();
      brand6.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      sheet6.getRange("H"+firstEmptyRow).activate();
      sheet6.getRange('Brand6!C66:C77').copyTo(sheet6.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        firstEmptyRow = firstEmptyRow+12;
        break;
      
    case "E!":
      brand6.getRange('C79').activate();
      var currentCell = brand6.getCurrentCell();
      brand6.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      sheet6.getRange("H"+firstEmptyRow).activate();
      sheet6.getRange('Brand6!C79:C90').copyTo(sheet6.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        firstEmptyRow = firstEmptyRow+12;
        break;
    case "Universo":
      brand6.getRange('C92').activate();
      var currentCell = brand6.getCurrentCell();
      brand6.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      sheet6.getRange("H"+firstEmptyRow).activate();
      sheet6.getRange('Brand6!C92:C103').copyTo(sheet6.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        firstEmptyRow = firstEmptyRow+12;
        break;  
    case "Telemundo":
      brand6.getRange('C105').activate();
      var currentCell = brand6.getCurrentCell();
      brand6.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      sheet6.getRange("H"+firstEmptyRow).activate();
      sheet6.getRange('Brand6!C105:C116').copyTo(sheet6.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        firstEmptyRow = firstEmptyRow+12;
        break;
      
    case "USA":
      brand6.getRange('C118').activate();
      var currentCell = brand6.getCurrentCell();
      brand6.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      sheet6.getRange("H"+firstEmptyRow).activate();
      sheet6.getRange('Brand6!C118:C129').copyTo(sheet6.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        firstEmptyRow = firstEmptyRow+12;
        break;
    case "Golf":
      brand6.getRange('C131').activate();
      var currentCell = brand6.getCurrentCell();
      brand6.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      sheet6.getRange("H"+firstEmptyRow).activate();
      sheet6.getRange('Brand6!C131:C142').copyTo(sheet6.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        firstEmptyRow = firstEmptyRow+12;
        break;
    case "SyFy":
      brand6.getRange('C144').activate();
      var currentCell = brand6.getCurrentCell();
      brand6.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      sheet6.getRange("H"+firstEmptyRow).activate();
      sheet6.getRange('Brand6!C144:C155').copyTo(sheet6.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        firstEmptyRow = firstEmptyRow+12;
        break;
      
    case "Universal Kids":
      brand6.getRange('C157').activate();
      var currentCell = brand6.getCurrentCell();
      brand6.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
      currentCell.activateAsCurrentCell();
      sheet6.getRange("H"+firstEmptyRow).activate();
      sheet6.getRange('Brand6!C157:C168').copyTo(sheet6.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        firstEmptyRow = firstEmptyRow+12;
        break;
    default :
        break;
    }
  }
   Logger.log("Brand 6's Inserted...");
}

function airDate(){
  // Prompt User to enter Air Date
  Logger.log("Getting User input for AirDate");
  var ui = SpreadsheetApp.getUi();
  var date = ui.prompt("Please Enter the Air Date\nM/dd/YYYY").getResponseText(); // Store UserInput as Variable

  Logger.log("Creating VODdoc for "+date);
  // Return Air Date Value
  return date;  
}




