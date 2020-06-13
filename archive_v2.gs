// This Script Admin,Owner and Developer is Haris Nasir.
// Any and all changes must be approved by Haris Nasir due to the complex build of this script.
// Thursday 06/13/2020 15:43 ET

// ** Changes Notes **
// Make function to create TodaySheetName from todays date
// Make funtion getSpreadsheetID flexiable to get id of any URL
//  * This requires our initial user input of Destination URL to be moved into startArchive before Var SSID = getSpreadsheetID();

// Archiving Assets once all Syndication Timestamps are complete
// We must save todays assets from Builder into the Spreadsheet correspondent to Current Month.
// EX: "C3 Checks JUNE 2020"
//      This doc has a tab (sheet) for each day; asset's metadata and syndication timestamps are saved here.
//      If you notice, it is the exact same data we see in "Builder", without the build buttons in row 1.
// This Spreadsheet Doc will serve as a archive for the entire months assets, a new sheet for each day.
// Therefore a new Spreadsheet must be made every month to archive the month's assets.

// To Begin, We must ask the User for the URL of the Spreadsheet Doc of This month.
// From the URL we will be fetching the Spreadsheet ID.
// The Spreadsheet ID will gain us access to the data inside of it.
// We can create sheet's and manipulate the sheet data inside of it, in our case, save today's asset's as an archive.
// We must First Get a list of all the sheets inside that Spreadsheet,
//     * Then check if a spreadsheet with today's date as the name already exists.
//     * If yes, Save todays assets from "Builder" Range 'A3:M+Lastrow' into that sheet.
//     * If no, Create sheet with todays date as Name + Save todays assets from "Builder" Range 'A3:M+Lastrow' into that sheet.

function startArchive(){
  
  // Get the name for Today's Archive sheet
     var TodaySheetName = dateAsname();
  
  // Get Spreadsheet ID
     var ui = SpreadsheetApp.getUi();
     var url = ui.prompt("URL of this Month's Asset DOC").getResponseText(); // Store UserInput as Variable
     var SSID = getSpreadsheetID(url); // Archive Spreadsheet's Spreadsheet-ID
  
  // Check if Todays Sheet Name exists in Spreadsheet.
     var Exists = chkExists(SSID,TodaySheetName);
     Logger.log("TodaySheetName Already Exists? "+Exists);
  
  // If Exists == true 
  // Copy assets into todays sheet
  // else if Exists == false
  // Create a sheet with todays date as name
  // Then copy assets into todays sheet
     if(Exists == true){
       archiveAssets(SSID,TodaySheetName); 
       // Auto resize Column Width and Align Text Center in Destination
       resizeCenter(SSID,TodaySheetName);
     }
     else if (Exists == false){
      // createArchiveSheet(SSID,TodaySheetName);
      // archiveAssets(SSID,TodaySheetName);
       // Auto resize Column Width and Align Text Center in Destination
      // resizeCenter(SSID,TodaySheetName);
       Logger.log("We need to create Sheet");
     }  
}

function getSpreadsheetID(url) {
  
  // We want to get the Spreadsheet ID for this Months C3 Assets Doc to archive our assets into.
  // Google Script's has built in function to get sheetID's but not the Spreadsheet ID.
  // This function Splits through the URL and finds the Spreadsheet ID.
  
  // DELETE var ui = SpreadsheetApp.getUi();
  // DELETE var url = ui.prompt("URL of this Month's Asset DOC").getResponseText(); // Store UserInput as Variable
  var url = url;
  
  var firstPart = url.toString().split("/d/");
  var secondPart = firstPart[1].split("/edit");
  var SpreadsheetID = secondPart[0];
  
  Logger.log("Spreadsheet ID: "+SpreadsheetID);  
  return SpreadsheetID;
}

function chkExists(spreadsheetID,todaysheetname) {
  var sheetname = todaysheetname; // Sheet name we want to know if already exists.
  
  var ssid = spreadsheetID; // spreadsheet ID of where we will get all sheet names from.
  var TargetSpreadsheet = SpreadsheetApp.openById(ssid); // making Variable for that Spreadsheet to get all sheet names from.
  
  var exists = false;     // By default, sheetname will be flagged as False, Does not exist.          
  
  var out = []; // Will hold list of Sheet names that are in Spreadsheet
  
  var sheets = TargetSpreadsheet.getSheets(); // Gets all the sheet's, as an Array
  for (var i=0 ; i<sheets.length ; i++) out.push( [ sheets[i].getName() ] ) // Gets the Name of the sheet's and pushes into Array "out".
   //Logger.log("List of SheetNames "+out); 
  
  // Check if sheetname exists in List of sheets.
  // Remember, When we get sheet names from Target Spreadsheet, it returns each sheet name as Array.
  // Our Array "out" is now a Two dimensional array,
  // EX: out = [["6/2/20"], ["6/3/20"], ["6/4/20"]];
  
  for(i=0;i< out.length;i++){  // For each Array in Array:out
    var val = out[i][0]; // index 0 will always be Name of Sheet.
    Logger.log("Their Name: "+val+" Our Name: "+sheetname);
   
    if (val == sheetname){  // compare sheetname from Doc with the name we are looking for
      exists = true; // if exist, flag "exists" as true and end loop
      break;
    }
    else{
     exists = false; // if it does not exsit, flag as false; 
    }
  }
  
  return exists;

}

function archiveAssets(ssid,todaysheetname){
  // Copies over all Assets from Builder to sheet in Target Spreadsheet to Archive today's assets.
  var ssid = ssid; // Spreadsheet ID of where target sheet exists
  var todaysheetname = todaysheetname; // Name of the sheet to paste into.
  var destinationSS = SpreadsheetApp.openById(ssid); // The Spreadsheet to Paste into as a Variable 
  var sheet = destinationSS.getSheetByName(todaysheetname); // The Sheet in Spreadsheet to Paste into as a Variable
  
  var source = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Builder"); // The sheet to copy asset's from.
  var lastRowNum = source.getLastRow(); // The Row number for the last asset in list.

  for (var row = 3;row<=lastRowNum;row++){   // start at row 3 and increment till the last row
      var myRow = source.getRange("A"+row+":E"+row).getValues(); // Loops through each row, creating Array within Array [[MSNBC,Series,Title,AssetID,Timestamp]]
      myRow.sort();
    for (var x = 0;x<myRow.length;x++){  // x is to loop through each subArray(1) in Array, Bascially getting that one Array that holds our values.
      for (var i=0;i<myRow[x].length;i++){ // i is to loop through each value in subArray
      var Val = myRow[x][i];  // same as myRow[0][i] 
        sheet.getRange(row-1,i+1).setValue(Val);  // copy to destination starting at row-1 (Row 2), Column: i+1
      }
    }
  } 
}

function resizeCenter(ssid,todaysheetname){
  // Resize Columns and text-align center 
  var ssid = ssid;
  var todaysheetname = todaysheetname;
  
  var destinationSS = SpreadsheetApp.openById(ssid); // The Spreadsheet to Paste into as a Variable 
  var sheet = destinationSS.getSheetByName(todaysheetname); // The Sheet in Spreadsheet to Paste into as a Variable
  
  var lastRow = sheet.getLastRow();
  
  // center text
  var all = sheet.getRange("A1:M"+lastRow).activate();
  all.setHorizontalAlignment("center"); // Alligning text to center  
  // resize column
  sheet.autoResizeColumns(1,13); // Resizing Column 1-7
  
  sheet.getRange("A1").activate(); // Unselect all columns and rows by selecting FirstEmptyRow
    Logger.log("Text alignment and Column Resize complete.");
}

function createArchiveSheet(ssid,todaysheetname){
  var ssid = ssid;
  var todaysheetname = todaysheetname;
  
  var destinationSS = SpreadsheetApp.openById(ssid); // The Spreadsheet to Paste into as a Variable 
  var sheet = destinationSS.getSheetByName(todaysheetname); // The Sheet in Spreadsheet to Paste into as a Variable
  
  // We want to copy Entire Row 2 from "Builder"  which is the header row for Row 1 in Archive sheet
  // To do that we need the Spreadsheet ID of Active Spreadsheet
  var sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var url = sourceSpreadsheet.getUrl(); // Gets URL of "Builder" Spreadsheet
  var sourceSSID = getSpreadsheetID(url); // SpreadsheetID of Active Spreadsheet
  
  Logger.log("Source SS ID: "+sourceSSID);
  
  // Create the sheet with the date name
  sheet.insertSheet(todaysheetname);
   
}

function dateAsname(){
  // Objective is to get Date in format m/d/yy
  // ex; 6/6/20 & 6/10/20
  
  // Create a new Date object
     var date = new Date();
  
  // Get the Various parts as simplified digit values, no decimal points or leading zeros
     var month = date.getMonth()+1;month = month.toString().slice(-1);
     var day = date.getDate(); if(day.toString().length==1){var day = day.toString().slice(-1);}
     var year = date.getFullYear().toString().slice(-2);

        // Logger.log("\nMonth: "+month+"\nDay: "+day+"\nYear: "+year);

     // Return the Date that will be used as the name for the sheet
     var d = month+"/"+day+"/"+year;
        //Logger.log(d);
   return d
}



/*




// ** Start of Build Syndication **

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
  
  // Autoresize All column width's for Sheet7
     Logger.log("Resizing Columns to fit text's...");
     Utilities.sleep(100); // Pause for 100 milliseconds to avoid sheet freezing before adding border
     autoResize(sheet7,lastRow);

} // End BuildSyndication()

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
} // End of chkEmpty

function autoResize(sheet,lastRow){
  // For sheet7
  // Selecting all rows and columns and auto resizing the Column width 
  // Allign text to center 
  
  var target = sheet;
  var endRow = lastRow;
  
  // Resize Columns to fit Text
  target.getRange(1, 1, target.getMaxRows(), target.getMaxColumns()).activate(); // Start at top-left corner, selecting all Rows and Columns with data in it.
  target.autoResizeColumns(1, 26); // Resizing Column Width
  
  // Align Text - Center
  var all = target.getRange(1, 1, target.getMaxRows(), target.getMaxColumns()).activate(); // Start at top-left corner, selecting all Rows and Columns with data in it.
  all.setHorizontalAlignment("center"); // Alligning text to center
  
  // Add Border all around to display grid view
  target.getRange("A3:G"+endRow).activate();
  target.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  target.getRange('A3').activate(); // Unselect all columns and rows by selecting cell A2
    Logger.log("Text alignment and Column Resize complete.");
}

// ** End of Building Syndication **
//--------------------------------------------------------------------------------------------------------------
// ** Start of Build VODdoc **

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
     Utilities.sleep(100); // Pause for 100 milliseconds to avoid sheet freezing.
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
} // End of showProtected

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
} // End of checkFilter

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
} // End of ridDouble

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
} // End createVOD

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

// ** End of Building VODdoc **
//--------------------------------------------------------------------------------------------------------------
// ** Start of MagicMaker **

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
     Utilities.sleep(100); // Pause for 100 milliseconds to avoid sheet freezing before adding border
     autoResize(sheet7,lastRow);

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
     Utilities.sleep(100); // Pause for 100 milliseconds to avoid sheet freezing.
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

// ** End of MagicMaker **

// ** COPY PASTE Column Functions for Syndication Sheet
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

// ** Brand 6's for VODdoc
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
} // End of Brand 6's

function airDate(){
  // Prompt User to enter Air Date
  Logger.log("Getting User input for AirDate");
  var ui = SpreadsheetApp.getUi();
  var date = ui.prompt("Please Enter the Air Date\nM/dd/YYYY").getResponseText(); // Store UserInput as Variable

  Logger.log("Creating VODdoc for "+date);
  // Return Air Date Value
  return date;  
}

*/
