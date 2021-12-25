// Have upto 51 rows in Builder and add a Protection Range to B1:G51, no one can edit but me, they can only edit column A becuase they need to add red background to the rows they want to ignore. 
// Activate "Drive API". Go to in the script editor to: Resources > Advanced Google services ... > Drive API > On


// // use chkText method (at the end of script) before switch method for syndication coloring.
// // Updated 11/24/21 14:15
// // AirDate changed in Spreadssheet:"DTS C3-Checlist" Sheet:"MasterC3" Cell:C3
// // Will update the assets in Spreadssheet:"DTS C3-Checlist" Sheet:"Prep C3 VOD Checks" 
// // And in our Automation Doc Spreadsheet:"VOD Checks (Automated)2021" Sheet:"Builder" & Sheet:"Prep C3 VOD Checks" 

// // BUG : creating extra rows in C3Checks tab (as many asset rows copied in + 1)  ...WHY?
// // Line 428 when doing brand 6 not looping to next network, stays stuck on first network name but loops correct amount of times. 
// //     Trying to remove white spaces from network name, maybe remove white spaces within PrepC3 tab


// Make documentation of set up process

// // DTS C3-Checklist
// // TRAFICREPORT: Must have the correct date range in cell B2
// // This will update assets and allow us to pull AssetID's
// // MasterC3: Put in the Airdate 
// // This will update assets in MasterC3 AND in our "Builder" sheet
// // Go back to our Builder Doc to fill out syndication Timestamps

// // New Workflow
// // 1. Update the Airdate in "DTS C3-Checklist" : "MasterC3" : Cell C3
// //   * Cell Values to keep updated for accurate asset Listing
// //     - DTS C3-Checlist : TRAFFICREPORT : Cell B2  should have the current weeks range
// //       > To get the the options of Range values open Spreadsheet "Traffic Report" can get the Tab names on bottom.
// //       > https://docs.google.com/spreadsheets/d/13PWvIyuvOYsKkzqPfG5gJxwkbQ3aq9f7W-p328VJCYg/edit#gid=1320183517
// //     - DTS C3-Checlist : TELEMUNDO : Cell B2  should have current Month Year 
// //       > ie "September 2020"
// // 2. Once MasterC3 Airdate is updated you should see list below the airdate Update with new assets.
// //   * Assets Automatically update in our "VOD Checks (Automated) 2021" : "Builder" & "Prep C3 VOD Checks"
// //     > In "Prep C3 VOD Checks" our Checks Rough Draft should automatically populate In Columns K:P
// //     > If you happen to find any issues in asset metadata, it's best to first troubleshoot in "DTS C3-Checklist" : "MasterC3"
// //     > All of our data originates from MasterC3 so fixing typos in our "Builder" or "Prep C3 VOD Checks" will not make legitimate changes.
// // 3. We are Ready to Filter out old dates from C3Checks and paste in our new assets list.
// //   * Click the "Build VOD Doc" Button

// Global Sheet Values
  // 'DTS-C3 Checklist' Global Sheet Variables
  var SheetID = "1tHT_n9l-IiwqyWZFTnaULsZ_W8wOAbzJQpW1TLLRuBE" ;
  var MasterC3 = SpreadsheetApp.openById(SheetID).getSheetByName("MasterC3");
  var Prep = SpreadsheetApp.openById(SheetID).getSheetByName("Prep C3 VOD Checks");

  // 'VOD CHECKS(Automated)' Global Sheet Variables
  var C3Checks = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("C3Checks");
  var PrepC3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Prep C3 VOD Checks");
  var Brand6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Brand6");
  var SyndicationNotes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SyndicationNotes");
  var Builder = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Builder");
  var Log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LOGS");

  // Daily Asset Archiving related Global Variables
  let DailyArchiveFolderId = "1nYFn71YCdRdF8kWiyPYN09Q8dDv6g_lY";
  let DailyArchiveFolder = DriveApp.getFolderById(DailyArchiveFolderId);

// // STARTER() - Starts the building of VOD doc process
// // SYNDICATOR() - Starts the building of Syndication tab
// //------------------------------------------------------------------------
// // Initilizer - gets the process going to create the VOD Doc. 

function confirmC3(){
  // Alert user to confirm moving forward with script.
  // alert(title, prompt, buttons)

  // Did user mark bkgrd color red for Network column of row's they wish to ignore?
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Do you wish to continue?', 'For rows you wish to leave out of the C3 Checks,\nYou must make the \'Network\' column\'s background color Red (#ff0000).\n\nClick Yes\nIf you have already color coded Network columns OR you just wish to continue building C3 Checks Doc.\n\nClick No\nIf you have to stop the script to make a row\'s network Background color Red (#ff0000).', ui.ButtonSet.YES_NO);

  // Process the user's response.
  // If Yes, continue running next step of script.
  // If No, exit script, allow user to make edits.
  // else default exit.
  if (response == ui.Button.NO) {
    // Networks are NOT color coded, prompt user to do so.
    let msg = "Script Terminated, Please make your edits.";
    updateMsg(msg);
    updateLog(msg);
    return;
  }else if(response == ui.Button.YES){
    // Networks are color coded continue to build C3
    let msg = "Building C3 Doc...stand by.";
    updateMsg(msg);
    updateLog(msg);

    // Begin building c3 doc
    STARTHERE();
  }else{
    let msg = "<exited>";
    updateMsg(msg);
    updateLog(msg);
    return;
  }
} // end userPermission()

  /* CONSTRUCTOR */
  function C3Info(){
    /*----- First step is to gather data -----*/
    // The Airdate in cell B1  
    this.getAirDate = Builder.getRange('B1').getDisplayValue();
    
    //Last Row with data in Builder
    this.lastRowBuilder = getLastAssetRow(Builder);

    // Row's to ignore in builder.(if any) : parameter lastRowBuilder : returns array with range's in column A ie; [A2,A12] or empty array if none found.
    this.rowsToIgnore = findIgnores(this.lastRowBuilder);

    // Total Number of Assets
   this.totalAssets = (this.lastRowBuilder - 2) - this.rowsToIgnore.length;

    // If the Last Row with data is row 1, that is the heading so the rest of the sheet is clear.
   this.lastRowPrep = getLastAssetRow(PrepC3);

   // If the Last Row with data is row 1, that is the heading so the rest of the sheet is clear.
  this.lastRowSyndication = getLastAssetRow(SyndicationNotes);
   
  //let msg = "Constructor:\n Air Date: "+this.getAirDate+" Last Row Builder: "+this.lastRowBuilder+"\nRows To Ignore: "+this.rowsToIgnore+" Total Assets: "+this.totalAssets+"\nLast Row PrepC3: "+this.lastRowPrep;
 // updateLog(msg);
  }

// Initilizer for creating C3 Doc.
function STARTHERE(){
  const c3Data = new C3Info();

  // Step 1: Prepare the PrepC3 tab
  if (c3Data.lastRowPrep == 1){
    // row 1 is heading row so theres no data in PrepC3
    // no nead to clear sheet.
    Logger.log("PrepC3 Last Row:"+c3Data.lastRowPrep + ": The sheet does not have data, already clear...");
  }else if(c3Data.lastRowPrep > 1){
    // the last row with data is not the heading (row 1)
    Logger.log("PrepC3 Last Row:"+c3Data.lastRowPrep + ": The sheet has data, clearing data now...");
    // clear prepC3 sheet
    clearPrepC3(c3Data.lastRowPrep);
  }else{
    // last row with data is < 1 which means thats an error
    // do nothing
    let msg = "Error when searching for last row in PrepC3.";
    updateMsg(msg);
    updateLog(msg);
    return;
  }

  // Step 2: Copy over assets from Builder to PrepC3
  if(c3Data.rowsToIgnore.length == 0){
    // if array is empty, no rows to ignore, Copy over Builder assets as normal
    //Logger.log("No Rows to ignore in Builder, copying from Builder Normally.");
    let msg = "No Rows to ignore in Builder, copying from Builder Normally.";
    updateLog(msg);
    updateLog("Airdate: "+c3Data.getAirDate+" Total # of Assets: "+c3Data.totalAssets);
    copyNormalPrepC3(c3Data.getAirDate, c3Data.totalAssets);
  }else if(c3Data.rowsToIgnore.length >= 1){
    // if array has atleast 1 or more rows to ignore, copy over assets from Builder but ignore these rows.
    //Logger.log("Some rows to ignore in Builder,copying from Builder without ignores.");
    let msg = "Some rows to ignore in Builder,copying from Builder without ignores.";
    updateLog(msg);
    updateLog("Airdate: "+c3Data.getAirDate+" Last Row # Builder: "+c3Data.lastRowBuilder+" The Rows To Ignore in Builder: "+c3Data.rowsToIgnore);
    copyIgnoresPrepC3(c3Data.getAirDate, c3Data.lastRowBuilder, c3Data.rowsToIgnore);
  }else{
    // unkown data, do nothing
    let msg = "Error when searching if there are Rows to Ignore.";
    updateMsg(msg);
    updateLog(msg);
  }

  // Step 3: trim whitespaces from network column (Column C) in PrepC3
  cleanNetworks(c3Data.totalAssets);

  // Step 4: Get a array of all unique dates in Column B of C3Checks tab
  let datesList = getUniqueDates();

  // Step 5: Check if airdate already exists in C3CHecks. if it does, terminate all steps beyond this point and alert user, else continue
  Logger.log("Checking if AirDate already exists in C3Checks.");
  let existDateC3Checks = checkDateC3Checks(c3Data.getAirDate, datesList); // returns true if airdate already exists in C3Checks else returns false

  if( existDateC3Checks != true){
    // AirDate does not already exist in C3Checks so we can continue creating the C3Checks doc.
    // Step 6: Prepare C3Checks, set filter to show only rows with dates in array datesList;
    hideDates(datesList); 

    // Step 7: add 1 extra empty row at the bottom of the sheet , function takes in any sheet name as parameter.
    // returns row # of new empty row we will start pasting new assets into. 
    let newRowC3 = addExtraRow(C3Checks);

    // Step 8: Ready to copy asset list from PrepC3 into C3Checks
    // returns row # of last asset in C3 Checks that we pasted
    let newLastRowC3 = CopyFromPrepC3(c3Data.totalAssets, newRowC3);

    // Step 9:  Change Format of Column B to Date
    // returns true or false depending on success of function
    let formatColStatus = formatColumn_Date(newRowC3, newLastRowC3);

    // Step 10: If format for column A&B was successfull then do Step 8
    let emptyRowsToDelete = false;
    if(formatColStatus == true){
      // succesfully formatted column A&B so now check if there are any extra empty rows that might have been created by script
      // return true if we did have any empty rows and deletes them
      // return false if we did not have any empty rows to delete.
      emptyRowsToDelete = removeEmptyRows(newRowC3);
    }else if(formatColStatus == false){
      // formatting column A&B was unsuccessful so we will not have any extra empty rows to delete
      let msg = "Did not add formulas to Column A&B in C3Checks.";
      updateLog(msg);
    }else{
      // unkown error
    }

    // if emptyRowsToDelete is true, then our row number for last asset might have change, get new lastRowC3 value
    if(emptyRowsToDelete == true){
      Logger.log("true, we did delete row/rows!");
      newLastRowC3 = C3Checks.getLastRow();
    }else if(emptyRowsToDelete == false){
      // did not have to delete empty rows so newLastRowC3 is still the same
    }else{
      // Unkown || No return value from removeEmptyRows()
    }

    // Step 11: Add Formulas from Column H13:M13 to our new rows, These formulas are the data validation and status1,2,3 timestamp column formulas.
   addStatus_Columns(newRowC3,newLastRowC3);

    // Step 12: C3Checks Doc is created, hide the formula rows 1:13 in filterview
    datesList.push(['5/14/2020']);
    hideDates(datesList);
    let msg = "Formula Rows H13:M13 are now hidden in C3Checks.";
    updateLog(msg);

    // Step 13: Add Brand 6's
    insert_Brand6(c3Data.totalAssets, newRowC3);
    msg = "C3 Doc is Ready !!";
    updateMsg(msg);
    updateLog(msg);
  }else{
    // airDate did exist in C3Checks already so we will alert user and avoid recreating C3Checks Doc with same assets.
    Logger.log("Not Creating C3Checks becuase airDate already existed.");
    msg = "AirDate: "+c3Data.getAirDate+" Already Exists in C3Checks.";
    updateMsg(msg);
    updateLog(msg);
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Script Terminated', "The AirDate: "+c3Data.getAirDate+" Already exists in C3Checks tab.", ui.ButtonSet.OK);
  }
} // end STARTHERE()

// Functions for creating C3 Doc
function updateMsg(msg){
  // Takes a message and updates cell D1 with current message 
  Builder.getRange('D1').setValue(msg);
  Builder.getRange('D1').activate();
} // end updateMsg()

function updateLog(msg){
  // @param msg  : the message to log into LOGS tab
  var day = new Date();
  var timestamp = Utilities.formatDate(day, 'America/New_York', "MM-dd-yyyy HH:mm:ss");
  var timerow = Log.getLastRow()+1;
  Log.getRange('A'+timerow).setValue(timestamp);
  Log.getRange('B'+timerow).setValue(msg);
} // end updateLog()

function getLastAssetRow(SheetName) {
  // Get the Row number of the last asset in a given SheetName then return the row number
  // Requires atleast 1 empty row at the bottom of the sheet
  var sheet = SheetName;
  var column = sheet.getRange('B:B');
  // get all data in one call (which will be in array format) stored in a array called values : 2D array
  var values = column.getValues(); 
  // starting at the first array, check the first index (0) 
  // if its not empty, move on to next array
  // When an array with first index empty is found - while loop will end 
  // ct will not increment which means ct will be the value of the last row that had data.
  var ct = 0;
  while ( values[ct][0] != "") {
    ct++;
  }
  return(ct);
} // end getLastAssetRow()

function findIgnores(lastRowBuilder){
  // @param: lastRowBuilder
  // Find the Rows in Builder that are background red in Network column.
  // return an array of the rows if any found
  // Checking for background color in Builder column A - if #ff0000 (red), we will ignore row when creating C3 Checks Doc.
  let ignoreRows = []; // initialize array of rows to ignore
  
  // Loop through each column starting at row 3 column A and check for background color Red's #ff0000.
  for(var i = 3; i <= lastRowBuilder; i++){
    let currentCell = "A"+i;
    let currentCellColor = Builder.getRange(currentCell).getBackground();
    
    if(currentCellColor == "#ff0000"){
      ignoreRows.push(currentCell);
    }else{
      // Do nothing
    }
  }
  return ignoreRows;
} // end findIgnores()

function clearPrepC3(lastRowPrep){
  // @param : lastRowPrep
  // This function gets called only if we have to clear PrepC3 sheet, so we already know data exists under the header row.
  // Clear the entire range Column B2:F lastRowPrep
  PrepC3.getRange("B2:F"+lastRowPrep).setValue("");
} // end clearPrepC3()

function copyNormalPrepC3(airDate,totalAssets){
  // @param : airDate; the airdate to copy into Column B.
  // @param : totalAssets; total # assets we are copying over from Builder.
  let builderLastRow = totalAssets + 2;
  let currentRange = "A3:D"+builderLastRow;
  let newLastRowPrepC3 = totalAssets + 1;
  
  Builder.getRange(currentRange).copyTo(PrepC3.getRange("C2"), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  PrepC3.getRange("B2:B"+newLastRowPrepC3).setValue(airDate);
} // end copyNormalPrepC3()

function copyIgnoresPrepC3(airDate,lastRowBuilder,ignoreRows){
  // @param : airDate; the airdate to copy into Column B.
  // @param : lastRowBuilder; the last row# in Builder.
  // @param : array of row's we will be ignoring in Builder when copying over assets to PrepC3.
  let startRowBuilder = 3;
  let startRowPrepC3 = 2; // initilize empty row #, this will be incremented each time a row is copied over to PrepC3 from Builder.

  // Loop through each row in Builder starting at row startRowBuilder and Column A
  for(let i = startRowBuilder; i<= lastRowBuilder; i++){
    let startCell = "A"+i;
    // if startCell is in array ignoreRows then do nothing, ignore the row move to next.
    if(ignoreRows.includes(startCell)){
      // Logger.log(startCell+ " Exist in array");
      // If this cell is in array, we skip this row and don't copy it.
    }else{
      // copy this row from Builder to PrepC3 and increment startRowPrepC3 by 1
      // Logger.log("Copying "+startCell+" To PrepC3 row: "+startRowPrepC3);
      Builder.getRange(startCell+":D"+i).copyTo(PrepC3.getRange("C"+startRowPrepC3), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      startRowPrepC3++;
    }
  }
  let endRow = startRowPrepC3 - 1; // by the end of for-loop we have the row number of the next empty row, by subtracting 1, we get the row # of the last row we copied into. ( last row in PrepC3 with data)
  // Copy the airdate into range B2:B+endRow in PrepC3
  PrepC3.getRange("B2:B"+endRow).setValue(airDate);
} // end copyWithoutIgnoresPrepC3()

function cleanNetworks(totalAssets){
  // @param : totalAssets; the total # of assets we copied into PrepC3
  // Get each network value i Col C and clean the white spaces from beginning and end.
  let lastRowPrep = totalAssets + 1; // total # of assets plus heading row (1) will give us the last row in PrepC3
  let networks = PrepC3.getRange("C2:C"+lastRowPrep).getValues(); // 2D array [[NBC], [USA]]
  let cleanNetworks = []; // initilize 1D array clean networks. (clear of white spaces)

  // perform trim()
  for(let i = 0; i< networks.length; i++){
    let network = networks[i][0];
    //Logger.log(network);
    cleanNetworks.push(network.trim());
  }

  let pasteRow = 2; // start pasting back cleanNetworks beginning on row 2
  // Paste back cleanNetworks into Column C
  for(let i = 0; i<= cleanNetworks.length; i++){
    // i is the index of array cleanNetworks 
    PrepC3.getRange("C"+pasteRow).setValue(cleanNetworks[i]);
    pasteRow++;
  } 

  let msg = "cleaned whitespaces from Network column in PrepC3.";
  updateLog(msg);
} // end cleanNetworks()

function getUniqueDates(){
  // Get all Available Values in Column B, loop through and make array of unique dates/remove duplicates
  // return array cleanDates[]
  C3Checks.getRange('B:B').activate();
  C3Checks.getActiveRangeList().setNumberFormat('M/d/yyyy');
   // ---------- GET UNIQUE DATE VALUES ---------- //
  let focusCol = 1;
  let values = C3Checks.getRange("B14:B").getValues(); // Gets all Values in Column B
  let allDates = []; // Holds dates which contain duplicates
  let cleanDates = []; // Holds dates after filtering out duplicates
  // Custome format each date in Values so we can predict how to read the values ie; 5/14/2020
  for(let r=0; r<values.length; r++){
    let cell = values[r][0]; // values is a 2D array, r is the index to itterate through and we will always need the 0 of each r
    if(cell == ""){
      // cell is empty, don't try to format
    }else{
      let m = cell.getMonth() + 1;
      let d = cell.getDate() ;
      let y = cell.getFullYear() ;
      let tempDate = m + "/" + d + "/" + y;
      allDates.push(tempDate);
    }
  }
  // Start comparing date 1 with date 2 so we can make a list of unique dates (list of dates without duplicates)
  for(var n=0; n<allDates.length; n++){
    var Date1 = allDates[n]; // assign first value in list
    var cnt = n+1;
    var Date2 = allDates[cnt]; // assign second value in list
   
    // Compare values, if value is same, move on,
    // eventually value n and value cnt will be two different dates 
    // in which case we will save value n in array cleanDates[]
    if(Date1 == Date2){ 
      //Logger.log(Date1+" Same Dates");
      // If both values are same, skip, don't do anything
    }else{
      //Logger.log(Date1+" Added to List");
      // Date1 and Date2 are different, save Date1 into our cleanDate array.
      // In the First instance of both dates not being same, save the first date value as a unique value in cleanDates
      cleanDates.push(Date1); // Unique dates gets pushed to clean list
    }
  }
  //Logger.log("Unique dates list complete...");
  //Logger.log(cleanDates);
  return cleanDates;
} // end getUniqueDates()

function checkDateC3Checks(airDate, datesList){
  // @param : airDate : the airdate of todays assets
  // @param : datesList : an array of all unique airDates in C3Checks
  // check if todays airdate is in datesList, return true if it does, else return false.

  if(datesList.includes(airDate)){
    Logger.log("airDate already exists in C3Checks");
    return true;
  }else{
    Logger.log("AirDate not found in C3Checks");
    return false;
  }
}

function hideDates(datesList){
  /*
   NOW that the "Prep C3 VOD Checks" has the New Assets,
   We must prepare our "C3Checks" before pasting the assets in.
   We will start by Updating/Creating Filter View to show only 5/14/2020 in Column B
  */
  // Create/Update Filter View to
  // Hide all Values Except for 5/14/2020

  // set date format to column B to remove any strings
  let filter = C3Checks.getFilter();
 
  // Now we have an array cleanDates that holds all unique dates in column B of our C3Checks sheet.
  // We will use these dates to create a filter view of our sheet.
  /* ---------- CHECK FILTER , UPDATE Filter if already active or Create Filter if Filter not found ---------- */
  if (filter !== null) {
    // IF filter is NOT OFF aka filter IS ON.
    // IF FILTER Already exist, GET VALUES AND HIDE ALL EXCEPT 5/14/2020
    Logger.log("Filter on Column B found, Updating filter...");   
    //C3Checks.getRange('B:B').activate();
    // Add our array of dates into existing filter view.
    var criteria = SpreadsheetApp.newFilterCriteria()
    .setHiddenValues(datesList)
    .build();
    C3Checks.getFilter().setColumnFilterCriteria(2, criteria);
    let msg = "C3Checks already has filter, updating filter view...";
    updateLog(msg);
  }else{
    // Filter was not found so
    // CREATE FILTER TO HIDE OLD DATES
    Logger.log("No Filters found. Creating New Filter to Hide list of Unique dates except 5/14/2020..."); 
    C3Checks.getRange('B:B').createFilter();
    // Build Filter View
    var criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(datesList)
    .build();
    // Apply Filter View
    C3Checks.getFilter().setColumnFilterCriteria(2, criteria);
    // Log
    let msg = "C3Checks was not in filter view, created filter view...";
    updateLog(msg);
  }
} // end hideDates()

function addExtraRow(sheet){
  // currently using to add 1 extra empty row in c3Checks after hideDates().
  // However, Function will be written to take in any sheet name, get the lastrow and add one empty row below it.
  sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  let newEmptyRow = sheet.getLastRow() + 1;
  sheet.getRange("B"+newEmptyRow).activate();
  return newEmptyRow;
} // end addExtraRow()

function CopyFromPrepC3(totalAssets, newRowC3){
  // @param: totalAssets; total assets in for todays C3
  // @param: newRowC3; the row # the empty row where we will begin to paste new C3's
  // Copy rough draft from Our VOD Checks (Automated)2021 "Prep C3..." into "C3Checks"
  /*
    How many rows do we need to copy from K2-P? 
    ? = Number of Assets * 12
     Each asset has 12 rows (1 for each device).
     Muliplying the number of assets and adding the heading row number = will give us the last row to copy till
     This completes our ? in the range K2:P?  for  "VOD Check (Automated)2021 - Prep C3 VOD Checks
  */
  let lastRowToCopy = (totalAssets * 12) + 1; // the +1 is to account for the header row # so we accuratley aquire the row # of the Last Asset in our List
  //Logger.log("LastRowToCopy = "+lastRowToCopy);
  // Copying assets
  try{
    PrepC3.getRange("K2:P"+ lastRowToCopy).copyTo(C3Checks.getRange("B"+newRowC3), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  }catch(e){
    let msg = "Error when copying Assets into C3Checks: "+e;
    updateLog(msg);
  }
  let msg = "Copied all assets from PrepC3 to C3Checks.";
  updateLog(msg);
  Logger.log(msg);
  C3Checks.getRange("B"+newRowC3).activate();
  // get the last row # in C3 Checks
  return C3Checks.getLastRow();
} // end CopyFromPrepC3()

function formatColumn_Date(newRowC3, newLastRowC3) {
  // @param: newRowC3; the first row # we pasted new assets into in C3Checks
  // @param: newLastRowC3; the last row # we pasted new assets into in C3Checks
  // Format Column B of C3 Checks to date format so we can extra data for Column A in next step
  try{
    let i = newRowC3;
    while(i <= newLastRowC3){
      C3Checks.getRange("B"+i).setNumberFormat('M/d/yyyy');
      C3Checks.getRange("A"+i).setFormula('=TEXT(B'+newRowC3+',"MMMM")');
      i++;
    }
    Logger.log("Formatted column A&B for rows  "+newRowC3+" : "+newLastRowC3);
  }catch(e){
    // error trying to format
    Logger.log(e);
    return false;
  }
  // no errors when trying to format column B of C3 Checks
  return true;
} // end formatColumn_Date()

function removeEmptyRows(newRowC3){
  // Check each row, if empty cell, delete row becuase its an extra row added during Format of column B.
  // if you delete a row, increment to counter and return so we can update total Row count.
    C3Checks.getRange("B"+newRowC3).getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
    C3Checks.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  let trueLastRowC3 = C3Checks.getCurrentCell().getRow();
  let numRowsRemoved = 0; // initilize how many rows we delete
  let rowsToDelete = []; // which exact rows are we deleting. 
  let data = C3Checks.getRange("B"+newRowC3+":B"+trueLastRowC3).getValues();
  
  for(let i=0; i<data.length; i++){
    if(data[i] == "" || data[i] == null){
      // empty cell
      numRowsRemoved++;
      let deletedRow = newRowC3+i;
      rowsToDelete.push(deletedRow);
     // C3Checks.getRange("B"+deletedRow).setValue("Deleted");
    }else{
      // not a empty cell, do nothing 
    }
  }
  Logger.log("Checked Values for rows: "+newRowC3+":"+trueLastRowC3);
  Logger.log("How many rows deleted: "+numRowsRemoved+"\nWhich Rows Were Deleted: "+rowsToDelete);
  if(numRowsRemoved != 0){
    // we do need to delete one or more rows. return true
    for (var i = rowsToDelete.length - 1; i>=0; i--) {
      C3Checks.deleteRow(rowsToDelete[i]); 
    }
    return true;
  }else{
    // no rows to delete. return false;
    return false;
  }
} // end removeEmptyRows()

function addStatus_Columns(newRowC3,newLastRowC3){
  // @param: newRowC3; the first row # we pasted new assets into in C3Checks
  // @param: newLastRowC3; the last row # we pasted new assets into in C3Checks
  // Add the new formulas of H13:M13 into those rows.
  try{
    let range = 'H'+  newRowC3  +  ':M'  +  newLastRowC3;
    C3Checks.getRange('H13:M13').copyTo(C3Checks.getRange(range), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    let msg = "Status1,2,3 Formulas added to new assets from H13:M13 in C3 Checks.";
    updateLog(msg);
  }
  catch(e){
    let msg = "Error when adding formulas to new assets from H13:M13 in C3 Checks: "+e;
    updateLog(msg);
  }
} // end addStatus_Columns()

function insert_Brand6(totalAssets, newRowC3){ 
  // @param: totalAssets; total assets in for todays C3
  // @param: newRowC3; the first row # we pasted new assets into in C3Checks
  // Copy status's from Brand6 into C3Check Columns H
  let startRow_num = newRowC3;

 // Repeat the process each time for each asset, so i <= Total Number of Assets
 for(var i=1; i<=totalAssets; i++){
    // Column C in C3Checks has the network values
    // The Network value
    let rangeDestination = "H"+startRow_num;
    var network = C3Checks.getRange("C"+startRow_num).getValue();
    // removing white spaces and turning lowercase to avoid case sensitive errors
    //var networkLowercase = network.toString().toLowerCase().replace(/\s*/g,"").trim();
    var networkLowercase = network.toString().toLowerCase();
    Logger.log("Starting at row"+startRow_num+" # of Assets: "+totalAssets);
    Logger.log(networkLowercase+ "  "+networkLowercase.length);

    // Switch method will CopyPaste 12 rows from Brand6 based on the network value in C3Checks
    switch(networkLowercase){
    case "bravo":
      C3Checks.getRange(rangeDestination).activate();
      C3Checks.getRange('Brand6!C1:C12').copyTo(C3Checks.getRange(rangeDestination), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        // Increase the start row in C3Checks by 12 to move on to the next asset
        startRow_num = startRow_num+12;
        break;
    case "cnbc":
      C3Checks.getRange(rangeDestination).activate();
      C3Checks.getRange('Brand6!C14:C25').copyTo(C3Checks.getRange(rangeDestination), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        startRow_num = startRow_num+12;
        break;
    case "msnbc":
      C3Checks.getRange(rangeDestination).activate();
      C3Checks.getRange('Brand6!C27:C38').copyTo(C3Checks.getRange(rangeDestination), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        startRow_num = startRow_num+12;
        break;   
    case "nbc":
      C3Checks.getRange(rangeDestination).activate();
      C3Checks.getRange('Brand6!C40:C51').copyTo(C3Checks.getRange(rangeDestination), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        startRow_num = startRow_num+12;
        break;
    case "nbc news":
      C3Checks.getRange(rangeDestination).activate();
      C3Checks.getRange('Brand6!C53:C64').copyTo(C3Checks.getRange(rangeDestination), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        startRow_num = startRow_num+12;
        break;     
    case "oxygen":
      C3Checks.getRange(rangeDestination).activate();
      C3Checks.getRange('Brand6!C66:C77').copyTo(C3Checks.getRange(rangeDestination), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        startRow_num = startRow_num+12;
        break;      
    case "e!":
      C3Checks.getRange(rangeDestination).activate();
      C3Checks.getRange('Brand6!C79:C90').copyTo(C3Checks.getRange(rangeDestination), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        startRow_num = startRow_num+12;
        break;
    case "universo":
      C3Checks.getRange(rangeDestination).activate();
      C3Checks.getRange('Brand6!C92:C103').copyTo(C3Checks.getRange(rangeDestination), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        startRow_num = startRow_num+12;
        break;  
    case "telemundo":
      C3Checks.getRange(rangeDestination).activate();
      C3Checks.getRange('Brand6!C105:C116').copyTo(C3Checks.getRange(rangeDestination), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        startRow_num = startRow_num+12;
        break;
    case "usa":
      C3Checks.getRange(rangeDestination).activate();
      C3Checks.getRange('Brand6!C118:C129').copyTo(C3Checks.getRange(rangeDestination), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        startRow_num = startRow_num+12;
        break;
    case "golf":
      C3Checks.getRange(rangeDestination).activate();
      C3Checks.getRange('Brand6!C131:C142').copyTo(C3Checks.getRange(rangeDestination), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        startRow_num = startRow_num+12;
        break;
    case "syfy":
      C3Checks.getRange(rangeDestination).activate();
      C3Checks.getRange('Brand6!C144:C155').copyTo(C3Checks.getRange(rangeDestination), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        startRow_num = startRow_num+12;
        break;
    case "universal kids":
      C3Checks.getRange(rangeDestination).activate();
      C3Checks.getRange('Brand6!C157:C168').copyTo(C3Checks.getRange(rangeDestination), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        startRow_num = startRow_num+12;
        break;
    case "uni kids":
      C3Checks.getRange(rangeDestination).activate();
      C3Checks.getRange('Brand6!C157:C168').copyTo(C3Checks.getRange(rangeDestination), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        startRow_num = startRow_num+12;
        break;    
    default :
      C3Checks.getRange(rangeDestination).activate();
        break;
    }
  }
  Logger.log("Brand 6's Inserted...");
  updateLog("Brand 6's Inserted.");
} // end insert_Brand6()

// ---------- SYNDICATION NOTES START ---------- //

function confirmSyndication(){
  // Alert user to confirm moving forward with script.
  // alert(title, prompt, buttons)

  // Are the syndication timestamps up to date?
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are the Syndication Timestamps up-to-date ?', 'You can still update timestamps later for status 2 & 3', ui.ButtonSet.YES_NO);

  // Process the user's response.
  // If Yes, continue running next step of script.
  // If No, exit script, allow user to update syndication timestamps.
  // else default exit.
  if (response == ui.Button.NO) {
    // Syndication Timestamps are not updated
    let msg = "Script Terminated, Please update Syndication Timestamps.";
    updateMsg(msg);
    updateLog(msg);
    return;
  }else if(response == ui.Button.YES){
    // Networks are color coded continue to build C3
    let msg = "Building Syndication Notes...stand by.";
    updateMsg(msg);
    updateLog(msg);

    // Begin building c3 doc
    STARTSYNDICATION();
  }else{
    let msg = "<exited>";
    updateMsg(msg);
    updateLog(msg);
    return;
  }
} // end confirmSyndication()

// Build SyndicationNotes sheet //
function STARTSYNDICATION(){
  // get Constructor data
  const c3Data = new C3Info();
  let totalAssets = c3Data.totalAssets;
  let lastRowBuilder = c3Data.lastRowBuilder;

  SyndicationNotes.getRange("A2").activate();

  // Step 1: if Syndication Notes isn't already clear, clear it
  if(c3Data.lastRowSyndication != 1){
    // Last row is not the header, so there is data in sheet
    let msg = "Clearing SyndicationNotes tab";
    updateLog(msg);
    clearSyndication(c3Data.lastRowSyndication);
  }else{
    // last row is the header so the data is already clear
    let msg = "Syndication Notes Already Clear";
    updateLog(msg);
  }

  // Step 2: Copying from Builder into SynidcationNotes
  // If rowsToIgnore array is 0, we don't have to ignore any assets from list, Copy form Builder into Synidcation normally
  // else rowsToIgnore > 0 we do have rows to ignore so copy without ignores
  let newlastRowSyndication = 2;
  if(c3Data.rowsToIgnore.length == 0 ){
    let msg = "No rows in builder to ignore, copying Normally";
    updateLog(msg);
    newlastRowSyndication = copySyndicationNormal(c3Data.lastRowBuilder); // returns the lastRow of Syndiciation Notes AFTER assets are pasted
  }else if(c3Data.rowsToIgnore.length > 0){
    let msg = "We do have to ignore some rows in Builder.";
    updateLog(msg);
    newlastRowSyndication = copySyndicationIgnore(c3Data.rowsToIgnore, c3Data.lastRowBuilder);
  }else{
    // last row is the header so the data is already clear
    let msg = "Unkownvalue of rowsToIgnore, copying Normally";
    updateLog(msg);
    newlastRowSyndication = copySyndicationNormal(c3Data.lastRowBuilder);
  }

  // Step 3: Center the text and add borders to the range in Syndication Notes
  // Center the Text and Add Borders to SyndicationNotes
  updateLog("Formatting Syndication Notes.");
  SyndicationNotes.getRange("A2").activate();
  centerBorder(newlastRowSyndication);
  // resize the Columns Widths to show all texts in Syndication Notes
  resizeColumnWidth(newlastRowSyndication);
  // check the value of each cell and assign a background color accordingly 
  colorAssign(newlastRowSyndication);

  // Syndicaiton Notes created.
  let msg = "Syndication Notes Ready!";
  updateMsg(msg);
  updateLog(msg);
}

function clearSyndication(lastRowSyndication){
  // Syndication sheet was not empty so take lastrow and delete everything from A3:F lastrow
  SyndicationNotes.getRange("A2:F"+lastRowSyndication).clear();
  SyndicationNotes.getRange("A2:F"+lastRowSyndication).clearDataValidations();
} // end clearSyndication()

/*   Column Infomation  */
// Network and Series
//    Builder : Columns A:B
//    SyndicationNotes : Columns A:B
// CMC-C | CMC-D | MPX-TVE | VOD50-C
//    Builder : Columns H:K
//    SyndicationNotes : Columns C:F

function copySyndicationNormal(lastRowBuilder){
  // Row# : 3 - First Row we begin copying from in Builder.
  // Builder Range for Network and Series
  let range1 = "A3:B"+lastRowBuilder;
  // Builder Range for CMC-C | CMC-D | MPX-TVE | VOD50-C
  let range2 = "H3:K"+lastRowBuilder;
  // Builder Range for Dish - DirecTV-DAI - DirecTV-NBC@VOD50-C - Charter-NBC
  //let range3 = "Builder!K3:N"+lastRowBuilder;
  // range3 depricated as of 11/20/21 update

  // Copy Network and Series
  Builder.getRange(range1).copyTo(SyndicationNotes.getRange("A2"), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  // Copy Timestamps 
  Builder.getRange(range2).copyTo(SyndicationNotes.getRange("C2"), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  let msg = "Copied Syndications to Syndication Notes Normally.";
  updateLog(msg);
  let lastRowSyndication = SyndicationNotes.getLastRow();
  return lastRowSyndication;
}

function copySyndicationIgnore(rowsToIgnore, lastRowBuilder){
  // @param : rowsToIgnore : array of row's we will be ignoring in Builder when copying over assets to PrepC3.
  // @param : lastRowBuilder; the last row# in Builder.
  let startRowBuilder = 3;
  let startRowSyndication = 2; // initilize empty row #, this will be incremented each time a row is copied over to SyndicationNotes from Builder.

  // Loop through each row in Builder starting at row startRowBuilder and Column A
  for(let i = startRowBuilder; i<= lastRowBuilder; i++){
    let startCell = "A"+i;
    // if startCell is in array rowsToIgnore then do nothing, ignore the row move to next.
    if(rowsToIgnore.includes(startCell)){
      // Logger.log(startCell+ " Exist in array");
      // If this cell is in array, we skip this row and don't copy it.
    }else{
      // copy this row from Builder to SyndicationNotes and increment startRowSyndication by 1
      // Logger.log("Copying "+startCell+" To SyndicationNotes row: "+startRowSyndication);
      // copy Network & Series
      Builder.getRange(startCell+":B"+i).copyTo(SyndicationNotes.getRange("A"+startRowSyndication), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      Builder.getRange("H"+i+":K"+i).copyTo(SyndicationNotes.getRange("C"+startRowSyndication), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      startRowSyndication++;
    }
  } // done copying assets
  let msg = "Copied Over Assets to SyndicationNotes without ignores.";
  updateLog(msg);
  let lastRowSyndication = SyndicationNotes.getLastRow();
  return lastRowSyndication;
}

function centerBorder(newlastRowSyndication){
  // @param : newlastRowSyndication  :  the last row in Syndication Notes after we pasted the assets.
  let range = "A2:F"+newlastRowSyndication;
  // Center Text of Range
  SyndicationNotes.getRange(range).setHorizontalAlignment("center"); 
  // Add Border to Range
  SyndicationNotes.getRange(range).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);   Logger.log("Text center and border complete!.");
}

function resizeColumnWidth(newlastRowSyndication){
  // @param : newlastRowSyndication  :  the last row in Syndication Notes after we pasted the assets.
  // resize the column width to show all text in cells
  var column1 = 1;
  var column2 = 6;
  SyndicationNotes.autoResizeColumns(column1,column2);
  Logger.log("Column resize complete!.");
}

function colorAssign(newlastRowSyndication){
  // @param : newlastRowSyndication  :  the last row in Syndication Notes after we pasted the assets.
  // Change color of everything we pasted - to Green
  SyndicationNotes.getRange("A2:F"+newlastRowSyndication).setBackground('#00ff00');

  // Looping through each column: C-F (column index's 3-6)
  for(var col=3; col<=6; col++){
    // Looping through each row: starting at row 2 - till newlastRowSyndication
    for (var row=2; row<=newlastRowSyndication; row++){
      // convert row,col index to its range  ie; (row,col) = (3,2) in A1 Notation is B3
      let curentCell = SyndicationNotes.getRange(row,col).getA1Notation();
      //SyndicationNotes.getRange(curentCell).activate(); // Select starting cell
      // get the value of the cell and convert to lowercase
      let CellVal = SyndicationNotes.getRange(curentCell).getValue();
      // if it is not a timestamp, check if it is empty or requires data validation dropdown, if it is a timestamp (date) leave it alone
      if( Object.prototype.toString.call(CellVal) !== "[object Date]"){
        //Logger.log("It is NOT a Date");
        // Now it is either empty and we add N/A or we add the processing validation 
        if(CellVal == "" || CellVal.toLowerCase() == "n/a" || CellVal.toLowerCase() == "na"){
          // Empty cell/or N/A already in cell, add N/A with the color
          //Logger.log("It is an Empty Cell");
          SyndicationNotes.getRange(curentCell).setValue("N/A");
          SyndicationNotes.getRange(curentCell).setBackground('#b7b7b7'); // Change color to Gray
        }else{
          // not empty, not a timestamp, so we add the processing validation dropdown 
          //Logger.log("It is a processing issue");
          SyndicationNotes.getRange(curentCell).setBackground('#ffff00'); // Change color to yellow
          addProcessingValidation(curentCell);
        }
      }else{
        //Logger.log("it IS a date");
        // It is a date, don't do anything
      }
    }// End row 
  }// End col
  Logger.log("Color assign Complete!");
} // End colorAssign

function addProcessingValidation(curentCell){
  // given a cell, add the data validation to it
  //let curentCell = SyndicationNotes.getRange("C4"); // debugging purpose
  let Lists = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LISTS");
  let processingLastRow =  Lists.getRange("E2").getDataRegion().getLastRow(); // returns row # of last item in this specific column 
  let processingRange = "E2:E"+processingLastRow;
  let processingList = Lists.getRange(processingRange); 
  // Build datavalidation
  let processingValidation = SpreadsheetApp.newDataValidation().requireValueInRange(processingList).build();
  // Apply the dataValidation to a specific cell
  SyndicationNotes.getRange(curentCell).setDataValidation(processingValidation);
}

// ---------- Daily Asset Archiving START ---------- //
/*
  We archive the daily assets list and timestamps for record keeping.
  Previously we would manually create that doc, but now we copy assets to a sheet through the script. 
  * Folder "Daily Assets Archive" has Spreadsheets for every month, each month's Spreadsheet includes sheets(tabs) for each day which holds that days asset list and syndication time stamps
  * We will make a list of these Monthly Spreadsheets In order to keep track of them, this Master Daily Assets Spreadsheet Tracker will be called "Daily Archive Catalog Tracker"; Read ONLY.
*/
function confirmDailyArchive(){
  // Alert user to confirm moving forward with script.
  // alert(title, prompt, buttons)

  // Are the syndication timestamps up to date?
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Are you ready to archive today's Asset's list & Syndication timestamps?", "This should be done ONLY after we are done doing ALL checks on assets for the day.\nBest time to do thia is end of the day AFTER status 3 and Before creating the next days C3 doc.", ui.ButtonSet.YES_NO);

  // Process the user's response.
  // If Yes, continue running next step of script.
  // If No, exit script, allow user to update syndication timestamps.
  // else default exit.
  if (response == ui.Button.NO) {
    // Syndication Timestamps are not updated
    let msg = "Script Terminated, Archive Daily Assets Only After Status 3, Before creating next day C3 Doc.";
    updateMsg(msg);
    updateLog(msg);
    return;
  }else if(response == ui.Button.YES){
    // Networks are color coded continue to build C3
    let msg = "Archiving todays assets...stand by.";
    updateMsg(msg);
    updateLog(msg);

    // Start Daily Archive process
    STARTDAILYARCHIVE();
  }else{
    let msg = "<exited>";
    updateMsg(msg);
    updateLog(msg);
    return;
  }
} // end confirmArchive(

function STARTDAILYARCHIVE(){
  SpreadsheetApp.flush();
  // get Constructor data
  const c3Data = new C3Info();
  let airDate = c3Data.getAirDate; //   m/d/yyyy

  // Step 1: What will this airdates Spreadsheet file name be ? and expected tab name for days archive sheet
  // depends on the month and year of airDate
  let expectedFileName = getTodayArchiveFileName(airDate); // returns string of todays expected file name which should be in Folder: "Daily Assets Archive"
  // function to get expectedTabName as well, if file exists, the tab will be todays date ie; '12/23'  (month and date without the year)
  let expectedTabName = getTodayTabName(airDate); // returns todays expected tab name
  Logger.log("Expected Spreadsheet Name: "+ expectedFileName + " Expected sheet Name: "+expectedTabName);
  // Step 2: Get list of all files (spreadsheets) that already exist in Folder: "Daily Assets Archive"
  let allFilesNames = getAllFiles(); // returns an aray of all file names that exist in Daily Assets Archive folder
  //Logger.log(allFilesNames);

  // Step 3: if File already exists for this month, get the ID of that file, else create a file with name: expectedFileName in Folder: "Daily Assets Archive"
  if(allFilesNames.includes(expectedFileName)){
    // true if yes it exists, Get it's spreadsheet ID so we can check if expectedTabName exist's inside it.
    // get the files id
    Logger.log("This Months SS does exists");
    let todayFileId = getTodayFileId(expectedFileName); // returns the spreadsheet id of expectedFileName
    // check if today tab exists in file
    let todayTabExist = checkTodayTab(todayFileId, expectedTabName); // return true if exists, false if doesn't

    if(todayTabExist == false){
      Logger.log("this airdates tab does not exist");
      // have to create todays tab
      updateLog("This months archive SS already exists, but todays date tab does not. Creating Today's tab now.");
      createTodayTab(todayFileId, expectedTabName);
      updateLog("Created todays date tab, Ready to Archive Todays Assets.");
      // Archiving
      archiveTodaysAssets(c3Data.lastRowBuilder, todayFileId, expectedTabName );
    }else if (todayTabExist == true){
      Logger.log("this airdates tab already exist");
      // ready to copy assets
      updateLog("This months archive SS and todays date tab exist, Ready to Archive Todays Assets.");
      // Archiving 
      archiveTodaysAssets(c3Data.lastRowBuilder, todayFileId, expectedTabName );
    }else{
      // unkown error
    }
  }else{
    Logger.log("This months archive SS does NOT exist, creating now");
    // false, expectedFileName does not exist in Folder: "Daily Assets Archive" ()
    //Logger.log("It does NOT exist, lets create it");
    updateLog("This months Spreadsheet does not exist, creating new spreadsheet with name: "+expectedFileName);
    let todayFileId = createTodaySS(expectedFileName);
    // we just created the sheet so todayTab will not exist, create it.
    Logger.log("creating tab "+expectedTabName);
    createTodayTab(todayFileId, expectedTabName);
    updateLog("Created this months Archive Spreadsheet and todays date tab inside of it, Ready to Archive Todays Assets.");
    // Archiving 
    archiveTodaysAssets(c3Data.lastRowBuilder, todayFileId, expectedTabName );
  }

  let msg = "Todays Assets have been Archived Successfully,";
  updateLog(msg);

  Utilities.sleep(300);// pause for 300 milliseconds
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Todays Assets have been Archived in Folder 'Daily Assets Archive'", 'You can now change the airdate in MasterC3 of Spreadsheet DTS-C3 Checks to update assets list in Builder for tomorrows checks.',  ui.ButtonSet.OK);

  updateMsg("");
} // end STARTDAILYARCHIVE

function getTodayArchiveFileName(airDate){
// Name convention is : 'C3 Checks Month Year Archive'
  // extract month and year from airDate
  let airdateParts = airDate.split("/");
  let m_part = airdateParts[0];
  let y_part = airdateParts[2];
  
  // Convert Month to String 
  let monthStringList = {
    "1": "January",
    "2": "February",
    "3":"March",
    "4":"April",
    "5":"May",
    "6":"June",
    "7":"July",
    "8":"August",
    "9":"September",
    "10":"October",
    "11":"November",
    "12":"December"
  };
  // string version of month
  let m_part_string = monthStringList[m_part];
  // File Name of this airdates Archive Spreadsheet
  let todayArchiveFileName = "C3 Checks "+m_part_string+" "+y_part+" Archive";
  return todayArchiveFileName;
}

function getAllFiles(){
  // Function : getting all spreadsheet names in Daily Archive Folder
  Logger.log("Getting list of Spreadsheets...");
  let namesOfAllSpreadsheets = [];

  let files = DailyArchiveFolder.getFiles();
  while (files.hasNext()) {
    let file = files.next();
    namesOfAllSpreadsheets.push(file.getName());
  }
  return namesOfAllSpreadsheets;
}

function getTodayTabName(airDate){
  // Function : todays expectedTabName
  // split the airdate by / and store that into array
  let airDateParts = airDate.split("/");
  // create custom date with just the month and date and return
  let todayTabName = airDateParts[0]+"/"+airDateParts[1];
  return todayTabName;
}

function getTodayFileId(expectedFileName){
  // Function : file exist's get it's id
  let allFileIDs = []; // empty array to hold all file Id's that have the filename expectedFileName
  let allMatching = DailyArchiveFolder.getFilesByName(expectedFileName);
  while(allMatching.hasNext()){
    let file = allMatching.next();
    allFileIDs.push(file.getId());
  }
  Logger.log(allFileIDs);
  return allFileIDs[0];
} 

function checkTodayTab(todayFileId, expectedTabName){
  // Function : got file url for todays month, check for all tabs inside it for todays tab name
  let tabs = SpreadsheetApp.openById(todayFileId).getSheets();
  let allTabNames = [];
  let c = 0;
  while( c < tabs.length){
    allTabNames.push(tabs[c].getSheetName());
    c++;
  }
  Logger.log(allTabNames);
  if(allTabNames.includes(expectedTabName)){
    Logger.log("returning true");
    return true;
  }else{
    Logger.log("returning false");
    return false;
  }
}

function createTodaySS(expectedFileName){
  // Function : file for todays month does not exist, create it and return it's spreadsheet id
  var resource = {
    title: expectedFileName,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: DailyArchiveFolderId }]
  }
  var fileJson = Drive.Files.insert(resource)
  var fileId = fileJson.id
  // Logger.log(fileId);
  return fileId;
}

function createTodayTab(todayFileId, expectedTabName){
  // Function : file for todays month created, now create sheet(tab) for todays airdate(expectedTabName)
  let newSheet = SpreadsheetApp.openById(todayFileId).insertSheet();
  newSheet.setName(expectedTabName);
}

function archiveTodaysAssets(lastRowBuilder, todayFileId, expectedTabName ){
  // Function : today file and sheet found, archive the assets from today into it
  let source = Builder;
  let destination = SpreadsheetApp.openById(todayFileId).getSheetByName(expectedTabName);

  let sourceValues = source.getRange("A2:K"+lastRowBuilder).getValues(); // 2D array [ [NBC,asdsa,asdsad,sadasd], [USA,sadsad,asdsad,fdsfs] ]
  let destRow = 1; // the row we start pasting in the archive sheet

  // loop through each Array in the main sourceValues[]
  for(let i = 0; i < sourceValues.length; i++){
    let destCol = 1; // the col we start pasting in the arching sheet
    // for each array, start pasting on the current destLine
    for(let a = 0; a < sourceValues[i].length; a++){
      destination.getRange(destRow,destCol).setValue(sourceValues[i][a]);
      destCol++;
    }
    destRow++;
  }
  updateLog("Copied assets into archive, now formatting.");
  // format archive
  destination.getRange("A1:K"+destRow).setHorizontalAlignment("center"); // Alligning text to center  
  destination.getRange("A1:K"+destRow).setBorder(true, true, true, true, true, true); // setBorder(top, left, bottom, right, vertical, horizontal)
  destination.autoResizeColumns(1,11); // Resizing Column 1-11
  destination.getRange("A1:K1").setBackground("#4a86e8");
  Logger.log("Copied assets into Archive");
}

// DONE Archiving Daily Assets
// NOTES BELOW THIS LINE
/* 
The constructor variables we can use. Same as above where we define, but put a copy here just to view  
  // The Airdate in cell B1  
    c3Data.getAirDate
  //Last Row with data in Builder
    c3Data.lastRowBuilder
  // Row's to ignore in builder.(if any) : parameter lastRowBuilder : returns array with range's in column A ie; [A2,A12] or empty array if none found.
    c3Data.rowsToIgnore
  // Total Number of Assets
    c3Data.totalAssets
  // If the Last Row with data is row 1, that is the heading so the rest of the sheet is clear.
    c3Data.lastRowPrep
  // If the Last Row with data is row 1, that is the heading so the rest of the sheet is clear.
    c3Data.lastRowSyndication
*/

/*
  Formulas in Column A : =IFNA(ARRAYFORMULA(VLOOKUP($B1,IMPORTRANGE($P1,"MasterC3!A"&row()+3&":H"&row()+3),{2,3,4,5,6,7,8},0)))
 // Possible Formulas for Builder Columns, testing for potential
 // Only part that changes is the Column Letter under the star
 // New formula format per cell: =IFNA(IMPORTRANGE($P1,"MasterC3!B"&row()+3),"Input Manually In C3Checks")

              Builder Column  | MasterC3 Column    |   Formula in Builder Column
-----------------------------------------------------------------------
Network          :   A        :   B                :   =IFNA(IMPORTRANGE($P1,"MasterC3!B"&row()+3),"Input Manually In C3Checks")
Series           :   B        :   C                :   =IFNA(IMPORTRANGE($P1,"MasterC3!C"&row()+3),"Input Manually In C3Checks")
Title            :   C        :   D                :   =IFNA(IMPORTRANGE($P1,"MasterC3!D"&row()+3),"Input Manually In C3Checks")
Asset ID         :	 D        :   E                :   =IF(ISTEXT(IMPORTRANGE($P1,"MasterC3!E"&row()+3)), "Manually Enter In C3 Cheecks" ,IMPORTRANGE($P1,"MasterC3!E"&row()+3))
C3 HD Filename   :	 E        :   F                :   =IFNA(IMPORTRANGE($P1,"MasterC3!F"&row()+3),"Input Manually In C3Checks")
C3 SD Filename   :   F        :   G                :   =IFNA(IMPORTRANGE($P1,"MasterC3!G"&row()+3),"Input Manually In C3Checks")
Broadcast AirTime (POM/Time of Show) : G  : H      :   =IFNA(IMPORTRANGE($P1,"MasterC3!H"&row()+3),"Input Manually In C3Checks")
*/

// function addFormulasBuilder(){
//   // sample :  C3Checks.getRange("A"+i).setFormula('=TEXT(B'+newRowC3+',"MMMM")');
//   let original_formula = '=IFNA(ARRAYFORMULA(VLOOKUP($B1,IMPORTRANGE($P1,"MasterC3!A"&row()+3&":H"&row()+3),{2,3,4,5,6,7,8},0)))';
//   // let formula_col_A = '=IFNA(IMPORTRANGE($P1,"MasterC3!B"&row()+3),"Input Manually In C3Checks")';
//   // let formula_col_B = '=IFNA(IMPORTRANGE($P1,"MasterC3!C"&row()+3),"Input Manually In C3Checks")';
//   // let formula_col_C = '=IFNA(IMPORTRANGE($P1,"MasterC3!D"&row()+3),"Input Manually In C3Checks")';
//   // let formula_col_D = '=IF(ISTEXT(IMPORTRANGE($P1,"MasterC3!E"&row()+3)), "Manually Enter In C3 Cheecks" ,IMPORTRANGE($P1,"MasterC3!E"&row()+3))';
//   // let formula_col_E = '=IFNA(IMPORTRANGE($P1,"MasterC3!F"&row()+3),"Input Manually In C3Checks")';
//   // let formula_col_F = '=IFNA(IMPORTRANGE($P1,"MasterC3!G"&row()+3),"Input Manually In C3Checks")';
//   // let formula_col_G = '=IFNA(IMPORTRANGE($P1,"MasterC3!H"&row()+3),"Input Manually In C3Checks")';

//   let startRow = 3;
//   let endRow = 51;

//   for(let i = startRow; i<= endRow; i++){
//     // Builder.getRange("A"+i).setFormula(formula_col_A);
//     // Builder.getRange("B"+i).setFormula(formula_col_B);
//     // Builder.getRange("C"+i).setFormula(formula_col_C);
//     // Builder.getRange("D"+i).setFormula(formula_col_D);
//     // Builder.getRange("E"+i).setFormula(formula_col_E);
//     // Builder.getRange("F"+i).setFormula(formula_col_F);
//     // Builder.getRange("G"+i).setFormula(formula_col_G);
//     // original formula
//     Builder.getRange("A"+i).setFormula(original_formula);
//   }
//   Logger.log("Formulas Set");
// }