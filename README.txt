Haris Nasir - C3VODAutomated - Git Branch : archiveV1
Scripts are created for NBCU Digital NOC.
This automates the creation of the Daily C3 checks Doc and simplifies the processes by allowing L1's to focus on one Doc. 

New Feature archiveV1:
* The Builder Sheet is reused everyday, 
  erasing old assets and replaced with new assets by an L1-A.

  -- BEFORE THEY ERASE OLD ASSET'S --

* Once all Syndication timestamps are complete at the end of the day, an L1
  Has to Click Archive Button.
* They will need to provide the URL of the Google Spreadsheet correlating to 
  that Month.
 - This Archive Script Searches the Names of all Sheets in that Spreadsheet,
   Looking for a sheetname with todays date.
 - If it finds it, it will copy over all asset data from Builder into this sheet.
 - If does NOT find a sheet with todays date as the name, 
   It will create the sheet + copy over all asset data from Builder into this sheet. 


New Sheet comparison to Traditional Sheets:
"test" : C3 Checks *Month* 
"sheet7" : Syndication Notes
"sheet6" : C3 Checks (VOD Checks Doc)

Rolls and Responsibilities:
Evening L1 : L1-A
Overnight L1 : L1-B
Nextmorning L1 (morning of checks) : L1-C 

L1-A Will Copy over vod asset info that require checks the next day into the first sheet named "test"
   - This includes Episode related info, Asset File Name etc. as per usual.
   - No need to format or color code this sheet, this is done by the script.
   - No need to perform Syndication Updates by L1-A
   - Not Required to create the VOD-Doc anymore, done by script.
   - IMPORTANT : Do Not delete Row 1 in sheet's "test" and "sheet7"

L1-B Will Update Syndication timestamps on this singular "test" Sheet throughout the night.
   - No need to add formatting.
   - If a Asset is not expected to have a syndication time stamp (N/A), simply leave the cell empty, Script will search for empty cells and replace with "N/A".
   - Pending & Transcoding status for syndication can be adding although script does not taylor to these conditions in current version. 
   - Once Syndication Timestamps are updated, or are not complete but Partners expect updated syndications, Simply Click the "Build Syndication" button on Row1 of Sheet "test".
   - Since L1-A will not create VODdoc, you are not required to enter Brand-6 either.
   - In the event all Syndication Time stamps are complete, you may click the Big Red Button labeled "Build All" which will create "sheet7" and "sheet6" in one swoop.

L1-C Will create the VODdoc.
  - First you may update sheet "test" with syndications.
  - Not all syndication timestamps will be complete, if you do update any cell, click "Build Syndication" button to update sheet7.
  - You may now click "Build Doc" which will create the C3 vod doc performing all neccessary filtering and data validation insertions and formulas. 
  - Perform C3 sweeps as usual. 
