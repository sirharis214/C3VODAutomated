Haris Nasir - C3VODAutomated - Git Branch : magicmakerv6
Scripts are created for NBCU Digital NOC.
This automates the creation of the Daily C3 checks Doc and simplifies the processes by allowing L1's to focus on one Doc. 

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
