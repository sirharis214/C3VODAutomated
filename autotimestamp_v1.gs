// This Script Admin,Owner and Developer is Haris Nasir.
// Any and all changes must be approved by Haris Nasir due to the complex build of this script.
// Thursday 06/18/2020 12:50 ET

// ** NOTES git: autoTimestamp**
// These two functions automate the insertion of timestamp when 
//changing the status of asset during sweep.
// This checks to make sure the Cell's being editing are in columns H,J, or L and updates the column to the right with the current timestamp upon edit.
// If status is changed back to NULL, timestamp is erased.

function onEdit(cel) {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var activeCell = activeSheet.getActiveCell();
  var row = activeCell.getRow();
  var col = activeCell.getColumn();
  var val = activeCell.getValue();
  Logger.log(val);
  
 // We want to Make sure the script is being applied to Columns H,J or L Only (8,10,12)
  if(col == 8 || col == 10 || col == 12){
    if(val != "NULL"){ // If the value is anything but NULL , get the timestamp as a value
      var timestamp = getTimestamp();
      activeSheet.getRange(row, col+1).setValue(timestamp);
    }
    else if(val == "NULL"){
      activeSheet.getRange(row,col+1).setValue(" ");
    }
  }
}

function getTimestamp(){
 var date = new Date();
     var month = date.getMonth()+1;month = month.toString().slice(-1);
     var day = date.getDate(); if(day.toString().length==1){var day = day.toString().slice(-1);}
     var year = date.getFullYear().toString().slice(-2);
     var hr = date.getHours();
     //var min = date.getMinutes();
     var min = (date.getMinutes()<10?'0':'') + date.getMinutes();  //if the min's is 0-9 add a leading 0 to make it a double digit value
  
 var timestamp = month+"/"+day+"/"+year+" "+hr+":"+min;
 Logger.log(timestamp);
     return timestamp;
}
