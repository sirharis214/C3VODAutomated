function removeDup() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dup");
  Logger.clear();
  
  var values = sheet.getDataRange().getValues();
  var allDates = [];
  var cleanDates = [];
  for(var i=1;i<values.length;++i){
    var cell = values[i][0] ; // x is the index of the column starting from 1
    allDates.push(cell)
      //Logger.log(cell);
  }
  
  allDates.sort();
  for(var n=0;n<allDates.length;n++){
   var Val1 = allDates[n];
    var cnt = n+1;
   var Val2 = allDates[cnt];
   
    if(Val1 == Val2){
      Logger.log(Val1+" Already Exists");
    }
    else{
     Logger.log(Val1+" Added to List");
      cleanDates.push(Val1);
    }
  }
  
  Logger.log("The List After Removing Dup's:\n"+cleanDates);
  //Logger.log(allDates);
}


// Below is a rough draft

function Filter() {
  // Chances are the Rows that we need for Status Formulas are hidden.
  // Before Creating the Doc we need to filter the doc to display these rows. (1-13)
  
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet6");
  sheet6.getRange('A1').activate();
  //var criteria = SpreadsheetApp.newFilterCriteria();
  
  var criteria = SpreadsheetApp.newFilterCriteria();
  criteria.whenTextContains("Filler")
  .build();
  sheet6.getFilter().setColumnFilterCriteria(1, criteria);
};
