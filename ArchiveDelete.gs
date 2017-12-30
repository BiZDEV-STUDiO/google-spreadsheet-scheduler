function archiveReadyUnitTicketsDeleteOriginal() {
// The code below will check if a new ticket name was generated today, and 
// if so, it will use the name to create a new ticket. Then it will set values
// in the ticket pulled from the same row.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FieldSupervisor_AlbionTerrace");
  // figure out what the last row is
  var lastRow = sheet.getLastRow();
  // timestamp today's date
//  var today = sheet.getRange("CQ2");
//  var dateValue = new Date();
//  var checkDateValue = new Date();
//  sheet.getRange("CQ2").setValue(dateValue);
 
  // the rows are indexed starting at 1, and the first 2 rows
  // are the headers, so start with row 3
  var startRow = 3;
  
  // Now, grab the empty column column
  var range = sheet.getRange(3, 105, lastRow-startRow+1, 1);
  var numRows = range.getNumRows();
  var empty_column_values = range.getValues();
  
  // Now, grab the date ready value column
  range = sheet.getRange(3, 106, lastRow-startRow+1, 1);
  var date_ready_values = range.getValues();
  
  // Now, today's value column
  range = sheet.getRange(3, 107, lastRow-startRow+1, 1);
  var todays_date_values = range.getValues();
  
  // Now, grab the sheet name column
  range = sheet.getRange(3, 90, lastRow-startRow+1, 1);
  var sheet_name_values = range.getValues();
  
  // Now, grab the ticket archive folder name column
  range = sheet.getRange(3, 108, lastRow-startRow+1, 1);
  var archiveFolder_name_values = range.getValues();   
 
  var warning_count = 0;
  var msg = "";
   
  // Loop over the values
  for (var i = 0; i <= numRows - 1; i++) {
    var empty_column = empty_column_values[i][0];
    var date_ready = date_ready_values[i][0];
    var todays_date = todays_date_values[i][0];
    var sheet_name = sheet_name_values[i][0];
    var archiveFolder_name = archiveFolder_name_values[i][0];
    
    if(date_ready != '' && sheet_name != '' 
         && empty_column == '' && date_ready <= todays_date) {
  
      copyValues();
  var originalSpreadsheet = SpreadsheetApp.getActive();
  var newSpreadsheet = SpreadsheetApp.create("Spreadsheet to export");
  var sheetToCopy = ss.getSheetByName(sheet_name);
//  sheetToCopy = originalSpreadsheet.getActiveSheet();
  sheetToCopy.copyTo(newSpreadsheet);
  
//    var sheetToCopy = ss.getSheetByName('TicketMaster').copyTo(ss);
     SpreadsheetApp.flush(); // Utilities.sleep(2000);
     
//  var ticketProperty = sheetToCopy.getRange("D3");
//  var ticketUnit = sheetToCopy.getRange("D4");
//  var ticketMix = sheetToCopy.getRange("D5");
//  var ticketOldRent = sheetToCopy.getRange("B33");
//  var ticketNewRent = sheetToCopy.getRange("B34");
//  
//  
//  ticketProperty.setValue(property_name);
//  ticketUnit.setValue(unit_number);
//  ticketMix.setValue(unit_mix);
////  ticketOldRent.setValue(old_rent);
////  ticketNewRent.setValue(new_rent);
//  ticketOldRent.setValue("=FieldSupervisor_AlbionTerrace!"+ old_rent_cell);
//  ticketNewRent.setValue("=FieldSupervisor_AlbionTerrace!" + new_rent_cell);
//  sheetToCopy.setName(sheet_name);
  // My code - get the old and new rent values in b33 and b34
// which are pasted as formulas unless we copy and paste 
// the values before exporting and sending
  var rangeToCopy = sheetToCopy.getRange('B33:B34');
  var data = rangeToCopy.getValues();
  var tss = SpreadsheetApp.openById(newSpreadsheet.getId());
  

  // My code - get the new spreadsheet to be exported and paste old and new rent
  // values into the same corresponding range.
  var ts = newSpreadsheet.getSheets()[1].activate();
//  Logger.log(ts.getSheetName());
  ts.getRange(33,2,2,1).setValues(data);
//  Logger.log(data);


  // Find and delete the default "Sheet 1", after the copy to avoid triggering an apocalypse
  newSpreadsheet.getSheetByName('Sheet1').activate();
  newSpreadsheet.deleteActiveSheet();
  //
  ss.getSheetByName(sheet_name).activate();
  ss.deleteActiveSheet();

 DriveApp.getFileById(newSpreadsheet.getId()).setName(sheet_name);
 //create an archive folder if it doesn't exist and put the file in the folder
      var par_fdr = DriveApp.getFolderById('0B4R_lB18iUkGYVdhUVRIMjZjN3M');
      var fdr_name = archiveFolder_name;
      
      try {
        var newFdr = par_fdr.getFoldersByName(fdr_name).next();
      }
      catch(e) {
        var newFdr = par_fdr.createFolder(fdr_name);
      }
//end archive function
 // add the files to the correct folder
      var ticketsToFile = DriveApp.getFilesByName(sheet_name);
      
      while (ticketsToFile.hasNext()) {
        var ticketToFile = ticketsToFile.next();
        var dest_folder = par_fdr.getFoldersByName(fdr_name).next();
        dest_folder.addFile(ticketToFile);
      }       
    }
  }
  }
  
  function copyValues(){
// This code will create a timestamp on the day an new ticket sheet name is generated.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FieldSupervisor_AlbionTerrace");
  var lastRow = sheet.getLastRow()
  var dateString= new Date();
  dateString=new Date(dateString).toUTCString();
  dateString=dateString.split(' ').slice(0, 4).join(' ')

  for (var i = 1; i <= lastRow; i++){
// Estimated Costs Column
    var estimatedCostsRange = sheet.getRange([i], 96);
    var estimatedCostsValues = estimatedCostsRange.getValues();
// Copy Estimated Costs Column
    var copyEstimatedRange = sheet.getRange([i], 97);
    var copyEstimatedValues = copyEstimatedRange.getValues();
// Major TO items Column
    var majorItemsRange = sheet.getRange([i], 98);
    var majorItemsValues = majorItemsRange.getValues();
// Copy Major TO items Column
    var copyItemsRange = sheet.getRange([i], 99);
    var copyItemsValues = copyItemsRange.getValues();
// Date Ready Column
    var dateReadyRange = sheet.getRange([i], 15);
    var dateReadyValues = dateReadyRange.getValues();
// This Script Time Stamp Column
    var timeStampRange = sheet.getRange([i], 100);
    var timeStampValues = timeStampRange.getValues();
// Copy Completed Time Stamp Column
    var copyCompleteRange = sheet.getRange([i], 101);
    var copyCompleteValues = timeStampRange.getValues();     
    
    if (dateReadyValues != '' && copyCompleteValues == ''){ 
      copyEstimatedRange.setValue(estimatedCostsValues);
      copyItemsRange.setValue(majorItemsValues);
      timeStampRange.setValue(dateString);
      
    copyCompleted();
      }
    }
}

function copyCompleted(){
// This code will create a timestamp on the day an new ticket sheet name is generated.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FieldSupervisor_AlbionTerrace");
  var lastRow = sheet.getLastRow()
  var dateString= new Date();
  dateString=new Date(dateString).toUTCString();
  dateString=dateString.split(' ').slice(0, 4).join(' ')

  for (var i = 1; i <= lastRow; i++){
// Copy Values Script Time Stamp Column
    var timeStampRange = sheet.getRange([i], 100);
    var timeStampValues = timeStampRange.getValues();
// Copy Completed Date
    var timeRange = sheet.getRange([i], 101);
    var timeValues = timeRange.getValues();

    if (timeStampValues != '' && timeValues == ''){ // 
      timeRange.setValue(dateString);

      }
    }
} 

//function createArchiveFolder(){
// var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FieldSupervisor_AlbionTerrace");
//  // figure out what the last row is
//  var lastRow = sheet.getLastRow();
//  // timestamp today's date
////  var today = sheet.getRange("CQ2");
////  var dateValue = new Date();
////  var checkDateValue = new Date();
////  sheet.getRange("CQ2").setValue(dateValue);
// 
//  // the rows are indexed starting at 1, and the first 2 rows
//  // are the headers, so start with row 3
//  var startRow = 3;
//  
//  // Now, grab the empty column column
//  var range = sheet.getRange(3, 105, lastRow-startRow+1, 1);
//  var numRows = range.getNumRows();
//  var empty_column_values = range.getValues();
//  
//  // Now, grab the date ready value column
//  range = sheet.getRange(3, 106, lastRow-startRow+1, 1);
//  var date_ready_values = range.getValues();
//  
//  // Now, today's value column
//  range = sheet.getRange(3, 107, lastRow-startRow+1, 1);
//  var todays_date_values = range.getValues();
//  
//  // Now, grab the sheet name column
//  range = sheet.getRange(3, 90, lastRow-startRow+1, 1);
//  var sheet_name_values = range.getValues();
//  
//  // Now, grab the ticket archive folder name column
//  range = sheet.getRange(3, 108, lastRow-startRow+1, 1);
//  var archiveFolder_name_values = range.getValues();   
// 
//  var warning_count = 0;
//  var msg = "";
//   
//  // Loop over the values
//  for (var i = 0; i <= numRows - 1; i++) {
//    var empty_column = empty_column_values[i][0];
//    var date_ready = date_ready_values[i][0];
//    var todays_date = todays_date_values[i][0];
//    var sheet_name = sheet_name_values[i][0];
//    var archiveFolder_name = archiveFolder_name_values[i][0];
//    
//    if(date_ready != '' && sheet_name != '' 
//         && empty_column == '' && date_ready <= todays_date) {
//
//      var par_fdr = DriveApp.getFolderById('0B4R_lB18iUkGYVdhUVRIMjZjN3M');
//      var fdr_name = archiveFolder_name;
//      
//      try {
//        var newFdr = par_fdr.getFoldersByName(fdr_name).next();
//      }
//      catch(e) {
//        var newFdr = par_fdr.createFolder(fdr_name);
//      }
//      } 
//        
//
////   archiveTicketFolder.addFile(sheetInDrive); 
//}
//}
