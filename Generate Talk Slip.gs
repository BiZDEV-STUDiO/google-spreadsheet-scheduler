function sheetNameGeneratedTimestamp(){
// This code will create a timestamp on the day an new ticket sheet name is generated.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FieldSupervisor_AlbionTerrace");
  var lastRow = sheet.getLastRow()
  var dateString= new Date();
  dateString=new Date(dateString).toUTCString();
  dateString=dateString.split(' ').slice(0, 4).join(' ')

  for (var i = 1; i <= lastRow; i++){
// New Ticket Sheet Names
    var ticketNameRange = sheet.getRange([i], 90);
    var ticketNameValues = ticketNameRange.getValues();
// Timestamp date created
    var timeRange = sheet.getRange([i], 91);
    var timeValues = timeRange.getValues();

    if (ticketNameValues != '' && timeValues == ''){ // 
      timeRange.setValue(dateString);
    checkIfTicketSheetNameGenerated();
      }
    }
}
function checkIfTicketSheetNameGenerated() {

// The code below will check if a new ticket sheet name 
//  was generated and will place a timestamp next to timestamp above.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetWithCellForName = ss.getSheetByName('FieldSupervisor_AlbionTerrace')
  var lastRow = sheetWithCellForName.getLastRow();
  var dateString= new Date();
  dateString=new Date(dateString).toUTCString();
  dateString=dateString.split(' ').slice(0, 4).join(' ')
//  var todayDate = new Date();
  
    for (var i = 1; i <= lastRow; i++){
// New Ticket Sheet Names
    var readyRange = sheetWithCellForName.getRange([i], 90);
    var readyValues = readyRange.getValues();
// Date Name Created
    var timeRange = sheetWithCellForName.getRange([i], 91);
    var timeValues = timeRange.getValues();
// This Timestamp    
    var timeRange2 = sheetWithCellForName.getRange([i], 92);
    var timeValues2 = timeRange2.getValues();
    
    if (timeValues != '' && timeValues2 == ''){ // 
      timeRange2.setValue(dateString);
      
}
}
}


function createNewTicket() {
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
  
  // grab column 93 (the 'date name generated' column) 
  var range = sheet.getRange(3,93,lastRow-startRow+1,1 );
  var numRows = range.getNumRows();
  var date_generated_values = range.getValues();
   
  // Now, grab the date checked column
  range = sheet.getRange(3, 94, lastRow-startRow+1, 1);
  var date_checked_values = range.getValues();
  
  // Now, grab the ticket sheet names column
  range = sheet.getRange(3, 90, lastRow-startRow+1, 1);
  var sheet_name_values = range.getValues();
  
  // Now, grab the property column
  range = sheet.getRange(3, 1, lastRow-startRow+1, 1);
  var property_name_values = range.getValues();
  
  // Now, grab the unit number column
  range = sheet.getRange(3, 2, lastRow-startRow+1, 1);
  var unit_number_values = range.getValues();
  
  // Now, grab the apartment mix column
  range = sheet.getRange(3, 3, lastRow-startRow+1, 1);
  var unit_mix_values = range.getValues();
  
  // Now, grab the old rent column
  range = sheet.getRange(3, 8, lastRow-startRow+1, 1);
  var old_rent_values = range.getValues();

  
  // Now, grab the new rent column
  range = sheet.getRange(3, 9, lastRow-startRow+1, 1);
  var new_rent_values = range.getValues();
  
  // Now, grab the old rent cell values column
  range = sheet.getRange(3, 102, lastRow-startRow+1, 1);
  var old_rent_cellReference = range.getValues();
  
  // Now, grab the new rent cell values column
  range = sheet.getRange(3, 103, lastRow-startRow+1, 1);
  var new_rent_cellReference = range.getValues();

  
  
  // Now, grab the timestamp for this script column
  range = sheet.getRange(3, 95, lastRow-startRow+1, 1);
  var time_stamp_values = range.getValues();
   
  var warning_count = 0;
  var msg = "";
   
  // Loop over the values
  for (var i = 0; i <= numRows - 1; i++) {
    var date_generated = date_generated_values[i][0];
    var date_checked = date_checked_values[i][0];
    var sheet_name = sheet_name_values[i][0];
    var time_stamp = time_stamp_values[i][0];
    var property_name = property_name_values[i][0];
    var unit_number = unit_number_values[i][0];
    var unit_mix = unit_mix_values[i][0];
    var old_rent = old_rent_values[i][0];
    var new_rent = new_rent_values[i][0];
    var old_rent_cell = old_rent_cellReference[i][0];
    var new_rent_cell = new_rent_cellReference[i][0];

    
    if(time_stamp == '' && date_generated != '' && date_generated==date_checked) {
    
//    for (var i = 1; i <= lastRow; i++){
// New Ticket Sheet Names
//    var readyRange = sheet.getRange([i], 90);
//    var readyValues = readyRange.getValues();
//// Date Name Created
//    var timeRange = sheet.getRange([i], 91);
//    var timeValues = timeRange.getValues();
//// This Timestamp    
//    var timeRange2 = sheet.getRange([i], 92);
//    var timeValues2 = timeRange2.getValues();
   
    var sheetToCopy = ss.getSheetByName('TicketMaster').copyTo(ss);
     SpreadsheetApp.flush(); // Utilities.sleep(2000);
     
  var ticketProperty = sheetToCopy.getRange("D3");
  var ticketUnit = sheetToCopy.getRange("D4");
  var ticketMix = sheetToCopy.getRange("D5");
  var ticketOldRent = sheetToCopy.getRange("B33");
  var ticketNewRent = sheetToCopy.getRange("B34");
  
  
  ticketProperty.setValue(property_name);
  ticketUnit.setValue(unit_number);
  ticketMix.setValue(unit_mix);
//  ticketOldRent.setValue(old_rent);
//  ticketNewRent.setValue(new_rent);
  ticketOldRent.setValue("=FieldSupervisor_AlbionTerrace!"+ old_rent_cell);
  ticketNewRent.setValue("=FieldSupervisor_AlbionTerrace!" + new_rent_cell);
  sheetToCopy.setName(sheet_name);
  

    
 
  /* Make the new sheet active */
//  ss.setActiveSheet(sheetToCopy);
  ticketGeneratedTimestamp();
   
       
    }
  }
 }
 
function ticketGeneratedTimestamp(){
// This code will create a timestamp on the day an new ticket sheet name is generated.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FieldSupervisor_AlbionTerrace");
  var lastRow = sheet.getLastRow()
  var dateString= new Date();
  dateString=new Date(dateString).toUTCString();
  dateString=dateString.split(' ').slice(0, 4).join(' ')

  for (var i = 1; i <= lastRow; i++){
// New Ticket Sheet Names
    var ticketNameRange = sheet.getRange([i], 94);
    var ticketNameValues = ticketNameRange.getValues();
// Timestamp date created
    var timeRange = sheet.getRange([i], 95);
    var timeValues = timeRange.getValues();

    if (ticketNameValues != '' && timeValues == ''){ // 
      timeRange.setValue(dateString);

      }
    }
} 
 
 
//function cloneGoogleSheet() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var sheetWithCellForName = ss.getSheetByName('FieldSupervisorPFIReport')
//  var lastRow = sheetWithCellForName.getLastRow();
//  var dateString= new Date();
//  dateString=new Date(dateString).toUTCString();
//  dateString=dateString.split(' ').slice(0, 4).join(' ')
// 
//    for (var i = 1; i <= lastRow; i++){
//// New Ticket Sheet Names
//    var readyRange = sheetWithCellForName.getRange([i], 93);
//    var readyValues = readyRange.getValues();
//// Date Name Created
//    var timeRange = sheetWithCellForName.getRange([i], 94);
//    var timeValues = timeRange.getValues();
//// This Timestamp    
//    var timeRange2 = sheetWithCellForName.getRange([i], 92);
//    var timeValues2 = timeRange2.getValues();
//
//
//
//  var sheet = ss.getSheetByName('TicketMaster').copyTo(ss);
//  
//  /* Before cloning the sheet, delete any previous copy */
//  var old = ss.getSheetByName(name);
//  if (old) ss.deleteSheet(old); // or old.setName(new Name);
//  var company = "TicketMaster2";
//  
//  SpreadsheetApp.flush(); // Utilities.sleep(2000);
//  sheet.setName(company);
// 
//  /* Make the new sheet active */
//  ss.setActiveSheet(sheet);
// 
//}
//}

//var localTimeZone = "Europe/London";
//var dateFormatForFileNameString = "yyyy-MM-dd'at'HH:mm:ss";  

//function checkReminder() {
//  // get the spreadsheet object
//  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//  // set the first sheet as active
//  SpreadsheetApp.setActiveSheet(spreadsheet.getSheets()[0]);
//  // fetch this sheet
//  var sheet = spreadsheet.getActiveSheet();
//   
//  // figure out what the last row is
//  var lastRow = sheet.getLastRow();
// 
//  // the rows are indexed starting at 1, and the first row
//  // is the headers, so start with row 2
//  var startRow = 2;
// 
//  // grab column 5 (the 'days left' column) 
//  var range = sheet.getRange(2,5,lastRow-startRow+1,1 );
//  var numRows = range.getNumRows();
//  var days_left_values = range.getValues();
//   
//  // Now, grab the reminder name column
//  range = sheet.getRange(2, 1, lastRow-startRow+1, 1);
//  var reminder_info_values = range.getValues();
//   
//  var warning_count = 0;
//  var msg = "";
//   
//  // Loop over the days left values
//  for (var i = 0; i <= numRows - 1; i++) {
//    var days_left = days_left_values[i][0];
//    if(days_left == 7) {
//      // if it's exactly 7, do something with the data.
//      var reminder_name = reminder_info_values[i][0];
//       
//      msg = msg + "Reminder: "+reminder_name+" is due in "+days_left+" days.\n";
//      warning_count++;
//    }
//  }
//   
//  if(warning_count) {
//    MailApp.sendEmail("youremail@yourdomain.com,anotheremail@anotherdomain.com", 
//        "Reminder Spreadsheet Message", msg);
//  }
//   
//}
// 
