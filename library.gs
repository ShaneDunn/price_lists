function ExcelDateToJSDate(serial) {
   var utc_days  = Math.floor(serial - 25569);
   var utc_value = utc_days * 86400;                                        
   var date_info = new Date(utc_value * 1000);

   var fractional_day = serial - Math.floor(serial) + 0.0000001;

   var total_seconds = Math.floor(86400 * fractional_day);

   var seconds = total_seconds % 60;

   total_seconds -= seconds;

   var hours = Math.floor(total_seconds / (60 * 60));
   var minutes = Math.floor(total_seconds / 60) % 60;

   return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
}

/* =========== Spreadsheet Function ======================= */
/**
* Gets the Sheet Name of a selected Sheet.
*
* @param {number} option 0 - Current Sheet, 1 - All Sheets, 2 - Spreadsheet filename
* @return the current sheet name, or all sheet names, or the name of the spreadsheet based on the given parameter.
* @customfunction
*/
 
function SHEETNAME(option) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var thisSheet = sheet.getName();
  
  //Current option Sheet Name
  if(option === 0){
    return thisSheet;
  
  //All Sheet Names in Spreadsheet
  }else if(option === 1){
    var sheetList = [];
    ss.getSheets().forEach(function(val){
       sheetList.push(val.getName())
    });
    return sheetList;
  
  //The Spreadsheet File Name
  }else if(option === 2){
    return ss.getName();
  
  //Error  
  }else{
    return "#N/A";
  };
};

/* =========== Menu functions ======================= */

/**
 * Present the help page [help.html].
 * @private.
 */
function help() {
  var html = HtmlService.createHtmlOutputFromFile('help')
  .setTitle("Google Scripts Support")
  .setWidth(400)
  .setHeight(260);
  var ss = SpreadsheetApp.getActive();
  ss.show(html);
}

/**
 * Initialise Authorisation.
 * @private.
 */
function init() {
  return;
}

