
/* =========== Main CIM Load File Creation function =========== */

function createCIMload(e) {
  setupLog_();
  var i, config, configName;
  log_('Running on: ' + now);

  var configs = getConfigs_(getOrCreateSheet_(CONFIG_SHEET));

  if (!configs.length) {
    log_('No CIM Load configurations found');
  } else {
    log_('Found ' + configs.length + ' CIM Load configurations.');

    for (i = 0; config = configs[i]; ++i) {
      //Logger.log(config);
      configName = config.file_name;
      //Logger.log(configName);
      //Logger.log(config['sheet_name']);
      if (config['sheet_name']) {
        try {
          log_('Creating CIM Load for: ' + configName);
          saveAsCim(config);
        } catch (error) {
          log_('Error executing ' + configName + ': ' + error.message);
        }
      } else {
        log_('No sheet-name found: ' + configName);
      }
    }
  }
  now = new Date();
  log_('Script done: ' + now);

  // Update the user about the status of the queries.
  if( e === undefined ) {
    showLogDialog_();
    dumpLog_(getOrCreateSheet_(LOG_SHEET));
    dumpError_(getOrCreateSheet_(ERROR_SHEET));
  }
}

/* =========== Secondary CIM Load File Creation functions =========== */

function saveAsCim(config) {
  // Name the file
  var fileName = config['file_name']; //"apvomt.cim";
  // Convert the range data to CIM format
  log_('Creating CIM Load output from sheet: ' + config['sheet_name']);
  var csvFile = convertRangeToCimFile_(config);
  // Create a file in the root of google Drive with the given name and the CSV data
  log_('Creating CIM Load file in Drive. File Name: ' + fileName);
  DriveApp.createFile(fileName, csvFile);
}

function convertRangeToCimFile_(config) {
  // Get the range to be exported from the app
  var sheet_name      = config['sheet_name'];
  var exportRange     = config['exportRange'];
  var screenNameRow   = config['screenNameRow'];
  var screenNoRow     = config['screenNoRow'];
  var variableTypeRow = config['variableTypeRow'];
  var screenLenRow    = config['screenLenRow'];
  var program_name    = config['program_name'];
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  var rangeToExport = ss.getRange(exportRange);                 // SpreadsheetApp.getActiveSpreadsheet().getRange("A12:BE104");
  var screenName    = ss.getRange(screenNameRow).getValues();   // SpreadsheetApp.getActiveSpreadsheet().getRange("A3:BE3").getValues();
  var screenNo      = ss.getRange(screenNoRow).getValues();     // SpreadsheetApp.getActiveSpreadsheet().getRange("A3:BE3").getValues();
  var varType       = ss.getRange(variableTypeRow).getValues(); // SpreadsheetApp.getActiveSpreadsheet().getRange("A5:BE5").getValues();
  var screenLen     = ss.getRange(screenLenRow).getValues();    // SpreadsheetApp.getActiveSpreadsheet().getRange("A5:BE5").getValues();
  var progName      = ss.getRange(program_name).getValues();    // SpreadsheetApp.getActiveSpreadsheet().getRange("B2:B2").getValues();
  
  try {
    var dataToExport = rangeToExport.getDisplayValues();
    var csvFile = undefined;

    // Loop through the data in the range and build a string with the CSV data
    if (dataToExport.length > 1) {
      var csv = "";
      var vLine = "";
      var vField = "";
      for (var row = 0; row < dataToExport.length; row++) {
        vLine = "";
        for (var col = 0; col < dataToExport[row].length; col++) {
          vField = dataToExport[row][col].toString();

          var test1 = screenLen[0][col];
          if (screenLen[0][col] == "N/A") {
            vField = "";
          }
          else if (screenLen[0][col] != "-" && vField.length > parseInt(screenLen[0][col])){
            log_('WARNING: String Length greater than screen maximum. Field: ' + screenName[0][col] + ' Row: ' + row + ' Lenght: ' + vField.length + ' Max Lenght: ' + parseInt(screenLen[0][col]));
          }
          
          if (screenNo[0][col] == "C") {
            if (vField == "~") {
              vField = "\n";
            } else if (vField == ".") {
              vField = vField + "\n";
            } else {
              vField = "";
            }
          }

          if (vField.indexOf(" ") != -1) {
            vField = "\"" + vField + "\"";
          }
          else if (varType[0][col] == "c" && vField != "-" && vField != ";" && vField != "." && screenLen[0][col] != "N/A") {
            vField = "\"" + vField + "\"";
          }
          
          if (vField == ";" || vField == '";"') {
            vField = "";
          }
          
          if (vField != "" && vField != "\n") {
            vField += " ";
          }
          
          if (dataToExport[row].length-1 > col && (screenNo[0][col] != screenNo[0][col+1] && vField != "" && screenNo[0][col+1] != "C")) {
            vField = vField.trim();
            vField += "\n";
          }

          vLine += vField;
        }

        // Join each row's columns
        // Add a carriage return to end of each row, except for the last one
        if (row < dataToExport.length-1) {
          csv += vLine; // + "\n";
        }
        else {
          csv += vLine;
        }
      }
      csvFile = "@@batchload " + progName + ".p" + "\n" + csv + "\n" + "@@end" + "\n";
    }
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}


