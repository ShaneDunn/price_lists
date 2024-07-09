/**
 * Logging and configuration functions adapted from the script:
 * 'A Google Apps Script for importing CSV data into a Google Spreadsheet' by Ian Lewis.
 *  https://gist.github.com/IanLewis/8310540
 * @author ianmlewis@gmail.com (Ian Lewis)
 * @author dunn.shane@gmail.com (Shane Dunn)
 * De Bortoli Wines July 2017
*/
/* =========== Globals ======================= */
/**
 * The output text that should be displayed in the log.
 * @private.
 */
var logArray_;
var errorArray_;
var LOG_SHEET = 'Log';
var ERROR_SHEET = 'Error';
var VERBOSE = 3; // 1: INFO , 2: WARNING , 3: ERROR
/* =========== Logging functions ======================= */

/**
 * Clears the in app log.
 * @private.
 */
function setupLog_() {
  logArray_ = [];
  errorArray_ = [];
}

/**
 * Returns the log as a string.
 * @returns {string} The log.
 */
function getLog_() {
  return logArray_.join('\n'); 
}

/**
 * Returns the log as a HTML string.
 * @returns {string} The log.
 */
function getHTMLLog_() {
  /*
  var textarray = [];
  for (var i = 0; i < logArray_.length; i++) { 
    textarray[i] = "<p>" + logArray_[i] + "</p>"; 
  } 
  */
  return logArray_.join('<br>\n'); 
}

/**
 * Returns the error log as a string.
 * @returns {string} The error log.
 */
function getError_() {
  return errorArray_.join('\n'); 
}

/**
 * Appends a string as a new line to the log.
 * @param {String} value The value to add to the log.
 */
function log_(value) {
  logArray_.push(value);

  var now = new Date();
  if (value.substring(0, 4) == "INFO" && VERBOSE >= 1) {
    errorArray_.push([now, value]);
  }
  if (value.substring(0, 7) == "WARNING" && VERBOSE >= 2) {
    errorArray_.push([now, value]);
  }
  if (value.substring(0, 5) == "ERROR" && VERBOSE >= 3) {
    errorArray_.push([now, value]);
  }
  /*
  var app = UiApp.getActiveApplication();
  var foo = app.getElementById('log');
  foo.setText(getLog_());
  */
}

/**
 * Displays the log in memory to the user.
 */
function displayLog_() {
  var uiLog = UiApp.createApplication().setTitle('Report Status').setWidth(400).setHeight(500);
  var panel = uiLog.createVerticalPanel();
  uiLog.add(panel);

  var txtOutput = uiLog.createTextArea().setId('log').setWidth('400').setHeight('500').setValue(getLog_());
  panel.add(txtOutput);
  
  SpreadsheetApp.getActiveSpreadsheet().show(uiLog); 
}

function showLogDialog_() {
  var html = HtmlService.createHtmlOutput(getHTMLLog_())
      .setWidth(400)
      .setHeight(500);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'CIM Load File Status');
}

function dumpLog_(sheet) {
  var lastRow = sheet.getLastRow() + 1;
  if (logArray_.length != 0) {
    var array = logArray_.map(function (el) {
          return [el];
    });
    sheet.getRange(lastRow,1,array.length,array[0].length).setValues(array);
    lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow,1).setValue(".").setBackground("#ffe2c6");
  }
}

function dumpError_(sheet) {
  var lastRow = sheet.getLastRow() + 1;
  if (errorArray_.length != 0) {
    sheet.getRange(lastRow,1,errorArray_.length,errorArray_[0].length).setValues(errorArray_);
    lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow,1).setValue(".").setBackground("#ffe2c6");
    sheet.getRange(lastRow,2).setValue(" ").setBackground("#ffe2c6");
  }
}

/**
 * Adds a Heading to the Log sheet.
 */
function createLog() {
  loadNewLog(getOrCreateSheet_(LOG_SHEET));
}

function loadNewLog(sheet) {
  sheet.getRange(1, 1).setValue("Logs");
  sheet.getRange(1, 1).setBackground("#ffe2c6").setFontColor("#976759").setFontFamily('Verdana').setFontSize(12).setFontWeight("Bold");
}

/**
 * Adds a Heading to the Error sheet.
 */
function createError() {
  loadNewError(getOrCreateSheet_(ERROR_SHEET));
}

function loadNewError(sheet) {
  sheet.getRange(1, 1).setValue("Errors");
  sheet.getRange(1, 1).setBackground("#ffe2c6").setFontColor("#976759").setFontFamily('Verdana').setFontSize(12).setFontWeight("Bold");
  sheet.getRange(1, 2).setBackground("#ffe2c6").setFontColor("#976759").setFontFamily('Verdana').setFontSize(12).setFontWeight("Bold");
}
