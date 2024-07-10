/**
 * == A script to create the CIM load file for the supplier item and price screens in QAD. ==
 *
 * Logging and configuration functions adapted from the script:
 * 'A Google Apps Script for importing CSV data into a Google Spreadsheet' by Ian Lewis.
 *  https://gist.github.com/IanLewis/8310540
 * @author ianmlewis@gmail.com (Ian Lewis)
 * @author dunn.shane@gmail.com (Shane Dunn)
 * De Bortoli Wines July 2018
*/

/* =========== Globals ======================= */
var SPREADSHEET_ID = "1vPtIuT4HjOn4pcYzODDWweXOmxM-C7gGzQQ_SY3XqzM";
// var CONFIG_SHEET = 'Configuration';
var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
var now = new Date();

/* =========== Setup Menu ======================= */
/**
 * Create a Menu when the script loads. Adds a new configuration sheet if
 * one doesn't exist.
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('DBW Menu')
    .addItem('Create Excel Load Files', 'createExcelLoad')
    .addItem('Create CIM Load Files', 'createCIMload')
    .addSeparator()
    .addItem('Help and Support »', 'help')
    .addSeparator()
    .addItem('Authorise', 'init')
    .addToUi();

  var sheet = getOrCreateSheet_(CONFIG_SHEET);
  if (sheet.getRange(1, 1).getValue() == ""){
    loadNewConfiguration(sheet);
  }
  sheet = getOrCreateSheet_(LOG_SHEET);
  if (sheet.getRange(1, 1).getValue() == ""){
    loadNewLog(sheet);
  }
  sheet = getOrCreateSheet_(ERROR_SHEET);
  if (sheet.getRange(1, 1).getValue() == ""){
    loadNewError(sheet);
  }
}
