/**
 * Logging and configuration functions adapted from the script:
 * 'A Google Apps Script for importing CSV data into a Google Spreadsheet' by Ian Lewis.
 *  https://gist.github.com/IanLewis/8310540
 * @author ianmlewis@gmail.com (Ian Lewis)
 * @author dunn.shane@gmail.com (Shane Dunn)
 * De Bortoli Wines July 2017
*/
/* =========== Globals ======================= */
var CONFIG_SHEET = 'Configuration';

/* =========== Configuration functions ======================= */
/**
 * Returns the values from 2 columns from the csvconfig sheet starting at
 * colIndex, as key-value pairs. Key-values are only returned if they do
 * not contain the empty string or have a boolean value of false.
 * If the key is start-date or end-date and the value is an instance of
 * the date object, the value will be converted to a string in yyyy-MM-dd.
 * If the key is start-index or max-results and the type of the value is
 * number, the value will be parsed into a string.
 * If value is "ColumnNames", the subsequent key value pairs are treated
 * as an array of values (used for column names)
 * @param {number} colIndex The column index to return values from.
 * @return {object} The values starting in colIndex and the following column
       as key-value pairs.
 */
function getConfigsStartingAtCol_(sheet, colIndex) {
  var config = {}, rowIndex, key, value, columnDef, tblName;
  var range = sheet.getRange(1, colIndex, sheet.getLastRow(), 2);

  columnDef = false;

  for (rowIndex = 2; rowIndex <= range.getLastRow(); ++rowIndex) {
    key = range.getCell(rowIndex, 1).getValue();
    value = escapeQuotes(range.getCell(rowIndex, 2).getValue());
    if (value) {
      var trailNum = getTrailingNumber_(key)
      if ( columnDef && trailNum ) {
        config[tblName][trailNum] = escapeQuotes(value);
      } else {
        columnDef = false;
      }
      if ( columnDef || value == "ColumnNames") {
        if ( value == "ColumnNames") {
          tblName = key;
          columnDef = true;
          config[tblName] = [];
        }
      } else {
        config[key] = value;
      }
    }
  }
  return config;
}

/**
 * Returns an array of config objects. This reads the csvconfig sheet
 * and tries to extract adjacent column names that end with the same
 * number. For example Names1 : Values1. Then both columns are used
 * to define key-value pairs for the coniguration object. The first
 * column defines the keys, and the adjacent column values define
 * each keys values.
 * @param {Sheet} The csvconfig sheet from which to read configurations.
 * @returns {Array} An array of API query configuration object.
 */
function getConfigs_(sheet) {

  var configs = [], colIndex;
  // There must be at least 2 columns.
  if (sheet.getLastColumn() < 2) {
    return configs;
  }

  var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  var firstColValue, firstColNum, secondColValue, secondColNum;

  // Test the name of each column to see if it has an adjacent column that ends
  // in the same number. ie xxxx555 : yyyy555.
  // Since we check 2 columns at a time, we don't need to check the last column,
  // as there is no second column to also check. 
  for (colIndex = 1; colIndex <= headerRange.getNumColumns() - 1; ++colIndex) {
    firstColValue = headerRange.getCell(1, colIndex).getValue();
    firstColNum = getTrailingNumber_(firstColValue);
    
    secondColValue = headerRange.getCell(1, colIndex + 1).getValue();
    secondColNum = getTrailingNumber_(secondColValue);
  
    if (firstColNum && secondColNum && firstColNum === secondColNum) {
      configs.push(getConfigsStartingAtCol_(sheet, colIndex)); 
    }
  }
  return configs;  
}

/**
 * Returns 1 greater than the largest trailing number in the header row.
 * @param {Object} sheet The sheet in which to find the last number.
 * @returns {Number} The next largest trailing number.
 */
function getLastNumber_(sheet) {
  var maxNumber = 0;
  
  var lastColIndex = sheet.getLastColumn();

  if (lastColIndex > 0) {
    var range = sheet.getRange(1, 1, 1, lastColIndex);

    for (var colIndex = 1; colIndex < sheet.getLastColumn(); ++colIndex) {
      var value = range.getCell(1, colIndex).getValue();
      var headerNumber = getTrailingNumber_(value);
      if (headerNumber) {
        var number = parseInt(headerNumber, 10);
        maxNumber = number > maxNumber ? number : maxNumber;                                  
      }
    }
  }
  return maxNumber + 1;
}

/**
 * Returns the trailing number on a string. For example the
 * input: xxxx555 will return 555. Inputs with no trailing numbers
 * return undefined. Trailing whitespace is not ignored.
 * @param {string} input The input to parse.
 * @resturns {number} The trailing number on the input as a string.
 *     undefined if no number was found.
 */
function getTrailingNumber_(input) {
  // Match at one or more digits at the end of the string.
  var pattern = /(\d+)$/;
  var result = pattern.exec(input);
  if (result) {
    // Return the matched number.
    return result[0];
  }
  return undefined;
}

function escapeQuotes(value) {
  if (!value) {
    return "";
  }
  if (typeof value != 'string') {
    value = value.toString();
  }
  return value.replace(/\\/g, '\\\\').replace(/'/g, "\\\'");
}

function getOrCreateSheet_(sheet_name) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName(sheet_name);
  if (!sheet) {
    sheet = activeSpreadsheet.insertSheet(sheet_name, 0);
  }
  return sheet;
}

/**
 * Adds a configuration to the spreadsheet.
 */
function createConfig() {
  loadNewConfiguration(getOrCreateSheet_(CONFIG_SHEET));
}

function loadNewConfiguration(sheet) {
  var headerNumber = getLastNumber_(sheet);
  var config = [
    ["file_name_" + headerNumber, "value_" + headerNumber],
    ['file_name', 'iclomt_cim.txt'],
    ['sheet_name', 'CIM Load'],
    ['exportRange', 'A12:BE104'],
    ['screenNoRow', 'A3:BE3'],
    ['variableTypeRow', 'A5:BE5'],
    ['program_name', 'B2:B2']
  ];
  //Logger.log(config);
  sheet.getRange(1, sheet.getLastColumn() + 1, config.length, 2).setValues(config);
}

/*
  var config = [
    ['ConfigKey', 'Value'],
    ['data-table-name', 'Flocom Data - Data'],
    ['event-table-name', 'Flocom Data - Events'],
    ['load-log-sheet-name', 'Log'],
    ['error-sheet-name', 'errors'],
    ['data-table-name-columns', 'ColumnNames'],
    ['data-table-column1', 'Reading Date/Time'],
    ['data-table-column2', 'Status'],
    ['data-table-column3', 'Flow Rate (Ml/day)'],
    ['data-table-column4', 'Actual Date/Time'],
    ['data-table-column5', 'Last Total'],
    ['data-table-column6', 'Ml/min'],
    ['data-table-column7', 'Interval (min)'],
    ['data-table-column8', 'Flow (Ml)'],
    ['data-table-column9', 'Adjusted Total'],
    ['data-table-column10', 'Cumulative Total'],
    ['event-table-name-columns', 'ColumnNames'],
    ['event-table-column1', 'Event Date/Time'],
    ['event-table-column2', 'Event'],
    ['event-table-column3', 'Result'],
  ];
*/

function testConfig() {
  Logger.log(getConfigs_(getOrCreateSheet_(CONFIG_SHEET)));
}
