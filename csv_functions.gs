// From http://www.greywyvern.com/?post=258 - Brian Huisman AKA GreyWyvern
String.prototype.splitCSV = function(sep) {
  for (var foo = this.split(sep = sep || ","), x = foo.length - 1, tl; x >= 0; x--) {
    if (foo[x].replace(/"\s+$/, '"').charAt(foo[x].length - 1) == '"') {
      if ((tl = foo[x].replace(/^\s+"/, '"')).length > 1 && tl.charAt(0) == '"') {
        foo[x] = foo[x].replace(/^\s*"|"\s*$/g, '').replace(/""/g, '"');
      } else if (x) {
        foo.splice(x - 1, 2, [foo[x - 1], foo[x]].join(sep));
      } else foo = foo.shift().split(sep).concat(foo);
    } else foo[x].replace(/""/g, '"');
  } return foo;
};


function saveAsTabDelimitedTextFile() {
  // get Spreadsheet Name
  var fileName = SpreadsheetApp.getActiveSheet().getSheetName();

  // Add the ".txt" extension to the file name
  fileName = fileName + ".txt";

  // Convert the range data to tab-delimited format
  var txtFile = convertRangeToTxtFile_(fileName);

  // Delete existing file
  deleteDocByName(fileName);

  // Create a file in the Docs List with the given name and the data
  DocsList.createFile(fileName, txtFile);
}

function convertRangeToTxtFile_(txtFileName) {
  try {
    var txtFile = undefined;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var rows = sheet.getDataRange();
    var data = rows.getValues();

    // Loop through the data in the range and build a string with the data
    if (data.length > 1) {
      var txt = "";
      for (var row = 0; row < data.length; row++) {
        // Join each row's columns and add a carriage return to end of each row
        txt += data[row].join("\t") + "\r\n";
      }
      txtFile = txt;
    }
    return txtFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}

//return ContentService.createTextOutput("bbb,aaa,ccc").downloadAsFile("MyData.csv").setMimeType(ContentService.MimeType.CSV);

function saveAsCSV2() {
  
  // Name the file
  var fileName = "apvomt2.csv";
  // Convert the range data to CSV format
  var rangeToExport = SpreadsheetApp.getActiveSpreadsheet().getRange("A11:BE103");
  var csvFile = exportToCsv(fileName, rangeToExport);
  // Create a file in the root of my Drive with the given name and the CSV data
  DriveApp.createFile(fileName, csvFile);
}


function exportToCsv(filename, rows) {
        var processRow = function (row) {
            var finalVal = '';
            for (var j = 0; j < row.length; j++) {
                var innerValue = row[j] === null ? '' : row[j].toString();
                if (row[j] instanceof Date) {
                    innerValue = row[j].toLocaleString();
                };
                var result = innerValue.replace(/"/g, '""');
                if (result.search(/("|,|\n)/g) >= 0)
                    result = '"' + result + '"';
                if (j > 0)
                    finalVal += ',';
                finalVal += result;
            }
            return finalVal + '\n';
        };

        var csvFile = '';
        for (var i = 0; i < rows.length; i++) {
            csvFile += processRow(rows[i]);
        }

        /*
        var blob = new Blob([csvFile], { type: 'text/csv;charset=utf-8;' });
        if (navigator.msSaveBlob) { // IE 10+
            navigator.msSaveBlob(blob, filename);
        } else {
            var link = document.createElement("a");
            if (link.download !== undefined) { // feature detection
                // Browsers that support HTML5 download attribute
                var url = URL.createObjectURL(blob);
                link.setAttribute("href", url);
                link.setAttribute("download", filename);
                link.style.visibility = 'hidden';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            }
        }
        */
      return csvFile;
    }
