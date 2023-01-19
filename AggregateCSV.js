//This function import data from the csv files in the folder with id '1pP-c1fM7wN9CD6ZDTAgJGdtSt8PwZdB1'
function importCSV() {
  //The folder id where the csv files are located
  var folderId = "Source Folder ID";

  //Getting the active sheet of the current spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();

  //Getting the folder by folderId
  var folder = DriveApp.getFolderById(folderId);

  //Getting all the files inside the folder
  var files = folder.getFiles();

  //Empty array to store the data from the csv files
  var data = [];

  //Getting the values of the first column of the sheet
  var keywords = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();

  //Flattening the 2D array to 1D
  keywords = keywords.flat();

  //An array with the common headings that will be used to check against the headers in the csv files
  var commonHeadings = [
    "Keyword",
    "Volume",
    "Position",
    "Est. Visits",
    "CPC",
    "Paid Difficulty",
    "SEO Difficulty",
    "URL",
  ];

  //Creating a map of the common headings
  var map = {};
  for (var i = 0; i < commonHeadings.length; i++) {
    map[commonHeadings[i]] = commonHeadings[i];
  }

  // array to store the processed csv files
  var processedFiles = [];

  //Iterating through each file in the folder
  while (files.hasNext()) {
    var file = files.next();

    //Checking if the file has the csv extension
    if (file.getName().endsWith(".csv")) {
      //Getting the content of the csv file
      var csv = file.getBlob().getDataAsString();

      //Parsing the content of the csv file into a 2D array
      var csvData = Utilities.parseCsv(csv);

      //Getting the headers of the csv file
      var headers = csvData[0];

      // Create a 2D array to store the data
      var data = [
        [
          "Keyword",
          "Volume",
          "Position",
          "Est. Visits",
          "CPC",
          "Paid Difficulty",
          "SEO Difficulty",
          "URL",
        ],
      ];

      // Iterate through the csv files and add the data to the 2D array
      while (files.hasNext()) {
        var file = files.next();
        if (file.getName().endsWith(".csv")) {
          var csv = file.getBlob().getDataAsString();
          var csvData = Utilities.parseCsv(csv);
          var headers = csvData[0];
          for (var i = 1; i < csvData.length; i++) {
            var row = csvData[i];
            var obj = {};
            for (var j = 0; j < headers.length; j++) {
              if (headers[j] in map) {
                obj[map[headers[j]]] = row[j];
              }
            }
            if (!keywords.includes(obj["Keyword"])) {
              data.push([
                obj["Keyword"],
                obj["Volume"],
                obj["Position"],
                obj["Est. Visits"],
                obj["CPC"],
                obj["Paid Difficulty"],
                obj["SEO Difficulty"],
                obj["URL"],
              ]);
              keywords.push(obj["Keyword"]);
            }
          }
          // push current processed file to processedFiles array
          processedFiles.push(file);
        }
      }

      // Get the last row of the sheet
      var lastRow = sheet.getLastRow();

      // Write the data to the sheet starting from the last row + 1
      sheet
        .getRange(lastRow + 1, 1, data.length, data[0].length)
        .setValues(data);

      // move all processed files to the processed folder
      var parsedFolder = DriveApp.getFolderById(
        "Processed CSV Files folder ID"
      );
      for (var i = 0; i < processedFiles.length; i++) {
        processedFiles[i].moveTo(parsedFolder);
      }
    }
  }
}
