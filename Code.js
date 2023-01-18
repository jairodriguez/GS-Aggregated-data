function aggregateData() {
  // Get the folder containing the CSV files
  var folder = DriveApp.getFolderById("1pP-c1fM7wN9CD6ZDTAgJGdtSt8PwZdB1");

  // Create an empty array to store the data
  var data = [];

  // Create an empty array to store the headings
  var headings = [];

  // Get all the files in the folder
  var files = folder.getFiles();

  // Loop through the files
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    // Check if the file is a CSV
    if (fileName.substring(fileName.length - 4) === ".csv") {
      // Get the data from the CSV
      var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());

      // Add the headings from the current file to the headings array
      for (var i = 0; i < csvData[0].length; i++) {
        if (!headings.includes(csvData[0][i])) {
          headings.push(csvData[0][i]);
        }
      }

      // Loop through the rows of data
      for (var i = 1; i < csvData.length; i++) {
        var row = csvData[i];

        // Check if the row already exists in the data array
        var duplicate = false;
        for (var j = 0; j < data.length; j++) {
          if (row.join() === data[j].join()) {
            duplicate = true;
            break;
          }
        }

        // If the row is not a duplicate, add it to the data array
        if (!duplicate) {
          data.push(row);
        }
      }
    }
  }

  // Get the active sheet
  var sheet = SpreadsheetApp.getActiveSheet();

  // Clear the sheet
  sheet.clear();

  // Add the headings to the sheet
  sheet.appendRow(headings);

  // Add the data to the sheet
  for (var i = 0; i < data.length; i++) {
    sheet.appendRow(data[i]);
  }
}
