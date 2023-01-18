//This function import data from the csv files in the folder with id '1pP-c1fM7wN9CD6ZDTAgJGdtSt8PwZdB1'
function importCSV() {
  
  //The folder id where the csv files are located
  var folderId = '1pP-c1fM7wN9CD6ZDTAgJGdtSt8PwZdB1';
  
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
  var commonHeadings = ["Keyword", "Volume", "Position", "Est. Visits", "CPC", "Paid Difficulty", "SEO Difficulty"];
  
  //Creating a map of the common headings
  var map = {};
  for (var i = 0; i < commonHeadings.length; i++) {
    map[commonHeadings[i]] = commonHeadings[i];
  }

  //Iterating through each file in the folder
  while (files.hasNext()) {
    var file = files.next();
    
    //Checking if the file has the csv extension
    if (file.getName().endsWith('.csv')) {
      
      //Getting the content of the csv file
      var csv = file.getBlob().getDataAsString();
      
      //Parsing the content of the csv file into a 2D array
      var csvData = Utilities.parseCsv(csv);
      
      //Getting the headers of the csv file
      var headers = csvData[0];
      
      //Iterating through each row of the csv file
      for (var i = 1; i < csvData.length; i++) {
        var row = csvData[i];
        var obj = {};
        
        //Iterating through each column of the current row
        for (var j = 0; j < headers.length; j++) {
          
          //Checking if the current header is present in the map object
          if (headers[j] in map) {
            obj[map[headers[j]]] = row[j];
          }
        }
        //Checking if the current row is already present in the sheet
        if (!keywords.includes(obj["Keyword"])) {
          
          //If it's not present, adding the current row to the data array
          data.push(obj);
          
          //Appending the current row to the sheet
          sheet.appendRow
          ([obj["Keyword"], obj["Volume"], obj["Position"], obj["Est. Visits"], obj["CPC"], obj["Paid Difficulty"], obj["SEO Difficulty"]]);
          //Adding the keyword to the keywords array
          keywords.push(obj["Keyword"]);
        }
      }
    }
    // Move the file to the parsed folder
    var parsedFolder = DriveApp.getFolderById("1TgdPCEOGSNm4xE1psbaCfsWUT6_BJANf");
    file.getParents().next().removeFile(file);
    parsedFolder.addFile(file);
    
  }
}
