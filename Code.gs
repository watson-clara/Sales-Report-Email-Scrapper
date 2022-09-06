
function main() {
  Logger.log("STARTING");
  var name1 = getPaulsReportFromEmail();
  var name = getSalesReportSS(name1);
  getData(name1, name);
}

function getSalesReportSS(name) {
  // calls specific folder in google drive
  const fldr = DriveApp.getFolderById("1K5ufYV2JYtje96eOtLP-VJxDzj7pdIOb");
  Logger.log("got folder");
  // gets the file in the specific folder
  const files = fldr.getFilesByName(name);
  Logger.log("got files");
  var file = files.next();
  // get file id
  var id = file.getId();
  // get data in file
  var blob = file.getBlob();
  // sleep
  Utilities.sleep(1000);
  // creates new google sheet file 
  var newFile = {
    title: name,
    parents: [{ id: "1K5ufYV2JYtje96eOtLP-VJxDzj7pdIOb" }],
    mimeType: MimeType.GOOGLE_SHEETS
  };
  Logger.log("new SS file created");
  // copies the xlsx content to the new file 
  Drive.Files.insert(newFile, blob);
  // deletes the xlsx file 
  Drive.Files.remove(id);
  Logger.log("old files deleted")
  // returns new name 
  return name
}


function sortArray(allsheets) {
  // creates empty array to store sheet names
  var sheetNameArray = [];
  // loop through sheets and add each name to array
  for (var i = 0; i < allsheets.length; i++) {
    sheetNameArray.push(allsheets[i].getName());
  }
  // sort the array 
  sheetNameArray.sort(function (a, b) {
    return a.localeCompare(b);
  });
  Logger.log(sheetNameArray);
  // return sorted array
  return sheetNameArray;
}


function getData(name, date) {
  // calls the function to get day of week from the date 
  var day = getWeekDayName(date);
  // calls specific folder in google drive
  const fldr = DriveApp.getFolderById("1K5ufYV2JYtje96eOtLP-VJxDzj7pdIOb");
  Logger.log("got folder");
  // gets the file in the specific folder
  const files = fldr.getFiles()
  // finds next file
  while (files.hasNext()) {
    var file = files.next();
    // if the file has the right name then get the url
    if (file.getName() == name) {
      var url = file.getUrl();
      Logger.log(url);
      Logger.log(file.getName());
      break
    }
  }
  // open spreadsheet with id
  var paulsSS = SpreadsheetApp.openByUrl(url);
  // gets individual sheets within pauls spreadsheet
  var source = paulsSS.getSheets();
  //  template spreadsheet
  var template = SpreadsheetApp.openById("1r6sXjeSnrRiTcfYe7oqmoXlTiOE7_DHit6kQMj-QUww");
  // makes copy of report template spreadsheet and renames it 
  var ssT = template.copy("SALES REPORT " + date);
  Logger.log(ssT.getName());
  // open new final spreadsheet 
  var ssF = SpreadsheetApp.openByUrl(ssT.getUrl());
  // get array of names
  var sheetsF = ssF.getSheetByName("SALES");
  // gets first range of names and saves it 
  var range1 = sheetsF.getRange("A3:A27");
  var values1 = range1.getRichTextValues();
  // gets second range of names and saves it 
  var range2 = sheetsF.getRange("A39:A45");
  var values2 = range2.getRichTextValues();
  // gets second range of names and saves it 
  var range3 = sheetsF.getRange("A52:A59");
  var values3 = range3.getRichTextValues();
  //calls helper function
  getDataHelper(paulsSS, sheetsF, day, source, values1, 1);
  getDataHelper(paulsSS, sheetsF, day, source, values2, 2);
  getDataHelper(paulsSS, sheetsF, day, source, values3, 3);
}


function getDataHelper(paulsSS, sheetsF, day, source, values, num) {
  for (var s = 0 in source) {
    // get the name of each sheet from the array of sheet names
    var sheetName = source[s].getName()
    // gets individual sheet by sheet name obtaind from array
    var sheetR = paulsSS.getSheetByName(sheetName);
    // creates the total variable
    var total;
    // since the names from the template are saved in a 2D array we just use a nest loop
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        // gets the text from each cell saved in the 2D array
        var textO = values[i][j].getText();
        // cnverts text to lower case
        text = textO.toString().toLowerCase();
        // checks if the name is a individual or a team
        if (text.split(" ")[0] == "team") {
          // calls function to put the data in the new SS3
          putDataReps(sheetName.toLowerCase().split(" ")[1] == text.split(" ")[1], sheetR, sheetsF, text, textO, total, day);
        } else {
          if (num == 1) {
            // calls function to put the data in the new SS
            putDataReps(sheetName.toLowerCase().startsWith(text.split(" ")[0]) && sheetName.toLowerCase().split(" ")[1] == text.split(" ")[1], sheetR, sheetsF, text, textO, total, day);
          } else if (num == 2) {
            putDataDialers(sheetR, sheetName, sheetsF, text, textO, total, day);
          } else if (num == 3) {
            putDataDealDesk(sheetR, sheetName, sheetsF, text, textO, total, day);
          }
        }

      }
    }

  }
}
function putDataDealDesk(sheetR, sheetName, sheetsF, text, textO, total, day) {
  if (sheetName.toLowerCase().startsWith(text.split(" ")[1]) && sheetName.toLowerCase().split(" ")[1] == text.split(" ")[2]) {
    Logger.log("DEAL DESK");
    // looks for these words in Pauls SS by searching for them 
    var tosearch = "SPLIT COMMISSIONS";
    var tf = sheetR.createTextFinder(tosearch);
    var c = tf.findNext();
    if (c != null) {
      c = c.offset(1, 0);
      var row = c.getRow();
      Logger.log(row);

     
      Logger.log(sheetR.getName() + " = " + text);
      var cell = sheetR.setCurrentCell(sheetR.getRange(row, 5));

      total = 0;
      while (!cell.isBlank()) {
        if (cell.getBackground() != '#92d050'){
          total = total + cell.getValue();
        }
        Logger.log("while " + cell.getRow());
        cell = cell.offset(1, 0);
        
      }
      Logger.log(text + " " + total);
      // gets the colum of the day of the week we are finding 
      var col = sheetsF.createTextFinder(day).findNext().getColumn();
      // gets the row of the person or team's name 
      var row = sheetsF.createTextFinder(textO).findNext().getRow();
      // cross references row and colum and sets that cell as the current
      var curr = sheetsF.setCurrentCell(sheetsF.getRange(row, col));
      // adds the data to the current cell
      curr.setValue(total);
    }
  }
}


function putDataDialers(sheetR, sheetName, sheetsF, text, textO, total, day) {
  if (sheetName.toLowerCase().startsWith(text.split(" ")[0]) && sheetName.toLowerCase().split(" ")[1] == text.split(" ")[1]) {
    if (sheetName != "Mike Andrews") {
      Logger.log("DIALERS");
      Logger.log(sheetName);
      Logger.log(text);
      Logger.log(sheetR.getName() + " = " + text);
      var cell = sheetR.setCurrentCell(sheetR.getRange('E1'));
      var red = { Value: cell.getBackground() == '#ff0000' ? true : false };
      var green = { Value: cell.getBackground() == '#92d050' ? true : false };
      // move down one cell until we find a green or red cell
      while (cell.getBackground() != '#92d050') {
        cell = cell.offset(1, 0);
      }
      // get the value in the cell 
      total = cell.getValue();
      Logger.log(text + " " + total);
      // gets the colum of the day of the week we are finding 
      var col = sheetsF.createTextFinder(day).findNext().getColumn();
      // gets the row of the person or team's name 
      var row = sheetsF.createTextFinder(textO).findNext().getRow();
      // cross references row and colum and sets that cell as the current
      var curr = sheetsF.setCurrentCell(sheetsF.getRange(row, col));
      // adds the data to the current cell
      curr.setValue(total);
    }
  }
}

function putDataReps(logicalStatement, sheetR, sheetsF, text, textO, total, day) {
  // uses the loical statement passed thrugh as a parameter to test the if statement 
  if (logicalStatement) {
    Logger.log(sheetR.getName() + " = " + text);
    // looks for these words in Pauls SS by searching for them 
    var tosearch = "Grand Total";
    var tf = sheetR.createTextFinder(tosearch);
    var cell = tf.findNext();
    // checks if cell is null to avoid error aka search words dont appear on SS
    // or that name doesn't have a matching SS
    if (cell == null) {
      // if there is no cell then set total to 0
      total = 0;
    } else {
      // gets the cell to the right of the grand total cell
      cell = cell.offset(0, 1);
      // checks to see if the background is red
      if (cell.getBackground() == '#ff0000') {
        // if it is then get the value
        total = cell.getValue();
      } else {
        // overwise we are looking for a green cell and we must move down one cell until we find it 
        while (cell.getBackground() != '#92d050') {
          cell = cell.offset(1, 0);
        }
      }
    }
    // get the value in the cell 
    total = cell.getValue();
    // harmon has a different procedure so we must not enter data for team harmon
    if (!text.includes("team harmon")) {
      Logger.log(text + " " + total);
      Logger.log(day);
      Logger.log(textO);
      // gets the colum of the day of the week we are finding 
      var col = sheetsF.createTextFinder(day).findNext().getColumn();
      // gets the row of the person or team's name 
      var row = sheetsF.createTextFinder(textO).findNext().getRow();
      // cross references row and colum and sets that cell as the current
      var curr = sheetsF.setCurrentCell(sheetsF.getRange(row, col));
      // adds the data to the current cell
      curr.setValue(total);
    }
  }
}


function getWeekDayName(date) {
  // get current date
  var exampleDate = new Date(date);
  Logger.log('date is: ' + exampleDate);
  // get the weekday number from the current date
  var dayOfWeek = exampleDate.getDay();
  Logger.log('weekday number is: ' + dayOfWeek);
  // use a 'switch' statement to calculate the weekday name from the weekday number
  switch (5) {
    case 0:
      day = "Sunday";
      break;
    case 1:
      day = "Monday";
      break;
    case 2:
      day = "Tuesday";
      break;
    case 3:
      day = "Wednesday";
      break;
    case 4:
      day = "Thursday";
      break;
    case 5:
      day = "Friday";
      break;
    case 6:
      day = "Saturday";
  }
  // log the output to show the weekday name
  Logger.log('weekday name is: ' + day);
  // returns the day in uppercase becasue that is how it appears on the SS
  return day.toUpperCase();
}

function getPaulsReportFromEmail() {
  // gets the folder for storing pauls spreadsheet with the raw data
  var folder = DriveApp.getFolderById('1K5ufYV2JYtje96eOtLP-VJxDzj7pdIOb');
  // searches for emails that meet the query and creates a list
  var query = 'from:"forward.cw.tax@gmail.com" is:unread';
  var results = Gmail.Users.Messages.list("me", { q: query });
  // loops through the list of emails
  results.messages.forEach(function (m) {
    // gets the id of each email
    var msg = GmailApp.getMessageById(m.id);
    Logger.log(m.id);
    // gets the attachement for each eamil
    var a = msg.getAttachments()[0];
    // gets name of attachement 
    var fileName = a.getName();
    fileName = fileName.split(" ")[0];
    // copies the data from the attachement and creates a new file and puts the data in it and renames it 
    var file = folder.createFile(a.copyBlob()).setName(fileName);
    //msg.markMessageRead();
    // gets the name of the new file made becasue returning fileName was giving an error
    name = file.getName();

  });
  return name;
}

