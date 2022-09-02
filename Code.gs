
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
  var day = getWeekDayName(date);
  // calls specific folder in google drive
  const fldr = DriveApp.getFolderById("1K5ufYV2JYtje96eOtLP-VJxDzj7pdIOb");
  Logger.log("got folder");
  // gets the file in the specific folder
  const files = fldr.getFiles()
  while (files.hasNext()) {
    var file = files.next();
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
  // gets range where names are and copies data
  var range = sheetsF.getRange("A3:A27");
  var values = range.getRichTextValues();
  // makes array to store names
  var reps = new Array()
  // saves names to the array
  for (var s = 0 in source) {
    // sleep
    Utilities.sleep(1000);
    var sheetName = source[s].getName()
    // gets individual sheet in raw data spreadsheet by index in array
    var sheetR = paulsSS.getSheetByName(sheetName);
    var total = "0"
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        var textO = values[i][j].getText();
        text = textO.toString().toLowerCase();
        if (text.split(" ")[0] == "team") {
          if (sheetName.toLowerCase().split(" ")[1] == text.split(" ")[1]) {
            Logger.log(sheetR.getName() + " = " + text);
            var tosearch = "Grand Total";
            var tf = sheetR.createTextFinder(tosearch);
            var cell = tf.findNext();

            if (cell == null) {
              Logger.log("cell == null ??");
              total = sheetsF.setCurrentCell(sheetsF.getRange('Y3')).getValue();
            } else {
              cell = cell.offset(0, 1);
              if (cell.getBackground() == '#ff0000') {
                total = cell.getValue();
              } else {
                while (cell.getBackground() != '#92d050') {
                  cell = cell.offset(1, 0);
                }
              }
            }
            total = cell.getValue();
            if (text.split(" ")[1] != "harmon") {
              Logger.log(text + " " + total);
              var col = sheetsF.createTextFinder(day).findNext().getColumn();
              var row = sheetsF.createTextFinder(textO).findNext().getRow();
              var curr = sheetsF.setCurrentCell(sheetsF.getRange(row, col));
              curr.setValue(total);
            }
          }
        } else {
          if (sheetName.toLowerCase().startsWith(text.split(" ")[0]) && sheetName.toLowerCase().split(" ")[1] == text.split(" ")[1]) {
            Logger.log(sheetR.getName() + " = " + text);
            var tosearch = "Grand Total";
            var tf = sheetR.createTextFinder(tosearch);
            var cell = tf.findNext();
            if (cell == null) {
              Logger.log("cell == null ??");
              total = sheetsF.setCurrentCell(sheetsF.getRange('Y3')).getValue();
            } else {
              cell = cell.offset(0, 1);
              if (cell.getBackground() == '#ff0000') {
                total = cell.getValue();
              } else {
                while (cell.getBackground() != '#92d050') {
                  cell = cell.offset(1, 0);
                }
              }
            }
            total = cell.getValue();
            if (text != "team harmon") {
              Logger.log(text + " " + total);
              var col = sheetsF.createTextFinder(day).findNext().getColumn();
              var row = sheetsF.createTextFinder(textO).findNext().getRow();
              var curr = sheetsF.setCurrentCell(sheetsF.getRange(row, col));
              curr.setValue(total);
            }
          }
        }
      }
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
  switch (dayOfWeek) {
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
  return day.toUpperCase();
}

function getPaulsReportFromEmail() {
  var query = 'from:"forward.cw.tax@gmail.com" is:unread';
  var userId = "me";
  var folder = DriveApp.getFolderById('1K5ufYV2JYtje96eOtLP-VJxDzj7pdIOb');
  var results = Gmail.Users.Messages.list("me", { q: query });
  results.messages.forEach(function (m) {
    Logger.log(m.id);
    var msg = GmailApp.getMessageById(m.id);
    var a = msg.getAttachments()[0];
    var fileName = a.getName();
    fileName = fileName.split(" ")[0];
    var file = folder.createFile(a.copyBlob()).setName(fileName);
    //msg.markMessageRead();
    name = file.getName();

  });
  return name;
}

