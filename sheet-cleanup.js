// This should have 1 header row.
var DB_SHEET_NAME = "BD"
// 0-indexed position of the email column in DB_SHEET
var COL_EMAIL = 0;

// This tab should just have a single column of email addresses, with no header
// These are the emails that will be removed by removeUnsubscribed()
var DELETE_SHEET_NAME = "To Delete";

// Sheet to archive unsubscribed entries
var UNSUBSCRIBED_SHEET_NAME = "Unsubscribed";

function removeDuplicates() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var newData = new Array();
  var deletedRows = 0;
  for(i in data){
    // Keep the header row.
    if (i == 0) {
      newData.push(data[i]);
      continue;
    }
    
    // If it's a new email compared to the previous row, keep this row.
    if (data[i][COL_EMAIL] != data[i-1][COL_EMAIL]) {
      newData.push(data[i]);
    } else {
      deletedRows += 1;
    }
  }
  
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
  Logger.log("Kept %s rows and deleted %s rows.", newData.length, deletedRows);
}



function removeUnsubscribed() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var deleteSheet = ss.getSheetByName(DELETE_SHEET_NAME);
  var deleteEmails = deleteSheet.getDataRange().getValues();
  // Flatten from 2d array (because data is in columns) to 1d
  // https://stackoverflow.com/questions/10865025/merge-flatten-an-array-of-arrays-in-javascript
  deleteEmails = [].concat.apply([], deleteEmails);
  // Quickly lowercase everything
  deleteEmails = deleteEmails.join('|').toLowerCase().split('|');
  deleteEmails = deleteEmails.sort();
  
  var dbSheet = ss.getSheetByName(DB_SHEET_NAME);
  var dbData = dbSheet.getDataRange().getValues();
  function sortByEmail(a, b) {
    a = a[COL_EMAIL].toLowerCase();
    b = b[COL_EMAIL].toLowerCase();
    return (a > b) - (a < b);
  }
  dbData.sort(sortByEmail);
  
  var newDbData = [dbData[0]]; // Save the header row
  var newUnsubData = []; // No header row because we append to the end of the sheet
  
  var dbCounter = 1; // Start comparing after header row
  var deleteCounter = 0;
  
  while(dbCounter < dbData.length && deleteCounter < deleteEmails.length) {
    if (dbData[dbCounter][COL_EMAIL].toLowerCase() == deleteEmails[deleteCounter]) {
      newUnsubData.push(dbData[dbCounter]);
      dbCounter++;
      deleteCounter++;
    } else {
      if (dbData[dbCounter][COL_EMAIL].toLowerCase() < deleteEmails[deleteCounter]) {
        newDbData.push(dbData[dbCounter]);
        dbCounter++;
      } else {
        deleteCounter++;
      }
    }
  }
  
  dbSheet.clearContents();
  dbSheet.getRange(2, 1, newDbData.length, newDbData[0].length).setValues(newDbData);
  
  var unsubSheet = ss.getSheetByName(UNSUBSCRIBED_SHEET_NAME);
  var unsubNextRow = unsubSheet.getLastRow() + 1
  unsubSheet.getRange(unsubNextRow, 1, newUnsubData.length, newUnsubData[0].length).setValues(newUnsubData);
}
