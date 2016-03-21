//adds Create PDFs menu and Item to Sheet
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Create PDFs').addItem('CreatePDFs', 'createInfoSheet').addToUi();
}


/* function to read sheet row by row, pull data from specified cells and save it, create a new sheet from template, 
   add saved data to new sheet, save new sheet as PDF in specified Google Drive Folder, delete new sheet. Loop for entire sheet
   until last row of data*/
function createInfoSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssKey = 'YOUR_KEY_HERE' // enter your spreadsheet key here
  //site list worksheet
  var siteList = ss.getSheetByName('Sites');
  //template worksheet
  var template = ss.getSheetByName('Template');
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Start row', 'Enter the row to start at', ui.ButtonSet.OK_CANCEL);
  
  // ask user what row to start at, else start at beginning row
  if (response.getSelectedButton() == ui.Button.OK) {
    var startRow = response.getResponseText(); 
  }  else {
    Logger.log("User did not enter a row");
    return;
  }
  
  //number of rows, set to the value of the last row with data
  var numRows = siteList.getLastRow();
  //var numRows = 3;
  //gets the range of data to process, B:Z, (starting row, starting column, num of rows to process, num of coll to process) 
  var dataRange = siteList.getRange(startRow, 1, numRows, 26);
  
  //fetch values for the range
  var data = dataRange.getValues();
  
  var folder = DriveApp.getFolderById('YOUR_FOLDER_ID_HERE'); //enter your Google Drive Folder ID Here
  
  /* loops through each row and pulls the revelant data from the range, storing them as variables, 
  creates a new sheet with the agency name as the name from the template sheet, fills the template, saves as a PDF to Google
  drive folder by ID, deletes the filled in sheet*/
  for (var i = 0; i < data.length; i++) {
    var column = data[i]
    
    //pulls values from site list and saves as variable to be copied to new shee later
    var agencyName = column[7]
    var agencyType = column[20]
    var agencyDescription = column[21]
    var agencyContactName = column[17]
    var agencyContactPhoneNumber = column[18]
    var agencyContactEmail = column[11]
    if (column[19] == '') {
      var agencyServiceAddress = column[13] + ' ' +column[14] + ', ' +column[15] + ' ' +column[16]
    }
    else {
      var agencyServiceAddress = column[19]
    }
    
    //creates a new copy of the template file with the agency name as the name
    var newSheet = ss.insertSheet(agencyName, {template: template});
    //var newSheet = ss.getSheetByName(agencyName);
    
    //get sheet ID for new sheet
    var id = newSheet.getSheetId();
    
    //fill new sheet with values from site list
    newSheet.getRange("A3").setValue(agencyName);
    newSheet.getRange("A9").setValue(agencyType);
    newSheet.getRange("A15").setValue(agencyDescription);
    newSheet.getRange("A21").setValue(agencyContactName);
    newSheet.getRange("A23").setValue(agencyContactPhoneNumber);
    newSheet.getRange("A25").setValue(agencyContactEmail);
    newSheet.getRange("A29").setValue(agencyServiceAddress);
    
    //refresh spreadsheet to save new values
    SpreadsheetApp.flush();
    
    //define the params URL to fetch
    var params = '?gid='+id
                 + '&fitw=true' //fit width of page
                 + '&exportFormat=pdf' //export as pdf
                 + '&format=pdf' //save as pdf
                 + '&size=Letter' //letter sized paper
                 + '&portrait=true' //portrait orietation 
                 + '&sheetnames=false' //print sheetname
                 + '&printtitle=false' //print title of worksheet
                 + '&gridlines=false'; //print gridlines
    
    var options = {headers: {'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()}}
    
    //add wait to avoice HTTP Response 429
    var backoff = 1
    Utilities.sleep((Math.pow(2,backoff)*1000) + (Math.round(Math.random() * 1000)));
    
    //creates a variable holding a random number, if random number is < .75, then increase backoff variable by 1 giving 75% each run increases backoff
    var random = 0
    random = Math.random()
    if (random < 0.75) {
      ++backoff
    }
    
    //fetching file url and naming it with the agency name
    var blob = UrlFetchApp.fetch("https://docs.google.com/spreadsheets/d/"+ssKey+"/export"+params, options);
    blob = blob.getBlob().setName(agencyName);
    
    //return file
    folder.createFile(blob);
    
    //delete new sheet
    ss.deleteSheet(ss.getSheetByName(agencyName))  

}
  //exit when done
  return;
}
