/* The purpose of this script is to take data from the attached spreadsheet
and create new folders in a named folder in Google Drive. The folder name is taken from cell C2. 
The engine of this code was taken from +Tony Hirst's blog post 
http://blog.ouseful.info/2010/03/04/maintaining-google-calendars-from-a-google-spreadsheet/ 

I haved tailored one or two things in the meantime. Using code mainly from:
http://stackoverflow.com/questions/18044152/google-script-create-and-publicly-share-a-folder-within-existing-parent
Finally, and many thanks to +Alexander Ivanov who tidied up my code including some much needed error handling.
https://plus.google.com/115872203173779243836/posts/4MZzkCox9S9
Eamon Collins
22/10/2014 - 23/10/2014 */

function onOpen() { //This is the new standard script for the onOpen trigger that creates a menu item.
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('GDrive')
      .addItem('Create new Folders', 'crtGdriveFolder')
      .addToUi();
}

function crtGdriveFolder() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = sheet.getLastRow();   // Number of rows to process
  var maxRows = Math.min(numRows,20); //Limit the number of rows to prevent enormous number of folder creations
  var folderid = sheet.getRange("C2").getValue();
  var root = sheet.getRange("D2").getValue();
  var dataRange = sheet.getRange(startRow, 1, maxRows, 2); //startRow, startCol, endRow, endCol
  var data = dataRange.getValues();
  var folderIterator = DriveApp.getFoldersByName(folderid); //get the file iterator
  
  if(!folderIterator.hasNext()) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Folder not found!');
    return;
  }
  
  var parentFolder = folderIterator.next();
  
  if(folderIterator.hasNext()) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Folder has a non-unique name!');
    return;
  }
  
  for (i in data) {
    var row = data[i];
    var name = row[0];      // column A
    var desc = row[1];      // column B
    

    
    
  if(root == "N" && name != "") {
      var idNewFolder = parentFolder.createFolder(name).setDescription(desc).getId();
      Utilities.sleep(100);
      var newFolder = DriveApp.getFolderById(idNewFolder); 

      
      } if(root == "Y" && name != "") {
          var idNewFolder = DriveApp.createFolder(name).setDescription(desc).getId();
          Utilities.sleep(100);
          var newFolder = DriveApp.getFolderById(idNewFolder);
         
          }
            
    }
};﻿
