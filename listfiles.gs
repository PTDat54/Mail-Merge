


function listAllFilesInFolder() {
 
var ui = SpreadsheetApp.getUi(); // Same variations.

  var counter = 0;
  var result = ui.prompt(
      'Folder to display',
      'Please enter folder ID',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.insertSheet();
  sheet.setName('File List ' + Math.random());
  sheet.appendRow(["Your Key","File Name", "File ID", "Last Updated", "Size"]); 
  var folderId = text;
  // Log the name of every file in the folder.
  var files = DriveApp.getFolderById(folderId).getFiles();
  while (files.hasNext()) 
        {
         var file = files.next();
         sheet.appendRow(["",file.getName(),file.getId(),file.getLastUpdated(),file.getSize()]);
         counter++;
        }
   SpreadsheetApp.getActiveSpreadsheet().toast((counter) + ' files listed','Complete');   
    
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
   
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.   
  }

}//end function