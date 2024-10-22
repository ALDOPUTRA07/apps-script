function onOpen() {
  // This function is triggered when the spreadsheet is opened.
  // It calls the function to create a custom menu in the UI.
  createEmptyMenu();
}

function createEmptyMenu() {
  // Create a custom menu named "⚙ Additional Menu" in the Google Sheets UI.
  var menu = SpreadsheetApp.getUi().createMenu("⚙ Additional Menu");
  
  // Add a menu item "Hide Column" that calls the hideColumn function.
  menu.addItem("Hide Column", "hideColumn");
  
  // Add the constructed menu to the UI.
  menu.addToUi();
}

function hideColumn() {
  // Get the active spreadsheet and the specific sheet named 'Data'.
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Data');

  // Get the last populated row and last populated column in the sheet.
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  // Ensure all columns are visible before applying the hide logic.
  sheet.showColumns(1, lastColumn);

  // Loop through columns starting from column 4 to the last column.
  for (var i = 4; i <= lastColumn; i++) {
    
    // Check if the header value in the current column is not today or yesterday's date.
    if (
      (sheet.getRange(1, i).getDisplayValues()[0][0] != 
        Utilities.formatDate(new Date(), "GMT+7", "dd MMM YY").toString()) 
      &&
      (sheet.getRange(1, i).getDisplayValues()[0][0] != 
        Utilities.formatDate(new Date(new Date().getTime() - 1 * 24 * 60 * 60 * 1000), "GMT+7", "dd MMM YY").toString())
    ) {
      // Hide the column if it doesn't match the required dates.
      sheet.hideColumns(i, 1);
      
      // Log the header value of the hidden column for debugging purposes.
      Logger.log(sheet.getRange(1, i).getDisplayValues()[0][0]);
    } 
  }
}
