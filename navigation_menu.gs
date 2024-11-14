/**
 * Adds custom menu items to the Google Sheets UI when the document is opened.
 */
function onOpen() {
  // Create an "Admin" submenu with items for different departments
  var adminMenu = SpreadsheetApp.getUi().createMenu("Admin")
    .addItem("HR", "hr")               // Links to the HR sheet
    .addItem("Legal", "legal")         // Links to the Legal sheet
    .addItem("Facilities", "facilities") // Links to the Facilities sheet
    .addItem("Executive", "executive") // Links to the Executive sheet
    .addItem("PR", "pr");              // Links to the PR sheet
  
  // Create a "Departments" menu and add both department items and the Admin submenu
  SpreadsheetApp.getUi().createMenu("Departments")
    .addItem("Sales", "sales")         // Links to the Sales sheet
    .addItem("Marketing", "marketing") // Links to the Marketing sheet
    .addItem("Finance", "finance")     // Links to the Finance sheet
    .addItem("Engineering", "engineering") // Links to the Engineering sheet
    .addSubMenu(adminMenu)             // Adds the Admin submenu under Departments
    .addToUi();                        // Adds the Departments menu to the UI
}

/**
 * Activates a specified sheet by name.
 * @param {string} sheetName - The name of the sheet to activate.
 */
function setActiveSpreadsheet(sheetName) {
  SpreadsheetApp.getActive().getSheetByName(sheetName).activate();
}

// Below are individual functions to activate specific sheets based on department name

/**
 * Activates the Sales sheet.
 */
function sales() {
  setActiveSpreadsheet("Sales");
}

/**
 * Activates the Marketing sheet.
 */
function marketing() {
  setActiveSpreadsheet("Marketing");
}

/**
 * Activates the Finance sheet.
 */
function finance() {
  setActiveSpreadsheet("Finance");
}

/**
 * Activates the HR sheet.
 */
function hr() {
  setActiveSpreadsheet("HR");
}

/**
 * Activates the Engineering sheet.
 */
function engineering() {
  setActiveSpreadsheet("Engineering");
}

/**
 * Activates the Legal sheet.
 */
function legal() {
  setActiveSpreadsheet("Legal");
}

/**
 * Activates the Facilities sheet.
 */
function facilities() {
  setActiveSpreadsheet("Facilities");
}

/**
 * Activates the Executive sheet.
 */
function executive() {
  setActiveSpreadsheet("Executive");
}

/**
 * Activates the PR sheet.
 */
function pr() {
  setActiveSpreadsheet("PR");
}
