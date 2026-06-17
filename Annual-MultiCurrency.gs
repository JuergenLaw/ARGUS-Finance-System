/**
 * System Admin Menu
 * Creates a custom administrative menu when the Annual spreadsheet opens.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('System Admin')
    .addItem('Initialize Annual Triggers', 'setupAutomatedTriggers')
    .addToUi();
}

/**
 * Trigger Initialization
 * Programmatically provisions installable onEdit triggers for the active Annual file.
 */
function setupAutomatedTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetFunctionName = 'runAnnualFinancialSystem';
  
  // Check for existing triggers to prevent duplicates
  const existingTriggers = ScriptApp.getProjectTriggers();
  let triggerExists = false;
  
  for (let i = 0; i < existingTriggers.length; i++) {
    if (existingTriggers[i].getHandlerFunction() === targetFunctionName) {
      triggerExists = true;
      break;
    }
  }
  
  if (triggerExists) {
    ss.toast("Triggers are already active for this workbook.", "System Notice", 5);
    return;
  }
  
  // Create the installable onEdit trigger
  try {
    ScriptApp.newTrigger(targetFunctionName)
             .forSpreadsheet(ss)
             .onEdit()
             .create();
             
    ss.toast("Installable triggers successfully mapped to Annual Master.", "Setup Complete", 5);
  } catch (err) {
    SpreadsheetApp.getUi().alert("Setup failed: " + err.toString());
  }
}

/**
 * Master Trigger Orchestrator
 * Coordinates execution sequence for all underlying Annual financial automation routines.
 */
function runAnnualFinancialSystem(e) {
  handleCurrencyFormatter(e);   // Manages local sheet number formats
  handleDropdownSync(e);        // Synchronizes local sheet dropdowns
  handleGlobalDropdownSync(e);  // Distributes state updates outwards to monthly sheets
}

/**
 * Local Currency Formatter
 * Instantly formats active accounting sheets when the local toggle is changed.
 */
function handleCurrencyFormatter(e) {
  const sheet = e.source.getActiveSheet();
  const cell = e.range;
  const sheetName = sheet.getName();
  const value = cell.getValue().toString().trim();

  const config = {
    "INITIAL BALANCE SHEET": { cell: "F4", ranges: ["C5:D14"] },
    "ACCOUNTS SUMMARY": { cell: "A1", ranges: ["D6:D", "H6:H", "L6:L", "P6:P"] },
    "TRIAL BALANCE": { cell: "F4", ranges: ["C5:D36"] },
    "FINANCIAL STATEMENTS": { cell: "G9", ranges: ["D5:E29", "H5:H7", "M5:M14", "R5:R14"] },
    "CLOSING ENTRIES": { cell: "H4", ranges: ["E5:F34"] },
    "EXECUTIVE SUMMARY": { cell: "G4", ranges: ["B5:E18"] }
  };

  if (!config[sheetName]) return;

  const { cell: triggerCell, ranges } = config[sheetName];
  if (cell.getA1Notation() !== triggerCell) return;

  const format = (value === "Total in IDR") ? "Rp#,##0.00" : "$#,##0.00";

  ranges.forEach(r => sheet.getRange(r).setNumberFormat(format));
}

/**
 * Local Dropdown Synchronization
 * Mirror-syncs all localized currency dropdowns within the active workbook.
 */
function handleDropdownSync(e) {
  const editedCell = e.range;
  const sheetName = editedCell.getSheet().getName();
  const ss = e.source;
  const newValue = editedCell.getValue();

  const map = {
    "INITIAL BALANCE SHEET": "F4",
    "ACCOUNTS SUMMARY": "A1",
    "TRIAL BALANCE": "F4",
    "FINANCIAL STATEMENTS": "G9",
    "EXECUTIVE SUMMARY": "G4"
  };

  if (map[sheetName] !== editedCell.getA1Notation()) return;

  for (let name in map) {
    const targetSheet = ss.getSheetByName(name);
    if (!targetSheet) continue;

    const targetCell = targetSheet.getRange(map[name]);

    if (targetCell.getValue() !== newValue) {
      targetCell.setValue(newValue);
    }
  }
}

/**
 * Core Layout Formatting Engine
 * Formats designated data columns across all active accounting tabs.
 */
function applyCurrencyFormat(ss, newValue) {
  const format = (newValue === "Total in IDR") ? "Rp#,##0.00" : "$#,##0.00";

  const config = {
    "INITIAL BALANCE SHEET": ["C5:D14"],
    "ACCOUNTS SUMMARY": ["D6:D", "H6:H", "L6:L", "P6:P"],
    "TRIAL BALANCE": ["C5:D36"],
    "FINANCIAL STATEMENTS": ["D5:E29", "H5:H7", "M5:M14", "R5:R14"],
    "CLOSING ENTRIES": ["E5:F34"],
    "EXECUTIVE SUMMARY": ["B5:E18"]
  };

  for (let s in config) {
    const sh = ss.getSheetByName(s);
    if (!sh) continue;

    config[s].forEach(r => {
      sh.getRange(r).setNumberFormat(format);
    });
  }
}

/**
 * Network Sync: Distributed Global State Update
 * Reads monthly registry IDs from FRONT COVER and pushes currency updates to them.
 */
function handleGlobalDropdownSync(e) {
  const sheet = e.range.getSheet();
  const cell = e.range;

  const triggers = {
    "INITIAL BALANCE SHEET": "F4",
    "ACCOUNTS SUMMARY": "A1",
    "TRIAL BALANCE": "F4",
    "FINANCIAL STATEMENTS": "G9",
    "EXECUTIVE SUMMARY": "G4"
  };

  if (triggers[sheet.getName()] !== cell.getA1Notation()) return;

  const newValue = cell.getValue();
  const ssAnnual = e.source;

  // Extract sibling Monthly Spreadsheet IDs from launchpad matrix
  const frontCover = ssAnnual.getSheetByName("FRONT COVER");
  if (!frontCover) return;

  const spreadsheetIds = frontCover.getRange("B34:B45").getValues()
                                    .map(row => row[0].toString().trim())
                                    .filter(id => id !== "" && !id.toLowerCase().includes("id_"));

  const monthlyMap = {
    "INITIAL BALANCE SHEET": "F4",
    "JOURNAL ENTRIES": "M4",
    "GENERAL LEDGER": "A1",
    "TRIAL BALANCE": "F4",
    "FINANCIAL STATEMENTS": "G9",
    "CLOSING ENTRIES": "H4",
    "TRIAL BALANCE AFTER CLOSING": "F4"
  };

  spreadsheetIds.forEach(id => {
    try {
      const ssTarget = SpreadsheetApp.openById(id);

      for (let sheetName in monthlyMap) {
        const targetSheet = ssTarget.getSheetByName(sheetName);
        if (!targetSheet) continue;

        const targetCell = targetSheet.getRange(monthlyMap[sheetName]);
        if (targetCell.getValue() !== newValue) {
          targetCell.setValue(newValue);
        }
      }
      applyCurrencyFormat(ssTarget, newValue);

    } catch (err) {
      Logger.log("Error updating monthly ID from annual execution: " + id);
    }
  });

  ssAnnual.toast("Pushed currency sync to all active monthly ledgers.", "Global Sync Success");
}
