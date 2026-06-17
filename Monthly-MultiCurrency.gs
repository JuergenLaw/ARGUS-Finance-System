/**
 * System Admin Menu
 * Creates a custom administrative menu when the spreadsheet opens.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('System Admin')
    .addItem('Initialize Monthly Triggers', 'setupAutomatedTriggers')
    .addToUi();
}

/**
 * Trigger Initialization
 * Programmatically provisions installable onEdit triggers for the active sheet.
 */
function setupAutomatedTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetFunctionName = 'runFinancialSystem';
  
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
             
    ss.toast("Installable triggers successfully mapped to this workbook.", "Setup Complete", 5);
  } catch (err) {
    SpreadsheetApp.getUi().alert("Setup failed: " + err.toString());
  }
}

/**
 * Master Trigger Orchestrator
 * Coordinates execution sequence for all underlying financial automation routines.
 */
function runFinancialSystem(e) {
  handleJournalSubsidiary(e);   // Processes conditional data validation rules
  handleCurrencyFormatter(e);   // Manages local sheet number formats
  handleDropdownSync(e);        // Synchronizes local sheet dropdowns
  handleJournalCurrency(e);     // Evaluates dynamic row-by-row journal formats
  handleGlobalDropdownSync(e);  // Distributes state updates across workbook network
}

/**
 * Script 1: Journal Subsidiary Validation
 * Dynamically updates sub-account validation dropdowns based on account type.
 */
function handleJournalSubsidiary(e) {
  const sheet = e.source.getActiveSheet();
  const editedCell = e.range;
  const row = editedCell.getRow();
  const col = editedCell.getColumn();

  if (sheet.getName() !== "JOURNAL ENTRIES") return;
  if (col !== 5 || row < 5) return;

  const nonAssetKeywords = ["expenses", "income", "gain", "loss"];

  const validationMap = {
    "Cash on hand": ["CAD on Hand", "IDR on Hand"],
    "Cash in bank": ["Bank-CAD", "BCA J. (S)-IDR", "BCA J. (F)-IDR", "BCA J. (E)-IDR"],
    "Electronic wallets": ["E Wallet-CAD", "GoPay-IDR", "OVO-IDR"],
    "Accounts receivable": ["A/R (PAR)-CAD", "A/R (Vhier)-IDR", "A/R (Edy/Yenni)-IDR"],
    "Interest receivable": ["I/R (PIR)-CAD", "I/R (Vhier-IDR)"],
    "Commodities": ["Gold (ANTM)-IDR"],
    "Accounts payable": ["A/P (PAP)-CAD", "A/P (Vhier)-IDR", "A/P (Edy/Yenni)-IDR"]
  };

  const target = sheet.getRange(row, 6);
  const input = editedCell.getValue();

  if (!input) {
    target.clearContent();
    target.clearDataValidations();
    return;
  }

  const inputStr = input.toString().trim().toLowerCase();

  if (inputStr.startsWith("(")) {
    target.clearContent();
    return;
  }

  if (nonAssetKeywords.some(k => inputStr.includes(k))) {
    target.setValue("Not Asset/Liability");
    return;
  }

  const keys = Object.keys(validationMap).sort((a, b) => b.length - a.length);

  for (let key of keys) {
    if (inputStr.includes(key.toLowerCase())) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(validationMap[key], true)
        .setAllowInvalid(true)
        .build();

      target.clearContent();
      target.setDataValidation(rule);
      return;
    }
  }

  target.setValue("Not Recognizable");
}

/**
 * Script 2: Local Currency Formatter
 * Instantly formats active accounting sheets when the local toggle is changed.
 */
function handleCurrencyFormatter(e) {
  const sheet = e.source.getActiveSheet();
  const cell = e.range;
  const sheetName = sheet.getName();
  const value = cell.getValue().toString().trim();

  const config = {
    "INITIAL BALANCE SHEET": { cell: "F4", ranges: ["C5:D14"] },
    "JOURNAL ENTRIES": { cell: "M4", ranges: ["K5:L"] },
    "GENERAL LEDGER": { cell: "A1", ranges: ["E:H", "N:Q", "W:Z"] },
    "TRIAL BALANCE": { cell: "F4", ranges: ["C5:D36"] },
    "FINANCIAL STATEMENTS": { cell: "G9", ranges: ["D5:E29", "H5:H7", "M5:M14", "R5:R14"] },
    "CLOSING ENTRIES": { cell: "H4", ranges: ["E5:F34"] },
    "TRIAL BALANCE AFTER CLOSING": { cell: "F4", ranges: ["C5:D14"] }
  };

  if (!config[sheetName]) return;

  const { cell: triggerCell, ranges } = config[sheetName];
  if (cell.getA1Notation() !== triggerCell) return;

  const format = (value === "Total in IDR") ? "Rp#,##0.00" : "$#,##0.00";

  ranges.forEach(r => sheet.getRange(r).setNumberFormat(format));
}

/**
 * Core Layout Formatting Engine
 * Formats designated data columns across all active accounting tabs.
 */
function applyCurrencyFormat(ss, newValue) {
  const format = (newValue === "Total in IDR") ? "Rp#,##0.00" : "$#,##0.00";

  const config = {
    "INITIAL BALANCE SHEET": ["C5:D14"],
    "TRIAL BALANCE": ["C5:D36"],
    "FINANCIAL STATEMENTS": ["D5:E29", "H5:H7", "M5:M14", "R5:R14"],
    "CLOSING ENTRIES": ["E5:F34"],
    "JOURNAL ENTRIES": ["K5:L"],
    "GENERAL LEDGER": ["E:H", "N:Q", "W:Z"],
    "TRIAL BALANCE AFTER CLOSING": ["C5:D14"],
    "ACCOUNTS SUMMARY": ["D6:D", "H6:H", "L6:L", "P6:P"],
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
 * Script 3: Local Dropdown Synchronization
 * Mirror-syncs all localized currency dropdowns within the active workbook.
 */
function handleDropdownSync(e) {
  const editedCell = e.range;
  const sheetName = editedCell.getSheet().getName();
  const ss = e.source;
  const newValue = editedCell.getValue();

  const map = {
    "INITIAL BALANCE SHEET": "F4",
    "JOURNAL ENTRIES": "M4",
    "GENERAL LEDGER": "A1",
    "TRIAL BALANCE": "F4",
    "FINANCIAL STATEMENTS": "G9",
    "CLOSING ENTRIES": "H4",
    "TRIAL BALANCE AFTER CLOSING": "F4"
  };

  if (map[sheetName] !== editedCell.getA1Notation()) return;

  applyCurrencyFormat(ss, newValue);

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
 * Script 4: Journal Entry Currency Rule
 * Evaluates split-currency entries row-by-row on transaction ledger journals.
 */
function handleJournalCurrency(e) {
  const sheet = e.source.getActiveSheet();
  const cell = e.range;

  if (sheet.getName() !== "JOURNAL ENTRIES") return;

  const col = cell.getColumn();
  const row = cell.getRow();
  const lastRow = sheet.getLastRow();

  if (col === 6 && row >= 5) {
    let startRow = row;
    while (startRow > 5 && sheet.getRange(startRow - 1, 6).getValue() !== "") startRow--;

    let endRow = row;
    while (endRow <= lastRow && sheet.getRange(endRow, 6).getValue() !== "") endRow++;
    endRow--;

    let currency = "CAD";

    for (let r = startRow; r <= endRow; r++) {
      const val = sheet.getRange(r, 6).getValue().toString().toUpperCase();
      if (val.includes("IDR")) { currency = "IDR"; break; }
      if (val.includes("CAD")) { currency = "CAD"; break; }
    }

    const format = (currency === "IDR") ? "Rp#,##0.00" : "$#,##0.00";

    sheet.getRange(startRow, 9, endRow - startRow + 1, 2)
         .setNumberFormat(format);
  }

  if (cell.getA1Notation() === "M4") {
    const display = sheet.getRange("M4").getValue();
    const format = (display === "Total in IDR") ? "Rp#,##0.00" : "$#,##0.00";

    sheet.getRange(5, 11, lastRow - 4, 2)
         .setNumberFormat(format);
  }
}

/**
 * Network Sync: Distributed Global State Update
 * Distributes structural currency changes across Annual and Monthly files.
 */
function handleGlobalDropdownSync(e) {
  const sheet = e.range.getSheet();
  const cell = e.range;

  const triggers = {
    "INITIAL BALANCE SHEET": "F4",
    "JOURNAL ENTRIES": "M4",
    "GENERAL LEDGER": "A1",
    "TRIAL BALANCE": "F4",
    "FINANCIAL STATEMENTS": "G9",
    "CLOSING ENTRIES": "H4",
    "TRIAL BALANCE AFTER CLOSING": "F4"
  };

  if (triggers[sheet.getName()] !== cell.getA1Notation()) return;

  const newValue = cell.getValue();
  const ssLocal = e.source;
  const currentId = ssLocal.getId();

  const frontCover = ssLocal.getSheetByName("FRONT COVER");
  if (!frontCover) return; 

  const annualId = frontCover.getRange("B32").getValue().toString().trim();
  if (!annualId || annualId === "" || annualId.toLowerCase() === "id_annual") {
    ssLocal.toast("Global sync skipped: Annual ID missing in FRONT COVER!B32.", "Notice");
    return;
  }

  const annualMap = {
    "INITIAL BALANCE SHEET": "F4",
    "ACCOUNTS SUMMARY": "A1",
    "TRIAL BALANCE": "F4",
    "FINANCIAL STATEMENTS": "G9",
    "EXECUTIVE SUMMARY": "G4"
  };

  const monthlyMap = {
    "INITIAL BALANCE SHEET": "F4",
    "JOURNAL ENTRIES": "M4",
    "GENERAL LEDGER": "A1",
    "TRIAL BALANCE": "F4",
    "FINANCIAL STATEMENTS": "G9",
    "CLOSING ENTRIES": "H4",
    "TRIAL BALANCE AFTER CLOSING": "F4"
  };

  try {
    const ssAnnual = SpreadsheetApp.openById(annualId);

    // Update Master Annual Summary sheets
    for (let sheetName in annualMap) {
      const targetSheet = ssAnnual.getSheetByName(sheetName);
      if (!targetSheet) continue;

      const targetCell = targetSheet.getRange(annualMap[sheetName]);
      if (targetCell.getValue() !== newValue) {
        targetCell.setValue(newValue);
      }
    }
    applyCurrencyFormat(ssAnnual, newValue);

    // Process Sibling Monthly Workbooks from Launchpad Matrix
    const annualFrontCover = ssAnnual.getSheetByName("FRONT COVER");
    if (annualFrontCover) {
      const spreadsheetIds = annualFrontCover.getRange("B34:B45").getValues()
                                        .map(row => row[0].toString().trim())
                                        .filter(id => id !== "" && id !== currentId && !id.toLowerCase().includes("id_"));

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
          
        } catch (siblingErr) {
          Logger.log("Skipped or failed updating sibling ID: " + id);
        }
      });
    }

    ssLocal.toast("Successfully synchronized currency states across the entire annual network!", "Network Sync Complete");

  } catch (err) {
    Logger.log("Error during distribution sync: " + err);
    ssLocal.toast("Network sync failed. Check Annual ID in FRONT COVER!B32.", "Execution Error");
  }
}
