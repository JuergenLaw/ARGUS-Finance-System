function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const editedCell = e.range;
  const row = editedCell.getRow();
  const col = editedCell.getColumn();
  const sheetName = "JOURNAL ENTRIES";

  if (sheet.getName() !== sheetName) return;

  // === Constants ===
  const nonAssetKeywords = ["expenses", "income", "gain", "loss"];
  const validationMap = {
    "Cash on hand": ["IDR on Hand"],
    "Cash in bank": ["BCA J. (S)", "BCA J. (F)", "BCA J. (E)"],
    "Electronic wallets": ["GoPay", "OVO", "Flazz"],
    "Accounts receivable": ["A/R (X)", "A/R (Y)", "A/R (Z)"],
    "Interest receivable": ["I/R (X)"],
    "Commodities": ["Gold (ANTM)"],
    "Accounts payable": ["A/P (X)", "A/P (Y)", "A/P (Z)"]
  };

  // Prepare normalized key map for case-insensitive matching
  const normalizedMap = {};
  const keys = Object.keys(validationMap).map(k => k.toString());
  // Sort by longest key first so "cash in bank" matches before shorter keys
  keys.sort((a, b) => b.length - a.length);
  for (let k of keys) {
    normalizedMap[k.toLowerCase().trim()] = validationMap[k];
  }

  /**
   * Apply validation to `cell` using `options`.
   * NOTE: setAllowInvalid(true) so pre-existing cell values won't trigger
   * "violates data validation" immediately after the rule is applied.
   */
  function applyValidation(cell, options) {
  // Always clear the old value first when applying a new rule
  cell.clearContent();

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(true)
    .build();

  cell.clearDataValidations();
  cell.setDataValidation(rule);
}

  // === Column E → Column F (validation logic) ===
  if (col === 5 && row >= 5) {
    const input = editedCell.getValue();
    const target = sheet.getRange(row, 6);

    if (!input) {
      target.clearContent();
      target.clearDataValidations();
      return;
    }

    const inputStr = input.toString().trim();
    const lowerInput = inputStr.toLowerCase();
    target.clearDataValidations();

    if (inputStr.startsWith("(")) {
      target.clearContent();
      return;
    }

    // If non-asset keyword found
    if (nonAssetKeywords.some(k => lowerInput.includes(k.toLowerCase()))) {
      target.setValue("Not Asset/Liability");
      return;
    }

    // Try to match a defined key (case-insensitive)
    for (let normKey in normalizedMap) {
      if (lowerInput === normKey || lowerInput.includes(normKey)) {
        applyValidation(target, normalizedMap[normKey]);
        return;
      }
    }

    // Fallback
    target.setValue("Not Recognizable");
    return;
  }

  // === Column O → Column P (simple classification) ===
  if (col === 15) {
    const input = editedCell.getValue();
    const target = sheet.getRange(row, 16);

    if (!input) {
      target.clearContent();
      return;
    }

    const inputStr = input.toString().trim();
    const lowerInput = inputStr.toLowerCase();

    if (inputStr.startsWith("(")) {
      target.clearContent();
      return;
    }

    if (nonAssetKeywords.some(k => lowerInput.includes(k.toLowerCase()))) {
      target.setValue("Not Asset/Liability");
      return;
    }

    for (let normKey in normalizedMap) {
      if (lowerInput === normKey || lowerInput.includes(normKey)) {
        const originalKey = keys.find(k => k.toLowerCase().trim() === normKey);
        target.setValue(originalKey || normKey);
        return;
      }
    }

    target.setValue("Not Recognizable");
  }
}
