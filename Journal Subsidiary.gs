function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const editedCell = e.range;
  const row = editedCell.getRow();
  const col = editedCell.getColumn();
  const sheetName = "JOURNAL ENTRIES";

  if (sheet.getName() !== sheetName) return;
  if (row < 5) return;

  // === Constants ===
  const nonAssetKeywords = ["expenses", "income", "gain", "loss"];

  // === IDR MAP (existing one) ===
  const validationMapIDR = {
    "Cash on hand": ["IDR on Hand"],
    "Cash in bank": ["BCA J. (S)", "BCA J. (F)", "BCA J. (E)"],
    "Electronic wallets": ["GoPay", "OVO"],
    "Accounts receivable": ["A/R (Vhier)", "A/R (Edy/Yenni)"],
    "Interest receivable": ["I/R (Vhier)"],
    "Commodities": ["Gold (ANTM)"],
    "Accounts payable": ["A/P (Vhier)", "A/P (Edy/Yenni)"]
  };

  // === CAD MAP (same structure, YOU will edit subsidiaries) ===
  const validationMapCAD = {
    "Cash on hand": ["CAD on Hand"],
    "Cash in bank": ["CAD Bank"],
    "Electronic wallets": ["CAD E Wallet"],
    "Accounts receivable": ["A/R (PAR))"],
    "Interest receivable": ["I/R (PIR)"],
    "Accounts payable": ["A/P (PAP)"]
  };

  function buildNormalizedMap(map) {
    const normalized = {};
    const keys = Object.keys(map).map(k => k.toString());

    keys.sort((a, b) => b.length - a.length);

    for (let k of keys) {
      normalized[k.toLowerCase().trim()] = map[k];
    }

    return normalized;
  }

  const normalizedIDR = buildNormalizedMap(validationMapIDR);
  const normalizedCAD = buildNormalizedMap(validationMapCAD);

  function applyValidation(cell, options) {
    cell.clearContent();

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options, true)
      .setAllowInvalid(true)
      .build();

    cell.clearDataValidations();
    cell.setDataValidation(rule);
  }

  function process(input, targetCell, normalizedMap) {
    if (!input) {
      targetCell.clearContent();
      targetCell.clearDataValidations();
      return;
    }

    const inputStr = input.toString().trim();
    const lowerInput = inputStr.toLowerCase();

    targetCell.clearDataValidations();

    if (inputStr.startsWith("(")) {
      targetCell.clearContent();
      return;
    }

    if (nonAssetKeywords.some(k => lowerInput.includes(k))) {
      targetCell.setValue("Not Asset/Liability");
      return;
    }

    for (let normKey in normalizedMap) {
      if (lowerInput === normKey || lowerInput.includes(normKey)) {
        applyValidation(targetCell, normalizedMap[normKey]);
        return;
      }
    }

    targetCell.setValue("Not Recognizable");
  }

  // === CAD TABLE (E → F) ===
  if (col === 5) {
    const target = sheet.getRange(row, 6);
    process(editedCell.getValue(), target, normalizedCAD);
  }

  // === IDR TABLE (O → P) ===
  if (col === 15) {
    const target = sheet.getRange(row, 16);
    process(editedCell.getValue(), target, normalizedIDR);
  }
}