function onEdit(e) {
  const editedCell = e.range;
  const sheetName = editedCell.getSheet().getName();
  const ss = e.source;

  const newValue = editedCell.getValue();

  // === Mapping: sheet → its control cell ===
  const map = {
    "INITIAL BALANCE SHEET": "F4",
    "JOURNAL ENTRIES": "M4",
    "GENERAL LEDGER": "A1",
    "TRIAL BALANCE": "F4",
    "FINANCIAL STATEMENTS": "G9",
    "CLOSING ENTRIES": "H4",
    "TRIAL BALANCE AFTER CLOSING": "F4"
  };

  // === Check if edited cell is one of the mapped control cells ===
  if (map[sheetName] !== editedCell.getA1Notation()) return;

  // === Sync to all other sheets ===
  for (let name in map) {
    const targetSheet = ss.getSheetByName(name);
    if (!targetSheet) continue;

    const targetCell = targetSheet.getRange(map[name]);

    // Prevent unnecessary overwrite (important)
    if (targetCell.getValue() !== newValue) {
      targetCell.setValue(newValue);
    }
  }
}