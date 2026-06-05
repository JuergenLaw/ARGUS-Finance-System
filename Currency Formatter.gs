function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const cell = e.range;
  const sheetName = sheet.getName();
  const value = cell.getValue().toString().trim();

  // =========================
  // CONFIG (all sheets here)
  // =========================
  const config = {
    "INITIAL BALANCE SHEET": {
      cell: "F4",
      ranges: ["C5:D14"]
    },
    "JOURNAL ENTRIES": {
      cell: "J1",
      ranges: ["H5:I"]
    },
    "GENERAL LEDGER": {
      cell: "A1",
      ranges: ["E:H", "N:Q", "W:Z"]
    },
    "TRIAL BALANCE": {
      cell: "F4",
      ranges: ["C5:D36"]
    },
    "FINANCIAL STATEMENT": {
      cell: "G9",
      ranges: ["D5:E29", "H5:H7", "M5:M14", "R5:R14"]
    },
    "CLOSING ENTRIES": {
      cell: "H4",
      ranges: ["E5:F34"]
    },
    "TRIAL BALANCE AFTER CLOSING": {
      cell: "F4",
      ranges: ["C5:D14"]
    }
  };

  // =========================
  // CHECK IF SHEET IS IN CONFIG
  // =========================
  if (!config[sheetName]) return;

  const { cell: triggerCell, ranges } = config[sheetName];

  // Only trigger if correct dropdown edited
  if (cell.getA1Notation() !== triggerCell) return;

  // =========================
  // FORMAT DECISION
  // =========================
  let format = "$#,##0.00"; // default CAD

  if (value === "Total in IDR") {
    format = "Rp#,##0.00";
  }

  // =========================
  // APPLY FORMATTING
  // =========================
  ranges.forEach(r => {
    sheet.getRange(r).setNumberFormat(format);
  });
}