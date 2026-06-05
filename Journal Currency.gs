function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const cell = e.range;

  if (sheet.getName() !== "JOURNAL ENTRIES") return;

  const col = cell.getColumn();
  const row = cell.getRow();
  const lastRow = sheet.getLastRow();

  // =========================
  // 🟢 PART 1: I:J formatting (based on transaction currency)
  // =========================
  if (col === 6 && row >= 5) {

    // === Find start of transaction ===
    let startRow = row;
    while (startRow > 5 && sheet.getRange(startRow - 1, 6).getValue() !== "") {
      startRow--;
    }

    // === Find end of transaction ===
    let endRow = row;
    while (endRow <= lastRow && sheet.getRange(endRow, 6).getValue() !== "") {
      endRow++;
    }
    endRow--;

    // === Detect currency inside block ===
    let currency = "";

    for (let r = startRow; r <= endRow; r++) {
      const val = sheet.getRange(r, 6).getValue().toString().trim().toUpperCase();

      if (val.includes("IDR")) {
        currency = "IDR";
        break;
      }

      if (val.includes("CAD")) {
        currency = "CAD";
        break;
      }
    }

    // Default to CAD
    if (currency === "") currency = "CAD";

    const formatIJ = (currency === "IDR")
      ? "Rp#,##0.00"
      : "$#,##0.00";

    // Apply to I:J (columns 9–10)
    sheet.getRange(startRow, 9, endRow - startRow + 1, 2)
         .setNumberFormat(formatIJ);
  }

  // =========================
  // 🔵 PART 2: K:L formatting (based on M4 toggle)
  // =========================
  if (cell.getA1Notation() === "M4") {
    const display = sheet.getRange("M4").getValue().toString().trim();

    const formatKL = (display === "Total in IDR")
      ? "Rp#,##0.00"
      : "$#,##0.00";

    // Apply to entire K:L columns (11–12)
    sheet.getRange(5, 11, lastRow - 4, 2)
         .setNumberFormat(formatKL);
  }
}