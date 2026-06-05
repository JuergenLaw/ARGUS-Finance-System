# Personal Financial Ledger Automation

An interconnected accounting network that links 12 individual monthly sheets to a central annual master dashboard using Google Apps Script. The system handles bi-directional currency switching, automated ledger formatting, and multi-workbook synchronization entirely in the background.

---

## Features

* **Bi-Directional Currency Sync:** Flip the currency dropdown on any sheet (monthly or annual), and the script instantly broadcasts the state change across the entire network.
* **Dynamic Formatting Engine:** Automatically re-formats accounting tables between Indonesian Rupiah (`Rp#,##0.00`) and Canadian Dollars (`$#,##0.00`) based on the active selection.
* **Decoupled Architecture:** Zero hardcoded spreadsheet IDs. The system reads connections dynamically from a launchpad on the `FRONT COVER` tab, making year-over-year rollouts or duplication easy.
* **Programmatic Trigger Setup:** Includes a custom `System Admin` menu to provision background installable triggers automatically, removing the need to configure them manually in the Google Cloud console.

---

## Project Structure

The underlying scripts rely on explicit tab names to process formatting changes safely without throwing null errors.

### Monthly Sheets
* `FRONT COVER` — Contains system configurations and network links.
* `JOURNAL ENTRIES` & `GENERAL LEDGER` — Transactional input ledgers.
* `TRIAL BALANCE` & `FINANCIAL STATEMENTS` — Core financial reporting.
* `CLOSING ENTRIES` & `TRIAL BALANCE AFTER CLOSING` — Month-end close tabs.

### Annual Master
* `FRONT COVER` — The master configuration hub holding the spreadsheet IDs for Jan–Dec (rows `B34:B45`).
* `ACCOUNTS SUMMARY` & `EXECUTIVE SUMMARY` — Multi-month data aggregations.

---

## How to Set It Up

1. **Link the Network:** Open your Annual Master sheet, go to the `FRONT COVER`, and paste your monthly Google Spreadsheet IDs into column B (rows 34–45).
2. **Apply the Script:** Paste the monthly codebase into your monthly sheets and the annual codebase into your master sheet.
3. **Initialize Triggers:** * Refresh your spreadsheet browser tab and wait a few seconds for the UI to register.
   * Click the custom menu: **System Admin** ➡️ **Initialize Monthly Triggers**.
   * Authorize the standard Google security permissions on the initial run.
4. **Test the Sync:** Once the success toast notification appears, the automation is live. Change a currency dropdown on any statement tab to see the network update.

---

## Design Constraints & Rules

* **Tab Names are Strict:** The formatting loops search for exact tab strings (like `"JOURNAL ENTRIES"`). If you rename a tab, the script will safely bypass it to avoid breaking the workbook.
* **Copying Files Drops Triggers:** Google Sheets strips out installable background triggers when you duplicate a file. When creating a new month from your template, remember to re-run step 3 (**System Admin ➡️ Initialize Monthly Triggers**).
* **Top-Left Anchor Validation:** The sheet reads year parameters from the merged `B31:C31` cell on the `FRONT COVER`. Ensure your data values remain anchored in the top-left cell (`B31`) to protect data integrity.

---

## Templates

If you want to look at the layout or test the scripts, feel free to check out these viewing samples and copy them:
* [Monthly Ledger Template ↗](PASTE_YOUR_SANITIZED_MONTHLY_LINK_HERE)
* [Annual Master Dashboard ↗](PASTE_YOUR_SANITIZED_ANNUAL_LINK_HERE)
