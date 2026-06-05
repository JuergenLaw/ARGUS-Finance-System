# Dynamic Multi-Spreadsheet Financial Ledger System

An automated, interconnected accounting and financial ledger network built using **Google Sheets** and **Google Apps Script**. This ecosystem connects 12 individual Monthly Spreadsheets to a central Annual Master Spreadsheet, featuring real-time, bi-directional synchronization, dynamic currency formatting, and automated background triggers.

---

## Key Features

* **Bi-Directional Currency Sync:** Changing the global currency dropdown in *any* sheet (Monthly or Annual) automatically signals and updates the entire network.
* **Intelligent Formatting Engine:** Dynamically transforms accounting ranges between Indonesian Rupiah (`Rp#,##0.00`) and Canadian Dollars (`$#,##0.00`) depending on the active state.
* **Zero Hardcoded IDs:** Fully decoupled architecture. All spreadsheet connections are read dynamically from a `FRONT COVER` launchpad, allowing effortless system duplication or year-over-year rollover.
* **One-Click System Initialization:** Custom `⚙️ System Admin` native UI menu allows users to programmatically install complex background `onEdit` triggers without touching code.

---

## System Architecture & File Layout

The system relies on exact tab structural naming to safely execute formatting passes without throwing errors.

### Monthly Spreadsheet Tabs
* `FRONT COVER` — Holds configuration metadata and system links.
* `JOURNAL ENTRIES` & `GENERAL LEDGER` — Raw transactional input data.
* `TRIAL BALANCE` & `FINANCIAL STATEMENTS` — Core financial reporting sheets.
* `CLOSING ENTRIES` & `TRIAL BALANCE AFTER CLOSING` — Period-end processing.

### Annual Master Spreadsheet Tabs
* `FRONT COVER` — The Master Launchpad containing the Spreadsheet IDs for rows `B34:B45` (January–December).
* `ACCOUNTS SUMMARY` & `EXECUTIVE SUMMARY` — High-level multi-month aggregations.

---

## Setup & Installation Instructions

To deploy this system across your workbook network:

1. **Populate the Launchpad:** In the Annual Master Spreadsheet, navigate to the `FRONT COVER` and input the respective Google Spreadsheet IDs for each month in column B (rows 34–45).
2. **Open Apps Script:** Paste the monthly codebase into your Monthly template and the annual codebase into your Annual Master sheet.
3. **Initialize Background Triggers:**
   * Refresh your spreadsheet browser tab.
   * Wait 3–5 seconds for the custom native menu to load.
   * Click **`⚙️ System Admin`** ➡️ **`🚀 Initialize Triggers`**.
   * Grant the standard Google security permissions on the initial run.
4. **Success:** Once the toast notification appears in the bottom right corner, your background engine is officially active.

---

## Limitations & Governance Rules

To ensure the ecosystem doesn't drop network requests, adhere to the following governance rules:

* **Tab Renaming Prohibited:** The underlying script relies on exact string matches (e.g., `"JOURNAL ENTRIES"`). Renaming any core tracking sheet will cause the formatting engine to safely bypass that tab to prevent workbook corruption.
* **File Duplication Trigger Loss:** Google Sheets intentionally strips background installable triggers during a file copy operation for account security. When generating a new month from your template, **you must re-run step 3 (Initialize Triggers)** in the new file.
* **Cell Merge Preservation:** The year formatting function reads directly from cell `B31` on the `FRONT COVER`. While `B31` is merged with `C31`, data integrity must remain anchored in the top-left cell (`B31`).

---

## 📜 License

**Personal / Proprietary License** This codebase and architectural design are built exclusively for private financial tracking and proprietary accounting usage. All rights reserved.
