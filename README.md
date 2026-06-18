# ARGUS: Accounting Report & GUidance System

An automated double-entry accounting system built entirely within Google Sheets and Google Apps Script. ARGUS automates a simplified version of the accounting cycle, from daily journal entries to ledger posting, trial balancing, and financial statements.

This system is divided into 2 separate sub-systems: *Monthly Financial Reports (MFRs)* for recording transactions, and *Annual Financial Summaries (AFSs)* for combining individual MFRs into a centralized yearly evaluation.

---

## Single-Currency Setup Features

* **Automated Bookkeeping Cycle**: Records transactions through `JOURNAL ENTRIES` and have them automatically configured from `GENERAL LEDGER` and `SUBSIDIARY LEDGER` all the way to `TRIAL BALANCE AFTER CLOSING`, ensuring every transaction remains balanced and traceable.
* **Simple Error-Handling**: Built-in safeguards (specifically in `JOURNAL ENTRIES`) that stop you from inputting incorrect or unrecognized accounts and subsidiaries, making sure the rest of the automation works well.
* **Automatic Carry-Forward Pipeline**: Allows the importing of the previous month's closing balances via simple URL linkages so it moves continuously without manual data-entry errors.
* **AFS Integration**: Combines data from 12 monthly MFRs into a centralized annual summary system so financial data can be analysed through automated summaries, trend visualizations, and comparative tables.

---

## Multi-Currency Setup Features (Active WIP Branch)

* **Bi-Directional Currency Sync:** Flip the currency dropdown on any sheet (monthly or annual), and the script instantly broadcasts the state change across the entire network.
<img width="481" height="413" alt="Currency-Sync" src="https://github.com/user-attachments/assets/94f4cfcd-8f4f-4605-b50e-457b20a6c3c8" />

* **Dynamic Formatting Engine:** Automatically re-formats accounting tables between Indonesian Rupiah (`Rp#,##0.00`) and Canadian Dollars (`$#,##0.00`) based on the active selection.
* **Decoupled Architecture:** Zero hardcoded spreadsheet IDs. The system reads connections dynamically from a launchpad on the `FRONT COVER` tab, making year-over-year rollouts or duplication easy.
* **Programmatic Trigger Setup:** Includes a custom `System Admin` menu to provision background installable triggers automatically, removing the need to configure them manually in the Google Cloud console.

---

## Spreadsheet Structure

### Monthly Financial Reports (MFRs)
* `FRONT COVER`: Contains system configurations and network links.
* `CHART OF ACCOUNTS`: Details the accounts and subsidiaries used across the system.
* `INITIAL BALANCE SHEET`: Pulls data from last month's spreadsheet.
* `JOURNAL ENTRIES`: Records financial transactions, *the only tab a user needs to edit*.
* `GENERAL LEDGER` & `SUBSIDIARY LEDGER`: Categorisation of accounts and subsidiaries used in the `JOURNAL ENTRIES`.
* `TRIAL BALANCE`: Balances the debit and credit, *if both are balanced, then no errors are made by the system so far*.
* `FINANCIAL STATEMENTS`: Checks net profit/loss and makes a new capital balance.
* `CLOSING ENTRIES`: Closes expenses, liabilities, and prive.
* `TRIAL BALANCE AFTER CLOSING`: Balances the debit and credit after closing, *if both are balanced, then no errors are made by the system so far*.

### Annual Financial Summaries (AFSs)
* `FRONT COVER`: Contains system configurations and network links.
* `CHART OF ACCOUNTS`: Details the accounts and subsidiaries used across the system.
* `INITIAL BALANCE SHEET`: Pulls data from January's spreadsheet.
* `ACCOUNTS SUMMARY` & `SUBSIDIARY SUMMARY`: Multi-month data summaries.
* `TRIAL BALANCE`: Balances the debit and credit, *if both are balanced, then no errors are made by the system so far*.
* `FINANCIAL STATEMENTS`: Checks net profit/loss and makes a new capital balance.
* `EXECUTIVE SUMMARY`: Evaluates all financial happenings throughout the year.

---

## Custom Builds Service

While ARGUS is 100% free and open-source for anyone who wants to download, study, or manually customize it, I understand that businesses and busy individuals often need tailored systems.

If you want a custom-tailored financial tracking setup without getting your hands dirty in Google Apps Script, you can hire me to build it for you.

I offer custom services including:
* Designing custom `CHART OF ACCOUNTS` tailored to your business or personal assets.
* Re-wiring the scripts to support custom local currencies and local tax codes.
* Integrating custom financial reporting dashboards, visual KPI tracking, and analytical charts.

To discuss a custom setup, feel free to open a discussion here on GitHub or reach out to me directly via:
* E-mail: juergenllaturisa@gmail.com
* LinkedIn: Juergen Law

---

## How to Set It Up

[ARGUS Guide Series ↗]

---

## Templates

[Monthly Financial Report Template-SingleCurrency ↗]
[Annual Financial Summary Template-SingleCurrency ↗]

[Monthly Financial Report Template-MultiCurrency (WIP) ↗]
[Annual Financial Summary Template-MultiCurrency (WIP) ↗]

*If you encounter any bugs or errors, please contact me so I can fix it for you!*
