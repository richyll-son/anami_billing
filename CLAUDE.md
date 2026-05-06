# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Google Apps Script billing system for **AHNHAI (Anami Homes North Homeowners Association)**, Consolacion Cebu. Manages monthly water billing and association dues for ~300 units across 2 phases. Runs inside a target Google Spreadsheet via Apps Script.

## Development Workflow

Uses [Clasp](https://github.com/google/clasp) to sync between local `.gs` files and the Apps Script editor.

```bash
npm install -g @google/clasp   # one-time
clasp login                    # one-time
clasp push                     # deploy local files → Apps Script
clasp pull                     # pull latest from editor
clasp open                     # open editor in browser
clasp logs                     # stream Stackdriver logs
```

No build step. Files go directly to Google's V8 runtime.

## File Map

| File | Responsibility |
|---|---|
| `Code.gs` | Constants (SH, WL, DL, column maps), `onOpen` menu, `onEdit` trigger, shared utilities |
| `Setup.gs` | `setupSystem()`, create all 11 sheets, `importMasterlistFromSource()`, `populateWaterInputTable()` |
| `BillingEngine.gs` | `processAndGenerateBills()`, rate computation, penalty check, `recalcWaterBalances()`, `recalcDuesBalances()` |
| `PaymentProcessor.gs` | `postPayments()`, oldest-first credit application |
| `LedgerManager.gs` | `refreshUnitLedger()` display, `handleUnitLedgerEdit()` write-back |
| `SummaryGenerator.gs` | `refreshMonthlySummary()` |
| `PrintGenerator.gs` | `regenerateBillPrint()` for Phase 1 & 2 (batch-built, 38 rows/bill) |
| `WebApp.gs` | `doGet()` entry point; all `webapp_*` server-side functions (no UI dialogs, return `{ok, msg, ...}`) |
| `Index.html` | Bootstrap 5.3 SPA — 6 tabs: Dashboard, Readings, Payments, Ledger, Summary, Setup |

## Sheet Architecture

**9 visible sheets** (in order):
1. `Masterlist` — editable unit registry; synced from source on demand
2. `Water Reading Input` — monthly form (rows 1–13) + unit reading table (row 14+)
3. `Water Reading Data Store` — append-only historical readings
4. `Rate Calculator` — one row per billing cycle
5. `Unit Ledger` — display view with B1 dropdown; refreshes on selection
6. `Central Payment Log` — admin enters payments here; posted via menu
7. `Monthly Summary` — rebuilt each billing run
8. `Phase 1 Bill Print` / `Phase 2 Bill Print` — all bills stacked vertically

**2 hidden data stores** (prefixed `_`):
- `_WaterLedger` — 19 cols per row per unit (see `WL.*` constants in Code.gs)
- `_DuesLedger` — 9 cols per row per unit (see `DL.*` constants in Code.gs)

## Key Business Rules

- **Rate** = (MCWD + Electricity + Manpower) ÷ Total Consumption of all units
- **Water Bill** = Consumption × Rate, minimum ₱250
- **Penalty** = 5% × previous balance, only if previous due date has passed
- **Bill Date** = 10th of month after reading month
- **Due Date** = 2nd of month, 2 months after reading month
- **Bill Number** = `[YEAR][MONTH2D][PHASE][BLOCK2D][LOT2D]00`
- **Association Dues** = ₱500/unit/month flat (combined lots count as one unit)
- **Add-on MCWD** = per-unit carry-forward amount stored in `_WaterLedger` ADDON column

## Balance Calculation

Running balance per unit — recalculated in one batch write after any change:

```
BALANCE[i] = BALANCE[i-1] + DEBIT[i] + PENALTY[i] + ADDON[i] - CREDIT[i]  (min 0)
```

In the bill:
- **Arrears** = BALANCE − current DEBIT − current PENALTY − current ADDON
- **TOTAL WATER DUE** = BALANCE (the full running balance)

## Source Spreadsheets (read-only)

- Masterlist: `1b-sOFs61PLmv8JcCtFftbxwE6zS8f5yA96nPahFFEyk` (gid 586990653)
- Historical water data: `1FiFcIazPqHCUAvCp57mzrsxAw8ZlMiEK`

Never write to these. `SRC_MASTERLIST_ID` / `SRC_WATER_STORE_ID` in `Code.gs`.

## Adding to Apps Script

Copy each `.gs` file into the Apps Script editor as a separate script file (File → New script file), and add `Index.html` as an HTML file. File order does not matter — all files share one global scope.

After adding all files, run **AHNHAI Billing System → Initial Setup** once. Then each month: fill Water Reading Input → Process & Generate Bills.

## Web App Deployment

The system can be served as a free GAS web app via `doGet()` in `WebApp.gs`.

1. In the Apps Script editor: **Deploy → New deployment → Web app**
2. Execute as: **Me** | Who has access: **Anyone** (or restrict as needed)
3. Copy the deployment URL — this is the app's permanent URL

All UI-less server functions are in `WebApp.gs` prefixed `webapp_*`. They never call `SpreadsheetApp.getUi()` / `alert_()`, and always return `{ok: bool, msg: string, ...}` so the browser can handle errors gracefully.

The `<?= SRC_MASTERLIST_ID ?>` scriptlet in `Index.html` (Setup tab) injects the source file ID at serve time — this is a GAS template evaluation, not a JavaScript variable.
