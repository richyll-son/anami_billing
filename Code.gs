// ============================================================
// AHNHAI Billing System — Code.gs
// Constants · onOpen menu · onEdit trigger · shared utilities
// ============================================================

// ── Source spreadsheets (read-only) ─────────────────────────
var SRC_MASTERLIST_ID  = '1b-sOFs61PLmv8JcCtFftbxwE6zS8f5yA96nPahFFEyk';
var SRC_WATER_STORE_ID = '1FiFcIazPqHCUAvCp57mzrsxAw8ZlMiEK';

// ── Sheet name registry ──────────────────────────────────────
var SH = {
  MASTERLIST : 'Masterlist',
  W_INPUT    : 'Water Reading Input',
  W_STORE    : 'Water Reading Data Store',
  RATE_CALC  : 'Rate Calculator',
  UNIT_LEDGER: 'Unit Ledger',
  PAY_LOG    : 'Central Payment Log',
  SUMMARY    : 'Monthly Summary',
  P1_PRINT   : 'Phase 1 Bill Print',
  P2_PRINT   : 'Phase 2 Bill Print',
  _WL        : '_WaterLedger',   // hidden data store
  _DL        : '_DuesLedger',    // hidden data store
};

// ── Business constants ───────────────────────────────────────
var MONTHS         = ['January','February','March','April','May','June',
                      'July','August','September','October','November','December'];
var ASSOC_DUES     = 500;
var MANPOWER_DEF   = 35000;
var MIN_WATER_BILL = 250;
var PENALTY_RATE   = 0.05;

// ── Layout constants ─────────────────────────────────────────
// Water Reading Input: rows 1-13 are form fields; data table starts at row 14
var INPUT_TABLE_START = 14;
// Water Input table columns (1-based)
var WI_COL = { UNIT:1, METER:2, OWNER:3, CUR:4, PREV:5, CONS:6 };
// Unit Ledger layout
var UL_DATA_ROW   = 6;   // first data row in Unit Ledger display
var UL_WATER_COL  = 1;   // water ledger starts at column 1
var UL_DUES_COL   = 17;  // dues ledger starts at column 17

// _WaterLedger column indices (0-based)
var WL = {
  UNIT:0, YEAR:1, MONTH:2, BILL_DATE:3, PREV_DATE:4, PRESENT_DATE:5,
  PREV_RDG:6, CUR_RDG:7, RATE:8, DUE_DATE:9, PENALTY:10,
  DEBIT:11, CREDIT:12, BALANCE:13, ADDON:14,
  OR:15, REMARKS:16, BILL_NO:17, PAY_DATE:18
};
var WL_COLS = 19;

// _DuesLedger column indices (0-based)
var DL = {
  UNIT:0, YEAR:1, MONTH:2, PAY_DATE:3,
  DEBIT:4, CREDIT:5, BALANCE:6, OR:7, REMARKS:8
};
var DL_COLS = 9;

// Unit Ledger display → _WaterLedger column map (display col 1-based → WL index)
// -1 = read-only / calculated
var UL_WATER_MAP = [
  null,   // 0  (unused, 1-based)
  18,     // 1  Payment Date    → WL.PAY_DATE
  3,      // 2  Bill Date       → WL.BILL_DATE
  4,      // 3  Prev Date       → WL.PREV_DATE
  5,      // 4  Present Date    → WL.PRESENT_DATE
  6,      // 5  Prev Reading    → WL.PREV_RDG
  7,      // 6  Present Reading → WL.CUR_RDG
  8,      // 7  Rate/Cubic      → WL.RATE
  9,      // 8  Due Date        → WL.DUE_DATE
  10,     // 9  Penalty         → WL.PENALTY
  11,     // 10 Debit           → WL.DEBIT   (triggers recalc)
  12,     // 11 Credit          → WL.CREDIT  (triggers recalc)
  -1,     // 12 Balance         → read-only
  14,     // 13 Add-On MCWD     → WL.ADDON
  15,     // 14 OR Number       → WL.OR
  16      // 15 Remarks         → WL.REMARKS
];

// Unit Ledger display → _DuesLedger column map (display col relative to UL_DUES_COL)
var UL_DUES_MAP = [
  null,  // 0 (unused)
  3,     // 1 (+16=col17) Payment Date → DL.PAY_DATE
  2,     // 2 (+16=col18) Month        → DL.MONTH  (shouldn't edit but allow)
  4,     // 3 (+16=col19) Debit        → DL.DEBIT  (triggers recalc)
  5,     // 4 (+16=col20) Credit       → DL.CREDIT (triggers recalc)
  -1,    // 5 (+16=col21) Balance      → read-only
  7,     // 6 (+16=col22) OR Number    → DL.OR
  8      // 7 (+16=col23) Remarks      → DL.REMARKS
];

// ── Menu ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AHNHAI Billing System')
    .addItem('▶ Initial Setup (first run only)',  'setupSystem')
    .addSeparator()
    .addItem('⚡ Process & Generate Bills',       'processAndGenerateBills')
    .addItem('💰 Post Payments',                   'postPayments')
    .addSeparator()
    .addItem('↺  Refresh Summary',               'refreshMonthlySummary')
    .addItem('↻  Import Masterlist',              'importMasterlistFromSource')
    .addItem('↻  Refresh Reading Input Table',   'populateWaterInputTable')
    .addToUi();
}

// ── onEdit trigger ───────────────────────────────────────────
function onEdit(e) {
  try {
    var sh = e.range.getSheet();
    if (sh.getName() === SH.UNIT_LEDGER) handleUnitLedgerEdit(e);
  } catch (err) {
    Logger.log('onEdit: ' + err.message);
  }
}

// ── Spreadsheet accessors ────────────────────────────────────
function ss_()       { return SpreadsheetApp.getActiveSpreadsheet(); }
function getSheet_(n){ return ss_().getSheetByName(n); }
function orCreate_(n){
  var s = ss_().getSheetByName(n);
  return s || ss_().insertSheet(n);
}

// ── Number / string formatting ───────────────────────────────
function fmt2(n) {
  var v = parseFloat(n);
  return isNaN(v) ? '0.00' : v.toFixed(2);
}

function fmtDate(d) {
  if (!d) return '';
  if (typeof d === 'string') return d;
  var mm  = String(d.getMonth() + 1).padStart(2, '0');
  var dd  = String(d.getDate()).padStart(2, '0');
  var yy  = d.getFullYear();
  return mm + '/' + dd + '/' + yy;
}

function parseDate(s) {
  if (!s) return null;
  if (s instanceof Date) return new Date(s.getTime());
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function toNum(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }

// ── Unit ID helpers ──────────────────────────────────────────
function buildUID(phase, block, lot) {
  return 'P' + phase + 'B' + block + 'L' + lot;
}

function parseUID(uid) {
  var m = String(uid).match(/^P(\d+)B(\d+)L(.+)$/);
  return m ? { phase: +m[1], block: +m[2], lot: m[3] } : null;
}

// ── Billing date helpers ─────────────────────────────────────
function getBillDate(yr, mo) {          // 10th of month after reading month
  var m = mo + 1, y = yr;
  if (m > 12) { m = 1; y++; }
  return new Date(y, m - 1, 10);
}

function getDueDate(yr, mo) {           // 2nd of month, 2 months after reading
  var m = mo + 2, y = yr;
  if (m > 12) { m -= 12; y++; }
  return new Date(y, m - 1, 2);
}

function getMonthNum(name) { return MONTHS.indexOf(name) + 1; }
function getMonthName(num) { return MONTHS[num - 1] || ''; }

// ── Bill number ──────────────────────────────────────────────
function buildBillNum(yr, mo, phase, block, lot) {
  var l = String(lot).split('&')[0];
  return String(yr) +
    String(mo).padStart(2, '0') +
    String(phase) +
    String(block).padStart(2, '0') +
    String(l).padStart(2, '0') +
    '00';
}

// ── Masterlist helpers ───────────────────────────────────────
function getMLRow(mlData, unitId) {
  for (var i = 1; i < mlData.length; i++) {
    if (mlData[i][0] === unitId) return mlData[i];
  }
  return null;
}

function ownerName(mlRow) {
  if (!mlRow) return '';
  return (String(mlRow[4]) + ', ' + String(mlRow[5])).replace(/(^,\s*|,\s*$)/g, '').trim();
}

// ── UI helpers ───────────────────────────────────────────────
function alert_(msg)        { SpreadsheetApp.getUi().alert(msg); }
function toast_(msg, title) { ss_().toast(msg, title || 'AHNHAI Billing', 4); }
function confirm_(title, msg) {
  var r = SpreadsheetApp.getUi().alert(title, msg,
    SpreadsheetApp.getUi().ButtonSet.YES_NO);
  return r === SpreadsheetApp.getUi().Button.YES;
}

function convertOldDataToNewDataStore() {
  try {
    /*
      OLD DATA FORMAT:
      A  Date generated
      B  Year
      C  Month
      D  MCWD Date from
      E  MCWD Date to
      F  MCWD Amount
      G  Electricity Date from
      H  Electricity Date to
      I  Electricity Amount
      J  Manpower
      K  Total Expense
      L  Total Water Consumption
      M  Rate/cubic meter
      N onward = Unit IDs

      NEW DATA FORMAT:
      A  DATE GENERATED
      B  YEAR
      C  MONTH
      D  UNIT ID
      E  METER NUMBER
      F  PREVIOUS READING
      G  CURRENT READING
      H  CONSUMPTION
      I  RATE/CUBIC
      J  WATER BILL AMOUNT
    */

    var OLD_DATA_SHEET_NAME = 'WaterReadingData'; // change if needed
    var NEW_DATA_SHEET_NAME = 'Water Reading Data Store'; // change if needed

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var oldSheet = ss.getSheetByName(OLD_DATA_SHEET_NAME);
    var newSheet = ss.getSheetByName(NEW_DATA_SHEET_NAME);

    if (!oldSheet) {
      alert_('Old data sheet was not found: ' + OLD_DATA_SHEET_NAME);
      return;
    }

    if (!newSheet) {
      newSheet = ss.insertSheet(NEW_DATA_SHEET_NAME);
    }

    var oldData = oldSheet.getDataRange().getValues();

    if (oldData.length < 2) {
      alert_('Old data sheet is empty.');
      return;
    }

    var header = oldData[0];

    // Unit ID columns start at column N, index 13.
    var UNIT_START_COL = 13;

    var output = [];

    output.push([
      'DATE GENERATED',
      'YEAR',
      'MONTH',
      'UNIT ID',
      'METER NUMBER',
      'PREVIOUS READING',
      'CURRENT READING',
      'CONSUMPTION',
      'RATE/CUBIC',
      'WATER BILL AMOUNT'
    ]);

    // Optional: get meter numbers from Masterlist.
    // This assumes MASTERLIST columns:
    // A = UNIT ID
    // J = METER NUMBER
    var meterMap = getMeterNumberMap_();

    for (var rowIndex = 1; rowIndex < oldData.length; rowIndex++) {
      var row = oldData[rowIndex];

      var dateGenerated = row[0];
      var year = row[1];
      var month = row[2];
      var rate = Number(row[12]) || 0;

      // Skip empty month rows.
      if (!dateGenerated && !year && !month) {
        continue;
      }

      for (var colIndex = UNIT_START_COL; colIndex < header.length; colIndex++) {
        var unitId = String(header[colIndex] || '').trim();

        if (!unitId) {
          continue;
        }

        // Skip non-unit columns, if any exist after the unit list.
        if (!isValidUnitId_(unitId)) {
          continue;
        }

        var currentReading = toNumber_(row[colIndex]);

        // Skip totally blank readings.
        if (row[colIndex] === '' || row[colIndex] === null) {
          continue;
        }

        var previousReading = 0;

        // Previous reading comes from the previous old-data row for the same unit.
        if (rowIndex > 1) {
          previousReading = toNumber_(oldData[rowIndex - 1][colIndex]);
        }

        var consumption = currentReading - previousReading;

        // Prevent negative consumption if old source has reset/error values.
        if (consumption < 0) {
          consumption = 0;
        }

        var waterBillAmount = roundMoney_(consumption * rate);

        output.push([
          dateGenerated,
          year,
          month,
          unitId,
          meterMap[unitId] || '',
          previousReading,
          currentReading,
          consumption,
          rate,
          waterBillAmount
        ]);
      }
    }

    newSheet.clearContents();

    newSheet
      .getRange(1, 1, output.length, output[0].length)
      .setValues(output);

    newSheet.setFrozenRows(1);
    newSheet.autoResizeColumns(1, output[0].length);

    toast_('Converted ' + (output.length - 1) + ' row(s).', 'Data Store');
    alert_('Converted ' + (output.length - 1) + ' row(s) from old data to the new data store.');

  } catch (err) {
    alert_('Error converting old data:\n' + err.message);
    Logger.log(err);
  }
}

function getMeterNumberMap_() {
  var map = {};

  try {
    var sheet = getSheet_(SH.MASTERLIST);

    if (!sheet || sheet.getLastRow() < 2) {
      return map;
    }

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();

    data.forEach(function(row) {
      var unitId = String(row[0] || '').trim();   // Column A
      var meterNo = String(row[9] || '').trim();  // Column J

      if (unitId) {
        map[unitId] = meterNo;
      }
    });

  } catch (e) {
    Logger.log('Meter number map skipped: ' + e.message);
  }

  return map;
}

function isValidUnitId_(value) {
  value = String(value || '').trim();

  // Accepts unit IDs like:
  // P1B1L1
  // P1B6L1&2
  // P2B3L24&25
  return /^P\d+B\d+L[\d&]+$/i.test(value);
}

function toNumber_(value) {
  if (value === '' || value === null || typeof value === 'undefined') {
    return 0;
  }

  if (typeof value === 'number') {
    return value;
  }

  value = String(value).replace(/,/g, '').trim();

  var num = Number(value);
  return isNaN(num) ? 0 : num;
}

function roundMoney_(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}