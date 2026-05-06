/*******************************************************
 * Migration.gs
 * AHNHAI Billing System
 *
 * Purpose:
 * Migrates old wide-format WaterReadingData into the new
 * row-based system data stores:
 *
 * 1. Water Reading Data Store
 * 2. _WaterLedger
 * 3. _DuesLedger
 * 4. Rate Calculator
 *
 * Also includes billable common accounts:
 * - GUARDHOUSE
 * - CLUBHOUSE
 * - CHAPEL
 *
 * Excludes non-billable/source-meter accounts:
 * - MCWD
 * - SUBMERSIBLE
 * - SUBMERSIBLE 2
 *
 * Water bill rule:
 * - If consumption = 0, water bill = 0
 * - If consumption > 0, water bill = max(MIN_WATER_BILL, consumption × rate)
 *
 * IMPORTANT:
 * - Make a backup copy of the spreadsheet before running.
 * - Run Initial Setup first.
 * - Import Masterlist first.
 * - Keep your old wide-format data in a sheet named:
 *   WaterReadingData
 *******************************************************/


function migrateOldWideDataToNewSystem() {
  /*
    OLD WATERREADINGDATA FORMAT:
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
    N onward = Unit IDs and common accounts

    NEW TARGETS:
    - Water Reading Data Store
    - _WaterLedger
    - _DuesLedger
    - Rate Calculator
  */

  var OLD_DATA_SHEET_NAME = 'WaterReadingData';

  /*
    SAFETY:
    false = append migrated rows after existing rows
    true  = clear existing new data stores first, then migrate

    Recommended:
    - Use true only if this is a fresh/new spreadsheet.
    - Make a backup copy before using true.
  */
  var REPLACE_EXISTING_NEW_DATA = false;

  /*
    If true, migrated water ledger debits follow your billing rule:
    - consumption = 0 → debit = 0
    - consumption > 0 → debit = max(MIN_WATER_BILL, consumption × rate)

    If false:
    - consumption = 0 → debit = 0
    - consumption > 0 → debit = consumption × rate
  */
  var APPLY_MIN_WATER_BILL = true;

  try {
    var ss = ss_();

    var oldSh = ss.getSheetByName(OLD_DATA_SHEET_NAME);
    if (!oldSh) {
      alert_('Old data sheet was not found: ' + OLD_DATA_SHEET_NAME);
      return;
    }

    var wsSh = getSheet_(SH.W_STORE);
    var wlSh = getSheet_(SH._WL);
    var dlSh = getSheet_(SH._DL);
    var rcSh = getSheet_(SH.RATE_CALC);

    if (!wsSh || !wlSh || !dlSh || !rcSh) {
      alert_(
        'Required new data store sheets are missing.\n\n' +
        'Run Initial Setup first before migration.'
      );
      return;
    }

    if (REPLACE_EXISTING_NEW_DATA) {
      clearDataRowsOnly_(wsSh);
      clearDataRowsOnly_(wlSh);
      clearDataRowsOnly_(dlSh);
      clearDataRowsOnly_(rcSh);
    } else {
      var hasExisting =
        wsSh.getLastRow() > 1 ||
        wlSh.getLastRow() > 1 ||
        dlSh.getLastRow() > 1 ||
        rcSh.getLastRow() > 1;

      if (hasExisting) {
        var proceed = confirmMigration_(
          'Existing Data Found',
          'The new data stores already contain data.\n\n' +
          'This migration will APPEND old records, not replace them.\n\n' +
          'Continue?'
        );

        if (!proceed) return;
      }
    }

    var oldData = oldSh.getDataRange().getValues();

    if (oldData.length < 2) {
      alert_('Old data sheet has no data to migrate.');
      return;
    }

    var header = oldData[0];

    // Old data unit/common-account columns start at column N, zero-based index 13.
    var UNIT_START_COL = 13;

    var meterMap = buildMasterlistMeterMap_();

    var newWsRows = [];
    var newWlRows = [];
    var newDlRows = [];

    /*
      Rate Calculator is migrated directly from WaterReadingData.
      This copies the old monthly rate/expense rows directly.
    */
    var newRateRows = migrateRateCalculatorFromWaterReadingData_(oldData);

    var lastReadingByUnit = {};
    var lastReadingDateByUnit = {};

    var migratedBillingPeriods = 0;
    var migratedReadingRows = 0;
    var skippedBaselineRows = 0;

    for (var rIndex = 1; rIndex < oldData.length; rIndex++) {
      var row = oldData[rIndex];

      var dateGenerated = row[0];
      var year = parseInt(row[1], 10);
      var monthName = String(row[2] || '').trim();

      if (!dateGenerated && !year && !monthName) {
        continue;
      }

      var monthNum = getMonthNum(monthName);

      if (!year || !monthName || monthNum === 0) {
        Logger.log('Skipped row ' + (rIndex + 1) + ': invalid year/month.');
        continue;
      }

      var totalExpense = toNum(row[10]);
      var totalCons = toNum(row[11]);
      var rate = toNum(row[12]);

      /*
        Important:
        Some old files have a first row with readings but no expenses/rate.
        Example: May 2025 can be a baseline row only.
      */
      var isBillingRow = totalExpense > 0 || totalCons > 0 || rate > 0;

      var presentDateStr = fmtDate(dateGenerated);
      var billDateStr = fmtDate(getBillDate(year, monthNum));
      var dueDateStr = fmtDate(getDueDate(year, monthNum));

      var rowHasAnyUnitReading = false;
      var rowBillingCount = 0;

      for (var cIndex = UNIT_START_COL; cIndex < header.length; cIndex++) {
        var unitId = String(header[cIndex] || '').trim();

        if (!unitId) continue;
        if (!isMigrationBillableAccount_(unitId)) continue;

        var rawReading = row[cIndex];

        if (
          rawReading === '' ||
          rawReading === null ||
          typeof rawReading === 'undefined'
        ) {
          continue;
        }

        var currentReading = toNum(rawReading);
        rowHasAnyUnitReading = true;

        var previousReading = lastReadingByUnit.hasOwnProperty(unitId)
          ? lastReadingByUnit[unitId]
          : 0;

        var previousDate = lastReadingDateByUnit[unitId] || '';

        var consumption = currentReading - previousReading;
        if (consumption < 0) consumption = 0;

        /*
          Always update last reading tracker.
          This allows a baseline row to become the previous reading
          for the next billing row.
        */
        lastReadingByUnit[unitId] = currentReading;
        lastReadingDateByUnit[unitId] = presentDateStr;

        if (!isBillingRow) {
          continue;
        }

        var rawBill = consumption * rate;

        /*
          Confirmed rule:
          If consumption = 0, water bill amount = 0.
          If consumption > 0, water bill amount = max(₱250, consumption × rate).
        */
        var waterDebit = 0;

        if (consumption > 0) {
          waterDebit = APPLY_MIN_WATER_BILL
            ? Math.max(MIN_WATER_BILL, rawBill)
            : rawBill;
        }

        waterDebit = roundMoney_(waterDebit);

        var parsed = parseUID(unitId);

        var billNo = parsed
          ? buildBillNum(year, monthNum, parsed.phase, parsed.block, parsed.lot)
          : buildCommonAccountBillNum_(year, monthNum, unitId);

        var meterNo = meterMap[unitId] || '';

        // 1. New Water Reading Data Store row
        newWsRows.push([
          presentDateStr,
          year,
          monthName,
          unitId,
          meterNo,
          previousReading,
          currentReading,
          consumption,
          rate,
          waterDebit
        ]);

        // 2. Hidden _WaterLedger row
        newWlRows.push([
          unitId,                 // UNIT ID
          year,                   // YEAR
          monthName,              // MONTH
          billDateStr,            // BILL DATE
          previousDate,           // PREV READING DATE
          presentDateStr,         // PRESENT READING DATE
          previousReading,        // PREV READING
          currentReading,         // PRESENT READING
          rate,                   // RATE/CUBIC
          dueDateStr,             // DUE DATE
          0,                      // PENALTY
          waterDebit,             // DEBIT
          0,                      // CREDIT
          0,                      // BALANCE, recalculated later
          0,                      // ADDON MCWD
          '',                     // OR NUMBER
          'Migrated from old WaterReadingData',
          billNo,
          ''                      // PAYMENT DATE
        ]);

        // Association dues: skip for common accounts (GUARDHOUSE, CLUBHOUSE, CHAPEL)
        if (!isCommonBillableAccount_(unitId)) {
          newDlRows.push([
            unitId,
            year,
            monthName,
            '',
            ASSOC_DUES,
            0,
            0,
            '',
            'Migrated from old WaterReadingData'
          ]);
        }

        migratedReadingRows++;
        rowBillingCount++;
      }

      if (!isBillingRow && rowHasAnyUnitReading) {
        skippedBaselineRows++;
      }

      if (isBillingRow && rowBillingCount > 0) {
        migratedBillingPeriods++;
      }
    }

    if (newWsRows.length === 0) {
      alert_(
        'No billing rows were migrated.\n\n' +
        'Check that WaterReadingData has rate/expense values and unit readings.'
      );
      return;
    }

    appendMigrationRows_(wsSh, newWsRows);
    appendMigrationRows_(wlSh, newWlRows);
    appendMigrationRows_(dlSh, newDlRows);
    appendMigrationRows_(rcSh, newRateRows);

    recalcWaterBalances();
    recalcDuesBalances();

    refreshMonthlySummary();
    populateWaterInputTable();

    try {
      regenerateLatestMigratedBillPrint_(newRateRows);
    } catch (printErr) {
      Logger.log('Bill print regeneration skipped: ' + printErr.message);
    }

    alert_(
      'Migration complete!\n\n' +
      'Billing periods migrated: ' + migratedBillingPeriods + '\n' +
      'Unit/common account billing rows migrated: ' + migratedReadingRows + '\n' +
      'Rate Calculator rows migrated: ' + newRateRows.length + '\n' +
      'Baseline rows used as previous readings: ' + skippedBaselineRows + '\n\n' +
      'Included common accounts:\n' +
      '- GUARDHOUSE\n' +
      '- CLUBHOUSE\n' +
      '- CHAPEL\n\n' +
      'Excluded source-meter accounts:\n' +
      '- MCWD\n' +
      '- SUBMERSIBLE\n' +
      '- SUBMERSIBLE 2\n\n' +
      'Water bill rule applied:\n' +
      '- Consumption 0 = ₱0 water bill\n' +
      '- Consumption above 0 = ₱250 minimum or computed bill, whichever is higher\n\n' +
      'Updated:\n' +
      '- Water Reading Data Store\n' +
      '- _WaterLedger\n' +
      '- _DuesLedger\n' +
      '- Rate Calculator\n' +
      '- Monthly Summary\n' +
      '- Water Reading Input'
    );

  } catch (err) {
    Logger.log(err);
    alert_('Migration failed:\n' + err.message);
  }
}


/**
 * Migrates Rate Calculator directly from old WaterReadingData.
 */
function migrateRateCalculatorFromWaterReadingData_(oldData) {
  /*
    Source: old WaterReadingData wide-format sheet

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
  */

  var rows = [];

  for (var i = 1; i < oldData.length; i++) {
    var r = oldData[i];

    var dateGenerated = r[0];
    var year = r[1];
    var month = r[2];

    var mcwdFrom = r[3];
    var mcwdTo = r[4];
    var mcwdAmount = toNum(r[5]);

    var electricityFrom = r[6];
    var electricityTo = r[7];
    var electricityAmount = toNum(r[8]);

    var manpower = toNum(r[9]);
    var totalExpense = toNum(r[10]);
    var totalConsumption = toNum(r[11]);
    var ratePerCubic = toNum(r[12]);

    // Skip totally blank rows.
    if (!dateGenerated && !year && !month) {
      continue;
    }

    /*
      Skip baseline rows.
      Example: first month may have readings but no expense/rate data.
    */
    if (
      !mcwdAmount &&
      !electricityAmount &&
      !manpower &&
      !totalExpense &&
      !totalConsumption &&
      !ratePerCubic
    ) {
      continue;
    }

    // If total expense is missing but components are present, compute it.
    if (!totalExpense) {
      totalExpense = mcwdAmount + electricityAmount + manpower;
    }

    // If rate is missing but expense and consumption exist, compute it.
    if (!ratePerCubic && totalExpense && totalConsumption) {
      ratePerCubic = totalExpense / totalConsumption;
    }

    rows.push([
      dateGenerated ? fmtDate(dateGenerated) : '',
      year,
      month,
      mcwdFrom ? fmtDate(mcwdFrom) : '',
      mcwdTo ? fmtDate(mcwdTo) : '',
      mcwdAmount,
      electricityFrom ? fmtDate(electricityFrom) : '',
      electricityTo ? fmtDate(electricityTo) : '',
      electricityAmount,
      manpower,
      totalExpense,
      totalConsumption,
      ratePerCubic
    ]);
  }

  return rows;
}


/**
 * Regenerates bill print sheets using the latest migrated rate row.
 */
function regenerateLatestMigratedBillPrint_(rateRows) {
  if (!rateRows || rateRows.length === 0) return;

  var latest = rateRows[rateRows.length - 1];

  var latestYear = latest[1];
  var latestMonthName = latest[2];
  var latestMonthNum = getMonthNum(latestMonthName);

  var latestMcwdFrom = latest[3];
  var latestMcwdTo = latest[4];
  var latestMcwdAmt = latest[5];

  var latestElecFrom = latest[6];
  var latestElecTo = latest[7];
  var latestElecAmt = latest[8];

  var latestManpower = latest[9];
  var latestTotalCons = latest[11];
  var latestRate = latest[12];

  regenerateBillPrint(
    latestYear,
    latestMonthNum,
    latestMonthName,
    latestRate,
    latestTotalCons,
    latestMcwdAmt,
    latestElecAmt,
    latestManpower,
    latestMcwdFrom,
    latestMcwdTo,
    latestElecFrom,
    latestElecTo
  );
}


/**
 * Clears only data rows, preserving headers.
 */
function clearDataRowsOnly_(sh) {
  if (!sh || sh.getLastRow() < 2) return;

  sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
    .clearContent()
    .clearFormat();
}


/**
 * Appends migration rows to a target sheet.
 */
function appendMigrationRows_(sh, rows) {
  if (!sh || !rows || rows.length === 0) return;

  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length)
    .setValues(rows);
}


/**
 * Builds a UNIT ID / COMMON ACCOUNT ID → METER NUMBER map from Masterlist.
 */
function buildMasterlistMeterMap_() {
  var map = {};

  var sh = getSheet_(SH.MASTERLIST);
  if (!sh || sh.getLastRow() < 2) return map;

  /*
    Masterlist expected columns:
    A = UNIT ID
    J = METER NUMBER
  */
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();

  data.forEach(function(row) {
    var unitId = String(row[0] || '').trim();
    var meterNo = String(row[9] || '').trim();

    if (unitId) {
      map[unitId] = meterNo;
    }
  });

  return map;
}


/**
 * Determines which old WaterReadingData columns should be migrated
 * as billable accounts.
 *
 * Includes:
 * - Regular unit IDs: P1B1L1, P1B6L1&2, P2B4L19&20
 * - Common billable accounts: GUARDHOUSE, CLUBHOUSE, CHAPEL
 *
 * Excludes:
 * - MCWD
 * - SUBMERSIBLE
 * - SUBMERSIBLE 2
 */
function isMigrationBillableAccount_(value) {
  value = String(value || '').trim().toUpperCase();

  if (!value) return false;

  // Regular homeowner units.
  if (/^P\d+B\d+L[\d&]+$/i.test(value)) {
    return true;
  }

  // Billable common/community accounts.
  var billableCommonAccounts = [
    'GUARDHOUSE',
    'CLUBHOUSE',
    'CHAPEL'
  ];

  return billableCommonAccounts.indexOf(value) !== -1;
}


/**
 * Explicit helper for common billable accounts.
 */
function isCommonBillableAccount_(value) {
  value = String(value || '').trim().toUpperCase();

  return [
    'GUARDHOUSE',
    'CLUBHOUSE',
    'CHAPEL'
  ].indexOf(value) !== -1;
}


/**
 * Builds bill numbers for common accounts that do not follow
 * the P#B#L# unit ID pattern.
 */
function buildCommonAccountBillNum_(year, monthNum, unitId) {
  var cleanId = String(unitId || '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, '-');

  return 'WB-' + year + '-' + String(monthNum).padStart(2, '0') + '-' + cleanId;
}


/**
 * Confirmation wrapper.
 * Uses SpreadsheetApp.getUi() when available.
 */
function confirmMigration_(title, message) {
  try {
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(title, message, ui.ButtonSet.YES_NO);
    return result === ui.Button.YES;
  } catch (e) {
    /*
      Fallback for contexts where UI is unavailable.
      Returning true allows migration to continue.
      If you want stricter behavior, change this to false.
    */
    return true;
  }
}