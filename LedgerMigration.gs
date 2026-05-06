/*******************************************************
 * LedgerMigration.gs
 * AHNHAI Billing System
 *
 * Purpose:
 * Migrates old per-unit ledger sheets into the new system:
 *
 * 1. Monthly Dues  → _DuesLedger
 * 2. Water Bill    → _WaterLedger
 *
 * Source structure:
 * - Phase 1 folder contains block spreadsheets.
 * - Phase 2 folder contains block spreadsheets.
 * - Common accounts folder contains GUARDHOUSE, CLUBHOUSE, CHAPEL.
 * - Each spreadsheet contains unit tabs.
 *
 * Important behavior:
 * - Clears only _WaterLedger and _DuesLedger.
 * - Does NOT clear Water Reading Data Store.
 * - Does NOT clear Rate Calculator.
 * - Copies old balances exactly.
 * - Copies old dates exactly.
 * - Copies old remarks exactly.
 * - Copies old add-on values into ADDON MCWD.
 * - GUARDHOUSE, CLUBHOUSE, and CHAPEL migrate water only.
 * - GUARDHOUSE, CLUBHOUSE, and CHAPEL do NOT migrate monthly dues.
 * - Creates a Migration Report sheet.
 *******************************************************/


function migrateOldUnitLedgersFromFolders() {
  var SOURCE_FOLDERS = [
    {
      label: 'Phase 1',
      folderId: '1KwWboN0tvj2_rmns6qrp2LSn26Sjyeu9',
      migrateDues: true,
      migrateWater: true
    },
    {
      label: 'Phase 2',
      folderId: '15YDk9ZNxRmeRftZndNjjXkHFowOp_qFO',
      migrateDues: true,
      migrateWater: true
    },
    {
      label: 'Common Accounts',
      folderId: '1Ecbb72bavWiLddfiPtH4Ymg8m51z1zDC',
      migrateDues: false,
      migrateWater: true
    }
  ];

  try {
    var targetSS = ss_();

    var waterLedgerSheet = getSheet_(SH._WL);
    var duesLedgerSheet = getSheet_(SH._DL);

    if (!waterLedgerSheet || !duesLedgerSheet) {
      alert_(
        'Required target ledger sheets are missing.\n\n' +
        'Please run Initial Setup first.'
      );
      return;
    }

    var proceed = lm_confirm_(
      'Migrate Old Unit Ledgers',
      'This will CLEAR and REBUILD only these two sheets:\n\n' +
      '- _WaterLedger\n' +
      '- _DuesLedger\n\n' +
      'It will scan the Phase 1, Phase 2, and Common Account folders.\n\n' +
      'Continue?'
    );

    if (!proceed) return;

    lm_clearDataRowsOnly_(waterLedgerSheet);
    lm_clearDataRowsOnly_(duesLedgerSheet);

    var allWaterRows = [];
    var allDuesRows = [];
    var reportRows = [];

    reportRows.push([
      'TIMESTAMP',
      'FOLDER GROUP',
      'FILE NAME',
      'FILE ID',
      'SHEET NAME',
      'UNIT ID',
      'STATUS',
      'DUES ROWS',
      'WATER ROWS',
      'MESSAGE'
    ]);

    SOURCE_FOLDERS.forEach(function(source) {
      var folder;

      try {
        folder = DriveApp.getFolderById(source.folderId);
      } catch (folderErr) {
        reportRows.push(lm_reportRow_(
          source.label,
          '',
          source.folderId,
          '',
          '',
          'ERROR',
          0,
          0,
          'Cannot open folder: ' + folderErr.message
        ));
        return;
      }

      var files = folder.getFiles();

      while (files.hasNext()) {
        var file = files.next();

        if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
          reportRows.push(lm_reportRow_(
            source.label,
            file.getName(),
            file.getId(),
            '',
            '',
            'SKIPPED',
            0,
            0,
            'Not a native Google Sheets file.'
          ));
          continue;
        }

        var sourceSS;

        try {
          sourceSS = SpreadsheetApp.openById(file.getId());
        } catch (openErr) {
          reportRows.push(lm_reportRow_(
            source.label,
            file.getName(),
            file.getId(),
            '',
            '',
            'ERROR',
            0,
            0,
            'Cannot open spreadsheet: ' + openErr.message
          ));
          continue;
        }

        var sheets = sourceSS.getSheets();

        sheets.forEach(function(sheet) {
          var sheetName = sheet.getName();
          var unitId = lm_normalizeUnitId_(sheetName);

          if (!lm_isBillableLedgerTab_(unitId)) {
            reportRows.push(lm_reportRow_(
              source.label,
              file.getName(),
              file.getId(),
              sheetName,
              '',
              'SKIPPED',
              0,
              0,
              'Sheet name is not a billable unit/common account.'
            ));
            return;
          }

          var isCommon = lm_isCommonAccount_(unitId);

          try {
            var parsed = lm_parseOldUnitLedgerSheet_(
              sheet,
              unitId,
              {
                migrateDues: source.migrateDues && !isCommon,
                migrateWater: source.migrateWater
              }
            );

            Array.prototype.push.apply(allDuesRows, parsed.duesRows);
            Array.prototype.push.apply(allWaterRows, parsed.waterRows);

            reportRows.push(lm_reportRow_(
              source.label,
              file.getName(),
              file.getId(),
              sheetName,
              unitId,
              'IMPORTED',
              parsed.duesRows.length,
              parsed.waterRows.length,
              parsed.message
            ));

          } catch (sheetErr) {
            reportRows.push(lm_reportRow_(
              source.label,
              file.getName(),
              file.getId(),
              sheetName,
              unitId,
              'ERROR',
              0,
              0,
              sheetErr.message
            ));
          }
        });
      }
    });

    if (allWaterRows.length > 0) {
      lm_appendRows_(waterLedgerSheet, allWaterRows);
    }

    if (allDuesRows.length > 0) {
      lm_appendRows_(duesLedgerSheet, allDuesRows);
    }

    lm_writeMigrationReport_(targetSS, reportRows);

    /*
      Important:
      Do NOT call recalcWaterBalances() or recalcDuesBalances().
      The instruction is to copy old balances exactly.
    */

    try {
      refreshMonthlySummary();
    } catch (summaryErr) {
      Logger.log('Monthly Summary refresh skipped: ' + summaryErr.message);
    }

    alert_(
      'Old ledger migration complete!\n\n' +
      'Water ledger rows imported: ' + allWaterRows.length + '\n' +
      'Dues ledger rows imported: ' + allDuesRows.length + '\n\n' +
      'A Migration Report sheet was created/updated.\n\n' +
      'Note: Balances were copied exactly from the old ledgers.'
    );

  } catch (err) {
    Logger.log(err);
    alert_('Ledger migration failed:\n' + err.message);
  }
}


/**
 * Parses one old unit ledger sheet.
 */
function lm_parseOldUnitLedgerSheet_(sheet, unitId, options) {
  var values = sheet.getDataRange().getValues();

  if (!values || values.length === 0) {
    return {
      duesRows: [],
      waterRows: [],
      message: 'Empty sheet.'
    };
  }

  var headerInfo = lm_findLedgerHeaderInfo_(values);

  if (!headerInfo) {
    throw new Error('Could not find old ledger table headers.');
  }

  var duesRows = [];
  var waterRows = [];

  if (options.migrateDues) {
    duesRows = lm_parseDuesRows_(values, headerInfo, unitId);
  }

  if (options.migrateWater) {
    waterRows = lm_parseWaterRows_(values, headerInfo, unitId);
  }

  var msgParts = [];

  if (!options.migrateDues) {
    msgParts.push('Dues skipped.');
  }

  if (!options.migrateWater) {
    msgParts.push('Water skipped.');
  }

  if (msgParts.length === 0) {
    msgParts.push('Imported successfully.');
  }

  return {
    duesRows: duesRows,
    waterRows: waterRows,
    message: msgParts.join(' ')
  };
}


/**
 * Finds header row and column mappings for the old side-by-side ledgers.
 */
function lm_findLedgerHeaderInfo_(values) {
  for (var r = 0; r < Math.min(values.length, 30); r++) {
    var row = values[r];
    var norms = row.map(function(v) {
      return lm_normHeader_(v);
    });

    var billDateCol = lm_findExactNorm_(norms, 'BILLDATE');

    if (billDateCol === -1) {
      continue;
    }

    var waterStartCol = lm_findPaymentDateBefore_(norms, billDateCol);
    if (waterStartCol === -1) {
      waterStartCol = billDateCol;
    }

    var duesEndCol = waterStartCol - 1;

    var duesMonthCol = lm_findExactNormInRange_(norms, 'MONTH', 0, duesEndCol);
    var duesDebitCol = lm_findContainsAllInRange_(norms, ['DEBIT', 'MONTHLY', 'DUES'], 0, duesEndCol);

    if (duesMonthCol === -1 || duesDebitCol === -1) {
      /*
        Water-only sheets are still valid for GUARDHOUSE, CLUBHOUSE, CHAPEL.
        So do not reject yet.
      }

      */
    }

    var waterDebitCol = lm_findContainsAllInRange_(norms, ['DEBIT', 'WATER', 'BILL'], waterStartCol, row.length - 1);
    var presentReadingCol = lm_findContainsAllInRange_(norms, ['PRESENT', 'METER', 'READING'], waterStartCol, row.length - 1);

    if (waterDebitCol === -1 && presentReadingCol === -1) {
      continue;
    }

    return {
      headerRow: r,

      dues: {
        paymentDate: lm_findExactNormInRange_(norms, 'PAYMENTDATE', 0, duesEndCol),
        month: duesMonthCol,
        debit: duesDebitCol,
        credit: lm_findContainsAllInRange_(norms, ['CREDIT', 'PAYMENTS'], 0, duesEndCol),
        balance: lm_findExactNormInRange_(norms, 'BALANCE', 0, duesEndCol),
        orNumber: lm_findContainsAnyInRange_(norms, ['REFERENC', 'ORNUMBER'], 0, duesEndCol),
        remarks: lm_findExactNormInRange_(norms, 'REMARKS', 0, duesEndCol)
      },

      water: {
        paymentDate: lm_findExactNormInRange_(norms, 'PAYMENTDATE', waterStartCol, row.length - 1),
        billDate: billDateCol,
        prevReadingDate: lm_findContainsAllInRange_(norms, ['PREVIOUS', 'READING', 'DATE'], waterStartCol, row.length - 1),
        presentReadingDate: lm_findContainsAllInRange_(norms, ['PRESENT', 'READING', 'DATE'], waterStartCol, row.length - 1),
        prevReading: lm_findContainsAllInRange_(norms, ['PREVIOUS', 'METER', 'READING'], waterStartCol, row.length - 1),
        presentReading: presentReadingCol,
        rate: lm_findContainsAllInRange_(norms, ['RATE', 'CUBIC'], waterStartCol, row.length - 1),
        dueDate: lm_findExactNormInRange_(norms, 'DUEDATE', waterStartCol, row.length - 1),
        penalty: lm_findExactNormInRange_(norms, 'PENALTY', waterStartCol, row.length - 1),
        debit: waterDebitCol,
        credit: lm_findContainsAllInRange_(norms, ['CREDIT', 'PAYMENTS'], waterStartCol, row.length - 1),
        balance: lm_findExactNormInRange_(norms, 'BALANCE', waterStartCol, row.length - 1),
        addon: lm_findContainsAllInRange_(norms, ['ADD', 'ON'], waterStartCol, row.length - 1),
        orNumber: lm_findContainsAnyInRange_(norms, ['REFERENC', 'ORNUMBER'], waterStartCol, row.length - 1),
        remarks: lm_findExactNormInRange_(norms, 'REMARKS', waterStartCol, row.length - 1)
      }
    };
  }

  return null;
}


/**
 * Parses old Monthly Dues rows into _DuesLedger format.
 *
 * New _DuesLedger format:
 * A UNIT ID
 * B YEAR
 * C MONTH
 * D PAYMENT DATE
 * E DEBIT
 * F CREDIT
 * G BALANCE
 * H OR NUMBER
 * I REMARKS
 */
function lm_parseDuesRows_(values, headerInfo, unitId) {
  var c = headerInfo.dues;

  if (c.month === -1) {
    return [];
  }

  var rows = [];

  for (var r = headerInfo.headerRow + 1; r < values.length; r++) {
    var row = values[r];

    var monthRaw = lm_getCell_(row, c.month);

    if (lm_isBlank_(monthRaw)) {
      continue;
    }

    var parsedMonth = lm_parseDuesMonth_(monthRaw);

    rows.push([
      unitId,
      parsedMonth.year,
      parsedMonth.month,
      lm_getCell_(row, c.paymentDate),
      lm_getCell_(row, c.debit),
      lm_getCell_(row, c.credit),
      lm_getCell_(row, c.balance),
      lm_getCell_(row, c.orNumber),
      lm_getCell_(row, c.remarks)
    ]);
  }

  return rows;
}


/**
 * Parses old Water Bill rows into _WaterLedger format.
 *
 * New _WaterLedger format:
 * A  UNIT ID
 * B  YEAR
 * C  MONTH
 * D  BILL DATE
 * E  PREV READING DATE
 * F  PRESENT READING DATE
 * G  PREV READING
 * H  PRESENT READING
 * I  RATE/CUBIC
 * J  DUE DATE
 * K  PENALTY
 * L  DEBIT
 * M  CREDIT
 * N  BALANCE
 * O  ADDON MCWD
 * P  OR NUMBER
 * Q  REMARKS
 * R  BILL NUMBER
 * S  PAYMENT DATE
 */
function lm_parseWaterRows_(values, headerInfo, unitId) {
  var c = headerInfo.water;
  var rows = [];

  for (var r = headerInfo.headerRow + 1; r < values.length; r++) {
    var row = values[r];

    var paymentDate = lm_getCell_(row, c.paymentDate);
    var billDate = lm_getCell_(row, c.billDate);
    var prevReadingDate = lm_getCell_(row, c.prevReadingDate);
    var presentReadingDate = lm_getCell_(row, c.presentReadingDate);
    var prevReading = lm_getCell_(row, c.prevReading);
    var presentReading = lm_getCell_(row, c.presentReading);
    var rate = lm_getCell_(row, c.rate);
    var dueDate = lm_getCell_(row, c.dueDate);
    var penalty = lm_getCell_(row, c.penalty);
    var debit = lm_getCell_(row, c.debit);
    var credit = lm_getCell_(row, c.credit);
    var balance = lm_getCell_(row, c.balance);
    var addon = lm_getCell_(row, c.addon);
    var orNumber = lm_getCell_(row, c.orNumber);
    var remarks = lm_getCell_(row, c.remarks);

    var hasWaterData = !lm_allBlank_([
      paymentDate,
      billDate,
      prevReadingDate,
      presentReadingDate,
      prevReading,
      presentReading,
      rate,
      dueDate,
      penalty,
      debit,
      credit,
      balance,
      addon,
      orNumber,
      remarks
    ]);

    if (!hasWaterData) {
      continue;
    }

    var derived = lm_deriveYearMonthFromWaterRow_(billDate, presentReadingDate, paymentDate);
    var billNo = lm_buildBillNumberForLedger_(unitId, derived.year, derived.month);

    rows.push([
      unitId,
      derived.year,
      derived.month,
      billDate,
      prevReadingDate,
      presentReadingDate,
      prevReading,
      presentReading,
      rate,
      dueDate,
      penalty,
      debit,
      credit,
      balance,
      addon,
      orNumber,
      remarks,
      billNo,
      paymentDate
    ]);
  }

  return rows;
}


/**
 * Derives YEAR and MONTH from:
 * Bill Date → Present Reading Date → Payment Date
 */
function lm_deriveYearMonthFromWaterRow_(billDate, presentReadingDate, paymentDate) {
  var candidates = [billDate, presentReadingDate, paymentDate];

  for (var i = 0; i < candidates.length; i++) {
    var d = lm_toDate_(candidates[i]);

    if (d) {
      return {
        year: d.getFullYear(),
        month: lm_monthName_(d.getMonth() + 1)
      };
    }
  }

  return {
    year: '',
    month: ''
  };
}


/**
 * Parses dues month values:
 * - July 2022 → YEAR 2022, MONTH July
 * - Registration → YEAR blank, MONTH Registration
 */
function lm_parseDuesMonth_(value) {
  if (value instanceof Date && !isNaN(value)) {
    return {
      year: value.getFullYear(),
      month: lm_monthName_(value.getMonth() + 1)
    };
  }

  var s = String(value || '').trim();

  if (!s) {
    return {
      year: '',
      month: ''
    };
  }

  var m = s.match(/^([A-Za-z]+)\s+(\d{4})$/);

  if (m) {
    return {
      year: parseInt(m[2], 10),
      month: lm_titleCase_(m[1])
    };
  }

  return {
    year: '',
    month: s
  };
}


/**
 * Builds bill number for migrated water ledger rows.
 */
function lm_buildBillNumberForLedger_(unitId, year, monthName) {
  if (!year || !monthName) return '';

  var monthNum = lm_monthNumber_(monthName);
  if (!monthNum) return '';

  if (lm_isRegularUnitId_(unitId)) {
    var parsed = parseUID(unitId);

    if (parsed) {
      return buildBillNum(year, monthNum, parsed.phase, parsed.block, parsed.lot);
    }
  }

  return lm_buildCommonAccountBillNum_(year, monthNum, unitId);
}


/**
 * Builds bill number for common account ledgers.
 */
function lm_buildCommonAccountBillNum_(year, monthNum, unitId) {
  var cleanId = String(unitId || '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, '-');

  return 'WB-' + year + '-' + String(monthNum).padStart(2, '0') + '-' + cleanId;
}


/**
 * Clears data rows only, preserving headers.
 */
function lm_clearDataRowsOnly_(sh) {
  if (!sh || sh.getLastRow() < 2) return;

  sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
    .clearContent()
    .clearFormat();
}


/**
 * Appends rows to a sheet.
 */
function lm_appendRows_(sh, rows) {
  if (!sh || !rows || rows.length === 0) return;

  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length)
    .setValues(rows);
}


/**
 * Writes migration report.
 */
function lm_writeMigrationReport_(ss, reportRows) {
  var name = 'Migration Report';
  var sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
  }

  sh.clearContents();
  sh.clearFormats();

  sh.getRange(1, 1, reportRows.length, reportRows[0].length)
    .setValues(reportRows);

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, reportRows[0].length);
}


/**
 * Creates one report row.
 */
function lm_reportRow_(folderGroup, fileName, fileId, sheetName, unitId, status, duesCount, waterCount, message) {
  return [
    new Date(),
    folderGroup || '',
    fileName || '',
    fileId || '',
    sheetName || '',
    unitId || '',
    status || '',
    duesCount || 0,
    waterCount || 0,
    message || ''
  ];
}


/**
 * Valid tab names:
 * - P1B1L1
 * - P2B2L7&8
 * - GUARDHOUSE
 * - CLUBHOUSE
 * - CHAPEL
 */
function lm_isBillableLedgerTab_(value) {
  value = lm_normalizeUnitId_(value);

  if (lm_isRegularUnitId_(value)) return true;
  if (lm_isCommonAccount_(value)) return true;

  return false;
}


function lm_isRegularUnitId_(value) {
  value = lm_normalizeUnitId_(value);
  return /^P\d+B\d+L[\d&]+$/i.test(value);
}


function lm_isCommonAccount_(value) {
  value = lm_normalizeUnitId_(value);

  return [
    'GUARDHOUSE',
    'CLUBHOUSE',
    'CHAPEL'
  ].indexOf(value) !== -1;
}


function lm_normalizeUnitId_(value) {
  return String(value || '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, '');
}


/**
 * Header normalization.
 */
function lm_normHeader_(value) {
  return String(value || '')
    .toUpperCase()
    .replace(/\s+/g, '')
    .replace(/[^A-Z0-9]/g, '');
}


function lm_findExactNorm_(norms, exact) {
  for (var i = 0; i < norms.length; i++) {
    if (norms[i] === exact) return i;
  }

  return -1;
}


function lm_findExactNormInRange_(norms, exact, start, end) {
  start = Math.max(0, start);
  end = Math.min(norms.length - 1, end);

  for (var i = start; i <= end; i++) {
    if (norms[i] === exact) return i;
  }

  return -1;
}


function lm_findContainsAllInRange_(norms, parts, start, end) {
  start = Math.max(0, start);
  end = Math.min(norms.length - 1, end);

  for (var i = start; i <= end; i++) {
    var ok = true;

    for (var p = 0; p < parts.length; p++) {
      if (norms[i].indexOf(parts[p]) === -1) {
        ok = false;
        break;
      }
    }

    if (ok) return i;
  }

  return -1;
}


function lm_findContainsAnyInRange_(norms, parts, start, end) {
  start = Math.max(0, start);
  end = Math.min(norms.length - 1, end);

  for (var i = start; i <= end; i++) {
    for (var p = 0; p < parts.length; p++) {
      if (norms[i].indexOf(parts[p]) !== -1) {
        return i;
      }
    }
  }

  return -1;
}


/**
 * Finds Payment Date column immediately before Bill Date.
 */
function lm_findPaymentDateBefore_(norms, billDateCol) {
  for (var i = billDateCol - 1; i >= 0; i--) {
    if (norms[i] === 'PAYMENTDATE') {
      return i;
    }
  }

  return -1;
}


function lm_getCell_(row, col) {
  if (col === -1 || col === null || typeof col === 'undefined') return '';
  return row[col];
}


function lm_isBlank_(value) {
  return value === '' || value === null || typeof value === 'undefined';
}


function lm_allBlank_(values) {
  for (var i = 0; i < values.length; i++) {
    if (!lm_isBlank_(values[i])) {
      return false;
    }
  }

  return true;
}


/**
 * Tries to convert a value to Date for deriving year/month only.
 * Does not change the value that gets migrated.
 */
function lm_toDate_(value) {
  if (value instanceof Date && !isNaN(value)) {
    return value;
  }

  if (lm_isBlank_(value)) {
    return null;
  }

  var s = String(value).trim();

  if (!s) {
    return null;
  }

  var d = new Date(s);

  if (isNaN(d)) {
    return null;
  }

  return d;
}


function lm_monthName_(monthNum) {
  var names = [
    '',
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December'
  ];

  return names[monthNum] || '';
}


function lm_monthNumber_(monthName) {
  var s = String(monthName || '').trim().toLowerCase();

  var names = {
    january: 1,
    february: 2,
    march: 3,
    april: 4,
    may: 5,
    june: 6,
    july: 7,
    august: 8,
    september: 9,
    october: 10,
    november: 11,
    december: 12
  };

  return names[s] || 0;
}


function lm_titleCase_(value) {
  value = String(value || '').toLowerCase();

  if (!value) return '';

  return value.charAt(0).toUpperCase() + value.slice(1);
}


/**
 * Confirmation wrapper.
 */
function lm_confirm_(title, message) {
  try {
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(title, message, ui.ButtonSet.YES_NO);
    return result === ui.Button.YES;
  } catch (e) {
    return true;
  }
}