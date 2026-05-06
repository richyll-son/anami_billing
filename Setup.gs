// ============================================================
// AHNHAI Billing System — Setup.gs
// Initial setup: create all 9 visible + 2 hidden sheets,
// apply headers / data-validation, import masterlist.
// ============================================================

function setupSystem() {
  if (!confirm_('Initial Setup',
    'This will create all required sheets and import the Masterlist.\n\nContinue?')) return;

  toast_('Creating sheets…', 'Setup');
  createAllSheets();
  toast_('Importing masterlist…', 'Setup');
  importMasterlistFromSource();
  toast_('Populating reading input table…', 'Setup');
  populateWaterInputTable();
  alert_('Setup complete! All sheets created and Masterlist imported.\n\n' +
         'Next step: fill in Water Reading Input, then use\n"Process & Generate Bills" from the menu.');
}

// ── Master sheet-creation orchestrator ───────────────────────
function createAllSheets() {
  _createHiddenSheets();
  _setupMasterlistSheet();
  _setupWaterInputSheet();
  _setupWaterStoreSheet();
  _setupRateCalcSheet();
  _setupUnitLedgerSheet();
  _setupPaymentLogSheet();
  _setupSummarySheet();
  _setupPrintSheet(SH.P1_PRINT, 'PHASE 1');
  _setupPrintSheet(SH.P2_PRINT, 'PHASE 2');
  _reorderSheets();
}

function _reorderSheets() {
  var order = [
    SH.MASTERLIST, SH.W_INPUT, SH.W_STORE, SH.RATE_CALC,
    SH.UNIT_LEDGER, SH.PAY_LOG, SH.SUMMARY, SH.P1_PRINT, SH.P2_PRINT,
    SH._WL, SH._DL
  ];
  var sp = ss_();
  order.forEach(function(name, i) {
    var sh = sp.getSheetByName(name);
    if (sh) { sp.setActiveSheet(sh); sp.moveActiveSheet(i + 1); }
  });
  sp.setActiveSheet(sp.getSheetByName(SH.MASTERLIST));
}

// ── Hidden ledger stores ──────────────────────────────────────
function _createHiddenSheets() {
  var wlH = ['UNIT ID','YEAR','MONTH','BILL DATE','PREV READING DATE','PRESENT READING DATE',
             'PREV READING','PRESENT READING','RATE/CUBIC','DUE DATE','PENALTY',
             'DEBIT','CREDIT','BALANCE','ADDON MCWD','OR NUMBER','REMARKS',
             'BILL NUMBER','PAYMENT DATE'];
  var dlH = ['UNIT ID','YEAR','MONTH','PAYMENT DATE','DEBIT','CREDIT',
             'BALANCE','OR NUMBER','REMARKS'];

  _initHiddenSheet(SH._WL, wlH, '#cccccc');
  _initHiddenSheet(SH._DL, dlH, '#cccccc');
}

function _initHiddenSheet(name, headers, bg) {
  var sh = orCreate_(name);
  sh.clearContents(); sh.clearFormats();
  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers]).setFontWeight('bold').setBackground(bg);
  sh.setFrozenRows(1);
  sh.hideSheet();
}

// ── Sheet 1: Masterlist ───────────────────────────────────────
function _setupMasterlistSheet() {
  var sh = orCreate_(SH.MASTERLIST);
  sh.clearContents(); sh.clearFormats();

  var H = ['UNIT ID','PHASE','BLOCK','LOT','LAST NAME','FIRST NAME',
           'DATE ACCEPTED','CONTACT NUMBER','EMAIL','METER NUMBER','STUBOUT NUMBER','ACTIVE'];
  sh.getRange(1, 1, 1, H.length).setValues([H])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.setFrozenRows(1);

  // Tip note
  sh.getRange('N1').setValue('Use menu: Import Masterlist to sync from source');
  sh.getRange('N1').setFontColor('#c62828').setFontWeight('bold');

  sh.setColumnWidth(1, 100); sh.setColumnWidth(5, 140); sh.setColumnWidth(6, 130);
  sh.setColumnWidth(9, 200); sh.setColumnWidth(10, 130); sh.setColumnWidth(11, 120);
}

// ── Sheet 2: Water Reading Input ──────────────────────────────
function _setupWaterInputSheet() {
  var sh = orCreate_(SH.W_INPUT);
  sh.clearContents(); sh.clearFormats();

  sh.getRange('A1').setValue('WATER READING INPUT')
    .setFontSize(14).setFontWeight('bold').setBackground('#e3f2fd');
  sh.getRange('A1:F1').setBackground('#e3f2fd');

  var labels = [
    [3,  'A', 'Year:'],
    [4,  'A', 'Month (reading month):'],
    [5,  'A', 'MCWD Date From:'],
    [6,  'A', 'MCWD Date To:'],
    [7,  'A', 'MCWD Amount (₱):'],
    [8,  'A', 'Electricity Date From:'],
    [9,  'A', 'Electricity Date To:'],
    [10, 'A', 'Electricity Amount (₱):'],
    [11, 'A', 'Manpower Amount (₱):'],
  ];
  labels.forEach(function(l) {
    sh.getRange(l[0], 1).setValue(l[2]).setFontWeight('bold');
  });

  var now = new Date();
  sh.getRange('B3').setValue(now.getFullYear());
  sh.getRange('B4').setValue(MONTHS[now.getMonth()]);
  sh.getRange('B11').setValue(MANPOWER_DEF);

  // Dropdowns
  var yearList = ['2024','2025','2026','2027','2028','2029','2030'];
  sh.getRange('B3').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(yearList, true).build());
  sh.getRange('B4').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(MONTHS, true).build());

  // Table header row 13
  var TH = ['UNIT ID','METER NUMBER','HOMEOWNER NAME',
            'CURRENT READING','PREVIOUS READING','CONSUMPTION (m³)'];
  sh.getRange(13, 1, 1, TH.length).setValues([TH])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.setFrozenRows(13);

  // Alert note
  sh.getRange('H3').setValue('→  Use menu: Process & Generate Bills')
    .setFontWeight('bold').setFontColor('#c62828');

  sh.setColumnWidth(1, 105); sh.setColumnWidth(2, 135); sh.setColumnWidth(3, 180);
  sh.setColumnWidth(4, 150); sh.setColumnWidth(5, 150); sh.setColumnWidth(6, 140);
}

// ── Sheet 3: Water Reading Data Store ─────────────────────────
function _setupWaterStoreSheet() {
  var sh = orCreate_(SH.W_STORE);
  sh.clearContents(); sh.clearFormats();

  var H = ['DATE GENERATED','YEAR','MONTH','UNIT ID','METER NUMBER',
           'PREVIOUS READING','CURRENT READING','CONSUMPTION','RATE/CUBIC','WATER BILL AMOUNT'];
  sh.getRange(1, 1, 1, H.length).setValues([H])
    .setFontWeight('bold').setBackground('#0d652d').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 130); sh.setColumnWidth(5, 130);
}

// ── Sheet 4: Rate Calculator ──────────────────────────────────
function _setupRateCalcSheet() {
  var sh = orCreate_(SH.RATE_CALC);
  sh.clearContents(); sh.clearFormats();

  var H = ['DATE GENERATED','YEAR','MONTH',
           'MCWD DATE FROM','MCWD DATE TO','MCWD AMOUNT',
           'ELECTRICITY DATE FROM','ELECTRICITY DATE TO','ELECTRICITY AMOUNT',
           'MANPOWER','TOTAL EXPENSE','TOTAL CONSUMPTION','RATE/CUBIC'];
  sh.getRange(1, 1, 1, H.length).setValues([H])
    .setFontWeight('bold').setBackground('#7b2d8b').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

// ── Sheet 5: Unit Ledger ──────────────────────────────────────
function _setupUnitLedgerSheet() {
  var sh = orCreate_(SH.UNIT_LEDGER);
  sh.clearContents(); sh.clearFormats();

  // Selector row
  sh.getRange('A1').setValue('Select Unit:').setFontWeight('bold').setFontSize(11);
  sh.getRange('B1').setValue('-- Select Unit ID --')
    .setBackground('#fff9c4').setFontWeight('bold');
  sh.getRange('A2').setValue('Homeowner:').setFontWeight('bold');

  // Section titles row 4
  sh.getRange(4, UL_WATER_COL, 1, 15)
    .merge().setValue('WATER BILL LEDGER')
    .setFontWeight('bold').setBackground('#1565c0').setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sh.getRange(4, UL_DUES_COL, 1, 7)
    .merge().setValue('ASSOCIATION DUES LEDGER')
    .setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // Column headers row 5
  var WH = ['Payment Date','Bill Date','Prev Reading Date','Present Reading Date',
            'Prev Reading','Present Reading','Rate/Cubic','Due Date','Penalty',
            'Debit','Credit','Balance','Add-On MCWD','OR Number','Remarks'];
  sh.getRange(5, UL_WATER_COL, 1, WH.length).setValues([WH])
    .setFontWeight('bold').setBackground('#bbdefb').setWrap(true);

  var DH = ['Payment Date','Month','Debit (₱500)','Credit','Balance','OR Number','Remarks'];
  sh.getRange(5, UL_DUES_COL, 1, DH.length).setValues([DH])
    .setFontWeight('bold').setBackground('#c8e6c9').setWrap(true);

  sh.setFrozenRows(5);
  sh.setRowHeight(5, 36);

  // Column widths
  var wCols = [100,90,130,130,100,110,80,90,70,90,90,90,90,90,130];
  wCols.forEach(function(w, i) { sh.setColumnWidth(UL_WATER_COL + i, w); });
  var dCols = [90,80,80,80,80,90,130];
  dCols.forEach(function(w, i) { sh.setColumnWidth(UL_DUES_COL + i, w); });

  // Freeze col A-B widths
  sh.setColumnWidth(1, 120);
}

// ── Sheet 6: Central Payment Log ─────────────────────────────
function _setupPaymentLogSheet() {
  var sh = orCreate_(SH.PAY_LOG);
  sh.clearContents(); sh.clearFormats();

  var H = ['DATE RECORDED','UNIT ID','HOMEOWNER NAME','PAYMENT DATE',
           'AMOUNT PAID','OR NUMBER','PAYMENT TYPE',
           'WATER AMOUNT','DUES AMOUNT','REMARKS','STATUS'];
  sh.getRange(1, 1, 1, H.length).setValues([H])
    .setFontWeight('bold').setBackground('#bf360c').setFontColor('#ffffff');
  sh.setFrozenRows(1);

  sh.getRange('G2:G1000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Water','Dues','Both'], true).build());

  sh.getRange('M1').setValue('→  Use menu: Post Payments')
    .setFontWeight('bold').setFontColor('#c62828');

  [1,2,3,4,5,6,7,8,9,10,11].forEach(function(c, i) {
    sh.setColumnWidth(c, [130,100,180,130,110,100,100,110,100,160,100][i]);
  });
}

// ── Sheet 7: Monthly Summary ──────────────────────────────────
function _setupSummarySheet() {
  var sh = orCreate_(SH.SUMMARY);
  sh.clearContents(); sh.clearFormats();

  sh.getRange('A1').setValue('ANAMI HOMES NORTH HOMEOWNERS ASSOCIATION — MONTHLY BILLING SUMMARY')
    .setFontSize(13).setFontWeight('bold');
  sh.getRange('A2').setValue('Use menu: Refresh Summary to update')
    .setFontColor('#777777').setFontStyle('italic');
  sh.getRange('A3').setValue('Filter: [All]   Phase 1   Phase 2  — use Data > Filter views');

  var H = ['UNIT ID','HOMEOWNER NAME','METER NO','CONSUMPTION','RATE',
           'WATER BILL','PENALTY','ARREARS (WATER)','ADD-ON MCWD','TOTAL WATER DUE',
           'ASSOC DUE','ARREARS (DUES)','TOTAL DUES DUE','GRAND TOTAL DUE',
           'PAYMENT STATUS','DUE DATE'];
  sh.getRange(4, 1, 1, H.length).setValues([H])
    .setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
  sh.setFrozenRows(4);
  sh.setColumnWidth(2, 180); sh.setColumnWidth(16, 110);
}

// ── Sheets 8-9: Phase Print ───────────────────────────────────
function _setupPrintSheet(name, phase) {
  var sh = orCreate_(name);
  sh.clearContents(); sh.clearFormats();
  sh.getRange('A1').setValue(phase + ' BILL PRINT  — Auto-generated. Do not edit manually.')
    .setFontColor('#888888').setFontStyle('italic');
  sh.setColumnWidth(1, 420);
  sh.setColumnWidth(2, 180);
}

// ── Import Masterlist from source ─────────────────────────────
var MIME_GSHEET = 'application/vnd.google-apps.spreadsheet';
var MIME_XLSX   = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
var MIME_XLS    = 'application/vnd.ms-excel';

function importMasterlistFromSource() {
  try {
    var SRC_MASTERLIST_GID = 586990653;

    // ── Verify file exists first ──
    var file;
    try {
      file = DriveApp.getFileById(SRC_MASTERLIST_ID);
    } catch (e) {
      alert_('Cannot find source masterlist in Google Drive.\n\n' +
             'File ID: ' + SRC_MASTERLIST_ID + '\n\n' +
             'Make sure the file is shared with your Google account and the ID in Code.gs is correct.');
      return;
    }

    // ── Check file type before opening with SpreadsheetApp ──
    var mime = file.getMimeType();

    if (mime === MIME_XLSX || mime === MIME_XLS) {
      alert_('Source masterlist is still an Excel file, not a native Google Spreadsheet.\n\n' +
             'File name: "' + file.getName() + '"\n\n' +
             'To fix:\n' +
             '1. Open the file in Google Drive.\n' +
             '2. Click File → Save as Google Sheets.\n' +
             '3. Open the newly created Google Sheets version.\n' +
             '4. Copy the new file ID from the URL.\n' +
             '5. Update SRC_MASTERLIST_ID in Code.gs.\n' +
             '6. Run Import Masterlist again.');
      return;
    }

    if (mime !== MIME_GSHEET) {
      alert_('Source file is not a native Google Spreadsheet.\n\n' +
             'File name: "' + file.getName() + '"\n' +
             'File type: ' + mime + '\n\n' +
             'Please update SRC_MASTERLIST_ID to point to a real Google Sheets file.');
      return;
    }

    // ── Open source spreadsheet only after confirming it is Google Sheets ──
    var srcSS;
    try {
      srcSS = SpreadsheetApp.openById(SRC_MASTERLIST_ID);
    } catch (e) {
      alert_('The source file exists, but Google Sheets cannot open it.\n\n' +
             'File ID: ' + SRC_MASTERLIST_ID + '\n' +
             'File name: "' + file.getName() + '"\n' +
             'File type: ' + mime + '\n\n' +
             'Please check that this is a native Google Spreadsheet and that your account has access.');
      return;
    }

    // ── Find the exact tab by gid ──
    var srcSheet = null;
    var sheets = srcSS.getSheets();

    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === SRC_MASTERLIST_GID) {
        srcSheet = sheets[i];
        break;
      }
    }

    if (!srcSheet) {
      alert_('Source masterlist tab was not found.\n\n' +
             'Spreadsheet ID: ' + SRC_MASTERLIST_ID + '\n' +
             'Expected tab gid: ' + SRC_MASTERLIST_GID + '\n\n' +
             'Please check that the source link uses the correct gid.');
      return;
    }

    var srcData = srcSheet.getDataRange().getValues();

    if (srcData.length < 2) {
      alert_('Source masterlist tab is empty.');
      return;
    }

    var destSheet = getSheet_(SH.MASTERLIST);
    var existing = {};

    if (destSheet.getLastRow() > 1) {
      destSheet.getRange(2, 1, destSheet.getLastRow() - 1, 1)
        .getValues()
        .flat()
        .forEach(function(v) {
          if (v) existing[v] = true;
        });
    }

    // Source columns (0-based):
    // PHASE=0, BLOCK=1, LOT=2, LAST=3, FIRST=4,
    // DATE=5, CONTACT=6, EMAIL=7, METER=8, STUBOUT=9
    var newRows = [];

    for (var rIndex = 1; rIndex < srcData.length; rIndex++) {
      var r = srcData[rIndex];

      if (!r[0] && !r[1] && !r[2]) continue;

      var uid = buildUID(r[0], r[1], String(r[2]));

      if (existing[uid]) continue;

      newRows.push([
        uid,
        r[0],
        r[1],
        String(r[2]),
        r[3] || '',
        r[4] || '',
        r[5] || '',
        r[6] || '',
        r[7] || '',
        r[8] || '',
        r[9] || '',
        'Yes'
      ]);
    }

    if (newRows.length === 0) {
      alert_('No new units found in source masterlist.');
      return;
    }

    var startRow = destSheet.getLastRow() + 1;

    destSheet.getRange(startRow, 1, newRows.length, 12).setValues(newRows);

    // Active dropdown for new rows
    destSheet.getRange(startRow, 12, newRows.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Yes', 'No'], true)
        .build()
    );

    _refreshUnitLedgerDropdown();

    toast_('Imported ' + newRows.length + ' unit(s).', 'Masterlist');
    alert_('Imported ' + newRows.length + ' new unit(s) from source masterlist.');

  } catch (err) {
    alert_('Error importing masterlist:\n' + err.message +
           '\n\nCheck that you have access to the source spreadsheet.');
    Logger.log(err);
  }
}

// Rebuild the Unit Ledger B1 dropdown from current masterlist
function _refreshUnitLedgerDropdown() {
  var mlSh = getSheet_(SH.MASTERLIST);
  var ulSh = getSheet_(SH.UNIT_LEDGER);
  if (!mlSh || !ulSh || mlSh.getLastRow() < 2) return;

  var unitIds = mlSh.getRange(2, 1, mlSh.getLastRow() - 1, 1)
    .getValues().flat().filter(function(v) { return v !== ''; });
  if (unitIds.length === 0) return;

  // GAS data validation limit: 500 items per rule; chunk if needed
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(unitIds.slice(0, 500), true).build();
  ulSh.getRange('B1').setDataValidation(rule);
}

// ── Populate Water Reading Input table ────────────────────────
function populateWaterInputTable() {
  var inputSh = getSheet_(SH.W_INPUT);
  var mlSh    = getSheet_(SH.MASTERLIST);
  if (!inputSh || !mlSh || mlSh.getLastRow() < 2) return;

  var mlData = mlSh.getRange(2, 1, mlSh.getLastRow() - 1, 12).getValues();

  // Last current reading per unit from Water Store
  var lastReading = {};
  var wsSh = getSheet_(SH.W_STORE);
  if (wsSh && wsSh.getLastRow() > 1) {
    var wsData = wsSh.getRange(2, 4, wsSh.getLastRow() - 1, 4).getValues();
    // cols: UNIT_ID(0), METER(1), PREV_READING(2), CUR_READING(3)
    wsData.forEach(function(r) {
      if (r[0]) lastReading[r[0]] = r[3]; // keep last (data is appended chronologically)
    });
  }

  // Build rows for active units
  var rows = [];
  mlData.forEach(function(r) {
    var uid = r[0];
    if (!uid || String(r[11]).toLowerCase() === 'no') return;
    var prev = (lastReading[uid] !== undefined) ? lastReading[uid] : '';
    rows.push([uid, r[9], ownerName(r), '', prev, '']); // D (current) blank for admin to fill
  });

  if (rows.length === 0) return;

  // Clear old table rows (keep header row 13 intact)
  var last = inputSh.getLastRow();
  if (last >= INPUT_TABLE_START) {
    inputSh.getRange(INPUT_TABLE_START, 1, last - INPUT_TABLE_START + 1, 6).clearContent();
  }

  inputSh.getRange(INPUT_TABLE_START, 1, rows.length, 6).setValues(rows);

  // Consumption formula (col F = col 6): =IF(AND(ISNUMBER(D#),ISNUMBER(E#)),MAX(0,D#-E#),"")
  for (var i = 0; i < rows.length; i++) {
    var row = INPUT_TABLE_START + i;
    inputSh.getRange(row, WI_COL.CONS).setFormula(
      '=IF(AND(ISNUMBER(D' + row + '),ISNUMBER(E' + row + ')),MAX(0,D' + row + '-E' + row + '),"")');
  }

  toast_('Reading input table refreshed (' + rows.length + ' units).', 'Setup');
}
