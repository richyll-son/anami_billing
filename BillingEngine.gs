// ============================================================
// AHNHAI Billing System — BillingEngine.gs
// Process & Generate Bills: validate → compute rate →
// check penalties → write ledgers → refresh downstream sheets
// ============================================================

function processAndGenerateBills() {
  var inSh = getSheet_(SH.W_INPUT);
  if (!inSh) { alert_('Water Reading Input sheet not found.'); return; }

  // ── 1. Read form header inputs ───────────────────────────
  var year      = parseInt(inSh.getRange('B3').getValue(), 10);
  var monthName = String(inSh.getRange('B4').getValue()).trim();
  var month     = getMonthNum(monthName);

  if (!year || month === 0) {
    alert_('Please set Year and Month before generating bills.'); return;
  }

  var mcwdFrom = inSh.getRange('B5').getValue();
  var mcwdTo   = inSh.getRange('B6').getValue();
  var mcwdAmt  = toNum(inSh.getRange('B7').getValue());
  var elecFrom = inSh.getRange('B8').getValue();
  var elecTo   = inSh.getRange('B9').getValue();
  var elecAmt  = toNum(inSh.getRange('B10').getValue());
  var manpower = toNum(inSh.getRange('B11').getValue()) || MANPOWER_DEF;

  if (mcwdAmt === 0 && elecAmt === 0) {
    if (!confirm_('Warning', 'MCWD and Electricity amounts are both 0. Continue?')) return;
  }

  // ── 2. Read table data ───────────────────────────────────
  var lastRow = inSh.getLastRow();
  if (lastRow < INPUT_TABLE_START) {
    alert_('No reading data found. Run "Refresh Reading Input Table" from the menu first,\n' +
           'then enter Current Readings in column D.'); return;
  }

  var tableVals = inSh.getRange(INPUT_TABLE_START, 1,
    lastRow - INPUT_TABLE_START + 1, 6).getValues();

  var readings = [];
  tableVals.forEach(function(r) {
    var uid = String(r[WI_COL.UNIT - 1]).trim();
    if (!uid) return;
    var cur  = parseFloat(r[WI_COL.CUR  - 1]);
    if (isNaN(cur)) return;
    var prev = parseFloat(r[WI_COL.PREV - 1]);
    var cons = isNaN(prev) ? cur : Math.max(0, cur - prev);
    readings.push({
      unitId   : uid,
      meterNo  : r[WI_COL.METER - 1],
      ownerStr : r[WI_COL.OWNER - 1],
      prev     : isNaN(prev) ? 0 : prev,
      cur      : cur,
      cons     : cons
    });
  });

  if (readings.length === 0) {
    alert_('No valid readings found.\n' +
           'Make sure Unit IDs are in column A and Current Readings are in column D.'); return;
  }

  // ── 3. Check for duplicate billing period ────────────────
  var wsSh = getSheet_(SH.W_STORE);
  if (wsSh && wsSh.getLastRow() > 1) {
    var wsCheck = wsSh.getRange(2, 2, wsSh.getLastRow() - 1, 2).getValues();
    var dup = wsCheck.some(function(r) {
      return parseInt(r[0], 10) === year && r[1] === monthName;
    });
    if (dup && !confirm_('Duplicate Warning',
      'Bills for ' + monthName + ' ' + year + ' have already been generated.\nRegenerate?')) return;
  }

  // ── 4. Compute rate ──────────────────────────────────────
  var totalCons    = readings.reduce(function(s, r) { return s + r.cons; }, 0);
  if (totalCons === 0) {
    alert_('Total consumption is 0. Cannot compute rate.'); return;
  }
  var totalExpense = mcwdAmt + elecAmt + manpower;
  var rate         = totalExpense / totalCons;
  var billDt       = getBillDate(year, month);
  var dueDt        = getDueDate(year, month);
  var now          = new Date();

  toast_('Computing penalties…', 'Processing');

  // ── 5. Load existing water ledger (for penalty check) ───
  var wlSh   = getSheet_(SH._WL);
  var wlData = (wlSh && wlSh.getLastRow() > 1)
    ? wlSh.getRange(2, 1, wlSh.getLastRow() - 1, WL_COLS).getValues()
    : [];

  // Index: last balance and due date per unit
  var unitLastState = {};  // unitId → {balance, dueDate, addon}
  wlData.forEach(function(r) {
    var uid = r[WL.UNIT];
    if (!uid) return;
    unitLastState[uid] = {
      balance : toNum(r[WL.BALANCE]),
      dueDate : r[WL.DUE_DATE],
      addon   : toNum(r[WL.ADDON])
    };
  });

  // ── 6. Build rows for _WaterLedger, _DuesLedger, Water Store
  var mlSh   = getSheet_(SH.MASTERLIST);
  var mlData = (mlSh && mlSh.getLastRow() > 1)
    ? mlSh.getDataRange().getValues() : [[]];

  var presentDateStr = fmtDate(now);
  var billDateStr    = fmtDate(billDt);
  var dueDateStr     = fmtDate(dueDt);

  var newWlRows  = [];
  var newDlRows  = [];
  var newWsRows  = [];

  readings.forEach(function(r) {
    var state   = unitLastState[r.unitId] || { balance: 0, dueDate: '', addon: 0 };
    var prevBal = state.balance;

    // Penalty: 5% of previous balance if due date has passed
    var penalty = 0;
    if (prevBal > 0) {
      var prevDue = parseDate(state.dueDate);
      if (prevDue && now > prevDue) {
        penalty = Math.round(prevBal * PENALTY_RATE * 100) / 100;
      }
    }

    var rawBill  = r.cons * rate;
    var debit    = Math.max(MIN_WATER_BILL, rawBill);
    debit        = Math.round(debit * 100) / 100;
    var addon    = state.addon;          // carry forward per-unit addon

    var parsed   = parseUID(r.unitId);
    var billNo   = parsed
      ? buildBillNum(year, month, parsed.phase, parsed.block, parsed.lot) : '';

    // _WaterLedger row — BALANCE will be recalculated in batch below
    newWlRows.push([
      r.unitId,        // UNIT ID
      year,            // YEAR
      monthName,       // MONTH
      billDateStr,     // BILL DATE
      '',              // PREV READING DATE  (set from last Water Store entry, or blank on first)
      presentDateStr,  // PRESENT READING DATE
      r.prev,          // PREV READING
      r.cur,           // PRESENT READING
      rate,            // RATE/CUBIC
      dueDateStr,      // DUE DATE
      penalty,         // PENALTY
      debit,           // DEBIT
      0,               // CREDIT (not yet paid)
      0,               // BALANCE (recalculated below)
      addon,           // ADDON MCWD
      '',              // OR NUMBER
      '',              // REMARKS
      billNo,          // BILL NUMBER
      ''               // PAYMENT DATE
    ]);

    // _DuesLedger row
    newDlRows.push([
      r.unitId, year, monthName,
      '',        // PAYMENT DATE
      ASSOC_DUES,
      0,         // CREDIT
      0,         // BALANCE (recalculated below)
      '',        // OR NUMBER
      ''         // REMARKS
    ]);

    // Water Reading Data Store row
    newWsRows.push([
      presentDateStr, year, monthName,
      r.unitId, r.meterNo,
      r.prev, r.cur, r.cons,
      rate, debit
    ]);
  });

  toast_('Writing ledgers…', 'Processing');

  // ── 7. Append rows to hidden ledgers ─────────────────────
  _appendRows(SH._WL, newWlRows);
  _appendRows(SH._DL, newDlRows);

  // ── 8. Batch recalculate all balances ────────────────────
  recalcWaterBalances();
  recalcDuesBalances();

  // ── 9. Append to Water Store ─────────────────────────────
  if (newWsRows.length > 0 && wsSh) {
    wsSh.getRange(wsSh.getLastRow() + 1, 1, newWsRows.length, newWsRows[0].length)
      .setValues(newWsRows);
  }

  // ── 10. Append to Rate Calculator ─────────────────────────
  _appendRateCalcRow(now, year, monthName, mcwdFrom, mcwdTo, mcwdAmt,
    elecFrom, elecTo, elecAmt, manpower, totalExpense, totalCons, rate);

  toast_('Refreshing summary & bill prints…', 'Processing');

  // ── 11. Refresh downstream sheets ────────────────────────
  refreshMonthlySummary();
  regenerateBillPrint(year, month, monthName, rate, totalCons,
    mcwdAmt, elecAmt, manpower, mcwdFrom, mcwdTo, elecFrom, elecTo);

  // ── 12. Refresh input table with new "previous" readings ─
  populateWaterInputTable();

  alert_('Bills generated successfully!\n\n' +
    'Period:             ' + monthName + ' ' + year + '\n' +
    'Units billed:       ' + readings.length + '\n' +
    'Rate / m³:          ₱' + fmt2(rate) + '\n' +
    'Total Expense:      ₱' + fmt2(totalExpense) + '\n' +
    'Total Consumption:  ' + totalCons + ' m³\n\n' +
    'Bill Date: ' + billDateStr + '   Due Date: ' + dueDateStr);
}

// ── Helpers ──────────────────────────────────────────────────

function _appendRows(sheetName, rows) {
  if (rows.length === 0) return;
  var sh = getSheet_(sheetName);
  if (!sh) return;
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function _appendRateCalcRow(now, year, monthName,
    mcwdFrom, mcwdTo, mcwdAmt, elecFrom, elecTo, elecAmt,
    manpower, totalExpense, totalCons, rate) {
  var sh = getSheet_(SH.RATE_CALC);
  if (!sh) return;
  sh.getRange(sh.getLastRow() + 1, 1, 1, 13).setValues([[
    fmtDate(now), year, monthName,
    mcwdFrom ? fmtDate(mcwdFrom) : '',
    mcwdTo   ? fmtDate(mcwdTo)   : '',
    mcwdAmt,
    elecFrom ? fmtDate(elecFrom) : '',
    elecTo   ? fmtDate(elecTo)   : '',
    elecAmt,
    manpower, totalExpense, totalCons, rate
  ]]);
}

// ── Balance recalculation (called after any debit/credit change) ──

function recalcWaterBalances() {
  var sh = getSheet_(SH._WL);
  if (!sh || sh.getLastRow() < 2) return;

  var numRows = sh.getLastRow() - 1;
  // Read only needed columns: UNIT(A=1), PENALTY(K=11), DEBIT(L=12), CREDIT(M=13), ADDON(O=15)
  var data = sh.getRange(2, 1, numRows, WL_COLS).getValues();

  var unitBal = {};  // running balance per unit
  var balCol  = new Array(numRows);

  for (var i = 0; i < data.length; i++) {
    var uid    = data[i][WL.UNIT];
    if (!uid) { balCol[i] = 0; continue; }
    var prev   = unitBal[uid] || 0;
    var debit  = toNum(data[i][WL.DEBIT]);
    var pen    = toNum(data[i][WL.PENALTY]);
    var credit = toNum(data[i][WL.CREDIT]);
    var addon  = toNum(data[i][WL.ADDON]);
    var bal    = prev + debit + pen + addon - credit;
    if (bal < 0) bal = 0;
    bal = Math.round(bal * 100) / 100;
    unitBal[uid] = bal;
    balCol[i] = bal;
  }

  // Write all balances in one batch (column N = index 14 = col 14)
  sh.getRange(2, WL.BALANCE + 1, numRows, 1)
    .setValues(balCol.map(function(v) { return [v]; }));
}

function recalcDuesBalances() {
  var sh = getSheet_(SH._DL);
  if (!sh || sh.getLastRow() < 2) return;

  var numRows = sh.getLastRow() - 1;
  var data = sh.getRange(2, 1, numRows, DL_COLS).getValues();

  var unitBal = {};
  var balCol  = new Array(numRows);

  for (var i = 0; i < data.length; i++) {
    var uid    = data[i][DL.UNIT];
    if (!uid) { balCol[i] = 0; continue; }
    var prev   = unitBal[uid] || 0;
    var debit  = toNum(data[i][DL.DEBIT]);
    var credit = toNum(data[i][DL.CREDIT]);
    var bal    = prev + debit - credit;
    if (bal < 0) bal = 0;
    bal = Math.round(bal * 100) / 100;
    unitBal[uid] = bal;
    balCol[i] = bal;
  }

  sh.getRange(2, DL.BALANCE + 1, numRows, 1)
    .setValues(balCol.map(function(v) { return [v]; }));
}
