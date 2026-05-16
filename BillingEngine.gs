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

  // Add-on MCWD cutoff: read from B12 (month name) and C12 (year)
  var addonUntilMonth = String(inSh.getRange('B12').getValue()).trim();
  var addonUntilYear  = parseInt(inSh.getRange('C12').getValue(), 10);
  var addonCutoffRank = (addonUntilMonth && addonUntilYear)
    ? (addonUntilYear * 100 + getMonthNum(addonUntilMonth)) : 0;
  var billingRank = year * 100 + month;
  var applyAddon  = (addonCutoffRank === 0) || (billingRank <= addonCutoffRank);

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
    lastRow - INPUT_TABLE_START + 1, 7).getValues();

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
      cons     : cons,
      addon    : toNum(r[WI_COL.ADDON - 1])
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

  // ── 5b. Load pending meter changes for this billing period ─
  var mcSh   = getSheet_(SH._METER_CHG);
  var mcData = (mcSh && mcSh.getLastRow() > 1)
    ? mcSh.getRange(2, 1, mcSh.getLastRow() - 1, MC_COLS).getValues() : [];

  // Build map: unitId → sorted array of {rowIdx(1-based), dateChanged, oldFinal, newMeter, newStart}
  var meterChangeMap = {};
  mcData.forEach(function(mc, i) {
    var uid    = String(mc[MC.UNIT]   || '').trim();
    var mcYr   = parseInt(mc[MC.YEAR], 10);
    var mcMo   = String(mc[MC.MONTH]  || '').trim();
    var status = String(mc[MC.STATUS] || '').trim().toLowerCase();
    if (!uid || mcYr !== year || mcMo !== monthName || status !== 'pending') return;
    if (!meterChangeMap[uid]) meterChangeMap[uid] = [];
    var rawDate = mc[MC.DATE];
    meterChangeMap[uid].push({
      rowIdx  : i + 2,
      date    : (rawDate instanceof Date) ? rawDate : new Date(String(rawDate)),
      oldFinal: toNum(mc[MC.OLD_FINAL]),
      newMeter: String(mc[MC.NEW_METER] || '').trim(),
      newStart: toNum(mc[MC.NEW_START])
    });
  });
  // Sort each unit's changes by date ascending
  Object.keys(meterChangeMap).forEach(function(uid) {
    meterChangeMap[uid].sort(function(a, b) { return a.date - b.date; });
  });

  // ── 6. Build rows for _WaterLedger, _DuesLedger, Water Store
  var mlSh   = getSheet_(SH.MASTERLIST);
  var mlData = (mlSh && mlSh.getLastRow() > 1)
    ? mlSh.getDataRange().getValues() : [[]];

  // Coverage From/To: actual meter reading dates entered by admin (B13, B14)
  var coverageFromRaw = inSh.getRange('B13').getValue();
  var coverageToRaw   = inSh.getRange('B14').getValue();
  var coverageFromStr = coverageFromRaw ? fmtDate(coverageFromRaw) : '';
  var coverageToStr   = coverageToRaw   ? fmtDate(coverageToRaw)   : fmtDate(now);

  var presentDateStr = fmtDate(now);
  var billDateStr    = fmtDate(billDt);
  var dueDateStr     = fmtDate(dueDt);

  var newWlRows  = [];
  var newDlRows  = [];
  var newWsRows  = [];

  // Pre-load dues rates and BOD data once to avoid per-unit sheet reads
  var duesRatesSh   = getSheet_(SH._DUES_RATES);
  var duesRatesData = (duesRatesSh && duesRatesSh.getLastRow() > 1)
    ? duesRatesSh.getRange(2, 1, duesRatesSh.getLastRow() - 1, DR_COLS).getValues() : [];
  var bodSh   = getSheet_(SH._BOD);
  var bodData = (bodSh && bodSh.getLastRow() > 1)
    ? bodSh.getRange(2, 1, bodSh.getLastRow() - 1, BOD_COLS).getValues() : [];

  readings.forEach(function(r) {
    var state   = unitLastState[r.unitId] || { balance: 0, dueDate: '', addon: 0 };
    var prevBal = state.balance;

    // Penalty: 5% of previous balance (positive = owes), starting 1 day after due date
    var penalty = 0;
    if (prevBal > 0) {
      var prevDue = parseDate(state.dueDate);
      if (prevDue) {
        var penaltyDate = new Date(prevDue.getTime());
        penaltyDate.setDate(penaltyDate.getDate() + 1);
        if (now >= penaltyDate) {
          penalty = Math.round(prevBal * PENALTY_RATE * 100) / 100;
        }
      }
    }

    // Apply meter change chain if pending changes exist for this unit
    var effCons    = r.cons;
    var effMeter   = r.meterNo;
    var changeWarn = false;
    var changes    = meterChangeMap[r.unitId];
    if (changes && changes.length > 0) {
      var workingPrev = r.prev;
      var chainCons   = 0;
      changes.forEach(function(ch) {
        if (ch.oldFinal < workingPrev) {
          Logger.log('WARNING: Meter change for ' + r.unitId + ': oldFinal (' + ch.oldFinal +
                     ') < workingPrev (' + workingPrev + '). Consumption floored at 0.');
          changeWarn = true;
        }
        chainCons   += Math.max(0, ch.oldFinal - workingPrev);
        workingPrev  = ch.newStart;
      });
      chainCons += Math.max(0, r.cur - workingPrev);
      effCons   = Math.round(chainCons * 100) / 100;
      effMeter  = changes[changes.length - 1].newMeter;
    }

    var rawBill     = effCons * rate;
    var debit       = effCons === 0 ? 0 : Math.max(MIN_WATER_BILL, rawBill);
    debit           = Math.round(debit * 100) / 100;
    var isMinCharge = effCons > 0 && rawBill < MIN_WATER_BILL;
    var addon       = applyAddon ? toNum(r.addon) : 0;

    var parsed   = parseUID(r.unitId);
    var billNo   = parsed
      ? buildBillNum(year, month, parsed.phase, parsed.block, parsed.lot) : '';

    var baseRemark = changeWarn ? 'Meter changed x' + changes.length + ' [CHECK: oldFinal < prev]'
                   : (changes && changes.length > 0 ? 'Meter changed x' + changes.length : '');
    var remark = isMinCharge
      ? (baseRemark ? baseRemark + '; Min. charge applied' : 'Min. charge applied')
      : baseRemark;

    // _WaterLedger row — BALANCE will be recalculated in batch below
    newWlRows.push([
      r.unitId,        // UNIT ID
      year,            // YEAR
      monthName,       // MONTH
      billDateStr,     // BILL DATE
      coverageFromStr, // PREV READING DATE
      coverageToStr,   // PRESENT READING DATE
      r.prev,          // PREV READING (original — for reference)
      r.cur,           // PRESENT READING (new meter current)
      rate,            // RATE/CUBIC
      dueDateStr,      // DUE DATE
      penalty,         // PENALTY
      debit,           // DEBIT
      0,               // CREDIT (not yet paid)
      0,               // BALANCE (recalculated below)
      addon,           // ADDON MCWD
      '',              // OR NUMBER
      remark,          // REMARKS
      billNo,          // BILL NUMBER
      ''               // PAYMENT DATE
    ]);

    // _DuesLedger row
    var duesDebit = isBODExempt_(r.unitId, year, month, bodData) ? 0 : getDuesRate_(year, month, duesRatesData);
    newDlRows.push([
      r.unitId, year, monthName,
      '',        // PAYMENT DATE
      duesDebit,
      0,         // CREDIT
      0,         // BALANCE (recalculated below)
      '',        // OR NUMBER
      ''         // REMARKS
    ]);

    // Water Reading Data Store row (use effCons and effMeter for meter-change units)
    newWsRows.push([
      presentDateStr, year, monthName,
      r.unitId, effMeter,
      r.prev, r.cur, effCons,
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

  // ── 12. Mark processed meter changes + update Masterlist ─
  var mcChangedUnits = Object.keys(meterChangeMap);
  if (mcChangedUnits.length > 0 && mcSh) {
    // Batch-update STATUS column in _MeterChanges (one write instead of N)
    var statusVals = mcSh.getRange(2, MC.STATUS + 1, mcSh.getLastRow() - 1, 1).getValues();
    mcChangedUnits.forEach(function(uid) {
      meterChangeMap[uid].forEach(function(ch) {
        statusVals[ch.rowIdx - 2][0] = 'Processed';
      });
    });
    mcSh.getRange(2, MC.STATUS + 1, statusVals.length, 1).setValues(statusVals);

    // Batch-update Masterlist meter number column (col J = 10)
    if (mlSh && mlSh.getLastRow() > 1) {
      var mlMeters = mlSh.getRange(2, 10, mlSh.getLastRow() - 1, 1).getValues();
      mcChangedUnits.forEach(function(uid) {
        var lastNewMeter = meterChangeMap[uid][meterChangeMap[uid].length - 1].newMeter;
        for (var mi = 1; mi < mlData.length; mi++) {
          if (String(mlData[mi][0] || '').trim() === uid) {
            mlMeters[mi - 1][0] = lastNewMeter;
            break;
          }
        }
      });
      mlSh.getRange(2, 10, mlMeters.length, 1).setValues(mlMeters);
    }
  }

  // ── 13. Refresh input table with new "previous" readings ─
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

  var unitBal  = {};  // running balance per unit
  var unitSeen = {};  // first row per unit: preserve stored opening balance
  var balCol   = new Array(numRows);

  for (var i = 0; i < data.length; i++) {
    var uid    = data[i][WL.UNIT];
    if (!uid) { balCol[i] = 0; continue; }

    if (!unitSeen[uid]) {
      // First row: use stored balance as the opening balance (may be manually set)
      unitSeen[uid] = true;
      var opening = toNum(data[i][WL.BALANCE]);
      unitBal[uid] = opening;
      balCol[i]    = opening;
      continue;
    }

    var prev   = unitBal[uid] || 0;
    var debit  = toNum(data[i][WL.DEBIT]);
    var pen    = toNum(data[i][WL.PENALTY]);
    var credit = toNum(data[i][WL.CREDIT]);
    var addon  = toNum(data[i][WL.ADDON]);
    var bal    = prev + debit + pen + addon - credit;
    bal = Math.round(bal * 100) / 100;
    // Negative balance = credit carry-forward (unit overpaid), intentional
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
    var bal    = Math.round((prev + debit - credit) * 100) / 100;
    unitBal[uid] = bal;
    balCol[i] = bal;
  }

  sh.getRange(2, DL.BALANCE + 1, numRows, 1)
    .setValues(balCol.map(function(v) { return [v]; }));
}

// Returns the association dues amount for a given year+month.
// Looks up _DuesRates sheet; most-recently-added matching row wins.
// Falls back to most recent row amount, then ASSOC_DUES constant.
// Pass preloadedData to avoid a sheet read when calling inside a loop.
function getDuesRate_(year, month, preloadedData) {
  var rows = preloadedData;
  if (!rows) {
    var sh = getSheet_(SH._DUES_RATES);
    if (!sh || sh.getLastRow() < 2) return ASSOC_DUES;
    rows = sh.getRange(2, 1, sh.getLastRow() - 1, DR_COLS).getValues();
  }
  if (!rows.length) return ASSOC_DUES;

  var rank = year * 100 + month;
  var match = null;

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var fromRank = toNum(r[DR.FROM_YEAR]) * 100 + getMonthNum(String(r[DR.FROM_MONTH]));
    var toMonth  = String(r[DR.TO_MONTH] || '').trim();
    var toYear   = toNum(r[DR.TO_YEAR]);
    var toRank   = (toMonth && toYear) ? (toYear * 100 + getMonthNum(toMonth)) : 999999;

    if (rank >= fromRank && rank <= toRank) {
      match = toNum(r[DR.AMOUNT]);
    }
  }

  if (match !== null) return match;

  // Fall back to most recent entry amount
  for (var j = rows.length - 1; j >= 0; j--) {
    var amt = toNum(rows[j][DR.AMOUNT]);
    if (amt > 0) return amt;
  }

  return ASSOC_DUES;
}

// Returns true if the given unitId has an active BOD exemption for year+month.
// Pass preloadedData to avoid a sheet read when calling inside a loop.
function isBODExempt_(unitId, year, month, preloadedData) {
  var rows = preloadedData;
  if (!rows) {
    var sh = getSheet_(SH._BOD);
    if (!sh || sh.getLastRow() < 2) return false;
    rows = sh.getRange(2, 1, sh.getLastRow() - 1, BOD_COLS).getValues();
  }
  if (!rows.length) return false;

  var rank = year * 100 + month;
  var uid  = String(unitId || '').trim().toUpperCase();

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var rowUnit = String(r[BOD.UNIT] || '').trim().toUpperCase();
    if (rowUnit !== uid) continue;

    var fromRank = toNum(r[BOD.FROM_YEAR]) * 100 + getMonthNum(String(r[BOD.FROM_MONTH]));
    var toMonth  = String(r[BOD.TO_MONTH] || '').trim();
    var toYear   = toNum(r[BOD.TO_YEAR]);
    var toRank   = (toMonth && toYear) ? (toYear * 100 + getMonthNum(toMonth)) : 999999;

    if (rank >= fromRank && rank <= toRank) return true;
  }

  return false;
}

// ── Recalculate Penalties for all existing _WaterLedger rows ──
// Zeros out any penalty on a row where the previous unclamped running balance
// was ≤ 0 (unit was in credit/overpayment), then rebuilds all balances.
function recalcPenalties() {
  if (!confirm_('Recalculate Penalties',
      'This will scan all water ledger rows and zero out penalties on rows where\n' +
      'the unit had a credit balance before that billing period.\n\n' +
      'Proceed?')) return;

  var sh = getSheet_(SH._WL);
  if (!sh || sh.getLastRow() < 2) {
    alert_('Water Ledger is empty.'); return;
  }

  var numRows = sh.getLastRow() - 1;
  var data    = sh.getRange(2, 1, numRows, WL_COLS).getValues();

  // Track UNCLAMPED running balance per unit so we know true credit state
  var unitBal   = {};   // uid → unclamped balance before this row
  var penCol    = new Array(numRows);
  var zeroed    = 0;

  for (var i = 0; i < data.length; i++) {
    var uid  = String(data[i][WL.UNIT] || '').trim();
    if (!uid) { penCol[i] = null; continue; }

    var prevUnclamped = unitBal[uid] || 0;
    var debit  = toNum(data[i][WL.DEBIT]);
    var pen    = toNum(data[i][WL.PENALTY]);
    var credit = toNum(data[i][WL.CREDIT]);
    var addon  = toNum(data[i][WL.ADDON]);

    // If the unit had no debt (<= 0 = paid/credit), this row's penalty should be 0
    if (prevUnclamped <= 0 && pen > 0) {
      data[i][WL.PENALTY] = 0;
      pen = 0;
      penCol[i] = 0;
      zeroed++;
    } else {
      penCol[i] = null; // no change
    }

    // Advance running balance: positive = owes, negative = credit
    unitBal[uid] = prevUnclamped + debit + pen + addon - credit;
  }

  if (zeroed > 0) {
    var penVals = data.map(function(row) { return [toNum(row[WL.PENALTY])]; });
    sh.getRange(2, WL.PENALTY + 1, numRows, 1).setValues(penVals);

    // Rebuild all balances
    recalcWaterBalances();
  }

  alert_('Penalty recalculation complete.\n\nRows corrected: ' + zeroed +
         '\nAll water balances have been recalculated.');
}

// ── Full recompute: penalties + balances ─────────────────────
// Formula: balance = prev + debit + penalty + addon − credit
// Positive balance = unit owes money; negative = credit carry-forward.
// First row per unit: stored balance preserved as opening balance.
// Penalty: preserved at 0 if manually waived; otherwise recomputed as
//   5% × prevBal when prevBal > 0 AND previous row's due date < today.
function recomputeAllBalances() {
  if (!confirm_('Recompute All Balances',
      'This will rewrite PENALTY and BALANCE for all water ledger rows,\n' +
      'and BALANCE for all dues ledger rows.\n\n' +
      '  balance = prev + debit + penalty + addon − credit\n\n' +
      'First row per unit keeps its stored opening balance.\n' +
      'All penalties are recomputed (5% when previous balance > 0 and due date passed).\n\n' +
      'Proceed?')) return;

  toast_('Recomputing water ledger…', 'Recompute');

  var wlSh = getSheet_(SH._WL);
  if (wlSh && wlSh.getLastRow() > 1) {
    var numRows = wlSh.getLastRow() - 1;
    var data    = wlSh.getRange(2, 1, numRows, WL_COLS).getValues();
    var today   = new Date();

    var unitBal     = {};   // uid → running balance
    var unitSeen    = {};   // first row per unit: preserve stored opening balance
    var unitDueDate = {};   // uid → due date of previous row
    var penCol      = new Array(numRows);
    var balCol      = new Array(numRows);

    for (var i = 0; i < data.length; i++) {
      var uid    = String(data[i][WL.UNIT] || '').trim();
      if (!uid) { penCol[i] = null; balCol[i] = 0; continue; }

      if (!unitSeen[uid]) {
        // First row: preserve stored opening balance, skip penalty recompute
        unitSeen[uid] = true;
        var opening = toNum(data[i][WL.BALANCE]);
        unitBal[uid] = opening;
        balCol[i]    = opening;
        penCol[i]    = null;  // no change to first row's penalty
        var dd0 = data[i][WL.DUE_DATE];
        if (dd0) unitDueDate[uid] = dd0;
        continue;
      }

      var prevBal  = unitBal[uid] || 0;
      var debit    = toNum(data[i][WL.DEBIT]);
      var credit   = toNum(data[i][WL.CREDIT]);
      var addon    = toNum(data[i][WL.ADDON]);

      // Always recompute: apply penalty only if prevBal > 0 and previous due date passed
      var newPen = 0;
      if (prevBal > 0) {
        var prevDue = parseDate(unitDueDate[uid]);
        if (prevDue) {
          var cutoff = new Date(prevDue.getTime());
          cutoff.setDate(cutoff.getDate() + 1);
          if (today >= cutoff) {
            newPen = Math.round(prevBal * PENALTY_RATE * 100) / 100;
          }
        }
      }
      var pen = newPen;
      penCol[i] = pen;
      data[i][WL.PENALTY] = pen;

      var bal = Math.round((prevBal + debit + pen + addon - credit) * 100) / 100;
      unitBal[uid] = bal;
      balCol[i] = bal;

      // Track this row's due date for the next row of this unit
      var dd = data[i][WL.DUE_DATE];
      if (dd) unitDueDate[uid] = dd;
    }

    // Write penalty column (only changed rows)
    var penRange = wlSh.getRange(2, WL.PENALTY + 1, numRows, 1);
    var penVals  = penRange.getValues();
    for (var k = 0; k < numRows; k++) {
      if (penCol[k] !== null) penVals[k][0] = penCol[k];
    }
    penRange.setValues(penVals);

    // Write balance column
    wlSh.getRange(2, WL.BALANCE + 1, numRows, 1)
      .setValues(balCol.map(function(v) { return [v]; }));
  }

  toast_('Recomputing dues ledger…', 'Recompute');
  recalcDuesBalances();

  alert_('Recompute complete.\nWater: penalties and balances updated.\nDues: balances updated.');
}

// ── Unit tests: minimum water bill rule ──────────────────────
// Run from the Apps Script editor (▶) or via the menu.
// Results are logged to the Apps Script console (View → Logs).
function testBillingMin() {
  var PASS = 'PASS';
  var FAIL = 'FAIL';
  var results = [];
  var allPass = true;

  function check(label, actual, expected) {
    var ok = Math.abs(actual - expected) < 0.001;
    if (!ok) allPass = false;
    results.push('[' + (ok ? PASS : FAIL) + ']  ' + label +
      '\n       expected: ' + expected + '  got: ' + actual);
    return ok;
  }

  function calcBill(cons, rate) {
    var raw = cons * rate;
    return cons === 0 ? 0 : Math.max(MIN_WATER_BILL, raw);
  }

  // 1. Zero consumption → ₱0 (minimum does NOT apply)
  check('0 m³ × ₱50 → ₱0  (min does not apply at 0 cons)', calcBill(0, 50), 0);

  // 2. Low consumption whose amount is below ₱250 → ₱250
  check('2 m³ × ₱50 = ₱100  → min applies → ₱250', calcBill(2, 50), 250);

  // 3. Consumption result exactly at ₱250 → ₱250
  check('5 m³ × ₱50 = ₱250  → exactly at minimum → ₱250', calcBill(5, 50), 250);

  // 4. Normal consumption above ₱250 → actual amount, no rounding to min
  check('10 m³ × ₱50 = ₱500 → above min → ₱500', calcBill(10, 50), 500);

  // 5. Just below ₱250 → minimum applies
  check('4 m³ × ₱60 = ₱240  → min applies → ₱250', calcBill(4, 60), 250);

  // 6. Just above ₱250 → actual amount
  check('6 m³ × ₱50 = ₱300  → above min → ₱300', calcBill(6, 50), 300);

  // 7. Constant sanity: MIN_WATER_BILL = 250
  check('MIN_WATER_BILL constant = 250', MIN_WATER_BILL, 250);

  // 8. isMinCharge flag: true when cons > 0 and raw < 250
  var raw8 = 2 * 50;
  var isMin8 = 2 > 0 && raw8 < MIN_WATER_BILL;
  check('isMinCharge = true when raw (100) < MIN (250)', isMin8 ? 1 : 0, 1);

  // 9. isMinCharge flag: false when cons = 0
  var isMin9 = 0 > 0 && (0 * 50) < MIN_WATER_BILL;
  check('isMinCharge = false when cons = 0', isMin9 ? 1 : 0, 0);

  // 10. isMinCharge flag: false when raw >= 250
  var raw10 = 10 * 50;  // 500
  var isMin10 = 10 > 0 && raw10 < MIN_WATER_BILL;
  check('isMinCharge = false when raw (500) >= MIN (250)', isMin10 ? 1 : 0, 0);

  var passed = results.filter(function(r) { return r.indexOf('[' + PASS + ']') === 0; }).length;
  var total  = results.length;

  Logger.log('\n=== testBillingMin ===');
  results.forEach(function(r) { Logger.log(r); });
  Logger.log('=== ' + passed + '/' + total + ' passed' + (allPass ? ' — ALL PASS' : ' — FAILURES ABOVE') + ' ===\n');

  if (allPass) {
    alert_('testBillingMin: All ' + total + ' tests passed!');
  } else {
    alert_('testBillingMin: ' + (total - passed) + ' test(s) FAILED.\nSee Apps Script Logs for details.');
  }
}
