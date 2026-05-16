// ============================================================
// AHNHAI Billing System — WebApp.gs
// doGet entry point + all webapp_* server-side functions
// called from the browser via google.script.run
// ============================================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('AHNHAI Billing System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Used by HtmlService templates: <?!= include('Stylesheet') ?>
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ── Dashboard analytics helpers ──────────────────────────────
function webapp_isCommonAccount_(unitId) {
  var v = String(unitId || '').trim().toUpperCase().replace(/\s+/g, '');
  return ['GUARDHOUSE', 'CLUBHOUSE', 'CHAPEL'].indexOf(v) !== -1;
}

function webapp_isRegularUnit_(unitId) {
  return /^P\d+B\d+L[\d&]+$/i.test(String(unitId || '').trim());
}

function webapp_getGroup_(unitId) {
  unitId = String(unitId || '').trim().toUpperCase().replace(/\s+/g, '');
  if (webapp_isCommonAccount_(unitId)) return { phase: 'Common', block: 'Common Accounts', common: true };
  var m = unitId.match(/^P(\d+)B(\d+)L/i);
  if (!m) return { phase: 'Other', block: 'Other', common: false };
  return {
    phase: 'Phase ' + m[1],
    block: 'Block ' + m[2],
    common: false
  };
}

function webapp_buildCommonBillNum_(year, monthNum, unitId) {
  var cleanId = String(unitId || '').trim().toUpperCase().replace(/\s+/g, '-');
  return 'WB-' + year + '-' + String(monthNum).padStart(2, '0') + '-' + cleanId;
}

function webapp_emptyTotals_() {
  return {
    payments: 0,
    receivables: 0,
    overpayments: 0,
    waterPayments: 0,
    waterReceivables: 0,
    waterOverpayments: 0,
    duesPayments: 0,
    duesReceivables: 0,
    duesOverpayments: 0
  };
}

function webapp_addToTotals_(target, source) {
  Object.keys(source).forEach(function(k) {
    target[k] = (target[k] || 0) + (toNum(source[k]) || 0);
  });
}

function webapp_roundTotals_(obj) {
  Object.keys(obj).forEach(function(k) {
    if (typeof obj[k] === 'number') obj[k] = Math.round(obj[k] * 100) / 100;
  });
  return obj;
}

function webapp_periodKey_(year, month) {
  var y = parseInt(year, 10);
  var m = String(month || '').trim();
  var n = getMonthNum(m);
  if (!y || !n) return m || 'Unspecified';
  return y + '-' + String(n).padStart(2, '0');
}

function webapp_periodLabel_(year, month) {
  var y = parseInt(year, 10);
  var m = String(month || '').trim();
  if (!y || !m) return m || 'Unspecified';
  return m.substring(0, 3) + ' ' + y;
}

function webapp_getDashboardSummary() {
  var mlSh = getSheet_(SH.MASTERLIST);
  var wlSh = getSheet_(SH._WL);
  var dlSh = getSheet_(SH._DL);

  var master = {};
  var unitCount = 0;
  if (mlSh && mlSh.getLastRow() > 1) {
    mlSh.getRange(2, 1, mlSh.getLastRow() - 1, 12).getValues().forEach(function(r) {
      var uid = String(r[0] || '').trim();
      if (!uid || String(r[11] || 'Yes').toLowerCase() === 'no') return;
      master[uid] = {
        id: uid,
        name: ownerName(r),
        phase: r[1],
        block: r[2],
        common: webapp_isCommonAccount_(uid)
      };
      unitCount++;
    });
  }

  var byUnit = {};
  Object.keys(master).forEach(function(uid) {
    byUnit[uid] = webapp_emptyTotals_();
  });

  var monthlyMap = {};

  function ensureUnit(uid) {
    uid = String(uid || '').trim();
    if (!uid) return null;
    if (!byUnit[uid]) byUnit[uid] = webapp_emptyTotals_();
    return byUnit[uid];
  }

  function addMonthly(year, month, field, value) {
    var key = webapp_periodKey_(year, month);
    if (!/^\d{4}-\d{2}$/.test(key)) return;
    if (!monthlyMap[key]) {
      monthlyMap[key] = { key: key, label: webapp_periodLabel_(year, month), payments: 0, receivables: 0, overpayments: 0, waterPayments: 0, duesPayments: 0, waterReceivables: 0, duesReceivables: 0 };
    }
    monthlyMap[key][field] += toNum(value);
  }

  // Find global latest billing month across all WL rows
  var wlLatestRank = 0, wlLatestYear = 0, wlLatestMonth = 0;
  var wlAllRows = (wlSh && wlSh.getLastRow() > 1)
    ? wlSh.getRange(2, 1, wlSh.getLastRow() - 1, WL_COLS).getValues() : [];
  wlAllRows.forEach(function(r) {
    var yr = parseInt(r[WL.YEAR], 10);
    var mo = getMonthNum(String(r[WL.MONTH] || '').trim());
    var rank = yr * 100 + mo;
    if (rank > wlLatestRank) { wlLatestRank = rank; wlLatestYear = yr; wlLatestMonth = mo; }
  });

  // Payments = all-time sum of credits; receivables = balances from global latest billing month only
  var waterLast = {};  // uid → { balance, year, month } for units billed in the global latest month
  wlAllRows.forEach(function(r) {
    var uid = String(r[WL.UNIT] || '').trim();
    var t = ensureUnit(uid); if (!t) return;
    var credit = toNum(r[WL.CREDIT]);
    t.waterPayments += credit;
    t.payments += credit;
    addMonthly(r[WL.YEAR], r[WL.MONTH], 'payments', credit);
    addMonthly(r[WL.YEAR], r[WL.MONTH], 'waterPayments', credit);
    var yr = parseInt(r[WL.YEAR], 10);
    var mo = getMonthNum(String(r[WL.MONTH] || '').trim());
    if (yr === wlLatestYear && mo === wlLatestMonth) {
      waterLast[uid] = { balance: toNum(r[WL.BALANCE]), year: r[WL.YEAR], month: r[WL.MONTH] };
    }
  });

  // Dues balance is cumulative running total; last row per unit = total owed
  var duesLast = {};  // uid → { balance, year, month }
  if (dlSh && dlSh.getLastRow() > 1) {
    dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues().forEach(function(r) {
      var uid = String(r[DL.UNIT] || '').trim();
      var t = ensureUnit(uid); if (!t) return;
      var credit = toNum(r[DL.CREDIT]);
      var bal    = toNum(r[DL.BALANCE]);  // cumulative running balance
      t.duesPayments += credit;
      t.payments += credit;
      duesLast[uid] = { balance: bal, year: r[DL.YEAR], month: r[DL.MONTH] };
      addMonthly(r[DL.YEAR], r[DL.MONTH], 'payments', credit);
      addMonthly(r[DL.YEAR], r[DL.MONTH], 'duesPayments', credit);
    });
  }

  Object.keys(byUnit).forEach(function(uid) {
    var t = byUnit[uid];
    var wLast = waterLast[uid] || null;
    var wBal  = wLast ? wLast.balance : 0;
    var dBal  = duesLast[uid] ? toNum(duesLast[uid].balance) : 0;

    // positive = owes (receivable), negative = credit (overpayment)
    if (wBal > 0) { t.waterReceivables += wBal; t.receivables += wBal; }
    if (wBal < 0) { t.waterOverpayments += Math.abs(wBal); t.overpayments += Math.abs(wBal); }
    if (dBal > 0) { t.duesReceivables += dBal; t.receivables += dBal; }
    if (dBal < 0) { t.duesOverpayments += Math.abs(dBal); t.overpayments += Math.abs(dBal); }

    if (wLast && wBal !== 0) {
      if (wBal > 0) { addMonthly(wLast.year, wLast.month, 'receivables', wBal); addMonthly(wLast.year, wLast.month, 'waterReceivables', wBal); }
      if (wBal < 0) addMonthly(wLast.year, wLast.month, 'overpayments', Math.abs(wBal));
    }
    var dLast = duesLast[uid];
    if (dLast) {
      if (dBal > 0) { addMonthly(dLast.year, dLast.month, 'receivables', dBal); addMonthly(dLast.year, dLast.month, 'duesReceivables', dBal); }
      if (dBal < 0) addMonthly(dLast.year, dLast.month, 'overpayments', Math.abs(dBal));
    }
  });

  var total = webapp_emptyTotals_();
  var phases = { 'Phase 1': webapp_emptyTotals_(), 'Phase 2': webapp_emptyTotals_() };
  var blocks = {};

  Object.keys(byUnit).forEach(function(uid) {
    var group = webapp_getGroup_(uid);
    var t = byUnit[uid];
    webapp_addToTotals_(total, t);

    if (!phases[group.phase]) phases[group.phase] = webapp_emptyTotals_();
    webapp_addToTotals_(phases[group.phase], t);

    var blockKey = group.phase + ' ' + group.block;
    if (!blocks[blockKey]) blocks[blockKey] = webapp_emptyTotals_();
    webapp_addToTotals_(blocks[blockKey], t);
  });

  var phaseRows = Object.keys(phases).sort().map(function(k) {
    var o = webapp_roundTotals_(phases[k]);
    o.label = k;
    return o;
  });

  var blockRows = Object.keys(blocks).sort(function(a, b) {
    function rank(x) {
      var m = x.match(/Phase\s+(\d+)\s+Block\s+(\d+)/i);
      return m ? (parseInt(m[1], 10) * 100 + parseInt(m[2], 10)) : 9999;
    }
    return rank(a) - rank(b);
  }).map(function(k) {
    var o = webapp_roundTotals_(blocks[k]);
    o.label = k;
    return o;
  });

  // Always produce exactly 12 months ending at the current calendar month,
  // with zeros for months that have no ledger data yet.
  var now2 = new Date();
  var curYr = now2.getFullYear();
  var curMo = now2.getMonth() + 1; // 1-based
  var monthlyRows = [];
  for (var mi = 11; mi >= 0; mi--) {
    var mYr = curYr, mMo = curMo - mi;
    if (mMo <= 0) { mMo += 12; mYr--; }
    var mKey  = mYr + '-' + String(mMo).padStart(2, '0');
    var mName = getMonthName(mMo);
    var entry = monthlyMap[mKey] || {
      key: mKey, label: webapp_periodLabel_(mYr, mName),
      payments: 0, receivables: 0, overpayments: 0,
      waterPayments: 0, duesPayments: 0,
      waterReceivables: 0, duesReceivables: 0
    };
    monthlyRows.push(webapp_roundTotals_(entry));
  }

  return {
    unitCount: unitCount,
    total: webapp_roundTotals_(total),
    phases: phaseRows,
    blocks: blockRows,
    monthly: monthlyRows
  };
}

// ── Dues by Year ─────────────────────────────────────────────
function webapp_getDuesYears() {
  var dlSh = getSheet_(SH._DL);
  if (!dlSh || dlSh.getLastRow() < 2) return [];
  var years = {};
  dlSh.getRange(2, DL.YEAR + 1, dlSh.getLastRow() - 1, 1).getValues().forEach(function(r) {
    var y = parseInt(r[0], 10);
    if (y > 2000) years[y] = true;
  });
  return Object.keys(years).map(Number).sort(function(a, b) { return b - a; }); // newest first
}

function webapp_getDuesByYear(year) {
  // year = 0 means "All" — no year filter, aggregate across all years
  var normYear = parseInt(year, 10) || 0;

  var dlSh = getSheet_(SH._DL);
  var empty = { ok: true, year: normYear, totalDebit: 0, totalCredit: 0,
                totalUncollected: 0, overallReceivables: 0, phases: [], blocks: [], monthly: [] };
  if (!dlSh || dlSh.getLastRow() < 2) return empty;

  var totalDebit = 0, totalCredit = 0;
  var phases = {}, blocks = {};
  var monthly = {};
  MONTHS.forEach(function(m) { monthly[m] = { debit: 0, credit: 0 }; });

  // Track last balance per unit for overallReceivables
  var unitLastBal = {};

  dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues().forEach(function(r) {
    var uid = String(r[DL.UNIT] || '').trim();
    if (!uid) return;
    var group = webapp_getGroup_(uid);

    // Always track last balance for overallReceivables (all units, all years)
    if (!group.common && group.phase !== 'Other') {
      unitLastBal[uid] = toNum(r[DL.BALANCE]);
    }

    // Year filter (0 = all years)
    if (normYear !== 0 && toNum(r[DL.YEAR]) !== normYear) return;
    if (group.common || group.phase === 'Other') return;

    var debit  = toNum(r[DL.DEBIT]);
    var credit = toNum(r[DL.CREDIT]);
    var month  = String(r[DL.MONTH] || '').trim();

    totalDebit  += debit;
    totalCredit += credit;

    var ph = group.phase;
    if (!phases[ph]) phases[ph] = { debit: 0, credit: 0 };
    phases[ph].debit  += debit;
    phases[ph].credit += credit;

    var blk = ph + ' ' + group.block;
    if (!blocks[blk]) blocks[blk] = { debit: 0, credit: 0 };
    blocks[blk].debit  += debit;
    blocks[blk].credit += credit;

    if (monthly[month] !== undefined) {
      monthly[month].debit  += debit;
      monthly[month].credit += credit;
    }
  });

  // overallReceivables = sum of positive running balances across all units (all years)
  var overallReceivables = 0;
  Object.keys(unitLastBal).forEach(function(uid) {
    var bal = unitLastBal[uid];
    if (bal > 0) overallReceivables += bal;
  });

  function rnd(v) { return Math.round(v * 100) / 100; }

  var monthlyRows = MONTHS.map(function(m) {
    return {
      label       : m.substring(0, 3),
      debit       : rnd(monthly[m].debit),
      credit      : rnd(monthly[m].credit),
      uncollected : rnd(Math.max(0, monthly[m].debit - monthly[m].credit))
    };
  });

  var phaseRows = Object.keys(phases).sort().map(function(k) {
    return { label: k, debit: rnd(phases[k].debit), credit: rnd(phases[k].credit), uncollected: rnd(Math.max(0, phases[k].debit - phases[k].credit)) };
  });

  var blockRows = Object.keys(blocks).sort(function(a, b) {
    function rank(x) {
      var m = x.match(/Phase\s+(\d+)\s+Block\s+(\d+)/i);
      return m ? (parseInt(m[1], 10) * 100 + parseInt(m[2], 10)) : 9999;
    }
    return rank(a) - rank(b);
  }).map(function(k) {
    return { label: k, debit: rnd(blocks[k].debit), credit: rnd(blocks[k].credit), uncollected: rnd(Math.max(0, blocks[k].debit - blocks[k].credit)) };
  });

  return {
    ok                : true,
    year              : normYear,
    totalDebit        : rnd(totalDebit),
    totalCredit       : rnd(totalCredit),
    totalUncollected  : rnd(Math.max(0, totalDebit - totalCredit)),
    overallReceivables: rnd(overallReceivables),
    phases            : phaseRows,
    blocks            : blockRows,
    monthly           : monthlyRows
  };
}

// ── Status / Dashboard ────────────────────────────────────────
function webapp_getStatus() {
  var mlSh = getSheet_(SH.MASTERLIST);
  var rcSh = getSheet_(SH.RATE_CALC);

  var unitCount = 0;
  if (mlSh && mlSh.getLastRow() > 1) {
    mlSh.getRange(2, 1, mlSh.getLastRow() - 1, 12).getValues().forEach(function(r) {
      if (r[0] && String(r[11] || 'Yes').toLowerCase() !== 'no') unitCount++;
    });
  }
  var lastPeriod = '—';
  var lastRate   = 0;

  if (rcSh && rcSh.getLastRow() > 1) {
    var rcRow  = rcSh.getRange(rcSh.getLastRow(), 1, 1, 13).getValues()[0];
    lastPeriod = rcRow[2] + ' ' + rcRow[1];
    lastRate   = toNum(rcRow[12]);
  }

  var unpostedCount = 0;
  var paySh = getSheet_(SH.PAY_LOG);
  if (paySh && paySh.getLastRow() > 1) {
    paySh.getRange(2, 11, paySh.getLastRow() - 1, 1).getValues().forEach(function(r) {
      if (String(r[0]).trim().toLowerCase() !== 'posted') unpostedCount++;
    });
  }

  return {
    setupDone        : unitCount > 0,
    unitCount        : unitCount,
    lastPeriod       : lastPeriod,
    lastRate         : lastRate,
    unpostedPayments : unpostedCount,
    spreadsheetUrl   : ss_().getUrl(),
    spreadsheetName  : ss_().getName()
  };
}

// ── Water Reading Input ───────────────────────────────────────
function webapp_getWaterInputData() {
  var sh = getSheet_(SH.W_INPUT);
  var now = new Date();

  var config = {
    year    : now.getFullYear(),
    month   : MONTHS[now.getMonth()],
    mcwdAmt: 0, elecFrom: '', elecTo: '', elecAmt: 0,
    manpower: MANPOWER_DEF, addonGlobal: 0
  };

  if (sh) {
    var get = function(cell) { return sh.getRange(cell).getValue(); };
    config.year         = get('B3') || config.year;
    config.month        = get('B4') || config.month;
    config.mcwdAmt      = toNum(get('B7'));
    config.elecFrom     = get('B8') ? fmtDate(get('B8')) : '';
    config.elecTo       = get('B9') ? fmtDate(get('B9')) : '';
    config.elecAmt      = toNum(get('B10'));
    config.manpower     = toNum(get('B11')) || MANPOWER_DEF;
    config.coverageFrom = _readSheetDate_(sh, 'B13');
    config.coverageTo   = _readSheetDate_(sh, 'B14');
    config.addonGlobal  = toNum(get('B15'));
  }

  var table = [];
  if (sh && sh.getLastRow() >= INPUT_TABLE_START) {
    sh.getRange(INPUT_TABLE_START, 1, sh.getLastRow() - INPUT_TABLE_START + 1, 7).getValues()
      .forEach(function(r) {
        if (!r[0]) return;
        table.push({
          unitId   : r[0],
          meterNo  : r[1] || '',
          ownerName: r[2] || '',
          cur      : r[3] !== '' && r[3] !== null ? toNum(r[3]) : '',
          prev     : r[4] !== '' && r[4] !== null ? toNum(r[4]) : '',
          addon    : r[6] !== '' && r[6] !== null ? toNum(r[6]) : 0
        });
      });
  }

  return { config: config, table: table };
}

function webapp_generateBills(config, readings) {
  if (!readings || readings.length === 0) return { ok: false, msg: 'No readings received.' };

  var year      = parseInt(config.year, 10);
  var monthName = String(config.month).trim();
  var month     = getMonthNum(monthName);
  if (!year || !month) return { ok: false, msg: 'Invalid year or month.' };

  // Check if bills for this period were already generated
  var wlChk = getSheet_(SH._WL);
  if (wlChk && wlChk.getLastRow() > 1) {
    var chkData = wlChk.getRange(2, WL.YEAR + 1, wlChk.getLastRow() - 1, 2).getValues();
    var alreadyExists = chkData.some(function(r) {
      return parseInt(r[0], 10) === year && String(r[1]).trim() === monthName;
    });
    if (alreadyExists) {
      return { ok: false, msg: 'Bills for ' + monthName + ' ' + year + ' have already been generated. Delete the existing entries first if you need to re-generate.' };
    }
  }

  var mcwdAmt  = toNum(config.mcwdAmt);
  var elecAmt  = toNum(config.elecAmt);
  var manpower = toNum(config.manpower) || MANPOWER_DEF;
  var totalExpense = mcwdAmt + elecAmt + manpower;

  var coverageFromStr = String(config.coverageFrom || '').trim();
  var coverageToStr   = String(config.coverageTo   || '').trim() || fmtDate(new Date());

  var readingObjs = readings.map(function(r) {
    var prev = toNum(r.prev);
    var cur  = toNum(r.cur);
    var cons = Math.max(0, cur - prev);
    return {
      unitId : String(r.unitId).trim(),
      meterNo: r.meterNo || '',
      prev   : prev,
      cur    : cur,
      cons   : cons,
      addon  : toNum(r.addon),
      remarks: String(r.remarks || '').trim()
    };
  }).filter(function(r) { return r.unitId && r.cur !== ''; });

  var totalCons = readingObjs.reduce(function(s, r) { return s + r.cons; }, 0);
  if (totalCons <= 0) return { ok: false, msg: 'Total consumption is zero. Cannot compute rate.' };

  var rate = totalExpense / totalCons;
  var now = new Date();
  var billDateStr    = fmtDate(getBillDate(year, month));
  var dueDateStr     = fmtDate(getDueDate(year, month));
  var presentDateStr = fmtDate(now);
  var penalty        = 0;

  var newWlRows = [];
  var newDlRows = [];
  var newWsRows = [];

  readingObjs.forEach(function(r) {
    var waterRaw    = r.cons * rate;
    var debit       = r.cons === 0 ? 0 : Math.max(MIN_WATER_BILL, waterRaw);
    var isMinCharge = r.cons > 0 && waterRaw < MIN_WATER_BILL;

    var parsed = parseUID(r.unitId);
    var billNo = parsed
      ? buildBillNum(year, month, parsed.phase, parsed.block, parsed.lot)
      : webapp_buildCommonBillNum_(year, month, r.unitId);

    var autoRem = isMinCharge ? 'Min. charge applied' : '';
    var userRem = r.remarks || '';
    var fullRem = userRem && autoRem ? userRem + '; ' + autoRem : (userRem || autoRem);

    newWlRows.push([
      r.unitId, year, monthName, billDateStr,
      coverageFromStr, coverageToStr,
      r.prev, r.cur, rate, dueDateStr,
      penalty, debit, 0, 0, r.addon,
      '', fullRem, billNo, ''
    ]);

    if (!webapp_isCommonAccount_(r.unitId)) {
      newDlRows.push([r.unitId, year, monthName, '', ASSOC_DUES, 0, 0, '', '']);
    }

    newWsRows.push([
      presentDateStr, year, monthName, r.unitId, r.meterNo,
      r.prev, r.cur, r.cons, rate, debit
    ]);
  });

  _appendRows(SH._WL, newWlRows);
  _appendRows(SH._DL, newDlRows);
  recalcWaterBalances();
  recalcDuesBalances();

  var wsSh = getSheet_(SH.W_STORE);
  if (wsSh && newWsRows.length > 0) {
    wsSh.getRange(wsSh.getLastRow() + 1, 1, newWsRows.length, newWsRows[0].length).setValues(newWsRows);
  }

  _appendRateCalcRow(now, year, monthName,
    config.mcwdFrom, config.mcwdTo, mcwdAmt,
    config.elecFrom, config.elecTo, elecAmt,
    manpower, totalExpense, totalCons, rate);

  var inSh = getSheet_(SH.W_INPUT);
  if (inSh) {
    inSh.getRange('B3').setValue(year);
    inSh.getRange('B4').setValue(monthName);
    inSh.getRange('B7').setValue(mcwdAmt);
    inSh.getRange('B10').setValue(elecAmt);
    inSh.getRange('B11').setValue(manpower);
    if (coverageFromStr) inSh.getRange('B13').setValue(coverageFromStr);
    if (coverageToStr)   inSh.getRange('B14').setValue(coverageToStr);
  }

  refreshMonthlySummary();
  regenerateBillPrint(year, month, monthName, rate, totalCons,
    mcwdAmt, elecAmt, manpower,
    config.mcwdFrom, config.mcwdTo, config.elecFrom, config.elecTo);
  populateWaterInputTable();

  return {
    ok          : true,
    msg         : 'Bills generated!',
    period      : monthName + ' ' + year,
    unitCount   : readingObjs.length,
    rate        : rate,
    totalExpense: totalExpense,
    totalCons   : totalCons,
    billDate    : billDateStr,
    dueDate     : dueDateStr
  };
}

// ── Payments ──────────────────────────────────────────────────
function webapp_getPaymentLog() {
  var sh = getSheet_(SH.PAY_LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues().map(function(r) {
    return {
      dateRec  : r[0]  ? fmtDate(r[0])  : '',
      unitId   : r[1],
      ownerName: r[2],
      payDate  : r[3]  ? fmtDate(r[3])  : '',
      amount   : r[4],
      or       : r[5],
      type     : r[6],
      waterAmt : r[7],
      duesAmt  : r[8],
      remarks  : r[9],
      status   : r[10]
    };
  });
}

function webapp_payDateInfo_(dateVal, year, month) {
  var d = null;
  if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
    d = dateVal;
  } else if (dateVal && String(dateVal).trim()) {
    var parsed = new Date(dateVal);
    if (!isNaN(parsed.getTime())) d = parsed;
  }
  if (d) return { str: fmtDate(d), sort: d.getTime() };

  var y = parseInt(year, 10);
  var m = getMonthNum(String(month || ''));
  if (y && m) {
    var fb = new Date(y, m - 1, 1);
    return { str: fmtDate(fb), sort: fb.getTime() };
  }
  return { str: '', sort: 0 };
}

function webapp_getAllPayments(unitId) {
  var filterUid = String(unitId || '').trim();

  var masterNames = {};
  var mlSh = getSheet_(SH.MASTERLIST);
  if (mlSh && mlSh.getLastRow() > 1) {
    mlSh.getRange(2, 1, mlSh.getLastRow() - 1, 12).getValues().forEach(function(r) {
      var uid = String(r[0] || '').trim();
      if (uid) masterNames[uid] = ownerName(r);
    });
  }

  var rows = [];

  // Central Payment Log
  var plSh = getSheet_(SH.PAY_LOG);
  if (plSh && plSh.getLastRow() > 1) {
    plSh.getRange(2, 1, plSh.getLastRow() - 1, 11).getValues().forEach(function(r, i) {
      var uid = String(r[1] || '').trim();
      if (!uid) return;
      if (filterUid && uid !== filterUid) return;

      var pd = webapp_payDateInfo_(r[3], null, null);
      if (!pd.sort && r[0]) pd.sort = new Date(r[0]).getTime() || 0;

      rows.push({
        source   : 'log',
        rowNum   : i + 2,
        dateRec  : r[0] ? fmtDate(r[0]) : '',
        unitId   : uid,
        ownerName: r[2] || masterNames[uid] || '',
        payDate  : pd.str,
        amount   : toNum(r[4]),
        or       : String(r[5] || ''),
        type     : String(r[6] || ''),
        waterAmt : toNum(r[7]),
        duesAmt  : toNum(r[8]),
        remarks  : String(r[9] || ''),
        status   : String(r[10] || 'Unposted'),
        sortKey  : pd.sort
      });
    });
  }

  // _WaterLedger credits
  var wlSh = getSheet_(SH._WL);
  if (wlSh && wlSh.getLastRow() > 1) {
    wlSh.getRange(2, 1, wlSh.getLastRow() - 1, WL_COLS).getValues().forEach(function(r) {
      var uid = String(r[WL.UNIT] || '').trim();
      if (!uid) return;
      if (filterUid && uid !== filterUid) return;
      var credit = toNum(r[WL.CREDIT]);
      if (credit <= 0) return;

      var pd = webapp_payDateInfo_(r[WL.PAY_DATE], r[WL.YEAR], r[WL.MONTH]);
      rows.push({
        source   : 'water',
        rowNum   : null,
        dateRec  : 'Migrated',
        unitId   : uid,
        ownerName: masterNames[uid] || '',
        payDate  : pd.str,
        amount   : credit,
        or       : String(r[WL.OR] || ''),
        type     : 'Water (migrated)',
        waterAmt : credit,
        duesAmt  : 0,
        remarks  : String(r[WL.REMARKS] || ''),
        status   : 'Migrated',
        sortKey  : pd.sort
      });
    });
  }

  // _DuesLedger credits
  var dlSh = getSheet_(SH._DL);
  if (dlSh && dlSh.getLastRow() > 1) {
    dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues().forEach(function(r) {
      var uid = String(r[DL.UNIT] || '').trim();
      if (!uid) return;
      if (filterUid && uid !== filterUid) return;
      var credit = toNum(r[DL.CREDIT]);
      if (credit <= 0) return;

      var pd = webapp_payDateInfo_(r[DL.PAY_DATE], r[DL.YEAR], r[DL.MONTH]);
      rows.push({
        source   : 'dues',
        rowNum   : null,
        dateRec  : 'Migrated',
        unitId   : uid,
        ownerName: masterNames[uid] || '',
        payDate  : pd.str,
        amount   : credit,
        or       : String(r[DL.OR] || ''),
        type     : 'Dues (migrated)',
        waterAmt : 0,
        duesAmt  : credit,
        remarks  : String(r[DL.REMARKS] || ''),
        status   : 'Migrated',
        sortKey  : pd.sort
      });
    });
  }

  rows.sort(function(a, b) { return b.sortKey - a.sortKey; });
  return rows;
}

function webapp_correctPayment(rowNum, data) {
  var sh = getSheet_(SH.PAY_LOG);
  if (!sh) return { ok: false, msg: 'Payment Log sheet not found.' };
  if (rowNum < 2 || rowNum > sh.getLastRow()) return { ok: false, msg: 'Invalid row.' };

  var existing = sh.getRange(rowNum, 1, 1, 11).getValues()[0];
  var type     = String(data.type || existing[6] || 'Water');
  var waterAmt = type === 'Both' ? toNum(data.waterAmt) : 0;
  var duesAmt  = type === 'Both' ? toNum(data.duesAmt)  : 0;
  var payDate  = data.payDate ? new Date(data.payDate) : existing[3];

  sh.getRange(rowNum, 1, 1, 11).setValues([[
    existing[0],
    existing[1],
    existing[2],
    payDate,
    toNum(data.amount),
    String(data.or || ''),
    type,
    waterAmt,
    duesAmt,
    String(data.remarks || ''),
    'Unposted'
  ]]);
  return { ok: true, msg: 'Payment corrected.' };
}

function webapp_addPayment(p) {
  var sh = getSheet_(SH.PAY_LOG);
  if (!sh) return { ok: false, msg: 'Payment Log sheet not found.' };

  if (!p.ownerName) {
    var mlSh = getSheet_(SH.MASTERLIST);
    if (mlSh && mlSh.getLastRow() > 1) {
      var row = getMLRow(mlSh.getDataRange().getValues(), p.unitId);
      if (row) p.ownerName = ownerName(row);
    }
  }

  var newRow = sh.getLastRow() + 1;
  sh.getRange(newRow, 1, 1, 11).setValues([[
    fmtDate(new Date()),
    p.unitId,
    p.ownerName || '',
    p.payDate   || fmtDate(new Date()),
    p.amount,
    p.or        || '',
    p.type,
    p.waterAmt  || 0,
    p.duesAmt   || 0,
    p.remarks   || '',
    'Unposted'
  ]]);

  // Apply credit immediately to ledgers
  var wlSh   = getSheet_(SH._WL);
  var dlSh   = getSheet_(SH._DL);
  var wlData = (wlSh && wlSh.getLastRow() > 1)
    ? wlSh.getRange(2, 1, wlSh.getLastRow() - 1, WL_COLS).getValues() : [];
  var dlData = (dlSh && dlSh.getLastRow() > 1)
    ? dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues() : [];

  var wlDirty = false, dlDirty = false;
  var payType = String(p.type || '').trim();
  var payDate = p.payDate || fmtDate(new Date());
  var orNum   = p.or      || '';
  var remarks = p.remarks || '';

  try {
    if (payType === 'Water' || payType === 'Both') {
      var wAmt = payType === 'Both' ? toNum(p.waterAmt) : toNum(p.amount);
      if (wAmt > 0) { _applyWaterCredit(wlData, p.unitId, wAmt, payDate, orNum, remarks); wlDirty = true; }
    }
    if (payType === 'Dues' || payType === 'Both') {
      var dAmt = payType === 'Both' ? toNum(p.duesAmt) : toNum(p.amount);
      if (dAmt > 0) {
        _applyDuesCredit(dlData, p.unitId, dAmt, payDate, orNum, remarks,
          p.duesMonth || '', p.duesYear || 0);
        dlDirty = true;
      }
    }

    if (wlDirty && wlData.length > 0) wlSh.getRange(2, 1, wlData.length, WL_COLS).setValues(wlData);
    if (dlDirty && dlData.length > 0) dlSh.getRange(2, 1, dlData.length, DL_COLS).setValues(dlData);
    if (wlDirty) recalcWaterBalances();
    if (dlDirty) recalcDuesBalances();

    sh.getRange(newRow, 11)
      .setValue('Posted').setBackground('#c8e6c9').setFontColor('#1b5e20').setFontWeight('bold');
  } catch (e) {
    Logger.log('webapp_addPayment credit error: ' + e.message);
    sh.getRange(newRow, 11).setValue('Error').setBackground('#ffcdd2');
    return { ok: false, msg: 'Payment saved but credit failed: ' + e.message };
  }

  return { ok: true };
}

// Returns dues ledger rows for a unit where no credit has been posted yet
// (CREDIT = 0), sorted oldest to newest. Used to populate the dues month
// selector in the payment form.
function webapp_getDuesUnpaidMonths(unitId) {
  try {
    var normId = String(unitId || '').trim();
    if (!normId) return { ok: false, msg: 'No unit ID.' };

    var dlSh = getSheet_(SH._DL);
    if (!dlSh || dlSh.getLastRow() < 2) return { ok: true, months: [] };

    var dlData = dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues();
    var months = [];

    dlData.forEach(function(r) {
      if (String(r[DL.UNIT] || '').trim() !== normId) return;
      if (toNum(r[DL.DEBIT])  <= 0) return;           // no bill generated yet
      if (toNum(r[DL.CREDIT]) > 0) return;            // already has a payment
      var month = String(r[DL.MONTH] || '').trim();
      var year  = toNum(r[DL.YEAR]);
      if (!month || !year) return;
      months.push({
        month : month,
        year  : year,
        label : month + ' ' + year,
        rank  : year * 100 + getMonthNum(month)
      });
    });

    months.sort(function(a, b) { return a.rank - b.rank; });

    return {
      ok: true,
      months: months.map(function(m) {
        return { month: m.month, year: m.year, label: m.label };
      })
    };
  } catch (e) {
    Logger.log('webapp_getDuesUnpaidMonths: ' + e.message);
    return { ok: false, msg: e.message };
  }
}

function webapp_postAllPayments() {
  var paySh = getSheet_(SH.PAY_LOG);
  if (!paySh || paySh.getLastRow() < 2) return { ok: false, msg: 'No payment records found.' };

  var numRows = paySh.getLastRow() - 1;
  var data    = paySh.getRange(2, 1, numRows, 11).getValues();

  var unposted = [];
  data.forEach(function(r, i) {
    var uid    = String(r[1]).trim();
    var status = String(r[10]).trim().toLowerCase();
    if (!uid || status === 'posted') return;
    unposted.push({
      rowIdx  : i + 2,
      unitId  : uid,
      payDate : r[3],
      totalAmt: toNum(r[4]),
      orNum   : r[5],
      payType : String(r[6]).trim(),
      waterAmt: toNum(r[7]),
      duesAmt : toNum(r[8]),
      remarks : String(r[9]).trim()
    });
  });

  if (unposted.length === 0) return { ok: false, msg: 'No unposted payments found.' };

  var wlSh   = getSheet_(SH._WL);
  var dlSh   = getSheet_(SH._DL);
  var wlData = (wlSh && wlSh.getLastRow() > 1)
    ? wlSh.getRange(2, 1, wlSh.getLastRow() - 1, WL_COLS).getValues() : [];
  var dlData = (dlSh && dlSh.getLastRow() > 1)
    ? dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues() : [];

  var posted = 0, errors = 0, wlDirty = false, dlDirty = false;

  unposted.forEach(function(p) {
    try {
      if (p.payType === 'Water' || p.payType === 'Both') {
        var wAmt = p.payType === 'Both' ? p.waterAmt : p.totalAmt;
        if (wAmt > 0) { _applyWaterCredit(wlData, p.unitId, wAmt, p.payDate, p.orNum, p.remarks); wlDirty = true; }
      }
      if (p.payType === 'Dues' || p.payType === 'Both') {
        var dAmt = p.payType === 'Both' ? p.duesAmt : p.totalAmt;
        if (dAmt > 0) { _applyDuesCredit(dlData, p.unitId, dAmt, p.payDate, p.orNum, p.remarks); dlDirty = true; }
      }
      paySh.getRange(p.rowIdx, 11)
        .setValue('Posted').setBackground('#c8e6c9').setFontColor('#1b5e20').setFontWeight('bold');
      posted++;
    } catch (e) {
      Logger.log(e); errors++;
      paySh.getRange(p.rowIdx, 11).setValue('Error').setBackground('#ffcdd2');
    }
  });

  if (wlDirty && wlData.length > 0) wlSh.getRange(2, 1, wlData.length, WL_COLS).setValues(wlData);
  if (dlDirty && dlData.length > 0) dlSh.getRange(2, 1, dlData.length, DL_COLS).setValues(dlData);
  if (wlDirty) recalcWaterBalances();
  if (dlDirty) recalcDuesBalances();
  refreshMonthlySummary();

  return { ok: true, posted: posted, errors: errors };
}

// ── Unit Ledger ───────────────────────────────────────────────
function webapp_getUnitIds() {
  var sh = getSheet_(SH.MASTERLIST);
  if (!sh || sh.getLastRow() < 2) return [];

  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues()
    .filter(function(r) { return r[0] && String(r[11] || 'Yes').toLowerCase() !== 'no'; })
    .map(function(r) {
      var id = String(r[0] || '').trim();
      var name = ownerName(r);
      if (!name && webapp_isCommonAccount_(id)) name = id;
      return { id: id, name: name, common: webapp_isCommonAccount_(id) };
    });

  ids.sort(function(a, b) {
    if (a.common && !b.common) return 1;
    if (!a.common && b.common) return -1;
    return String(a.id).localeCompare(String(b.id), undefined, { numeric: true });
  });

  return ids;
}

function webapp_getUnitLedger(unitId) {
  var result = { water: [], dues: [], ownerName: '', unitId: unitId };

  var mlSh = getSheet_(SH.MASTERLIST);
  if (mlSh && mlSh.getLastRow() > 1) {
    var row = getMLRow(mlSh.getDataRange().getValues(), unitId);
    if (row) result.ownerName = ownerName(row);
  }

  var normId = String(unitId || '').trim();

  var wlSh = getSheet_(SH._WL);
  if (wlSh && wlSh.getLastRow() > 1) {
    wlSh.getRange(2, 1, wlSh.getLastRow() - 1, WL_COLS).getValues().forEach(function(r, i) {
      if (String(r[WL.UNIT] || '').trim() !== normId) return;
      result.water.push({
        sheetRow   : i + 2,
        year       : r[WL.YEAR],
        month      : r[WL.MONTH],
        billDate   : fmtDate(r[WL.BILL_DATE]),
        dueDate    : fmtDate(r[WL.DUE_DATE]),
        prevDate   : fmtDate(r[WL.PREV_DATE]),
        presentDate: fmtDate(r[WL.PRESENT_DATE]),
        prevRdg    : r[WL.PREV_RDG],
        curRdg     : r[WL.CUR_RDG],
        rate       : r[WL.RATE],
        penalty    : r[WL.PENALTY],
        debit      : r[WL.DEBIT],
        credit     : r[WL.CREDIT],
        balance    : r[WL.BALANCE],
        addon      : r[WL.ADDON],
        or         : r[WL.OR],
        remarks    : r[WL.REMARKS],
        payDate    : fmtDate(r[WL.PAY_DATE]),
        billNo     : r[WL.BILL_NO]
      });
    });
  }

  var dlSh = getSheet_(SH._DL);
  if (dlSh && dlSh.getLastRow() > 1) {
    dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues().forEach(function(r, i) {
      if (String(r[DL.UNIT] || '').trim() !== normId) return;
      result.dues.push({
        sheetRow: i + 2,
        year   : r[DL.YEAR],
        month  : r[DL.MONTH],
        debit  : r[DL.DEBIT],
        credit : r[DL.CREDIT],
        balance: r[DL.BALANCE],
        payDate: fmtDate(r[DL.PAY_DATE]),
        or     : r[DL.OR],
        remarks: r[DL.REMARKS]
      });
    });
  }

  return result;
}

function webapp_updateWaterRow(sheetRow, data) {
  var sh = getSheet_(SH._WL);
  if (!sh) return { ok: false, msg: '_WaterLedger not found.' };
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { ok: false, msg: 'Invalid row.' };

  var row = sh.getRange(sheetRow, 1, 1, WL_COLS).getValues()[0];
  row[WL.PAY_DATE] = data.payDate ? new Date(data.payDate) : '';
  row[WL.CREDIT]   = toNum(data.credit);
  row[WL.PENALTY]  = toNum(data.penalty);
  row[WL.ADDON]    = toNum(data.addon);
  row[WL.OR]       = String(data.or || '');
  row[WL.REMARKS]  = String(data.remarks || '');
  sh.getRange(sheetRow, 1, 1, WL_COLS).setValues([row]);
  recalcWaterBalances();
  return { ok: true };
}

function webapp_updateDuesRow(sheetRow, data) {
  var sh = getSheet_(SH._DL);
  if (!sh) return { ok: false, msg: '_DuesLedger not found.' };
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { ok: false, msg: 'Invalid row.' };

  var row = sh.getRange(sheetRow, 1, 1, DL_COLS).getValues()[0];
  row[DL.PAY_DATE] = data.payDate ? new Date(data.payDate) : '';
  row[DL.CREDIT]   = toNum(data.credit);
  row[DL.OR]       = String(data.or || '');
  row[DL.REMARKS]  = String(data.remarks || '');
  sh.getRange(sheetRow, 1, 1, DL_COLS).setValues([row]);
  recalcDuesBalances();
  return { ok: true };
}

// ── Monthly Summary ───────────────────────────────────────────
function webapp_getMonthlySummary() {
  var mlSh = getSheet_(SH.MASTERLIST);
  var wlSh = getSheet_(SH._WL);
  var dlSh = getSheet_(SH._DL);

  var masterRows = (mlSh && mlSh.getLastRow() > 1)
    ? mlSh.getRange(2, 1, mlSh.getLastRow() - 1, 12).getValues() : [];
  var waterRows = (wlSh && wlSh.getLastRow() > 1)
    ? wlSh.getRange(2, 1, wlSh.getLastRow() - 1, WL_COLS).getValues() : [];
  var duesRows = (dlSh && dlSh.getLastRow() > 1)
    ? dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues() : [];

  var units = [];
  var unitExists = {};
  masterRows.forEach(function(r) {
    var unitId = String(r[0] || '').trim();
    var active = String(r[11] || 'Yes').trim().toLowerCase();
    if (!unitId || active === 'no') return;
    units.push({
      unitId: unitId,
      name: ownerName(r) || sg_commonName_(unitId),
      meter: r[9] || '',
      common: sg_isCommonAccount_(unitId)
    });
    unitExists[unitId] = true;
  });

  waterRows.forEach(function(r) {
    var uid = String(r[WL.UNIT] || '').trim();
    if (!uid || unitExists[uid]) return;
    units.push({ unitId: uid, name: sg_commonName_(uid), meter: '', common: sg_isCommonAccount_(uid) });
    unitExists[uid] = true;
  });
  duesRows.forEach(function(r) {
    var uid = String(r[DL.UNIT] || '').trim();
    if (!uid || unitExists[uid]) return;
    units.push({ unitId: uid, name: sg_commonName_(uid), meter: '', common: sg_isCommonAccount_(uid) });
    unitExists[uid] = true;
  });

  units.sort(function(a, b) {
    if (a.common && !b.common) return 1;
    if (!a.common && b.common) return -1;
    return String(a.unitId).localeCompare(String(b.unitId), undefined, { numeric: true });
  });

  var latestWaterByUnit = {};
  waterRows.forEach(function(r) {
    var uid = String(r[WL.UNIT] || '').trim();
    if (uid) latestWaterByUnit[uid] = r;
  });
  var latestDuesByUnit = {};
  duesRows.forEach(function(r) {
    var uid = String(r[DL.UNIT] || '').trim();
    if (uid) latestDuesByUnit[uid] = r;
  });

  var period = sg_getLatestPeriod_(waterRows, duesRows);

  var rows = [];
  units.forEach(function(u) {
    var w = latestWaterByUnit[u.unitId] || null;
    var d = latestDuesByUnit[u.unitId] || null;

    var prevRdg    = w ? toNum(w[WL.PREV_RDG]) : 0;
    var curRdg     = w ? toNum(w[WL.CUR_RDG])  : 0;
    var cons       = w ? Math.max(0, curRdg - prevRdg) : 0;
    var rate       = w ? toNum(w[WL.RATE])      : 0;
    var waterBill  = w ? toNum(w[WL.DEBIT])     : 0;
    var penalty    = w ? toNum(w[WL.PENALTY])   : 0;
    var addon      = w ? toNum(w[WL.ADDON])     : 0;
    var waterBalance = w ? toNum(w[WL.BALANCE]) : 0;

    var waterArrears = waterBalance > 0 ? waterBalance : 0;
    var totalWater   = waterArrears;

    var duesDebit   = d ? toNum(d[DL.DEBIT])   : 0;
    var duesBalance = d ? toNum(d[DL.BALANCE])  : 0;
    var assocDue    = u.common ? 0 : duesDebit;
    var duesArrears = u.common ? 0 : (duesBalance > 0 ? duesBalance : 0);
    var totalDues   = duesArrears;

    var grandTotal = totalWater + totalDues;
    var status = grandTotal <= 0 ? 'Paid' : 'Unpaid';
    if (grandTotal > 0) {
      var hasAnyCredit = (w && toNum(w[WL.CREDIT]) > 0) || (d && toNum(d[DL.CREDIT]) > 0);
      status = hasAnyCredit ? 'Partial' : 'Unpaid';
    }

    var dueDate = w ? fmtDate(w[WL.DUE_DATE]) : '';

    rows.push({
      unitId: u.unitId, name: u.name, meter: u.meter,
      cons: cons, rate: rate, waterBill: waterBill,
      penalty: penalty, waterArrears: waterArrears, addon: addon,
      totalWater: totalWater, assocDue: assocDue, duesArrears: duesArrears,
      totalDues: totalDues, grandTotal: grandTotal, status: status,
      dueDate: dueDate,
      group: u.common ? 'common' : (String(u.unitId).indexOf('P1') === 0 ? '1' : (String(u.unitId).indexOf('P2') === 0 ? '2' : 'other'))
    });
  });

  return { rows: rows, period: period };
}

// ── Setup utilities ───────────────────────────────────────────
function webapp_runSetup() {
  try {
    createAllSheets();
    importMasterlistFromSource();
    populateWaterInputTable();
    return { ok: true, msg: 'Setup complete.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function webapp_importMasterlist() {
  try {
    importMasterlistFromSource();
    var mlSh  = getSheet_(SH.MASTERLIST);
    var count = mlSh ? Math.max(0, mlSh.getLastRow() - 1) : 0;
    return { ok: true, msg: 'Masterlist imported.', unitCount: count };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function webapp_refreshReadingTable() {
  try {
    populateWaterInputTable();
    return { ok: true };
  } catch (e) { return { ok: false, msg: e.message }; }
}

// ── Masterlist ────────────────────────────────────────────────
function webapp_getMasterlist() {
  var sh = getSheet_(SH.MASTERLIST);
  if (!sh || sh.getLastRow() < 2) return { ok: true, rows: [] };
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues();
  var rows = data.map(function(r, i) {
    return {
      sheetRow    : i + 2,
      unitId      : String(r[0]  || ''),
      phase       : String(r[1]  || ''),
      block       : String(r[2]  || ''),
      lot         : String(r[3]  || ''),
      lastName    : String(r[4]  || ''),
      firstName   : String(r[5]  || ''),
      dateAccepted: fmtDate(r[6]),
      contact     : String(r[7]  || ''),
      email       : String(r[8]  || ''),
      meterNo     : String(r[9]  || ''),
      stubout     : String(r[10] || ''),
      active      : String(r[11] || 'Yes')
    };
  });
  return { ok: true, rows: rows };
}

function webapp_updateMasterlistRow(sheetRow, data) {
  var sh = getSheet_(SH.MASTERLIST);
  if (!sh) return { ok: false, msg: 'Masterlist not found.' };
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { ok: false, msg: 'Invalid row.' };
  var row = sh.getRange(sheetRow, 1, 1, 12).getValues()[0];
  // Unit ID (col 0) is read-only — skip
  row[1]  = data.phase    !== undefined ? data.phase    : row[1];
  row[2]  = data.block    !== undefined ? data.block    : row[2];
  row[3]  = data.lot      !== undefined ? data.lot      : row[3];
  row[4]  = data.lastName !== undefined ? data.lastName : row[4];
  row[5]  = data.firstName!== undefined ? data.firstName: row[5];
  row[6]  = data.dateAccepted ? (new Date(data.dateAccepted)) : row[6];
  row[7]  = data.contact  !== undefined ? data.contact  : row[7];
  row[8]  = data.email    !== undefined ? data.email    : row[8];
  row[9]  = data.meterNo  !== undefined ? data.meterNo  : row[9];
  row[10] = data.stubout  !== undefined ? data.stubout  : row[10];
  row[11] = data.active   !== undefined ? data.active   : row[11];
  sh.getRange(sheetRow, 1, 1, 12).setValues([row]);
  return { ok: true };
}

function webapp_addMasterlistRow(data) {
  var sh = getSheet_(SH.MASTERLIST);
  if (!sh) return { ok: false, msg: 'Masterlist not found.' };
  var uid = String(data.unitId || '').trim();
  if (!uid) return { ok: false, msg: 'Unit ID is required.' };
  // Prevent duplicates
  if (sh.getLastRow() > 1) {
    var existing = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i][0]).trim().toUpperCase() === uid.toUpperCase())
        return { ok: false, msg: 'Unit ID already exists.' };
    }
  }
  sh.appendRow([uid, data.phase||'', data.block||'', data.lot||'',
                data.lastName||'', data.firstName||'',
                data.dateAccepted ? new Date(data.dateAccepted) : '',
                data.contact||'', data.email||'',
                data.meterNo||'', data.stubout||'',
                data.active||'Yes']);
  return { ok: true };
}

// ── Board of Directors ────────────────────────────────────────
function webapp_getBOD() {
  var sh = getSheet_(SH._BOD);
  if (!sh) { _ensureBODSheet_(); sh = getSheet_(SH._BOD); }
  if (!sh || sh.getLastRow() < 2) return { ok: true, rows: [] };
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  var rows = data.map(function(r, i) {
    return {
      sheetRow      : i + 2,
      name          : r[0], position     : r[1], unit         : r[2],
      fromMonth     : r[3], fromYear     : r[4],
      toMonth       : r[5], toYear       : r[6]
    };
  });
  return { ok: true, rows: rows };
}

function webapp_saveBOD(sheetRow, data) {
  var sh = getSheet_(SH._BOD);
  if (!sh) { _ensureBODSheet_(); sh = getSheet_(SH._BOD); }
  var row = [
    String(data.name     || ''), String(data.position  || ''),
    String(data.unit     || ''), String(data.fromMonth || ''),
    toNum(data.fromYear),        String(data.toMonth   || ''),
    data.toYear ? toNum(data.toYear) : ''
  ];
  if (sheetRow === -1) {
    sh.appendRow(row);
  } else {
    if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { ok: false, msg: 'Invalid row.' };
    sh.getRange(sheetRow, 1, 1, 7).setValues([row]);
  }
  return { ok: true };
}

function webapp_deleteBOD(sheetRow) {
  var sh = getSheet_(SH._BOD);
  if (!sh) return { ok: false, msg: 'BOD sheet not found.' };
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { ok: false, msg: 'Invalid row.' };
  sh.deleteRow(sheetRow);
  return { ok: true };
}

function _ensureBODSheet_() {
  var bodH = ['NAME','POSITION','UNIT','EXEMPT_FROM_MONTH','EXEMPT_FROM_YEAR',
              'EXEMPT_TO_MONTH','EXEMPT_TO_YEAR'];
  _initHiddenSheet(SH._BOD, bodH, '#ffe0b2');
}

// ── Monthly Dues Rates ────────────────────────────────────────
function webapp_getDuesRates() {
  var sh = getSheet_(SH._DUES_RATES);
  if (!sh) { _ensureDuesRatesSheet_(); sh = getSheet_(SH._DUES_RATES); }
  if (!sh || sh.getLastRow() < 2) return { ok: true, rows: [] };
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  var rows = data.map(function(r, i) {
    return {
      sheetRow : i + 2,
      amount   : r[0], fromMonth: r[1], fromYear : r[2],
      toMonth  : r[3], toYear   : r[4]
    };
  });
  return { ok: true, rows: rows };
}

function webapp_saveDuesRate(sheetRow, data) {
  var sh = getSheet_(SH._DUES_RATES);
  if (!sh) { _ensureDuesRatesSheet_(); sh = getSheet_(SH._DUES_RATES); }
  var row = [
    toNum(data.amount),      String(data.fromMonth || ''),
    toNum(data.fromYear),    String(data.toMonth   || ''),
    data.toYear ? toNum(data.toYear) : ''
  ];
  if (sheetRow === -1) {
    sh.appendRow(row);
  } else {
    if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { ok: false, msg: 'Invalid row.' };
    sh.getRange(sheetRow, 1, 1, 5).setValues([row]);
  }
  return { ok: true };
}

function webapp_deleteDuesRate(sheetRow) {
  var sh = getSheet_(SH._DUES_RATES);
  if (!sh) return { ok: false, msg: 'Dues rates sheet not found.' };
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { ok: false, msg: 'Invalid row.' };
  sh.deleteRow(sheetRow);
  return { ok: true };
}

function _ensureDuesRatesSheet_() {
  var drH = ['AMOUNT','FROM_MONTH','FROM_YEAR','TO_MONTH','TO_YEAR'];
  _initHiddenSheet(SH._DUES_RATES, drH, '#e8f5e9');
}

// ── Bills: Return distinct billing periods, newest first ─────
function webapp_getBillingPeriods() {
  var sh = getSheet_(SH._WL);
  if (!sh || sh.getLastRow() < 2) return { ok: true, periods: [] };

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  var seen = {};
  var periods = [];

  data.forEach(function(r) {
    var yr = parseInt(r[WL.YEAR], 10);
    var mo = String(r[WL.MONTH] || '').trim();
    if (!yr || !mo) return;
    var key = yr + '|' + mo;
    if (!seen[key]) {
      seen[key] = true;
      periods.push({ year: yr, month: mo, monthNum: getMonthNum(mo) });
    }
  });

  periods.sort(function(a, b) {
    if (b.year !== a.year) return b.year - a.year;
    return b.monthNum - a.monthNum;
  });

  return { ok: true, periods: periods.map(function(p) { return { year: p.year, month: p.month }; }) };
}

// ── Bills: Build bill data for all units in a billing period ─
function webapp_getBillData(year, month) {
  year = parseInt(year, 10);
  var monthNum = getMonthNum(month);

  var wlSh = getSheet_(SH._WL);
  if (!wlSh || wlSh.getLastRow() < 2) return { ok: false, msg: 'Water ledger is empty.' };
  var wlData = wlSh.getRange(2, 1, wlSh.getLastRow() - 1, WL_COLS).getValues();

  var dlSh   = getSheet_(SH._DL);
  var dlData = (dlSh && dlSh.getLastRow() > 1)
    ? dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues() : [];

  var mlSh  = getSheet_(SH.MASTERLIST);
  var mlRaw = (mlSh && mlSh.getLastRow() > 1) ? mlSh.getDataRange().getValues() : [[]];
  var mlMap = {};
  for (var mi = 1; mi < mlRaw.length; mi++) {
    var mr = mlRaw[mi];
    var muid = String(mr[0] || '').trim();
    if (muid) mlMap[muid] = mr;
  }

  // Rate data for this period from Rate Calculator
  var rcSh = getSheet_(SH.RATE_CALC);
  var rateData = { mcwdAmt: 0, elecAmt: 0, manpower: 0, totalExpense: 0, totalCons: 0, rate: 0 };
  if (rcSh && rcSh.getLastRow() > 1) {
    var rcRows = rcSh.getRange(2, 1, rcSh.getLastRow() - 1, 13).getValues();
    for (var ri = rcRows.length - 1; ri >= 0; ri--) {
      if (parseInt(rcRows[ri][1], 10) === year && String(rcRows[ri][2]).trim() === month) {
        rateData = {
          mcwdAmt: toNum(rcRows[ri][5]), elecAmt: toNum(rcRows[ri][8]),
          manpower: toNum(rcRows[ri][9]), totalExpense: toNum(rcRows[ri][10]),
          totalCons: toNum(rcRows[ri][11]), rate: toNum(rcRows[ri][12])
        };
        break;
      }
    }
  }

  // Scan all WL rows in order to track running balance per unit.
  // First row per unit: use stored opening balance (may be manually set).
  var unitRunBal  = {};   // uid → running balance after last processed row
  var unitWlSeen  = {};   // uid → first row seen flag
  var periodRows  = {};   // uid → { wlIdx, newBal, prevBal }

  for (var i = 0; i < wlData.length; i++) {
    var r      = wlData[i];
    var uid    = String(r[WL.UNIT] || '').trim();
    if (!uid) continue;

    var prevBal, newBal;
    if (!unitWlSeen[uid]) {
      unitWlSeen[uid] = true;
      var storedBal = toNum(r[WL.BALANCE]);
      var d0 = toNum(r[WL.DEBIT]), p0 = toNum(r[WL.PENALTY]);
      var a0 = toNum(r[WL.ADDON]), c0 = toNum(r[WL.CREDIT]);
      prevBal = storedBal - d0 - p0 - a0 + c0;  // reverse old formula to get opening
      newBal  = storedBal;
    } else {
      prevBal = unitRunBal[uid];
      newBal  = Math.round((prevBal + toNum(r[WL.DEBIT]) + toNum(r[WL.PENALTY]) +
                            toNum(r[WL.ADDON]) - toNum(r[WL.CREDIT])) * 100) / 100;
    }
    unitRunBal[uid] = newBal;

    if (parseInt(r[WL.YEAR], 10) === year && String(r[WL.MONTH] || '').trim() === month) {
      periodRows[uid] = { wlIdx: i, newBal: newBal, prevBal: prevBal };
    }
  }

  // Disconnection flag: both calendar months immediately before the current period
  // must have debit > 0 AND (debit + penalty + addon - credit) > 0
  var p1Mo = monthNum - 1, p1Yr = year;
  if (p1Mo < 1) { p1Mo += 12; p1Yr--; }
  var p2Mo = monthNum - 2, p2Yr = year;
  if (p2Mo < 1) { p2Mo += 12; p2Yr--; }

  // unitPrev[uid][1|2] = true (unpaid) / false (paid or zero-bill) / undefined (no entry)
  var unitPrev1 = {}, unitPrev2 = {};
  for (var j = 0; j < wlData.length; j++) {
    var u2    = String(wlData[j][WL.UNIT] || '').trim();
    if (!u2) continue;
    var rYr   = parseInt(wlData[j][WL.YEAR], 10);
    var rMo   = getMonthNum(String(wlData[j][WL.MONTH] || '').trim());
    var deb2  = toNum(wlData[j][WL.DEBIT]);
    if (deb2 <= 0) continue;   // skip zero-bill months
    var net2  = deb2 + toNum(wlData[j][WL.PENALTY]) +
                toNum(wlData[j][WL.ADDON]) - toNum(wlData[j][WL.CREDIT]);
    if (rYr === p1Yr && rMo === p1Mo) unitPrev1[u2] = net2 > 0.005;
    if (rYr === p2Yr && rMo === p2Mo) unitPrev2[u2] = net2 > 0.005;
  }

  // Build dues summary per unit: cumulative running balance (prev + debit - credit).
  // Track running balance through all rows; capture prevBalance and newBalance for billing period.
  var unitDuesRunBal = {};  // uid → running balance up to current row
  var unitDues = {};        // uid → { prevBalance, currentDues, newBalance }
  for (var k = 0; k < dlData.length; k++) {
    var dr      = dlData[k];
    var du      = String(dr[DL.UNIT] || '').trim();
    if (!du) continue;
    var dDeb    = toNum(dr[DL.DEBIT]);
    var dCre    = toNum(dr[DL.CREDIT]);
    var dPrev   = unitDuesRunBal[du] || 0;
    var dNet    = Math.round((dPrev + dDeb - dCre) * 100) / 100;
    unitDuesRunBal[du] = dNet;
    var rowYr   = parseInt(dr[DL.YEAR], 10);
    var rowMoN  = getMonthNum(String(dr[DL.MONTH] || '').trim());
    var rowRank = rowYr * 100 + rowMoN;
    var curRank = year * 100 + monthNum;

    if (rowRank === curRank) {
      unitDues[du] = {
        prevBalance : dPrev,
        currentDues : dDeb,
        newBalance  : dNet
      };
    }
  }

  var bills = [];
  Object.keys(periodRows).forEach(function(uid) {
    var pr  = periodRows[uid];
    var wlR = wlData[pr.wlIdx];
    var ml  = mlMap[uid];

    var deb   = toNum(wlR[WL.DEBIT]);
    var pen   = toNum(wlR[WL.PENALTY]);
    var addon = toNum(wlR[WL.ADDON]);
    // totalWaterDue = running balance after this row (negative = owes, positive = credit)
    var originalTotal = pr.newBal;
    // Arrears = prevBal before this row (negative = owes, positive = credit carry-forward)
    var arrears = Math.round(pr.prevBal * 100) / 100;

    var lastName  = ml ? String(ml[4] || '').trim() : '';
    var firstName = ml ? String(ml[5] || '').trim() : '';
    var owner     = lastName ? (lastName + ', ' + firstName).toUpperCase() : uid;
    var dues      = unitDues[uid] || { prevBalance: 0, currentDues: 0, newBalance: 0 };

    // Coverage date range from WL.PREV_DATE / WL.PRESENT_DATE
    var covFrom = _fmtShortDate_(wlR[WL.PREV_DATE]);
    var covTo   = _fmtShortDate_(wlR[WL.PRESENT_DATE]);
    var covRange = (covFrom && covTo) ? (covFrom + ' - ' + covTo)
                 : (covFrom || covTo || (month + ' ' + year));

    bills.push({
      unitId          : uid,
      phase           : ml ? String(ml[1]  || '').trim() : '',
      owner           : owner,
      meterNo         : ml ? String(ml[9]  || '').trim() : '',
      stubout         : ml ? String(ml[10] || '').trim() : '',
      billNo          : String(wlR[WL.BILL_NO]  || '').trim(),
      billDate        : _fmtBillDate_(wlR[WL.BILL_DATE]),
      dueDate         : _fmtBillDate_(wlR[WL.DUE_DATE]),
      coverage        : covRange,
      prevReadingDate : covFrom,
      curReadingDate  : covTo,
      prevReading     : toNum(wlR[WL.PREV_RDG]),
      curReading      : toNum(wlR[WL.CUR_RDG]),
      consumption     : Math.max(0, toNum(wlR[WL.CUR_RDG]) - toNum(wlR[WL.PREV_RDG])),
      rate            : toNum(wlR[WL.RATE]),
      arrears         : arrears,
      debit           : Math.round(deb * 100) / 100,
      penalty         : Math.round(pen * 100) / 100,
      totalWaterDue   : originalTotal,
      disconnection   : (unitPrev1[uid] === true && unitPrev2[uid] === true),
      duesPrevBalance : dues.prevBalance,
      duesCurrentDues : dues.currentDues,
      duesNewBalance  : dues.newBalance !== undefined ? dues.newBalance : 0
    });
  });

  // Natural numeric sort: P1B1L1 → P1B1L2 → … → P2… → Common accounts last
  function _billSortKey_(uid) {
    var v = String(uid || '').trim().toUpperCase().replace(/\s+/g, '');
    if (webapp_isCommonAccount_(v)) return [999, 999, 999];
    var m = v.match(/^P(\d+)B(\d+)L(\d+)/);
    if (m) return [parseInt(m[1], 10), parseInt(m[2], 10), parseInt(m[3], 10)];
    return [998, 998, 998];
  }
  bills.sort(function(a, b) {
    var ka = _billSortKey_(a.unitId), kb = _billSortKey_(b.unitId);
    return ka[0] - kb[0] || ka[1] - kb[1] || ka[2] - kb[2];
  });

  return { ok: true, period: { year: year, month: month }, rateData: rateData, bills: bills };
}

// Read a cell value and return it as a formatted date string only if it parses
// to a reasonable year (2000–2099). Returns '' for empty cells, non-date strings,
// numbers (e.g., stale meter numbers), or out-of-range years.
function _readSheetDate_(sh, cell) {
  var raw = sh.getRange(cell).getValue();
  if (!raw) return '';
  var d = (raw instanceof Date) ? raw : new Date(String(raw));
  if (isNaN(d.getTime())) return '';
  var yr = d.getFullYear();
  if (yr < 2000 || yr > 2099) return '';
  return fmtDate(d);
}

// Format a date value (Date object or string) as "Month DD, YYYY"
function _fmtBillDate_(v) {
  if (!v) return '';
  var d = (v instanceof Date) ? v : new Date(String(v));
  if (isNaN(d.getTime())) return String(v);
  var MN = ['January','February','March','April','May','June',
            'July','August','September','October','November','December'];
  return MN[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
}

// Format a date value as "MMM DD, YYYY" (3-letter month, 2-digit day)
function _fmtShortDate_(v) {
  if (!v) return '';
  var d = (v instanceof Date) ? v : new Date(String(v));
  if (isNaN(d.getTime())) return String(v);
  var MN = ['Jan','Feb','Mar','Apr','May','Jun',
            'Jul','Aug','Sep','Oct','Nov','Dec'];
  var dd = String(d.getDate()).padStart(2, '0');
  return MN[d.getMonth()] + ' ' + dd + ', ' + d.getFullYear();
}

// ── Delinquent Payers ─────────────────────────────────────────
function webapp_getDelinquentPayers() {
  var mlSh = getSheet_(SH.MASTERLIST);
  var wlSh = getSheet_(SH._WL);
  var dlSh = getSheet_(SH._DL);

  var nameMap = {}, phaseMap = {}, blockMap = {};
  if (mlSh && mlSh.getLastRow() > 1) {
    mlSh.getRange(2, 1, mlSh.getLastRow() - 1, 12).getValues().forEach(function(r) {
      var uid = String(r[0] || '').trim();
      if (!uid) return;
      nameMap[uid]  = ownerName(r);
      phaseMap[uid] = String(r[1] || '');
      blockMap[uid] = String(r[2] || '');
    });
  }

  // Water delinquents — positive balance in latest month, last payment >= 2 months before latest
  var wdLatestRank = 0, wdLatestYear = 0, wdLatestMonth = 0;
  var wdAllRows = (wlSh && wlSh.getLastRow() > 1)
    ? wlSh.getRange(2, 1, wlSh.getLastRow() - 1, WL_COLS).getValues() : [];
  wdAllRows.forEach(function(r) {
    var yr = parseInt(r[WL.YEAR], 10);
    var mo = getMonthNum(String(r[WL.MONTH] || '').trim());
    var rank = yr * 100 + mo;
    if (rank > wdLatestRank) { wdLatestRank = rank; wdLatestYear = yr; wdLatestMonth = mo; }
  });

  // Threshold: 1st day of (latest month - 2); payments before this date are "stale"
  var wdThreshYear = wdLatestYear, wdThreshMonth = wdLatestMonth - 2;
  if (wdThreshMonth <= 0) { wdThreshMonth += 12; wdThreshYear--; }
  var wdThreshDate = new Date(wdThreshYear, wdThreshMonth - 1, 1);

  var wDelinq = {};  // uid → { balance, lastPay (Date or null) }
  wdAllRows.forEach(function(r) {
    var uid = String(r[WL.UNIT] || '').trim();
    if (!uid) return;
    if (!wDelinq[uid]) wDelinq[uid] = { balance: 0, lastPay: null };
    var yr = parseInt(r[WL.YEAR], 10);
    var mo = getMonthNum(String(r[WL.MONTH] || '').trim());
    if (yr === wdLatestYear && mo === wdLatestMonth) {
      wDelinq[uid].balance = toNum(r[WL.BALANCE]);
    }
    var pd = r[WL.PAY_DATE];
    if (pd && toNum(r[WL.CREDIT]) > 0) {
      var pdDate = parseDate(pd);
      if (pdDate && (!wDelinq[uid].lastPay || pdDate > wDelinq[uid].lastPay)) {
        wDelinq[uid].lastPay = pdDate;
      }
    }
  });

  var waterList = [];
  Object.keys(wDelinq).forEach(function(uid) {
    var s = wDelinq[uid];
    if (s.balance <= 0.005) return;                          // must owe in latest month
    if (s.lastPay && s.lastPay >= wdThreshDate) return;     // paid recently enough
    waterList.push({
      unitId: uid, ownerName: nameMap[uid] || '',
      phase: phaseMap[uid] || '', block: blockMap[uid] || '',
      balance: Math.round(s.balance * 100) / 100,
      lastPayDate: s.lastPay ? fmtDate(s.lastPay) : ''
    });
  });
  waterList.sort(function(a, b) { return b.balance - a.balance; });

  // Dues delinquents — last row per unit has cumulative running balance (total owed)
  var dStats = {};
  if (dlSh && dlSh.getLastRow() > 1) {
    dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues().forEach(function(r) {
      var uid = String(r[DL.UNIT] || '').trim();
      if (!uid) return;
      if (!dStats[uid]) dStats[uid] = { months: 0, balance: 0, lastPay: '' };
      var deb = toNum(r[DL.DEBIT]), cre = toNum(r[DL.CREDIT]);
      if (deb > 0 && (deb - cre) > 0.005) dStats[uid].months++;
      dStats[uid].balance = toNum(r[DL.BALANCE]);  // last row overwrites → cumulative total
      var pd = r[DL.PAY_DATE];
      if (pd && cre > 0) {
        var ds = fmtDate(pd);
        if (!dStats[uid].lastPay || ds > dStats[uid].lastPay) dStats[uid].lastPay = ds;
      }
    });
  }

  var duesList = [];
  Object.keys(dStats).forEach(function(uid) {
    var s = dStats[uid];
    if (s.months <= 0 || s.balance <= 0.005) return;  // positive = owes
    duesList.push({
      unitId: uid, ownerName: nameMap[uid] || '',
      phase: phaseMap[uid] || '', block: blockMap[uid] || '',
      balance: Math.round(s.balance * 100) / 100,
      delinqMonths: s.months, lastPayDate: s.lastPay
    });
  });
  duesList.sort(function(a, b) { return b.delinqMonths - a.delinqMonths; });

  return { ok: true, water: waterList, dues: duesList };
}

// ── Meter Changes ─────────────────────────────────────────────
function webapp_getMeterChangeData() {
  var result = { units: [], changes: [], currentYear: 0, currentMonth: '' };

  var mlSh = getSheet_(SH.MASTERLIST);
  if (mlSh && mlSh.getLastRow() > 1) {
    mlSh.getRange(2, 1, mlSh.getLastRow() - 1, 12).getValues().forEach(function(r) {
      var uid = String(r[0] || '').trim();
      if (!uid || String(r[11] || 'Yes').toLowerCase() === 'no') return;
      if (webapp_isCommonAccount_(uid)) return;
      result.units.push({ id: uid, name: ownerName(r), meterNo: String(r[9] || '').trim() });
    });
    result.units.sort(function(a, b) {
      return String(a.id).localeCompare(String(b.id), undefined, { numeric: true });
    });
  }

  var inSh = getSheet_(SH.W_INPUT);
  if (inSh) {
    result.currentYear  = parseInt(inSh.getRange('B3').getValue(), 10) || new Date().getFullYear();
    result.currentMonth = String(inSh.getRange('B4').getValue() || '').trim() ||
                          MONTHS[new Date().getMonth()];
  }

  var mcSh = getSheet_(SH._METER_CHG);
  if (mcSh && mcSh.getLastRow() > 1) {
    mcSh.getRange(2, 1, mcSh.getLastRow() - 1, MC_COLS).getValues().forEach(function(r, i) {
      var uid = String(r[MC.UNIT] || '').trim();
      if (!uid) return;
      result.changes.push({
        sheetRow       : i + 2,
        unitId         : uid,
        dateChanged    : r[MC.DATE]      ? fmtDate(r[MC.DATE]) : '',
        oldMeterNo     : String(r[MC.OLD_METER] || ''),
        oldFinalReading: toNum(r[MC.OLD_FINAL]),
        newMeterNo     : String(r[MC.NEW_METER] || ''),
        newStartReading: toNum(r[MC.NEW_START]),
        billingYear    : r[MC.YEAR]  || '',
        billingMonth   : String(r[MC.MONTH] || ''),
        status         : String(r[MC.STATUS] || 'Pending')
      });
    });
  }

  return { ok: true, data: result };
}

function webapp_saveMeterChange(data) {
  try {
    var uid = String(data.unitId || '').trim();
    if (!uid) return { ok: false, msg: 'Unit ID is required.' };

    var billingYear  = parseInt(data.billingYear, 10);
    var billingMonth = String(data.billingMonth || '').trim();
    if (!billingYear || !billingMonth) return { ok: false, msg: 'Billing year and month are required.' };

    var newMeterNo = String(data.newMeterNo || '').trim();
    if (!newMeterNo) return { ok: false, msg: 'New meter number is required.' };

    // Block if this unit+period already has a billed WL entry
    var wlSh = getSheet_(SH._WL);
    if (wlSh && wlSh.getLastRow() > 1) {
      var wlData = wlSh.getRange(2, 1, wlSh.getLastRow() - 1, 3).getValues();
      for (var i = 0; i < wlData.length; i++) {
        if (String(wlData[i][0] || '').trim() === uid &&
            parseInt(wlData[i][1], 10) === billingYear &&
            String(wlData[i][2] || '').trim() === billingMonth) {
          return { ok: false, msg: uid + ' has already been billed for ' +
                   billingMonth + ' ' + billingYear + '. Cannot log a meter change for a billed period.' };
        }
      }
    }

    var mcSh = getSheet_(SH._METER_CHG);
    if (!mcSh) return { ok: false, msg: '_MeterChanges sheet not found.' };

    var dateChanged = data.dateChanged ? new Date(data.dateChanged) : new Date();
    var oldMeterNo  = String(data.oldMeterNo || '').trim();
    var oldFinal    = toNum(data.oldFinalReading);
    var newStart    = toNum(data.newStartReading);

    mcSh.appendRow([uid, dateChanged, oldMeterNo, oldFinal,
                    newMeterNo, newStart, billingYear, billingMonth, 'Pending']);
    populateWaterInputTable();
    return { ok: true, msg: 'Meter change logged for ' + uid + '.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function webapp_updateMeterChange(sheetRow, data) {
  try {
    var mcSh = getSheet_(SH._METER_CHG);
    if (!mcSh) return { ok: false, msg: '_MeterChanges sheet not found.' };
    if (sheetRow < 2 || sheetRow > mcSh.getLastRow()) return { ok: false, msg: 'Invalid row.' };

    var existing = mcSh.getRange(sheetRow, 1, 1, MC_COLS).getValues()[0];
    if (String(existing[MC.STATUS] || '').trim() !== 'Pending')
      return { ok: false, msg: 'Only Pending changes can be edited.' };

    var dateChanged = data.dateChanged ? new Date(data.dateChanged) : existing[MC.DATE];
    var newMeterNo  = String(data.newMeterNo || '').trim();
    if (!newMeterNo) return { ok: false, msg: 'New meter number is required.' };

    mcSh.getRange(sheetRow, 1, 1, MC_COLS).setValues([[
      String(data.unitId    || existing[MC.UNIT]),
      dateChanged,
      String(data.oldMeterNo || existing[MC.OLD_METER] || ''),
      toNum(data.oldFinalReading),
      newMeterNo,
      toNum(data.newStartReading),
      parseInt(data.billingYear, 10) || existing[MC.YEAR],
      String(data.billingMonth  || existing[MC.MONTH] || ''),
      'Pending'
    ]]);
    return { ok: true, msg: 'Meter change updated.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function webapp_deleteMeterChange(sheetRow) {
  try {
    var mcSh = getSheet_(SH._METER_CHG);
    if (!mcSh) return { ok: false, msg: '_MeterChanges sheet not found.' };
    if (sheetRow < 2 || sheetRow > mcSh.getLastRow()) return { ok: false, msg: 'Invalid row.' };

    var existing = mcSh.getRange(sheetRow, 1, 1, MC_COLS).getValues()[0];
    if (String(existing[MC.STATUS] || '').trim() !== 'Pending')
      return { ok: false, msg: 'Only Pending changes can be deleted.' };

    mcSh.deleteRow(sheetRow);
    return { ok: true, msg: 'Meter change deleted.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}


// ── Data Reset & Import from Drive folders ───────────────────
// The source files use the old per-unit format:
//   Each Drive file = one block spreadsheet.
//   Each sheet/tab  = one unit (named P1B1L1, P2B3L5, GUARDHOUSE, etc.)
//   Each sheet      = side-by-side dues (left) + water (right) ledger table.
// Parsing is handled by lm_parseOldUnitLedgerSheet_() in LedgerMigration.gs.

function webapp_previewImportFolders() {
  try {
    var SOURCE_FOLDERS = [
      { label: 'Phase 1',         folderId: IMPORT_FOLDERS['Phase 1'] },
      { label: 'Phase 2',         folderId: IMPORT_FOLDERS['Phase 2'] },
      { label: 'Common Accounts', folderId: IMPORT_FOLDERS['Common Accounts'] }
    ];
    var results = [];
    var totalUnits = 0;

    SOURCE_FOLDERS.forEach(function(source) {
      var folder;
      try { folder = DriveApp.getFolderById(source.folderId); }
      catch (e) { results.push({ folder: source.label, name: '(cannot open folder)', units: [], skipped: [e.message] }); return; }

      var files = folder.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        var fileInfo = { folder: source.label, name: file.getName(), id: file.getId(), units: [], skipped: [] };

        if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
          fileInfo.skipped.push('Not a Google Sheets file');
          results.push(fileInfo);
          continue;
        }

        var sourceSS;
        try { sourceSS = SpreadsheetApp.openById(file.getId()); }
        catch (e) { fileInfo.skipped.push('Cannot open: ' + e.message); results.push(fileInfo); continue; }

        sourceSS.getSheets().forEach(function(sh) {
          var unitId = lm_normalizeUnitId_(sh.getName());
          if (lm_isBillableLedgerTab_(unitId)) {
            fileInfo.units.push(unitId);
            totalUnits++;
          } else {
            fileInfo.skipped.push(sh.getName());
          }
        });
        results.push(fileInfo);
      }
    });

    return { ok: true, files: results, totalUnits: totalUnits };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function webapp_clearLedgerData() {
  try {
    var targets = [SH._WL, SH._DL, SH.PAY_LOG, SH.RATE_CALC];
    var cleared = [];
    targets.forEach(function(name) {
      var sh = getSheet_(name);
      if (!sh) return;
      var last = sh.getLastRow();
      if (last > 1) {
        sh.getRange(2, 1, last - 1, Math.max(sh.getLastColumn(), 1)).clearContent();
      }
      cleared.push(name);
    });
    return { ok: true, msg: 'Cleared: ' + cleared.join(', ') };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function webapp_importFromFolders() {
  try {
    var SOURCE_FOLDERS = [
      { label: 'Phase 1',         folderId: IMPORT_FOLDERS['Phase 1'],         migrateDues: true,  migrateWater: true },
      { label: 'Phase 2',         folderId: IMPORT_FOLDERS['Phase 2'],         migrateDues: true,  migrateWater: true },
      { label: 'Common Accounts', folderId: IMPORT_FOLDERS['Common Accounts'], migrateDues: false, migrateWater: true }
    ];

    var wlSh = getSheet_(SH._WL);
    var dlSh = getSheet_(SH._DL);
    if (!wlSh || !dlSh) return { ok: false, msg: 'Ledger sheets not found. Run Setup first.' };

    var allWaterRows = [];
    var allDuesRows  = [];
    var reportRows   = [['TIMESTAMP','FOLDER GROUP','FILE NAME','FILE ID','SHEET NAME','UNIT ID','STATUS','DUES ROWS','WATER ROWS','MESSAGE']];

    SOURCE_FOLDERS.forEach(function(source) {
      var folder;
      try { folder = DriveApp.getFolderById(source.folderId); }
      catch (e) { reportRows.push(lm_reportRow_(source.label, '', source.folderId, '', '', 'ERROR', 0, 0, 'Cannot open folder: ' + e.message)); return; }

      var files = folder.getFiles();
      while (files.hasNext()) {
        var file = files.next();

        if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
          reportRows.push(lm_reportRow_(source.label, file.getName(), file.getId(), '', '', 'SKIPPED', 0, 0, 'Not a Google Sheets file.'));
          continue;
        }

        var sourceSS;
        try { sourceSS = SpreadsheetApp.openById(file.getId()); }
        catch (e) { reportRows.push(lm_reportRow_(source.label, file.getName(), file.getId(), '', '', 'ERROR', 0, 0, 'Cannot open: ' + e.message)); continue; }

        sourceSS.getSheets().forEach(function(sheet) {
          var sheetName = sheet.getName();
          var unitId    = lm_normalizeUnitId_(sheetName);

          if (!lm_isBillableLedgerTab_(unitId)) {
            reportRows.push(lm_reportRow_(source.label, file.getName(), file.getId(), sheetName, '', 'SKIPPED', 0, 0, 'Not a billable unit/account tab.'));
            return;
          }

          var isCommon = lm_isCommonAccount_(unitId);
          try {
            var parsed = lm_parseOldUnitLedgerSheet_(sheet, unitId, {
              migrateDues : source.migrateDues && !isCommon,
              migrateWater: source.migrateWater
            });
            Array.prototype.push.apply(allWaterRows, parsed.waterRows);
            Array.prototype.push.apply(allDuesRows,  parsed.duesRows);
            reportRows.push(lm_reportRow_(source.label, file.getName(), file.getId(), sheetName, unitId, 'IMPORTED', parsed.duesRows.length, parsed.waterRows.length, parsed.message));
          } catch (e) {
            reportRows.push(lm_reportRow_(source.label, file.getName(), file.getId(), sheetName, unitId, 'ERROR', 0, 0, e.message));
          }
        });
      }
    });

    if (allWaterRows.length > 0) lm_appendRows_(wlSh, allWaterRows);
    if (allDuesRows.length  > 0) lm_appendRows_(dlSh,  allDuesRows);

    recalcWaterBalances();
    recalcDuesBalances();

    try { lm_writeMigrationReport_(ss_(), reportRows); } catch (e) { /* non-fatal */ }

    return {
      ok: true,
      msg: 'Imported ' + allWaterRows.length + ' water rows and ' + allDuesRows.length + ' dues rows. See Migration Report sheet for details.',
      waterRows: allWaterRows.length,
      duesRows: allDuesRows.length
    };
  } catch (e) { return { ok: false, msg: e.message }; }
}