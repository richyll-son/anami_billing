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
    if (!monthlyMap[key]) {
      monthlyMap[key] = { key: key, label: webapp_periodLabel_(year, month), payments: 0, receivables: 0, overpayments: 0, waterPayments: 0, duesPayments: 0, waterReceivables: 0, duesReceivables: 0 };
    }
    monthlyMap[key][field] += toNum(value);
  }

  var waterLast = {};
  if (wlSh && wlSh.getLastRow() > 1) {
    wlSh.getRange(2, 1, wlSh.getLastRow() - 1, WL_COLS).getValues().forEach(function(r) {
      var uid = String(r[WL.UNIT] || '').trim();
      var t = ensureUnit(uid); if (!t) return;
      var credit = toNum(r[WL.CREDIT]);
      var bal = toNum(r[WL.BALANCE]);
      t.waterPayments += credit;
      t.payments += credit;
      waterLast[uid] = { balance: bal, year: r[WL.YEAR], month: r[WL.MONTH] };
      addMonthly(r[WL.YEAR], r[WL.MONTH], 'payments', credit);
      addMonthly(r[WL.YEAR], r[WL.MONTH], 'waterPayments', credit);
    });
  }

  var duesLast = {};
  if (dlSh && dlSh.getLastRow() > 1) {
    dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues().forEach(function(r) {
      var uid = String(r[DL.UNIT] || '').trim();
      var t = ensureUnit(uid); if (!t) return;
      var credit = toNum(r[DL.CREDIT]);
      var bal = toNum(r[DL.BALANCE]);
      t.duesPayments += credit;
      t.payments += credit;
      duesLast[uid] = { balance: bal, year: r[DL.YEAR], month: r[DL.MONTH] };
      addMonthly(r[DL.YEAR], r[DL.MONTH], 'payments', credit);
      addMonthly(r[DL.YEAR], r[DL.MONTH], 'duesPayments', credit);
    });
  }

  Object.keys(byUnit).forEach(function(uid) {
    var t = byUnit[uid];
    var wBal = waterLast[uid] ? toNum(waterLast[uid].balance) : 0;
    var dBal = duesLast[uid] ? toNum(duesLast[uid].balance) : 0;

    if (wBal > 0) { t.waterReceivables += wBal; t.receivables += wBal; }
    if (wBal < 0) { t.waterOverpayments += Math.abs(wBal); t.overpayments += Math.abs(wBal); }
    if (dBal > 0) { t.duesReceivables += dBal; t.receivables += dBal; }
    if (dBal < 0) { t.duesOverpayments += Math.abs(dBal); t.overpayments += Math.abs(dBal); }

    var wLast = waterLast[uid];
    if (wLast) {
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

    // Dashboard money summaries intentionally exclude common accounts.
    if (group.common) return;

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

  var monthlyRows = Object.keys(monthlyMap).sort().map(function(k) {
    return webapp_roundTotals_(monthlyMap[k]);
  }).slice(-12);

  return {
    unitCount: unitCount,
    total: webapp_roundTotals_(total),
    phases: phaseRows,
    blocks: blockRows,
    monthly: monthlyRows
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
    mcwdFrom: '', mcwdTo: '', mcwdAmt: 0,
    elecFrom: '', elecTo: '', elecAmt: 0,
    manpower: MANPOWER_DEF
  };

  if (sh) {
    var get = function(cell) { return sh.getRange(cell).getValue(); };
    config.year     = get('B3') || config.year;
    config.month    = get('B4') || config.month;
    config.mcwdFrom = get('B5') ? fmtDate(get('B5')) : '';
    config.mcwdTo   = get('B6') ? fmtDate(get('B6')) : '';
    config.mcwdAmt  = toNum(get('B7'));
    config.elecFrom = get('B8') ? fmtDate(get('B8')) : '';
    config.elecTo   = get('B9') ? fmtDate(get('B9')) : '';
    config.elecAmt  = toNum(get('B10'));
    config.manpower = toNum(get('B11')) || MANPOWER_DEF;
  }

  var table = [];
  if (sh && sh.getLastRow() >= INPUT_TABLE_START) {
    sh.getRange(INPUT_TABLE_START, 1, sh.getLastRow() - INPUT_TABLE_START + 1, 5).getValues()
      .forEach(function(r) {
        if (!r[0]) return;
        table.push({
          unitId   : r[0],
          meterNo  : r[1] || '',
          ownerName: r[2] || '',
          cur      : r[3] !== '' && r[3] !== null ? toNum(r[3]) : '',
          prev     : r[4] !== '' && r[4] !== null ? toNum(r[4]) : ''
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

  var mcwdAmt  = toNum(config.mcwdAmt);
  var elecAmt  = toNum(config.elecAmt);
  var manpower = toNum(config.manpower) || MANPOWER_DEF;
  var totalExpense = mcwdAmt + elecAmt + manpower;

  var readingObjs = readings.map(function(r) {
    var prev = toNum(r.prev);
    var cur  = toNum(r.cur);
    var cons = Math.max(0, cur - prev);
    return {
      unitId: String(r.unitId).trim(),
      meterNo: r.meterNo || '',
      prev: prev,
      cur: cur,
      cons: cons
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
  var addon          = 0;

  var newWlRows = [];
  var newDlRows = [];
  var newWsRows = [];

  readingObjs.forEach(function(r) {
    var waterRaw = r.cons * rate;
    var debit = r.cons === 0 ? 0 : Math.max(MIN_WATER_BILL, waterRaw);

    var parsed = parseUID(r.unitId);
    var billNo = parsed
      ? buildBillNum(year, month, parsed.phase, parsed.block, parsed.lot)
      : webapp_buildCommonBillNum_(year, month, r.unitId);

    newWlRows.push([
      r.unitId, year, monthName, billDateStr,
      '', presentDateStr,
      r.prev, r.cur, rate, dueDateStr,
      penalty, debit, 0, 0, addon,
      '', '', billNo, ''
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

  sh.getRange(sh.getLastRow() + 1, 1, 1, 11).setValues([[
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
  return { ok: true };
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

  var wlSh = getSheet_(SH._WL);
  if (wlSh && wlSh.getLastRow() > 1) {
    wlSh.getRange(2, 1, wlSh.getLastRow() - 1, WL_COLS).getValues().forEach(function(r) {
      if (r[WL.UNIT] !== unitId) return;
      result.water.push({
        year       : r[WL.YEAR],
        month      : r[WL.MONTH],
        billDate   : r[WL.BILL_DATE],
        dueDate    : r[WL.DUE_DATE],
        prevDate   : r[WL.PREV_DATE],
        presentDate: r[WL.PRESENT_DATE],
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
        payDate    : r[WL.PAY_DATE],
        billNo     : r[WL.BILL_NO]
      });
    });
  }

  var dlSh = getSheet_(SH._DL);
  if (dlSh && dlSh.getLastRow() > 1) {
    dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues().forEach(function(r) {
      if (r[DL.UNIT] !== unitId) return;
      result.dues.push({
        year   : r[DL.YEAR],
        month  : r[DL.MONTH],
        debit  : r[DL.DEBIT],
        credit : r[DL.CREDIT],
        balance: r[DL.BALANCE],
        payDate: r[DL.PAY_DATE],
        or     : r[DL.OR],
        remarks: r[DL.REMARKS]
      });
    });
  }

  return result;
}

// ── Monthly Summary ───────────────────────────────────────────
function webapp_getMonthlySummary() {
  var sh = getSheet_(SH.SUMMARY);
  if (!sh || sh.getLastRow() < 5) return { rows: [], period: '' };

  var period = String(sh.getRange('A3').getValue());
  var rows = sh.getRange(5, 1, Math.max(1, sh.getLastRow() - 5), 16).getValues()
    .filter(function(r) { return r[0] && String(r[0]) !== 'TOTALS'; })
    .map(function(r) {
      return {
        unitId      : r[0],  name        : r[1],  meter      : r[2],
        cons        : r[3],  rate        : r[4],  waterBill  : r[5],
        penalty     : r[6],  waterArrears: r[7],  addon      : r[8],
        totalWater  : r[9],  assocDue    : r[10], duesArrears: r[11],
        totalDues   : r[12], grandTotal  : r[13], status     : r[14],
        dueDate     : r[15],
        group       : webapp_isCommonAccount_(r[0]) ? 'common' : (String(r[0]).indexOf('P1') === 0 ? '1' : (String(r[0]).indexOf('P2') === 0 ? '2' : 'other'))
      };
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