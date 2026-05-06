// ============================================================
// AHNHAI Billing System — SummaryGenerator.gs
// Monthly Summary generator + print/dashboard-friendly helpers
// ============================================================

function refreshMonthlySummary() {
  var ss = ss_();

  var mlSh = getSheet_(SH.MASTERLIST);
  var wlSh = getSheet_(SH._WL);
  var dlSh = getSheet_(SH._DL);
  var sumSh = getSheet_(SH.SUMMARY);

  if (!mlSh || !wlSh || !dlSh || !sumSh) {
    throw new Error('Required sheets are missing. Please run Initial Setup first.');
  }

  var masterRows = mlSh.getLastRow() > 1
    ? mlSh.getRange(2, 1, mlSh.getLastRow() - 1, 12).getValues()
    : [];

  var waterRows = wlSh.getLastRow() > 1
    ? wlSh.getRange(2, 1, wlSh.getLastRow() - 1, WL_COLS).getValues()
    : [];

  var duesRows = dlSh.getLastRow() > 1
    ? dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues()
    : [];

  var units = [];

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
  });

  /*
    Include any unit/account that exists in ledgers even if it is missing
    from Masterlist, so migrated old ledger data still appears.
  */
  var unitExists = {};
  units.forEach(function(u) {
    unitExists[u.unitId] = true;
  });

  waterRows.forEach(function(r) {
    var uid = String(r[WL.UNIT] || '').trim();
    if (!uid || unitExists[uid]) return;

    units.push({
      unitId: uid,
      name: sg_commonName_(uid),
      meter: '',
      common: sg_isCommonAccount_(uid)
    });

    unitExists[uid] = true;
  });

  duesRows.forEach(function(r) {
    var uid = String(r[DL.UNIT] || '').trim();
    if (!uid || unitExists[uid]) return;

    units.push({
      unitId: uid,
      name: sg_commonName_(uid),
      meter: '',
      common: sg_isCommonAccount_(uid)
    });

    unitExists[uid] = true;
  });

  units.sort(function(a, b) {
    if (a.common && !b.common) return 1;
    if (!a.common && b.common) return -1;
    return String(a.unitId).localeCompare(String(b.unitId), undefined, { numeric: true });
  });

  var latestWaterByUnit = {};
  var latestDuesByUnit = {};

  waterRows.forEach(function(r, idx) {
    var uid = String(r[WL.UNIT] || '').trim();
    if (!uid) return;

    latestWaterByUnit[uid] = {
      row: r,
      index: idx
    };
  });

  duesRows.forEach(function(r, idx) {
    var uid = String(r[DL.UNIT] || '').trim();
    if (!uid) return;

    latestDuesByUnit[uid] = {
      row: r,
      index: idx
    };
  });

  var latestPeriod = sg_getLatestPeriod_(waterRows, duesRows);

  var out = [];

  out.push([
    'MONTHLY BILLING SUMMARY'
  ]);

  out.push([
    'Generated: ' + fmtDate(new Date())
  ]);

  out.push([
    latestPeriod || ''
  ]);

  out.push([
    'Unit',
    'Homeowner',
    'Meter',
    'Cons (m³)',
    'Rate',
    'Water Bill',
    'Penalty',
    'Arrears(W)',
    'Add-On',
    'Total Water',
    'Assoc Due',
    'Arrears(D)',
    'Total Dues',
    'Grand Total',
    'Status',
    'Due Date'
  ]);

  var totals = {
    cons: 0,
    waterBill: 0,
    penalty: 0,
    waterArrears: 0,
    addon: 0,
    totalWater: 0,
    assocDue: 0,
    duesArrears: 0,
    totalDues: 0,
    grandTotal: 0
  };

  units.forEach(function(u) {
    var wInfo = latestWaterByUnit[u.unitId];
    var dInfo = latestDuesByUnit[u.unitId];

    var w = wInfo ? wInfo.row : null;
    var d = dInfo ? dInfo.row : null;

    var prevRdg = w ? toNum(w[WL.PREV_RDG]) : 0;
    var curRdg = w ? toNum(w[WL.CUR_RDG]) : 0;
    var cons = w ? Math.max(0, curRdg - prevRdg) : 0;

    var rate = w ? toNum(w[WL.RATE]) : 0;
    var waterBill = w ? toNum(w[WL.DEBIT]) : 0;
    var penalty = w ? toNum(w[WL.PENALTY]) : 0;
    var addon = w ? toNum(w[WL.ADDON]) : 0;
    var waterBalance = w ? toNum(w[WL.BALANCE]) : 0;

    /*
      Since migrated ledgers carry running balances copied exactly,
      summary arrears should use positive latest balance only.
      Negative balance is overpayment and should not inflate receivables.
    */
    var waterArrears = waterBalance > 0 ? waterBalance : 0;
    var totalWater = waterArrears;

    var duesDebit = d ? toNum(d[DL.DEBIT]) : 0;
    var duesBalance = d ? toNum(d[DL.BALANCE]) : 0;

    /*
      Common accounts do not have monthly dues.
    */
    var assocDue = u.common ? 0 : duesDebit;
    var duesArrears = u.common ? 0 : (duesBalance > 0 ? duesBalance : 0);
    var totalDues = duesArrears;

    var grandTotal = totalWater + totalDues;

    var status = grandTotal <= 0 ? 'Paid' : 'Unpaid';
    if (grandTotal > 0) {
      var hasAnyCredit =
        (w && toNum(w[WL.CREDIT]) > 0) ||
        (d && toNum(d[DL.CREDIT]) > 0);

      status = hasAnyCredit ? 'Partial' : 'Unpaid';
    }

    var dueDate = w ? w[WL.DUE_DATE] : '';

    out.push([
      u.unitId,
      u.name,
      u.meter,
      cons,
      rate,
      waterBill,
      penalty,
      waterArrears,
      addon,
      totalWater,
      assocDue,
      duesArrears,
      totalDues,
      grandTotal,
      status,
      dueDate
    ]);

    totals.cons += cons;
    totals.waterBill += waterBill;
    totals.penalty += penalty;
    totals.waterArrears += waterArrears;
    totals.addon += addon;
    totals.totalWater += totalWater;
    totals.assocDue += assocDue;
    totals.duesArrears += duesArrears;
    totals.totalDues += totalDues;
    totals.grandTotal += grandTotal;
  });

  out.push([
    'TOTALS',
    '',
    '',
    totals.cons,
    '',
    totals.waterBill,
    totals.penalty,
    totals.waterArrears,
    totals.addon,
    totals.totalWater,
    totals.assocDue,
    totals.duesArrears,
    totals.totalDues,
    totals.grandTotal,
    '',
    ''
  ]);

  sumSh.clearContents();
  sumSh.clearFormats();

  sumSh.getRange(1, 1, out.length, 16).setValues(out);

  sumSh.getRange(1, 1, 1, 16)
    .merge()
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#1565c0')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  sumSh.getRange(2, 1, 1, 16)
    .merge()
    .setFontStyle('italic')
    .setHorizontalAlignment('center');

  sumSh.getRange(3, 1, 1, 16)
    .merge()
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sumSh.getRange(4, 1, 1, 16)
    .setFontWeight('bold')
    .setBackground('#263238')
    .setFontColor('#ffffff');

  if (out.length > 5) {
    sumSh.getRange(5, 4, out.length - 4, 1).setNumberFormat('#,##0');
    sumSh.getRange(5, 5, out.length - 4, 10).setNumberFormat('₱#,##0.00;[Red](₱#,##0.00)');
  }

  sumSh.getRange(out.length, 1, 1, 16)
    .setFontWeight('bold')
    .setBackground('#e3f2fd');

  sumSh.setFrozenRows(4);
  sumSh.autoResizeColumns(1, 16);

  return {
    ok: true,
    rows: Math.max(0, out.length - 5),
    period: latestPeriod
  };
}


function sg_getLatestPeriod_(waterRows, duesRows) {
  var last = null;

  waterRows.forEach(function(r) {
    var y = parseInt(r[WL.YEAR], 10);
    var m = getMonthNum(r[WL.MONTH]);

    if (y && m) {
      var rank = y * 100 + m;
      if (!last || rank > last.rank) {
        last = {
          rank: rank,
          label: r[WL.MONTH] + ' ' + y
        };
      }
    }
  });

  duesRows.forEach(function(r) {
    var y = parseInt(r[DL.YEAR], 10);
    var m = getMonthNum(r[DL.MONTH]);

    if (y && m) {
      var rank = y * 100 + m;
      if (!last || rank > last.rank) {
        last = {
          rank: rank,
          label: r[DL.MONTH] + ' ' + y
        };
      }
    }
  });

  return last ? last.label : '';
}


function sg_isCommonAccount_(unitId) {
  var v = String(unitId || '').trim().toUpperCase().replace(/\s+/g, '');

  return [
    'GUARDHOUSE',
    'CLUBHOUSE',
    'CHAPEL'
  ].indexOf(v) !== -1;
}


function sg_commonName_(unitId) {
  var v = String(unitId || '').trim().toUpperCase().replace(/\s+/g, '');

  if (v === 'GUARDHOUSE') return 'Guardhouse';
  if (v === 'CLUBHOUSE') return 'Clubhouse';
  if (v === 'CHAPEL') return 'Chapel';

  return '';
}


/**
 * Optional menu wrapper, if your existing menu calls this.
 */
function refreshSummary() {
  return refreshMonthlySummary();
}