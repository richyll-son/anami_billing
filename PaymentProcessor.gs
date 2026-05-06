// ============================================================
// AHNHAI Billing System — PaymentProcessor.gs
// Post Payments: apply credits to oldest unpaid bills,
// recalculate running balances, mark rows as Posted.
// ============================================================

function postPayments() {
  var paySh = getSheet_(SH.PAY_LOG);
  if (!paySh || paySh.getLastRow() < 2) {
    alert_('No payment records found in Central Payment Log.'); return;
  }

  // Read all payment log rows (skip header)
  var numRows = paySh.getLastRow() - 1;
  var data    = paySh.getRange(2, 1, numRows, 11).getValues();

  // Col indices (0-based): DATE_REC=0, UNIT=1, OWNER=2, PAY_DATE=3, AMT=4,
  //                         OR=5, TYPE=6, WATER_AMT=7, DUES_AMT=8, REMARKS=9, STATUS=10
  var unposted = [];
  data.forEach(function(r, i) {
    var status = String(r[10]).trim().toLowerCase();
    var uid    = String(r[1]).trim();
    if (!uid || status === 'posted') return;
    unposted.push({
      rowIdx   : i + 2,        // 1-based sheet row
      unitId   : uid,
      payDate  : r[3],
      totalAmt : toNum(r[4]),
      orNum    : r[5],
      payType  : String(r[6]).trim(),
      waterAmt : toNum(r[7]),
      duesAmt  : toNum(r[8]),
      remarks  : String(r[9]).trim()
    });
  });

  if (unposted.length === 0) {
    alert_('No unposted payments found. All rows are already marked Posted.'); return;
  }

  if (!confirm_('Post Payments', 'About to post ' + unposted.length + ' payment(s). Continue?'))
    return;

  // Load ledger data once (batch)
  var wlSh   = getSheet_(SH._WL);
  var dlSh   = getSheet_(SH._DL);
  var wlData = (wlSh && wlSh.getLastRow() > 1)
    ? wlSh.getRange(2, 1, wlSh.getLastRow() - 1, WL_COLS).getValues() : [];
  var dlData = (dlSh && dlSh.getLastRow() > 1)
    ? dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues() : [];

  var posted  = 0;
  var errors  = 0;
  var wlDirty = false;
  var dlDirty = false;

  unposted.forEach(function(p) {
    try {
      if (p.payType === 'Water' || p.payType === 'Both') {
        var wAmt = (p.payType === 'Both') ? p.waterAmt : p.totalAmt;
        if (wAmt > 0) {
          _applyWaterCredit(wlData, p.unitId, wAmt, p.payDate, p.orNum, p.remarks);
          wlDirty = true;
        }
      }
      if (p.payType === 'Dues' || p.payType === 'Both') {
        var dAmt = (p.payType === 'Both') ? p.duesAmt : p.totalAmt;
        if (dAmt > 0) {
          _applyDuesCredit(dlData, p.unitId, dAmt, p.payDate, p.orNum, p.remarks);
          dlDirty = true;
        }
      }

      // Mark row as Posted (green)
      paySh.getRange(p.rowIdx, 11)
        .setValue('Posted')
        .setBackground('#c8e6c9').setFontColor('#1b5e20').setFontWeight('bold');
      posted++;

    } catch (err) {
      Logger.log('postPayments error [' + p.unitId + ']: ' + err.message);
      paySh.getRange(p.rowIdx, 11)
        .setValue('Error: ' + err.message)
        .setBackground('#ffcdd2').setFontColor('#b71c1c');
      errors++;
    }
  });

  // Write modified ledger data back in one batch
  if (wlDirty && wlData.length > 0) {
    wlSh.getRange(2, 1, wlData.length, WL_COLS).setValues(wlData);
  }
  if (dlDirty && dlData.length > 0) {
    dlSh.getRange(2, 1, dlData.length, DL_COLS).setValues(dlData);
  }

  // Recalculate all running balances
  if (wlDirty) recalcWaterBalances();
  if (dlDirty) recalcDuesBalances();

  // Refresh Unit Ledger display if a unit is selected
  var ulSh = getSheet_(SH.UNIT_LEDGER);
  if (ulSh) {
    var sel = ulSh.getRange('B1').getValue();
    if (sel && sel !== '-- Select Unit ID --') refreshUnitLedger(sel);
  }

  refreshMonthlySummary();

  alert_('Payment posting complete!\n\nPosted: ' + posted + '\nErrors: ' + errors);
}

// ── Apply credit to oldest unpaid water bills (in-memory) ────
function _applyWaterCredit(wlData, unitId, amount, payDate, orNum, remarks) {
  var remaining = amount;
  var payDateStr = fmtDate(payDate || new Date());

  for (var i = 0; i < wlData.length && remaining > 0; i++) {
    if (wlData[i][WL.UNIT] !== unitId) continue;

    var debit   = toNum(wlData[i][WL.DEBIT]);
    var pen     = toNum(wlData[i][WL.PENALTY]);
    var addon   = toNum(wlData[i][WL.ADDON]);
    var credit  = toNum(wlData[i][WL.CREDIT]);
    var billDue = debit + pen + addon;
    var unpaid  = Math.max(0, billDue - credit);

    if (unpaid <= 0) continue;

    var apply = Math.min(remaining, unpaid);
    wlData[i][WL.CREDIT]   = Math.round((credit + apply) * 100) / 100;
    wlData[i][WL.PAY_DATE] = payDateStr;

    if (!wlData[i][WL.OR])      wlData[i][WL.OR]      = orNum;
    if (!wlData[i][WL.REMARKS]) wlData[i][WL.REMARKS]  = remarks || ('Paid: ' + payDateStr);

    remaining = Math.round((remaining - apply) * 100) / 100;
  }

  // If there's still remaining credit (overpayment), note it on the last bill for this unit
  if (remaining > 0) {
    for (var j = wlData.length - 1; j >= 0; j--) {
      if (wlData[j][WL.UNIT] === unitId) {
        var note = 'Overpayment: ₱' + fmt2(remaining);
        wlData[j][WL.REMARKS] = wlData[j][WL.REMARKS]
          ? wlData[j][WL.REMARKS] + ' | ' + note : note;
        break;
      }
    }
  }
}

// ── Apply credit to oldest unpaid dues bills (in-memory) ─────
function _applyDuesCredit(dlData, unitId, amount, payDate, orNum, remarks) {
  var remaining  = amount;
  var payDateStr = fmtDate(payDate || new Date());

  for (var i = 0; i < dlData.length && remaining > 0; i++) {
    if (dlData[i][DL.UNIT] !== unitId) continue;

    var debit  = toNum(dlData[i][DL.DEBIT]);
    var credit = toNum(dlData[i][DL.CREDIT]);
    var unpaid = Math.max(0, debit - credit);
    if (unpaid <= 0) continue;

    var apply = Math.min(remaining, unpaid);
    dlData[i][DL.CREDIT]   = Math.round((credit + apply) * 100) / 100;
    dlData[i][DL.PAY_DATE] = payDateStr;

    if (!dlData[i][DL.OR])      dlData[i][DL.OR]      = orNum;
    if (!dlData[i][DL.REMARKS]) dlData[i][DL.REMARKS]  = remarks || ('Paid: ' + payDateStr);

    remaining = Math.round((remaining - apply) * 100) / 100;
  }
}
