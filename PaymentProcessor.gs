// ============================================================
// AHNHAI Billing System — PaymentProcessor.gs
// Post Payments: apply water credit to the most recent bill,
// dues credit oldest-first; recalculate running balances.
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

// ── Apply credit to the most recent water bill only (in-memory) ──
// The entire payment goes to the last ledger row for the unit.
// If the payment exceeds the balance on that row, the surplus stays
// as credit there and carries forward as a negative balance to the
// next billing period (handled automatically by recalcWaterBalances).
function _applyWaterCredit(wlData, unitId, amount, payDate, orNum, remarks) {
  var payDateStr = fmtDate(payDate || new Date());
  var orStr      = String(orNum || '').trim();
  var normId     = String(unitId || '').trim();

  // Find the last (most recent) row for this unit
  var lastIdx = -1;
  for (var i = wlData.length - 1; i >= 0; i--) {
    if (String(wlData[i][WL.UNIT] || '').trim() === normId) {
      lastIdx = i;
      break;
    }
  }

  if (lastIdx === -1) return; // no ledger entry for this unit yet

  // Apply the full payment to the most recent row
  wlData[lastIdx][WL.CREDIT]   = Math.round((toNum(wlData[lastIdx][WL.CREDIT]) + amount) * 100) / 100;
  wlData[lastIdx][WL.PAY_DATE] = payDateStr;

  var existOr = String(wlData[lastIdx][WL.OR] || '').trim();
  wlData[lastIdx][WL.OR] = (existOr && existOr !== orStr)
    ? existOr + ' / ' + orStr : orStr;

  var autoRemark = orStr ? 'Pmt – ' + orStr : 'Pmt';
  var existRem   = String(wlData[lastIdx][WL.REMARKS] || '').trim();
  wlData[lastIdx][WL.REMARKS] = existRem ? existRem + ' | ' + autoRemark : autoRemark;
}

// ── Apply credit to dues (in-memory) ─────────────────────────
// If targetMonth + targetYear are given, applies the full amount to that
// specific row (partial payment allowed). Falls back to oldest-first
// if the target row is not found or no target is specified.
function _applyDuesCredit(dlData, unitId, amount, payDate, orNum, remarks, targetMonth, targetYear) {
  var payDateStr  = fmtDate(payDate || new Date());
  var orStr       = String(orNum || '').trim();
  var normId      = String(unitId || '').trim();

  // ── Targeted: apply to a specific billing month ─────────────
  if (targetMonth && targetYear) {
    var normMonth = String(targetMonth).trim();
    var normYear  = parseInt(targetYear, 10);
    for (var t = 0; t < dlData.length; t++) {
      if (String(dlData[t][DL.UNIT]  || '').trim() !== normId)    continue;
      if (String(dlData[t][DL.MONTH] || '').trim() !== normMonth)  continue;
      if (toNum(dlData[t][DL.YEAR])  !== normYear)                 continue;

      dlData[t][DL.CREDIT]   = Math.round((toNum(dlData[t][DL.CREDIT]) + amount) * 100) / 100;
      dlData[t][DL.PAY_DATE] = payDateStr;

      var eo = String(dlData[t][DL.OR] || '').trim();
      dlData[t][DL.OR] = (eo && eo !== orStr) ? eo + ' / ' + orStr : orStr;

      var autoRemark = orStr ? 'Pmt – ' + orStr : 'Pmt';
      var er = String(dlData[t][DL.REMARKS] || '').trim();
      dlData[t][DL.REMARKS] = er ? er + ' | ' + autoRemark : autoRemark;
      return;
    }
    // Target row not found — fall through to oldest-first
  }

  // ── Oldest-first (default) ───────────────────────────────────
  var toApply = [];
  var tempRem = amount;
  for (var i = 0; i < dlData.length && tempRem > 0; i++) {
    if (String(dlData[i][DL.UNIT] || '').trim() !== normId) continue;
    var debit  = toNum(dlData[i][DL.DEBIT]);
    var credit = toNum(dlData[i][DL.CREDIT]);
    var unpaid = Math.max(0, debit - credit);
    if (unpaid <= 0) continue;
    var apply = Math.round(Math.min(tempRem, unpaid) * 100) / 100;
    toApply.push({ idx: i, apply: apply });
    tempRem = Math.round((tempRem - apply) * 100) / 100;
  }

  var total = toApply.length;

  toApply.forEach(function(entry, n) {
    var i      = entry.idx;
    dlData[i][DL.CREDIT]   = Math.round((toNum(dlData[i][DL.CREDIT]) + entry.apply) * 100) / 100;
    dlData[i][DL.PAY_DATE] = payDateStr;

    var existOr = String(dlData[i][DL.OR] || '').trim();
    dlData[i][DL.OR] = (existOr && existOr !== orStr) ? existOr + ' / ' + orStr : orStr;

    var autoRemark = 'Pmt ' + (n + 1) + ' of ' + total + ' – ' + orStr;
    var existRem   = String(dlData[i][DL.REMARKS] || '').trim();
    dlData[i][DL.REMARKS] = existRem ? existRem + ' | ' + autoRemark : autoRemark;
  });

  // Overpayment: carry excess on last row
  if (tempRem > 0) {
    for (var j = dlData.length - 1; j >= 0; j--) {
      if (String(dlData[j][DL.UNIT] || '').trim() === normId) {
        dlData[j][DL.CREDIT] = Math.round((toNum(dlData[j][DL.CREDIT]) + tempRem) * 100) / 100;
        var note = 'Advance ₱' + fmt2(tempRem) + ' – ' + orStr;
        var er2  = String(dlData[j][DL.REMARKS] || '').trim();
        dlData[j][DL.REMARKS] = er2 ? er2 + ' | ' + note : note;
        dlData[j][DL.PAY_DATE] = payDateStr;
        var eo2 = String(dlData[j][DL.OR] || '').trim();
        dlData[j][DL.OR] = (eo2 && eo2 !== orStr) ? eo2 + ' / ' + orStr : orStr;
        break;
      }
    }
  }
}
