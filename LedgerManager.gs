// ============================================================
// AHNHAI Billing System — LedgerManager.gs
// Unit Ledger: populate display when dropdown changes,
// write-back to hidden stores when admin edits a cell.
// ============================================================

// Called by onEdit when the Unit Ledger sheet changes
function handleUnitLedgerEdit(e) {
  var row = e.range.getRow();
  var col = e.range.getColumn();
  var sh  = e.range.getSheet();

  // ── Dropdown changed → refresh display ───────────────────
  if (row === 1 && col === 2) {
    var uid = sh.getRange('B1').getValue();
    if (uid && uid !== '-- Select Unit ID --') {
      refreshUnitLedger(uid);
    }
    return;
  }

  // ── Data area edited → write back to hidden ledger ────────
  if (row < UL_DATA_ROW) return;

  var dispRowIdx = row - UL_DATA_ROW; // 0-based data row in display

  if (col >= UL_WATER_COL && col <= UL_WATER_COL + 14) {
    var dispColOffset = col - UL_WATER_COL + 1; // 1-based within water section
    _writeBackWater(sh, dispRowIdx, dispColOffset, e.value);

  } else if (col >= UL_DUES_COL && col <= UL_DUES_COL + 6) {
    var dispDuesOffset = col - UL_DUES_COL + 1; // 1-based within dues section
    _writeBackDues(sh, dispRowIdx, dispDuesOffset, e.value);
  }
}

// ── Refresh Unit Ledger display for a given unit ─────────────
function refreshUnitLedger(unitId) {
  var sh = getSheet_(SH.UNIT_LEDGER);
  if (!sh) return;
  if (!unitId) unitId = sh.getRange('B1').getValue();
  if (!unitId || unitId === '-- Select Unit ID --') return;

  // Set homeowner name
  var mlSh = getSheet_(SH.MASTERLIST);
  if (mlSh && mlSh.getLastRow() > 1) {
    var mlData = mlSh.getDataRange().getValues();
    var mlRow  = getMLRow(mlData, unitId);
    sh.getRange('B2').setValue(mlRow ? ownerName(mlRow) : '');
  }

  // Clear current data area (cols 1-23, from data row to last row)
  var lastRow = sh.getLastRow();
  if (lastRow >= UL_DATA_ROW) {
    sh.getRange(UL_DATA_ROW, 1, lastRow - UL_DATA_ROW + 1, UL_DUES_COL + 6)
      .clearContent().clearFormat();
  }

  // ── Load water ledger rows for this unit ─────────────────
  var wlSh   = getSheet_(SH._WL);
  var wlRows = [];
  if (wlSh && wlSh.getLastRow() > 1) {
    wlSh.getRange(2, 1, wlSh.getLastRow() - 1, WL_COLS).getValues()
      .forEach(function(r) {
        if (r[WL.UNIT] !== unitId) return;
        // Display order: PayDate|BillDate|PrevDate|PresentDate|PrevRdg|CurRdg|Rate|DueDate|
        //                Penalty|Debit|Credit|Balance|Addon|OR|Remarks
        wlRows.push([
          r[WL.PAY_DATE],
          r[WL.BILL_DATE],
          r[WL.PREV_DATE],
          r[WL.PRESENT_DATE],
          r[WL.PREV_RDG],
          r[WL.CUR_RDG],
          r[WL.RATE],
          r[WL.DUE_DATE],
          r[WL.PENALTY],
          r[WL.DEBIT],
          r[WL.CREDIT],
          r[WL.BALANCE],
          r[WL.ADDON],
          r[WL.OR],
          r[WL.REMARKS]
        ]);
      });
  }

  if (wlRows.length > 0) {
    var wlRange = sh.getRange(UL_DATA_ROW, UL_WATER_COL, wlRows.length, 15);
    wlRange.setValues(wlRows);

    // Colour Balance column based on value
    for (var i = 0; i < wlRows.length; i++) {
      var balCell = sh.getRange(UL_DATA_ROW + i, UL_WATER_COL + 11); // col 12
      var bal     = toNum(wlRows[i][11]);
      balCell.setBackground(bal <= 0 ? '#c8e6c9' : '#ffcdd2');
    }
  }

  // ── Load dues ledger rows for this unit ───────────────────
  var dlSh   = getSheet_(SH._DL);
  var dlRows = [];
  if (dlSh && dlSh.getLastRow() > 1) {
    dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues()
      .forEach(function(r) {
        if (r[DL.UNIT] !== unitId) return;
        // Display: PayDate|Month|Debit|Credit|Balance|OR|Remarks
        dlRows.push([
          r[DL.PAY_DATE],
          r[DL.MONTH],
          r[DL.DEBIT],
          r[DL.CREDIT],
          r[DL.BALANCE],
          r[DL.OR],
          r[DL.REMARKS]
        ]);
      });
  }

  if (dlRows.length > 0) {
    var dlRange = sh.getRange(UL_DATA_ROW, UL_DUES_COL, dlRows.length, 7);
    dlRange.setValues(dlRows);

    for (var j = 0; j < dlRows.length; j++) {
      var dBalCell = sh.getRange(UL_DATA_ROW + j, UL_DUES_COL + 4); // col 5 within dues
      var dBal     = toNum(dlRows[j][4]);
      dBalCell.setBackground(dBal <= 0 ? '#c8e6c9' : '#ffcdd2');
    }
  }
}

// ── Write-back: water ledger cell edited in display ──────────
function _writeBackWater(sh, dispRowIdx, dispColOffset, newValue) {
  var uid = sh.getRange('B1').getValue();
  if (!uid) return;

  var wlSh = getSheet_(SH._WL);
  if (!wlSh || wlSh.getLastRow() < 2) return;

  var wlData = wlSh.getRange(2, 1, wlSh.getLastRow() - 1, WL_COLS).getValues();
  var count  = 0;

  for (var i = 0; i < wlData.length; i++) {
    if (wlData[i][WL.UNIT] !== uid) continue;
    if (count < dispRowIdx) { count++; continue; }

    // Map display column → _WaterLedger column index
    var wlColIdx = UL_WATER_MAP[dispColOffset];
    if (wlColIdx === null || wlColIdx === undefined) return;
    if (wlColIdx === -1) return; // read-only (Balance)

    // Write the single cell
    wlSh.getRange(i + 2, wlColIdx + 1).setValue(newValue);

    // Trigger balance recalc if a financial column changed
    if (dispColOffset === 9 || dispColOffset === 10 || dispColOffset === 13) {
      // Debit, Credit, or Addon changed
      recalcWaterBalances();
      // Silently refresh (re-colour balance cells only, avoid full re-render loop)
      _recolourWaterBalances(sh, uid);
    }
    return;
  }
}

// ── Write-back: dues ledger cell edited in display ────────────
function _writeBackDues(sh, dispRowIdx, dispDuesOffset, newValue) {
  var uid = sh.getRange('B1').getValue();
  if (!uid) return;

  var dlSh = getSheet_(SH._DL);
  if (!dlSh || dlSh.getLastRow() < 2) return;

  var dlData = dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues();
  var count  = 0;

  for (var i = 0; i < dlData.length; i++) {
    if (dlData[i][DL.UNIT] !== uid) continue;
    if (count < dispRowIdx) { count++; continue; }

    var dlColIdx = UL_DUES_MAP[dispDuesOffset];
    if (dlColIdx === null || dlColIdx === undefined) return;
    if (dlColIdx === -1) return; // read-only (Balance)

    dlSh.getRange(i + 2, dlColIdx + 1).setValue(newValue);

    if (dispDuesOffset === 3 || dispDuesOffset === 4) {
      // Debit or Credit changed
      recalcDuesBalances();
      _recolourDuesBalances(sh, uid);
    }
    return;
  }
}

// ── Re-colour balance cells after in-place recalc ────────────
function _recolourWaterBalances(sh, uid) {
  var wlSh = getSheet_(SH._WL);
  if (!wlSh) return;
  var wlData = wlSh.getRange(2, WL.BALANCE + 1, wlSh.getLastRow() - 1, 1).getValues();
  var rows   = wlSh.getRange(2, WL.UNIT + 1, wlSh.getLastRow() - 1, 1).getValues();
  var dispRow = UL_DATA_ROW;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] !== uid) continue;
    var bal = toNum(wlData[i][0]);
    sh.getRange(dispRow, UL_WATER_COL + 11)
      .setValue(bal)
      .setBackground(bal <= 0 ? '#c8e6c9' : '#ffcdd2');
    dispRow++;
  }
}

function _recolourDuesBalances(sh, uid) {
  var dlSh = getSheet_(SH._DL);
  if (!dlSh) return;
  var dlData = dlSh.getRange(2, DL.BALANCE + 1, dlSh.getLastRow() - 1, 1).getValues();
  var rows   = dlSh.getRange(2, DL.UNIT + 1, dlSh.getLastRow() - 1, 1).getValues();
  var dispRow = UL_DATA_ROW;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] !== uid) continue;
    var bal = toNum(dlData[i][0]);
    sh.getRange(dispRow, UL_DUES_COL + 4)
      .setValue(bal)
      .setBackground(bal <= 0 ? '#c8e6c9' : '#ffcdd2');
    dispRow++;
  }
}
