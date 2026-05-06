// ============================================================
// AHNHAI Billing System — PrintGenerator.gs
// Build Phase 1 and Phase 2 Bill Print sheets.
// All 300 bills stacked vertically, one bill per unit.
// Each bill is ~36 rows.  Column A = label/content,
// Column B = right-side value (for two-column rows).
// Entire sheet is written in two SetValues calls per phase
// (values then formats are batched for speed).
// ============================================================

var BILL_ROWS = 38; // rows allocated per bill (including separator)

function regenerateBillPrint(year, month, monthName, rate, totalCons,
    mcwdAmt, elecAmt, manpower, mcwdFrom, mcwdTo, elecFrom, elecTo) {
  _buildPhasePrint(1, year, month, monthName, rate, totalCons,
    mcwdAmt, elecAmt, manpower, mcwdFrom, mcwdTo, elecFrom, elecTo);
  _buildPhasePrint(2, year, month, monthName, rate, totalCons,
    mcwdAmt, elecAmt, manpower, mcwdFrom, mcwdTo, elecFrom, elecTo);
}

function _buildPhasePrint(phase, year, month, monthName, rate, totalCons,
    mcwdAmt, elecAmt, manpower, mcwdFrom, mcwdTo, elecFrom, elecTo) {

  var shName = (phase === 1) ? SH.P1_PRINT : SH.P2_PRINT;
  var sh     = getSheet_(shName);
  if (!sh) return;

  sh.clearContents(); sh.clearFormats();
  sh.setColumnWidth(1, 420);
  sh.setColumnWidth(2, 200);

  // ── Load source data ──────────────────────────────────────
  var mlSh   = getSheet_(SH.MASTERLIST);
  var wlSh   = getSheet_(SH._WL);
  var dlSh   = getSheet_(SH._DL);

  if (!mlSh || mlSh.getLastRow() < 2) return;

  var mlData = mlSh.getDataRange().getValues();
  var wlData = (wlSh && wlSh.getLastRow() > 1)
    ? wlSh.getRange(2, 1, wlSh.getLastRow() - 1, WL_COLS).getValues() : [];
  var dlData = (dlSh && dlSh.getLastRow() > 1)
    ? dlSh.getRange(2, 1, dlSh.getLastRow() - 1, DL_COLS).getValues() : [];

  // ── Index water and dues data for current period ──────────
  var waterIdx = {};   // uid → last matching _WL row
  var duesIdx  = {};   // uid → last matching _DL row
  wlData.forEach(function(r) {
    if (r[WL.UNIT] && parseInt(r[WL.YEAR],10) === year && r[WL.MONTH] === monthName) {
      waterIdx[r[WL.UNIT]] = r;
    }
  });
  dlData.forEach(function(r) {
    if (r[DL.UNIT] && parseInt(r[DL.YEAR],10) === year && r[DL.MONTH] === monthName) {
      duesIdx[r[DL.UNIT]] = r;
    }
  });

  // ── Get previous balances (arrears) per unit ──────────────
  // Arrears = balance carried in BEFORE this month's debit
  var prevWaterBal = {};   // uid → balance before this period
  var prevDuesBal  = {};

  // For each unit, the "previous balance" = current BALANCE - curDebit - curPenalty - curAddon
  Object.keys(waterIdx).forEach(function(uid) {
    var r   = waterIdx[uid];
    var bal = toNum(r[WL.BALANCE]);
    var deb = toNum(r[WL.DEBIT]);
    var pen = toNum(r[WL.PENALTY]);
    var add = toNum(r[WL.ADDON]);
    prevWaterBal[uid] = Math.max(0, bal - deb - pen - add);
  });
  Object.keys(duesIdx).forEach(function(uid) {
    var r   = duesIdx[uid];
    var bal = toNum(r[DL.BALANCE]);
    var deb = toNum(r[DL.DEBIT]);
    prevDuesBal[uid] = Math.max(0, bal - deb);
  });

  // Utility amounts for bill header
  var utilityAmt = mcwdAmt + elecAmt;
  var mcwdFromStr = mcwdFrom ? fmtDate(mcwdFrom) : '';
  var mcwdToStr   = mcwdTo   ? fmtDate(mcwdTo)   : '';
  var elecFromStr = elecFrom ? fmtDate(elecFrom)  : '';
  var elecToStr   = elecTo   ? fmtDate(elecTo)    : '';

  // ── Build value arrays and format metadata ────────────────
  var allValsA   = [];  // column A values (one element per sheet row)
  var allValsB   = [];  // column B values
  var boldRows   = {};  // row (1-based) → true
  var bgRows     = {};  // row → hex bg colour
  var totalRows  = 0;

  // ── Loop over masterlist units for this phase ─────────────
  mlData.slice(1).forEach(function(mlRow) {
    var uid = mlRow[0];
    if (!uid) return;
    var parsed = parseUID(uid);
    if (!parsed || parsed.phase !== phase) return;
    if (String(mlRow[11]).toLowerCase() === 'no') return;

    var lastName  = mlRow[4];
    var firstName = mlRow[5];
    var meterNo   = mlRow[9];
    var stubOut   = mlRow[10];
    var billNo    = buildBillNum(year, month, parsed.phase, parsed.block, parsed.lot);

    var wlRow = waterIdx[uid];
    var dlRow = duesIdx[uid];

    var prevReading     = wlRow ? toNum(wlRow[WL.PREV_RDG])    : 0;
    var curReading      = wlRow ? toNum(wlRow[WL.CUR_RDG])     : 0;
    var cons            = Math.max(0, curReading - prevReading);
    var rawBill         = cons * rate;
    var waterBill       = Math.max(MIN_WATER_BILL, rawBill);
    var penalty         = wlRow ? toNum(wlRow[WL.PENALTY])     : 0;
    var addon           = wlRow ? toNum(wlRow[WL.ADDON])       : 0;
    var waterArrears    = prevWaterBal[uid] || 0;
    var waterTotal      = wlRow ? toNum(wlRow[WL.BALANCE])     : waterBill + penalty + addon;
    var presentDateStr  = wlRow ? wlRow[WL.PRESENT_DATE]       : '';
    var prevDateStr     = wlRow ? wlRow[WL.PREV_DATE]          : '';
    var dueDateStr      = wlRow ? wlRow[WL.DUE_DATE]           : fmtDate(getDueDate(year, month));
    var billDateStr     = wlRow ? wlRow[WL.BILL_DATE]          : fmtDate(getBillDate(year, month));

    var duesDebit       = dlRow ? toNum(dlRow[DL.DEBIT])       : ASSOC_DUES;
    var duesArrears     = prevDuesBal[uid] || 0;
    var duesTotal       = dlRow ? toNum(dlRow[DL.BALANCE])     : duesDebit + duesArrears;

    // "Total Bill" = current water charges (already includes minimum)
    var totalBill = Math.round((waterBill + penalty) * 100) / 100;

    // Overpayment: if credit > debit for this bill, excess shows here
    var overpay = wlRow ? Math.max(0, toNum(wlRow[WL.CREDIT]) - toNum(wlRow[WL.DEBIT]) - penalty) : 0;
    if (overpay > 0) totalBill = Math.max(0, totalBill - overpay);

    // ── Build this bill's rows  [colA, colB] ─────────────────
    var startRow = totalRows + 1; // 1-based
    var bill = _makeBillLines(
      lastName, firstName, uid, billNo, billDateStr, meterNo, stubOut,
      prevDateStr, presentDateStr, prevReading, curReading, cons,
      rate, totalCons, utilityAmt, mcwdFromStr, mcwdToStr, elecFromStr, elecToStr,
      manpower, rawBill, waterBill, penalty, overpay, totalBill,
      waterArrears, addon, waterTotal, dueDateStr,
      duesArrears, duesDebit, duesTotal, monthName
    );

    bill.forEach(function(line, idx) {
      allValsA.push(line[0]);
      allValsB.push(line[1]);

      var absRow = startRow + idx;
      if (line[2]) boldRows[absRow] = true;
      if (line[3]) bgRows[absRow]   = line[3];
    });

    // Padding rows to reach BILL_ROWS (including separator)
    while (bill.length < BILL_ROWS - 1) {
      allValsA.push(''); allValsB.push('');
      bill.push(['','',false,'']);
    }

    // Separator row
    allValsA.push(''); allValsB.push('');
    var sepRow = startRow + BILL_ROWS - 1;
    bgRows[sepRow] = '#cccccc';

    totalRows += BILL_ROWS;
  });

  if (totalRows === 0) return;

  // ── Write values in 2 batch calls ────────────────────────
  sh.getRange(1, 1, totalRows, 1).setValues(allValsA.map(function(v) { return [v]; }));
  sh.getRange(1, 2, totalRows, 1).setValues(allValsB.map(function(v) { return [v]; }));

  // ── Apply formatting in ranges ────────────────────────────
  Object.keys(boldRows).forEach(function(r) {
    sh.getRange(+r, 1, 1, 2).setFontWeight('bold');
  });
  Object.keys(bgRows).forEach(function(r) {
    sh.getRange(+r, 1, 1, 2).setBackground(bgRows[r]);
  });
}

// ── Build one bill's lines as [[colA, colB, bold, bgColor], …] ──
function _makeBillLines(lastName, firstName, uid, billNo, billDateStr,
    meterNo, stubOut, prevDateStr, presentDateStr, prevReading, curReading, cons,
    rate, totalCons, utilityAmt, mcwdFrom, mcwdTo, elecFrom, elecTo,
    manpower, rawBill, waterBill, penalty, overpay, totalBill,
    waterArrears, addon, waterTotal, dueDateStr,
    duesArrears, duesDebit, duesTotal, monthName) {

  var T = true, F = false;
  var HDR1  = '#1a237e';
  var HDR2  = '#e8eaf6';
  var SEC   = '#e3f2fd';
  var SEC2  = '#e8f5e9';
  var TOTAL = '#fff9c4';
  var GRAND = '#ffccbc';
  var LINE  = '#eeeeee';

  return [
    // Row 1: top border
    ['═══════════════════════════════════════════════════════════════════', '', T, HDR1],
    // Row 2-3: header
    ['ANAMI HOMES NORTH HOMEOWNERS ASSOCIATION — BILLING STATEMENT',       '', T, HDR1],
    ['Brgy. Jugan, Consolacion, Cebu 6001  |  Tel: (032) XXX-XXXX',       '', F, HDR1],
    // Row 4: blank
    ['', '', F, '#ffffff'],
    // Row 5-7: unit info
    ['Name: ' + lastName + ', ' + firstName,          'Bill No.:  ' + billNo,   F, HDR2],
    ['Address: ' + uid,                                'Bill Date: ' + billDateStr, F, HDR2],
    ['Meter No.: ' + meterNo + '   Stub-out: ' + stubOut,
      'Coverage: ' + prevDateStr + ' – ' + presentDateStr,                        F, HDR2],
    // Row 8: blank
    ['', '', F, '#ffffff'],
    // Row 9-12: reading table
    ['Service Period', 'Meter Reading | Due Date',                          T, SEC],
    ['Present:  ' + presentDateStr,   String(curReading)  + '   |  Due: ' + dueDateStr, F, '#ffffff'],
    ['Previous: ' + prevDateStr,      String(prevReading),                              F, '#ffffff'],
    ['Total Water Consumed:',         cons + ' m³',                         T, LINE],
    // Row 13: blank
    ['', '', F, '#ffffff'],
    // Row 14-17: cost breakdown
    ['Total Water Pump Electricity & Maintenance', '₱ ' + fmt2(utilityAmt), F, '#ffffff'],
    ['AHNHAI Admin Charge (Manpower)',             '₱ ' + fmt2(manpower),   F, '#ffffff'],
    ['Total Consumption (all units):',             String(totalCons) + ' m³', F, '#ffffff'],
    ['Effective Rate per Cubic Meter:',            '₱ ' + fmt2(rate),        T, LINE],
    // Row 18: blank
    ['', '', F, '#ffffff'],
    // Row 19: water bill header
    ['WATER BILL', '', T, SEC],
    // Row 20-26: water bill breakdown
    ['Current Bill:  ' + cons + ' m³  @ ₱' + fmt2(rate),    '= ₱ ' + fmt2(rawBill), F, '#ffffff'],
    ['Minimum Rate Charge:',                                  '₱ ' + fmt2(waterBill), F, '#ffffff'],
    ['Plus: Penalty (5% of unpaid balance):',                 '₱ ' + fmt2(penalty),   F, '#ffffff'],
    ['Less: Overpayment:',                                    '₱ ' + fmt2(overpay),    F, '#ffffff'],
    ['Total Bill:',                                           '₱ ' + fmt2(totalBill), T, TOTAL],
    ['Arrears (previous unpaid balance):',                    '₱ ' + fmt2(waterArrears), F, '#ffffff'],
    ['Add-on MCWD 10/10:',                                    '₱ ' + fmt2(addon),      F, '#ffffff'],
    ['TOTAL WATER AMOUNT DUE:',                               '₱ ' + fmt2(waterTotal), T, GRAND],
    // Row 28: blank
    ['', '', F, '#ffffff'],
    // Row 29-31: reminders
    ['IMPORTANT REMINDERS:', '', T, '#fff3e0'],
    ['• Minimum charge of ₱250 applies if Current Bill is below ₱250.', '', F, '#fff3e0'],
    ['• 5% penalty on unpaid balance for late payments.  Non-payment for 2 months = disconnection.', '', F, '#fff3e0'],
    // Row 32: blank
    ['', '', F, '#ffffff'],
    // Row 33-36: association dues
    ['ASSOCIATION DUES', '', T, SEC2],
    ['Arrears:',                   '₱ ' + fmt2(duesArrears), F, '#ffffff'],
    ['Due (' + monthName + '):',  '₱ ' + fmt2(duesDebit),   F, '#ffffff'],
    ['TOTAL DUES AMOUNT DUE:',     '₱ ' + fmt2(duesTotal),   T, SEC2],
    // Row 37: payment instructions
    ['PLEASE PAY THRU: CASH at AHNHAI Office/Clubhouse  |  BDO SM Consolacion Acnt No. 007688009327  |  Accnt Name: Anami Homes North (AHNHAI)', '', F, '#fce4ec'],
  ];
  // Note: BILL_ROWS = 38; row 38 is the separator added in the caller
}
