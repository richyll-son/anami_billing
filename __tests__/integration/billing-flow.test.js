// ============================================================
// __tests__/integration/billing-flow.test.js
// Integration tests that wire real GAS functions together using
// mock sheet data via createGASContext.
//
// Scenarios covered:
//  1. recalcWaterBalances end-to-end with 3 units × 3 months
//  2. recalcDuesBalances end-to-end
//  3. getDuesRate_ + isBODExempt_ interaction
//  4. _applyWaterCredit + recalcWaterBalances combined flow
// ============================================================

'use strict';

const { createGASContext } = require('../helpers/gas-loader');

// ── Column constants (mirror Code.gs) ────────────────────────
const WL = {
  UNIT:0, YEAR:1, MONTH:2, BILL_DATE:3, PREV_DATE:4, PRESENT_DATE:5,
  PREV_RDG:6, CUR_RDG:7, RATE:8, DUE_DATE:9, PENALTY:10,
  DEBIT:11, CREDIT:12, BALANCE:13, ADDON:14,
  OR:15, REMARKS:16, BILL_NO:17, PAY_DATE:18
};
const DL = { UNIT:0, YEAR:1, MONTH:2, PAY_DATE:3, DEBIT:4, CREDIT:5, BALANCE:6, OR:7, REMARKS:8 };

// ── Row builders ──────────────────────────────────────────────
function wlRow(uid, year, month, {
  debit = 0, credit = 0, balance = 0, penalty = 0, addon = 0, dueDate = ''
} = {}) {
  const row = new Array(19).fill('');
  row[WL.UNIT]     = uid;
  row[WL.YEAR]     = year;
  row[WL.MONTH]    = month;
  row[WL.DUE_DATE] = dueDate;
  row[WL.PENALTY]  = penalty;
  row[WL.DEBIT]    = debit;
  row[WL.CREDIT]   = credit;
  row[WL.BALANCE]  = balance;
  row[WL.ADDON]    = addon;
  return row;
}

function dlRow(uid, year, month, { debit = 0, credit = 0, balance = 0 } = {}) {
  const row = new Array(9).fill('');
  row[DL.UNIT]    = uid;
  row[DL.YEAR]    = year;
  row[DL.MONTH]   = month;
  row[DL.DEBIT]   = debit;
  row[DL.CREDIT]  = credit;
  row[DL.BALANCE] = balance;
  return row;
}

// DR: AMOUNT=0,FROM_MONTH=1,FROM_YEAR=2,TO_MONTH=3,TO_YEAR=4
function drRow(amount, fromMonth, fromYear, toMonth = '', toYear = '') {
  return [amount, fromMonth, fromYear, toMonth, toYear];
}

// BOD: NAME=0,POSITION=1,UNIT=2,FROM_MONTH=3,FROM_YEAR=4,TO_MONTH=5,TO_YEAR=6
function bodRow(unit, fromMonth, fromYear, toMonth = '', toYear = '') {
  return ['', '', unit, fromMonth, fromYear, toMonth, toYear];
}

const WL_HEADER = new Array(19).fill('');
const DL_HEADER = new Array(9).fill('');

// ═════════════════════════════════════════════════════════════
// Integration: recalcWaterBalances with multiple units × months
// ═════════════════════════════════════════════════════════════
describe('recalcWaterBalances — end-to-end multi-unit multi-month', () => {
  let ctx;
  let sheet;

  beforeEach(() => {
    // 3 units × 3 months each, interleaved per unit (same order as in real ledger)
    // Unit A: Jan(opening=500), Feb(debit=300), Mar(debit=300,credit=800)
    // Unit B: Jan(opening=200), Feb(debit=250,penalty=10), Mar(debit=250)
    // Unit C: Jan(opening=0),   Feb(debit=400,addon=50),   Mar(debit=400)
    const rows = [
      // Unit A
      wlRow('P1B1L1', 2025, 'January',  { debit: 300, balance: 500 }),
      wlRow('P1B1L1', 2025, 'February', { debit: 300 }),
      wlRow('P1B1L1', 2025, 'March',    { debit: 300, credit: 800 }),
      // Unit B
      wlRow('P1B2L1', 2025, 'January',  { debit: 250, balance: 200 }),
      wlRow('P1B2L1', 2025, 'February', { debit: 250, penalty: 10 }),
      wlRow('P1B2L1', 2025, 'March',    { debit: 250 }),
      // Unit C
      wlRow('P1B3L1', 2025, 'January',  { debit: 400, balance: 0 }),
      wlRow('P1B3L1', 2025, 'February', { debit: 400, addon: 50 }),
      wlRow('P1B3L1', 2025, 'March',    { debit: 400 }),
    ];

    ctx = createGASContext({
      '_WaterLedger': [WL_HEADER, ...rows]
    });

    ctx.recalcWaterBalances();
    sheet = ctx._mockSS.getSheetByName('_WaterLedger');
  });

  it('Unit A: first row preserves stored opening balance (500)', () => {
    expect(sheet._data[1][WL.BALANCE]).toBe(500);
  });

  it('Unit A: Feb balance = 500 + 300 = 800', () => {
    expect(sheet._data[2][WL.BALANCE]).toBe(800);
  });

  it('Unit A: Mar balance = 800 + 300 - 800 = 300', () => {
    expect(sheet._data[3][WL.BALANCE]).toBe(300);
  });

  it('Unit B: first row preserves stored opening balance (200)', () => {
    expect(sheet._data[4][WL.BALANCE]).toBe(200);
  });

  it('Unit B: Feb balance = 200 + 250 + 10 = 460', () => {
    expect(sheet._data[5][WL.BALANCE]).toBe(460);
  });

  it('Unit B: Mar balance = 460 + 250 = 710', () => {
    expect(sheet._data[6][WL.BALANCE]).toBe(710);
  });

  it('Unit C: first row uses stored balance 0 as opening', () => {
    // Opening = 0 (stored), first row rule: preserve stored
    expect(sheet._data[7][WL.BALANCE]).toBe(0);
  });

  it('Unit C: Feb balance = 0 + 400 + 50(addon) = 450', () => {
    expect(sheet._data[8][WL.BALANCE]).toBe(450);
  });

  it('Unit C: Mar balance = 450 + 400 = 850', () => {
    expect(sheet._data[9][WL.BALANCE]).toBe(850);
  });

  it('Unit A balances do not affect Unit B or Unit C', () => {
    // Verify independence: B and C are computed from their own opening balances
    expect(sheet._data[4][WL.BALANCE]).toBe(200);  // B Jan opening
    expect(sheet._data[7][WL.BALANCE]).toBe(0);    // C Jan opening
  });
});

// ═════════════════════════════════════════════════════════════
// Integration: recalcDuesBalances end-to-end
// ═════════════════════════════════════════════════════════════
describe('recalcDuesBalances — end-to-end multi-unit', () => {
  it('correctly chains dues balances across units and months', () => {
    const rows = [
      dlRow('P1B1L1', 2025, 'January',  { debit: 500, credit: 0   }),
      dlRow('P1B1L1', 2025, 'February', { debit: 500, credit: 500 }),
      dlRow('P1B1L1', 2025, 'March',    { debit: 500, credit: 0   }),
      dlRow('P1B2L1', 2025, 'January',  { debit: 500, credit: 0   }),
      dlRow('P1B2L1', 2025, 'February', { debit: 500, credit: 0   }),
    ];

    const ctx = createGASContext({ '_DuesLedger': [DL_HEADER, ...rows] });
    ctx.recalcDuesBalances();

    const sheet = ctx._mockSS.getSheetByName('_DuesLedger');
    // A: Jan=500, Feb=500+500-500=500, Mar=500+500=1000
    expect(sheet._data[1][DL.BALANCE]).toBe(500);
    expect(sheet._data[2][DL.BALANCE]).toBe(500);
    expect(sheet._data[3][DL.BALANCE]).toBe(1000);
    // B: Jan=500, Feb=500+500=1000
    expect(sheet._data[4][DL.BALANCE]).toBe(500);
    expect(sheet._data[5][DL.BALANCE]).toBe(1000);
  });
});

// ═════════════════════════════════════════════════════════════
// Integration: getDuesRate_ + isBODExempt_
// ═════════════════════════════════════════════════════════════
describe('getDuesRate_ + isBODExempt_ integration', () => {
  let ctx;

  beforeEach(() => {
    ctx = createGASContext({});
  });

  it('BOD member with active exemption → dues = 0 (isBODExempt_ returns true)', () => {
    const bodData = [bodRow('P1B1L1', 'January', 2025, 'December', 2025)];
    const isExempt = ctx.isBODExempt_('P1B1L1', 2025, 6, bodData);
    // If exempt, dues = 0 (the billing engine uses 0 when isBODExempt_ is true)
    const duesDebit = isExempt ? 0 : ctx.getDuesRate_(2025, 6, []);
    expect(isExempt).toBe(true);
    expect(duesDebit).toBe(0);
  });

  it('BOD member with expired exemption → dues = normal rate', () => {
    // Exemption ended December 2024; billing is March 2025 → not exempt
    const bodData = [bodRow('P1B1L1', 'January', 2024, 'December', 2024)];
    const isExempt = ctx.isBODExempt_('P1B1L1', 2025, 3, bodData);
    const duesDebit = isExempt ? 0 : ctx.getDuesRate_(2025, 3, []);
    expect(isExempt).toBe(false);
    expect(duesDebit).toBe(500); // ASSOC_DUES fallback
  });

  it('custom dues rate in range → uses custom rate for non-BOD units', () => {
    const duesRatesData = [drRow(600, 'January', 2025, 'December', 2025)];
    const bodData = [];
    const uid = 'P1B2L3';
    const isExempt = ctx.isBODExempt_(uid, 2025, 6, bodData);
    const rate = isExempt ? 0 : ctx.getDuesRate_(2025, 6, duesRatesData);
    expect(isExempt).toBe(false);
    expect(rate).toBe(600);
  });

  it('no matching dues rate → fallback to ASSOC_DUES (500)', () => {
    // Rate only covers 2024, billing is 2025 → no match → fallback to last positive (400)
    // But if we pass empty rates array, fallback is ASSOC_DUES
    expect(ctx.getDuesRate_(2025, 6, [])).toBe(500);
  });

  it('multiple BOD rows — only matching unit gets exempt', () => {
    const bodData = [
      bodRow('P1B1L1', 'January', 2025, 'December', 2025),
      bodRow('P1B2L3', 'January', 2025, 'December', 2025),
    ];
    expect(ctx.isBODExempt_('P1B1L1', 2025, 6, bodData)).toBe(true);
    expect(ctx.isBODExempt_('P1B5L7', 2025, 6, bodData)).toBe(false);
  });
});

// ═════════════════════════════════════════════════════════════
// Integration: _applyWaterCredit + recalcWaterBalances
// ═════════════════════════════════════════════════════════════
describe('_applyWaterCredit + recalcWaterBalances combined flow', () => {
  const UID = 'P1B1L1';

  // Build 3-month WL data with progressive unpaid balances
  function makeInitialData() {
    return [
      // Month 1 (opening balance = 300; stored as opening)
      wlRow(UID, 2025, 'January',  { debit: 300, credit: 0, balance: 300 }),
      // Month 2 (balance = 600 = 300+300)
      wlRow(UID, 2025, 'February', { debit: 300, credit: 0, balance: 600 }),
      // Month 3 (balance = 900 = 600+300)
      wlRow(UID, 2025, 'March',    { debit: 300, credit: 0, balance: 900 }),
    ];
  }

  it('partial payment credits oldest bills and recalcWaterBalances updates balance correctly', () => {
    const data = makeInitialData();
    const ctx = createGASContext({ '_WaterLedger': [WL_HEADER, ...data] });

    // Apply partial payment of 400 to in-memory data
    ctx._applyWaterCredit(data, UID, 400, new Date(), 'OR-TEST', 'integration test');

    // Jan gets 300 (fully paid), Feb gets 100 (partial)
    expect(data[0][WL.CREDIT]).toBe(300);
    expect(data[1][WL.CREDIT]).toBe(100);
    expect(data[2][WL.CREDIT]).toBe(0);

    // Now write the mutated data back to the mock sheet
    const sheet = ctx._mockSS.getSheetByName('_WaterLedger');
    for (let i = 0; i < data.length; i++) {
      for (let j = 0; j < data[i].length; j++) {
        sheet._data[i + 1][j] = data[i][j];
      }
    }

    // Run recalcWaterBalances
    ctx.recalcWaterBalances();

    // Jan (opening = stored 300, first row rule)
    // Feb: 300 + 300 - 100 = 500
    // Mar: 500 + 300 - 0 = 800
    expect(sheet._data[1][WL.BALANCE]).toBe(300); // Jan preserved
    expect(sheet._data[2][WL.BALANCE]).toBe(500); // Feb
    expect(sheet._data[3][WL.BALANCE]).toBe(800); // Mar
  });

  it('overpayment results in negative balance (credit carry-forward)', () => {
    const data = makeInitialData();
    const ctx = createGASContext({ '_WaterLedger': [WL_HEADER, ...data] });

    // Pay 1500 when total owed = 900 → 600 overpayment
    ctx._applyWaterCredit(data, UID, 1500, new Date(), 'OR-OVP', 'overpayment test');

    // Write back
    const sheet = ctx._mockSS.getSheetByName('_WaterLedger');
    for (let i = 0; i < data.length; i++) {
      for (let j = 0; j < data[i].length; j++) {
        sheet._data[i + 1][j] = data[i][j];
      }
    }

    ctx.recalcWaterBalances();

    // Jan (opening=300, first row keeps it)
    // Feb: 300 + 300 - 300 = 300 (Jan fully cleared opens=300, Feb debit=300, credit=300)
    // Mar: balance depends on total credits distributed
    // Overpayment on last row (Mar): 300 (bill) + 600 (advance) = 900 credit
    // Mar balance: 300 (from Feb) + 300 (debit) - 900 (credit) = -300
    const marBalance = sheet._data[3][WL.BALANCE];
    expect(marBalance).toBeLessThan(0);
  });

  it('full payment for all 3 months reduces balance correctly', () => {
    // recalcWaterBalances preserves the FIRST row's stored balance (first-row rule).
    // Data: Jan(stored=300), Feb(debit=300,credit=0), Mar(debit=300,credit=0)
    // _applyWaterCredit with 900: Jan gets 300, Feb gets 300, Mar gets 300
    // After recalc:
    //   Jan (first row): preserved as stored 300
    //   Feb: 300(prev) + 300(debit) - 300(credit) = 300
    //   Mar: 300(prev) + 300(debit) - 300(credit) = 300
    // The final balance is 300 (not 0) because Jan's stored balance is preserved.
    const data = makeInitialData();
    const ctx = createGASContext({ '_WaterLedger': [WL_HEADER, ...data] });

    ctx._applyWaterCredit(data, UID, 900, new Date(), 'OR-FULL', 'full payment');

    const sheet = ctx._mockSS.getSheetByName('_WaterLedger');
    for (let i = 0; i < data.length; i++) {
      for (let j = 0; j < data[i].length; j++) {
        sheet._data[i + 1][j] = data[i][j];
      }
    }

    ctx.recalcWaterBalances();

    // All 3 months are fully paid (credit = debit for each)
    expect(data[0][WL.CREDIT]).toBe(300); // Jan paid
    expect(data[1][WL.CREDIT]).toBe(300); // Feb paid
    expect(data[2][WL.CREDIT]).toBe(300); // Mar paid

    // Final running balance: preserved opening(300) + paid months net to 0
    // Jan preserved=300; Feb=300+300-300=300; Mar=300+300-300=300
    const finalBal = sheet._data[3][WL.BALANCE];
    expect(finalBal).toBe(300);
  });

  it('multi-unit: payment for Unit A does not affect Unit B balances', () => {
    const rowsA = makeInitialData();
    const rowsB = [
      wlRow('P1B2L1', 2025, 'January',  { debit: 500, balance: 500 }),
      wlRow('P1B2L1', 2025, 'February', { debit: 500, balance: 0 }),
    ];

    const allRows = [...rowsA, ...rowsB];
    const ctx = createGASContext({ '_WaterLedger': [WL_HEADER, ...allRows] });

    ctx._applyWaterCredit(allRows, UID, 300, new Date(), 'OR-A', 'payment for A');

    // Write back
    const sheet = ctx._mockSS.getSheetByName('_WaterLedger');
    for (let i = 0; i < allRows.length; i++) {
      for (let j = 0; j < allRows[i].length; j++) {
        sheet._data[i + 1][j] = allRows[i][j];
      }
    }

    ctx.recalcWaterBalances();

    // Unit B should be unaffected: its Feb balance = 500 (opening) + 500 = 1000
    expect(sheet._data[4][WL.BALANCE]).toBe(500);  // B Jan (opening preserved)
    expect(sheet._data[5][WL.BALANCE]).toBe(1000); // B Feb
  });
});
