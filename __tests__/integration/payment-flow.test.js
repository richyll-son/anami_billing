// ============================================================
// __tests__/integration/payment-flow.test.js
// End-to-end payment posting flow tests.
//
// Tests combine _applyWaterCredit + _applyDuesCredit to simulate
// the full payment cycle used by postPayments() and webapp_addPayment().
// ============================================================

'use strict';

const { createGASContext } = require('../helpers/gas-loader');

// ── Column constants ──────────────────────────────────────────
const WL = {
  UNIT:0, YEAR:1, MONTH:2, BILL_DATE:3, PREV_DATE:4, PRESENT_DATE:5,
  PREV_RDG:6, CUR_RDG:7, RATE:8, DUE_DATE:9, PENALTY:10,
  DEBIT:11, CREDIT:12, BALANCE:13, ADDON:14,
  OR:15, REMARKS:16, BILL_NO:17, PAY_DATE:18
};
const DL = { UNIT:0, YEAR:1, MONTH:2, PAY_DATE:3, DEBIT:4, CREDIT:5, BALANCE:6, OR:7, REMARKS:8 };

// ── Row builders ──────────────────────────────────────────────
function wlRow(uid, year, month, { debit = 0, credit = 0, balance = 0, penalty = 0, addon = 0 } = {}) {
  const row = new Array(19).fill('');
  row[WL.UNIT]    = uid;
  row[WL.YEAR]    = year;
  row[WL.MONTH]   = month;
  row[WL.PENALTY] = penalty;
  row[WL.DEBIT]   = debit;
  row[WL.CREDIT]  = credit;
  row[WL.BALANCE] = balance;
  row[WL.ADDON]   = addon;
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

const WL_HEADER = new Array(19).fill('');
const DL_HEADER = new Array(9).fill('');

// ═════════════════════════════════════════════════════════════
// Both payment type: _applyWaterCredit + _applyDuesCredit
// ═════════════════════════════════════════════════════════════
describe('"Both" payment type: water + dues credits applied together', () => {
  const UID      = 'P1B1L1';
  const PAY_DATE = new Date(2025, 2, 15);
  const OR_NUM   = 'OR-BOTH-001';

  let ctx, wlData, dlData;

  beforeEach(() => {
    // 2 months of unpaid water bills
    wlData = [
      wlRow(UID, 2025, 'January',  { debit: 300 }),
      wlRow(UID, 2025, 'February', { debit: 300 }),
    ];
    // 2 months of unpaid dues
    dlData = [
      dlRow(UID, 2025, 'January',  { debit: 500 }),
      dlRow(UID, 2025, 'February', { debit: 500 }),
    ];
    ctx = createGASContext({
      '_WaterLedger': [WL_HEADER, ...wlData],
      '_DuesLedger':  [DL_HEADER, ...dlData]
    });
  });

  it('water credit is applied to oldest unpaid water bills', () => {
    const waterAmt = 400; // covers Jan(300) + partial Feb(100)
    ctx._applyWaterCredit(wlData, UID, waterAmt, PAY_DATE, OR_NUM, 'Both payment');

    expect(wlData[0][WL.CREDIT]).toBe(300); // Jan fully paid
    expect(wlData[1][WL.CREDIT]).toBe(100); // Feb partial
  });

  it('dues credit is applied to oldest unpaid dues bills', () => {
    const duesAmt = 700; // covers Jan(500) + partial Feb(200)
    ctx._applyDuesCredit(dlData, UID, duesAmt, PAY_DATE, OR_NUM, 'Both payment');

    expect(dlData[0][DL.CREDIT]).toBe(500); // Jan fully paid
    expect(dlData[1][DL.CREDIT]).toBe(200); // Feb partial
  });

  it('OR numbers are recorded on all affected water rows', () => {
    ctx._applyWaterCredit(wlData, UID, 600, PAY_DATE, OR_NUM, 'Both payment');

    // Both rows should have OR_NUM
    expect(wlData[0][WL.OR]).toBe(OR_NUM);
    expect(wlData[1][WL.OR]).toBe(OR_NUM);
  });

  it('OR numbers are recorded on all affected dues rows', () => {
    ctx._applyDuesCredit(dlData, UID, 1000, PAY_DATE, OR_NUM, 'Both payment');

    expect(dlData[0][DL.OR]).toBe(OR_NUM);
    expect(dlData[1][DL.OR]).toBe(OR_NUM);
  });

  it('remarks include payment annotation on water rows', () => {
    ctx._applyWaterCredit(wlData, UID, 600, PAY_DATE, OR_NUM, 'Both');

    // First row should say "Pmt 1 of 2"
    expect(wlData[0][WL.REMARKS]).toContain('Pmt 1 of 2');
    expect(wlData[1][WL.REMARKS]).toContain('Pmt 2 of 2');
  });

  it('remarks include payment annotation on dues rows', () => {
    ctx._applyDuesCredit(dlData, UID, 1000, PAY_DATE, OR_NUM, 'Both');

    expect(dlData[0][DL.REMARKS]).toContain('Pmt 1 of 2');
    expect(dlData[1][DL.REMARKS]).toContain('Pmt 2 of 2');
  });

  it('combined payment does not cross-contaminate water and dues', () => {
    // Apply water payment
    ctx._applyWaterCredit(wlData, UID, 600, PAY_DATE, OR_NUM, 'Both');
    // Apply dues payment
    ctx._applyDuesCredit(dlData, UID, 500, PAY_DATE, OR_NUM, 'Both');

    // Water data untouched by dues call
    expect(wlData[0][WL.CREDIT]).toBe(300);
    // Dues data untouched by water call
    expect(dlData[0][DL.CREDIT]).toBe(500);
    expect(dlData[1][DL.CREDIT]).toBe(0); // only 500 applied to dues, Jan fully covered
  });
});

// ═════════════════════════════════════════════════════════════
// Multi-unit payment isolation
// ═════════════════════════════════════════════════════════════
describe('Multi-unit payment isolation', () => {
  const UID_A = 'P1B1L1';
  const UID_B = 'P1B2L1';
  const PAY_DATE = new Date(2025, 2, 15);

  it('payment for Unit A does not affect Unit B water data', () => {
    const wlData = [
      wlRow(UID_A, 2025, 'January', { debit: 300 }),
      wlRow(UID_A, 2025, 'February', { debit: 300 }),
      wlRow(UID_B, 2025, 'January', { debit: 400 }),
      wlRow(UID_B, 2025, 'February', { debit: 400 }),
    ];

    const ctx = createGASContext({});
    ctx._applyWaterCredit(wlData, UID_A, 600, PAY_DATE, 'OR-A', 'Unit A payment');

    // Unit A: fully paid
    expect(wlData[0][WL.CREDIT]).toBe(300);
    expect(wlData[1][WL.CREDIT]).toBe(300);

    // Unit B: untouched
    expect(wlData[2][WL.CREDIT]).toBe(0);
    expect(wlData[3][WL.CREDIT]).toBe(0);
  });

  it('payment for Unit B does not affect Unit A dues data', () => {
    const dlData = [
      dlRow(UID_A, 2025, 'January', { debit: 500 }),
      dlRow(UID_B, 2025, 'January', { debit: 500 }),
    ];

    const ctx = createGASContext({});
    ctx._applyDuesCredit(dlData, UID_B, 500, PAY_DATE, 'OR-B', 'Unit B dues');

    // Unit A untouched
    expect(dlData[0][DL.CREDIT]).toBe(0);
    // Unit B fully paid
    expect(dlData[1][DL.CREDIT]).toBe(500);
  });

  it('two units can be paid in the same session without interference', () => {
    const wlData = [
      wlRow(UID_A, 2025, 'January', { debit: 300 }),
      wlRow(UID_B, 2025, 'January', { debit: 400 }),
    ];

    const ctx = createGASContext({});
    ctx._applyWaterCredit(wlData, UID_A, 300, PAY_DATE, 'OR-A', 'A');
    ctx._applyWaterCredit(wlData, UID_B, 400, PAY_DATE, 'OR-B', 'B');

    expect(wlData[0][WL.CREDIT]).toBe(300); // A paid
    expect(wlData[1][WL.CREDIT]).toBe(400); // B paid
  });
});

// ═════════════════════════════════════════════════════════════
// Overpayment handling
// ═════════════════════════════════════════════════════════════
describe('Overpayment handling', () => {
  const UID      = 'P1B1L1';
  const PAY_DATE = new Date(2025, 2, 15);
  const OR_NUM   = 'OR-OVP';

  it('excess credited to last water row when overpaying', () => {
    const wlData = [
      wlRow(UID, 2025, 'January',  { debit: 300 }),
      wlRow(UID, 2025, 'February', { debit: 300 }),
    ];

    const ctx = createGASContext({});
    // Total owed = 600, pay 800 → 200 overpayment
    ctx._applyWaterCredit(wlData, UID, 800, PAY_DATE, OR_NUM, 'overpayment');

    // Last row should have: 300 (bill cleared) + 200 (advance) = 500
    expect(wlData[1][WL.CREDIT]).toBe(500);
  });

  it('overpayment remarks note "Advance ₱X" on last row', () => {
    const wlData = [
      wlRow(UID, 2025, 'January', { debit: 300 }),
    ];

    const ctx = createGASContext({});
    ctx._applyWaterCredit(wlData, UID, 500, PAY_DATE, OR_NUM, 'advance payment');

    // The 200 overpayment should be noted in remarks
    expect(wlData[0][WL.REMARKS]).toContain('Advance');
    expect(wlData[0][WL.REMARKS]).toContain('200.00');
  });

  it('excess credited to last dues row when overpaying dues', () => {
    const dlData = [
      dlRow(UID, 2025, 'January',  { debit: 500 }),
      dlRow(UID, 2025, 'February', { debit: 500 }),
    ];

    const ctx = createGASContext({});
    // Total owed = 1000, pay 1300 → 300 overpayment
    ctx._applyDuesCredit(dlData, UID, 1300, PAY_DATE, OR_NUM, 'dues overpayment');

    // Feb (last row): 500 (cleared) + 300 (advance) = 800
    expect(dlData[1][DL.CREDIT]).toBe(800);
    expect(dlData[1][DL.REMARKS]).toContain('Advance');
  });

  it('overpayment on single-row ledger: all excess on same row', () => {
    const wlData = [wlRow(UID, 2025, 'January', { debit: 300 })];

    const ctx = createGASContext({});
    ctx._applyWaterCredit(wlData, UID, 1000, PAY_DATE, OR_NUM, 'big advance');

    // Single row: 300 paid + 700 advance = 1000 total credit
    expect(wlData[0][WL.CREDIT]).toBe(1000);
    expect(wlData[0][WL.REMARKS]).toContain('700.00');
  });

  it('recalcWaterBalances after overpayment correctly computes balance', () => {
    // Jan: stored balance=300, debit=300, credit=0
    // Feb: debit=300, credit=0, balance=600
    // Pay 900 for 600 owed → 300 overpayment added to last row
    // After _applyWaterCredit:
    //   Jan credit=300, Feb credit=300 (cleared) + 300 (advance) = 600
    // After recalcWaterBalances:
    //   Jan (first row preserved): balance = 300 (stored)
    //   Feb: prev=300, debit=300, credit=600 → 300+300-600 = 0
    // Balance reaches 0 (unit fully paid including advance)
    const data = [
      wlRow(UID, 2025, 'January',  { debit: 300, credit: 0, balance: 300 }),
      wlRow(UID, 2025, 'February', { debit: 300, credit: 0, balance: 600 }),
    ];
    const ctx = createGASContext({ '_WaterLedger': [WL_HEADER, ...data] });

    ctx._applyWaterCredit(data, UID, 900, PAY_DATE, OR_NUM, 'overpay');

    const sheet = ctx._mockSS.getSheetByName('_WaterLedger');
    for (let i = 0; i < data.length; i++) {
      for (let j = 0; j < data[i].length; j++) {
        sheet._data[i + 1][j] = data[i][j];
      }
    }

    ctx.recalcWaterBalances();

    // Jan credit = 300 (fully paid), Feb credit = 600 (300 cleared + 300 advance)
    expect(data[0][WL.CREDIT]).toBe(300);
    expect(data[1][WL.CREDIT]).toBe(600);

    // Final balance: Jan(300 preserved) + Feb(300+300-600=0) = 0
    const finalBal = sheet._data[2][WL.BALANCE];
    expect(finalBal).toBe(0);
  });
});

// ═════════════════════════════════════════════════════════════
// Edge cases in payment flow
// ═════════════════════════════════════════════════════════════
describe('Edge cases in payment flow', () => {
  it('zero water amount does not modify any row', () => {
    const wlData = [wlRow('P1B1L1', 2025, 'January', { debit: 300 })];
    const ctx = createGASContext({});
    ctx._applyWaterCredit(wlData, 'P1B1L1', 0, new Date(), 'OR-000', 'zero');
    expect(wlData[0][WL.CREDIT]).toBe(0);
    expect(wlData[0][WL.OR]).toBe('');
  });

  it('zero dues amount does not modify any row', () => {
    const dlData = [dlRow('P1B1L1', 2025, 'January', { debit: 500 })];
    const ctx = createGASContext({});
    ctx._applyDuesCredit(dlData, 'P1B1L1', 0, new Date(), 'OR-000', 'zero');
    expect(dlData[0][DL.CREDIT]).toBe(0);
    expect(dlData[0][DL.OR]).toBe('');
  });

  it('payment with penalty and addon: unpaid = debit + penalty + addon - credit', () => {
    const wlData = [
      wlRow('P1B1L1', 2025, 'January', { debit: 300, penalty: 15, addon: 50, credit: 0 })
    ];
    const ctx = createGASContext({});
    // unpaid = 300 + 15 + 50 = 365
    ctx._applyWaterCredit(wlData, 'P1B1L1', 365, new Date(), 'OR-PAP', 'full');
    expect(wlData[0][WL.CREDIT]).toBe(365);
  });

  it('already-paid rows are skipped even when they have a non-zero credit', () => {
    const wlData = [
      wlRow('P1B1L1', 2025, 'January',  { debit: 300, credit: 300 }), // paid
      wlRow('P1B1L1', 2025, 'February', { debit: 300, credit: 0   }), // unpaid
    ];
    const ctx = createGASContext({});
    ctx._applyWaterCredit(wlData, 'P1B1L1', 300, new Date(), 'OR-NEW', '');

    expect(wlData[0][WL.CREDIT]).toBe(300); // unchanged
    expect(wlData[1][WL.CREDIT]).toBe(300); // newly paid
  });
});
