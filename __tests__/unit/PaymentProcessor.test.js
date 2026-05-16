// ============================================================
// __tests__/unit/PaymentProcessor.test.js
// Tests for _applyWaterCredit and _applyDuesCredit from
// PaymentProcessor.gs.
//
// Both functions mutate in-memory 2-D arrays (wlData / dlData)
// without touching sheets. We construct minimal arrays using the
// WL / DL column constants and verify mutations after each call.
//
// WL column constants (0-based):
//   UNIT=0,YEAR=1,MONTH=2,BILL_DATE=3,PREV_DATE=4,
//   PRESENT_DATE=5,PREV_RDG=6,CUR_RDG=7,RATE=8,DUE_DATE=9,
//   PENALTY=10,DEBIT=11,CREDIT=12,BALANCE=13,ADDON=14,
//   OR=15,REMARKS=16,BILL_NO=17,PAY_DATE=18
//
// DL column constants (0-based):
//   UNIT=0,YEAR=1,MONTH=2,PAY_DATE=3,
//   DEBIT=4,CREDIT=5,BALANCE=6,OR=7,REMARKS=8
// ============================================================

'use strict';

const { createGASContext } = require('../helpers/gas-loader');

// ── WL row helper ─────────────────────────────────────────────
function wlRow(uid, year, month, {
  debit = 0, credit = 0, balance = 0, penalty = 0, addon = 0,
  or_ = '', remarks = '', payDate = ''
} = {}) {
  const row = new Array(19).fill('');
  row[0]  = uid;      // UNIT
  row[1]  = year;     // YEAR
  row[2]  = month;    // MONTH
  row[10] = penalty;  // PENALTY
  row[11] = debit;    // DEBIT
  row[12] = credit;   // CREDIT
  row[13] = balance;  // BALANCE
  row[14] = addon;    // ADDON
  row[15] = or_;      // OR
  row[16] = remarks;  // REMARKS
  row[18] = payDate;  // PAY_DATE
  return row;
}

// ── DL row helper ─────────────────────────────────────────────
function dlRow(uid, year, month, {
  debit = 0, credit = 0, balance = 0, or_ = '', remarks = ''
} = {}) {
  const row = new Array(9).fill('');
  row[0] = uid;
  row[1] = year;
  row[2] = month;
  row[4] = debit;
  row[5] = credit;
  row[6] = balance;
  row[7] = or_;
  row[8] = remarks;
  return row;
}

// ── Shorthand column accessors ────────────────────────────────
const WL = {
  UNIT:0, YEAR:1, MONTH:2, BILL_DATE:3, PREV_DATE:4, PRESENT_DATE:5,
  PREV_RDG:6, CUR_RDG:7, RATE:8, DUE_DATE:9, PENALTY:10,
  DEBIT:11, CREDIT:12, BALANCE:13, ADDON:14,
  OR:15, REMARKS:16, BILL_NO:17, PAY_DATE:18
};
const DL = { UNIT:0, YEAR:1, MONTH:2, PAY_DATE:3, DEBIT:4, CREDIT:5, BALANCE:6, OR:7, REMARKS:8 };

// Load context once — _applyWaterCredit / _applyDuesCredit are pure in-memory ops
let ctx;
beforeAll(() => {
  ctx = createGASContext({});
});

// ═════════════════════════════════════════════════════════════
// _applyWaterCredit
// ═════════════════════════════════════════════════════════════
describe('_applyWaterCredit', () => {
  const OR_NUM  = 'OR-001';
  const PAY_DATE = new Date(2025, 2, 15); // March 15, 2025
  const REMARKS  = 'Test payment';

  it('happy path: exact payment for one bill (credit = debit + penalty + addon)', () => {
    const uid = 'P1B1L1';
    // Unpaid: debit=300, penalty=0, addon=0 → unpaid=300
    const data = [wlRow(uid, 2025, 'January', { debit: 300 })];
    ctx._applyWaterCredit(data, uid, 300, PAY_DATE, OR_NUM, REMARKS);

    expect(data[0][WL.CREDIT]).toBe(300);
    expect(data[0][WL.OR]).toBe(OR_NUM);
    expect(data[0][WL.PAY_DATE]).toBeTruthy();
    expect(data[0][WL.REMARKS]).toContain('Pmt 1 of 1');
  });

  it('partial payment: amount < unpaid → credits as much as possible', () => {
    const uid = 'P1B1L1';
    const data = [wlRow(uid, 2025, 'January', { debit: 300 })];
    ctx._applyWaterCredit(data, uid, 150, PAY_DATE, OR_NUM, REMARKS);

    expect(data[0][WL.CREDIT]).toBe(150);
    // Bill not fully paid
  });

  it('multi-bill payment: oldest-first, remainder applied to next bill', () => {
    const uid = 'P1B1L1';
    // Two months, each with debit=300 unpaid
    const data = [
      wlRow(uid, 2025, 'January',  { debit: 300 }),
      wlRow(uid, 2025, 'February', { debit: 300 }),
    ];
    // Pay 400: January gets 300, February gets 100
    ctx._applyWaterCredit(data, uid, 400, PAY_DATE, OR_NUM, REMARKS);

    expect(data[0][WL.CREDIT]).toBe(300); // January fully paid
    expect(data[1][WL.CREDIT]).toBe(100); // February partially paid
  });

  it('overpayment: excess added to LAST row of unit as advance credit', () => {
    const uid = 'P1B1L1';
    const data = [wlRow(uid, 2025, 'January', { debit: 300 })];
    // Pay 500: 300 clears the bill, 200 is overpayment
    ctx._applyWaterCredit(data, uid, 500, PAY_DATE, OR_NUM, REMARKS);

    // Credit on last row = 300 (from main apply) + 200 (overpayment) = 500
    expect(data[0][WL.CREDIT]).toBe(500);
    expect(data[0][WL.REMARKS]).toContain('Advance');
  });

  it('skips already-paid rows (debit + penalty + addon - credit ≤ 0)', () => {
    const uid = 'P1B1L1';
    const data = [
      wlRow(uid, 2025, 'January',  { debit: 300, credit: 300 }), // already paid
      wlRow(uid, 2025, 'February', { debit: 250 }),               // unpaid
    ];
    ctx._applyWaterCredit(data, uid, 250, PAY_DATE, OR_NUM, REMARKS);

    // January credit must stay at 300 (not increased)
    expect(data[0][WL.CREDIT]).toBe(300);
    // February should receive the 250
    expect(data[1][WL.CREDIT]).toBe(250);
  });

  it('sets PAY_DATE on all affected rows', () => {
    const uid = 'P1B1L1';
    const data = [
      wlRow(uid, 2025, 'January',  { debit: 300 }),
      wlRow(uid, 2025, 'February', { debit: 300 }),
    ];
    ctx._applyWaterCredit(data, uid, 600, PAY_DATE, OR_NUM, REMARKS);

    expect(data[0][WL.PAY_DATE]).toBeTruthy();
    expect(data[1][WL.PAY_DATE]).toBeTruthy();
  });

  it('OR number is appended if different from existing OR', () => {
    const uid = 'P1B1L1';
    const data = [wlRow(uid, 2025, 'January', { debit: 300, or_: 'OR-000' })];
    ctx._applyWaterCredit(data, uid, 300, PAY_DATE, 'OR-001', REMARKS);

    // Existing 'OR-000' is different from 'OR-001' → both kept
    expect(data[0][WL.OR]).toBe('OR-000 / OR-001');
  });

  it('OR number is not duplicated if already the same', () => {
    const uid = 'P1B1L1';
    const data = [wlRow(uid, 2025, 'January', { debit: 300, or_: 'OR-001' })];
    ctx._applyWaterCredit(data, uid, 300, PAY_DATE, 'OR-001', REMARKS);

    expect(data[0][WL.OR]).toBe('OR-001');
  });

  it('remarks include "Pmt N of M" annotation', () => {
    const uid = 'P1B1L1';
    const data = [
      wlRow(uid, 2025, 'January',  { debit: 300 }),
      wlRow(uid, 2025, 'February', { debit: 300 }),
    ];
    ctx._applyWaterCredit(data, uid, 600, PAY_DATE, OR_NUM, REMARKS);

    expect(data[0][WL.REMARKS]).toContain('Pmt 1 of 2');
    expect(data[1][WL.REMARKS]).toContain('Pmt 2 of 2');
  });

  it('zero amount does nothing', () => {
    const uid = 'P1B1L1';
    const data = [wlRow(uid, 2025, 'January', { debit: 300 })];
    ctx._applyWaterCredit(data, uid, 0, PAY_DATE, OR_NUM, REMARKS);

    // No credit applied
    expect(data[0][WL.CREDIT]).toBe(0);
    expect(data[0][WL.OR]).toBe('');
  });

  it('unit not found → no changes', () => {
    const data = [wlRow('P1B1L1', 2025, 'January', { debit: 300 })];
    ctx._applyWaterCredit(data, 'P9B9L9', 300, PAY_DATE, OR_NUM, REMARKS);

    // Overpayment also goes to last row of P9B9L9 which doesn't exist
    // so data[0] is unchanged
    expect(data[0][WL.CREDIT]).toBe(0);
  });

  it('includes debit, penalty, and addon in unpaid calculation', () => {
    const uid = 'P1B1L1';
    // unpaid = 300 (debit) + 15 (penalty) + 50 (addon) = 365
    const data = [wlRow(uid, 2025, 'January', { debit: 300, penalty: 15, addon: 50 })];
    ctx._applyWaterCredit(data, uid, 365, PAY_DATE, OR_NUM, REMARKS);

    expect(data[0][WL.CREDIT]).toBe(365);
  });
});

// ═════════════════════════════════════════════════════════════
// _applyDuesCredit
// ═════════════════════════════════════════════════════════════
describe('_applyDuesCredit', () => {
  const OR_NUM   = 'OR-D01';
  const PAY_DATE = new Date(2025, 2, 15);
  const REMARKS  = 'Dues payment';

  it('happy path: exact payment for one dues bill', () => {
    const uid = 'P1B1L1';
    const data = [dlRow(uid, 2025, 'January', { debit: 500 })];
    ctx._applyDuesCredit(data, uid, 500, PAY_DATE, OR_NUM, REMARKS);

    expect(data[0][DL.CREDIT]).toBe(500);
    expect(data[0][DL.OR]).toBe(OR_NUM);
    expect(data[0][DL.REMARKS]).toContain('Pmt 1 of 1');
  });

  it('partial payment: amount < unpaid', () => {
    const uid = 'P1B1L1';
    const data = [dlRow(uid, 2025, 'January', { debit: 500 })];
    ctx._applyDuesCredit(data, uid, 200, PAY_DATE, OR_NUM, REMARKS);

    expect(data[0][DL.CREDIT]).toBe(200);
  });

  it('multi-bill: oldest first, remainder to next', () => {
    const uid = 'P1B1L1';
    const data = [
      dlRow(uid, 2025, 'January',  { debit: 500 }),
      dlRow(uid, 2025, 'February', { debit: 500 }),
    ];
    ctx._applyDuesCredit(data, uid, 700, PAY_DATE, OR_NUM, REMARKS);

    expect(data[0][DL.CREDIT]).toBe(500);
    expect(data[1][DL.CREDIT]).toBe(200);
  });

  it('overpayment: excess credited to last row', () => {
    const uid = 'P1B1L1';
    const data = [dlRow(uid, 2025, 'January', { debit: 500 })];
    ctx._applyDuesCredit(data, uid, 800, PAY_DATE, OR_NUM, REMARKS);

    // 500 (bill) + 300 (overpayment) = 800 on the last (only) row
    expect(data[0][DL.CREDIT]).toBe(800);
    expect(data[0][DL.REMARKS]).toContain('Advance');
  });

  it('skips already-paid dues rows', () => {
    const uid = 'P1B1L1';
    const data = [
      dlRow(uid, 2025, 'January',  { debit: 500, credit: 500 }), // paid
      dlRow(uid, 2025, 'February', { debit: 500 }),               // unpaid
    ];
    ctx._applyDuesCredit(data, uid, 500, PAY_DATE, OR_NUM, REMARKS);

    expect(data[0][DL.CREDIT]).toBe(500); // unchanged
    expect(data[1][DL.CREDIT]).toBe(500); // paid
  });

  it('sets PAY_DATE on affected rows', () => {
    const uid = 'P1B1L1';
    const data = [dlRow(uid, 2025, 'January', { debit: 500 })];
    ctx._applyDuesCredit(data, uid, 500, PAY_DATE, OR_NUM, REMARKS);

    expect(data[0][DL.PAY_DATE]).toBeTruthy();
  });

  it('remarks include "Pmt N of M" for multi-month payments', () => {
    const uid = 'P1B1L1';
    const data = [
      dlRow(uid, 2025, 'January',  { debit: 500 }),
      dlRow(uid, 2025, 'February', { debit: 500 }),
      dlRow(uid, 2025, 'March',    { debit: 500 }),
    ];
    ctx._applyDuesCredit(data, uid, 1500, PAY_DATE, OR_NUM, REMARKS);

    expect(data[0][DL.REMARKS]).toContain('Pmt 1 of 3');
    expect(data[1][DL.REMARKS]).toContain('Pmt 2 of 3');
    expect(data[2][DL.REMARKS]).toContain('Pmt 3 of 3');
  });

  it('zero amount does nothing', () => {
    const uid = 'P1B1L1';
    const data = [dlRow(uid, 2025, 'January', { debit: 500 })];
    ctx._applyDuesCredit(data, uid, 0, PAY_DATE, OR_NUM, REMARKS);

    expect(data[0][DL.CREDIT]).toBe(0);
  });

  it('unit not found → no changes', () => {
    const data = [dlRow('P1B1L1', 2025, 'January', { debit: 500 })];
    ctx._applyDuesCredit(data, 'P9B9L9', 500, PAY_DATE, OR_NUM, REMARKS);

    expect(data[0][DL.CREDIT]).toBe(0);
  });

  it('OR appended when different from existing', () => {
    const uid = 'P1B1L1';
    const data = [dlRow(uid, 2025, 'January', { debit: 500, or_: 'OR-000' })];
    ctx._applyDuesCredit(data, uid, 500, PAY_DATE, 'OR-D01', REMARKS);

    expect(data[0][DL.OR]).toBe('OR-000 / OR-D01');
  });

  it('handles 3 months of dues arrears with single lump sum', () => {
    const uid = 'P1B1L1';
    // 3 months × ₱500 = ₱1500 owed
    const data = [
      dlRow(uid, 2025, 'January',  { debit: 500 }),
      dlRow(uid, 2025, 'February', { debit: 500 }),
      dlRow(uid, 2025, 'March',    { debit: 500 }),
    ];
    ctx._applyDuesCredit(data, uid, 1500, PAY_DATE, OR_NUM, REMARKS);

    expect(data[0][DL.CREDIT]).toBe(500);
    expect(data[1][DL.CREDIT]).toBe(500);
    expect(data[2][DL.CREDIT]).toBe(500);
  });
});
