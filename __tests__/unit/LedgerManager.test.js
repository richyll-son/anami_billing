// ============================================================
// __tests__/unit/LedgerManager.test.js
// Smoke tests for LedgerManager.gs's _writeBackWater / _writeBackDues.
//
// Regression coverage for a bug found during manual QA: the water
// write-back recalc trigger checked display columns 9/10/13
// (Penalty/Debit/Addon) instead of 10/11/13 (Debit/Credit/Addon),
// so editing the Credit cell directly in the Unit Ledger sheet
// silently never recomputed the running balance.
// ============================================================

'use strict';

const { createGASContext } = require('../helpers/gas-loader');

// WL column constants (0-based), mirrors Code.gs
const WL = {
  UNIT:0, YEAR:1, MONTH:2, BILL_DATE:3, PREV_DATE:4, PRESENT_DATE:5,
  PREV_RDG:6, CUR_RDG:7, RATE:8, DUE_DATE:9, PENALTY:10,
  DEBIT:11, CREDIT:12, BALANCE:13, ADDON:14,
  OR:15, REMARKS:16, BILL_NO:17, PAY_DATE:18
};
const DL = { UNIT:0, YEAR:1, MONTH:2, PAY_DATE:3, DEBIT:4, CREDIT:5, BALANCE:6, OR:7, REMARKS:8 };

function wlRow(uid, year, month, { debit = 0, credit = 0, balance = 0, penalty = 0, addon = 0 } = {}) {
  const row = new Array(19).fill('');
  row[WL.UNIT] = uid; row[WL.YEAR] = year; row[WL.MONTH] = month;
  row[WL.PENALTY] = penalty; row[WL.DEBIT] = debit; row[WL.CREDIT] = credit;
  row[WL.BALANCE] = balance; row[WL.ADDON] = addon;
  return row;
}

const WL_HEADER = new Array(19).fill('');
const UID = 'P1B1L1';

// Unit Ledger display sheet mock — only B1 (unit selector) matters for write-back
function ulSheetData() {
  const row1 = [];
  row1[1] = UID; // B1 (0-based col index 1)
  return [row1];
}

describe('_writeBackWater', () => {
  function setup() {
    return createGASContext({
      'Unit Ledger': ulSheetData(),
      '_WaterLedger': [
        WL_HEADER,
        wlRow(UID, 2025, 'January', { balance: 200 }),          // opening row
        wlRow(UID, 2025, 'February', { debit: 300, balance: 0 }) // to be recalculated
      ]
    });
  }

  it('editing Credit (display col 11) writes the value AND triggers recalcWaterBalances', () => {
    const ctx = setup();
    const ulSheet = ctx.getSheet_(ctx.SH.UNIT_LEDGER);

    // Row index 1 = second row for this unit (February), pay off the 300 debit
    ctx._writeBackWater(ulSheet, 1, 11, 300);

    const wlSheet = ctx.getSheet_(ctx.SH._WL);
    const febRow = wlSheet._data[2]; // header + Jan + Feb

    expect(febRow[WL.CREDIT]).toBe(300);
    // Balance must be recomputed: 200 (prev) + 300 (debit) + 0 (penalty) + 0 (addon) - 300 (credit) = 200
    expect(febRow[WL.BALANCE]).toBe(200);
  });

  it('editing Debit (display col 10) still triggers recalcWaterBalances (regression guard)', () => {
    const ctx = setup();
    const ulSheet = ctx.getSheet_(ctx.SH.UNIT_LEDGER);

    ctx._writeBackWater(ulSheet, 1, 10, 500); // bump Feb debit 300 → 500

    const wlSheet = ctx.getSheet_(ctx.SH._WL);
    const febRow = wlSheet._data[2];

    expect(febRow[WL.DEBIT]).toBe(500);
    // 200 (prev) + 500 (debit) + 0 + 0 - 0 (credit still 0) = 700
    expect(febRow[WL.BALANCE]).toBe(700);
  });

  it('editing Add-On MCWD (display col 13) still triggers recalcWaterBalances', () => {
    const ctx = setup();
    const ulSheet = ctx.getSheet_(ctx.SH.UNIT_LEDGER);

    ctx._writeBackWater(ulSheet, 1, 13, 50);

    const wlSheet = ctx.getSheet_(ctx.SH._WL);
    const febRow = wlSheet._data[2];

    expect(febRow[WL.ADDON]).toBe(50);
    // 200 (prev) + 300 (debit) + 0 (penalty) + 50 (addon) - 0 (credit) = 550
    expect(febRow[WL.BALANCE]).toBe(550);
  });

  it('editing Balance (display col 12, read-only) does not write anything', () => {
    const ctx = setup();
    const ulSheet = ctx.getSheet_(ctx.SH.UNIT_LEDGER);

    ctx._writeBackWater(ulSheet, 1, 12, 9999);

    const wlSheet = ctx.getSheet_(ctx.SH._WL);
    const febRow = wlSheet._data[2];

    expect(febRow[WL.BALANCE]).toBe(0); // unchanged from initial data (no recalc ran)
  });
});

describe('_writeBackDues', () => {
  function dlRow(uid, year, month, { debit = 0, credit = 0, balance = 0 } = {}) {
    const row = new Array(9).fill('');
    row[DL.UNIT] = uid; row[DL.YEAR] = year; row[DL.MONTH] = month;
    row[DL.DEBIT] = debit; row[DL.CREDIT] = credit; row[DL.BALANCE] = balance;
    return row;
  }
  const DL_HEADER = new Array(9).fill('');

  it('editing Credit (display col 4) triggers recalcDuesBalances', () => {
    // recalcDuesBalances (unlike recalcWaterBalances) has no "first row
    // preserves stored opening balance" rule — every row is recomputed
    // as prev + debit - credit from unitBal starting at 0. Single-row
    // scenario avoids relying on that non-existent special case.
    const ctx = createGASContext({
      'Unit Ledger': ulSheetData(),
      '_DuesLedger': [
        DL_HEADER,
        dlRow(UID, 2025, 'January', { debit: 500, balance: 500 })
      ]
    });
    const ulSheet = ctx.getSheet_(ctx.SH.UNIT_LEDGER);

    ctx._writeBackDues(ulSheet, 0, 4, 500);

    const dlSheet = ctx.getSheet_(ctx.SH._DL);
    const janRow = dlSheet._data[1];

    expect(janRow[DL.CREDIT]).toBe(500);
    // 0 (prev) + 500 (debit) - 500 (credit) = 0
    expect(janRow[DL.BALANCE]).toBe(0);
  });
});
