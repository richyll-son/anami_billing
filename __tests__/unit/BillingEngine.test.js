// ============================================================
// __tests__/unit/BillingEngine.test.js
// Tests for billing math functions in BillingEngine.gs.
//
// getDuesRate_ and isBODExempt_ are tested with preloaded data
// so no sheet access is needed.
//
// recalcWaterBalances / recalcDuesBalances are tested with
// mock sheet data injected via sheetsMap.
//
// Column constants (0-based):
//   WL: UNIT=0,YEAR=1,MONTH=2,BILL_DATE=3,PREV_DATE=4,
//       PRESENT_DATE=5,PREV_RDG=6,CUR_RDG=7,RATE=8,DUE_DATE=9,
//       PENALTY=10,DEBIT=11,CREDIT=12,BALANCE=13,ADDON=14,
//       OR=15,REMARKS=16,BILL_NO=17,PAY_DATE=18  (WL_COLS=19)
//   DL: UNIT=0,YEAR=1,MONTH=2,PAY_DATE=3,DEBIT=4,
//       CREDIT=5,BALANCE=6,OR=7,REMARKS=8  (DL_COLS=9)
//   DR: AMOUNT=0,FROM_MONTH=1,FROM_YEAR=2,TO_MONTH=3,TO_YEAR=4
//   BOD: NAME=0,POSITION=1,UNIT=2,FROM_MONTH=3,FROM_YEAR=4,
//        TO_MONTH=5,TO_YEAR=6
// ============================================================

'use strict';

const { createGASContext, createMockSheet } = require('../helpers/gas-loader');

// ── Helper: build a minimal WL row (19 cols) ─────────────────
// All unused fields default to '' / 0.
function wlRow(uid, year, month, {
  debit = 0, credit = 0, balance = 0, penalty = 0, addon = 0,
  dueDate = '', payDate = '', rate = 0,
  prevRdg = 0, curRdg = 0
} = {}) {
  const row = new Array(19).fill('');
  row[0]  = uid;      // UNIT
  row[1]  = year;     // YEAR
  row[2]  = month;    // MONTH
  row[9]  = dueDate;  // DUE_DATE
  row[10] = penalty;  // PENALTY
  row[11] = debit;    // DEBIT
  row[12] = credit;   // CREDIT
  row[13] = balance;  // BALANCE (will be overwritten by recalc)
  row[14] = addon;    // ADDON
  row[18] = payDate;  // PAY_DATE
  row[8]  = rate;     // RATE
  row[6]  = prevRdg;  // PREV_RDG
  row[7]  = curRdg;   // CUR_RDG
  return row;
}

// ── Helper: build a minimal DL row (9 cols) ──────────────────
function dlRow(uid, year, month, { debit = 0, credit = 0, balance = 0 } = {}) {
  const row = new Array(9).fill('');
  row[0] = uid;
  row[1] = year;
  row[2] = month;
  row[4] = debit;
  row[5] = credit;
  row[6] = balance;
  return row;
}

// ── Helper: build a _DuesRates row (5 cols) ──────────────────
// DR: AMOUNT=0,FROM_MONTH=1,FROM_YEAR=2,TO_MONTH=3,TO_YEAR=4
function drRow(amount, fromMonth, fromYear, toMonth = '', toYear = '') {
  return [amount, fromMonth, fromYear, toMonth, toYear];
}

// ── Helper: build a _BOD row (7 cols) ────────────────────────
// BOD: NAME=0,POSITION=1,UNIT=2,FROM_MONTH=3,FROM_YEAR=4,TO_MONTH=5,TO_YEAR=6
function bodRow(unit, fromMonth, fromYear, toMonth = '', toYear = '') {
  return ['', '', unit, fromMonth, fromYear, toMonth, toYear];
}

// ═════════════════════════════════════════════════════════════
// getDuesRate_
// ═════════════════════════════════════════════════════════════
describe('getDuesRate_', () => {
  let ctx;

  beforeAll(() => {
    ctx = createGASContext({});
  });

  it('returns ASSOC_DUES (500) when preloadedData is an empty array', () => {
    // No custom rates configured → fallback to constant
    expect(ctx.getDuesRate_(2025, 3, [])).toBe(500);
  });

  it('returns the matching rate when billing period is within range', () => {
    // Rate of 600 active from January 2025 to December 2025
    const data = [drRow(600, 'January', 2025, 'December', 2025)];
    expect(ctx.getDuesRate_(2025, 6, data)).toBe(600);
  });

  it('returns rate when toMonth/toYear are empty (open-ended → toRank=999999)', () => {
    // Open-ended rate: applies forever from January 2025 onward
    const data = [drRow(750, 'January', 2025, '', '')];
    expect(ctx.getDuesRate_(2030, 1, data)).toBe(750);
  });

  it('last matching row wins when multiple rows match', () => {
    // Two overlapping rates — the later one in the array wins
    const data = [
      drRow(600, 'January', 2025, 'December', 2025),
      drRow(700, 'January', 2025, 'June', 2025)  // narrower range, still matches March 2025
    ];
    // Both rows match 2025-03; the last one in iteration order wins
    expect(ctx.getDuesRate_(2025, 3, data)).toBe(700);
  });

  it('falls back to most-recent row amount when no match', () => {
    // Rates only for 2024; billing is in 2025 — no match → last positive amount
    const data = [
      drRow(400, 'January', 2024, 'December', 2024)
    ];
    expect(ctx.getDuesRate_(2025, 3, data)).toBe(400);
  });

  it('returns ASSOC_DUES when all amounts are 0', () => {
    const data = [drRow(0, 'January', 2025, 'December', 2025)];
    // No match AND no positive amount in fallback → ASSOC_DUES
    // The rate is 0, which is not > 0, so fallback loop skips it
    expect(ctx.getDuesRate_(2026, 1, data)).toBe(500);
  });

  it('returns ASSOC_DUES when preloadedData is null (triggers sheet read path)', () => {
    // When null is passed, it tries to read from sheet; sheet returns null → returns ASSOC_DUES
    // We need a context where the sheet does not exist
    const ctxNoSheet = createGASContext({});
    expect(ctxNoSheet.getDuesRate_(2025, 3, null)).toBe(500);
  });
});

// ═════════════════════════════════════════════════════════════
// isBODExempt_
// ═════════════════════════════════════════════════════════════
describe('isBODExempt_', () => {
  let ctx;

  beforeAll(() => {
    ctx = createGASContext({});
  });

  it('returns false for empty data', () => {
    expect(ctx.isBODExempt_('P1B1L1', 2025, 3, [])).toBe(false);
  });

  it('returns true when unit billing period is within exempt range', () => {
    // Unit is BOD from January 2025 to December 2025
    const data = [bodRow('P1B1L1', 'January', 2025, 'December', 2025)];
    expect(ctx.isBODExempt_('P1B1L1', 2025, 6, data)).toBe(true);
  });

  it('returns false when billing period is before exemption start', () => {
    // Exemption starts March 2025, billing is February 2025
    const data = [bodRow('P1B1L1', 'March', 2025, 'December', 2025)];
    expect(ctx.isBODExempt_('P1B1L1', 2025, 2, data)).toBe(false);
  });

  it('returns false when billing period is after exemption end', () => {
    // Exemption ends December 2025, billing is January 2026
    const data = [bodRow('P1B1L1', 'January', 2025, 'December', 2025)];
    expect(ctx.isBODExempt_('P1B1L1', 2026, 1, data)).toBe(false);
  });

  it('returns false for a different unit not in BOD data', () => {
    const data = [bodRow('P1B1L1', 'January', 2025, 'December', 2025)];
    expect(ctx.isBODExempt_('P2B3L5', 2025, 6, data)).toBe(false);
  });

  it('is case-insensitive for unit ID matching', () => {
    // Stored as uppercase, queried as lowercase → should still match
    const data = [bodRow('P1B1L1', 'January', 2025, 'December', 2025)];
    expect(ctx.isBODExempt_('p1b1l1', 2025, 6, data)).toBe(true);
  });

  it('returns true for open-ended exemption (no toMonth/toYear)', () => {
    // No end date → toRank=999999 → matches any future period
    const data = [bodRow('P1B2L3', 'January', 2025, '', '')];
    expect(ctx.isBODExempt_('P1B2L3', 2099, 12, data)).toBe(true);
  });

  it('returns false when null preloadedData and no BOD sheet exists', () => {
    const ctxNoSheet = createGASContext({});
    expect(ctxNoSheet.isBODExempt_('P1B1L1', 2025, 3, null)).toBe(false);
  });
});

// ═════════════════════════════════════════════════════════════
// recalcWaterBalances
// ═════════════════════════════════════════════════════════════
describe('recalcWaterBalances', () => {
  // WL columns (0-based): UNIT=0,YEAR=1,MONTH=2,...,PENALTY=10,DEBIT=11,
  // CREDIT=12,BALANCE=13,ADDON=14,...
  // Sheet has 1 header row + data rows.

  function makeWLSheet(dataRows) {
    // header + data rows for _WaterLedger
    return [['UNIT', 'YEAR', 'MONTH', ...new Array(16).fill('')]].concat(dataRows);
  }

  it('single unit, single row: balance preserved as opening (first row rule)', () => {
    // The first row for a unit always keeps its stored balance as opening.
    // The recalc loop preserves it and sets balCol[0] = stored balance.
    const uid = 'P1B1L1';
    const row = wlRow(uid, 2025, 'January', { debit: 300, credit: 0, balance: 300, penalty: 0, addon: 0 });
    const ctx = createGASContext({ '_WaterLedger': makeWLSheet([row]) });

    ctx.recalcWaterBalances();

    const sheet = ctx._mockSS.getSheetByName('_WaterLedger');
    // Balance column (14, 1-based) of data row 1 = the stored opening balance
    const written = sheet._data[1][13]; // row index 1 = first data row
    expect(written).toBe(300);
  });

  it('single unit, multiple rows: balance = prev + debit + penalty + addon - credit', () => {
    const uid = 'P1B1L1';
    // Row 1 (opening): stored balance = 300
    // Row 2: debit=250, penalty=15, addon=50, credit=0
    //   expected balance = 300 + 250 + 15 + 50 - 0 = 615
    const rows = [
      wlRow(uid, 2025, 'January',  { debit: 300, credit: 0,   balance: 300, penalty: 0,  addon: 0  }),
      wlRow(uid, 2025, 'February', { debit: 250, credit: 0,   balance: 0,   penalty: 15, addon: 50 }),
    ];
    const ctx = createGASContext({ '_WaterLedger': makeWLSheet(rows) });
    ctx.recalcWaterBalances();

    const sheet = ctx._mockSS.getSheetByName('_WaterLedger');
    const balance2 = sheet._data[2][13]; // third row in sheet (index 2)
    expect(balance2).toBe(615);
  });

  it('single unit, three rows: running balance chains correctly', () => {
    const uid = 'P1B1L1';
    // Opening: 300; Feb: +250 debit, -300 credit → 250; Mar: +250 debit → 500
    const rows = [
      wlRow(uid, 2025, 'January',  { debit: 300, credit: 0,   balance: 300 }),
      wlRow(uid, 2025, 'February', { debit: 250, credit: 300, balance: 0   }),
      wlRow(uid, 2025, 'March',    { debit: 250, credit: 0,   balance: 0   }),
    ];
    const ctx = createGASContext({ '_WaterLedger': makeWLSheet(rows) });
    ctx.recalcWaterBalances();

    const sheet = ctx._mockSS.getSheetByName('_WaterLedger');
    // After Jan: 300
    // After Feb: 300 + 250 - 300 = 250
    // After Mar: 250 + 250 = 500
    expect(sheet._data[1][13]).toBe(300); // Jan (opening)
    expect(sheet._data[2][13]).toBe(250); // Feb
    expect(sheet._data[3][13]).toBe(500); // Mar
  });

  it('multiple units: independent running balances', () => {
    const rowA1 = wlRow('P1B1L1', 2025, 'January', { debit: 400, balance: 400 });
    const rowB1 = wlRow('P1B2L1', 2025, 'January', { debit: 300, balance: 300 });
    const rowA2 = wlRow('P1B1L1', 2025, 'February', { debit: 200, balance: 0 });
    const rowB2 = wlRow('P1B2L1', 2025, 'February', { debit: 100, credit: 300, balance: 0 });

    const ctx = createGASContext({ '_WaterLedger': makeWLSheet([rowA1, rowB1, rowA2, rowB2]) });
    ctx.recalcWaterBalances();

    const sheet = ctx._mockSS.getSheetByName('_WaterLedger');
    // P1B1L1 Feb: 400 + 200 = 600
    // P1B2L1 Feb: 300 + 100 - 300 = 100
    expect(sheet._data[3][13]).toBe(600); // A2
    expect(sheet._data[4][13]).toBe(100); // B2
  });

  it('allows negative balance (credit carry-forward, not clamped)', () => {
    const uid = 'P1B1L1';
    const rows = [
      wlRow(uid, 2025, 'January',  { debit: 300, credit: 0,   balance: 300 }),
      wlRow(uid, 2025, 'February', { debit: 250, credit: 700, balance: 0   }),
    ];
    const ctx = createGASContext({ '_WaterLedger': makeWLSheet(rows) });
    ctx.recalcWaterBalances();

    const sheet = ctx._mockSS.getSheetByName('_WaterLedger');
    // Feb: 300 + 250 - 700 = -150 (negative = credit carry-forward, intentional)
    expect(sheet._data[2][13]).toBe(-150);
  });

  it('empty uid rows receive balance 0', () => {
    const rowEmpty = new Array(19).fill(''); // uid = ''
    rowEmpty[11] = 500; // debit
    const ctx = createGASContext({ '_WaterLedger': makeWLSheet([rowEmpty]) });
    ctx.recalcWaterBalances();

    const sheet = ctx._mockSS.getSheetByName('_WaterLedger');
    expect(sheet._data[1][13]).toBe(0);
  });

  it('rounds balance to 2 decimal places', () => {
    const uid = 'P1B1L1';
    const rows = [
      wlRow(uid, 2025, 'January',  { debit: 100.001, balance: 100.001 }),
      wlRow(uid, 2025, 'February', { debit: 50.007,  balance: 0 }),
    ];
    const ctx = createGASContext({ '_WaterLedger': makeWLSheet(rows) });
    ctx.recalcWaterBalances();

    const sheet = ctx._mockSS.getSheetByName('_WaterLedger');
    const feb = sheet._data[2][13];
    // 100.001 + 50.007 = 150.008 → rounded → 150.01
    expect(feb).toBe(150.01);
  });

  it('does nothing when sheet has fewer than 2 rows', () => {
    // Only header row → getLastRow() returns 1, which is < 2, so function returns early
    const ctx = createGASContext({ '_WaterLedger': [['UNIT', 'YEAR']] });
    expect(() => ctx.recalcWaterBalances()).not.toThrow();
  });
});

// ═════════════════════════════════════════════════════════════
// recalcDuesBalances
// ═════════════════════════════════════════════════════════════
describe('recalcDuesBalances', () => {
  function makeDLSheet(dataRows) {
    return [['UNIT', 'YEAR', 'MONTH', 'PAY_DATE', 'DEBIT', 'CREDIT', 'BALANCE', 'OR', 'REMARKS']]
      .concat(dataRows);
  }

  it('single unit running balance: balance = prev + debit - credit', () => {
    const uid = 'P1B1L1';
    const rows = [
      dlRow(uid, 2025, 'January',  { debit: 500, credit: 0,   balance: 0 }),
      dlRow(uid, 2025, 'February', { debit: 500, credit: 500, balance: 0 }),
      dlRow(uid, 2025, 'March',    { debit: 500, credit: 0,   balance: 0 }),
    ];
    const ctx = createGASContext({ '_DuesLedger': makeDLSheet(rows) });
    ctx.recalcDuesBalances();

    const sheet = ctx._mockSS.getSheetByName('_DuesLedger');
    // Jan:  0 + 500 - 0   = 500
    // Feb: 500 + 500 - 500 = 500
    // Mar: 500 + 500 - 0   = 1000
    expect(sheet._data[1][6]).toBe(500);
    expect(sheet._data[2][6]).toBe(500);
    expect(sheet._data[3][6]).toBe(1000);
  });

  it('multiple units: independent running balances', () => {
    const rowsA = [
      dlRow('P1B1L1', 2025, 'January', { debit: 500, credit: 0 }),
      dlRow('P1B1L1', 2025, 'February', { debit: 500, credit: 500 }),
    ];
    const rowsB = [
      dlRow('P1B2L1', 2025, 'January', { debit: 500, credit: 0 }),
      dlRow('P1B2L1', 2025, 'February', { debit: 500, credit: 0 }),
    ];
    const ctx = createGASContext({ '_DuesLedger': makeDLSheet([...rowsA, ...rowsB]) });
    ctx.recalcDuesBalances();

    const sheet = ctx._mockSS.getSheetByName('_DuesLedger');
    // A Feb: 500 + 500 - 500 = 500
    // B Feb: 500 + 500 - 0   = 1000
    expect(sheet._data[2][6]).toBe(500);
    expect(sheet._data[4][6]).toBe(1000);
  });

  it('does not clamp negative balance (credit carry-forward allowed)', () => {
    const uid = 'P1B1L1';
    const rows = [
      dlRow(uid, 2025, 'January',  { debit: 500, credit: 0   }),
      dlRow(uid, 2025, 'February', { debit: 500, credit: 1500 }),
    ];
    const ctx = createGASContext({ '_DuesLedger': makeDLSheet(rows) });
    ctx.recalcDuesBalances();

    const sheet = ctx._mockSS.getSheetByName('_DuesLedger');
    // Jan: 500; Feb: 500 + 500 - 1500 = -500 (allowed)
    expect(sheet._data[2][6]).toBe(-500);
  });

  it('does nothing when sheet has fewer than 2 rows', () => {
    const ctx = createGASContext({ '_DuesLedger': [['UNIT', 'YEAR']] });
    expect(() => ctx.recalcDuesBalances()).not.toThrow();
  });

  it('rounds dues balance to 2 decimal places', () => {
    const uid = 'P1B1L1';
    const rows = [
      dlRow(uid, 2025, 'January', { debit: 333.333, credit: 0 }),
      dlRow(uid, 2025, 'February', { debit: 333.333, credit: 0 }),
    ];
    const ctx = createGASContext({ '_DuesLedger': makeDLSheet(rows) });
    ctx.recalcDuesBalances();

    const sheet = ctx._mockSS.getSheetByName('_DuesLedger');
    // Cumulative rounding:
    //   Jan: Math.round(333.333 * 100) / 100 = 333.33 (stored as prev for Feb)
    //   Feb: Math.round((333.33 + 333.333) * 100) / 100 = Math.round(66663.3) / 100 = 666.66
    // (prev is the already-rounded value 333.33, not the raw 333.333)
    expect(sheet._data[2][6]).toBe(666.66);
  });
});
