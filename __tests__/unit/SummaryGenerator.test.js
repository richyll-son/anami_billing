// ============================================================
// __tests__/unit/SummaryGenerator.test.js
// Tests for pure helper functions in SummaryGenerator.gs:
//   sg_isCommonAccount_
//   sg_commonName_
//   sg_getLatestPeriod_
//
// These functions have no sheet access and can be tested with
// a single context load.
// ============================================================

'use strict';

const { createGASContext } = require('../helpers/gas-loader');

// WL constants (0-based) for building WL-row stubs
const WL = { UNIT:0, YEAR:1, MONTH:2 };
const DL = { UNIT:0, YEAR:1, MONTH:2 };

// One-time context load — pure functions, no sheet mutations
let ctx;
beforeAll(() => {
  ctx = createGASContext({});
});

// ── Helper: minimal WL row ────────────────────────────────────
function wlRow(uid, year, month) {
  const row = new Array(19).fill('');
  row[WL.UNIT]  = uid;
  row[WL.YEAR]  = year;
  row[WL.MONTH] = month;
  return row;
}

// ── Helper: minimal DL row ────────────────────────────────────
function dlRow(uid, year, month) {
  const row = new Array(9).fill('');
  row[DL.UNIT]  = uid;
  row[DL.YEAR]  = year;
  row[DL.MONTH] = month;
  return row;
}

// ═════════════════════════════════════════════════════════════
// sg_isCommonAccount_
// ═════════════════════════════════════════════════════════════
describe('sg_isCommonAccount_', () => {
  it('returns true for "GUARDHOUSE"', () => {
    expect(ctx.sg_isCommonAccount_('GUARDHOUSE')).toBe(true);
  });

  it('returns true for "CLUBHOUSE"', () => {
    expect(ctx.sg_isCommonAccount_('CLUBHOUSE')).toBe(true);
  });

  it('returns true for "CHAPEL"', () => {
    expect(ctx.sg_isCommonAccount_('CHAPEL')).toBe(true);
  });

  it('is case-insensitive: "guardhouse" → true', () => {
    expect(ctx.sg_isCommonAccount_('guardhouse')).toBe(true);
  });

  it('strips spaces: "GUARD HOUSE" → true', () => {
    // The function does .replace(/\s+/g,'') before checking
    expect(ctx.sg_isCommonAccount_('GUARD HOUSE')).toBe(true);
  });

  it('returns false for a normal unit P1B1L1', () => {
    expect(ctx.sg_isCommonAccount_('P1B1L1')).toBe(false);
  });

  it('returns false for empty string', () => {
    expect(ctx.sg_isCommonAccount_('')).toBe(false);
  });

  it('returns false for an unknown string', () => {
    expect(ctx.sg_isCommonAccount_('OTHER')).toBe(false);
  });

  it('returns false for null', () => {
    expect(ctx.sg_isCommonAccount_(null)).toBe(false);
  });

  it('returns false for undefined', () => {
    expect(ctx.sg_isCommonAccount_(undefined)).toBe(false);
  });
});

// ═════════════════════════════════════════════════════════════
// sg_commonName_
// ═════════════════════════════════════════════════════════════
describe('sg_commonName_', () => {
  it('"GUARDHOUSE" → "Guardhouse"', () => {
    expect(ctx.sg_commonName_('GUARDHOUSE')).toBe('Guardhouse');
  });

  it('"CLUBHOUSE" → "Clubhouse"', () => {
    expect(ctx.sg_commonName_('CLUBHOUSE')).toBe('Clubhouse');
  });

  it('"CHAPEL" → "Chapel"', () => {
    expect(ctx.sg_commonName_('CHAPEL')).toBe('Chapel');
  });

  it('is case-insensitive for lookup ("clubhouse" → "Clubhouse")', () => {
    // The function uppercases the input before comparing
    expect(ctx.sg_commonName_('clubhouse')).toBe('Clubhouse');
  });

  it('returns "" for an unknown ID', () => {
    expect(ctx.sg_commonName_('P1B1L1')).toBe('');
  });

  it('returns "" for empty string', () => {
    expect(ctx.sg_commonName_('')).toBe('');
  });

  it('returns "" for null', () => {
    expect(ctx.sg_commonName_(null)).toBe('');
  });
});

// ═════════════════════════════════════════════════════════════
// sg_getLatestPeriod_
// ═════════════════════════════════════════════════════════════
describe('sg_getLatestPeriod_', () => {
  it('returns empty string for empty arrays', () => {
    expect(ctx.sg_getLatestPeriod_([], [])).toBe('');
  });

  it('returns "Month YEAR" for a single water row', () => {
    const water = [wlRow('P1B1L1', 2025, 'March')];
    expect(ctx.sg_getLatestPeriod_(water, [])).toBe('March 2025');
  });

  it('returns the highest rank when multiple water rows exist', () => {
    const water = [
      wlRow('P1B1L1', 2025, 'January'),
      wlRow('P1B1L1', 2025, 'March'),
      wlRow('P1B2L1', 2025, 'February'),
    ];
    // March 2025 rank = 2025*100+3 = 202503 is highest
    expect(ctx.sg_getLatestPeriod_(water, [])).toBe('March 2025');
  });

  it('considers dues rows if they are later than water rows', () => {
    const water = [wlRow('P1B1L1', 2025, 'January')];
    const dues  = [dlRow('P1B1L1', 2025, 'April')]; // April > January
    expect(ctx.sg_getLatestPeriod_(water, dues)).toBe('April 2025');
  });

  it('returns water period if it is later than dues', () => {
    const water = [wlRow('P1B1L1', 2025, 'June')];
    const dues  = [dlRow('P1B1L1', 2025, 'February')];
    expect(ctx.sg_getLatestPeriod_(water, dues)).toBe('June 2025');
  });

  it('skips rows with invalid/missing year (parseInt → NaN, y=0, falsy)', () => {
    const water = [
      wlRow('P1B1L1', '',   'March'), // empty year
      wlRow('P1B1L1', 2025, 'January')
    ];
    expect(ctx.sg_getLatestPeriod_(water, [])).toBe('January 2025');
  });

  it('skips rows with invalid/missing month (getMonthNum → 0, falsy)', () => {
    const water = [
      wlRow('P1B1L1', 2025, ''),      // empty month → getMonthNum('') = 0
      wlRow('P1B1L1', 2025, 'June')
    ];
    expect(ctx.sg_getLatestPeriod_(water, [])).toBe('June 2025');
  });

  it('handles year rollover: December 2024 vs January 2025', () => {
    const water = [
      wlRow('P1B1L1', 2024, 'December'), // rank = 202412
      wlRow('P1B1L1', 2025, 'January'),  // rank = 202501
    ];
    expect(ctx.sg_getLatestPeriod_(water, [])).toBe('January 2025');
  });

  it('handles large dataset (100+ rows) without error', () => {
    const water = [];
    for (let m = 1; m <= 12; m++) {
      for (let u = 1; u <= 10; u++) {
        water.push(wlRow(`P1B1L${u}`, 2025, [
          'January','February','March','April','May','June',
          'July','August','September','October','November','December'
        ][m - 1]));
      }
    }
    // All 120 rows, highest is December 2025
    expect(ctx.sg_getLatestPeriod_(water, [])).toBe('December 2025');
  });
});
