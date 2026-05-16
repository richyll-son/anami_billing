// ============================================================
// __tests__/unit/Code.test.js
// Tests for all pure utility functions defined in Code.gs.
// These functions have no side-effects and no sheet access.
// We load them once into a GAS context and call them directly.
// ============================================================

'use strict';

const { createGASContext } = require('../helpers/gas-loader');

// Load GAS context once for this file — pure functions are stateless
let ctx;

beforeAll(() => {
  ctx = createGASContext({});
});

// ── fmt2 ─────────────────────────────────────────────────────
describe('fmt2', () => {
  it('formats a positive integer to 2 decimal places', () => {
    expect(ctx.fmt2(100)).toBe('100.00');
  });

  it('formats a positive float to 2 decimal places', () => {
    expect(ctx.fmt2(123.456)).toBe('123.46'); // toFixed rounds
  });

  it('formats zero correctly', () => {
    expect(ctx.fmt2(0)).toBe('0.00');
  });

  it('formats a negative number', () => {
    expect(ctx.fmt2(-50.5)).toBe('-50.50');
  });

  it('returns "0.00" for NaN', () => {
    expect(ctx.fmt2(NaN)).toBe('0.00');
  });

  it('returns "0.00" for undefined', () => {
    expect(ctx.fmt2(undefined)).toBe('0.00');
  });

  it('returns "0.00" for null', () => {
    expect(ctx.fmt2(null)).toBe('0.00');
  });

  it('parses a numeric string before formatting', () => {
    expect(ctx.fmt2('250')).toBe('250.00');
  });

  it('returns "0.00" for non-numeric string', () => {
    expect(ctx.fmt2('hello')).toBe('0.00');
  });
});

// ── fmtDate ──────────────────────────────────────────────────
describe('fmtDate', () => {
  it('formats a Date object as MM/DD/YYYY', () => {
    // Use UTC-safe construction to avoid timezone issues in CI
    const d = new Date(2025, 2, 15); // March 15, 2025 (month is 0-based)
    expect(ctx.fmtDate(d)).toBe('03/15/2025');
  });

  it('passes through a string value unchanged', () => {
    expect(ctx.fmtDate('01/15/2025')).toBe('01/15/2025');
  });

  it('returns empty string for null', () => {
    expect(ctx.fmtDate(null)).toBe('');
  });

  it('returns empty string for undefined', () => {
    expect(ctx.fmtDate(undefined)).toBe('');
  });

  it('returns empty string for empty string', () => {
    expect(ctx.fmtDate('')).toBe('');
  });

  it('pads single-digit month and day', () => {
    const d = new Date(2025, 0, 5); // January 5, 2025
    expect(ctx.fmtDate(d)).toBe('01/05/2025');
  });

  it('handles December correctly (month 11 → padded as 12)', () => {
    const d = new Date(2025, 11, 31); // December 31, 2025
    expect(ctx.fmtDate(d)).toBe('12/31/2025');
  });
});

// ── parseDate ─────────────────────────────────────────────────
describe('parseDate', () => {
  it('parses a valid date string', () => {
    const result = ctx.parseDate('2025-03-15');
    expect(result).toBeInstanceOf(Date);
    expect(result.getFullYear()).toBe(2025);
  });

  it('returns a copy of a Date object', () => {
    const original = new Date(2025, 2, 15);
    const result = ctx.parseDate(original);
    expect(result).toBeInstanceOf(Date);
    expect(result.getTime()).toBe(original.getTime());
    // Must be a copy, not the same reference
    expect(result).not.toBe(original);
  });

  it('returns null for an invalid string', () => {
    expect(ctx.parseDate('not-a-date')).toBeNull();
  });

  it('returns null for null input', () => {
    expect(ctx.parseDate(null)).toBeNull();
  });

  it('returns null for undefined input', () => {
    expect(ctx.parseDate(undefined)).toBeNull();
  });

  it('returns null for empty string', () => {
    expect(ctx.parseDate('')).toBeNull();
  });

  it('parses MM/DD/YYYY format string', () => {
    const result = ctx.parseDate('01/15/2025');
    expect(result).toBeInstanceOf(Date);
    expect(isNaN(result.getTime())).toBe(false);
  });
});

// ── toNum ─────────────────────────────────────────────────────
describe('toNum', () => {
  it('returns a number as-is', () => {
    expect(ctx.toNum(42)).toBe(42);
  });

  it('parses a numeric string', () => {
    expect(ctx.toNum('3.14')).toBe(3.14);
  });

  it('returns 0 for empty string', () => {
    expect(ctx.toNum('')).toBe(0);
  });

  it('returns 0 for null', () => {
    expect(ctx.toNum(null)).toBe(0);
  });

  it('returns 0 for undefined', () => {
    expect(ctx.toNum(undefined)).toBe(0);
  });

  it('returns 0 for non-numeric string', () => {
    expect(ctx.toNum('hello')).toBe(0);
  });

  it('returns 0 for NaN', () => {
    expect(ctx.toNum(NaN)).toBe(0);
  });

  it('handles negative numbers', () => {
    expect(ctx.toNum(-100)).toBe(-100);
  });
});

// ── buildUID ─────────────────────────────────────────────────
describe('buildUID', () => {
  it('builds a standard UID', () => {
    expect(ctx.buildUID(1, 2, 3)).toBe('P1B2L3');
  });

  it('handles double-digit block and lot', () => {
    expect(ctx.buildUID(2, 10, 25)).toBe('P2B10L25');
  });

  it('handles combined lot notation', () => {
    // buildUID just concatenates — combined lot is stored as-is
    expect(ctx.buildUID(1, 6, '1&2')).toBe('P1B6L1&2');
  });

  it('handles phase 2', () => {
    expect(ctx.buildUID(2, 1, 1)).toBe('P2B1L1');
  });
});

// ── parseUID ─────────────────────────────────────────────────
describe('parseUID', () => {
  it('parses a standard UID into phase, block, lot', () => {
    const result = ctx.parseUID('P1B2L3');
    expect(result).toEqual({ phase: 1, block: 2, lot: '3' });
  });

  it('parses a UID with double-digit numbers', () => {
    const result = ctx.parseUID('P2B10L25');
    expect(result).toEqual({ phase: 2, block: 10, lot: '25' });
  });

  it('parses a combined lot UID', () => {
    const result = ctx.parseUID('P1B6L1&2');
    expect(result).toEqual({ phase: 1, block: 6, lot: '1&2' });
  });

  it('returns null for invalid format', () => {
    expect(ctx.parseUID('GUARDHOUSE')).toBeNull();
  });

  it('returns null for empty string', () => {
    expect(ctx.parseUID('')).toBeNull();
  });

  it('returns null for undefined', () => {
    expect(ctx.parseUID(undefined)).toBeNull();
  });
});

// ── getBillDate ───────────────────────────────────────────────
describe('getBillDate', () => {
  it('returns the 10th of the next month for a mid-year reading', () => {
    // Reading month: October (10) → Bill date: November 10
    const d = ctx.getBillDate(2025, 10);
    expect(d.getMonth() + 1).toBe(11); // November
    expect(d.getDate()).toBe(10);
    expect(d.getFullYear()).toBe(2025);
  });

  it('rolls over to January of next year for December reading', () => {
    // Reading month: December (12) → Bill date: January 10, next year
    const d = ctx.getBillDate(2025, 12);
    expect(d.getMonth() + 1).toBe(1); // January
    expect(d.getDate()).toBe(10);
    expect(d.getFullYear()).toBe(2026);
  });

  it('handles November reading → December bill date in same year', () => {
    const d = ctx.getBillDate(2025, 11);
    expect(d.getMonth() + 1).toBe(12);
    expect(d.getFullYear()).toBe(2025);
    expect(d.getDate()).toBe(10);
  });

  it('handles January reading → February bill date', () => {
    const d = ctx.getBillDate(2025, 1);
    expect(d.getMonth() + 1).toBe(2);
    expect(d.getFullYear()).toBe(2025);
    expect(d.getDate()).toBe(10);
  });
});

// ── getDueDate ────────────────────────────────────────────────
describe('getDueDate', () => {
  it('returns 2nd of month, 2 months after reading month', () => {
    // Reading: October (10) → Due: December 2
    const d = ctx.getDueDate(2025, 10);
    expect(d.getMonth() + 1).toBe(12);
    expect(d.getDate()).toBe(2);
    expect(d.getFullYear()).toBe(2025);
  });

  it('rolls over year for November reading', () => {
    // Reading: November (11) → Due: January 2, next year
    const d = ctx.getDueDate(2025, 11);
    expect(d.getMonth() + 1).toBe(1);
    expect(d.getDate()).toBe(2);
    expect(d.getFullYear()).toBe(2026);
  });

  it('rolls over year for December reading', () => {
    // Reading: December (12) → Due: February 2, next year
    const d = ctx.getDueDate(2025, 12);
    expect(d.getMonth() + 1).toBe(2);
    expect(d.getDate()).toBe(2);
    expect(d.getFullYear()).toBe(2026);
  });

  it('handles mid-year reading', () => {
    // Reading: March (3) → Due: May 2
    const d = ctx.getDueDate(2025, 3);
    expect(d.getMonth() + 1).toBe(5);
    expect(d.getDate()).toBe(2);
    expect(d.getFullYear()).toBe(2025);
  });
});

// ── getMonthNum ───────────────────────────────────────────────
describe('getMonthNum', () => {
  it('returns 1 for January', () => expect(ctx.getMonthNum('January')).toBe(1));
  it('returns 6 for June',    () => expect(ctx.getMonthNum('June')).toBe(6));
  it('returns 12 for December', () => expect(ctx.getMonthNum('December')).toBe(12));
  it('returns 3 for March',   () => expect(ctx.getMonthNum('March')).toBe(3));

  it('returns 0 for an invalid month name (indexOf=-1, +1=0)', () => {
    // This is intentional GAS behavior: indexOf returns -1, +1 = 0
    expect(ctx.getMonthNum('InvalidMonth')).toBe(0);
  });

  it('returns 0 for empty string', () => {
    expect(ctx.getMonthNum('')).toBe(0);
  });

  it('returns 0 for undefined', () => {
    // 'undefined'.indexOf won't find it
    expect(ctx.getMonthNum(undefined)).toBe(0);
  });
});

// ── getMonthName ──────────────────────────────────────────────
describe('getMonthName', () => {
  it('returns "January" for index 1', () => expect(ctx.getMonthName(1)).toBe('January'));
  it('returns "June" for index 6',    () => expect(ctx.getMonthName(6)).toBe('June'));
  it('returns "December" for index 12', () => expect(ctx.getMonthName(12)).toBe('December'));

  it('returns empty string for out-of-range index 0', () => {
    expect(ctx.getMonthName(0)).toBe('');
  });

  it('returns empty string for out-of-range index 13', () => {
    expect(ctx.getMonthName(13)).toBe('');
  });

  it('returns empty string for negative index', () => {
    expect(ctx.getMonthName(-1)).toBe('');
  });
});

// ── buildBillNum ──────────────────────────────────────────────
describe('buildBillNum', () => {
  it('builds bill number for a standard unit', () => {
    // yr=2025, mo=3, phase=1, block=2, lot=5
    // → 2025 03 1 02 05 00
    expect(ctx.buildBillNum(2025, 3, 1, 2, 5)).toBe('2025031020500');
  });

  it('builds bill number for a double-digit block and lot', () => {
    // buildBillNum(yr, mo, phase, block, lot):
    // str(2025) + str(11).pad(2) + str(2) + str(10).pad(2) + str(24).pad(2) + '00'
    //   = '2025' + '11'   + '2'    + '10'    + '24'    + '00'
    //   = '2025112102400'  (phase is NOT padded)
    expect(ctx.buildBillNum(2025, 11, 2, 10, 24)).toBe('2025112102400');
  });

  it('uses only the first lot for combined lots', () => {
    // lot='1&2' → split on '&' → '1'
    // str(2025) + str(1).pad(2) + str(1) + str(6).pad(2) + str(1).pad(2) + '00'
    //   = '2025' + '01' + '1' + '06' + '01' + '00'  = '2025011060100'
    const billNum = ctx.buildBillNum(2025, 1, 1, 6, '1&2');
    expect(billNum).toBe('2025011060100');
    expect(billNum).not.toContain('&');
  });

  it('pads single-digit month with leading zero', () => {
    const billNum = ctx.buildBillNum(2025, 3, 1, 1, 1);
    // month 3 → '03'
    expect(billNum.substring(4, 6)).toBe('03');
  });
});

// ── getMLRow ──────────────────────────────────────────────────
describe('getMLRow', () => {
  const mlData = [
    ['Unit', 'Phase', 'Block', 'Lot', 'Last', 'First'], // header (row 0, skipped)
    ['P1B1L1', '1', '1', '1', 'Santos', 'Maria'],
    ['P1B1L2', '1', '1', '2', 'Cruz', 'Juan'],
    ['P1B2L3', '1', '2', '3', 'Reyes', 'Ana']
  ];

  it('finds and returns the row for a matching unit ID', () => {
    const row = ctx.getMLRow(mlData, 'P1B1L2');
    expect(row).toBeTruthy();
    expect(row[0]).toBe('P1B1L2');
    expect(row[4]).toBe('Cruz');
  });

  it('skips the first row (header)', () => {
    // The header has 'Unit' in col 0 — should NOT match 'Unit' as a unit ID
    // because the loop starts at i=1
    const row = ctx.getMLRow(mlData, 'Unit');
    expect(row).toBeNull();
  });

  it('returns null when unit ID is not found', () => {
    expect(ctx.getMLRow(mlData, 'P2B5L10')).toBeNull();
  });

  it('returns null for empty data', () => {
    expect(ctx.getMLRow([[]], 'P1B1L1')).toBeNull();
  });
});

// ── ownerName ─────────────────────────────────────────────────
describe('ownerName', () => {
  it('builds "LastName, FirstName" from columns 4 and 5', () => {
    const row = ['P1B1L1', '1', '1', '1', 'Santos', 'Maria'];
    expect(ctx.ownerName(row)).toBe('Santos, Maria');
  });

  it('returns empty string for null mlRow', () => {
    expect(ctx.ownerName(null)).toBe('');
  });

  it('handles only last name (no first name)', () => {
    const row = ['P1B1L1', '1', '1', '1', 'Santos', ''];
    // "Santos, " → trailing comma+space stripped → "Santos"
    expect(ctx.ownerName(row)).toBe('Santos');
  });

  it('handles only first name (no last name)', () => {
    const row = ['P1B1L1', '1', '1', '1', '', 'Maria'];
    // ", Maria" → leading comma+space stripped → "Maria"
    expect(ctx.ownerName(row)).toBe('Maria');
  });
});

// ── isValidUnitId_ ───────────────────────────────────────────
describe('isValidUnitId_', () => {
  it('accepts a standard unit ID P1B1L1', () => {
    expect(ctx.isValidUnitId_('P1B1L1')).toBe(true);
  });

  it('accepts a combined lot unit P1B6L1&2', () => {
    expect(ctx.isValidUnitId_('P1B6L1&2')).toBe(true);
  });

  it('accepts P2B3L24&25', () => {
    expect(ctx.isValidUnitId_('P2B3L24&25')).toBe(true);
  });

  it('rejects an empty string', () => {
    expect(ctx.isValidUnitId_('')).toBe(false);
  });

  it('rejects "GUARDHOUSE"', () => {
    expect(ctx.isValidUnitId_('GUARDHOUSE')).toBe(false);
  });

  it('rejects "ABC"', () => {
    expect(ctx.isValidUnitId_('ABC')).toBe(false);
  });

  it('rejects undefined', () => {
    expect(ctx.isValidUnitId_(undefined)).toBe(false);
  });

  it('is case-insensitive (lowercase p)', () => {
    // The regex uses /i flag
    expect(ctx.isValidUnitId_('p1b1l1')).toBe(true);
  });
});

// ── toNumber_ ────────────────────────────────────────────────
describe('toNumber_', () => {
  it('returns a number as-is', () => {
    expect(ctx.toNumber_(42)).toBe(42);
  });

  it('parses a numeric string', () => {
    expect(ctx.toNumber_('3.14')).toBe(3.14);
  });

  it('strips commas from strings like "1,234.56"', () => {
    expect(ctx.toNumber_('1,234.56')).toBe(1234.56);
  });

  it('returns 0 for empty string', () => {
    expect(ctx.toNumber_('')).toBe(0);
  });

  it('returns 0 for null', () => {
    expect(ctx.toNumber_(null)).toBe(0);
  });

  it('returns 0 for undefined', () => {
    expect(ctx.toNumber_(undefined)).toBe(0);
  });

  it('returns 0 for non-numeric string', () => {
    expect(ctx.toNumber_('hello')).toBe(0);
  });

  it('handles negative numbers', () => {
    expect(ctx.toNumber_(-50)).toBe(-50);
  });
});

// ── roundMoney_ ──────────────────────────────────────────────
describe('roundMoney_', () => {
  it('rounds to 2 decimal places', () => {
    expect(ctx.roundMoney_(123.456)).toBe(123.46);
  });

  it('handles zero', () => {
    expect(ctx.roundMoney_(0)).toBe(0);
  });

  it('handles negative values', () => {
    // JavaScript Math.round(-555.5) = -555 (rounds toward +infinity),
    // so roundMoney_(-5.555) = Math.round(-5.555 * 100) / 100 = Math.round(-555.5) / 100 = -555/100 = -5.55
    expect(ctx.roundMoney_(-5.555)).toBe(-5.55);
  });

  it('returns 0 for undefined', () => {
    expect(ctx.roundMoney_(undefined)).toBe(0);
  });

  it('returns 0 for null', () => {
    expect(ctx.roundMoney_(null)).toBe(0);
  });

  it('handles whole numbers correctly', () => {
    expect(ctx.roundMoney_(250)).toBe(250);
  });

  it('handles floating-point precision correctly', () => {
    // 0.1 + 0.2 = 0.30000000000000004 — roundMoney_ must fix it
    expect(ctx.roundMoney_(0.1 + 0.2)).toBe(0.3);
  });
});
