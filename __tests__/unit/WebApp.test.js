// ============================================================
// __tests__/unit/WebApp.test.js
// Tests for pure helper functions in WebApp.gs:
//   webapp_isCommonAccount_
//   webapp_isRegularUnit_
//   webapp_getGroup_
//   webapp_buildCommonBillNum_
//   webapp_emptyTotals_
//   webapp_addToTotals_
//   webapp_roundTotals_
//   webapp_periodKey_
//   webapp_periodLabel_
//
// None of these functions access sheets.
// ============================================================

'use strict';

const { createGASContext } = require('../helpers/gas-loader');

let ctx;
beforeAll(() => {
  ctx = createGASContext({});
});

// ═════════════════════════════════════════════════════════════
// webapp_isCommonAccount_
// ═════════════════════════════════════════════════════════════
describe('webapp_isCommonAccount_', () => {
  it('returns true for "GUARDHOUSE"', () => {
    expect(ctx.webapp_isCommonAccount_('GUARDHOUSE')).toBe(true);
  });

  it('returns true for "CLUBHOUSE"', () => {
    expect(ctx.webapp_isCommonAccount_('CLUBHOUSE')).toBe(true);
  });

  it('returns true for "CHAPEL"', () => {
    expect(ctx.webapp_isCommonAccount_('CHAPEL')).toBe(true);
  });

  it('is case-insensitive', () => {
    expect(ctx.webapp_isCommonAccount_('guardhouse')).toBe(true);
  });

  it('strips whitespace before matching', () => {
    expect(ctx.webapp_isCommonAccount_('GUARD HOUSE')).toBe(true);
  });

  it('returns false for a regular unit', () => {
    expect(ctx.webapp_isCommonAccount_('P1B1L1')).toBe(false);
  });

  it('returns false for empty string', () => {
    expect(ctx.webapp_isCommonAccount_('')).toBe(false);
  });

  it('returns false for null', () => {
    expect(ctx.webapp_isCommonAccount_(null)).toBe(false);
  });
});

// ═════════════════════════════════════════════════════════════
// webapp_isRegularUnit_
// ═════════════════════════════════════════════════════════════
describe('webapp_isRegularUnit_', () => {
  it('returns true for standard unit P1B1L1', () => {
    expect(ctx.webapp_isRegularUnit_('P1B1L1')).toBe(true);
  });

  it('returns true for combined lot P1B6L1&2', () => {
    expect(ctx.webapp_isRegularUnit_('P1B6L1&2')).toBe(true);
  });

  it('returns true for phase 2 double-digit unit', () => {
    expect(ctx.webapp_isRegularUnit_('P2B10L24')).toBe(true);
  });

  it('returns false for "GUARDHOUSE"', () => {
    expect(ctx.webapp_isRegularUnit_('GUARDHOUSE')).toBe(false);
  });

  it('returns false for empty string', () => {
    expect(ctx.webapp_isRegularUnit_('')).toBe(false);
  });

  it('returns false for null', () => {
    expect(ctx.webapp_isRegularUnit_(null)).toBe(false);
  });

  it('returns false for a random string "ABC"', () => {
    expect(ctx.webapp_isRegularUnit_('ABC')).toBe(false);
  });
});

// ═════════════════════════════════════════════════════════════
// webapp_getGroup_
// ═════════════════════════════════════════════════════════════
describe('webapp_getGroup_', () => {
  it('P1B2L3 → Phase 1, Block 2, not common', () => {
    const g = ctx.webapp_getGroup_('P1B2L3');
    expect(g).toEqual({ phase: 'Phase 1', block: 'Block 2', common: false });
  });

  it('P2B10L5 → Phase 2, Block 10, not common', () => {
    const g = ctx.webapp_getGroup_('P2B10L5');
    expect(g).toEqual({ phase: 'Phase 2', block: 'Block 10', common: false });
  });

  it('"GUARDHOUSE" → Common, Common Accounts, common=true', () => {
    const g = ctx.webapp_getGroup_('GUARDHOUSE');
    expect(g).toEqual({ phase: 'Common', block: 'Common Accounts', common: true });
  });

  it('"CLUBHOUSE" → common=true', () => {
    const g = ctx.webapp_getGroup_('CLUBHOUSE');
    expect(g.common).toBe(true);
  });

  it('unknown string → Other, Other, common=false', () => {
    const g = ctx.webapp_getGroup_('SOME_RANDOM_THING');
    expect(g).toEqual({ phase: 'Other', block: 'Other', common: false });
  });

  it('empty string → Other', () => {
    const g = ctx.webapp_getGroup_('');
    expect(g.phase).toBe('Other');
  });

  it('lowercase uid is handled (trimmed and uppercased internally)', () => {
    const g = ctx.webapp_getGroup_('p1b1l1');
    expect(g.phase).toBe('Phase 1');
    expect(g.block).toBe('Block 1');
  });
});

// ═════════════════════════════════════════════════════════════
// webapp_buildCommonBillNum_
// ═════════════════════════════════════════════════════════════
describe('webapp_buildCommonBillNum_', () => {
  it('builds "WB-YYYY-MM-UNITID" format', () => {
    expect(ctx.webapp_buildCommonBillNum_(2025, 3, 'GUARDHOUSE'))
      .toBe('WB-2025-03-GUARDHOUSE');
  });

  it('pads single-digit month with leading zero', () => {
    const result = ctx.webapp_buildCommonBillNum_(2025, 1, 'CHAPEL');
    expect(result).toBe('WB-2025-01-CHAPEL');
  });

  it('converts unitId to uppercase', () => {
    const result = ctx.webapp_buildCommonBillNum_(2025, 3, 'guardhouse');
    expect(result).toBe('WB-2025-03-GUARDHOUSE');
  });

  it('replaces spaces in unitId with dashes', () => {
    const result = ctx.webapp_buildCommonBillNum_(2025, 3, 'GUARD HOUSE');
    expect(result).toBe('WB-2025-03-GUARD-HOUSE');
  });
});

// ═════════════════════════════════════════════════════════════
// webapp_emptyTotals_
// ═════════════════════════════════════════════════════════════
describe('webapp_emptyTotals_', () => {
  it('returns an object with exactly 9 numeric keys all set to 0', () => {
    const t = ctx.webapp_emptyTotals_();
    const keys = Object.keys(t);
    expect(keys.length).toBe(9);
    keys.forEach(k => {
      expect(typeof t[k]).toBe('number');
      expect(t[k]).toBe(0);
    });
  });

  it('contains the expected keys', () => {
    const t = ctx.webapp_emptyTotals_();
    expect(t).toHaveProperty('payments', 0);
    expect(t).toHaveProperty('receivables', 0);
    expect(t).toHaveProperty('overpayments', 0);
    expect(t).toHaveProperty('waterPayments', 0);
    expect(t).toHaveProperty('waterReceivables', 0);
    expect(t).toHaveProperty('waterOverpayments', 0);
    expect(t).toHaveProperty('duesPayments', 0);
    expect(t).toHaveProperty('duesReceivables', 0);
    expect(t).toHaveProperty('duesOverpayments', 0);
  });
});

// ═════════════════════════════════════════════════════════════
// webapp_addToTotals_
// ═════════════════════════════════════════════════════════════
describe('webapp_addToTotals_', () => {
  it('adds numeric values from source to target', () => {
    const target = ctx.webapp_emptyTotals_();
    const source = { payments: 100, receivables: 500, waterPayments: 100 };
    ctx.webapp_addToTotals_(target, source);
    expect(target.payments).toBe(100);
    expect(target.receivables).toBe(500);
    expect(target.waterPayments).toBe(100);
    // Other keys unchanged
    expect(target.overpayments).toBe(0);
  });

  it('accumulates across multiple calls', () => {
    const target = ctx.webapp_emptyTotals_();
    ctx.webapp_addToTotals_(target, { payments: 100 });
    ctx.webapp_addToTotals_(target, { payments: 200 });
    expect(target.payments).toBe(300);
  });

  it('handles missing keys in source gracefully (target key unchanged)', () => {
    const target = ctx.webapp_emptyTotals_();
    // source has no 'payments' key
    ctx.webapp_addToTotals_(target, { someOtherKey: 999 });
    expect(target.payments).toBe(0);
  });

  it('treats non-numeric source values as 0', () => {
    const target = ctx.webapp_emptyTotals_();
    ctx.webapp_addToTotals_(target, { payments: 'hello' });
    // toNum('hello') = 0
    expect(target.payments).toBe(0);
  });
});

// ═════════════════════════════════════════════════════════════
// webapp_roundTotals_
// ═════════════════════════════════════════════════════════════
describe('webapp_roundTotals_', () => {
  it('rounds all numeric values to 2 decimal places', () => {
    const obj = { payments: 100.555, receivables: 200.004 };
    const result = ctx.webapp_roundTotals_(obj);
    expect(result.payments).toBe(100.56);
    expect(result.receivables).toBe(200);
  });

  it('ignores non-numeric values', () => {
    const obj = { payments: 100, label: 'Phase 1' };
    ctx.webapp_roundTotals_(obj);
    expect(obj.label).toBe('Phase 1'); // unchanged
  });

  it('returns the mutated object (for chaining)', () => {
    const obj = { payments: 100.555 };
    const ret = ctx.webapp_roundTotals_(obj);
    expect(ret).toBe(obj); // same reference
  });

  it('handles zero values', () => {
    const obj = { payments: 0, receivables: 0 };
    ctx.webapp_roundTotals_(obj);
    expect(obj.payments).toBe(0);
  });

  it('handles negative values', () => {
    // Math.round(-100.555 * 100) / 100:
    // -100.555 * 100 = -10055.5 → Math.round(-10055.5) = -10055 (rounds toward +infinity)
    // -10055 / 100 = -100.55
    const obj = { payments: -100.555 };
    ctx.webapp_roundTotals_(obj);
    expect(obj.payments).toBe(-100.55);
  });
});

// ═════════════════════════════════════════════════════════════
// webapp_periodKey_
// ═════════════════════════════════════════════════════════════
describe('webapp_periodKey_', () => {
  it('2025, "March" → "2025-03"', () => {
    expect(ctx.webapp_periodKey_(2025, 'March')).toBe('2025-03');
  });

  it('2025, "January" → "2025-01"', () => {
    expect(ctx.webapp_periodKey_(2025, 'January')).toBe('2025-01');
  });

  it('2025, "December" → "2025-12"', () => {
    expect(ctx.webapp_periodKey_(2025, 'December')).toBe('2025-12');
  });

  it('invalid year returns month name or "Unspecified"', () => {
    // When year is invalid (0 or NaN), the function returns m || 'Unspecified'
    const result = ctx.webapp_periodKey_(0, 'March');
    // !y is true → returns 'March' (the month string)
    expect(result).toBe('March');
  });

  it('invalid month name returns "Unspecified" when month is empty', () => {
    const result = ctx.webapp_periodKey_(2025, '');
    // m = '' → !n is true, returns '' || 'Unspecified' = 'Unspecified'
    expect(result).toBe('Unspecified');
  });

  it('zero year and empty month → "Unspecified"', () => {
    const result = ctx.webapp_periodKey_(0, '');
    expect(result).toBe('Unspecified');
  });
});

// ═════════════════════════════════════════════════════════════
// webapp_periodLabel_
// ═════════════════════════════════════════════════════════════
describe('webapp_periodLabel_', () => {
  it('2025, "March" → "Mar 2025"', () => {
    expect(ctx.webapp_periodLabel_(2025, 'March')).toBe('Mar 2025');
  });

  it('2025, "December" → "Dec 2025"', () => {
    expect(ctx.webapp_periodLabel_(2025, 'December')).toBe('Dec 2025');
  });

  it('2025, "January" → "Jan 2025"', () => {
    expect(ctx.webapp_periodLabel_(2025, 'January')).toBe('Jan 2025');
  });

  it('invalid year → returns month or "Unspecified"', () => {
    // When year is 0 → !y is true → returns m || 'Unspecified'
    // m = 'March' → returns 'March'
    const result = ctx.webapp_periodLabel_(0, 'March');
    expect(result).toBe('March');
  });

  it('empty month and valid year → "Unspecified"', () => {
    const result = ctx.webapp_periodLabel_(2025, '');
    expect(result).toBe('Unspecified');
  });

  it('both empty → "Unspecified"', () => {
    const result = ctx.webapp_periodLabel_(0, '');
    expect(result).toBe('Unspecified');
  });

  it('uses only first 3 chars of month name', () => {
    // "September" → "Sep"
    expect(ctx.webapp_periodLabel_(2025, 'September')).toBe('Sep 2025');
  });
});
