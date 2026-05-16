// ============================================================
// __tests__/helpers/gas-loader.js
//
// Utilities for loading GAS source files into an isolated
// vm.runInNewContext sandbox, and for building mock Sheet /
// Spreadsheet / Range objects that faithfully simulate the
// Google Sheets API surface used by the billing system.
// ============================================================

'use strict';

const vm   = require('vm');
const fs   = require('fs');
const path = require('path');

// Absolute path to the project root where the .gs files live
const ROOT = path.resolve(__dirname, '..', '..');

// Ordered list of .gs files to load (order matters: Code.gs must be first
// so constants like WL, DL, MONTHS, ASSOC_DUES etc. are defined before
// the other files reference them)
const GS_FILES = [
  'Code.gs',
  'BillingEngine.gs',
  'PaymentProcessor.gs',
  'LedgerManager.gs',
  'SummaryGenerator.gs',
  'Setup.gs',
  'WebApp.gs'
];

// ── Range mock ────────────────────────────────────────────────
/**
 * Creates a range mock that operates on a slice of the sheet's 2-D data array.
 *
 * @param {Array<Array<*>>}  sheetData  - Reference to the sheet's 2-D data array (mutated in place)
 * @param {number}           startRow   - 1-based row index
 * @param {number}           startCol   - 1-based column index
 * @param {number}           numRows    - number of rows to cover (default 1)
 * @param {number}           numCols    - number of columns to cover (default 1)
 * @returns {object}         range mock
 */
function createMockRange(sheetData, startRow, startCol, numRows, numCols) {
  numRows = numRows || 1;
  numCols = numCols || 1;

  const range = {
    // ── Read ────────────────────────────────────────────────
    getValue() {
      const rowIdx = startRow - 1;
      const colIdx = startCol - 1;
      if (rowIdx < 0 || rowIdx >= sheetData.length) return '';
      const row = sheetData[rowIdx];
      if (!row || colIdx < 0 || colIdx >= row.length) return '';
      return row[colIdx];
    },

    getValues() {
      const result = [];
      for (let r = 0; r < numRows; r++) {
        const rowIdx = startRow - 1 + r;
        const dataRow = sheetData[rowIdx] || [];
        const cols = [];
        for (let c = 0; c < numCols; c++) {
          const colIdx = startCol - 1 + c;
          cols.push(dataRow[colIdx] !== undefined ? dataRow[colIdx] : '');
        }
        result.push(cols);
      }
      return result;
    },

    // ── Write ───────────────────────────────────────────────
    setValue(value) {
      const rowIdx = startRow - 1;
      const colIdx = startCol - 1;
      // Expand sheetData if needed
      while (sheetData.length <= rowIdx) sheetData.push([]);
      while (sheetData[rowIdx].length <= colIdx) sheetData[rowIdx].push('');
      sheetData[rowIdx][colIdx] = value;
      return this;
    },

    setValues(values) {
      for (let r = 0; r < values.length; r++) {
        const rowIdx = startRow - 1 + r;
        while (sheetData.length <= rowIdx) sheetData.push([]);
        for (let c = 0; c < values[r].length; c++) {
          const colIdx = startCol - 1 + c;
          while (sheetData[rowIdx].length <= colIdx) sheetData[rowIdx].push('');
          sheetData[rowIdx][colIdx] = values[r][c];
        }
      }
      return this;
    },

    // ── Chainable no-op formatting methods ──────────────────
    // All formatting methods return `this` so fluent chains work.
    setBackground:         function() { return this; },
    setBackgrounds:        function() { return this; },
    setFontColor:          function() { return this; },
    setFontWeight:         function() { return this; },
    setFontSize:           function() { return this; },
    setFontStyle:          function() { return this; },
    setHorizontalAlignment:function() { return this; },
    setNumberFormat:       function() { return this; },
    setNumberFormats:      function() { return this; },
    setBorder:             function() { return this; },
    setWrap:               function() { return this; },
    setWrapStrategy:       function() { return this; },
    clearContent:          function() { return this; },
    clearFormat:           function() {
      // clear content in the range too when clearContent is called
      for (let r = 0; r < numRows; r++) {
        const rowIdx = startRow - 1 + r;
        if (rowIdx < sheetData.length) {
          for (let c = 0; c < numCols; c++) {
            const colIdx = startCol - 1 + c;
            if (sheetData[rowIdx] && colIdx < sheetData[rowIdx].length) {
              sheetData[rowIdx][colIdx] = '';
            }
          }
        }
      }
      return this;
    },
    merge:                 function() { return this; },
    setDataValidation:     function() { return this; },
    activate:              function() { return this; },
    autoResizeColumns:     function() { return this; },
    getContent:            function() { return ''; },
    getSheet:              function() { return null; } // overridden per sheet
  };

  return range;
}

// ── Sheet mock ────────────────────────────────────────────────
/**
 * Creates a sheet mock from a 2-D data array.
 * Row 1 = header (index 0 in array).
 *
 * @param {string}           name      - sheet name
 * @param {Array<Array<*>>}  initData  - initial 2-D data (may be empty [])
 * @returns {object}         sheet mock
 */
function createMockSheet(name, initData) {
  // Deep-copy so test mutations don't bleed between tests
  const data = (initData || []).map(row => row.slice());

  // Parse "A1" style cell address → { row, col } (both 1-based)
  function parseCellAddress(addr) {
    const m = String(addr).match(/^([A-Za-z]+)(\d+)$/);
    if (!m) return null;
    const colStr = m[1].toUpperCase();
    let col = 0;
    for (let i = 0; i < colStr.length; i++) {
      col = col * 26 + (colStr.charCodeAt(i) - 64);
    }
    return { row: parseInt(m[2], 10), col };
  }

  const sheet = {
    _name: name,
    _data: data,   // exposed for direct test assertions

    getName() { return name; },

    // ── getRange overloads ──────────────────────────────────
    // Supports:
    //   getRange(row, col)
    //   getRange(row, col, numRows)
    //   getRange(row, col, numRows, numCols)
    //   getRange('A1')   ← cell address string
    getRange(rowOrAddr, col, numRows, numCols) {
      if (typeof rowOrAddr === 'string') {
        const parsed = parseCellAddress(rowOrAddr);
        if (!parsed) throw new Error(`Cannot parse cell address: ${rowOrAddr}`);
        const r = createMockRange(data, parsed.row, parsed.col, 1, 1);
        r.getSheet = () => sheet;
        return r;
      }
      const r = createMockRange(
        data,
        rowOrAddr,
        col,
        numRows || 1,
        numCols || (numRows ? undefined : 1)
      );
      r.getSheet = () => sheet;
      return r;
    },

    // ── Row/column information ──────────────────────────────
    getLastRow() {
      // Return 1-based index of last row that has any non-empty cell
      let last = 0;
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (row && row.some(cell => cell !== '' && cell !== null && cell !== undefined)) {
          last = i + 1;
        }
      }
      return last;
    },

    getLastColumn() {
      let max = 0;
      for (const row of data) {
        if (row && row.length > max) max = row.length;
      }
      return max;
    },

    getMaxRows()    { return Math.max(data.length, 1000); },
    getMaxColumns() { return 26; },

    // ── Full-sheet range ────────────────────────────────────
    getDataRange() {
      const lr = sheet.getLastRow();
      const lc = sheet.getLastColumn();
      if (lr === 0 || lc === 0) return createMockRange(data, 1, 1, 1, 1);
      const r = createMockRange(data, 1, 1, lr, lc);
      r.getSheet = () => sheet;
      return r;
    },

    // ── Row mutation ────────────────────────────────────────
    appendRow(row) {
      data.push(row.slice());
      return this;
    },

    deleteRow(n) {
      // n is 1-based
      if (n >= 1 && n <= data.length) {
        data.splice(n - 1, 1);
      }
      return this;
    },

    insertRowAfter(rowNum) {
      data.splice(rowNum, 0, []);
      return this;
    },

    // ── Formatting (no-ops that return this) ────────────────
    clearContents()      { data.length = 0; return this; },
    clearFormats()       { return this; },
    setFrozenRows()      { return this; },
    setFrozenColumns()   { return this; },
    autoResizeColumns()  { return this; },
    setColumnWidth()     { return this; },
    hideSheet()          { return this; },
    showSheet()          { return this; },
    setTabColor()        { return this; },
    setName(n)           { /* read-only in tests */ return this; },
    activate()           { return this; },

    // ── Data validation (no-op) ─────────────────────────────
    setDataValidation()  { return this; }
  };

  return sheet;
}

// ── Spreadsheet mock ──────────────────────────────────────────
/**
 * Creates a spreadsheet mock from an object mapping sheet names to 2-D data arrays.
 *
 * @param {Object.<string, Array<Array<*>>>} sheetsMap
 * @returns {object} spreadsheet mock
 */
function createMockSpreadsheet(sheetsMap) {
  const sheets = {};
  Object.keys(sheetsMap || {}).forEach(name => {
    sheets[name] = createMockSheet(name, sheetsMap[name]);
  });

  const ss = {
    _sheets: sheets,

    getSheetByName(name) {
      return sheets[name] || null;
    },

    insertSheet(name) {
      sheets[name] = createMockSheet(name, []);
      return sheets[name];
    },

    getSheets() {
      return Object.values(sheets);
    },

    getUrl()  { return 'https://docs.google.com/spreadsheets/d/MOCK_ID'; },
    getName() { return 'AHNHAI Billing [TEST]'; },

    toast:         jest.fn(),
    setActiveSheet:jest.fn()
  };

  return ss;
}

// ── GAS Context ───────────────────────────────────────────────
/**
 * Creates a full GAS execution context by:
 *  1. Constructing a mock SpreadsheetApp pointing at sheetsMap
 *  2. Building all GAS global mocks
 *  3. Loading each .gs file in order via vm.runInNewContext
 *
 * The returned context object contains all GAS functions as properties,
 * so tests can call e.g. ctx.fmt2(…), ctx.recalcWaterBalances(), etc.
 *
 * @param {Object.<string, Array<Array<*>>>} sheetsMap   - initial sheet data
 * @param {Object}                           extraGlobals - additional globals to inject
 * @returns {object}  context (all GAS globals + all GAS functions)
 */
function createGASContext(sheetsMap, extraGlobals) {
  const mockSS = createMockSpreadsheet(sheetsMap || {});

  // Build a fresh UI mock per context so jest.fn() calls are isolated
  const mockUi = {
    alert:          jest.fn(),
    createMenu:     jest.fn().mockReturnThis(),
    addItem:        jest.fn().mockReturnThis(),
    addSeparator:   jest.fn().mockReturnThis(),
    addToUi:        jest.fn(),
    ButtonSet:      { YES_NO: 'YES_NO' },
    Button:         { YES: 'YES' }
  };

  const mockSpreadsheetApp = {
    getActiveSpreadsheet: jest.fn(() => mockSS),
    getUi:                jest.fn(() => mockUi),
    newDataValidation:    jest.fn(() => ({
      requireValueInList: jest.fn().mockReturnThis(),
      setAllowInvalid:    jest.fn().mockReturnThis(),
      build:              jest.fn(() => ({}))
    }))
  };

  // ── Base context ──────────────────────────────────────────
  const ctx = {
    // GAS globals
    SpreadsheetApp: mockSpreadsheetApp,
    Logger:         { log: jest.fn() },
    HtmlService: {
      createTemplateFromFile: jest.fn(() => ({
        evaluate: jest.fn(() => ({
          setTitle:             jest.fn().mockReturnThis(),
          setXFrameOptionsMode: jest.fn().mockReturnThis()
        }))
      })),
      createHtmlOutputFromFile: jest.fn(() => ({
        getContent: jest.fn(() => '')
      })),
      XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' }
    },
    DriveApp: {
      getFileById: jest.fn(() => ({
        getName: jest.fn(() => 'Mock File'),
        getId:   jest.fn(() => 'mock-file-id')
      }))
    },
    // Node built-ins that GAS scripts might accidentally reference
    console,
    parseInt,
    parseFloat,
    isNaN,
    Math,
    String,
    Number,
    Array,
    Object,
    Date,
    RegExp,
    JSON,

    // Expose mock objects so tests can reach into them
    _mockSS: mockSS,
    _mockUi: mockUi,
    _mockSpreadsheetApp: mockSpreadsheetApp,

    // Spread any extra globals the caller wants
    ...(extraGlobals || {})
  };

  // ── Load each .gs file into the context ───────────────────
  for (const file of GS_FILES) {
    const filePath = path.join(ROOT, file);
    if (!fs.existsSync(filePath)) {
      // Skip missing optional files (e.g. PrintGenerator.gs not listed in GS_FILES)
      continue;
    }
    const code = fs.readFileSync(filePath, 'utf8');
    try {
      vm.runInNewContext(code, ctx, { filename: file });
    } catch (err) {
      // Re-throw with the file name for easier debugging
      throw new Error(`Error loading ${file}: ${err.message}\n${err.stack}`);
    }
  }

  return ctx;
}

// ── Exports ───────────────────────────────────────────────────
module.exports = {
  createMockRange,
  createMockSheet,
  createMockSpreadsheet,
  createGASContext
};
