// ============================================================
// jest.setup.js
// Sets up GAS global mocks on the Node.js global object.
// These mocks are available in every test file automatically.
// Do NOT load .gs files here — that is done per-test in gas-loader.js.
// ============================================================

// ── Logger mock ──────────────────────────────────────────────
global.Logger = {
  log: jest.fn()
};

// ── Minimal SpreadsheetApp mock ──────────────────────────────
// Concrete spreadsheet instances are created in gas-loader.js.
// This stub prevents "SpreadsheetApp is not defined" errors in
// any code that runs before createGASContext is called.
global.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn(() => ({
    getSheetByName: jest.fn(() => null),
    insertSheet: jest.fn(),
    getUrl: jest.fn(() => 'https://docs.google.com/spreadsheets/d/test'),
    getName: jest.fn(() => 'Test Spreadsheet'),
    toast: jest.fn()
  })),
  getUi: jest.fn(() => ({
    alert: jest.fn(),
    createMenu: jest.fn().mockReturnThis(),
    addItem: jest.fn().mockReturnThis(),
    addSeparator: jest.fn().mockReturnThis(),
    addToUi: jest.fn(),
    ButtonSet: { YES_NO: 'YES_NO' },
    Button: { YES: 'YES' }
  })),
  newDataValidation: jest.fn(() => ({
    requireValueInList: jest.fn().mockReturnThis(),
    setAllowInvalid: jest.fn().mockReturnThis(),
    build: jest.fn(() => ({}))
  }))
};

// ── HtmlService mock ─────────────────────────────────────────
global.HtmlService = {
  createTemplateFromFile: jest.fn(() => ({
    evaluate: jest.fn(() => ({
      setTitle: jest.fn().mockReturnThis(),
      setXFrameOptionsMode: jest.fn().mockReturnThis()
    }))
  })),
  createHtmlOutputFromFile: jest.fn(() => ({
    getContent: jest.fn(() => '')
  })),
  XFrameOptionsMode: {
    ALLOWALL: 'ALLOWALL'
  }
};

// ── DriveApp mock ─────────────────────────────────────────────
global.DriveApp = {
  getFileById: jest.fn(() => ({
    getName: jest.fn(() => 'Mock File'),
    getId: jest.fn(() => 'mock-file-id')
  }))
};

// ── console mock (GAS uses Logger, but some utility code may use console) ──
// We leave console intact so test failures print meaningful messages.

// ── Utilities ─────────────────────────────────────────────────
// Clear all mock call counts before each test by default
beforeEach(() => {
  jest.clearAllMocks();
});
