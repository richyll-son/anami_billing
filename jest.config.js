/** @type {import('jest').Config} */
module.exports = {
  // Find all test files under __tests__/
  testMatch: [
    '<rootDir>/__tests__/**/*.test.js'
  ],

  // Global setup run after the test framework is installed
  // setupFilesAfterEachTestFramework: not needed — GAS globals are set per-test in gas-loader.js

  // No transform needed — the helper uses vm.runInNewContext,
  // and the test files themselves are plain CommonJS.
  // babel-jest handles ES-module syntax in test files only.
  transform: {
    '^.+\\.js$': 'babel-jest'
  },

  // Exclude the .gs source files from Jest's own module resolution
  // (they are loaded via vm.runInNewContext in gas-loader.js)
  testPathIgnorePatterns: ['/node_modules/'],

  // Coverage settings
  collectCoverageFrom: [
    '*.gs'
  ],

  coverageThreshold: {
    global: {
      lines: 80,
      functions: 80,
      branches: 60,
      statements: 80
    }
  },

  coverageReporters: ['text', 'lcov', 'html'],

  // Increase timeout for integration tests that do more work
  testTimeout: 15000,

  // Verbose output
  verbose: true
};
