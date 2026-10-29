'use strict';
// Loads Code.gs into a Node vm context with stand-ins for the Apps Script services.
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { spawnSync } = require('child_process');

const CODE = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');

function makeSheet(name) {
  const sheet = {
    name, cells: new Map(), notes: new Map(), frozenRows: 0, clears: 0, writes: [],
    clearContents() { this.cells.clear(); this.clears++; },
    setFrozenRows(n) { this.frozenRows = n; },
    getRange(row, col, numRows = 1, numCols = 1) {
      return {
        setValues: (values) => {
          if (values.length !== numRows) throw new Error('row count mismatch: ' + values.length + ' vs ' + numRows);
          values.forEach((vals, r) => {
            if (vals.length !== numCols) throw new Error('column count mismatch');
            vals.forEach((v, c) => sheet.cells.set((row + r) + ',' + (col + c), v));
          });
          sheet.writes.push({ row, col, numRows, numCols });
        },
        setNote: (text) => sheet.notes.set(row + ',' + col, text),
        clearNote: () => sheet.notes.delete(row + ',' + col),
      };
    },
    // Returns the grid as an array of rows (trailing empty rows trimmed).
    grid() {
      let maxR = 0, maxC = 0;
      for (const k of sheet.cells.keys()) { const [r, c] = k.split(',').map(Number); maxR = Math.max(maxR, r); maxC = Math.max(maxC, c); }
      const out = [];
      for (let r = 1; r <= maxR; r++) { const line = []; for (let c = 1; c <= maxC; c++) line.push(sheet.cells.has(r + ',' + c) ? sheet.cells.get(r + ',' + c) : ''); out.push(line); }
      return out;
    },
  };
  return sheet;
}

/**
 * fetchImpl(url, options) -> { code, body, headers }
 * overrides: values assigned to the config variables of Code.gs, e.g. { BASE_ID: 'app...' }
 */
function load({ fetchImpl, overrides = {}, token = 'test-token' } = {}) {
  const sheets = {};
  const calls = { fetch: [], sleep: [], triggers: [], deleted: [] };
  const existingTriggers = [];
  const spreadsheet = {
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => (sheets[n] = makeSheet(n)),
  };
  const props = token === null ? {} : { AIRTABLE_TOKEN: token };
  const sandbox = {
    console,
    Logger: { log: () => {} },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (k in props ? props[k] : null) }) },
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet, openById: () => spreadsheet, flush: () => {} },
    Utilities: { sleep: (ms) => calls.sleep.push(ms) },
    UrlFetchApp: {
      fetch(url, options) {
        calls.fetch.push({ url, options });
        const r = fetchImpl(url, options);
        return { getResponseCode: () => r.code, getContentText: () => r.body, getHeaders: () => r.headers || {} };
      },
    },
    ScriptApp: {
      getProjectTriggers: () => existingTriggers.slice(),
      deleteTrigger: (t) => { calls.deleted.push(t); existingTriggers.splice(existingTriggers.indexOf(t), 1); },
      newTrigger(fn) {
        const spec = { fn, steps: [] };
        const b = {
          timeBased: () => b,
          everyHours: (n) => { spec.steps.push(['everyHours', n]); return b; },
          everyDays: (n) => { spec.steps.push(['everyDays', n]); return b; },
          atHour: (n) => { spec.steps.push(['atHour', n]); return b; },
          create: () => { calls.triggers.push(spec); existingTriggers.push({ getHandlerFunction: () => fn, spec }); return spec; },
        };
        return b;
      },
    },
  };
  const ctx = vm.createContext(sandbox);
  vm.runInContext(CODE, ctx, { filename: 'Code.gs' });
  for (const [k, v] of Object.entries(overrides)) ctx[k] = v;
  return { ctx, sheets, calls, existingTriggers };
}

// Real network via curl (only used to record fixtures and by the optional live test).
function curlFetch(url, options) {
  const args = ['-sS', '-m', '60', '-w', '\n%{http_code}', '-H', 'Authorization: ' + options.headers.Authorization, url];
  const r = spawnSync('curl', args, { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
  if (r.status !== 0) throw new Error('curl failed: ' + r.stderr);
  const i = r.stdout.lastIndexOf('\n');
  return { code: Number(r.stdout.slice(i + 1)), body: r.stdout.slice(0, i) };
}

module.exports = { load, curlFetch };
