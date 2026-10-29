'use strict';
// Maintainer tool: records REAL Airtable API responses into test/fixtures/real-responses.json.
// Needs a token that can read the demo bases below:  AIRTABLE_TOKEN=... node test/record-fixtures.js
// The token is sent to Airtable only; it is never written to the fixture file.
const fs = require('fs');
const path = require('path');
const { load, curlFetch } = require('./harness');

const token = process.env.AIRTABLE_TOKEN;
if (!token) { console.error('Set AIRTABLE_TOKEN'); process.exit(1); }

const recorded = [];
const recordingFetch = (url, options) => {
  const r = curlFetch(url, options);
  recorded.push({ url, status: r.code, body: r.body });
  return r;
};

const scenarios = [
  // Linked-record fields (arrays), several pages: PAGE_SIZE 3 forces the offset cursor to be used.
  { BASE_ID: 'appgqeNCk6gr37gw8', TABLE: 'Work log', VIEW: '', PAGE_SIZE: 3 },
  // A view plus an explicit field list.
  { BASE_ID: 'appgqeNCk6gr37gw8', TABLE: 'Work log', VIEW: 'This week', FIELDS: ['Task', 'Hours', 'Status'] },
  // Different base: text, select, number, currency, date fields. Single page.
  { BASE_ID: 'appWNdNoD1nqbuqbj', TABLE: 'Items', VIEW: '' },
  // Real error responses: bad table, then bad token.
  { BASE_ID: 'appgqeNCk6gr37gw8', TABLE: 'No such table', VIEW: '', expectError: true },
  { BASE_ID: 'appgqeNCk6gr37gw8', TABLE: 'Work log', VIEW: '', expectError: true, token: 'invalid-token' },
];

for (const s of scenarios) {
  const { expectError, token: tok, ...overrides } = s;
  const h = load({ fetchImpl: (u, o) => recordingFetch(u, o), overrides, token: tok || token });
  try {
    const out = h.ctx.syncAirtable();
    console.log('recorded', overrides.TABLE, overrides.VIEW || '(all)', '->', JSON.stringify(out));
  } catch (e) {
    if (!expectError) throw e;
    console.log('recorded error', overrides.TABLE, '->', e.message);
  }
}
const dir = path.join(__dirname, 'fixtures');
fs.mkdirSync(dir, { recursive: true });
fs.writeFileSync(path.join(dir, 'real-responses.json'), JSON.stringify(recorded, null, 2) + '\n');
console.log('wrote', recorded.length, 'responses');
