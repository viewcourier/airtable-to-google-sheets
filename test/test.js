'use strict';
// Run: node test/test.js        (Node 18+, no dependencies)
// Loads Code.gs in a vm with stand-ins for UrlFetchApp, SpreadsheetApp, PropertiesService, ScriptApp.
// Airtable responses are REAL ones recorded by test/record-fixtures.js (test/fixtures/real-responses.json).
// Set AIRTABLE_TOKEN to also run one live check against the Airtable API.
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { load, curlFetch } = require('./harness');

const fixtures = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'real-responses.json'), 'utf8'));
const byUrl = new Map(fixtures.map((f) => [f.url, f]));
const replay = (url) => {
  const f = byUrl.get(url);
  if (!f) throw new Error('no recorded response for ' + url);
  return { code: f.status, body: f.body };
};
const json = (x) => JSON.stringify(x);
// Plain-object copy so deepStrictEqual does not trip over the vm realm.
const plain = (x) => JSON.parse(json(x));

const WORK = { BASE_ID: 'appgqeNCk6gr37gw8', TABLE: 'Work log', VIEW: '', PAGE_SIZE: 3 };
const workPages = fixtures.filter((f) => f.url.includes('/Work%20log?pageSize=3')).map((f) => JSON.parse(f.body));

let passed = 0, failed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('ok   - ' + name); }
  catch (e) { failed++; console.log('FAIL - ' + name + '\n       ' + (e && e.stack || e)); }
}

// ---- flattening ------------------------------------------------------------------------------
test('flattenValue_: scalars and empties', () => {
  const { ctx } = load();
  const f = ctx.flattenValue_;
  assert.strictEqual(f(null), '');
  assert.strictEqual(f(undefined), '');
  assert.strictEqual(f('text'), 'text');
  assert.strictEqual(f(3.5), 3.5);
  assert.strictEqual(f(0), 0);
  assert.strictEqual(f(true), true);
  assert.strictEqual(f(false), false);
});

test('flattenValue_: attachments become URLs joined with ", "', () => {
  const { ctx } = load();
  const value = [
    { id: 'attA', url: 'https://v5.airtableusercontent.com/a/report.pdf', filename: 'report.pdf', size: 10, type: 'application/pdf' },
    { id: 'attB', url: 'https://v5.airtableusercontent.com/b/logo.png', filename: 'logo.png', size: 20, type: 'image/png',
      thumbnails: { small: { url: 'https://v5.airtableusercontent.com/b/small.png', width: 1, height: 1 } } },
  ];
  assert.strictEqual(ctx.flattenValue_(value),
    'https://v5.airtableusercontent.com/a/report.pdf, https://v5.airtableusercontent.com/b/logo.png');
});

test('flattenValue_: arrays, linked records, collaborators, buttons, errors, objects', () => {
  const { ctx } = load();
  const f = ctx.flattenValue_;
  assert.strictEqual(f(['Red', 'Blue']), 'Red, Blue');
  assert.strictEqual(f(['recAAA', null, '', 'recBBB']), 'recAAA, recBBB');
  assert.strictEqual(f([1, 2.5]), '1, 2.5');
  assert.strictEqual(f([]), '');
  assert.strictEqual(f({ id: 'usrX', email: 'a@example.com', name: 'A Person' }), 'A Person');
  assert.strictEqual(f({ id: 'usrX', email: 'a@example.com' }), 'a@example.com');
  assert.strictEqual(f([{ id: 'u1', email: 'a@example.com', name: 'One' }, { id: 'u2', email: 'b@example.com', name: 'Two' }]), 'One, Two');
  assert.strictEqual(f({ label: 'Open', url: 'https://example.com/x' }), 'https://example.com/x');
  assert.strictEqual(f({ specialValue: NaN }), 'NaN');
  assert.strictEqual(f({ error: '#ERROR!' }), '#ERROR!');
  assert.strictEqual(f({ text: '12345', type: 'upce' }), '12345');
  assert.strictEqual(f({ state: 'generated', isStale: false, value: 'summary' }), 'summary');
  assert.strictEqual(f({ a: 1 }), '{"a":1}');
});

test('protectText_: formula-looking text stays text, numbers untouched', () => {
  const { ctx } = load();
  assert.strictEqual(ctx.protectText_('=SUM(A1:A2)'), "'=SUM(A1:A2)");
  assert.strictEqual(ctx.protectText_('+1 555'), "'+1 555");
  assert.strictEqual(ctx.protectText_('-x'), "'-x");
  assert.strictEqual(ctx.protectText_('@cmd'), "'@cmd");
  assert.strictEqual(ctx.protectText_('plain'), 'plain');
  assert.strictEqual(ctx.protectText_(-5), -5);
  const off = load({ overrides: { PROTECT_TEXT: false } });
  assert.strictEqual(off.ctx.protectText_('=1+1'), '=1+1');
});

// ---- URL building and backoff ----------------------------------------------------------------
test('buildUrl_: encodes table, view, fields and offset', () => {
  const { ctx } = load({ overrides: { BASE_ID: 'appABC', TABLE: 'My table/1', VIEW: 'Grid view', FIELDS: ['Name', 'Due & date'] } });
  assert.strictEqual(ctx.buildUrl_(''), 'https://api.airtable.com/v0/appABC/My%20table%2F1?pageSize=100&view=Grid%20view&fields%5B%5D=Name&fields%5B%5D=Due%20%26%20date');
  assert.ok(ctx.buildUrl_('itrX/recY').endsWith('&offset=itrX%2FrecY'));
});

test('backoffMs_: exponential, capped at 30 s, honours Retry-After', () => {
  const { ctx } = load();
  assert.deepStrictEqual([0, 1, 2, 3, 4, 5].map((n) => ctx.backoffMs_(n)), [2000, 4000, 8000, 16000, 30000, 30000]);
  assert.strictEqual(ctx.backoffMs_(0, { 'Retry-After': '7' }), 7000);
  assert.strictEqual(ctx.backoffMs_(0, { 'Retry-After': '9999' }), 60000);
});

// ---- full runs against real recorded responses ------------------------------------------------
test('sync: paginates with offset and writes every record (real multi-page response)', () => {
  const h = load({ fetchImpl: replay, overrides: WORK });
  const expected = workPages.reduce((n, p) => n + p.records.length, 0);
  assert.ok(workPages.length >= 3, 'fixture should have several pages');
  assert.ok(workPages.slice(0, -1).every((p) => p.offset), 'all but the last page carry an offset');
  const out = h.ctx.syncAirtable();
  assert.strictEqual(out.records, expected);
  assert.strictEqual(h.calls.fetch.length, workPages.length);
  assert.ok(!h.calls.fetch[0].url.includes('offset='));
  workPages.slice(0, -1).forEach((p, i) => assert.ok(h.calls.fetch[i + 1].url.includes('offset=' + encodeURIComponent(p.offset))));
  assert.strictEqual(h.calls.fetch[0].options.headers.Authorization, 'Bearer test-token');
  const grid = h.sheets['Airtable'].grid();
  assert.strictEqual(grid.length, expected + 1, 'header + one row per record');
  assert.strictEqual(h.calls.sleep.length, workPages.length - 1, 'pause between pages only');
});

test('sync: header is the union of fields; arrays flattened; empty record gives blank row', () => {
  const h = load({ fetchImpl: replay, overrides: WORK });
  h.ctx.syncAirtable();
  const grid = h.sheets['Airtable'].grid();
  const header = grid[0];
  assert.deepStrictEqual(plain(header.slice().sort()), ['Client', 'Date', 'Hours', 'Notes', 'Status', 'Task']);
  const all = workPages.flatMap((p) => p.records);
  const col = (n) => header.indexOf(n);
  all.forEach((rec, i) => {
    const row = grid[i + 1];
    assert.strictEqual(row[col('Task')], rec.fields.Task === undefined ? '' : rec.fields.Task);
    assert.strictEqual(row[col('Hours')], rec.fields.Hours === undefined ? '' : rec.fields.Hours);
    assert.strictEqual(row[col('Client')], (rec.fields.Client || []).join(', '));
    assert.ok(row.every((c) => typeof c !== 'object'), 'no objects/arrays left in cells');
  });
  const empty = all.findIndex((r) => Object.keys(r.fields).length === 0);
  assert.ok(empty >= 0, 'fixture contains a record with no fields (real Airtable behaviour)');
  assert.ok(grid[empty + 1].every((c) => c === ''));
  assert.ok(typeof grid[1][col('Hours')] === 'number');
});

test('sync: A1 note carries last sync time and record count; header row frozen', () => {
  const h = load({ fetchImpl: replay, overrides: WORK });
  h.ctx.syncAirtable();
  const sheet = h.sheets['Airtable'];
  const note = sheet.notes.get('1,1');
  assert.match(note, /^Last sync: \d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z \(16 records from Airtable\)$/);
  assert.strictEqual(sheet.frozenRows, 1);
});

test('sync: VIEW and FIELDS restrict rows and columns (real response)', () => {
  const h = load({ fetchImpl: replay, overrides: { BASE_ID: WORK.BASE_ID, TABLE: 'Work log', VIEW: 'This week', FIELDS: ['Task', 'Hours', 'Status'] } });
  const out = h.ctx.syncAirtable();
  assert.strictEqual(out.records, 7);
  const grid = h.sheets['Airtable'].grid();
  assert.deepStrictEqual(plain(grid[0]), ['Task', 'Hours', 'Status']);
  assert.strictEqual(grid.length, 8);
  assert.ok(h.calls.fetch[0].url.includes('view=This%20week'));
});

test('sync: different base/table (Items), custom SHEET_NAME', () => {
  const h = load({ fetchImpl: replay, overrides: { BASE_ID: 'appWNdNoD1nqbuqbj', TABLE: 'Items', SHEET_NAME: 'Stock' } });
  const out = h.ctx.syncAirtable();
  const expected = JSON.parse(byUrl.get('https://api.airtable.com/v0/appWNdNoD1nqbuqbj/Items?pageSize=100').body).records.length;
  assert.strictEqual(out.records, expected);
  assert.strictEqual(h.sheets['Stock'].grid().length, expected + 1);
  assert.ok(!h.sheets['Airtable']);
});

test('sync: re-running clears old data (smaller result leaves no leftovers)', () => {
  const h = load({ fetchImpl: replay, overrides: WORK });
  h.ctx.syncAirtable();
  assert.strictEqual(h.sheets['Airtable'].grid().length, 17);
  h.ctx.TABLE = 'Work log'; h.ctx.PAGE_SIZE = 100; h.ctx.VIEW = 'This week'; h.ctx.FIELDS = ['Task', 'Hours', 'Status'];
  h.ctx.syncAirtable();
  const sheet = h.sheets['Airtable'];
  assert.strictEqual(sheet.clears, 2);
  assert.strictEqual(sheet.grid().length, 8);
  assert.strictEqual(sheet.grid()[0].length, 3);
});

// ---- errors, retries, safety ------------------------------------------------------------------
test('real 403/401 responses raise a readable error and leave the sheet untouched', () => {
  const h = load({ fetchImpl: replay, overrides: WORK });
  h.ctx.syncAirtable();
  const before = json(h.sheets['Airtable'].grid());
  h.ctx.TABLE = 'No such table'; h.ctx.PAGE_SIZE = 100;
  assert.throws(() => h.ctx.syncAirtable(), /HTTP 403.*BASE_ID and TABLE/s);
  const bad = load({ fetchImpl: replay, overrides: { ...WORK, PAGE_SIZE: 100 }, token: 'invalid-token' });
  // the recorded 401 was made with that token; the stub does not look at the header
  assert.throws(() => bad.ctx.syncAirtable(), /HTTP 401.*AIRTABLE_TOKEN/s);
  assert.strictEqual(json(h.sheets['Airtable'].grid()), before);
  assert.strictEqual(h.sheets['Airtable'].clears, 1);
});

test('missing token / placeholder BASE_ID give clear errors without any request', () => {
  const noToken = load({ fetchImpl: () => { throw new Error('should not fetch'); }, token: null });
  assert.throws(() => noToken.ctx.syncAirtable(), /AIRTABLE_TOKEN is missing/);
  const placeholder = load({ fetchImpl: () => { throw new Error('should not fetch'); } });
  assert.throws(() => placeholder.ctx.syncAirtable(), /Set BASE_ID/);
});

test('429: backs off 2 s, 4 s then succeeds (synthetic 429 in front of a real page)', () => {
  let n = 0;
  const fetchImpl = (url) => (n++ < 2
    ? { code: 429, body: '{"error":{"type":"RATE_LIMIT_REACHED","message":"Too many requests"}}', headers: {} }
    : replay(url));
  const h = load({ fetchImpl, overrides: { ...WORK, PAGE_SIZE: 100, VIEW: 'This week', FIELDS: ['Task', 'Hours', 'Status'] } });
  const out = h.ctx.syncAirtable();
  assert.strictEqual(out.records, 7);
  assert.deepStrictEqual(plain(h.calls.sleep), [2000, 4000]);
  assert.strictEqual(h.calls.fetch.length, 3);
});

test('429 forever: gives up after MAX_RETRIES and does not touch the sheet', () => {
  const h = load({ fetchImpl: () => ({ code: 429, body: '{"error":{"type":"RATE_LIMIT_REACHED"}}' }), overrides: WORK });
  assert.throws(() => h.ctx.syncAirtable(), /HTTP 429/);
  assert.deepStrictEqual(plain(h.calls.sleep), [2000, 4000, 8000, 16000, 30000]);
  assert.strictEqual(h.calls.fetch.length, 6);
  assert.ok(!h.sheets['Airtable']);
});

test('5xx is retried; Retry-After header is respected', () => {
  let n = 0;
  const fetchImpl = (url) => {
    n++;
    if (n === 1) return { code: 503, body: 'unavailable' };
    if (n === 2) return { code: 429, body: '{}', headers: { 'Retry-After': '3' } };
    return replay(url);
  };
  const h = load({ fetchImpl, overrides: { ...WORK, PAGE_SIZE: 100, VIEW: 'This week', FIELDS: ['Task', 'Hours', 'Status'] } });
  h.ctx.syncAirtable();
  assert.deepStrictEqual(plain(h.calls.sleep), [2000, 3000]);
});

test('a failure on page 3 leaves previously synced data in place', () => {
  const h = load({ fetchImpl: replay, overrides: WORK });
  h.ctx.syncAirtable();
  const before = json(h.sheets['Airtable'].grid());
  let page = 0;
  h.ctx.UrlFetchApp.fetch = (url) => { if (++page === 3) return { getResponseCode: () => 500, getContentText: () => 'boom', getHeaders: () => ({}) }; return { getResponseCode: () => 200, getContentText: () => replay(url).body, getHeaders: () => ({}) }; };
  h.ctx.MAX_RETRIES = 0;
  assert.throws(() => h.ctx.syncAirtable(), /HTTP 500/);
  assert.strictEqual(json(h.sheets['Airtable'].grid()), before);
});

test('runtime guard: stops before the 6 minute limit without touching the sheet', () => {
  const h = load({ fetchImpl: replay, overrides: { ...WORK, MAX_RUNTIME_MS: -1 } });
  assert.throws(() => h.ctx.syncAirtable(), /Stopped after/);
  assert.ok(!h.sheets['Airtable']);
});

test('formula-looking text is protected in the written sheet', () => {
  const page = { records: [{ id: 'rec1', createdTime: '2026-01-01T00:00:00.000Z', fields: { Name: '=HYPERLINK("http://x")', Qty: 2, Note: '-1 off' } }] };
  const h = load({ fetchImpl: () => ({ code: 200, body: json(page) }), overrides: { BASE_ID: 'appTEST', TABLE: 'T' } });
  h.ctx.syncAirtable();
  assert.deepStrictEqual(plain(h.sheets['Airtable'].grid()), [['Name', 'Qty', 'Note'], ["'=HYPERLINK(\"http://x\")", 2, "'-1 off"]]);
});

test('attachment, collaborator and checkbox columns end to end (shapes from Airtable docs)', () => {
  const page = { records: [{ id: 'rec1', createdTime: '2026-01-01T00:00:00.000Z', fields: {
    Name: 'Row 1', Done: true,
    Files: [{ id: 'att1', url: 'https://v5.airtableusercontent.com/x/a.pdf', filename: 'a.pdf', size: 1, type: 'application/pdf' },
            { id: 'att2', url: 'https://v5.airtableusercontent.com/x/b.png', filename: 'b.png', size: 2, type: 'image/png' }],
    Owner: { id: 'usr1', email: 'owner@example.com', name: 'Owner' },
    Tags: ['a', 'b'] } }] };
  const h = load({ fetchImpl: () => ({ code: 200, body: json(page) }), overrides: { BASE_ID: 'appTEST', TABLE: 'T' } });
  h.ctx.syncAirtable();
  assert.deepStrictEqual(plain(h.sheets['Airtable'].grid()), [
    ['Name', 'Done', 'Files', 'Owner', 'Tags'],
    ['Row 1', true, 'https://v5.airtableusercontent.com/x/a.pdf, https://v5.airtableusercontent.com/x/b.png', 'Owner', 'a, b'],
  ]);
});

test('large tables are written in chunks of 2000 rows without losing any', () => {
  const records = Array.from({ length: 4500 }, (_, i) => ({ id: 'rec' + i, createdTime: '2026-01-01T00:00:00.000Z', fields: { N: i, S: 'row ' + i } }));
  const pages = []; for (let i = 0; i < records.length; i += 100) pages.push(records.slice(i, i + 100));
  let p = 0;
  const fetchImpl = () => { const body = { records: pages[p] }; if (++p < pages.length) body.offset = 'off' + p; return { code: 200, body: json(body) }; };
  const h = load({ fetchImpl, overrides: { BASE_ID: 'appTEST', TABLE: 'T' } });
  const out = h.ctx.syncAirtable();
  assert.strictEqual(out.records, 4500);
  assert.strictEqual(h.calls.fetch.length, 45);
  const sheet = h.sheets['Airtable'];
  assert.strictEqual(sheet.grid().length, 4501);
  assert.strictEqual(sheet.writes.length, 1 + 3, 'header + 3 chunks (2000, 2000, 500)');
  assert.deepStrictEqual(plain(sheet.grid()[4500]), [4499, 'row 4499']);
});

test('empty table: sheet is cleared and A1 still stamped', () => {
  const h = load({ fetchImpl: () => ({ code: 200, body: '{"records":[]}' }), overrides: { BASE_ID: 'appTEST', TABLE: 'T' } });
  const out = h.ctx.syncAirtable();
  assert.strictEqual(out.records, 0);
  assert.match(h.sheets['Airtable'].notes.get('1,1'), /0 records/);
});

// ---- triggers ---------------------------------------------------------------------------------
test('installTrigger: hourly by default, daily on request, replaces existing triggers', () => {
  const h = load();
  h.ctx.installTrigger();
  assert.deepStrictEqual(plain(h.calls.triggers[0]), { fn: 'syncAirtable', steps: [['everyHours', 1]] });
  h.ctx.installTrigger();
  assert.strictEqual(h.calls.deleted.length, 1);
  assert.strictEqual(h.existingTriggers.length, 1);
  h.ctx.TRIGGER_EVERY = 'daily'; h.ctx.TRIGGER_HOUR = 7;
  h.ctx.installTrigger();
  assert.deepStrictEqual(plain(h.calls.triggers[2]), { fn: 'syncAirtable', steps: [['everyDays', 1], ['atHour', 7]] });
  assert.strictEqual(h.existingTriggers.length, 1);
  h.ctx.removeTriggers();
  assert.strictEqual(h.existingTriggers.length, 0);
});

// ---- optional live check ----------------------------------------------------------------------
if (process.env.AIRTABLE_TOKEN) {
  test('LIVE: real Airtable API through Code.gs, same row count as the recorded run', () => {
    const h = load({ fetchImpl: (u, o) => curlFetch(u, o), overrides: WORK, token: process.env.AIRTABLE_TOKEN });
    const out = h.ctx.syncAirtable();
    assert.strictEqual(out.records, 16);
    assert.strictEqual(h.sheets['Airtable'].grid().length, 17);
  });
} else {
  console.log('skip - live check (set AIRTABLE_TOKEN to enable)');
}

console.log('\n' + passed + ' passed, ' + failed + ' failed');
process.exit(failed ? 1 : 0);
