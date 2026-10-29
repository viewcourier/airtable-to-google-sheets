/**
 * Airtable -> Google Sheets sync (Google Apps Script)
 * See README.md for setup.
 *
 * Copies one Airtable table (or one view of it) into one sheet of a Google
 * Spreadsheet. Run syncAirtable() by hand, or run installTrigger() once to
 * refresh the sheet on a schedule.
 *
 * The Airtable personal access token is NOT stored in this file. Add it under
 * Project Settings > Script properties as AIRTABLE_TOKEN.
 * Required token scope: data.records:read, with access to the base.
 *
 * Copyright (c) 2026 BASIT B.V. - MIT License
 */

// ---------------------------------------------------------------------------
// Configuration: edit these
// ---------------------------------------------------------------------------
var BASE_ID = 'appXXXXXXXXXXXXXX';  // starts with "app"; find it in the Airtable API docs URL
var TABLE = 'Table 1';              // table name or table ID (tblXXXXXXXXXXXXXX)
var VIEW = '';                      // view name or ID; '' = every record in the table
var SHEET_NAME = 'Airtable';        // sheet (tab) to write to; created if missing
var FIELDS = [];                    // e.g. ['Name', 'Status']; [] = every field, in the order Airtable returns them
var SPREADSHEET_ID = '';            // '' = the spreadsheet this script is attached to

var TRIGGER_EVERY = 'hourly';       // 'hourly' or 'daily' (used by installTrigger)
var TRIGGER_HOUR = 6;               // hour of day (0-23, script time zone) for 'daily'

// Tuning: normally leave as is
var PAGE_SIZE = 100;                // Airtable maximum
var PAUSE_MS = 250;                 // pause between pages, keeps under 5 requests per second
var MAX_RETRIES = 5;                // retries for 429 and 5xx responses
var MAX_RUNTIME_MS = 5 * 60 * 1000; // stop before Apps Script's 6 minute limit
var LIST_SEPARATOR = ', ';          // joins multi-value cells (multiple select, links, attachment URLs)
var PROTECT_TEXT = true;            // prefix text that starts with = + - @ so Sheets keeps it as text, not a formula

// ---------------------------------------------------------------------------
// Entry points
// ---------------------------------------------------------------------------

/** Fetch the table from Airtable and replace the contents of the sheet. */
function syncAirtable() {
  var token = PropertiesService.getScriptProperties().getProperty('AIRTABLE_TOKEN');
  if (!token) {
    throw new Error('AIRTABLE_TOKEN is missing. Add it under Project Settings > Script properties.');
  }
  if (!BASE_ID || BASE_ID === 'appXXXXXXXXXXXXXX') {
    throw new Error('Set BASE_ID at the top of the script.');
  }

  var started = Date.now();
  var records = fetchAllRecords_(token, started);
  var table = buildTable_(records, FIELDS);
  // The sheet is touched only after every page has loaded, so a failed or
  // timed-out run leaves the previous data in place.
  writeSheet_(table);

  var summary = { records: records.length, columns: table.header.length, seconds: Math.round((Date.now() - started) / 100) / 10 };
  Logger.log('Synced ' + summary.records + ' records, ' + summary.columns + ' columns in ' + summary.seconds + ' s');
  return summary;
}

/** Run once: replaces any existing syncAirtable trigger with a new time-based one. */
function installTrigger() {
  removeTriggers();
  var builder = ScriptApp.newTrigger('syncAirtable').timeBased();
  if (TRIGGER_EVERY === 'daily') {
    builder.everyDays(1).atHour(TRIGGER_HOUR);
  } else {
    builder.everyHours(1);
  }
  builder.create();
}

/** Removes every trigger that runs syncAirtable. */
function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'syncAirtable') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// ---------------------------------------------------------------------------
// Pure helpers (no Google services): unit-tested with Node, see test/
// ---------------------------------------------------------------------------

/**
 * Turn one Airtable cell value into something a spreadsheet cell can hold.
 *  - null/undefined      -> ''
 *  - string/number/bool  -> unchanged
 *  - attachments         -> their URLs, joined
 *  - arrays              -> items flattened and joined with LIST_SEPARATOR
 *  - collaborator        -> name, or email if there is no name
 *  - button              -> its URL
 *  - formula error       -> the error text (for example "#ERROR!" or "NaN")
 *  - anything else       -> JSON text
 */
function flattenValue_(value) {
  if (value === null || value === undefined) return '';
  var type = typeof value;
  if (type === 'string' || type === 'number' || type === 'boolean') return value;
  if (Array.isArray(value)) {
    return value
      .map(function (item) { return String(flattenValue_(item)); })
      .filter(function (text) { return text !== ''; })
      .join(LIST_SEPARATOR);
  }
  if (type === 'object') {
    if (typeof value.url === 'string') return value.url;                  // attachment, button
    if (typeof value.name === 'string' && value.name) return value.name;  // collaborator, user
    if (typeof value.email === 'string') return value.email;
    if (typeof value.specialValue !== 'undefined') return String(value.specialValue);
    if (typeof value.error === 'string') return value.error;
    if (typeof value.text === 'string') return value.text;                // barcode
    if (typeof value.value !== 'undefined' && value.value !== null && typeof value.value !== 'object') {
      return value.value;                                                  // AI text field
    }
    return JSON.stringify(value);
  }
  return String(value);
}

/** Keep text that looks like a formula (= + - @) as plain text. */
function protectText_(value) {
  if (PROTECT_TEXT && typeof value === 'string' && /^[=+\-@]/.test(value)) return "'" + value;
  return value;
}

/**
 * Build the header row and the data rows from Airtable records.
 * Airtable leaves empty fields out of a record, so with no explicit field list
 * the header is the union of all field names, in order of first appearance.
 */
function buildTable_(records, fieldNames) {
  var header = (fieldNames || []).slice();
  if (!header.length) {
    var seen = {};
    records.forEach(function (record) {
      Object.keys(record.fields || {}).forEach(function (name) {
        if (!seen.hasOwnProperty(name)) {
          seen[name] = true;
          header.push(name);
        }
      });
    });
  }
  var rows = records.map(function (record) {
    var fields = record.fields || {};
    return header.map(function (name) {
      return protectText_(flattenValue_(fields[name]));
    });
  });
  return { header: header, rows: rows };
}

/** URL for one page of records. */
function buildUrl_(offset) {
  var enc = encodeURIComponent;
  var query = ['pageSize=' + PAGE_SIZE];
  if (VIEW) query.push('view=' + enc(VIEW));
  FIELDS.forEach(function (name) { query.push('fields%5B%5D=' + enc(name)); });
  if (offset) query.push('offset=' + enc(offset));
  return 'https://api.airtable.com/v0/' + enc(BASE_ID) + '/' + enc(TABLE) + '?' + query.join('&');
}

/** Wait time before retry number `attempt` (0-based): 2, 4, 8, 16, 30 s, or Retry-After. */
function backoffMs_(attempt, headers) {
  var retryAfter = headers && (headers['Retry-After'] || headers['retry-after']);
  var seconds = parseInt(retryAfter, 10);
  if (seconds > 0) return Math.min(seconds, 60) * 1000;
  return Math.min(2000 * Math.pow(2, attempt), 30000);
}

function describeError_(code, body) {
  var message = '';
  try {
    var parsed = JSON.parse(body);
    message = parsed.error && (parsed.error.message || parsed.error.type || parsed.error);
    if (typeof message !== 'string') message = JSON.stringify(message);
  } catch (e) {
    message = String(body).slice(0, 200);
  }
  var hint = '';
  if (code === 401) hint = ' Check the AIRTABLE_TOKEN script property.';
  if (code === 403) hint = ' Check that the token has the data.records:read scope and access to this base, and that BASE_ID and TABLE exist.';
  if (code === 404) hint = ' Check BASE_ID, TABLE and VIEW (names are case sensitive).';
  if (code === 422) hint = ' Check VIEW and FIELDS.';
  message = String(message).replace(/[.\s]+$/, "");
  return 'Airtable returned HTTP ' + code + ': ' + message + '.' + hint;
}

// ---------------------------------------------------------------------------
// Google services
// ---------------------------------------------------------------------------

/** GET one URL, retrying 429 and 5xx with backoff. Returns parsed JSON. */
function fetchJson_(url, token) {
  for (var attempt = 0; ; attempt++) {
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    if (code === 200) return JSON.parse(response.getContentText());
    if ((code === 429 || code >= 500) && attempt < MAX_RETRIES) {
      Utilities.sleep(backoffMs_(attempt, response.getHeaders && response.getHeaders()));
      continue;
    }
    throw new Error(describeError_(code, response.getContentText()));
  }
}

/** Follow Airtable's offset cursor until the last page. */
function fetchAllRecords_(token, started) {
  var records = [];
  var offset = '';
  do {
    if (Date.now() - started > MAX_RUNTIME_MS) {
      throw new Error('Stopped after ' + Math.round(MAX_RUNTIME_MS / 60000) + ' minutes with ' + records.length +
        ' records loaded. The sheet was not changed. Use a smaller VIEW or fewer FIELDS.');
    }
    var page = fetchJson_(buildUrl_(offset), token);
    Array.prototype.push.apply(records, page.records || []);
    offset = page.offset || '';
    if (offset) Utilities.sleep(PAUSE_MS);
  } while (offset);
  return records;
}

/** Replace the sheet contents with the header and rows, and stamp A1 with the sync time. */
function writeSheet_(table) {
  var spreadsheet = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
  var stamp = 'Last sync: ' + new Date().toISOString() + ' (' + table.rows.length + ' records from Airtable)';

  sheet.clearContents();
  sheet.getRange(1, 1).clearNote();

  var columns = table.header.length;
  if (columns) {
    var CHUNK = 2000;
    sheet.getRange(1, 1, 1, columns).setValues([table.header]);
    for (var start = 0; start < table.rows.length; start += CHUNK) {
      var part = table.rows.slice(start, start + CHUNK);
      sheet.getRange(2 + start, 1, part.length, columns).setValues(part);
    }
    sheet.setFrozenRows(1);
  }
  sheet.getRange(1, 1).setNote(stamp);
  SpreadsheetApp.flush();
}
