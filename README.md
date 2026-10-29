# airtable-to-google-sheets

A single Google Apps Script file (`Code.gs`) that copies an Airtable table, or one view of a table, into a Google Sheet and refreshes it on a schedule. No add-on, no third-party service, no dependencies.

- Reads the Airtable token from Script Properties, never from the code.
- Pages through the Airtable API (100 records per request, `offset` cursor).
- Flattens linked records, multiple selects, collaborators, buttons and attachments into plain cell text. Attachments become their URLs, joined with `, `.
- Clears the old data and writes a header row plus all rows. The sheet is only touched after every page has loaded, so a failed or timed-out run leaves the previous data in place.
- Puts the last sync time and record count in a note on cell A1.
- Retries HTTP 429 (rate limit) and 5xx responses with backoff.
- Includes `installTrigger()` to run the sync hourly or daily.

Prefer not to edit code? A browser form that generates this script with your values filled in is here: [Airtable to Google Sheets script generator](https://viewcourier.com/tools/airtable-to-google-sheets?utm_source=github&utm_medium=readme).

## Setup

1. **Create an Airtable token.** At [airtable.com/create/tokens](https://airtable.com/create/tokens) create a personal access token with the scope `data.records:read` and access to the base you want to copy. Copy the token.
2. **Open the script editor.** In the Google Sheet choose *Extensions > Apps Script* and paste the contents of `Code.gs` over the default file.
3. **Add the token.** In the Apps Script editor open *Project Settings* (gear icon), scroll to *Script properties*, click *Add script property*, and set the name `AIRTABLE_TOKEN` and the value to your token.
4. **Edit the configuration at the top of `Code.gs`.** Set `BASE_ID` (starts with `app`), `TABLE` (name or `tbl...` ID), and optionally `VIEW` and `SHEET_NAME`. Then pick `syncAirtable` in the function list and click *Run*. Google asks you to authorize the script the first time. The sheet fills and cell A1 gets a note with the sync time.
5. **Schedule it (optional).** Set `TRIGGER_EVERY` to `'hourly'` or `'daily'` (and `TRIGGER_HOUR` for daily), pick `installTrigger` in the function list and click *Run* once. Run `removeTriggers` to stop the schedule.

## Configuration

| Constant | Meaning | Default |
| --- | --- | --- |
| `BASE_ID` | Airtable base ID, starts with `app` | placeholder, must be set |
| `TABLE` | Table name or ID | `Table 1` |
| `VIEW` | View name or ID. Only that view's records, in that view's order | `''` (all records) |
| `SHEET_NAME` | Sheet tab to write; created if it does not exist | `Airtable` |
| `FIELDS` | List of field names to copy, in that order | `[]` (every field) |
| `SPREADSHEET_ID` | Use only if the script is not attached to the target spreadsheet | `''` (attached one) |
| `TRIGGER_EVERY`, `TRIGGER_HOUR` | Used by `installTrigger()` | `'hourly'`, `6` |
| `LIST_SEPARATOR` | Separator for multi-value cells | `', '` |
| `PROTECT_TEXT` | Prefix text that starts with `=`, `+`, `-` or `@` with an apostrophe so Sheets stores it as text instead of evaluating it | `true` |

## How values are converted

| Airtable value | Cell |
| --- | --- |
| Text, number, checkbox | As is (checkbox true shows as TRUE; unchecked is empty because Airtable omits it) |
| Empty field | Empty cell |
| Multiple select, linked records, lookups | Items joined with `, ` (linked records appear as record IDs, as the API returns them) |
| Attachments | URLs joined with `, ` |
| Collaborator | Name, or email if there is no name |
| Button | Its URL |
| Formula error | The error text, for example `#ERROR!` |
| Anything else (objects) | JSON text |

Without `FIELDS`, the columns are the union of the field names that appear in the returned records, in order of first appearance. A field that is empty in every record has no column. Set `FIELDS` for a fixed set and order of columns.

Dates arrive from Airtable as text such as `2026-10-31`; Google Sheets converts text like that to dates when it is written, the same as if you typed it.

## Limits

- **Apps Script runtime: 6 minutes per run.** The script stops itself after 5 minutes (`MAX_RUNTIME_MS`) without changing the sheet. Google also caps total trigger runtime per day for each account; see Google's Apps Script quotas page. A page of 100 records takes about a second, so tens of thousands of records are the practical ceiling. Use a smaller `VIEW` for bigger tables.
- **Airtable API: 5 requests per second per base.** The script pauses 250 ms between pages and retries a 429 after 2, 4, 8, 16 and 30 seconds (or the `Retry-After` value if Airtable sends one).
- **One direction only.** Edits in the sheet are overwritten on the next sync. Put your own formulas, charts and pivot tables on other sheets and reference the synced sheet.
- **Attachment URLs expire** after about two hours. The sheet holds the links, not the files.
- A cell holds at most 50,000 characters; a spreadsheet at most 10 million cells.

## Tests

The Apps Script cannot run outside Google, so `test/test.js` loads `Code.gs` into a Node `vm` with stand-ins for `UrlFetchApp`, `SpreadsheetApp`, `PropertiesService`, `ScriptApp` and `Utilities`, and replays real Airtable API responses stored in `test/fixtures/real-responses.json` (multi-page results, a view with a field list, a different table, a real 401 and a real 403). Rate-limit and attachment cases use hand-written responses in the shape the Airtable docs describe.

```
node test/test.js
```

Needs Node 18 or newer, nothing to install. To also run one check against the live Airtable API, set `AIRTABLE_TOKEN` to a token that can read the recorded demo base; `test/record-fixtures.js` regenerates the fixtures. The Google-side behaviour (sheet writes, triggers) is exercised through the stand-ins only, so run the script once in a scratch spreadsheet before relying on it.

## FAQ

**Where is the token stored?** In Script Properties of the Apps Script project, not in the code. Anyone who can edit the script project can read it, so share the sheet with care. Use a token limited to `data.records:read` on the one base.

**I get "HTTP 403" or "HTTP 404".** The token has no access to that base, lacks the `data.records:read` scope, or `BASE_ID`, `TABLE` or `VIEW` is wrong. Names are case sensitive; a table ID (`tbl...`) is safer than a name.

**"Authorization is required" when running.** Run `syncAirtable` from the editor once and accept the Google permission prompt. The script needs access to spreadsheets and to external services (`UrlFetchApp`). A time-driven trigger cannot show the prompt.

**Can it sync several tables?** Copy the project or add more functions that set `TABLE` and `SHEET_NAME` before calling `syncAirtable()`; the simplest route is one Apps Script project per table.

**Can I get field types or the record ID as columns?** Not in this script. It writes the values Airtable returns. Add a formula field `RECORD_ID()` in Airtable if you need the ID in the sheet.

**Does it delete rows removed in Airtable?** Yes. Every run replaces the full contents of the sheet.

**Why do some cells start with an apostrophe in the data?** That is `PROTECT_TEXT`. Text such as `=SUM(1,2)` coming from Airtable would otherwise become a live formula in Sheets. Sheets hides the leading apostrophe in the cell. Set it to `false` if you want values written as they are.

## Email the data instead

Want it emailed as CSV, Excel or PDF on a schedule instead? [ViewCourier](https://viewcourier.com/?utm_source=github&utm_medium=readme&utm_campaign=airtable-to-google-sheets) sends an Airtable view to any list of recipients on a schedule, with no script to maintain.

## License

MIT. Copyright (c) 2026 BASIT B.V.
