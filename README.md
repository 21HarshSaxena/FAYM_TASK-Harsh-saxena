# Multi-Item Return Automation Agent

Reads pending return line items from a spreadsheet a local Excel file or a live Google Sheet, same interface either way checks each one's return eligibility, and (once the platform adapters are filled in) drives the Amazon/Flipkart return flow through a real, human supervised Chrome session, writing the outcome back into the row it came from.

Built for the Faym assignment. The 16 column schema, the `Status` / `Return Status` split, and the outcome vocabulary (`Placed`, `Already Cancelled & Refunded`, `Out of window`, `Not yet delivered`, `Support Needed`, `Failed`) all come from the real "Faym Status Test Orders" sheet, not guessed from the brief.

## Current status

**Working and tested today:** the task queue, the eligibility check, and both storage backends all covered by `test_skeleton.py`, with no live site, no browser, and no credentials involved.

**Not built yet:** the actual clicking. `AmazonAdapter` and `FlipkartAdapter`'s `detect_flow_type()` and `initiate_return()` are intentionally `NotImplementedError` stubs until real DOM selectors go in for each site. In practice:
- A row whose eligibility alone resolves it window's closed, delivery date's blank, or the dates just don't parse is handled correctly end to end, and the browser never opens.
- A row that's genuinely eligible reaches the adapter and raises `NotImplementedError`. That's the one piece of work left before a live run can place a real return.

## How it works

- **Eligibility is computed, not assumed.** Return Window is stored as free text ("10 Days", "7 Day"), so the deadline is `delivery_date + window_days`. A date without a year ("27 June") is parsed against the current year; a range ("5-6 July") uses the *later* date assuming the item arrived as late as stated is the direction that won't prematurely mark something out-of-window.
- **Blank vs. unparseable stay distinct.** A blank delivery date means `NOT_YET_DELIVERED` the row stays `Pending` and is retried on a future run. A delivery date that's present but doesn't match any known format means `UNKNOWN` flagged `Needs Review` instead of guessed at. The raw string is kept alongside the parsed date specifically so these two don't get conflated into "no date."
- **Pandas' `NaN`, not `None`.** Reading Excel with `dtype=str` still yields float `NaN` for blank cells (not the string `"nan"`) and `NaN` is truthy, so a naive `value or ""` doesn't catch it. `_clean_str()` normalizes this everywhere; found via a failing test, not by inspection.
- **Outcomes are per line item.** Each item lands on `Placed`, `Already Cancelled & Refunded`, `Out of window`, `Support Needed`, `Failed`, or the non-terminal `Not yet delivered`. An order only rolls up to "done" once every one of its line items has hit a real terminal state, so one item failing or needing a human never blocks the rest of the order.
- **Retries know the difference between "broke" and "not built yet."** A transient error gets up to 2 retries (3 attempts total) with a randomized 3–8 second backoff before falling to `Needs Review`. A `NotImplementedError` is deliberately let through instead it means the code isn't there yet, not that the site glitched, and it shouldn't be quietly retried or logged like a normal failure.
- **Non headless, human in the loop, nothing stored.** Every real run opens a visible Chrome window and reuses a saved login OTP is always typed by a human, once and a short randomized pause (0.8–2.4s) runs between items.

## Two storage backends, one interface

`ExcelTaskQueue` and `GoogleSheetTaskQueue` (`sheets_backend.py`) both implement the same four-method `TaskQueue` interface, so nothing else in the agent needs to know which one it's holding. They differ in one important way:

- **Excel** `write_back()` updates the in memory table, the file itself is only written once, in `save()`, after every item in the run has been processed.
- **Google Sheets** `write_back()` fires one batched API call *per item*, immediately, so `save()` is a no op. A crash mid run only risks the one row actively being worked on, not the returns already placed earlier in that same run worth having for automation that touches real refunds.

Setting up your own live sheet to test against, so nothing touches the real one, is a 10 minute one time thing the full click by click version is in `SETUP_SHEETS.md`.

## Setup

Python 3.10+ and Git.

```bash
git clone https://github.com/<your-username>/<your-repo>.git
cd <your-repo>
pip install -r requirements.txt
python -m playwright install chromium
```

`.gitignore` already covers `service_account.json`, `.env`, `*.xlsx`, `sessions/`, `venv/`, `__pycache__/`, and `agent_run.log` (the log file every run writes, alongside stdout) so credentials, saved logins, and local spreadsheet copies won't get committed by accident.

## Running the tests

```bash
python test_skeleton.py
```

This is the real "does this work" check today there's no `--dry-run` flag on the agent itself yet. It builds a fixture Excel file on the fly and runs all 4 rows through the full eligibility + outcome state machine with a `MockAdapter` standing in for a platform, then repeats the equivalent checks against a mocked Google Sheet no live API call either way. `ALL TESTS PASSED` at the end means the queue, the eligibility logic, and both backends are behaving correctly.

## Running it for real

**Excel (default):**
```bash
python agent_skeleton.py
```
Reads from `RETURNS_EXCEL_PATH` (defaults to `returns_tasks.xlsx` in the current folder):
```bash
export RETURNS_EXCEL_PATH=path/to/your_copy.xlsx      # Windows: set RETURNS_EXCEL_PATH=...
python agent_skeleton.py
```

**Google Sheets:**
```bash
export RETURNS_BACKEND=sheets
export RETURNS_SHEET_ID=your_sheet_id_here
python agent_skeleton.py
```
`GOOGLE_SERVICE_ACCOUNT_JSON` defaults to `service_account.json` in the current folder the same name `SETUP_SHEETS.md` has you save the downloaded key as.

**Before the first real run against a platform**, save a login session a manual, one time step by design:
```python
from agent_skeleton import save_login_session
save_login_session("Flipkart", "https://www.flipkart.com/account/login")
```
A visible Chrome window opens; log in (OTP included) by hand, then press Enter in the terminal. The session is saved to `sessions/flipkart_state.json` (configurable via `RETURNS_SESSION_DIR`) and reused on every run after that `run_agent()` won't start against a platform it has no saved session for.

## Adding a platform

1. Subclass `PlatformAdapter` and implement `detect_flow_type(page, order_id)` and `initiate_return(page, item)`.
2. Add it to the `ADAPTERS` dict, e.g. `"Myntra": MyntraAdapter`.

Any row whose `Platform` isn't a key in `ADAPTERS` gets flagged `Needs Review` automatically, rather than silently skipped.

## Troubleshooting quick reference

| Symptom | Cause / fix |
|---|---|
| `NotImplementedError` on an eligible row | Expected for now see "Current status", that platform's adapter isn't filled in yet |
| `RuntimeError: No saved session for <Platform>` | Run `save_login_session()` for that platform once first |
| `ValueError: ... is missing required columns` | Header row got edited or reordered compare against the 16 columns in `SETUP_SHEETS.md` |
| `KeyError: 'RETURNS_SHEET_ID'` | Set `RETURNS_SHEET_ID` before running with `RETURNS_BACKEND=sheets` |
| `PermissionError` on save | The `.xlsx` is open in Excel close it and rerun |
| "No pending tasks." | Every row's `Status` is something other than blank / `Pending` / `To Do` |

## Project layout

```
.
├── agent_skeleton.py     # task queue, eligibility engine, Excel backend, adapter interface + stubs
├── sheets_backend.py     # Google Sheets backend drop in replacement for ExcelTaskQueue
├── test_skeleton.py      # fixture-driven tests, both backends, no live API or browser needed
├── SETUP_SHEETS.md       # click by click guide to connecting a live Google Sheet
├── requirements.txt
└── .gitignore
```
