# Connecting the agent to a live Google Sheet

Doing this once. Takes about only 10 minutes.

## 1. Make your OWN sheet
1. Open that sheet (view-only is enough)
2. **File → Make a copy** → save it into your own Google Drive
3. This gives you a clean copy with the correct 16 columns already in place
   (Address, Contact Number, Product Link, Amount, No of Product, Order date,
   Order Id, Delivery date, Return Window, Status, Platform, Refund ID,
   Return Status, Refund Amount, Timestamp, Log) but none of the other
   candidate's Refund IDs/timestamps in it.
4. Delete rows you don't want to test against, or add your own up to you.
5. Copy the new sheet's ID from its URL:
   `docs.google.com/spreadsheets/d/`**`THIS-LONG-PART`**`/edit`
   Save that string somewhere you'll need it in step 9.

## 2. Google Cloud project

1. Go to **console.cloud.google.com** → sign in with any Google account
2. Top-left dropdown next to "Google Cloud" → **New Project**
3. Name it anything (e.g. `returns-agent`) → **Create** → wait a few seconds,
   then make sure the new project is selected in that same top-left dropdown

## 3. Turn on the Sheets API

1. Search bar at the top → type **Google Sheets API** → click the result
2. Click **Enable** (if it already says "Manage", it's already on — skip)

## 4. Create the service account (this is what your code logs in as)

1. Left sidebar (☰) → **IAM & Admin → Service Accounts**
2. **+ Create Service Account** (top)
3. Name it anything (e.g. `returns-agent-bot`) → **Create and Continue**
4. The "Grant access" and "Grant users access" steps → leave both blank →
   **Continue** → **Done**

## 5. Create its key (this is the file your code actually uses)

1. You're now looking at the Service Accounts list → click the one you just made
2. Top tab bar → **Keys**
3. **Add Key → Create new key → JSON → Create**
4. A `.json` file downloads automatically this is a password, treat it
   like one. **Never commit it to GitHub**
5. Move that downloaded file into your `returns agent` folder and rename it
   to exactly: `service_account.json`

## 6. Share your sheet with the service account

1. Open that `service_account.json` file in Notepad (or any text editor)
2. Find the line that says `"client_email": "..."` — copy that email address.
   It'll look like `returns-agent-bot@your-project-id.iam.gserviceaccount.com`
3. Go back to **your own copy** of the sheet (from step 1) → **Share** button
   (top right) → paste that email → set it to **Editor** → **Send**
   (if Google warns the recipient has no Google Workspace / can't be
   notified, that's expected for a service account send anyway)

## 7. Install the new packages
```
venv\Scripts\activate            (Windows)
pip install -r requirements.txt
```

## 8. Point the agent at your sheet

Two environment variables tell the agent which sheet to use. In the same
terminal, before running anything:

**Windows (Command Prompt):**
```
set RETURNS_BACKEND=sheets
set RETURNS_SHEET_ID=paste_your_sheet_id_from_step_1_here
python agent_skeleton.py
```

**Mac/Linux:**
```
export RETURNS_BACKEND=sheets
export RETURNS_SHEET_ID=paste_your_sheet_id_from_step_1_here
python3 agent_skeleton.py
```

Note: these `set`/`export` lines only last for the current terminal window
you'll need to rerun them (or add them to a `.env` file / your system's
environment variables) each time you open a fresh terminal.

## Sanity check it worked, before touching a real browser

```python
from sheets_backend import GoogleSheetTaskQueue
q = GoogleSheetTaskQueue(sheet_id="your_sheet_id", credentials_path="service_account.json")
items = q.pending_line_items()
print(f"Found {len(items)} pending rows")
for i in items:
    print(i.order_id, i.platform, i.return_window, i.delivery_date, i.task_state)
```

If this prints your rows with no error, the connection is good the only
remaining piece is the same one as before: `agent_skeleton.py`'s two
platform adapters still need real Amazon/Flipkart selectors filled in
(the `NotImplementedError` spots) before a live run can actually place a
return. Everything up to that point reading the sheet, eligibility
checks, writing results back is real and working now.

## Common errors

- **`PermissionError` / `403`** → you shared the sheet with the wrong email,
  or shared your Drive folder but not the sheet itself. Re check step 6.
- **`SpreadsheetNotFound`** → wrong Sheet ID, or you copied the shared
  sheet's ID instead of your own copy's. Re check step 1.5.
- **`ValueError: Sheet is missing required columns`** → your copy's header
  row got edited/reordered. Compare it against the 16 names listed in step 1.
