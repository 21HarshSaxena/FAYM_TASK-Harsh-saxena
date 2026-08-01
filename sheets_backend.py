"""
Google Sheets storage backend for the returns agent implements the same
TaskQueue interface as ExcelTaskQueue (agent_skeleton.py), so run_agent()
works identically regardless of which backend it's handed.

One time setup (see SETUP.md for the full click by click version):
  1. Google Cloud Console -> new project -> enable "Google Sheets API"
  2. IAM & Admin -> Service Accounts -> create one -> Keys -> Add key (JSON)
     -> download it, keep it OUT of git (.gitignore already covers this)
  3. Open your Google Sheet -> Share -> paste the service account's email
     (looks like xxx@yyy.iam.gserviceaccount.com) -> give it Editor access
  4. pip install gspread google-auth

Usage:
    from sheets_backend import GoogleSheetTaskQueue
    queue = GoogleSheetTaskQueue(sheet_id="...", credentials_path="service_account.json")
    run_agent(queue)   # from agent_skeleton.py — identical call whether Excel or Sheets

Design note on write timing: write_back() writes to the sheet IMMEDIATELY
(one batched API call covering that single row's changed cells), rather than
queuing everything for a final save(). If the process crashes mid run, only
the row it was actively processing is at risk of re attempt not every row
processed so far in that run, which would otherwise risk the agent refiring
an ALREADY SUCCESSFUL return against a live platform on the next run.
"""

from __future__ import annotations

import datetime as dt
from typing import Optional

import gspread
from google.oauth2.service_account import Credentials

from agent_skeleton import (
    LineItem, TaskQueue, TaskState, parse_loose_date, _safe_float, _safe_int,
)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

REQUIRED_HEADERS = [
    "Address", "Contact Number", "Product Link", "Amount", "No of Product",
    "Order date", "Order Id", "Delivery date", "Return Window", "Status",
    "Platform", "Refund ID", "Return Status", "Refund Amount", "Timestamp", "Log",
]
PENDING_VALUES = {"", "Pending", "To Do"}
TERMINAL_TASK_STATES = {TaskState.DONE.value, TaskState.NEEDS_REVIEW.value}


def records_to_line_items(records: list[dict], year_hint: int) -> list[LineItem]:
    """Pure function: sheet rows (as dicts from gspread's get_all_records(),
    which is already 1 row = 1 dict keyed by header) -> pending LineItems.

    Deliberately separated from any network call so this can be unit-tested
    against fixture data with no live Sheets connection — see test_skeleton.py.
    row_index is the real sheet row number (data starts at row 2: row 1 is
    the header), needed later so write_back() edits the correct row.
    """
    items = []
    for i, row in enumerate(records):
        status = str(row.get("Status") or "").strip()
        if status not in PENDING_VALUES:
            continue
        delivery_raw = str(row.get("Delivery date") or "")
        items.append(LineItem(
            row_index=i + 2,
            platform=str(row.get("Platform") or "").strip(),
            order_id=str(row.get("Order Id") or "").strip(),
            product_link=str(row.get("Product Link") or "").strip(),
            return_window=str(row.get("Return Window") or "").strip(),
            delivery_date=parse_loose_date(delivery_raw, year_hint),
            delivery_date_raw=delivery_raw,
            order_date=parse_loose_date(row.get("Order date"), year_hint),
            address=str(row.get("Address") or ""),
            contact_number=str(row.get("Contact Number") or ""),
            amount=_safe_float(row.get("Amount")),
            no_of_product=_safe_int(row.get("No of Product")),
            task_state=status,
        ))
    return items


def order_done_from_records(records: list[dict], order_id: str) -> bool:
    """Pure function version of the order-rollup check — same testability
    reasoning as records_to_line_items above."""
    rows = [r for r in records if str(r.get("Order Id", "")).strip() == order_id]
    if not rows:
        return False
    return all(str(r.get("Status", "")).strip() in TERMINAL_TASK_STATES for r in rows)


def build_write_back_batch(item: LineItem, header: list[str]) -> list[dict]:
    """Pure function: LineItem -> gspread batch_update payload. Split out so
    the cell-mapping logic (row_index + column position -> A1 range) can be
    tested without a live sheet."""
    values = {
        "Refund ID": item.refund_id or "",
        "Return Status": item.return_status or "",
        "Refund Amount": item.refund_amount if item.refund_amount is not None else "",
        "Status": item.task_state,
        "Timestamp": dt.datetime.now().isoformat(timespec="seconds"),
        "Log": item.note,
    }
    batch = []
    for col_name, value in values.items():
        col_idx = header.index(col_name) + 1  # gspread columns are 1-indexed
        a1 = gspread.utils.rowcol_to_a1(item.row_index, col_idx)
        batch.append({"range": a1, "values": [[value]]})
    return batch


class GoogleSheetTaskQueue(TaskQueue):
    def __init__(self, sheet_id: str, credentials_path: str, worksheet_name: str = "Sheet1"):
        creds = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
        client = gspread.authorize(creds)
        self.sheet = client.open_by_key(sheet_id).worksheet(worksheet_name)
        header = self.sheet.row_values(1)
        missing = [c for c in REQUIRED_HEADERS if c not in header]
        if missing:
            raise ValueError(f"Sheet is missing required columns: {missing}")
        self._header = header

    def pending_line_items(self) -> list[LineItem]:
        records = self.sheet.get_all_records()
        return records_to_line_items(records, year_hint=dt.date.today().year)

    def write_back(self, item: LineItem) -> None:
        batch = build_write_back_batch(item, self._header)
        self.sheet.batch_update(batch)  # one API call for this row's ~6 changed cells

    def order_is_fully_done(self, order_id: str) -> bool:
        records = self.sheet.get_all_records()
        return order_done_from_records(records, order_id)

    def save(self) -> None:
        # write_back() already persists per item (see module docstring) —
        # nothing left to flush. Kept as a no-op so run_agent()'s
        # queue.save() call works identically for either backend.
        pass
