"""
Browser Agent for Automated Multi-Item Returns
=====================================================================
Implements the orchestrator, per-line-item state machine, Excel I/O, and
platform-adapter interface described in Returns_Automation_Agent_Design.docx
(Sections 3-5, 7),updated against the REAL
column schema and real Return Status values observed in the live "Faym
Status Test Orders" sheet (Address, Contact Number, Product Link, Amount,
No of Product, Order date, Order Id, Delivery date, Return Window, Status,
Platform, Refund ID, Return Status, Refund Amount, Timestamp, Log).

Storage is now pluggable: ExcelTaskQueue (local file) or
GoogleSheetTaskQueue (sheets_backend.py) both implement the TaskQueue
interface, so run_agent() doesn't care which one it's handed.

Credentials: NEVER hardcoded. Login is a manual, human-in-the-loop step
(see save_login_session below); everything after that reuses the saved
session via Playwright's storage_state.
"""

from __future__ import annotations

import os
import re
import time
import random
import logging
import datetime as dt
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Optional

import pandas as pd
from playwright.sync_api import sync_playwright, BrowserContext, Page

# ---------------------------------------------------------------------------
EXCEL_PATH = os.environ.get("RETURNS_EXCEL_PATH", "returns_tasks.xlsx")
SESSION_DIR = Path(os.environ.get("RETURNS_SESSION_DIR", "./sessions"))
MAX_RETRIES = 2
RETRY_BACKOFF_SEC = (3, 8)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[logging.FileHandler("agent_run.log"), logging.StreamHandler()],
)
log = logging.getLogger("returns_agent")


# ---------------------------------------------------------------------------
# Status model split into two axes, matching how the real sheet actually
# uses its "Status" vs "Return Status" columns
# ---------------------------------------------------------------------------
class TaskState(str, Enum):
    """Coarse workflow state -> sheet's "Status" column."""
    PENDING = "Pending"
    IN_PROGRESS = "In Progress"
    DONE = "Done"
    NEEDS_REVIEW = "Needs Review"


class ReturnResult(str, Enum):
    """Fine-grained outcome -> sheet's "Return Status" column. Values below
    are exactly what's observed in the live sheet, not guessed from the brief."""
    PLACED = "Placed"
    ALREADY_REFUNDED = "Already Cancelled & Refunded"
    OUT_OF_WINDOW = "Out of window"
    NOT_YET_DELIVERED = "Not yet delivered"
    SUPPORT_NEEDED = "Support Needed"
    FAILED = "Failed"


# Terminal from the ORDER rollup perspective (Status can move to Done/Needs
# Review). NOT_YET_DELIVERED is deliberately excluded it means "come back
# on a later run", not "finished" the row should stay Pending, not terminal.
TERMINAL_RESULTS = {
    ReturnResult.PLACED, ReturnResult.ALREADY_REFUNDED,
    ReturnResult.OUT_OF_WINDOW, ReturnResult.SUPPORT_NEEDED, ReturnResult.FAILED,
}


@dataclass
class LineItem:
    row_index: int
    platform: str
    order_id: str
    product_link: str
    return_window: str          # e.g. "10 Days" free text as it appears in the sheet
    delivery_date: Optional[dt.date] = None
    delivery_date_raw: str = ""  # raw sheet text, kept so eligibility can tell
                                  # "blank" (not yet delivered) apart from
                                  # "present but unparseable" (needs review)
    order_date: Optional[dt.date] = None
    address: str = ""
    contact_number: str = ""
    amount: Optional[float] = None
    no_of_product: Optional[int] = None
    task_state: str = TaskState.PENDING.value
    attempt_count: int = 0
    refund_id: Optional[str] = None
    return_status: Optional[str] = None
    refund_amount: Optional[float] = None
    note: str = ""


# ---------------------------------------------------------------------------
# Eligibility 3 way, using the REAL fields (Return Window + Delivery date)
# instead of placeholder stub. Dates in the live sheet are messy on
# purpose (real data): "27 June" has no year, "5-6 July" is a range. Rather
# than silently guess wrong, unparseable dates fall through to NEEDS_REVIEW
# consistent with the design doc's "flag, don't guess" principle.
# ---------------------------------------------------------------------------
class Eligibility(str, Enum):
    ELIGIBLE = "eligible"
    NOT_YET_DELIVERED = "not_yet_delivered"
    OUT_OF_WINDOW = "out_of_window"
    UNKNOWN = "unknown"          # date couldn't be parsed confidently -> human review


def parse_return_window_days(text: str) -> Optional[int]:
    """'10 Days' / '7 Day' -> 10 / 7. None if unparseable."""
    if not text:
        return None
    m = re.search(r"(\d+)", str(text))
    return int(m.group(1)) if m else None


def parse_loose_date(text: str, year_hint: int) -> Optional[dt.date]:
    """Best-effort parse of the sheet's inconsistent date formats.
    '27 June' -> uses year_hint. '5-6 July' (a range) -> takes the LATER date
    (6 July): the safer assumption for eligibility is "delivered as late as
    stated", so we don't prematurely mark something out-of-window.
    Returns None if the text doesn't match a recognized shape — callers must
    NOT treat None as "not delivered"; see LineItem.delivery_date_raw."""
    if not text or not str(text).strip():
        return None
    text = str(text).strip()
    if "-" in text and re.search(r"[A-Za-z]", text):
        text = text.split("-")[-1].strip()  # "5-6 July" -> "6 July"
    for fmt in ("%d %B", "%d %b"):
        try:
            parsed = dt.datetime.strptime(text, fmt)
            return dt.date(year_hint, parsed.month, parsed.day)
        except ValueError:
            continue
    for fmt in ("%d %B %Y", "%d %b %Y"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def check_eligibility(item: LineItem, today: Optional[dt.date] = None) -> Eligibility:
    today = today or dt.date.today()
    if item.delivery_date is None:
        if not item.delivery_date_raw.strip():
            return Eligibility.NOT_YET_DELIVERED  # field genuinely blank
        return Eligibility.UNKNOWN                 # text present but unparseable don't guess
    window_days = parse_return_window_days(item.return_window)
    if window_days is None:
        return Eligibility.UNKNOWN
    deadline = item.delivery_date + dt.timedelta(days=window_days)
    return Eligibility.ELIGIBLE if today <= deadline else Eligibility.OUT_OF_WINDOW


# ---------------------------------------------------------------------------
# Storage backend interface Excel and Google Sheets both implement this,
# so run_agent() is backend.
# ---------------------------------------------------------------------------
class TaskQueue(ABC):
    @abstractmethod
    def pending_line_items(self) -> list[LineItem]: ...

    @abstractmethod
    def write_back(self, item: LineItem) -> None: ...

    @abstractmethod
    def order_is_fully_done(self, order_id: str) -> bool: ...

    @abstractmethod
    def save(self) -> None: ...


class ExcelTaskQueue(TaskQueue):
    """Local-file backend — unchanged in spirit from v1, updated to the real
    column names so the SAME schema works whether you're on Excel or Sheets."""
    REQUIRED_COLUMNS = [
        "Address", "Contact Number", "Product Link", "Amount", "No of Product",
        "Order date", "Order Id", "Delivery date", "Return Window", "Status",
        "Platform", "Refund ID", "Return Status", "Refund Amount", "Timestamp", "Log",
    ]
    PENDING_VALUES = {"", "Pending", "To Do"}

    def __init__(self, path: str):
        self.path = path
        self.df = pd.read_excel(path, dtype=str)
        missing = [c for c in self.REQUIRED_COLUMNS if c not in self.df.columns]
        if missing:
            raise ValueError(f"Excel is missing required columns: {missing}")

    def _row_to_item(self, idx: int, row: pd.Series) -> LineItem:
        order_year = dt.date.today().year
        delivery_raw = _clean_str(row.get("Delivery date"))
        return LineItem(
            row_index=idx,
            platform=_clean_str(row["Platform"]).strip(),
            order_id=_clean_str(row["Order Id"]).strip(),
            product_link=_clean_str(row["Product Link"]).strip(),
            return_window=_clean_str(row["Return Window"]).strip(),
            delivery_date=parse_loose_date(delivery_raw, order_year),
            delivery_date_raw=delivery_raw,
            order_date=parse_loose_date(_clean_str(row.get("Order date")), order_year),
            address=_clean_str(row.get("Address")),
            contact_number=_clean_str(row.get("Contact Number")),
            amount=_safe_float(row.get("Amount")),
            no_of_product=_safe_int(row.get("No of Product")),
            task_state=_clean_str(row.get("Status")) or TaskState.PENDING.value,
        )

    def pending_line_items(self) -> list[LineItem]:
        status_col = self.df["Status"].fillna("")
        pending = self.df[status_col.isin(self.PENDING_VALUES)]
        return [self._row_to_item(idx, row) for idx, row in pending.iterrows()]

    def write_back(self, item: LineItem) -> None:
        r = item.row_index
        self.df.at[r, "Refund ID"] = str(item.refund_id) if item.refund_id else ""
        self.df.at[r, "Return Status"] = str(item.return_status) if item.return_status else ""
        self.df.at[r, "Refund Amount"] = str(item.refund_amount) if item.refund_amount is not None else ""
        self.df.at[r, "Status"] = str(item.task_state)
        self.df.at[r, "Timestamp"] = dt.datetime.now().isoformat(timespec="seconds")
        self.df.at[r, "Log"] = item.note

    def order_is_fully_done(self, order_id: str) -> bool:
        rows = self.df[self.df["Order Id"] == order_id]
        if len(rows) == 0:
            return False
        return bool(rows["Status"].isin({TaskState.DONE.value, TaskState.NEEDS_REVIEW.value}).all())

    def save(self) -> None:
        self.df.to_excel(self.path, index=False)


def _clean_str(v) -> str:
    """pandas' read_excel(dtype=str) still yields float NaN for blank cells
    (not the string 'nan') — and NaN is truthy, so `v or ""` does NOT catch
    it. This normalizes any pandas cell value to a clean string, treating
    NaN/None as empty. Found via test failure, not inspection — see test_v2.py."""
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return str(v)


def _safe_float(v) -> Optional[float]:
    if v is None or (isinstance(v, float) and pd.isna(v)) or v in ("", "N/A"):
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def _safe_int(v) -> Optional[int]:
    f = _safe_float(v)
    return int(f) if f is not None else None


# ---------------------------------------------------------------------------
# Platform adapters
# ---------------------------------------------------------------------------
class ReturnOutcome:
    def __init__(self, success: bool, refund_id: str = None, refund_amount: float = None,
                 result: ReturnResult = None, transient_error: bool = False, note: str = ""):
        self.success = success
        self.refund_id = refund_id
        self.refund_amount = refund_amount
        self.result = result
        self.transient_error = transient_error
        self.note = note


class PlatformAdapter(ABC):
    name: str

    def __init__(self, context: BrowserContext):
        self.context = context
        self._batch_mode_cache: dict[str, bool] = {}

    @abstractmethod
    def detect_flow_type(self, page: Page, order_id: str) -> bool: ...

    @abstractmethod
    def initiate_return(self, page: Page, item: LineItem) -> ReturnOutcome: ...

    def is_batch_order(self, page: Page, order_id: str) -> bool:
        if order_id not in self._batch_mode_cache:
            self._batch_mode_cache[order_id] = self.detect_flow_type(page, order_id)
        return self._batch_mode_cache[order_id]

    @staticmethod
    def human_pause():
        time.sleep(random.uniform(0.8, 2.4))


class AmazonAdapter(PlatformAdapter):
    name = "Amazon"

    def detect_flow_type(self, page: Page, order_id: str) -> bool:
        raise NotImplementedError("Needs live-DOM selectors — design doc S4.1")

    def initiate_return(self, page: Page, item: LineItem) -> ReturnOutcome:
        raise NotImplementedError("Needs live-DOM selectors — design doc S4")


class FlipkartAdapter(PlatformAdapter):
    name = "Flipkart"

    def detect_flow_type(self, page: Page, order_id: str) -> bool:
        raise NotImplementedError("Needs live-DOM selectors — design doc S4.1")

    def initiate_return(self, page: Page, item: LineItem) -> ReturnOutcome:
        raise NotImplementedError("Needs live-DOM selectors — design doc S4")


ADAPTERS: dict[str, type[PlatformAdapter]] = {"Amazon": AmazonAdapter, "Flipkart": FlipkartAdapter}


# ---------------------------------------------------------------------------
# Session management manual, human in the loop login
# (incl. OTP), persisted via Playwright storage_state.
# ---------------------------------------------------------------------------
def storage_state_path(platform: str) -> Path:
    return SESSION_DIR / f"{platform.lower()}_state.json"


def save_login_session(platform: str, login_url: str):
    SESSION_DIR.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        page.goto(login_url)
        input(f"[{platform}] Complete login (incl. OTP, if prompted), then press Enter here...")
        context.storage_state(path=str(storage_state_path(platform)))
        browser.close()
    log.info(f"Session saved for {platform} -> {storage_state_path(platform)}")


def load_context(playwright, platform: str) -> BrowserContext:
    state_path = storage_state_path(platform)
    if not state_path.exists():
        raise RuntimeError(f"No saved session for {platform}. Run save_login_session() once first.")
    browser = playwright.chromium.launch(headless=False)
    return browser.new_context(storage_state=str(state_path))


# ---------------------------------------------------------------------------
# Orchestrator now driven by the 3 way eligibility check and the fuller
# ReturnResult vocabulary. NOT_YET_DELIVERED deliberately leaves the item
# Pending (retried on a future run) rather than closing it out.
# ---------------------------------------------------------------------------
def process_line_item(adapter: PlatformAdapter, page: Page, item: LineItem, queue: TaskQueue):
    item.task_state = TaskState.IN_PROGRESS.value
    log.info(f"[{item.platform}] Order {item.order_id}: starting")

    elig = check_eligibility(item)
    if elig == Eligibility.NOT_YET_DELIVERED:
        item.task_state = TaskState.PENDING.value  # stays pending -> retried later, not abandoned
        item.return_status = ReturnResult.NOT_YET_DELIVERED.value
        item.note = "Order is not yet delivered; will re-check on a future run."
        log.info("  -> Not yet delivered, left Pending for a later run")
        queue.write_back(item)
        return
    if elig == Eligibility.OUT_OF_WINDOW:
        item.task_state = TaskState.DONE.value
        item.return_status = ReturnResult.OUT_OF_WINDOW.value
        item.note = "Return window has closed"
        log.info("  -> Out of window (terminal); order continues with remaining items")
        queue.write_back(item)
        return
    if elig == Eligibility.UNKNOWN:
        item.task_state = TaskState.NEEDS_REVIEW.value
        item.return_status = ReturnResult.FAILED.value
        item.note = f"Could not confidently parse delivery date / return window — flagging rather than guessing (delivery_date={item.delivery_date!r}, return_window={item.return_window!r})"
        log.warning("  -> Unparseable dates, flagged for human review (not guessed)")
        queue.write_back(item)
        return

    while True:
        item.attempt_count += 1
        try:
            outcome = adapter.initiate_return(page, item)
        except NotImplementedError:
            raise
        except Exception as e:  # noqa: BLE001
            outcome = ReturnOutcome(success=False, transient_error=True, note=str(e))

        if outcome.success:
            item.task_state = TaskState.DONE.value
            item.refund_id = outcome.refund_id
            item.refund_amount = outcome.refund_amount
            item.return_status = ReturnResult.PLACED.value
            item.note = "Placed successfully"
            break
        if outcome.result in (ReturnResult.ALREADY_REFUNDED, ReturnResult.SUPPORT_NEEDED):
            item.task_state = (TaskState.DONE.value if outcome.result == ReturnResult.ALREADY_REFUNDED
                                else TaskState.NEEDS_REVIEW.value)
            item.return_status = outcome.result.value
            item.note = outcome.note
            break
        if outcome.transient_error and item.attempt_count <= MAX_RETRIES:
            delay = random.uniform(*RETRY_BACKOFF_SEC)
            log.warning(f"  transient error (attempt {item.attempt_count}/{MAX_RETRIES}): {outcome.note} — retrying in {delay:.1f}s")
            time.sleep(delay)
            continue
        item.task_state = TaskState.NEEDS_REVIEW.value
        item.return_status = ReturnResult.FAILED.value
        item.note = outcome.note or "Failed after retries exhausted"
        break

    queue.write_back(item)
    log.info(f"  -> Status={item.task_state} / Return Status={item.return_status}")


def run_agent(queue: TaskQueue):
    """Backend-agnostic: pass an ExcelTaskQueue or a GoogleSheetTaskQueue
    (sheets_backend.py) — everything below is identical either way."""
    items = queue.pending_line_items()
    if not items:
        log.info("No pending tasks.")
        return

    by_platform: dict[str, list[LineItem]] = {}
    for it in items:
        by_platform.setdefault(it.platform, []).append(it)

    with sync_playwright() as p:
        for platform, platform_items in by_platform.items():
            adapter_cls = ADAPTERS.get(platform)
            if not adapter_cls:
                log.error(f"No adapter for platform '{platform}' — flagging for review")
                for it in platform_items:
                    it.task_state = TaskState.NEEDS_REVIEW.value
                    it.note = f"No adapter for platform '{platform}'"
                    queue.write_back(it)
                continue

            context = load_context(p, platform)
            adapter = adapter_cls(context)
            page = context.new_page()

            orders_seen: set[str] = set()
            for item in platform_items:
                process_line_item(adapter, page, item, queue)
                adapter.human_pause()
                orders_seen.add(item.order_id)

            for order_id in orders_seen:
                if queue.order_is_fully_done(order_id):
                    log.info(f"Order {order_id}: all line items terminal -> Done")

            context.close()

    queue.save()
    log.info("Run complete.")


if __name__ == "__main__":
    backend = os.environ.get("RETURNS_BACKEND", "excel").lower()
    if backend == "sheets":
        from sheets_backend import GoogleSheetTaskQueue
        sheet_id = os.environ["RETURNS_SHEET_ID"]
        creds_path = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "service_account.json")
        q = GoogleSheetTaskQueue(sheet_id=sheet_id, credentials_path=creds_path)
    else:
        q = ExcelTaskQueue(EXCEL_PATH)
    run_agent(q)
