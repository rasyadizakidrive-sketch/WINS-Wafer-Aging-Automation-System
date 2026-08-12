"""
ZPPMYR0520 -- Production Order Variance Check reconciliation.

For a batch of Production Orders, reads the actual posted Target /
Yield / Scrap quantities from SAP and calculates the difference for
each, so this can flag over/under-posting without a person
manually opening this transaction, pressing SUM, and doing the
subtraction by hand for every order one at a time.

Processes ONE order per SAP execution, resetting to a fresh selection
screen between orders via StartTransaction() (the programmatic
equivalent of typing /nZPPMYR0520). This is a deliberate choice, not
the fastest possible approach: an earlier version tried entering all
orders at once via SAP's Multiple Selection popup (a real VBS
recording showed a batch producing one row per order in a single
execution), but filling that popup via its Paste toolbar button
produced a corrupted, single mangled value in practice rather than
correctly splitting the entered orders across rows -- and neither
recording available for this actually showed that specific paste
interaction succeeding, only the button being pressed. Rather than
keep guessing at which exact button or table-cell path makes that
work, this processes orders one at a time using only mechanisms
already confirmed reliable elsewhere in this project.

Field paths and grid interactions below are confirmed against a real
VBS recording of a single-order run (Plant P200, Order 16032574):
  1. Set Plant (S_WERKS-LOW) and Order (S_AUFNR-LOW).
  2. Clear the Start Date restriction via its context menu ("DELACTX"),
     not a plain .text = "" -- SAP handles a select-option range field
     differently from a normal text field for this.
  3. Execute (tbar[1]/btn[8]).
  4. On the results grid, double-click the GR Quantity (WEMNG) cell on
     row 0 to drill into individual goods-receipt line items.
  5. On the resulting detail grid, select both the Yield (LMNGA) and
     Scrap (XMNGA) columns together, then press SUM once (&MB_SUM) --
     both totals get calculated in the same pass.
  6. Target Qty (GAMNG) is read from row 0 of the ORIGINAL results
     grid, not the detail grid.

LMNGA and XMNGA are standard SAP order-confirmation fields
(AFRU-LMNGA "Yield to be confirmed", AFRU-XMNGA "Scrap to be
confirmed") -- independently confirmed against SAP's own data
dictionary, not just inferred from the recording. GAMNG matches the
same field name po_creation.py already uses elsewhere in this project
for a Production Order's target quantity, a second, independent
confirmation for that one. BUDAT (Posting Date) is likewise confirmed
directly from a VBS recording's selectColumn call on this same detail
grid, not guessed.

One thing still inferred rather than directly confirmed: exactly which
row the SUM total lands on in the detail grid (see read_tally_for_row()
below for why).
"""

import time
import os
from datetime import datetime

import config
import logger
from SAP.sap_manager import start_transaction


# ==========================================================
# FIELD PATHS -- confirmed against the VBS recording
# ==========================================================

FIELD_PLANT = "wnd[0]/usr/ctxtS_WERKS-LOW"
FIELD_ORDER = "wnd[0]/usr/ctxtS_AUFNR-LOW"
FIELD_START_DATE = "wnd[0]/usr/ctxtS_GSTRP-LOW"

EXECUTE_BUTTON = "wnd[0]/tbar[1]/btn[8]"

# The results grid after Execute -- always exactly one row (row 0),
# since exactly one order is entered per execution.
GRID_RESULTS = "wnd[0]/shellcont[0]/shell/shellcont[0]/shell"
COLUMN_GR_QTY = "WEMNG"     # Goods Receipt quantity -- double-click target
COLUMN_TARGET = "GAMNG"     # Target Qty, read from row 0

# The detail grid that appears after drilling into GR Quantity --
# genuinely a different grid object at a different path, not the same
# one reused.
GRID_DETAIL = "wnd[0]/shellcont[0]/shell/shellcont[1]/shell/shellcont[2]/shell"
COLUMN_YIELD = "LMNGA"      # AFRU-LMNGA, "Yield to be confirmed"
COLUMN_SCRAP = "XMNGA"      # AFRU-XMNGA, "Scrap to be confirmed"
COLUMN_POSTING_DATE = "BUDAT"   # Posting Date -- confirmed via a real
                                # VBS recording's selectColumn call on
                                # this exact detail grid, not a guess.


class POTallyReader:
    """
    Opens ZPPMYR0520 once (a dedicated SAP window, managed the same way
    as every other module here -- see sap_manager.py). tally_batch() is
    the entry point for processing multiple orders: it processes them
    one at a time, resetting to a fresh selection screen between each
    via StartTransaction() -- see tally_batch()'s docstring for why.
    """

    def __init__(self):
        self.session = None

    # --------------------------------------------------
    # OPEN
    # --------------------------------------------------
    def open_transaction(self):

        print(f"[INFO] Opening {config.PO_TALLY_TCODE}...")

        self.session = start_transaction("ZPPMYR0520", config.PO_TALLY_TCODE)

        self.wait_until_ready()

        print(f"[OK] {config.PO_TALLY_TCODE} Ready.")

    # --------------------------------------------------
    # WAIT UNTIL SAP READY
    # (same pattern already proven in po_creation.py)
    # --------------------------------------------------
    def wait_until_ready(self, timeout=10):

        start = time.time()

        while time.time() - start < timeout:
            try:
                if not self.session.Busy:
                    return
            except Exception:
                pass
            time.sleep(0.2)

        raise TimeoutError("SAP screen is not ready.")

    # --------------------------------------------------
    # CHECK POPUP
    # (same pattern already proven in po_creation.py)
    # --------------------------------------------------
    def check_popup(self):

        try:
            if self.session.Children.Count > 1:
                popup = self.session.findById("wnd[1]")
                msg = popup.Text
                try:
                    popup.sendVKey(0)
                except Exception:
                    pass
                print(f"[WARN] Popup : {msg}")
                return msg
        except Exception:
            pass

        return ""

    # --------------------------------------------------
    # ENTER ONE ORDER'S SELECTION CRITERIA
    # --------------------------------------------------
    def enter_selection_single(self, po_number, plant=None):
        """
        Sets Plant and ONE Order number directly into the main
        selection field -- confirmed working from the original
        single-order recording. Deliberately not using SAP's Multiple
        Selection popup: filling it via its Paste toolbar button
        produced a corrupted, single mangled value in practice rather
        than correctly splitting entered orders across multiple rows,
        and neither VBS recording actually shows that specific
        interaction succeeding, only the button being pressed. Rather
        than keep guessing at which exact button or table-cell path
        makes that work, tally_batch() calls this once per order
        instead, resetting the screen between orders via
        StartTransaction() -- slower, but every step here is something
        already confirmed working, not a new guess.
        """

        plant = plant or config.SAP_PLANT

        self.session.findById(FIELD_PLANT).text = plant

        self.session.findById(FIELD_ORDER).text = str(po_number).strip()

        # Clearing the Start Date restriction -- a select-option range
        # field needs its context-menu "Delete" action, not a plain
        # .text = "" the way an ordinary text field would;
        # selectContextMenuItem runs on the PARENT container
        # (wnd[0]/usr), not the field itself.
        try:
            date_field = self.session.findById(FIELD_START_DATE)
            date_field.setFocus()
            date_field.caretPosition = 0
            date_field.showContextMenu()
            self.session.findById("wnd[0]/usr").selectContextMenuItem("DELACTX")
        except Exception as e:
            print(f"[WARN] Could not clear Start Date restriction: {e}")

    # --------------------------------------------------
    # EXECUTE
    # --------------------------------------------------
    def execute(self):

        print("[INFO] Executing...")

        self.session.findById(EXECUTE_BUTTON).press()

        self.wait_until_ready()

        popup_message = self.check_popup()

        if popup_message:
            raise Exception(f"SAP rejected the selection: {popup_message}")

    # --------------------------------------------------
    # READ ONE ROW'S TARGET / YIELD / SCRAP
    # --------------------------------------------------
    def read_tally_for_row(self, row_index):
        """
        Reads Target/Yield/Scrap for the order at this specific row of
        the results grid, via COM (GetCellValue) rather than clipboard.

        Confirmed from the mass-tally recording: row 0 sets
        currentCellColumn = "WEMNG" explicitly before double-clicking;
        every row after that only needs currentCellRow set (the column
        selection carries over from before), then double-click again.

        One thing is inferred rather than directly confirmed: exactly
        which row the SUM total lands on in the detail grid. &MB_SUM on
        a GuiGridView without pre-existing row grouping appends a
        single total row at the very end, so this reads from
        grid.RowCount - 1 -- consistent with both recordings scrolling
        to a high row number (firstVisibleRow) right after pressing
        SUM, but neither recording shows the actual read, only the
        navigation toward it. If totals come back wrong on a real run,
        this is the one spot to check first.
        """

        results_grid = self.session.findById(GRID_RESULTS)

        if row_index == 0:
            results_grid.currentCellColumn = COLUMN_GR_QTY
        else:
            results_grid.currentCellRow = row_index

        results_grid.doubleClickCurrentCell()

        self.wait_until_ready()

        detail_grid = self.session.findById(GRID_DETAIL)

        detail_grid.setCurrentCell(-1, COLUMN_YIELD)
        detail_grid.selectColumn(COLUMN_YIELD)
        detail_grid.selectColumn(COLUMN_SCRAP)
        detail_grid.pressToolbarButton("&MB_SUM")

        # &MB_SUM restructures the grid (adds a total row); give SAP a
        # moment to settle before trusting anything about it.
        self.wait_until_ready()

        total_row = detail_grid.RowCount - 1

        yield_qty = self._parse_number(detail_grid.GetCellValue(total_row, COLUMN_YIELD))
        scrap_qty = self._parse_number(detail_grid.GetCellValue(total_row, COLUMN_SCRAP))

        # Latest Posting Date, not the last row's date -- a real
        # screenshot of this exact grid showed BUDAT out of order, so
        # scanning every line item for the maximum is correct
        # regardless of row order; reading the last row would not be.
        posting_date = self._latest_posting_date(detail_grid)

        target_qty = self._parse_number(results_grid.GetCellValue(row_index, COLUMN_TARGET))

        return {
            "target": target_qty,
            "yield_qty": yield_qty,
            "scrap_qty": scrap_qty,
            "posting_date": posting_date,
        }

    @staticmethod
    def _latest_posting_date(detail_grid):
        """
        Scans every line-item row for BUDAT and returns the latest as
        an actual datetime object -- not a string -- so Excel can
        store it as a real date value rather than text.

        Scrolls the grid in chunks via firstVisibleRow before reading
        each one, rather than calling GetCellValue across a large row
        range without ever bringing those rows into view -- a
        genuinely large grid (hundreds of rows per the SUM total)
        returned a date earlier than one clearly visible on screen,
        and GuiGridView can return a stale value for a row well
        outside the currently-rendered window without raising an
        exception, which is why nothing showed up as skipped or
        failed to parse before.

        Also re-reads RowCount after scrolling to the current last
        known row, since some large SAP grids under-report RowCount
        until they've actually been scrolled through at least once.

        Writes every row it reads to Logs/posting_date_debug.log --
        not just a sample -- since this app is packaged with
        console=False (no visible terminal), meaning ordinary print()
        diagnostics from an earlier attempt at this were being written
        to a console that doesn't exist anywhere the person running
        this could ever see or share. This file is the actual fix for
        that: if the wrong date still comes out after this, the log
        will show the exact raw text SAP returned for every row, which
        settles definitively whether the read itself is wrong or
        something later in the pipeline is.

        SAP's BUDAT display format here is dd.mm.yyyy (e.g.
        "02.08.2026"); any row that's blank or fails to parse is
        skipped rather than aborting the whole scan.
        """

        debug_lines = [f"\n=== Posting Date scan @ {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ==="]

        row_count_before = detail_grid.RowCount

        try:
            detail_grid.firstVisibleRow = max(0, row_count_before - 1)
        except Exception:
            pass

        row_count_after = detail_grid.RowCount
        total_row = row_count_after - 1  # exclude the SUM row

        debug_lines.append(f"RowCount before/after scroll-to-end: {row_count_before} / {row_count_after}")

        latest = None
        parsed_count = 0

        CHUNK_SIZE = 10  # comfortably covers what's rendered per scroll position

        row = 0
        while row < total_row:
            try:
                detail_grid.firstVisibleRow = row
            except Exception:
                pass

            chunk_end = min(row + CHUNK_SIZE, total_row)

            for r in range(row, chunk_end):
                try:
                    raw = detail_grid.GetCellValue(r, COLUMN_POSTING_DATE)
                    text = str(raw).strip()
                    if not text:
                        debug_lines.append(f"row {r}: (blank)")
                        continue
                    current = datetime.strptime(text, "%d.%m.%Y")
                    parsed_count += 1
                    is_new_max = latest is None or current > latest
                    debug_lines.append(f"row {r}: raw={raw!r} parsed={current.strftime('%d.%m.%Y')}" + (" <- NEW MAX" if is_new_max else ""))
                    if is_new_max:
                        latest = current
                except Exception as e:
                    debug_lines.append(f"row {r}: raw={raw!r} FAILED TO PARSE ({e})")
                    continue

            row = chunk_end

        skipped_count = total_row - parsed_count
        debug_lines.append(
            f"RESULT: {total_row} rows scanned, {parsed_count} parsed, "
            f"{skipped_count} skipped, latest = {latest.strftime('%d.%m.%Y') if latest else None}"
        )

        try:
            debug_path = os.path.join(logger.LOG_DIR, "posting_date_debug.log")
            os.makedirs(logger.LOG_DIR, exist_ok=True)
            with open(debug_path, "a", encoding="utf-8") as f:
                f.write("\n".join(debug_lines) + "\n")
        except Exception:
            pass  # diagnostic logging itself must never break the actual read

        return latest

    @staticmethod
    def _parse_number(raw):

        try:
            cleaned = str(raw).replace(",", "").strip()
            return float(cleaned)
        except (ValueError, TypeError):
            return 0.0

    @staticmethod
    def _status_for_difference(difference):

        if difference == 0:
            return config.PO_TALLY_STATUS_MATCH
        elif difference > 0:
            return config.PO_TALLY_STATUS_LESS
        else:
            return config.PO_TALLY_STATUS_OVER

    # --------------------------------------------------
    # TALLY A BATCH OF PRODUCTION ORDERS (the main entry point)
    # --------------------------------------------------
    def tally_batch(self, po_numbers, plant=None, on_row_start=None, on_row_complete=None, should_cancel=None):
        """
        Runs the full reconciliation for a batch of Production Orders --
        ONE PER SAP EXECUTION, resetting to a fresh selection screen
        between orders via StartTransaction() (the programmatic
        equivalent of typing /nZPPMYR0520, which works reliably
        regardless of whatever screen state a previous order's
        drill-down left behind). See enter_selection_single()'s
        docstring for why this replaced entering all orders at once via
        SAP's Multiple Selection popup -- that approach turned out to be
        unreliable in practice.

        on_row_start(index, total, po_number), if given, is called
        before processing that order (for "reading..." progress).
        on_row_complete(index, total, result), if given, is called
        right after that order's result is computed (for
        "updating..." progress, or writing that result out
        incrementally rather than waiting for the whole batch).
        should_cancel(), if given, is checked before each order and
        stops the loop early if it returns True -- whatever's already
        in results at that point is still returned, not discarded.

        Returns a list of result dicts, one per order that completed
        successfully (a single order's failure is logged and skipped,
        not fatal to the rest of the batch).
        """

        results = []

        for index, po_number in enumerate(po_numbers):

            if should_cancel and should_cancel():
                print(f"[INFO] Cancelled after {len(results)} of {len(po_numbers)} orders.")
                break

            po_number = str(po_number).strip()

            if on_row_start:
                on_row_start(index, len(po_numbers), po_number)

            try:
                # Reset to a clean, known selection screen before every
                # order -- see this method's docstring for why.
                self.session.StartTransaction(config.PO_TALLY_TCODE)
                self.wait_until_ready()

                self.enter_selection_single(po_number, plant)
                self.execute()

                values = self.read_tally_for_row(0)  # always row 0 -- exactly one order per execution

            except Exception as e:
                print(f"[ERROR] {po_number}: {e}")
                continue

            difference = values["target"] - values["yield_qty"] - values["scrap_qty"]
            status = self._status_for_difference(difference)

            print(
                f"[OK] {po_number} -- Target {values['target']:g}, "
                f"Yield {values['yield_qty']:g}, Scrap {values['scrap_qty']:g}, "
                f"Difference {difference:g} ({status})"
            )

            result = {
                "po": po_number,
                "target": values["target"],
                "yield_qty": values["yield_qty"],
                "scrap_qty": values["scrap_qty"],
                "difference": difference,
                "status": status,
                "posting_date": values.get("posting_date"),
            }

            results.append(result)

            if on_row_complete:
                on_row_complete(index, len(po_numbers), result)

        return results
