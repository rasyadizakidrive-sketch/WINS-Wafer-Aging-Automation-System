"""
Excel/po_tally_excel.py -- Production Order Variance Check sheet read/write.

Independent from excel_manager.py's Aging PO sheet logic, matching the
design discussion's explicit goal for this feature to be "one module yang
independent and professional" -- its own sheet, its own column layout,
its own connect/open logic, not bolted onto the existing Aging PO
column structure.

Deliberately uses the same LIVE win32com approach as excel_manager.py,
not openpyxl reading directly from disk (which gr_date_excel.py and
shipment_excel.py do elsewhere in this project). This workbook is
frequently open live in Excel at the same time this runs -- a direct
-from-disk read risks either a stale read (missing unsaved changes) or
an outright file-lock conflict, since Excel typically locks a file it
has open. Both are real, known risks this module avoids by staying on
the live session throughout.
"""

import os
import datetime

import win32com.client

import config
import logger

xlUp = -4162


def connect():
    """Connects to the live Excel application. Matches
    excel_manager.py's exact connection pattern, reimplemented here
    (not imported from there) to keep this module genuinely
    independent."""

    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        print("[OK] Connected to existing Excel.")
    except Exception:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        print("[OK] Excel opened.")

    # Without this, an unexpected Excel dialog would sit waiting for a
    # click that never comes, since this runs unattended.
    excel.DisplayAlerts = False

    return excel


def open_workbook(excel):

    for wb in excel.Workbooks:
        if wb.Name.lower() == config.EXCEL_NAME.lower():
            return wb

    return excel.Workbooks.Open(os.path.abspath(config.EXCEL_FILE))


def get_or_create_sheet(workbook):
    """
    Returns the Production Order Variance Check worksheet, creating it
    with headers if this is the first time it has ever run against this
    workbook.
    """

    for sheet in workbook.Worksheets:
        if sheet.Name == config.PO_TALLY_SHEET_NAME:
            return sheet

    print(f"[INFO] Creating new sheet: {config.PO_TALLY_SHEET_NAME}")

    sheet = workbook.Worksheets.Add()
    sheet.Name = config.PO_TALLY_SHEET_NAME

    for col_index, header in enumerate(config.PO_TALLY_HEADERS, start=1):
        sheet.Cells(1, col_index).Value = header

    return sheet


def _get_last_row(sheet):

    return sheet.Cells(
        sheet.Rows.Count,
        config.PO_TALLY_PO_COLUMN,
    ).End(xlUp).Row


def _normalize_po(value):
    """
    Normalizes a PO number for comparison, regardless of whether Excel
    stored it as text or auto-converted it to a number. This is a real,
    well-documented Excel COM behavior: assigning a purely-numeric
    string like "16032582" to a cell often gets silently stored as a
    float, which reads back as 16032582.0 rather than "16032582" -- and
    str(16032582.0) is "16032582.0", which never matches a plain string
    comparison against the original. Without this, every existing row
    would be invisible to _find_row_for_po(), which is exactly what was
    happening: every re-run of an already-tallied PO inserted a new row
    instead of finding and updating the existing one.
    """

    if value is None:
        return ""

    if isinstance(value, float) and value == int(value):
        return str(int(value))

    return str(value).strip()


def _find_row_for_po(sheet, po_number, last_row):
    """
    Scans the PO column for an existing row. A linear scan is fine here
    -- this sheet is realistically hundreds, not tens of thousands, of
    rows, and it matches the same approach excel_manager.py already
    uses for the Aging PO sheet.
    """

    target = _normalize_po(po_number)

    for row in range(2, last_row + 1):

        value = sheet.Cells(row, config.PO_TALLY_PO_COLUMN).Value

        if value not in (None, "") and _normalize_po(value) == target:
            return row

    return None


def upsert_result(sheet, result):
    """
    Inserts a new row for this PO, or updates the existing one if it's
    already in the sheet -- "kalau dah ada, update row; kalau belum,
    insert row", per the design discussion. `result` is expected to
    already have "material" and "batch" merged in by the caller (this
    module has no knowledge of COOIS, which is where those values
    actually come from -- keeping this module's own responsibility
    limited to the Excel sheet itself). Either missing just writes as
    blank, not an error.

    "Product" is derived here, not merged in by the caller -- it's a
    pure lookup from "material" (config.get_product_for_material()),
    not a separate value from any external source, so there's nothing
    for the caller to know about or pass in.

    "posting_date" comes from the SAP read itself (zppmyr0520.py's
    detail-grid scan for the latest BUDAT), not from COOIS -- it's
    already a datetime object (or None) by the time it reaches here.

    "bib" comes from mb52_bib.py's warehouse-batch match/sum, computed
    by the caller before this is called -- defaults to 0.0 if missing
    rather than raising.
    """

    last_row = _get_last_row(sheet)

    existing_row = _find_row_for_po(sheet, result["po"], last_row)

    row = existing_row if existing_row else last_row + 1

    now_text = datetime.datetime.now().strftime("%d %b %Y %H:%M")

    po_cell = sheet.Cells(row, config.PO_TALLY_PO_COLUMN)
    # Force Text format BEFORE assigning the value -- otherwise Excel
    # can silently store a numeric-looking PO as an actual number,
    # which is the root cause fixed above on the read side. Setting
    # this on write means future rows don't have the problem at all,
    # not just work around it after the fact.
    po_cell.NumberFormat = "@"
    po_cell.Value = result["po"]

    # Batch is an identifier like PO (e.g. "0001835200"), and equally
    # vulnerable to the same silent numeric auto-conversion.
    batch_cell = sheet.Cells(row, config.PO_TALLY_BATCH_COLUMN)
    batch_cell.NumberFormat = "@"
    batch_cell.Value = result.get("batch", "")

    material = result.get("material", "")

    sheet.Cells(row, config.PO_TALLY_PRODUCT_COLUMN).Value = config.get_product_for_material(material)

    # Material codes are identifiers, not quantities, and are equally
    # vulnerable to the same silent auto-conversion (e.g. "11000378"
    # would otherwise become 11000378.0).
    material_cell = sheet.Cells(row, config.PO_TALLY_MATERIAL_COLUMN)
    material_cell.NumberFormat = "@"
    material_cell.Value = material

    sheet.Cells(row, config.PO_TALLY_TIME_COLUMN).Value = now_text

    # A real date value, not text -- so sorting, filtering, and pivot
    # tables against this column work correctly, per the planning
    # discussion.
    #
    # NOT assigning the Python datetime object directly to .Value --
    # a real debug log caught pywintypes/win32com silently treating a
    # naive datetime as local time and converting it to its UTC
    # equivalent during the COM write (confirmed: midnight Aug 6 in
    # Malaysia's UTC+8 is exactly 16:00 Aug 5 UTC, which is exactly
    # what came back reading the cell immediately after writing it).
    # Excel cells have no timezone concept at all, so that shifted
    # value is what actually gets stored -- a day early, for any
    # UTC+ timezone, every time.
    #
    # Instead, computing Excel's own date serial number directly in
    # pure Python and assigning that plain number -- no datetime
    # object crosses the COM boundary at all, so there's no timezone
    # conversion left to happen. Excel's epoch is Dec 30 1899 (this
    # exact offset correctly absorbs Excel's well-known fake 1900 leap
    # day for any real date, verified against a known reference: Dec
    # 31 2020 = serial 44196). NumberFormat still controls how this
    # number displays -- it's still a real, sortable/filterable date
    # to Excel, just reached without a datetime object.
    EXCEL_DATE_EPOCH = datetime.datetime(1899, 12, 30)

    posting_date = result.get("posting_date")
    posting_date_cell = sheet.Cells(row, config.PO_TALLY_POSTING_DATE_COLUMN)

    # Logged BEFORE the write -- the value already sitting in this
    # cell, if any (relevant if this is an update to an existing row
    # rather than a fresh insert).
    try:
        previous_value = posting_date_cell.Value
    except Exception:
        previous_value = "<unreadable>"

    posting_date_cell.NumberFormat = "dd/mm/yyyy"
    posting_date_serial = (posting_date - EXCEL_DATE_EPOCH).days if posting_date else ""
    posting_date_cell.Value = posting_date_serial

    # Read back immediately after writing -- confirms the fix, and
    # keeps this diagnostic in place in case a different date issue
    # ever shows up again.
    try:
        readback_value = posting_date_cell.Value
    except Exception:
        readback_value = "<unreadable>"

    try:
        debug_path = os.path.join(logger.LOG_DIR, "posting_date_debug.log")
        os.makedirs(logger.LOG_DIR, exist_ok=True)
        with open(debug_path, "a", encoding="utf-8") as f:
            f.write(
                f"\n=== Excel write @ {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n"
                f"PO: {result.get('po')}\n"
                f"Row: {row} ({'UPDATE existing' if existing_row else 'INSERT new'})\n"
                f"Value in cell BEFORE this write: {previous_value!r}\n"
                f"posting_date received from SAP-side result dict: {posting_date!r}\n"
                f"Serial number written (Excel date, no datetime/timezone involved): {posting_date_serial!r}\n"
                f"Value ACTUALLY IN the cell immediately AFTER writing + reading back: {readback_value!r}\n"
            )
    except Exception:
        pass  # diagnostic logging itself must never break the actual write

    sheet.Cells(row, config.PO_TALLY_TARGET_COLUMN).Value = result["target"]
    sheet.Cells(row, config.PO_TALLY_YIELD_COLUMN).Value = result["yield_qty"]
    sheet.Cells(row, config.PO_TALLY_SCRAP_COLUMN).Value = result["scrap_qty"]
    sheet.Cells(row, config.PO_TALLY_DIFFERENCE_COLUMN).Value = result["difference"]
    sheet.Cells(row, config.PO_TALLY_BIB_COLUMN).Value = result.get("bib", 0.0)
    sheet.Cells(row, config.PO_TALLY_STATUS_COLUMN).Value = result["status"]

    action = "Updated" if existing_row else "Inserted"

    print(f"[OK] {action} row {row} for PO {result['po']} ({result['status']}).")

    return row


def update_po_tally(results):
    """
    High-level entry point: connects to the live Excel session, opens
    (or creates) the sheet, upserts every result, saves once
    at the end -- not once per row, so a batch of many POs doesn't
    trigger a slow, repeated disk write for each one.
    """

    if not results:
        print("[INFO] No variance check results to write.")
        return

    excel = connect()
    workbook = open_workbook(excel)
    sheet = get_or_create_sheet(workbook)

    for result in results:
        upsert_result(sheet, result)

    workbook.Save()

    print(f"[OK] Variance check sheet saved -- {len(results)} order(s) processed.")
