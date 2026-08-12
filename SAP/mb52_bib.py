"""
SAP/mb52_bib.py -- MB52 warehouse-stock read for the "BIB" column in
Production Order Variance Check.

Confirmed via a real VBS recording (navigating from COOIS directly
into MB52 via /nmb52), extracting the clean, successful path (the
recording also shows some exploratory Multiple Selection attempts
that got cancelled/closed -- those are not reproduced here):
  1. Clear any carried-over Material selection via context menu
     ("DELACTX"), matching the same defensive pattern used elsewhere
     in this project for select-option fields that don't reliably
     clear via a plain .text = "" assignment.
  2. Type the M6 material (11000378) into the main Material-Low field.
  3. Open Multiple Selection for Material, add the G12 material
     (11000390) as a second single value on the "Single Values" tab,
     confirm -- both materials end up selected together in one query,
     not two separate ones.
  4. Set Plant = P200, Storage Location = PHN1.
  5. Execute.

Grid path (wnd[0]/usr/cntlGRID1/shellcont/shell) matches
sap_reader.py's already-confirmed MB52 grid path exactly -- the same
underlying screen, cross-confirmed by two independent sources.

LABST (Unrestricted-use stock quantity) is already confirmed via
sap_reader.py's existing, working get_total_stock() / get_new_batches()
functions elsewhere in this project (the DOI feature) -- not a new
guess for this module.

Uses its own dedicated session key ("MB52_BIB"), deliberately NOT
"MB52" -- mb52_runner.py's existing DOI feature already owns that
session key with a different storage location (LRN1) and a
single-material selection. Sharing the same key here would risk one
feature's selection criteria silently clobbering the other's whenever
both run within the same app session.
"""

import config
from SAP.sap_manager import start_transaction, get_dedicated_session


FIELD_MATNR_LOW = "wnd[0]/usr/ctxtMATNR-LOW"
FIELD_MATNR_MULTI_BTN = "wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH"
FIELD_MULTI_SINGLE_VALUE = (
    "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/"
    "tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]"
)
FIELD_MULTI_CONFIRM_BTN = "wnd[1]/tbar[0]/btn[8]"
FIELD_WERKS_LOW = "wnd[0]/usr/ctxtWERKS-LOW"
FIELD_LGORT_LOW = "wnd[0]/usr/ctxtLGORT-LOW"
EXECUTE_BUTTON = "wnd[0]/tbar[1]/btn[8]"

GRID_RESULTS = "wnd[0]/usr/cntlGRID1/shellcont/shell"
COLUMN_MATERIAL = "MATNR"
COLUMN_BATCH = "CHARG"
COLUMN_UNRESTRICTED = "LABST"

MODULE_KEY = "MB52_BIB"

BATCH_SUFFIX_LENGTH = 4  # "similar" batches share this many trailing characters


def run_mb52_for_bib():
    """
    Opens MB52 in its own dedicated session and selects BOTH products
    (M6 + G12) together at Plant P200 / Storage Location PHN1, per the
    confirmed VBS recording. Called ONCE per Production Order Variance
    Check batch run, not once per order -- every order's BIB lookup
    reads from this same result set (see sum_unrestricted_for_batch()).
    """

    session = start_transaction(MODULE_KEY, "MB52")

    session.findById(FIELD_MATNR_LOW).showContextMenu()
    session.findById("wnd[0]/usr").selectContextMenuItem("DELACTX")

    session.findById(FIELD_MATNR_LOW).text = config.PRODUCTS["M6"]["material"]

    session.findById(FIELD_MATNR_MULTI_BTN).press()
    session.findById(FIELD_MULTI_SINGLE_VALUE).text = config.PRODUCTS["G12"]["material"]
    session.findById(FIELD_MULTI_CONFIRM_BTN).press()

    session.findById(FIELD_WERKS_LOW).text = config.SAP_PLANT
    session.findById(FIELD_LGORT_LOW).text = config.BIB_STORAGE_LOCATION

    session.findById(EXECUTE_BUTTON).press()

    print(
        f"[OK] MB52 (BIB) executed for M6 + G12 @ Plant {config.SAP_PLANT} "
        f"/ Storage {config.BIB_STORAGE_LOCATION}."
    )


def read_all_batches():
    """
    Reads every row of the MB52 results grid -- NOT filtered to a
    single material, unlike sap_reader.py's DOI reader, since BIB
    needs both products' warehouse batches together in one set to
    match against whichever order's own batch is being looked up.

    Returns a list of {"material", "batch", "unrestricted"} dicts, one
    per warehouse batch row. A row that fails to read cleanly (missing
    cell, unparseable quantity) is skipped rather than aborting the
    whole scan.
    """

    session, _ = get_dedicated_session(MODULE_KEY)
    grid = session.findById(GRID_RESULTS)

    rows = []

    for row in range(grid.RowCount):
        try:
            material = str(grid.GetCellValue(row, COLUMN_MATERIAL)).strip()
            batch = str(grid.GetCellValue(row, COLUMN_BATCH)).strip()
            raw_qty = grid.GetCellValue(row, COLUMN_UNRESTRICTED)
            unrestricted = float(str(raw_qty).replace(",", "").strip() or 0)
        except Exception:
            continue

        if not batch:
            continue

        rows.append({
            "material": material,
            "batch": batch,
            "unrestricted": unrestricted,
        })

    print(f"[OK] {len(rows)} warehouse batch row(s) read for BIB.")

    return rows


def sum_unrestricted_for_batch(warehouse_rows, target_batch, suffix_length=BATCH_SUFFIX_LENGTH):
    """
    Sums Unrestricted stock across every warehouse row whose batch
    shares the same last `suffix_length` characters as target_batch --
    "similar" batches, per the confirmed pattern: MB52 splits a single
    production batch across multiple physical carriers/cassettes with
    a varying prefix, but the trailing digits stay tied to whatever
    batch COOIS reports for that order.

    Confirmed directly from real data: COOIS batch "0001834275" (order
    16032626) and MB52 batches "1801834275" / "2501834275" all share
    the trailing "4275" -- the leading 2 digits are the only
    difference, and summing exactly those two MB52 rows' Unrestricted
    values (296 + 152 = 448) is what BIB should hold for that order.

    Returns 0.0 if target_batch is too short or nothing matches,
    rather than raising -- a missing match shouldn't break the rest of
    the tally for other orders.
    """

    target_batch = str(target_batch or "").strip()

    if len(target_batch) < suffix_length:
        return 0.0

    target_suffix = target_batch[-suffix_length:]

    total = 0.0
    for row in warehouse_rows:
        batch = row.get("batch", "")
        if len(batch) >= suffix_length and batch[-suffix_length:] == target_suffix:
            total += row.get("unrestricted", 0.0)

    return total
