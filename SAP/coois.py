import time

from SAP.sap_manager import start_transaction


def run_coois():

    print("[INFO] Opening COOIS...")

    session = start_transaction("COOIS", "COOIS")

    time.sleep(1)

    print("[INFO] Setting Selection Criteria...")

    session.findById(
        "wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP"
    ).key = "PPIOM000"

    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW"
    ).text = "P200"

    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_AUART-LOW"
    ).text = "ZMYW"

    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1"
    ).setFocus()

    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1"
    ).caretPosition = 0

    session.findById("wnd[0]").sendVKey(4)

    time.sleep(1)

    session.findById(
        "wnd[1]/usr/cntlCUSTOM_CONTAINER/shellcont/shell"
    ).setCurrentCell(1, "TXT30")

    session.findById(
        "wnd[1]/usr/cntlCUSTOM_CONTAINER/shellcont/shell"
    ).selectedRows = "1"

    session.findById(
        "wnd[1]/tbar[0]/btn[0]"
    ).press()

    print("[INFO] Executing COOIS...")

    session.findById(
        "wnd[0]/tbar[1]/btn[8]"
    ).press()

    time.sleep(1)

    print("[INFO] Expanding Navigation Profile...")

    grid = session.findById(
        "wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell"
    )

    grid.pressToolbarButton("&NAVIGATION_PROFILE_TOOLBAR_EXPAND")

    time.sleep(0.5)

    print("[INFO] Sorting Material Descending...")

    grid.setCurrentCell(-1, "MATNR")
    grid.selectColumn("MATNR")
    grid.pressToolbarButton("&SORT_DSC")

    print("[OK] Production Order List Loaded.")

    return True


def read_production_orders():
    """
    Reads Order / Material / Batch from the CURRENTLY OPEN COOIS
    session's ALV Grid, without executing COOIS again -- COOIS's own
    session is only read from here, never re-run or navigated away
    from, so it stays exactly as the user left it (they may still be
    looking at it).

    Raises a clear error if COOIS hasn't been run yet in this run,
    rather than silently opening (and re-executing) a fresh one -- that
    would defeat the entire point of reading what's already on screen.

    AUFNR (Order) and MATNR (Material) are standard SAP field names for
    these columns across the PP module, and COOIS itself is a standard
    transaction (not a custom one), so these are reasonably reliable --
    unlike the Z-transaction fields elsewhere in the Production Order
    Variance Check feature.

    CHARG (Batch) is read defensively since not every COOIS layout is
    guaranteed to include it as a column.
    """

    from SAP import sap_manager

    status = sap_manager.get_session_status()

    if not status.get("COOIS"):
        raise Exception(
            "COOIS is not currently open. Open Production Order List "
            "(COOIS) first, then try Production Order Variance Check again."
        )

    session, _ = sap_manager.get_dedicated_session("COOIS")

    grid = session.findById(
        "wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell"
    )

    row_count = grid.RowCount

    orders = []
    seen_pos = set()

    for row in range(row_count):

        po = grid.GetCellValue(row, "AUFNR")

        if po in (None, ""):
            continue

        po_clean = str(po).strip()

        # COOIS commonly lists one row per operation within an order,
        # not one row per order -- the same order number can legitimately
        # repeat across several rows. Without deduplicating, the
        # selection dialog would show confusing duplicate entries for
        # what's really the same Production Order.
        if po_clean in seen_pos:
            continue

        seen_pos.add(po_clean)

        try:
            material = grid.GetCellValue(row, "MATNR")
        except Exception:
            material = ""

        try:
            batch = grid.GetCellValue(row, "CHARG")
        except Exception:
            batch = ""

        orders.append({
            "po": po_clean,
            "material": str(material).strip() if material else "",
            "batch": str(batch).strip() if batch else "",
        })

    print(f"[OK] {len(orders)} Production Order(s) read from COOIS.")

    return orders