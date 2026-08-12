import config

from SAP.sap_manager import get_dedicated_session


def get_new_batches():

    # Reads from the SAME session mb52_runner.run_mb52() just used --
    # not just "whatever session happens to be first", which stopped
    # being a safe assumption once each module got its own dedicated
    # session. Passing "MB52" here must match the module_key mb52_runner
    # used when it called start_transaction().
    session, _ = get_dedicated_session("MB52")

    try:
        grid = session.findById(
            "wnd[0]/usr/cntlGRID1/shellcont/shell"
        )
        print("[OK] GRID1 Found")

    except Exception as e:
        print("[ERROR] GRID1 Not Found")
        print(e)

        raise

    sap_data = []

    for row in range(grid.RowCount):

        material = str(
            grid.GetCellValue(row, "MATNR")
        ).strip()

        # ==========================================
        # Filter Selected Product
        # ==========================================

        if material != config.TARGET_MATERIAL:

            continue

        sap_data.append({

            "material": material,

            "batch": str(
                grid.GetCellValue(row, "CHARG")
            ).strip(),

            "description": str(
                grid.GetCellValue(row, "MAKTX")
            ).strip(),

            "qty": str(
                grid.GetCellValue(row, "LABST")
            ).strip()

        })

    # ==========================================
    # Sort Batch Number
    # ==========================================

    sap_data.sort(
        key=lambda x: x["batch"]
    )

    print(
        f"[OK] {len(sap_data)} {config.PRODUCT_NAME} Batch Retrieved."
    )

    return sap_data

# ==========================================
# GET TOTAL STOCK
# ==========================================

def get_total_stock():

    # Same reasoning as get_new_batches() above: this must read from
    # the dedicated "MB52" session (the one doi.py's calculate_doi()
    # just ran MB52 in via run_mb52()), not an assumed first session.
    session, _ = get_dedicated_session("MB52")

    grid = session.findById(
        "wnd[0]/usr/cntlGRID1/shellcont/shell"
    )

    total_qty = 0

    for row in range(grid.RowCount):

        material = str(
            grid.GetCellValue(row, "MATNR")
        ).strip()

        if material != config.TARGET_MATERIAL:
            continue

        qty = grid.GetCellValue(
            row,
            "LABST"
        )

        if qty in (None, ""):
            continue

        try:
            qty = float(
                str(qty).replace(",", "")
            )
        except (ValueError, TypeError):
            qty = 0

        total_qty += qty

    print(f"[OK] Total Stock : {total_qty:,.0f} pcs")

    return total_qty
