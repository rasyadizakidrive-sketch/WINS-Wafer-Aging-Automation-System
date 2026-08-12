
import pyperclip

from SAP.sap_manager import start_transaction
from config import SAP_PLANT, SAP_STORAGE

from Excel.shipment_excel import get_batch_without_shipment
from Excel.excel_manager import ExcelManager


def open_shipment():

    session = start_transaction("Shipment", "ZPPMYR0490")

    print("[INFO] Setting Plant/Storage...")

    # Plant
    session.findById(
        "wnd[0]/usr/ctxtP_WERKS"
    ).text = SAP_PLANT

    # Storage
    session.findById(
        "wnd[0]/usr/ctxtS_LGORT-LOW"
    ).text = SAP_STORAGE

    print("[OK] Shipment Screen Ready.")

    return session


def open_batch_selection(session):

    print("[INFO] Opening Batch Multiple Selection...")

    session.findById(
        "wnd[0]/usr/btn%_S_CHARG_%_APP_%-VALU_PUSH"
    ).press()

    print("[OK] Batch Multiple Selection Opened.")


def paste_batch_from_clipboard(session, batch_list):

    print("[INFO] Preparing clipboard...")

    batch_text = "\r\n".join(batch_list)

    pyperclip.copy(batch_text)

    print(f"[OK] {len(batch_list)} batch copied to clipboard.")

    # Upload From Clipboard
    session.findById(
        "wnd[1]/tbar[0]/btn[24]"
    ).press()

    # Green Tick (Close Multiple Selection)
    session.findById(
        "wnd[1]/tbar[0]/btn[8]"
    ).press()

    print("[OK] Batch list imported into SAP.")


def execute_shipment(session):

    print("[INFO] Selecting Scope of List...")

    session.findById(
        "wnd[0]/usr/radRB6"
    ).setFocus()

    session.findById(
        "wnd[0]/usr/radRB6"
    ).select()

    print("[OK] Scope Selected.")

    print("[INFO] Opening Layout Selection...")

    session.findById(
        "wnd[0]/usr/ctxtP_VARI"
    ).setFocus()

    session.findById(
        "wnd[0]/usr/ctxtP_VARI"
    ).caretPosition = 0

    session.findById("wnd[0]").sendVKey(4)

    print("[OK] Layout Window Opened.")

    grid = session.findById(
        "wnd[1]/usr/cntlGRID/shellcont/shell"
    )

    grid.currentCellRow = 1
    grid.selectedRows = "1"
    grid.clickCurrentCell()

    print("[OK] SUP_BATCH Selected.")

    print("[INFO] Executing Shipment Report...")

    session.findById(
        "wnd[0]/tbar[1]/btn[8]"
    ).press()

    print("[OK] Shipment Report Executed.")

def read_shipment_result(session):

    print("[INFO] Reading Shipment Result...")

    grid = session.findById(
        "wnd[0]/shellcont/shell"
    )

    shipment = {}

    rows = grid.RowCount

    for row in range(rows):

        batch = str(
            grid.GetCellValue(row, "CHARG")
        ).strip()

        supplier = str(
            grid.GetCellValue(row, "LICHA")
        ).strip()

        shipment[batch] = supplier

    print(f"[OK] {len(shipment)} Shipment Found.")

    return shipment



# ==========================================================
# SHIPMENT AUTOMATION
# GUI (main_window_v2) will call this function
# ==========================================================

def run_shipment_automation():
    print(">>> RUN SHIPMENT <<<")
    import config

    print("=" * 50)
    print("PRODUCT :", config.PRODUCT_NAME)
    print("SHEET   :", config.SHEET_NAME)
    print("=" * 50)

    print("\n[INFO] Checking Batch Without Shipment...")

    batches = get_batch_without_shipment()

    if not batches:

        print("[INFO] No Batch Pending Shipment.")

        return {}

    print(f"[OK] {len(batches)} Batch Pending Shipment.")

    session = open_shipment()

    open_batch_selection(session)

    paste_batch_from_clipboard(session, batches)

    execute_shipment(session)

    try:

        shipment = read_shipment_result(session)

    except Exception:

        print("[WARN] No Shipment Returned.")

        shipment = {}

    print("\n[INFO] Opening Excel...")

    excel = ExcelManager()

    excel.connect()

    excel.open_workbook()

    if shipment:

        excel.update_shipment(shipment)

    else:

        print("[INFO] Nothing to update.")

    print("\n[OK] Process Completed.")

    return shipment


# ==========================================================
# Console Entry
# ==========================================================

if __name__ == "__main__":

    run_shipment_automation()

    print("=" * 50)
    print("Shipment Automation Completed")
    print("=" * 50)
