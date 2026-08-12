import time

from SAP.sap_manager import start_transaction


def run_co02(po_number):

    print("[INFO] Opening CO02...")

    session = start_transaction("CO02", "CO02")

    time.sleep(1)

    print(f"[INFO] Production Order : {po_number}")

    session.findById(
        "wnd[0]/usr/ctxtCAUFVD-AUFNR"
    ).text = str(po_number)

    session.findById("wnd[0]").sendVKey(0)

    time.sleep(1)

    print("[INFO] Reading Master Data...")

    session.findById(
        "wnd[0]/tbar[0]/btn[86]"
    ).press()

    time.sleep(0.5)

    print("[INFO] Saving Production Order...")

    session.findById(
        "wnd[0]/tbar[0]/btn[11]"
    ).press()

    time.sleep(0.5)

    print("[OK] Production Order Updated Successfully.")

    return True