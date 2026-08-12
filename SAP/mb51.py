import pyperclip

from SAP.sap_manager import start_transaction
import config

from Excel.gr_date_excel import get_batch_without_gr_date
from Excel.excel_manager import ExcelManager


def open_mb51():

    session = start_transaction("MB51", "MB51")

    print("[INFO] Setting Selection Criteria...")

    session.findById(
        "wnd[0]/usr/ctxtMATNR-LOW"
    ).text = config.SAP_MATERIAL

    session.findById(
        "wnd[0]/usr/ctxtWERKS-LOW"
    ).text = config.SAP_PLANT

    session.findById(
        "wnd[0]/usr/ctxtBWART-LOW"
    ).text = "101"

    session.findById(
    "wnd[0]/usr/ctxtAUFNR-LOW"
    ).text = ""

    print("[OK] MB51 Ready.")

    return session

# ==========================================================
# Batch Multiple Selection
# ==========================================================

def open_batch_selection(session):

    print("[INFO] Opening Batch Multiple Selection...")

    session.findById(
        "wnd[0]/usr/btn%_CHARG_%_APP_%-VALU_PUSH"
    ).press()

    print("[OK] Batch Multiple Selection Opened.")


def paste_batch_from_clipboard(session, batch_list):

    print("[INFO] Preparing Clipboard...")

    batch_text = "\r\n".join(batch_list)

    pyperclip.copy(batch_text)

    print(f"[OK] {len(batch_list)} Batch Copied.")

    # Upload From Clipboard
    session.findById(
        "wnd[1]/tbar[0]/btn[24]"
    ).press()

    # Green Tick
    session.findById(
        "wnd[1]/tbar[0]/btn[8]"
    ).press()

    print("[OK] Batch Imported Into SAP.")

    # ==========================================================
# Execute MB51
# ==========================================================

def execute_mb51(session):

    print("[INFO] Executing MB51 Report...")

    session.findById(
        "wnd[0]/tbar[1]/btn[8]"
    ).press()

    # Tukar kepada ALV Grid
    session.findById(
        "wnd[0]/tbar[1]/btn[48]"
    ).press()

    print("[OK] MB51 Report Executed.")

    # ==========================================================
# Read GR Date
# ==========================================================

def read_gr_date(session):

    print("[INFO] Reading GR Date...")

    grid = session.findById(
        "wnd[0]/usr/cntlGRID1/shellcont/shell"
    )

    gr_date = {}

    rows = grid.RowCount

    for row in range(rows):

        batch = str(
            grid.GetCellValue(row, "CHARG")
        ).strip()

        date = grid.GetCellValue(
            row,
            "BUDAT"
        )

        gr_date[batch] = date

    print(f"[OK] {len(gr_date)} GR Date Found.")

    return gr_date

# ==========================================================
# EXPLORE SAP CONTROLS
# ==========================================================

def explore_controls(obj, level=0):

    indent = " " * level

    try:
        print(indent + obj.Id)
    except Exception:
        return

    try:

        for i in range(obj.Children.Count):

            explore_controls(
                obj.Children(i),
                level + 2
            )

    except Exception:
        pass





# ==========================================================
# MB51 AUTOMATION
# ==========================================================

def run_mb51_automation():

    print("\n[INFO] Checking Batch Without GR Date...")

    batches = get_batch_without_gr_date()

    if not batches:

        print("[INFO] No Batch Pending GR Date.")

        return {}

    print(f"[OK] {len(batches)} Batch Pending GR Date.")

    session = open_mb51()

    open_batch_selection(session)

    paste_batch_from_clipboard(session, batches)

    execute_mb51(session)

    try:

        gr_date = read_gr_date(session)

    except Exception as e:

        print(f"[ERROR] {e}")

        gr_date = {}

    print("\n[INFO] Opening Excel...")

    excel = ExcelManager()

    excel.connect()

    excel.open_workbook()

    if gr_date:

        excel.update_gr_date(gr_date)

    else:

        print("[INFO] Nothing To Update.")

    print("\n[OK] MB51 Automation Completed.")

    return gr_date

def debug_mb51(session):

    grid = session.findById(
        "wnd[0]/usr/cntlGRID1/shellcont/shell"
    )

    print("=" * 60)
    print("Row Count :", grid.RowCount)
    print("Column Count :", grid.ColumnCount)
    print("=" * 60)

    for c in range(grid.ColumnCount):

        try:
            print(c, grid.ColumnOrder(c))
        except Exception as e:
            print(c, e)

    print("=" * 60)


# ==========================================================
# Console Entry
# ==========================================================

if __name__ == "__main__":

    run_mb51_automation()

    print("=" * 50)
    print("MB51 Automation Completed")
    print("=" * 50)

