from SAP.mb52_runner import run_mb52
from SAP.sap_reader import get_new_batches
from Excel.excel_manager import ExcelManager

import config


def print_banner():

    print("=" * 60)
    print("Wafer Aging Automation v1.0".center(60))
    print("=" * 60)
    print()
    print(f"Product        : {config.PRODUCT_NAME}")
    print(f"Material       : {config.SAP_MATERIAL}")
    print(f"Worksheet      : {config.SHEET_NAME}")
    print()
    print("=" * 60)


def compare_batches(sap_data, excel_batches):

    return [
        item for item in sap_data
        if item["batch"] not in excel_batches
    ]


def print_new_batches(new_batches):

    print()
    print("=" * 100)
    print(f"NEW {config.PRODUCT_NAME} BATCH")
    print("=" * 100)

    if not new_batches:

        print("[INFO] No New Batch Found.")

    else:

        print(f'{"Batch":<15}{"Qty":>15}    Description')
        print("-" * 100)

        for item in new_batches:

            print(
                f'{item["batch"]:<15}'
                f'{item["qty"]:>15}    '
                f'{item["description"]}'
            )

    print()
    print("=" * 100)
    print(f"Total New {config.PRODUCT_NAME} Batch : {len(new_batches)}")
    print("=" * 100)


def update_excel(excel, new_batches):

    if not new_batches:
        return

    ans = input(
        f"\nUpdate {config.PRODUCT_NAME} Excel? (Y/N): "
    )

    if ans.upper() == "Y":

        excel.write_batches(new_batches)

    else:

        print("[INFO] Update Cancelled.")


# ==========================================================
# MAIN AUTOMATION
# GUI akan panggil function ni
# ==========================================================

def run_automation():

    print("\n[INFO] Reading SAP...")

    run_mb52()

    sap_data = get_new_batches()

    print("\n[INFO] Opening Excel...")

    excel = ExcelManager()

    excel.connect()

    excel.open_workbook()

    excel_batches = excel.get_existing_batches()

    print("\n[INFO] Comparing Batch...")

    new_batches = compare_batches(
        sap_data,
        excel_batches
    )

    print_new_batches(new_batches)

    #update_excel(
     #   excel,
      #  new_batches
    #)

    print("\n[OK] Process Completed.")

    return new_batches


# ==========================================================
# Console Entry
# ==========================================================

def main():

    print_banner()

    run_automation()


if __name__ == "__main__":

    main()