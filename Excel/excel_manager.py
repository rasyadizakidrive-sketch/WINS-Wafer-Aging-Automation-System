import os
import win32com.client

import config

xlUp = -4162


class ExcelManager:

    def __init__(self):

        self.excel = None
        self.workbook = None
        self.sheet = None

    # ======================================================
    # CONNECT EXCEL
    # ======================================================

    def connect(self):

        try:

            self.excel = win32com.client.GetActiveObject(
                "Excel.Application"
            )

            print("[OK] Connected to existing Excel.")

        except Exception:

            self.excel = win32com.client.Dispatch(
                "Excel.Application"
            )

            self.excel.Visible = True

            print("[OK] Excel opened.")

        # Without this, an unexpected Excel dialog (file already open
        # elsewhere, a formatting/compatibility warning on save, etc.)
        # would sit waiting for a click that never comes, since this
        # runs unattended -- silently hanging the whole automation
        # rather than raising an error that can actually be surfaced.
        self.excel.DisplayAlerts = False

    # ======================================================
    # OPEN WORKBOOK
    # ======================================================

    def open_workbook(self):

        for wb in self.excel.Workbooks:

            if wb.Name.lower() == config.EXCEL_NAME.lower():

                self.workbook = wb
                break

        if self.workbook is None:

            self.workbook = self.excel.Workbooks.Open(
                os.path.abspath(config.EXCEL_FILE)
            )

        print(
            f"[INFO] Opening Worksheet : {config.SHEET_NAME}"
        )

        self.sheet = self.workbook.Worksheets(
            config.SHEET_NAME
        )

        print(
            f"[OK] Worksheet : {config.SHEET_NAME}"
        )

    # ======================================================
    # LAST USED ROW
    # ======================================================

    def get_last_row(self):

        return self.sheet.Cells(
            self.sheet.Rows.Count,
            config.BATCH_COLUMN
        ).End(xlUp).Row

    # ======================================================
    # READ EXISTING BATCH
    # ======================================================

    def get_existing_batches(self):

        batches = set()

        last_row = self.get_last_row()

        for row in range(2, last_row + 1):

            value = self.sheet.Cells(
                row,
                config.BATCH_COLUMN
            ).Value

            if value not in [None, ""]:

                batches.add(str(value).strip())

        return batches

    # ======================================================
    # WRITE NEW BATCH
    # ======================================================

    def write_batches(self, new_batches):

        if not new_batches:

            print("[INFO] No New Batch.")

            return

        row = self.get_last_row() + 1

        print()
        print(
            f"[INFO] Writing {config.PRODUCT_NAME} Excel..."
        )

        for item in new_batches:

            self.sheet.Cells(
                row,
                config.BATCH_COLUMN
            ).Value = item["batch"]

            self.sheet.Cells(
                row,
                config.DESCRIPTION_COLUMN
            ).Value = item["description"]

            self.sheet.Cells(
                row,
                config.QTY_COLUMN
            ).Value = item["qty"]

            # Default status for new batch
            self.sheet.Cells(
                row,
                config.STATUS_COLUMN
            ).Value = "Yet Firmed"

            print(f"✓ {item['batch']}")

            row += 1

        self.workbook.Save()

        print()

        print(
            f"[OK] {config.PRODUCT_NAME} Excel Saved."
        )

    # ======================================================
    # UPDATE SHIPMENT
    # ======================================================

    def update_shipment(self, shipment_dict):

        print()
        print("[INFO] Updating Shipment...")

        last_row = self.get_last_row()

        updated = 0

        for row in range(2, last_row + 1):

            batch = self.sheet.Cells(
                row,
                config.BATCH_COLUMN
            ).Value

            if batch in [None, ""]:
                continue

            batch = str(batch).strip()

            if batch in shipment_dict:

                self.sheet.Cells(
                    row,
                    config.SHIPMENT_COLUMN
                ).Value = shipment_dict[batch]

                updated += 1

                print(
                    f"✓ {batch} -> {shipment_dict[batch]}"
                )

        self.workbook.Save()

        print()

        print(f"[OK] Updated {updated} Shipment.")


    # ======================================================
    # UPDATE GR DATE
    # ======================================================

    def update_gr_date(self, gr_date_dict):

        print()
        print("[INFO] Updating GR Date...")

        last_row = self.get_last_row()

        updated = 0

        for row in range(2, last_row + 1):

            batch = self.sheet.Cells(
                row,
                config.BATCH_COLUMN
            ).Value

            if batch in [None, ""]:
                continue

            batch = str(batch).strip()

            if batch in gr_date_dict:

                self.sheet.Cells(
                    row,
                    config.GR_DATE_COLUMN
                ).Value = gr_date_dict[batch]

                updated += 1

                print(
                    f"✓ {batch} -> {gr_date_dict[batch]}"
                )

        self.workbook.Save()

        print()
        print(f"[OK] Updated {updated} GR Date.")

    def get_po_batches(self):

        sheet = self.sheet

        batches = []

        last_row = self.get_last_row()

        for row in range(3, last_row + 1):

            batch = sheet.Cells(row, config.BATCH_COLUMN).Value

            if batch in (None, ""):
                continue

            po = sheet.Cells(row, config.PO_COLUMN).Value
            material = sheet.Cells(row, config.DESCRIPTION_COLUMN).Value
            qty = sheet.Cells(row, config.QTY_COLUMN).Value
            shipment = sheet.Cells(row, config.SHIPMENT_COLUMN).Value
            status = sheet.Cells(row, config.STATUS_COLUMN).Value

            # Skip batch that already has Production Order
            if po not in (None, ""):
                continue

            # Only process Yet Firmed batch
            if str(status).strip().lower() != "yet firmed":
                continue

            try:
                qty = int(float(qty))
            except (TypeError, ValueError):
                qty = 0

            batches.append(
                {
                    "batch": str(batch).strip(),
                    "po": "",
                    "material": "" if material is None else str(material).strip(),
                    "qty": qty,
                    "shipment": "" if shipment is None else str(shipment).strip(),
                    "row": row,
                }
            )

        print(f"[OK] {len(batches)} Batch Pending Production Order.")

        return batches    

    # --------------------------------------------------
    # UPDATE PO NUMBER (+ START DATE)
    # --------------------------------------------------
    def update_po(self, row, po_number, start_date=None):

        print()
        print("[INFO] Updating Production Order...")

        if not po_number:
            raise ValueError("Production Order is empty.")

        if start_date:
            self.sheet.Cells(row, config.START_DATE_COLUMN).Value = start_date

        self.sheet.Cells(
            row,
            config.PO_COLUMN
        ).Value = str(po_number).strip()

        # Update Status
        self.sheet.Cells(
            row,
            config.STATUS_COLUMN
        ).Value = "Schedule"

        self.workbook.Save()

        print(f"[OK] Production Order saved to Excel.")
        print(f"[OK] Row {row} -> {po_number}")

        return True