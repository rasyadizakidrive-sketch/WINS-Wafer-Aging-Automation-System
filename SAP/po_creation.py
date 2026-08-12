import time
import config
import re

from SAP.sap_manager import start_transaction

class ProductionOrder:

    def __init__(self):

        self.session = None

    # --------------------------------------------------
    # OPEN CO01
    # --------------------------------------------------
    def open_co01(self):

        print("[INFO] Opening CO01...")

        self.session = start_transaction("CO01", "CO01")

        print("[OK] CO01 Ready.")

    # --------------------------------------------------
    # HEADER
    # --------------------------------------------------
    def enter_header(self, material, plant):

        print("[INFO] Filling Header...")

        order_type = "ZMYW"

        self.session.findById("wnd[0]/usr/ctxtCAUFVD-MATNR").text = material
        self.session.findById("wnd[0]/usr/ctxtCAUFVD-WERKS").text = plant
        self.session.findById("wnd[0]/usr/ctxtAUFPAR-PP_AUFART").text = order_type

        fld = self.session.findById("wnd[0]/usr/ctxtAUFPAR-PP_AUFART")
        fld.setFocus()
        fld.caretPosition = 4

        self.session.findById("wnd[0]").sendVKey(0)

        # Submitting the header transitions to a completely different
        # screen (the KOZE tab with quantity/dates) -- without waiting
        # for that to finish, the very next call (enter_quantity) would
        # try to find a control on a screen that hasn't finished
        # rendering yet, which is exactly what "The control could not
        # be found by id" means. Checking for a popup here also means
        # an invalid material/plant combination surfaces as a clear
        # message rather than that same cryptic error one line later.
        self.wait_until_ready()

        popup_message = self.check_popup()
        if popup_message:
            raise Exception(f"SAP rejected the header: {popup_message}")

        print("[OK] Header Completed.")

    # --------------------------------------------------
    # QUANTITY
    # --------------------------------------------------
    def enter_quantity(self, qty):

        self.session.findById(
            "wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/txtCAUFVD-GAMNG"
        ).text = f"{qty:,}"

    # --------------------------------------------------
    # START DATE
    # --------------------------------------------------
    def enter_start_date(self, date):

        field = self.session.findById(
            "wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/ctxtCAUFVD-GSTRP"
        )

        field.text = date

    # --------------------------------------------------
    # FINISH DATE
    # --------------------------------------------------
    def enter_finish_date(self, date):

        field = self.session.findById(
            "wnd[0]/usr/tabsTABSTRIP_0115/tabpKOZE/ssubSUBSCR_0115:SAPLCOKO1:0120/ctxtCAUFVD-GLTRP"
        )

        field.text = date

    # --------------------------------------------------
    # BATCH
    # --------------------------------------------------
    def enter_batch(self, batch):

        self.session.findById(
            "wnd[0]/tbar[1]/btn[6]"
        ).press()

        # This button press opens the batch/component allocation
        # sub-screen -- same gap as enter_header() above: without
        # waiting for it to render, the findById() right below can fail
        # with "control could not be found by id" if SAP hasn't
        # finished building it yet.
        self.wait_until_ready()

        popup_message = self.check_popup()
        if popup_message:
            raise Exception(f"SAP rejected the batch entry: {popup_message}")

        self.session.findById(
            "wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-CHARG[12,0]"
        ).text = batch

    # --------------------------------------------------
    # SAVE
    # --------------------------------------------------
    def save(self):

        print("[INFO] Saving Production Order...")

        self.session.findById(
            "wnd[0]/tbar[0]/btn[11]"
        ).press()

        print("[OK] Save Completed.")


    # --------------------------------------------------
    # GET PO NUMBER
    # --------------------------------------------------
    def get_po_number(self):

        status = self.get_status_bar()

        m = re.search(r'(\d{8})', status)

        if m:
            po = m.group(1)
            print(f"[OK] Production Order : {po}")
            return po

        raise Exception(f"Cannot read Production Order.\nStatus : {status}")        


    # --------------------------------------------------
    # WAIT UNTIL SAP READY
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
    # STATUS BAR
    # --------------------------------------------------
    def get_status_bar(self):

        try:
            return self.session.findById("wnd[0]/sbar").text.strip()
        except Exception:
            return ""

    # --------------------------------------------------
    # CHECK POPUP
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
    # CREATE PRODUCTION ORDER
    # --------------------------------------------------
    def create_production_order(self, data):

        print("=" * 60)
        print("CREATE PRODUCTION ORDER")
        print("=" * 60)

        self.open_co01()
        self.wait_until_ready()

        self.enter_header(
            data["material"],
            data["plant"]
        )

        self.enter_quantity(
            data["qty"]
        )

        self.enter_start_date(
            data["start_date"]
        )

        self.enter_finish_date(
            data["finish_date"]
        )

        self.enter_batch(
            data["batch"]
        )

        self.save()

        self.wait_until_ready()

        status = self.get_status_bar()

        if status:
            print(f"[SAP] {status}")

        self.check_popup()

        po_number = self.get_po_number()

        if not po_number:
            raise Exception("Production Order was not created.")

        return po_number