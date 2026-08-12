import time

from pywinauto import Application
from pywinauto import Desktop

SAP_LOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"

SAP_CONNECTION = "GERP EPA - Production"


def launch_sap():

    # ==========================================
    # Start SAP Logon (if not running)
    # ==========================================

    try:

        Application(backend="uia").connect(
            title="SAP Logon 770"
        )

        print("[INFO] SAP Logon already running.")

    except Exception:

        print("[INFO] Launch SAP Logon...")

        Application(backend="uia").start(
            SAP_LOGON_PATH
        )

    # ==========================================
    # Wait SAP Logon Window
    # ==========================================

    window = None

    for _ in range(30):

        try:

            window = Desktop(backend="uia").window(
                title="SAP Logon 770"
            )

            if window.exists():

                break

        except Exception:

            pass

        time.sleep(1)

    if window is None:

        raise Exception("Unable to detect SAP Logon window.")

    window.set_focus()

    time.sleep(2)

    print("[INFO] SAP Logon Connected.")

    # ==========================================
    # Read Connection List
    # ==========================================

    listbox = window.child_window(
        auto_id="1008",
        control_type="List"
    )

    print("[INFO] Reading SAP Connections...")

    found = False

    for item in listbox.descendants(control_type="ListItem"):

        try:

            name = item.window_text().strip()

            print(f"[FOUND] {name}")

            if name == SAP_CONNECTION:

                item.select()

                time.sleep(0.5)

                window.child_window(
                    title="Log On",
                    auto_id="1068",
                    control_type="Button"
                ).click_input()

                print("[OK] Connection Selected.")

                found = True

                break

        except Exception as e:

            print(e)

    if not found:

        raise Exception(
            f"Connection '{SAP_CONNECTION}' not found."
        )
