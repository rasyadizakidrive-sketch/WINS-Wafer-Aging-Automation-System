import time
import win32com.client


def login_sap(client, username, password, language="EN"):

    print("[INFO] Waiting Login Screen...")

    session = None

    # ==========================================
    # Wait until SAP Login Session is available
    # ==========================================

    for _ in range(30):

        try:

            SapGuiAuto = win32com.client.GetObject("SAPGUI")

            application = SapGuiAuto.GetScriptingEngine

            if application.Children.Count == 0:

                time.sleep(1)
                continue

            connection = application.Children(0)

            if connection.Children.Count == 0:

                time.sleep(1)
                continue

            session = connection.Children(0)

            break

        except:

            time.sleep(1)

    if session is None:

        raise Exception("Unable to detect SAP Login Screen.")

    print("[INFO] Filling Credentials...")

    session.findById(
        "wnd[0]/usr/txtRSYST-MANDT"
    ).text = client

    session.findById(
        "wnd[0]/usr/txtRSYST-BNAME"
    ).text = username

    session.findById(
        "wnd[0]/usr/pwdRSYST-BCODE"
    ).text = password

    session.findById(
        "wnd[0]/usr/txtRSYST-LANGU"
    ).text = language

    print("[INFO] Login SAP...")

    session.findById("wnd[0]").sendVKey(0)

    # SAP shows an error popup (wrong/expired password, locked account,
    # multiple-logon warning, etc.) as a SEPARATE window (wnd[1]) rather
    # than rejecting the input outright -- without checking for it, a
    # failed login looks identical to a successful one from here, and
    # the real reason only surfaces later as a generic, unhelpful
    # timeout several calls away from where it actually happened.
    time.sleep(1)

    has_popup = False
    popup_message = ""

    try:
        if session.Children.Count > 1:
            has_popup = True
            popup_message = session.findById("wnd[1]").Text
    except Exception:
        pass  # couldn't check -- proceed as if there's no popup rather than masking the real result

    if has_popup:
        raise Exception(f"SAP rejected the login: {popup_message}")

    print("[OK] Login Success.")