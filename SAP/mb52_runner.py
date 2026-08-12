from SAP.sap_manager import start_transaction

import config


def _set_text(session, path, value, description):
    """
    Sets an SAP GUI field's .text property, wrapped so a failure names
    the EXACT field, the value being assigned, and its Python type --
    rather than letting a bare COM property-set failure surface as a
    generic "Property '<unknown>.text' can not be set" with no
    indication of which of several fields on this screen caused it.

    Deliberately re-raises rather than swallowing the error -- this
    isn't about recovering from the failure, just making whatever
    caused it immediately identifiable in the Activity Log instead of
    requiring a screenshot-and-guess cycle to narrow down.
    """
    try:
        session.findById(path).text = value
    except Exception as e:
        raise RuntimeError(
            f"Failed setting {description} ({path}) to {value!r} "
            f"(type: {type(value).__name__}) -- {e}"
        )


def _set_checkbox(session, path, description):
    try:
        session.findById(path).selected = True
    except Exception as e:
        raise RuntimeError(f"Failed selecting checkbox {description} ({path}) -- {e}")


def run_mb52():

    print(f"[INFO] Opening MB52 ({config.PRODUCT_NAME})...")

    session = start_transaction("MB52", "MB52")

    # ==========================================
    # Fill Selection
    # ==========================================

    _set_text(session, "wnd[0]/usr/ctxtMATNR-LOW", config.SAP_MATERIAL, "Material")
    _set_text(session, "wnd[0]/usr/ctxtWERKS-LOW", config.SAP_PLANT, "Plant")
    _set_text(session, "wnd[0]/usr/ctxtLGORT-LOW", config.SAP_STORAGE, "Storage Location")
    _set_text(session, "wnd[0]/usr/ctxtMATART-LOW", config.SAP_MATERIAL_TYPE, "Material Type")

    # Clear Batch
    _set_text(session, "wnd[0]/usr/ctxtCHARG-LOW", "", "Batch Low (clear)")
    _set_text(session, "wnd[0]/usr/ctxtCHARG-HIGH", "", "Batch High (clear)")

    _set_text(session, "wnd[0]/usr/ctxtP_VARI", config.SAP_LAYOUT, "Layout Variant")

    # ==========================================
    # Checkbox
    # ==========================================

    _set_checkbox(session, "wnd[0]/usr/chkXMCHB", "Display Batches")
    _set_checkbox(session, "wnd[0]/usr/chkNOZERO", "No Zero Stock")
    _set_checkbox(session, "wnd[0]/usr/chkNOVALUES", "No Values")

    try:
        session.findById("wnd[0]/usr/radPA_FLT").select()
    except Exception as e:
        raise RuntimeError(f"Failed selecting radio button radPA_FLT -- {e}")

    print(f"[INFO] Execute MB52 ({config.SAP_MATERIAL})...")

    session.findById("wnd[0]").sendVKey(8)

    print(f"[OK] MB52 Ready ({config.PRODUCT_NAME}).")
