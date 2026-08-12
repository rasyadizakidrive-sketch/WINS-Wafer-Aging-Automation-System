# ===========================================
# PROJECT PATHS -- portable, not tied to any
# specific machine, username, or drive letter
# ===========================================

import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    # Running as a PyInstaller-built .exe -- __file__ would point
    # somewhere inside the bundle's internal files, not next to the
    # .exe itself. sys.executable is the actual .exe's location, which
    # (in --onedir mode) is where Data/Assets/Logs live alongside it.
    PROJECT_ROOT = Path(sys.executable).resolve().parent
else:
    PROJECT_ROOT = Path(__file__).resolve().parent

DATA_FOLDER = PROJECT_ROOT / "Data"
ASSETS_FOLDER = PROJECT_ROOT / "Assets"
LOGS_FOLDER = PROJECT_ROOT / "Logs"

# ===========================================
# EXCEL CONFIG
# ===========================================

EXCEL_NAME = "2026 Wafer Loading Plan.xlsx"

EXCEL_FILE = str(DATA_FOLDER / EXCEL_NAME)

# ===========================================
# PRODUCT CONFIG
# ===========================================

PRODUCTS = {

    "M6": {

        "material": "11000378",
        "sheet": "Aging_PO_M6",
        "description": "Wafer/P/M6/L.G/150/G/A/MY/MY/3S-CF/MY"

    },

    "G12": {

        "material": "11000390",
        "sheet": "Aging_PO_G12",
        "description": "Wafer/N/G12/T.R/130/G/A/MY/VT/3S-CF/VT"

    }

}

# ===========================================
# ACTIVE PRODUCT
# ===========================================

CURRENT_PRODUCT = "M6"

CURRENT = None
PRODUCT_NAME = None
TARGET_MATERIAL = None
SAP_MATERIAL = None
SHEET_NAME = None
PRODUCT_DESCRIPTION = None


def set_product(product):

    """
    Change active product dynamically.
    Used by GUI later.
    """

    global CURRENT_PRODUCT
    global CURRENT
    global PRODUCT_NAME
    global TARGET_MATERIAL
    global SAP_MATERIAL
    global SHEET_NAME
    global PRODUCT_DESCRIPTION

    if product not in PRODUCTS:

        raise ValueError(f"Unknown product : {product}")

    CURRENT_PRODUCT = product

    CURRENT = PRODUCTS[product]

    PRODUCT_NAME = product

    TARGET_MATERIAL = CURRENT["material"]

    SAP_MATERIAL = CURRENT["material"]

    SHEET_NAME = CURRENT["sheet"]

    PRODUCT_DESCRIPTION = CURRENT["description"]


# Initialize default product
set_product(CURRENT_PRODUCT)

# ===========================================
# EXCEL COLUMN
# ===========================================

BATCH_COLUMN = 3
DESCRIPTION_COLUMN = 4
QTY_COLUMN = 5
SHIPMENT_COLUMN = 6
PO_COLUMN = 2      # Column B
START_DATE_COLUMN = 1
STATUS_COLUMN = 9      # Column I
GR_DATE_COLUMN = 7      # Column G

# ===========================================
# SAP MB52 CONFIG
#
# Plant, storage locations, and layout variant are specific to this
# company's own SAP landscape -- loaded from .env, never hardcoded
# here. See .env.example for the expected keys.
# ===========================================

from dotenv import load_dotenv
import os

# Explicit path, not a bare load_dotenv() relying on the current
# working directory -- for a packaged .exe, CWD isn't guaranteed to be
# the app's own folder in every launch scenario (e.g. a desktop
# shortcut with a different "Start in" directory). Same reasoning
# PROJECT_ROOT itself already exists for above.
load_dotenv(PROJECT_ROOT / ".env")


def _require_env(key):
    """
    Reads a required environment variable, failing fast with a clear
    message rather than silently falling back to a blank or wrong
    value -- a missing SAP credential should surface immediately at
    startup, not as a confusing login failure three steps into a run.
    """
    value = os.getenv(key)
    if not value:
        raise RuntimeError(
            f"Missing required environment variable: {key}. "
            f"Copy .env.example to .env and fill in your own SAP details."
        )
    return value


SAP_PLANT = _require_env("SAP_PLANT")

SAP_STORAGE = _require_env("SAP_STORAGE")

# Storage location for the BIB column's MB52 query (Production Order
# Variance Check) -- distinct from SAP_STORAGE above, which belongs to
# the unrelated DOI feature. Kept as its own key so the two features
# can never accidentally share or clobber each other's selection
# criteria.
BIB_STORAGE_LOCATION = _require_env("SAP_BIB_STORAGE_LOCATION")

SAP_MATERIAL_TYPE = "ZROH"  # standard SAP material type (Raw Material) -- not company-specific

SAP_LAYOUT = _require_env("SAP_LAYOUT")

# ===========================================
# SAP LOGIN
#
# Credentials and connection name -- all from .env. These must never
# be hardcoded or committed, regardless of how convenient it seems
# during development.
#
# SAP_USERNAME and SAP_PASSWORD deliberately do NOT check for
# placeholder text here, unlike a stricter version considered earlier
# -- this .exe is built once and distributed to multiple PCA users,
# each filling in their own login afterward. A hard check here would
# incorrectly block building the shared .exe itself, since the
# template values are intentionally still placeholders at that point.
# ===========================================

SAP_CLIENT = _require_env("SAP_CLIENT")

SAP_USERNAME = _require_env("SAP_USERNAME")

SAP_PASSWORD = _require_env("SAP_PASSWORD")

SAP_LANGUAGE = "EN"  # not sensitive -- standard SAP language code

# SAP Logon's saved connection name (as it appears in the SAP Logon
# picker) -- consolidated here from sap_launcher.py, which previously
# hardcoded this directly. Company-specific, so it comes from .env
# too, not a plain constant.
SAP_CONNECTION = _require_env("SAP_CONNECTION")

# Standard install location for SAP GUI on Windows -- genuinely
# generic (not company-specific), so this stays a plain default rather
# than requiring .env. Override via SAP_LOGON_PATH in .env only if
# your own install lives somewhere non-standard.
SAP_LOGON_PATH = os.getenv(
    "SAP_LOGON_PATH",
    r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe",
)

# Production Order Material
PO_MATERIAL = {
    "M6": "CF02-0119",
    "G12": "CF03-0049",
}

MATERIAL_OPTIONS = {
    "M6": {
        "M6 Longi": "CF02-0119",
        "M6 Hemlock": "CF02-0125",
    },
    "G12": {
        "G12 Trina": "CF03-0049",
    },
}

# ===========================================
# PO TALLY
# ===========================================
# Reconciliation module: for each selected Production Order, reads
# actual posted Target/Yield/Scrap from ZPPMYR0520 and compares against
# plan. Independent of the Aging PO sheet/columns above -- its own
# sheet, own column layout, own module.

PO_TALLY_TCODE = "ZPPMYR0520"

PO_TALLY_SHEET_NAME = "PO Variance Check"

# 1-indexed, matching this file's existing column convention.
# | Check Time | Posting Date | PO | Batch | Product | Material | Target Qty | Yield Qty | Scrap Qty | Difference | BIB | Status |
#
# Posting Date is inserted right after Check Time (Column B), pushing
# every other column one position further right than the previous
# layout -- the same kind of deliberate insertion used for Batch and
# Product earlier, not an append. Any rows already written under the
# previous layout will have their old values sitting under these new
# column letters/headers -- this only affects data already in the
# sheet, not anything going forward, but it's worth knowing before
# reading old rows against the new header row.
#
# BIB is inserted at column K, pushing Status to L -- same insertion
# pattern again. BIB's own value (summed Unrestricted stock from MB52
# across the "last 4 similar" batches) is not written by upsert_result
# yet -- the column exists so the layout is ready, but the read/sum
# logic behind it is pending confirmation of what "similar" means for
# batch grouping and the exact MB52 field name for Unrestricted stock.
PO_TALLY_TIME_COLUMN = 1
PO_TALLY_POSTING_DATE_COLUMN = 2
PO_TALLY_PO_COLUMN = 3
PO_TALLY_BATCH_COLUMN = 4
PO_TALLY_PRODUCT_COLUMN = 5
PO_TALLY_MATERIAL_COLUMN = 6
PO_TALLY_TARGET_COLUMN = 7
PO_TALLY_YIELD_COLUMN = 8
PO_TALLY_SCRAP_COLUMN = 9
PO_TALLY_DIFFERENCE_COLUMN = 10
PO_TALLY_BIB_COLUMN = 11
PO_TALLY_STATUS_COLUMN = 12

PO_TALLY_HEADERS = (
    "Check Time", "Posting Date", "PO", "Batch", "Product", "Material",
    "Target Qty", "Yield Qty", "Scrap Qty", "Difference", "BIB", "Status",
)

PO_TALLY_STATUS_MATCH = "Matched"
PO_TALLY_STATUS_LESS = "Less Posting"
PO_TALLY_STATUS_OVER = "Over Posting"


def get_product_for_material(material_code):
    """
    Reverse-looks-up PRODUCTS (M6 -> 11000378, G12 -> 11000390, defined
    once above) to turn a raw material code back into its product
    name, e.g. for the PO Variance Check sheet's "Product" column.
    Reuses the existing PRODUCTS mapping rather than a second, separate
    one -- adding a new product there in the future means this picks
    it up automatically, with nothing else to remember to update.
    Returns the material code itself if it doesn't match any known
    product, rather than an empty string, so an unrecognized code is
    still visible in the sheet rather than silently disappearing.
    """

    material_code = str(material_code).strip()

    for product_name, info in PRODUCTS.items():
        if info.get("material") == material_code:
            return product_name

    return material_code