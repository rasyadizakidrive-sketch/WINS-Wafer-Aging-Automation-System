# WINS — Wafer Intelligent Navigation System

Internal manufacturing automation platform for **wafer loading and production planning** in a semiconductor manufacturing environment (Hanwha Q CELLS Malaysia).

Replaces a 17-step manual planning cycle (8 SAP screens, ~320 clicks, 6 Excel updates, 47 minutes) with a single desktop application: **7 steps, one window, ~4.75 minutes** — a 90% cycle-time reduction.

> **Note:** This is the source code of a company-internal tool. The executive showcase (HTML) lives in the separate [`wins-showcase`](https://github.com/rasyadizakidrive-sketch/wins-showcase) repo.

## What it automates

| Module | SAP transaction | Purpose |
|--------|-----------------|---------|
| MB52 | Inventory | Batch-level stock detection, BIB matching, DOI feed |
| MB51 | Goods movements | Goods receipt date extraction |
| Shipment | ZPPMYR0490 | Supplier batch matching, GR date sync to Excel |
| CO01 (PO Creation) | Production order | Single & bulk order creation with auto-retry |
| CO02 | Change order | Master data update |
| COOIS | Order info system | Released order monitoring (cached once per run) |
| PO Variance Check | ZPPMYR0520 | Target vs Yield/Scrap reconciliation (Matched / Over / Less Posting) |
| DOI Dashboard | MB52-based | Days-of-Inventory per product with trend sparkline |
| Audit Trail | — | Filterable time/module/status log, CSV export |

## Architecture

```
GUI (CustomTkinter)
   │
Automation Controller  (orchestrates modules, retry logic, central session mgmt)
   │
SAP Layer              (7 transactions via SAP GUI Scripting)
   │
Excel Engine           (win32com live workbook, avoids file-lock conflicts)
   │
Production Planner     (daily cycle, products M6 + G12)
```

Key design points:

- **Single COOIS snapshot** cached per run and shared by all modules — no module ever sees different data
- **Connection lifecycle management** in `SAP/sap_manager.py` (alive checks, auto-reconnect, central login)
- **Structured logging** with plain-language error mapping ("SAP Session Lost", one-click reconnect)
- **Excel file-lock avoidance** via live COM workbook session

## Setup

```bash
# 1. Configure credentials
cp .env.example .env
#    fill in SAP_CLIENT, SAP_USERNAME, SAP_PASSWORD, SAP_CONNECTION, plant codes

# 2. Install dependencies (Windows, Python 3.13)
pip install -r requirements.txt

# 3. Run
python gui.py            # GUI app
python main.py           # original console automation (wafer aging flow)
```

## Building the .exe

Windows only (PyInstaller builds are platform-specific; customtkinter needs `--onedir`):

```bat
build_exe.bat
REM or: pyinstaller WINS.spec
```

Output: `dist/WINS/WINS.exe` (with `Data/`, `Assets/`, `Logs/` alongside). The built exe is distributed via internal file share / GitHub Releases — **never committed to git** (see `.gitignore`).

## Tech stack

Python 3.13 · CustomTkinter · SAP GUI Scripting (COM) · Excel COM (win32com) · OpenPyXL · PyAutoGUI · pandas · matplotlib

## License

Internal use — Hanwha Q CELLS Malaysia. Not licensed for redistribution.
