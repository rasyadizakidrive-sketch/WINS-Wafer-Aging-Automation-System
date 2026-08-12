
import os
import sys
import json
from datetime import datetime


# logger.py lives at the project root, so its own directory IS the
# project root -- anchoring here means log files always land in the
# right place regardless of what the current working directory happens
# to be when the app is launched (a desktop shortcut's "Start in"
# folder, Task Scheduler, etc. don't always match the project folder).
#
# When packaged as a PyInstaller .exe, __file__ would point somewhere
# inside the bundle's internal files instead -- sys.executable's
# directory is where the actual .exe lives, which is where Logs/
# should sit alongside it.
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

LOG_DIR = os.path.join(BASE_DIR, "Logs")


def _log_path(when=None):
    """Path to today's (or a given day's) log file, e.g. Logs/2026-07-23.log."""

    when = when or datetime.now()

    os.makedirs(LOG_DIR, exist_ok=True)

    return os.path.join(LOG_DIR, when.strftime("%Y-%m-%d.log"))


def _counts_path(when=None):
    """Path to today's small JSON counts file, e.g. Logs/2026-07-23_counts.json."""

    when = when or datetime.now()

    os.makedirs(LOG_DIR, exist_ok=True)

    return os.path.join(LOG_DIR, when.strftime("%Y-%m-%d_counts.json"))


def _write(entry):


    path = _log_path()

    with open(path, "a", encoding="utf-8") as f:
        f.write(entry)

    return path

def _bump_count(key):
    """
    Increments today's success/failed counter by exactly 1, independent of
    the human-readable .log file. The dashboard reads from here rather
    than by counting substrings in the pretty log, so it can't be thrown
    off by anything happening to that text (duplicated content, unusual
    line endings, etc.) -- this file only ever gets +1 per actual call.
    """

    path = _counts_path()

    counts = {"success": 0, "failed": 0}

    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                counts.update(json.load(f))
        except Exception:
            pass

    counts[key] = counts.get(key, 0) + 1

    with open(path, "w", encoding="utf-8") as f:
        json.dump(counts, f)

    return counts


# ==========================================================
# SUCCESS / FAILURE ENTRIES
# ==========================================================

def log_success(batch, po_number, qty):

    lines = [
        "=" * 48,
        "Batch",
        str(batch),
        "",
        "PO",
        str(po_number),
        "",
        "Qty",
        f"{qty:,}",
        "",
        "Result",
        "SUCCESS",
        "",
        "Time",
        datetime.now().strftime("%H:%M:%S"),
        "=" * 48,
        ""
    ]

    entry = "\n".join(lines)

    path = _write(entry)
    _bump_count("success")

    return path

def log_failure(batch, reason):

    lines = [
        "=" * 48,
        "Batch",
        str(batch),
        "",
        "Result",
        "FAILED",
        "",
        "Reason",
        str(reason),
        "",
        "Time",
        datetime.now().strftime("%H:%M:%S"),
        "=" * 48,
        ""
    ]

    entry = "\n".join(lines)

    path = _write(entry)
    _bump_count("failed")

    return path


# ==========================================================
# TODAY'S SUMMARY (for the dashboard counters)
# ==========================================================

def summarize_today():
    """
    Returns (success_count, failed_count) for today, read from the small
    JSON counts file rather than by counting substrings in the pretty
    .log file.
    """

    path = _counts_path()

    if not os.path.exists(path):
        return 0, 0

    try:
        with open(path, "r", encoding="utf-8") as f:
            counts = json.load(f)
        return counts.get("success", 0), counts.get("failed", 0)
    except Exception:
        return 0, 0


def today_log_path():
    """Path to today's log file, for display in the summary popup."""
    return _log_path()

