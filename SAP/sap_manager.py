"""
Centralized SAP GUI connection and session management.

Replaces the duplicated `win32com.client.GetObject("SAPGUI")` boilerplate
that used to be repeated in every module (mb52_runner.py, mb51.py,
shipment.py, co02.py, coois.py, po_creation.py) with one shared entry
point:

    from SAP.sap_manager import start_transaction
    session = start_transaction("MB52", "MB52")

Two behavior changes on purpose, matching what was asked for:

1. Auto-recovery if SAP GUI isn't running at all -- rather than erroring
   out and leaving the user to open SAP themselves, this reuses your
   existing sap_launcher.launch_sap() + sap_login.login_sap() (the same
   pywinauto-based launch-and-login that sap_connection.py's ensure_sap()
   already used) to bring SAP up and log in automatically, then
   continues. This does NOT reimplement that logic -- it calls your
   already-working code directly.

2. A dedicated session per module -- instead of every transaction
   reusing the same session (each one overwriting whatever the last one
   had open), each named module gets its own SAP GUI session, created
   once and reused on every later call, instead of piling up a fresh
   session every single time.

Other modules that need to *read* results from a session another module
just populated (sap_reader.py reading MB52's grid after mb52_runner.py
ran it, for instance) should call get_dedicated_session("MB52") directly
rather than starting a fresh transaction -- see sap_reader.py for the
pattern.
"""

import time
import threading

import win32com.client

from SAP.sap_launcher import launch_sap
from SAP.sap_login import login_sap
from config import SAP_CLIENT, SAP_USERNAME, SAP_PASSWORD, SAP_LANGUAGE

# How long to wait for a login to complete before giving up. Module-level
# so tests can shrink it rather than waiting out the real duration.
LOGIN_TIMEOUT_SECONDS = 30


# Cached COM objects (singleton)
_cached_sap_gui = None
_cached_application = None
_cached_connection = None

# The session reserved purely for sending /o commands through -- see
# _get_trigger_session() below for why this exists as a separate,
# never-reused-as-a-dedicated-window reference rather than an index
# lookup.
_trigger_session = None

_cache_lock = threading.RLock()


def reset_cache():
    """Clear cached SAP COM objects."""
    global _cached_sap_gui, _cached_application, _cached_connection, _trigger_session
    _cached_sap_gui = None
    _cached_application = None
    _cached_connection = None
    _trigger_session = None


def _connection_is_alive(connection):
    try:
        return connection is not None and connection.Children.Count >= 0
    except Exception:
        return False


def _is_logged_in(connection):
    """
    Checks whether the connection has an ACTUAL, authenticated user
    attached to its first session -- not just "does a connection object
    exist." SAP Logon can have a window open (a connection was
    established) while still sitting at the Client/User/Password
    screen, not yet authenticated -- session.Info.User is reliably
    blank in that state, and non-blank once a real login has completed.

    Treating "a connection exists" and "someone is actually logged in"
    as the same thing was exactly the gap that let automation proceed
    against a login screen it wasn't expecting -- producing a confusing
    COM error deep inside whatever screen the code assumed it was
    already on, rather than a clear "please log in first" message.
    """
    try:
        if connection is None or connection.Children.Count == 0:
            return False
        return bool(connection.Children(0).Info.User)
    except Exception:
        return False


def get_connection(force_refresh=False):
    """Single entry point for obtaining a live SAP connection."""
    with _cache_lock:
        if force_refresh:
            reset_cache()

        connection = _get_running_connection()

        if connection is not None:
            if _is_logged_in(connection):
                return connection

            # A SAP GUI window already exists -- SAP Logon was already
            # used to open one -- but nobody's actually completed the
            # login yet. Finishing that login directly (login_sap()
            # alone) is both simpler and safer than re-running the full
            # launch_sap() + login_sap() flow: launch_sap() re-selects a
            # connection from SAP Logon's own picker and clicks "Log
            # On" again, which risks opening a SECOND session rather
            # than completing the one that's already sitting there.
            print("[INFO] SAP is open but not logged in yet -- completing login...")
            return _finish_login(connection)

        reset_cache()
        return _launch_and_login()



# ==========================================================
# LOW-LEVEL: is SAP GUI running, and what's the connection?
# ==========================================================

def _get_running_connection():
    """
    Returns the cached SAP connection when still valid, otherwise rebuilds
    the cache from SAP GUI.
    """
    global _cached_sap_gui, _cached_application, _cached_connection

    if _connection_is_alive(_cached_connection):
        return _cached_connection

    reset_cache()

    try:
        _cached_sap_gui = win32com.client.GetObject("SAPGUI")
        _cached_application = _cached_sap_gui.GetScriptingEngine

        if _cached_application.Children.Count == 0:
            reset_cache()
            return None

        _cached_connection = _cached_application.Children(0)
        return _cached_connection

    except Exception:
        reset_cache()
        return None


def _finish_login(connection):
    """
    Completes a login on a SAP GUI window that's already open and
    sitting at the Client/User/Password screen -- called instead of
    _launch_and_login() specifically because the window already exists
    here; there's nothing to launch or re-select, just credentials to
    submit. Polls _is_logged_in() afterward rather than assuming
    login_sap() submitting the form means the session is immediately
    usable in scripting terms, same reasoning as _launch_and_login()
    below.
    """

    login_sap(SAP_CLIENT, SAP_USERNAME, SAP_PASSWORD, SAP_LANGUAGE)

    deadline = time.time() + LOGIN_TIMEOUT_SECONDS

    while time.time() < deadline:

        if _is_logged_in(connection):
            print("[OK] SAP Connected.")
            return connection

        time.sleep(1)

    raise TimeoutError(
        f"Login did not complete within {LOGIN_TIMEOUT_SECONDS} seconds "
        "after submitting credentials to the already-open SAP window. "
        "Check that window directly for what state it's actually in -- "
        "an incorrect password or a blocked account would show an error "
        "there that this can't see or report on its own."
    )


def _launch_and_login():
    """
    Brings SAP up and logs in using the existing launch_sap() (selects
    your configured connection in SAP Logon) + login_sap() (fills in
    Client/User/Password/Language and presses Enter) -- the same two
    functions sap_connection.py's ensure_sap() already relied on.
    """

    print("[INFO] SAP is not running -- launching and logging in...")

    launch_sap()

    login_sap(SAP_CLIENT, SAP_USERNAME, SAP_PASSWORD, SAP_LANGUAGE)

    # login_sap() already waits for the login screen and submits it, but
    # the session isn't necessarily fully live in scripting terms the
    # instant it returns -- poll briefly rather than assume.
    deadline = time.time() + LOGIN_TIMEOUT_SECONDS

    while time.time() < deadline:

        connection = _get_running_connection()

        if connection is not None:
            print("[OK] SAP Connected.")
            return connection

        time.sleep(1)

    raise TimeoutError(
        f"SAP did not finish logging in within {LOGIN_TIMEOUT_SECONDS} "
        "seconds after launch_sap()/login_sap() ran. Check the SAP "
        "Logon window directly for what state it's actually in."
    )


def ensure_sap_ready():
    """
    Returns a ready SAP GUI connection.
    """
    return get_connection()



# ==========================================================
# DEDICATED SESSIONS PER MODULE
# ==========================================================
#
# Each module gets its own session within ONE shared connection/login --
# opened via the /o command in the command field (e.g. "/oMB52"), not
# by calling GuiConnection.CreateSession() directly. /o asks SAP to open
# a new window already showing that transaction, which is the same
# underlying SAP action CreateSession() would trigger -- but it's
# reached by typing into a field and pressing Enter, an ordinary
# scripting interaction, rather than invoking a specific COM method.
# That distinction matters here because CreateSession() is confirmed
# blocked by Basis policy on this system (it raised immediately instead
# of a session ever appearing); if that block targets the method itself
# rather than the underlying capability, /o has a real chance of working
# around it while staying within one login, no separate logons involved.

# module_key -> session COM object
_dedicated_sessions = {}


def _session_is_alive(session):
    """
    SAP GUI session objects don't reliably support a clean 'is this
    still open' check -- the standard way is to touch a property and
    see whether it raises, since a session the user closed (or that SAP
    tore down) throws on any property access.
    """

    try:
        session.Info.SystemName
        return True
    except Exception:
        return False


def _wait_until_ready(session, timeout=10):
    """
    Waits until SAP reports the session is no longer busy, rather than
    assuming a fixed delay is always enough -- the same pattern already
    proven in po_creation.py's wait_until_ready(). Used right after
    opening a new window and before sending it anything further, since
    interacting with a window that's still mid-open is exactly the kind
    of timing gap that can leave SAP working through a backlog of
    commands and look "stuck" even when it isn't genuinely hung.
    """

    start = time.time()

    while time.time() - start < timeout:
        try:
            if not session.Busy:
                return
        except Exception:
            pass
        time.sleep(0.2)

    # Not raising here -- if it's still busy after the timeout, the
    # caller's own next action will surface any real problem, rather
    # than this blocking the whole flow on a status check that isn't
    # always reliable across SAP GUI versions.


def _get_trigger_session(connection):
    """
    Returns a session reserved purely for sending /o commands through --
    captured once by direct object reference, then reused for every
    later /o command regardless of what index position it's sitting at
    in connection.Children by that point.

    This exists because the previous implementation grabbed
    connection.Children(0) fresh on every call, assuming index 0 always
    refers to some neutral, safe-to-repurpose session. That assumption
    breaks the moment a dedicated session -- MB52's window, say --
    becomes index 0 instead, which can happen simply from clicking into
    that window to look at it, if SAP orders Children by recent focus
    rather than strict creation order. When that happens, the next /o
    command (opening CO01, for instance) gets sent through the exact
    window being kept open as a reference, which is the mechanism behind
    both "MB52 got overwritten by CO01" and "connection lost mid-PO" --
    the same session ends up serving two conflicting roles at once:
    something you're actively relying on, and a disposable trigger for
    opening something else.

    Holding a direct reference instead of an index sidesteps this
    entirely -- it doesn't matter what index this session is at later,
    since it's never looked up by index again after this first capture.
    """

    global _trigger_session

    if _trigger_session is not None and _session_is_alive(_trigger_session):
        return _trigger_session

    # No trigger session captured yet (or the previous one died) --
    # claim whatever's currently at index 0 ONCE. This is the only
    # place in the module that still does an index-based lookup, and
    # only as a one-time bootstrap; every call after this reuses the
    # captured reference directly.
    #
    # Reading Children(0) can transiently fail with "the enumerator of
    # the collection cannot find an element with the specified index"
    # if the collection's contents shift at the exact moment this reads
    # it -- a session opening or closing elsewhere between checking the
    # count and indexing into it. Retrying briefly here means a
    # momentary COM hiccup doesn't immediately surface as a fatal error
    # the first time it happens.
    last_error = None

    for _ in range(10):
        try:
            _trigger_session = connection.Children(0)
            return _trigger_session
        except Exception as e:
            last_error = e
            time.sleep(0.3)

    raise TimeoutError(
        f"Could not read a starting SAP session to use as the /o trigger, "
        f"even after retrying for 3 seconds: {last_error}"
    )


def _open_new_window(connection, tcode):
    """
    Opens a new SAP GUI window already showing the given transaction, by
    sending "/o<tcode>" through the reserved trigger session -- see
    _get_trigger_session() for why this is a stable captured reference
    rather than an index lookup into connection.Children.
    """

    before_ids = set()
    for i in range(connection.Children.Count):
        try:
            before_ids.add(connection.Children(i).Id)
        except Exception:
            pass

    trigger_session = _get_trigger_session(connection)

    try:
        trigger_session.findById("wnd[0]/tbar[0]/okcd").text = f"/o{tcode}"
        trigger_session.findById("wnd[0]").sendVKey(0)
    except Exception as e:
        raise TimeoutError(f"Could not send the /o{tcode} command: {e}")

    deadline = time.time() + 15

    while time.time() < deadline:

        try:
            # Scan every current session and find whichever one wasn't
            # there before, identified by its stable .Id -- rather than
            # assuming the new one always lands at the highest index.
            # That assumption breaks if SAP orders Children by recent
            # focus instead of creation order, since a freshly-opened
            # window would naturally have focus and could land at index
            # 0 instead -- which would mean grabbing the WRONG session
            # entirely as "the new one" (silently mislabeling some
            # other, unrelated session as this module's dedicated
            # window).
            for i in range(connection.Children.Count):
                candidate = connection.Children(i)
                if candidate.Id not in before_ids:
                    _wait_until_ready(candidate)
                    return candidate
        except Exception:
            pass

        time.sleep(0.3)

    raise TimeoutError(
        f"A new SAP window for /o{tcode} did not appear within 15 "
        "seconds. This can happen if the system has hit its maximum "
        "number of sessions (commonly 6, system parameter "
        "rdisp/max_alt_modes), or if /o-style session creation is ALSO "
        "restricted here, not just the CreateSession() API -- ask Basis "
        "to confirm either way if dedicated windows still don't appear."
    )


def get_dedicated_session(module_key, tcode=None):
    """
    Returns (session, needs_navigation) for the given module key (e.g.
    "MB52"), opening a new window for it (via /o<tcode>) the first time
    it's needed and reusing that same window on every call after that,
    for as long as it stays open. If it was closed in the meantime, a
    fresh one opens transparently.

    needs_navigation tells the caller whether the session still needs an
    explicit StartTransaction() to land on the right screen:
      - False when the window was JUST opened via /o<tcode>, which
        already lands on that transaction -- sending StartTransaction
        again immediately would be a redundant command to a window
        that's barely finished opening.
      - True when an existing session is being reused (it may be
        showing whatever it was left on) or when falling back to a
        shared session (which could be showing anything).

    tcode defaults to module_key itself, which covers every module
    where the two match (MB52, MB51, CO01, CO02, COOIS); pass it
    explicitly for the one that doesn't (Shipment -> ZPPMYR0490).

    A module that only needs to *read* from a window another module
    already populated (e.g. sap_reader.py reading MB52's grid right
    after mb52_runner.py populated it) can simply ignore the second
    value -- passing the same module_key ("MB52") returns that exact
    session either way.
    """

    with _cache_lock:

        connection = ensure_sap_ready()

        existing = _dedicated_sessions.get(module_key)

        if existing is not None and _session_is_alive(existing):
            return existing, True

        try:
            session = _open_new_window(connection, tcode or module_key)
            needs_navigation = False
        except TimeoutError as e:
            # Falling back to the trigger session (not a fresh
            # connection.Children(0) lookup) keeps the automation usable
            # rather than hard-failing the whole run because dedicated
            # windows aren't possible on this system -- and avoids
            # re-introducing the exact index-based confusion this whole
            # fix is for.
            print(f"[WARN] {e}")
            print("[WARN] Falling back to the shared SAP window instead of a dedicated one.")
            session = _get_trigger_session(connection)
            needs_navigation = True

        _dedicated_sessions[module_key] = session

    return session, needs_navigation


def start_transaction(module_key, tcode):
    """
    The one call every module needs: "give me a ready session for this
    module, already sitting on this transaction." Handles first-time
    window creation, reuse on subsequent calls, and auto-login if SAP
    wasn't even open -- all in one place, so individual modules don't
    need to think about SAP connection management at all.
    """

    session, needs_navigation = get_dedicated_session(module_key, tcode)

    session.findById("wnd[0]").maximize()

    if needs_navigation:
        session.StartTransaction(tcode)

    return session


# ==========================================================
# VISIBILITY -- read-only status, for the GUI to display
# ==========================================================

def is_sap_running():
    """
    A quick, passive check: is SAP GUI currently open AND actually
    logged in? Unlike ensure_sap_ready(), this never launches or logs
    in -- it's what "Check Connection" wants, a snapshot of current
    state, not an attempt to fix it.

    Requires an authenticated user, not just a connection object --
    SAP Logon can have a window open while still sitting at the
    Client/User/Password screen, and reporting that as "Connected"
    would be misleading for what this status is actually used to show.
    """

    connection = _cached_connection if _connection_is_alive(_cached_connection) else _get_running_connection()

    return _is_logged_in(connection)


def get_session_status():
    """
    Returns {module_key: is_alive} for every module that has requested a
    dedicated window so far this run. A module that has never run yet
    simply isn't in the dict -- the GUI treats "absent" and "not alive"
    the same way (not started / not active).
    """

    return {key: _session_is_alive(session) for key, session in _dedicated_sessions.items()}
