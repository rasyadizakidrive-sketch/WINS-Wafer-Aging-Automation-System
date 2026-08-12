import sys
import types

if sys.platform == "win32":
    # Some Windows security configurations restrict registry access for
    # a freshly-built, unsigned .exe more than they would for python.exe
    # itself -- seen as "PermissionError: [WinError 5] Access is denied"
    # coming from babel's own timezone detection, which tkcalendar pulls
    # in indirectly (tkcalendar -> babel.dates -> babel.localtime) just
    # by being imported, for locale-aware date formatting WINS doesn't
    # actually rely on.
    #
    # This has to happen as a pre-injected stub, not a "let it import,
    # then patch it" fix -- babel.localtime calls get_localzone() at its
    # own module-import time (not lazily on first use), so by the time
    # the import finishes -- or fails -- it's too late to patch.
    # Registering a fake babel.localtime._win32 module here means
    # Python's import system uses this instead of ever running the real
    # module's own registry-touching code.
    _fake_babel_win32 = types.ModuleType("babel.localtime._win32")

    from datetime import timezone as _utc_timezone

    _fake_babel_win32._get_localzone = lambda: _utc_timezone.utc
    _fake_babel_win32.get_localzone_name = lambda: "UTC"
    _fake_babel_win32.tz_names = {}

    sys.modules["babel.localtime._win32"] = _fake_babel_win32

import customtkinter as ctk
from GUI.main_window_v2 import MainWindow

def main():

    # Forced, not "System" -- every color in this app (COLOR_BG,
    # COLOR_CARD, etc.) is hardcoded for a dark theme, so following the
    # OS's light/dark preference here would leave any customtkinter-
    # native widget that doesn't have an explicit color set (default
    # scrollbar styling, for instance) rendering with light-mode
    # defaults on a light-mode Windows machine, clashing with everything
    # else that's intentionally dark.
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    app = MainWindow()
    app.mainloop()

if __name__ == "__main__":
    main()
