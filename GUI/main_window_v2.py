from tkinter import messagebox
import customtkinter as ctk
import config
from datetime import datetime
import threading
import io
import sys
import os
import csv
import time
import json
from GUI.create_po_dialog import CreatePODialog, POSummaryDialog
from GUI.po_tally_dialog import POTallySelectionDialog, POTallyProgressDialog, POTallyCompletionDialog
from PIL import Image
import icons

from Excel.excel_manager import ExcelManager
from main import run_automation
from SAP.shipment import run_shipment_automation
from SAP.mb51 import run_mb51_automation
from SAP.doi import run_doi
from SAP.po_creation import ProductionOrder
import logger

# ==============================================================
# APP / BRAND CONSTANTS
# ==============================================================

APP_TITLE = "WINS Wafer Loading Automation"
APP_SUBTITLE = "Production Planning Automation System"
APP_VERSION = "v2.0"
USER_FIRST_NAME = "Rasyadi"

COMPANY_LINE_1 = "Cell Manufacturing Control"
COMPANY_LINE_2 = "2026"
COMPANY_LINE_3 = "HANWHA Q CELLS Malaysia"

# Project root is one level above this file's folder (…/Wafer_Aging_Automation/GUI/this_file.py).
# When packaged as a PyInstaller .exe, the .exe itself sits directly at
# the top of the dist folder (no GUI/ nesting) -- sys.executable's
# directory IS the project root in that case, not two levels up.
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOGO_PATH = os.path.join(BASE_DIR, "Assets", "qcells_logo.png")

# Colors
COLOR_BG_HEADER = "#101826"
COLOR_CARD = "#1B2434"
COLOR_CARD_BORDER = "#2A3547"
COLOR_TEXT = "#F1F5F9"
COLOR_MUTED = "#94A3B8"

COLOR_OK = "#22C55E"
COLOR_WARN = "#F59E0B"
COLOR_ERROR = "#EF4444"
COLOR_INFO = "#60A5FA"

COLOR_PRIMARY = "#2563EB"
COLOR_PRIMARY_HOVER = "#1D4ED8"

COLOR_SECONDARY = "#059669"
COLOR_SECONDARY_HOVER = "#047857"

COLOR_TERTIARY = "#7C3AED"
COLOR_TERTIARY_HOVER = "#6D28D9"

COLOR_QUATERNARY = "#0D9488"
COLOR_QUATERNARY_HOVER = "#0F766E"

FONT_FAMILY = "Segoe UI"


# ==============================================================
# THREAD-SAFE STDOUT CAPTURE
# ==============================================================

class _ThreadLocalStdout:
    """
    Lets each background task capture only its own print() output.

    Plain contextlib.redirect_stdout swaps sys.stdout process-wide (it's a
    single global, not per-thread), so if anything else happens to print
    while a task is running -- another thread, a library call, a second
    task -- its output silently lands in the wrong task's captured log.
    This keeps capture strictly per-thread instead.
    """

    def __init__(self, default_stream):
        self._default = default_stream
        self._local = threading.local()

    def register(self, stream):
        self._local.stream = stream

    def unregister(self):
        if hasattr(self._local, "stream"):
            del self._local.stream

    def write(self, text):
        getattr(self._local, "stream", self._default).write(text)

    def flush(self):
        getattr(self._local, "stream", self._default).flush()


if not isinstance(sys.stdout, _ThreadLocalStdout):
    sys.stdout = _ThreadLocalStdout(sys.stdout)


# ==============================================================
# SIMPLE TOOLTIP HELPER
# ==============================================================

class ToolTip:
    """Small hover tooltip for widgets. Fails silently if it can't render."""

    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, event=None):

        if self.tip or not self.text:
            return

        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6

            self.tip = ctk.CTkToplevel()
            self.tip.withdraw()
            self.tip.overrideredirect(True)
            self.tip.geometry(f"+{x}+{y}")

            label = ctk.CTkLabel(
                self.tip,
                text=self.text,
                fg_color="#0B1220",
                text_color=COLOR_TEXT,
                corner_radius=6,
                font=(FONT_FAMILY, 11),
            )
            label.pack(padx=8, pady=5)

            self.tip.deiconify()
        except Exception:
            self.tip = None

    def hide(self, event=None):

        if self.tip is not None:
            try:
                self.tip.destroy()
            except Exception:
                pass
            self.tip = None


# ==============================================================
# STATUS INDICATOR (dot + label, optional tinted pill background)
# ==============================================================

class StatusIndicator(ctk.CTkFrame):
    """
    A small colored dot + label, standing in for plain colored text with
    an emoji glued on the front. Emoji circles render inconsistently
    across systems/fonts; an actual drawn dot reads as a real status
    indicator instead, the way a professional dashboard would show
    connection/run state.

    With pill=True it also sits on a subtly tinted rounded background
    (used for short, bounded text like "Connected"/"Disconnected").
    With pill=False it's just the dot + label with no background, for
    places where the text length varies too much for a fixed pill shape
    to look right (e.g. a status card that also shows longer messages
    like "Checking SAP (MB52)...").

    Exposes configure(text=, text_color=) / cget(...) so it drops in
    wherever a plain CTkLabel used to be, without touching call sites.
    """

    _TINTS = {
        COLOR_OK: "#132A1D",
        COLOR_WARN: "#332611",
        COLOR_ERROR: "#331515",
    }
    _DEFAULT_TINT = "#1E293B"

    def __init__(self, parent, text, text_color, pill=True, font_size=12, **kwargs):

        self._pill = pill

        bg = self._TINTS.get(text_color, self._DEFAULT_TINT) if pill else "transparent"
        super().__init__(parent, corner_radius=(14 if pill else 0), fg_color=bg, **kwargs)

        dot_padx = (10, 6) if pill else (0, 6)
        label_padx = (0, 12) if pill else (0, 0)
        v_pad = 6 if pill else 0

        self._dot = ctk.CTkFrame(self, width=8, height=8, corner_radius=4, fg_color=text_color)
        self._dot.pack(side="left", padx=dot_padx, pady=v_pad)
        self._dot.pack_propagate(False)

        self._label = ctk.CTkLabel(
            self, text=text, font=(FONT_FAMILY, font_size, "bold"), text_color=text_color
        )
        self._label.pack(side="left", padx=label_padx, pady=v_pad)

    def configure(self, text=None, text_color=None, **kwargs):

        if text is not None:
            self._label.configure(text=text)

        if text_color is not None:
            self._label.configure(text_color=text_color)
            self._dot.configure(fg_color=text_color)
            if self._pill:
                super().configure(fg_color=self._TINTS.get(text_color, self._DEFAULT_TINT))

        if kwargs:
            super().configure(**kwargs)

    def cget(self, param):
        if param == "text":
            return self._label.cget("text")
        if param == "text_color":
            return self._label.cget("text_color")
        return super().cget(param)

    def set_dot_color(self, color):
        """Used for the pulse animation -- moves just the dot, leaving the label text/color alone."""
        self._dot.configure(fg_color=color)


# ==============================================================
# WELCOME SCREEN
# ==============================================================

class WelcomeScreen(ctk.CTk):
    """
    A brief personalized greeting shown on startup, before the main
    window appears. Distinct from a loading splash -- there's no
    progress bar and nothing being waited on, just a warm "you're back"
    moment before the dashboard takes over.

    Like any startup screen, this is its own temporary CTk root (the
    real MainWindow doesn't exist yet at this point) and fully destroys
    itself when done -- see main() at the bottom of this file for how
    the two are sequenced.
    """

    _DISPLAY_MS = 2400

    def __init__(self):

        super().__init__()

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.overrideredirect(True)
        self.configure(fg_color="#05070d")

        w, h = 440, 300
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = (screen_w - w) // 2
        y = (screen_h - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

        border = ctk.CTkFrame(
            self, fg_color="#05070d", border_width=1, border_color=COLOR_CARD_BORDER, corner_radius=14
        )
        border.pack(fill="both", expand=True)

        content = ctk.CTkFrame(border, fg_color="transparent")
        content.pack(expand=True)

        now = datetime.now()
        hour = now.hour

        if 5 <= hour < 12:
            greeting = "Good Morning"
        elif 12 <= hour < 17:
            greeting = "Good Afternoon"
        elif 17 <= hour < 21:
            greeting = "Good Evening"
        else:
            greeting = "Good Night"

        ctk.CTkLabel(
            content, text=greeting, font=(FONT_FAMILY, 15), text_color=COLOR_MUTED,
        ).pack(pady=(0, 2))

        ctk.CTkLabel(
            content, text=USER_FIRST_NAME, font=(FONT_FAMILY, 30, "bold"), text_color=COLOR_TEXT,
        ).pack(pady=(0, 18))

        ctk.CTkFrame(content, height=1, width=90, fg_color=COLOR_CARD_BORDER).pack(pady=(0, 16))

        ctk.CTkLabel(
            content, text=now.strftime("%A"), font=(FONT_FAMILY, 13), text_color=COLOR_MUTED,
        ).pack()
        ctk.CTkLabel(
            content, text=now.strftime("%d %B"), font=(FONT_FAMILY, 13), text_color=COLOR_MUTED,
        ).pack(pady=(0, 4))
        ctk.CTkLabel(
            content, text=now.strftime("%H:%M"), font=(FONT_FAMILY, 22, "bold"), text_color=COLOR_TEXT,
        ).pack(pady=(2, 18))

        ctk.CTkLabel(
            content, text="Ready to Start Planning", font=(FONT_FAMILY, 12), text_color=COLOR_MUTED,
        ).pack()

        self.after(self._DISPLAY_MS, self._finish)

    def _cancel_pending_after_calls(self):
        """Same ScalingTracker cleanup as the earlier splash screen -- see
        that class's docstring for why this matters before destroy()."""
        try:
            for after_id in self.tk.call("after", "info"):
                try:
                    self.after_cancel(after_id)
                except Exception:
                    pass
        except Exception:
            pass

    def _finish(self):
        self._cancel_pending_after_calls()
        self.destroy()


# ==============================================================
# MAIN WINDOW
# ==============================================================

class MainWindow(ctk.CTk):

    def __init__(self):

        super().__init__()

        self.new_batches = []
        self.last_created_order = None
        self._po_running = False
        self._po_dialog_open = False
        self._log_entries = []
        self._log_pending_entries = []
        self._log_render_job = None
        self._log_filtered_cache = []
        self.current_page = "dashboard"

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title(f"{APP_TITLE} {APP_VERSION}")
        self.geometry("1220x780")
        self.minsize(900, 560)
        self.resizable(True, True)
        self.configure(fg_color="#0B1220")

        # Header stays frozen at the very top, unchanged from before.
        self._build_header(self)

        # Footer and status bar both anchor to the bottom of the window.
        # pack(side="bottom") stacks in call order from the bottom upward,
        # so building the footer first puts it at the very bottom edge,
        # with the live status bar sitting just above it.
        self._build_footer(self)

        # Below the header: a persistent bottom status bar, and in between,
        # a sidebar (navigation + product + connection) beside a content
        # area that itself splits into a page (swapped by the sidebar nav)
        # and a log panel that stays visible no matter which page is open --
        # the same reasoning VS Code/JetBrains use for a persistent output
        # panel regardless of which file is open.
        self._build_status_bar(self)

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True)

        self._build_sidebar(body)

        content_row = ctk.CTkFrame(body, fg_color="transparent")
        content_row.pack(side="left", fill="both", expand=True)

        self.page_container = ctk.CTkFrame(content_row, fg_color="transparent")
        self.page_container.pack(side="left", fill="both", expand=True, padx=(20, 10), pady=(16, 12))

        log_wrap = ctk.CTkFrame(content_row, fg_color="transparent")
        log_wrap.pack(side="left", fill="both", expand=True, padx=(14, 22), pady=(16, 12))
        self._build_log(log_wrap)

        self.pages = {}
        self._build_page_dashboard()
        self._build_page_automation()
        self._build_page_po()
        self._show_page("dashboard")

        # Ctrl+1/2/3 to jump between pages without touching the mouse.
        self.bind("<Control-Key-1>", lambda e: self._show_page("dashboard"))
        self.bind("<Control-Key-2>", lambda e: self._show_page("automation"))
        self.bind("<Control-Key-3>", lambda e: self._show_page("po"))

        self._refresh_po_dashboard()
        self._refresh_sap_sessions()

        self.after(300, self._update_sidebar_chart)

        self.log("Wafer Aging Automation Ready.", level="ok", module="System")

    # ==========================================================
    # UI BUILDERS
    # ==========================================================

    def _build_header(self, parent):

        header = ctk.CTkFrame(
            parent,
            fg_color=COLOR_BG_HEADER,
            corner_radius=0,
        )
        header.pack(fill="x")

        # A plain flow layout (pack, with padding) rather than a fixed
        # height + hand-placed pixel coordinates -- the header sizes
        # itself to whatever the title/subtitle actually need. A fixed
        # height combined with absolute .place(y=) coordinates is exactly
        # what clipped the descender on the "g" in "Loading" before: the
        # numbers were tuned against one font rendering and broke on
        # another. This can't clip, because nothing is asserting a size
        # smaller than the content.
        row = ctk.CTkFrame(header, fg_color="transparent")
        row.pack(fill="x", padx=24, pady=16)

        # ==========================================================
        # QCELLS LOGO (falls back to a plain placeholder mark if the
        # asset isn't found, so the app still opens on a fresh checkout)
        # ==========================================================

        logo_loaded = False

        if os.path.exists(LOGO_PATH):

            try:
                self.logo_image = ctk.CTkImage(
                    light_image=Image.open(LOGO_PATH),
                    dark_image=Image.open(LOGO_PATH),
                    size=(150, 46),
                )

                ctk.CTkLabel(row, image=self.logo_image, text="").pack(side="left")
                logo_loaded = True

            except Exception as e:
                print(f"[WARN] Could not load logo: {e}")

        if not logo_loaded:

            logo = ctk.CTkFrame(
                row,
                width=50,
                height=50,
                corner_radius=12,
                fg_color=COLOR_PRIMARY,
            )
            logo.pack(side="left")
            logo.pack_propagate(False)

            ctk.CTkLabel(
                logo,
                text="Q",
                font=(FONT_FAMILY, 22, "bold"),
                text_color="white",
            ).pack(expand=True)

        title_block = ctk.CTkFrame(row, fg_color="transparent")
        title_block.pack(side="left", padx=(16, 0))

        ctk.CTkLabel(
            title_block,
            text=APP_TITLE,
            font=(FONT_FAMILY, 22, "bold"),
            text_color=COLOR_TEXT,
        ).pack(anchor="w")

        ctk.CTkLabel(
            title_block,
            text=APP_SUBTITLE,
            font=(FONT_FAMILY, 13),
            text_color=COLOR_MUTED,
        ).pack(anchor="w", pady=(3, 0))

        version_badge = ctk.CTkFrame(
            row,
            corner_radius=10,
            fg_color="#1E293B",
        )
        version_badge.pack(side="right")

        ctk.CTkLabel(
            version_badge,
            text=APP_VERSION,
            font=(FONT_FAMILY, 12, "bold"),
            text_color=COLOR_MUTED,
        ).pack(padx=12, pady=4)

    def _build_sidebar(self, parent):

        sidebar = ctk.CTkFrame(parent, fg_color=COLOR_BG_HEADER, corner_radius=0, width=260)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        # Subtle vertical gradient -- a soft "light from above" effect
        # that fades out within the first ~220px rather than stretching
        # across the whole window height, which would be barely
        # perceptible on a tall window and is how subtle gradients are
        # typically used in real UI design (VS Code's own sidebar is
        # mostly flat with tonal variation between regions, not an
        # obvious gradient painted across the entire height).
        # A Canvas redrawn on resize, not a static PIL image -- CTkFrame
        # itself has no gradient-fill support at all.
        gradient_canvas = ctk.CTkCanvas(sidebar, highlightthickness=0, bg=COLOR_BG_HEADER)
        gradient_canvas.place(x=0, y=0, relwidth=1, relheight=1)

        gradient_top = self._lighten(COLOR_BG_HEADER, 0.18)
        gradient_fade_height = 260

        def redraw_gradient(event=None):
            gradient_canvas.delete("gradient")
            width = sidebar.winfo_width()
            height = sidebar.winfo_height()
            if width <= 1 or height <= 1:
                return
            fade_to = min(gradient_fade_height, height)
            for y in range(fade_to):
                t = y / fade_to
                color = self._lerp_color(gradient_top, COLOR_BG_HEADER, t)
                gradient_canvas.create_line(0, y, width, y, fill=color, tags="gradient")
            if fade_to < height:
                gradient_canvas.create_rectangle(
                    0, fade_to, width, height, fill=COLOR_BG_HEADER, outline="", tags="gradient"
                )
            gradient_canvas.tag_lower("gradient")

        sidebar.bind("<Configure>", redraw_gradient)
        sidebar.after(50, redraw_gradient)

        # Scrollable rather than a plain fixed frame -- the DOI trend card
        # added real height to the sidebar's fixed content, and a plain
        # frame with no scrolling just silently clips whatever doesn't
        # fit (which is exactly what happened to Check Connection). This
        # guarantees everything stays reachable regardless of window
        # height, the same fix used for the page content elsewhere.
        inner = ctk.CTkScrollableFrame(sidebar, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=14, pady=14)

        ctk.CTkLabel(
            inner, text="PRODUCT", font=(FONT_FAMILY, 11, "bold"), text_color=COLOR_MUTED
        ).pack(anchor="w")

        self.product = ctk.CTkComboBox(
            inner,
            values=["M6", "G12"],
            command=self.change_product,
        )
        self.product.pack(fill="x", pady=(6, 4))
        self.product.set(config.PRODUCT_NAME)

        self.description_label = ctk.CTkLabel(
            inner,
            text=config.PRODUCT_DESCRIPTION,
            font=(FONT_FAMILY, 10.5),
            text_color=COLOR_MUTED,
            wraplength=172,
            justify="left",
        )
        self.description_label.pack(anchor="w", pady=(0, 12))

        ctk.CTkFrame(inner, height=1, fg_color=COLOR_CARD_BORDER).pack(fill="x", pady=(0, 10))

        self.nav_buttons = {}
        self.nav_icons = {}
        for key, label, icon_kind in (
            ("dashboard", "  Dashboard", "inventory"),
            ("automation", "  Automation", "refresh"),
            ("po", "  Production Orders", "list"),
        ):
            icon_inactive = icons.make_icon(icon_kind, color="#94A3B8", size=15)
            icon_active = icons.make_icon(icon_kind, color="#FFFFFF", size=15)
            self.nav_icons[key] = (icon_inactive, icon_active)

            btn = ctk.CTkButton(
                inner,
                text=label,
                image=icon_inactive,
                compound="left",
                anchor="w",
                height=38,
                corner_radius=8,
                font=(FONT_FAMILY, 13),
                fg_color="transparent",
                hover_color="#1E293B",
                text_color=COLOR_MUTED,
                cursor="hand2",
                command=lambda k=key: self._show_page(k),
            )
            btn.pack(fill="x", pady=2)
            self.nav_buttons[key] = btn

        ctk.CTkLabel(
            inner,
            text="DOI TREND",
            font=(FONT_FAMILY, 10, "bold"),
            text_color=COLOR_MUTED,
        ).pack(anchor="w", pady=(12, 6))

        trend_card = ctk.CTkFrame(
            inner,
            fg_color="#162033",
            corner_radius=10
        )
        trend_card.pack(fill="x", pady=(0, 12))

        value_row = ctk.CTkFrame(trend_card, fg_color="transparent")
        value_row.pack(fill="x", padx=12, pady=(9, 0))

        self.sidebar_value_label = ctk.CTkLabel(
            value_row, text="--", font=(FONT_FAMILY, 22, "bold"), text_color=COLOR_MUTED
        )
        self.sidebar_value_label.pack(side="left")

        ctk.CTkLabel(
            value_row, text=" Days", font=(FONT_FAMILY, 11), text_color=COLOR_MUTED
        ).pack(side="left", padx=(3, 0), pady=(7, 0))

        self.sidebar_delta_label = ctk.CTkLabel(
            value_row, text="", font=(FONT_FAMILY, 10.5), text_color=COLOR_MUTED
        )
        self.sidebar_delta_label.pack(side="right", pady=(5, 0))

        self.sidebar_chart = ctk.CTkCanvas(trend_card, width=160, height=46, bg="#162033", highlightthickness=0)
        self.sidebar_chart.pack(padx=12, pady=(6, 3))

        self.sidebar_last_update = ctk.CTkLabel(
            trend_card,
            text="No history yet",
            font=(FONT_FAMILY, 9.5),
            text_color=COLOR_MUTED,
        )
        self.sidebar_last_update.pack(anchor="w", padx=12, pady=(0, 9))

        ctk.CTkFrame(inner, height=1, fg_color=COLOR_CARD_BORDER).pack(fill="x", pady=(4, 10))

        self.check_conn_btn = ctk.CTkButton(
            inner,
            text=" Check Connection",
            image=icons.make_icon("plug", color="#CBD5E1", size=14),
            compound="left",
            height=30,
            font=(FONT_FAMILY, 11, "bold"),
            fg_color="#334155",
            hover_color="#475569",
            command=self.check_connection,
            cursor="hand2",
        )
        self.check_conn_btn.pack(fill="x", pady=(0, 6))

    def _show_page(self, key):

        self.current_page = key

        for k, frame in self.pages.items():
            if k == key:
                frame.pack(fill="both", expand=True)
            else:
                frame.pack_forget()

        for k, btn in self.nav_buttons.items():
            icon_inactive, icon_active = self.nav_icons[k]
            if k == key:
                btn.configure(fg_color=COLOR_PRIMARY, text_color="white", image=icon_active)
            else:
                btn.configure(fg_color="transparent", text_color=COLOR_MUTED, image=icon_inactive)

        self._animate_page_transition()

    def _animate_page_transition(self):
        """
        A subtle wipe-reveal when switching pages: a solid overlay
        matching the background briefly covers the new page, then
        shrinks away left-to-right. This is layered on top of the
        already-packed page via .place() rather than animating the
        page's own geometry, specifically so it never has to touch
        CTkScrollableFrame's internal pack/place redirection (which has
        already shown real fragility elsewhere in this project).
        """
        try:
            overlay = ctk.CTkFrame(self.page_container, fg_color="#0B1220", corner_radius=0)
            overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
            overlay.lift()
        except Exception:
            return

        steps = 8
        duration_ms = 160

        def step(i=0):
            if not overlay.winfo_exists():
                return
            if i >= steps:
                overlay.destroy()
                return
            frac = (i + 1) / steps
            try:
                overlay.place(relx=frac, rely=0, relwidth=1 - frac, relheight=1)
            except Exception:
                return
            self.after(max(1, duration_ms // steps), lambda: step(i + 1))

        step()

    def _build_footer(self, parent):

        footer = ctk.CTkFrame(parent, fg_color=COLOR_BG_HEADER, corner_radius=0)
        footer.pack(fill="x", side="bottom")

        divider = ctk.CTkFrame(footer, height=1, fg_color=COLOR_CARD_BORDER)
        divider.pack(fill="x")

        ctk.CTkLabel(
            footer,
            text=f"{COMPANY_LINE_1}  \u00b7  {COMPANY_LINE_2}  \u00b7  {COMPANY_LINE_3}",
            font=(FONT_FAMILY, 10.5),
            text_color=COLOR_MUTED,
        ).pack(pady=5)

    def _make_gradient_image(self, width, height, color_top, color_bottom):
        """
        Generates a vertical gradient as a CTkImage, sized at a fixed
        width meant to safely cover realistic window widths -- CTkImage
        renders at a fixed size, it doesn't automatically stretch to
        fill whatever a label's own dynamic width happens to be, so
        this can't just generate at "however wide the bar is right now"
        and expect it to keep working if the window is resized wider.

        Used sparingly, on surfaces with a genuinely fixed HEIGHT (a
        rounded CTkFrame draws its own shape internally via canvas,
        which a background image can't safely sit behind without
        fighting that corner rendering -- this only targets
        square-cornered, fixed-height surfaces where that risk doesn't
        apply).
        """
        # Only `height` pixels actually differ -- compute a 1px-wide
        # column directly, then stretch it horizontally via PIL, rather
        # than redundantly computing the same color for every pixel in
        # a full width x height loop.
        column = Image.new("RGB", (1, height))
        top = color_top.lstrip("#")
        bottom = color_bottom.lstrip("#")
        r1, g1, b1 = int(top[0:2], 16), int(top[2:4], 16), int(top[4:6], 16)
        r2, g2, b2 = int(bottom[0:2], 16), int(bottom[2:4], 16), int(bottom[4:6], 16)
        pixels = column.load()
        for y in range(height):
            t = y / max(height - 1, 1)
            r = int(r1 + (r2 - r1) * t)
            g = int(g1 + (g2 - g1) * t)
            b = int(b1 + (b2 - b1) * t)
            pixels[0, y] = (r, g, b)

        img = column.resize((width, height), Image.NEAREST)

        return ctk.CTkImage(light_image=img, dark_image=img, size=(width, height))

    def _build_status_bar(self, parent):

        bar = ctk.CTkFrame(parent, fg_color=COLOR_BG_HEADER, corner_radius=0, height=44)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)

        # A subtle vertical gradient -- VS Code-style depth on a status
        # bar, rather than a flat fill. Placed before `inner` so it
        # naturally sits behind it; a 1px-wide image stretched to fill
        # the bar's width via relwidth=1, since the gradient only varies
        # top-to-bottom, not left-to-right.
        gradient_img = self._make_gradient_image(
            3000, 44, self._lighten(COLOR_BG_HEADER, 0.05), self._darken(COLOR_BG_HEADER, 0.15)
        )
        gradient_bg = ctk.CTkLabel(bar, image=gradient_img, text="")
        gradient_bg._gradient_image_ref = gradient_img  # keep a reference alive
        gradient_bg.place(x=0, y=0, relwidth=1, relheight=1)

        inner = ctk.CTkFrame(bar, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=20, pady=6)

        self.status_value = StatusIndicator(
            inner, text="Status : READY", text_color=COLOR_OK, pill=False, font_size=12
        )
        self.status_value.pack(side="left")

        self.sap_status = StatusIndicator(
            inner, text="SAP : Connected", text_color=COLOR_OK, pill=False, font_size=11.5
        )
        self.sap_status.pack(side="left", padx=(18, 0))

        self.excel_status = StatusIndicator(
            inner, text="Excel : Connected", text_color=COLOR_OK, pill=False, font_size=11.5
        )
        self.excel_status.pack(side="left", padx=(14, 0))

        self.current_batch_label = ctk.CTkLabel(
            inner, text="", font=(FONT_FAMILY, 11), text_color=COLOR_MUTED
        )
        self.current_batch_label.pack(side="left", padx=(18, 0))

        progress_wrap = ctk.CTkFrame(inner, fg_color="transparent")
        progress_wrap.pack(side="right", fill="y")

        self.progress_percent = ctk.CTkLabel(
            progress_wrap, text="0%", font=(FONT_FAMILY, 11), text_color=COLOR_MUTED, width=32
        )
        self.progress_percent.pack(side="right", padx=(8, 0))

        self.progress = ctk.CTkProgressBar(progress_wrap, height=8, width=200)
        self.progress.pack(side="right")
        self.progress.set(0)

    def _add_hover_elevation(self, card, hover_border="#3B82F6", base_border=None):
        """
        'Elevation' on hover: a brighter, thicker border plus a subtly
        lightened background, eased over a few quick steps rather than
        an instant color snap. Real box-shadow/elevation isn't
        something CTkFrame can render, so this is the practical,
        still-convincing equivalent of what's actually available.

        Purely polling-driven -- deliberately NOT using Tk's Enter/Leave
        events, after two separate failure modes turned up testing the
        event-based version: (1) recursive binding to every descendant
        needs those descendants to already exist at bind time, but every
        card in this app adds its actual content (labels, buttons, the
        caption text) AFTER calling this function, so the recursive walk
        always found zero children and only the card's own bare frame
        ever got bound -- hovering the visible content never registered
        at all; (2) winfo_containing(), used to detect "pointer still
        somewhere inside this card", can be fooled by a widget nested
        inside a CTkScrollableFrame, since CTkScrollableFrame implements
        scrolling via an internal Canvas that winfo_containing resolves
        to instead of the actual widget on top of it -- and every page
        in this app is a CTkScrollableFrame. A direct bounding-box check
        on a fast timer sidesteps both problems at once, rather than
        patching around either individually.

        The transition uses a single "progress" value (0.0 = base
        state, 1.0 = fully hovered) rather than animating toward a
        fixed target from a fixed start -- if the mouse leaves mid-way
        through the enter animation, the new leave animation just
        continues smoothly from wherever progress currently is, instead
        of jumping or needing special-case handling for interruption.
        """

        base_border = base_border or card.cget("border_color")
        base_width = card.cget("border_width")

        base_fill_raw = card.cget("fg_color")
        # fg_color can be a (light_mode, dark_mode) tuple in
        # customtkinter's theming system -- normalize to a plain hex
        # string, since this app is always forced to dark mode anyway.
        base_fill = base_fill_raw[1] if isinstance(base_fill_raw, (tuple, list)) else base_fill_raw

        can_animate_fill = isinstance(base_fill, str) and base_fill.startswith("#")
        hover_fill = self._lighten(base_fill, 0.08) if can_animate_fill else base_fill

        card._hover_progress = 0.0
        card._hover_anim_id = None
        card._hover_last_desired = 0.0

        STEPS = 4
        STEP_MS = 20
        STEP_DELTA = 1.0 / STEPS

        def cancel_pending():
            if card._hover_anim_id:
                try:
                    card.after_cancel(card._hover_anim_id)
                except Exception:
                    pass
                card._hover_anim_id = None

        def apply_progress():
            t = card._hover_progress
            try:
                card.configure(
                    border_color=self._lerp_color(base_border, hover_border, t),
                    border_width=(base_width + 1) if t > 0.5 else base_width,
                )
                if can_animate_fill:
                    card.configure(fg_color=self._lerp_color(base_fill, hover_fill, t))
            except Exception:
                pass

        def step_toward(target):
            cancel_pending()

            def do_step():
                if target > card._hover_progress:
                    card._hover_progress = min(target, card._hover_progress + STEP_DELTA)
                else:
                    card._hover_progress = max(target, card._hover_progress - STEP_DELTA)

                apply_progress()

                if abs(card._hover_progress - target) > 0.001:
                    card._hover_anim_id = card.after(STEP_MS, do_step)
                else:
                    card._hover_anim_id = None

            do_step()

        # Kept for any code that calls these directly (tests, or a
        # deliberate programmatic hover) -- not used by the polling
        # loop itself, which calls step_toward directly.
        card._hover_on_enter = lambda _e=None: step_toward(1.0)
        card._hover_on_leave = lambda _e=None: step_toward(0.0)

        def poll():
            desired = 1.0 if self._is_pointer_within(card) else 0.0
            if desired != card._hover_last_desired:
                card._hover_last_desired = desired
                step_toward(desired)

        self._register_hover_poll(poll)

    def _register_hover_poll(self, poll_fn):
        """
        Registers a widget's periodic hover-state check. Starts a
        single shared polling loop the first time this is called,
        rather than one timer per widget -- one .after() chain serving
        every registered widget each cycle is cheap even with dozens of
        widgets, since each poll_fn is just a bounding-box comparison.
        """

        if not hasattr(self, "_hover_poll_list"):
            self._hover_poll_list = []

        self._hover_poll_list.append(poll_fn)

        if not getattr(self, "_hover_watchdog_running", False):
            self._hover_watchdog_running = True
            self._run_hover_watchdog()

    def _run_hover_watchdog(self):

        for poll_fn in list(self._hover_poll_list):
            try:
                poll_fn()
            except Exception:
                pass

        self.after(60, self._run_hover_watchdog)

    def _make_card(self, parent, title, indicator=False, indicator_text="", indicator_color=COLOR_TEXT):

        card = ctk.CTkFrame(
            parent,
            fg_color=COLOR_CARD,
            border_color=COLOR_CARD_BORDER,
            border_width=1,
            corner_radius=12,
        )
        self._add_hover_elevation(card)

        ctk.CTkLabel(
            card,
            text=title,
            font=(FONT_FAMILY, 12, "bold"),
            text_color=COLOR_MUTED,
        ).pack(anchor="w", padx=16, pady=(14, 2))

        if indicator:
            value_label = StatusIndicator(
                card, text=indicator_text, text_color=indicator_color, pill=False, font_size=16
            )
        else:
            value_label = ctk.CTkLabel(
                card,
                text="",
                font=(FONT_FAMILY, 16, "bold"),
                text_color=COLOR_TEXT,
            )
        value_label.pack(anchor="w", padx=16, pady=(0, 14))

        return card, value_label

    # ----------------------------------------------------------
    # PAGE: DASHBOARD
    # ----------------------------------------------------------

    def _build_page_dashboard(self):

        page = ctk.CTkScrollableFrame(self.page_container, fg_color="transparent")
        self.pages["dashboard"] = page

        ctk.CTkLabel(
            page, text="Dashboard", font=(FONT_FAMILY, 19, "bold"), text_color=COLOR_TEXT
        ).pack(anchor="w", pady=(0, 2))
        ctk.CTkLabel(
            page, text="Live production data at a glance.",
            font=(FONT_FAMILY, 12), text_color=COLOR_MUTED,
        ).pack(anchor="w", pady=(0, 16))

        # DOI gets a full-width card of its own -- it carries more content
        # (two products + timestamp + a refresh action) than fits well
        # squeezed into a quarter-column.
        doi_card = ctk.CTkFrame(
            page, fg_color=COLOR_CARD, border_color=COLOR_CARD_BORDER, border_width=1, corner_radius=12
        )
        doi_card.pack(fill="x", pady=(0, 12))
        self._add_hover_elevation(doi_card)

        ctk.CTkLabel(
            doi_card, text="DAYS OF INVENTORY", font=(FONT_FAMILY, 12, "bold"), text_color=COLOR_MUTED
        ).pack(anchor="w", padx=16, pady=(14, 6))

        doi_columns = ctk.CTkFrame(doi_card, fg_color="transparent")
        doi_columns.pack(fill="x", padx=16, pady=(0, 4))
        doi_columns.grid_columnconfigure(0, weight=1, uniform="doi")
        doi_columns.grid_columnconfigure(1, weight=1, uniform="doi")

        m6_col = ctk.CTkFrame(doi_columns, fg_color="transparent")
        m6_col.grid(row=0, column=0, sticky="w")

        self.doi_m6_indicator = StatusIndicator(
            m6_col, text="M6 : -- Days", text_color=COLOR_MUTED, pill=False, font_size=16
        )
        self.doi_m6_indicator.pack(anchor="w")

        self.m6_until_label = ctk.CTkLabel(
            m6_col, text=" Until -", image=icons.make_icon("calendar", color=COLOR_MUTED, size=12),
            compound="left", font=(FONT_FAMILY, 10.5), text_color=COLOR_MUTED
        )
        self.m6_until_label.pack(anchor="w", pady=(6, 0))

        self.m6_time_label = ctk.CTkLabel(
            m6_col, text=" -", image=icons.make_icon("clock", color=COLOR_MUTED, size=12),
            compound="left", font=(FONT_FAMILY, 10.5), text_color=COLOR_MUTED
        )
        self.m6_time_label.pack(anchor="w")

        g12_col = ctk.CTkFrame(doi_columns, fg_color="transparent")
        g12_col.grid(row=0, column=1, sticky="w")

        self.doi_g12_indicator = StatusIndicator(
            g12_col, text="G12 : -- Days", text_color=COLOR_MUTED, pill=False, font_size=16
        )
        self.doi_g12_indicator.pack(anchor="w")

        self.g12_until_label = ctk.CTkLabel(
            g12_col, text=" Until -", image=icons.make_icon("calendar", color=COLOR_MUTED, size=12),
            compound="left", font=(FONT_FAMILY, 10.5), text_color=COLOR_MUTED
        )
        self.g12_until_label.pack(anchor="w", pady=(6, 0))

        self.g12_time_label = ctk.CTkLabel(
            g12_col, text=" -", image=icons.make_icon("clock", color=COLOR_MUTED, size=12),
            compound="left", font=(FONT_FAMILY, 10.5), text_color=COLOR_MUTED
        )
        self.g12_time_label.pack(anchor="w")

        doi_footer_row = ctk.CTkFrame(doi_card, fg_color="transparent")
        doi_footer_row.pack(fill="x", padx=16, pady=(6, 14))

        self.doi_updated_label = ctk.CTkLabel(
            doi_footer_row, text="Last Updated : -", font=(FONT_FAMILY, 10.5), text_color=COLOR_MUTED
        )
        self.doi_updated_label.pack(side="left")

        self.doi_button = ctk.CTkButton(
            doi_footer_row,
            text=" Refresh DOI",
            image=icons.make_icon("refresh", color="#CBD5E1", size=14),
            compound="left",
            width=120,
            height=28,
            font=(FONT_FAMILY, 11, "bold"),
            fg_color="#334155",
            hover_color="#475569",
            cursor="hand2",
            command=self.check_doi,
        )
        self.doi_button.pack(side="right")

        cards_wrap = ctk.CTkFrame(page, fg_color="transparent")
        cards_wrap.pack(fill="x", pady=(0, 12))
        cards_wrap.grid_columnconfigure(0, weight=1, uniform="c")
        cards_wrap.grid_columnconfigure(1, weight=1, uniform="c")

        material_card, self.material_value = self._make_card(cards_wrap, "MATERIAL")
        material_card.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        sheet_card, self.sheet_value = self._make_card(cards_wrap, "WORKSHEET")
        sheet_card.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        self.material_value.configure(text=config.SAP_MATERIAL)
        self.sheet_value.configure(text=config.SHEET_NAME)

        # PO summary strip
        ctk.CTkLabel(
            page, text="PRODUCTION ORDER SUMMARY", font=(FONT_FAMILY, 12, "bold"), text_color=COLOR_MUTED
        ).pack(anchor="w", pady=(4, 8))

        po_strip = ctk.CTkFrame(
            page, fg_color=COLOR_CARD, border_color=COLOR_CARD_BORDER, border_width=1, corner_radius=12
        )
        po_strip.pack(fill="x")
        self._add_hover_elevation(po_strip)

        po_inner = ctk.CTkFrame(po_strip, fg_color="transparent")
        po_inner.pack(fill="x", padx=16, pady=14)
        po_inner.grid_columnconfigure(0, weight=1, uniform="po")
        po_inner.grid_columnconfigure(1, weight=1, uniform="po")

        self.pending_po_badge = StatusIndicator(po_inner, text="Pending PO : -", text_color=COLOR_MUTED)
        self.pending_po_badge.grid(row=0, column=0, sticky="w", pady=(0, 8))

        self.created_today_badge = StatusIndicator(po_inner, text="Created Today : 0", text_color=COLOR_MUTED)
        self.created_today_badge.grid(row=0, column=1, sticky="w", pady=(0, 8))

        self.po_success_badge = StatusIndicator(po_inner, text="Success : 0", text_color=COLOR_MUTED)
        self.po_success_badge.grid(row=1, column=0, sticky="w")

        self.po_failed_badge = StatusIndicator(po_inner, text="Failed : 0", text_color=COLOR_MUTED)
        self.po_failed_badge.grid(row=1, column=1, sticky="w")

        # SAP sessions -- each module below (MB52, MB51, Shipment, CO01,
        # CO02, COOIS) now gets its own dedicated SAP GUI session rather
        # than all of them overwriting a single shared one. This shows
        # which of those dedicated sessions are actually alive right now.
        ctk.CTkLabel(
            page, text="SAP SESSIONS", font=(FONT_FAMILY, 12, "bold"), text_color=COLOR_MUTED
        ).pack(anchor="w", pady=(16, 8))

        sessions_card = ctk.CTkFrame(
            page, fg_color=COLOR_CARD, border_color=COLOR_CARD_BORDER, border_width=1, corner_radius=12
        )
        sessions_card.pack(fill="x")
        self._add_hover_elevation(sessions_card)

        sessions_inner = ctk.CTkFrame(sessions_card, fg_color="transparent")
        sessions_inner.pack(fill="x", padx=16, pady=14)
        sessions_inner.grid_columnconfigure(0, weight=1, uniform="sess")
        sessions_inner.grid_columnconfigure(1, weight=1, uniform="sess")

        self.sap_session_badges = {}

        sap_module_keys = ("MB52", "MB51", "Shipment", "CO01", "CO02", "COOIS", "ZPPMYR0520")
        # Computed from the actual count rather than a hardcoded row
        # number -- this is exactly what went stale when ZPPMYR0520 was
        # added as a 7th item: the grid grew from 3 rows to 4, but a
        # fixed "i < 4" threshold (correct for the old 3-row layout)
        # was left behind, so the new not-last row silently lost its
        # bottom padding. Deriving it here means adding an 8th module
        # later can't reintroduce the same bug.
        sap_last_row = (len(sap_module_keys) - 1) // 2

        for i, module_key in enumerate(sap_module_keys):
            badge = StatusIndicator(sessions_inner, text=f"{module_key} : Not started", text_color=COLOR_MUTED)
            row = i // 2
            badge.grid(row=row, column=i % 2, sticky="w", padx=(0, 10), pady=(0, 8 if row < sap_last_row else 0))
            self.sap_session_badges[module_key] = badge

    def _refresh_sap_sessions(self):
        """
        Reflects the current state of each module's dedicated SAP
        session on the Dashboard. Safe to call anytime -- a module that
        has never run yet just shows "Not started", not an error.
        """

        try:
            from SAP import sap_manager
            status = sap_manager.get_session_status()
        except Exception:
            status = {}

        for module_key, badge in self.sap_session_badges.items():

            if module_key not in status:
                badge.configure(text=f"{module_key} : Not started", text_color=COLOR_MUTED)
            elif status[module_key]:
                badge.configure(text=f"{module_key} : Active", text_color=COLOR_OK)
            else:
                badge.configure(text=f"{module_key} : Closed", text_color=COLOR_MUTED)

    # ----------------------------------------------------------
    # PAGE: AUTOMATION
    # ----------------------------------------------------------

    def _build_page_automation(self):

        page = ctk.CTkScrollableFrame(self.page_container, fg_color="transparent")
        self.pages["automation"] = page

        ctk.CTkLabel(
            page, text="Automation", font=(FONT_FAMILY, 19, "bold"), text_color=COLOR_TEXT
        ).pack(anchor="w", pady=(0, 2))
        ctk.CTkLabel(
            page, text="Batch detection, shipment, and goods-receipt sync.",
            font=(FONT_FAMILY, 12), text_color=COLOR_MUTED,
        ).pack(anchor="w", pady=(0, 10))

        self.batch_count = StatusIndicator(page, text="New Batch Found : 0", text_color=COLOR_MUTED)
        self.batch_count.pack(anchor="w", pady=(0, 14))

        actions_card = ctk.CTkFrame(
            page, fg_color=COLOR_CARD, border_color=COLOR_CARD_BORDER, border_width=1, corner_radius=12
        )
        actions_card.pack(fill="both", expand=True)

        inner = ctk.CTkFrame(actions_card, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=16, pady=16)

        self.check_btn = self._build_action_button(
            inner,
            "  CHECK MB52",
            "SAP TCODE : MB52\nScan SAP for new incoming batches.",
            COLOR_PRIMARY,
            COLOR_PRIMARY_HOVER,
            self.check_mb52,
            icon_kind="search",
        )

        self.update_batch_btn = self._build_action_button(
            inner,
            "  UPDATE BATCH & QUANTITY",
            "SAP TCODE : MB52\nWrite newly found batches into the Aging Excel sheet.",
            COLOR_SECONDARY,
            COLOR_SECONDARY_HOVER,
            self.update_batch_quantity,
            icon_kind="box",
        )
        self.update_batch_btn.configure(state="disabled")

        self.update_shipment_btn = self._build_action_button(
            inner,
            "  UPDATE SHIPMENT NUMBER",
            "SAP TCODE : ZPPMYR0490\nFetch shipment numbers from SAP and update Excel.",
            COLOR_TERTIARY,
            COLOR_TERTIARY_HOVER,
            self.update_shipment_number,
            icon_kind="truck",
        )

        self.update_gr_date_btn = self._build_action_button(
            inner,
            "  UPDATE GR DATE",
            "SAP TCODE : MB51\nFetch Goods Receipt dates from SAP and update Excel.",
            COLOR_QUATERNARY,
            COLOR_QUATERNARY_HOVER,
            self.update_gr_date,
            icon_kind="calendar",
        )

    # ----------------------------------------------------------
    # PAGE: PRODUCTION ORDERS
    # ----------------------------------------------------------

    def _build_page_po(self):

        page = ctk.CTkScrollableFrame(self.page_container, fg_color="transparent")
        self.pages["po"] = page

        ctk.CTkLabel(
            page, text="Production Orders", font=(FONT_FAMILY, 19, "bold"), text_color=COLOR_TEXT
        ).pack(anchor="w", pady=(0, 2))
        ctk.CTkLabel(
            page, text="Create, update, and review released production orders.",
            font=(FONT_FAMILY, 12), text_color=COLOR_MUTED,
        ).pack(anchor="w", pady=(0, 16))

        actions_card = ctk.CTkFrame(
            page, fg_color=COLOR_CARD, border_color=COLOR_CARD_BORDER, border_width=1, corner_radius=12
        )
        actions_card.pack(fill="both", expand=True)

        inner = ctk.CTkFrame(actions_card, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=16, pady=16)

        self.create_po_btn = self._build_action_button(
            inner,
            "  CREATE PRODUCTION ORDER",
            "SAP TCODE : CO01\nCreate Production Orders from Yet Firmed batches.",
            COLOR_PRIMARY,
            COLOR_PRIMARY_HOVER,
            self.open_create_po,
            icon_kind="doc_plus",
        )

        self.co02_btn = self._build_action_button(
            inner,
            "  UPDATE PRODUCTION ORDER",
            "SAP TCODE : CO02\nRead Master Data and save Production Order.",
            COLOR_SECONDARY,
            COLOR_SECONDARY_HOVER,
            self.update_production_order,
            icon_kind="refresh",
        )

        self.coois_btn = self._build_action_button(
            inner,
            "  PRODUCTION ORDER LIST (RELEASED)",
            "SAP TCODE : COOIS\nDisplay Released Production Order List.",
            COLOR_TERTIARY,
            COLOR_TERTIARY_HOVER,
            self.production_order_list,
            icon_kind="list",
        )

        self.po_tally_btn = self._build_action_button(
            inner,
            "  PRODUCTION ORDER VARIANCE CHECK",
            f"SAP TCODE : {config.PO_TALLY_TCODE}\nReconcile Target vs Yield/Scrap for selected Production Orders.",
            COLOR_QUATERNARY,
            COLOR_QUATERNARY_HOVER,
            self.open_po_tally,
            icon_kind="check",
        )

    # ----------------------------------------------------------
    # SHARED ACTION BUTTON BUILDER
    # ----------------------------------------------------------

    def _add_button_lift(self, btn, slot, lift_px=2, rest_y=2):
        """
        Animates btn's position within its slot on hover -- a literal
        upward shift (not just a color change), with the growing gap
        beneath it (revealing the slot's darker background) reading as
        a shadow that deepens as the button lifts. Uses place() within
        a fixed-height slot rather than moving the button within its
        parent's normal pack flow, so this can't disturb sibling
        widgets (the caption label below it, for instance) the way
        adjusting pack padding dynamically would.

        CTkButton already handles its own fg_color/hover_color switch
        natively and efficiently -- this only adds the position shift
        on top of that, not a color animation (color and lift are two
        genuinely separate concerns here).

        Polling-driven for the same reason as _add_hover_elevation: see
        that function's docstring for why Enter/Leave binding proved
        unreliable enough here to replace outright, rather than patch
        around case by case.
        """

        STEPS = 4
        STEP_MS = 16
        btn._lift_anim_id = None
        btn._lift_y = float(rest_y)
        btn._lift_last_desired = rest_y

        def cancel_pending():
            if btn._lift_anim_id:
                try:
                    btn.after_cancel(btn._lift_anim_id)
                except Exception:
                    pass
                btn._lift_anim_id = None

        def apply_position():
            try:
                btn.place(x=0, y=int(round(btn._lift_y)), relwidth=1)
            except Exception:
                pass

        def step_toward(target_y):
            cancel_pending()
            step_size = max(lift_px, 1) / STEPS

            def do_step():
                if target_y > btn._lift_y:
                    btn._lift_y = min(target_y, btn._lift_y + step_size)
                else:
                    btn._lift_y = max(target_y, btn._lift_y - step_size)

                apply_position()

                if abs(btn._lift_y - target_y) > 0.01:
                    btn._lift_anim_id = btn.after(STEP_MS, do_step)
                else:
                    btn._lift_anim_id = None

            do_step()

        # Kept for any code that calls these directly (tests, or a
        # deliberate programmatic hover) -- not used by the polling
        # loop itself, which calls step_toward directly.
        btn._lift_on_enter = lambda _e=None: step_toward(rest_y - lift_px)
        btn._lift_on_leave = lambda _e=None: step_toward(rest_y)

        def poll():
            desired = (rest_y - lift_px) if self._is_pointer_within(btn) else rest_y
            if desired != btn._lift_last_desired:
                btn._lift_last_desired = desired
                step_toward(desired)

        self._register_hover_poll(poll)

    def _build_action_button(self, parent, text, caption, color, hover, command, icon_kind=None):

        block = ctk.CTkFrame(
            parent,
            fg_color=COLOR_CARD,
            border_color=COLOR_CARD_BORDER,
            border_width=1,
            corner_radius=10,
        )
        block.pack(fill="x", pady=6)
        self._add_hover_elevation(block, hover_border=color)

        inner = ctk.CTkFrame(block, fg_color="transparent")
        inner.pack(fill="x", padx=8, pady=8)

        icon_image = icons.make_icon(icon_kind, color="#FFFFFF", size=18) if icon_kind else None

        # A fixed-height "slot" for the button rather than packing it
        # directly -- gives the button room to shift upward on hover
        # via place() without affecting the caption label below it.
        # The slot's own darker background shows through the gap that
        # grows beneath the button as it lifts, reading as a shadow.
        BUTTON_HEIGHT = 44
        LIFT_PX = 2
        slot = ctk.CTkFrame(
            inner, fg_color=self._darken(COLOR_CARD, 0.45),
            height=BUTTON_HEIGHT + LIFT_PX * 2, corner_radius=8,
        )
        slot.pack(fill="x")
        slot.pack_propagate(False)

        btn = ctk.CTkButton(
            slot,
            text=text,
            image=icon_image,
            compound="left",
            height=BUTTON_HEIGHT,
            corner_radius=8,
            font=(FONT_FAMILY, 13, "bold"),
            fg_color=color,
            hover_color=hover,
            command=command,
            cursor="hand2",
        )
        btn.place(x=0, y=LIFT_PX, relwidth=1)
        self._add_button_lift(btn, slot, lift_px=LIFT_PX, rest_y=LIFT_PX)

        ctk.CTkLabel(
            inner,
            text=caption,
            font=(FONT_FAMILY, 11),
            text_color=COLOR_MUTED,
            wraplength=380,
            justify="left",
        ).pack(anchor="w", pady=(4, 0))

        return btn

    def _build_log(self, parent):

        wrap = ctk.CTkFrame(
            parent,
            fg_color=COLOR_CARD,
            border_color=COLOR_CARD_BORDER,
            border_width=1,
            corner_radius=12,
        )
        wrap.pack(fill="both", expand=True)

        inner = ctk.CTkFrame(wrap, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=16, pady=16)

        header_row = ctk.CTkFrame(inner, fg_color="transparent")
        header_row.pack(fill="x")

        ctk.CTkLabel(
            header_row,
            text="Activity Log",
            font=(FONT_FAMILY, 13, "bold"),
            text_color=COLOR_TEXT,
        ).pack(side="left")

        btn_group = ctk.CTkFrame(header_row, fg_color="transparent")
        btn_group.pack(side="right")

        for text, cmd in (
            ("Clear", self.clear_log),
            ("Copy", self._copy_log),
            ("Export", self._export_log),
        ):
            ctk.CTkButton(
                btn_group, text=text, width=58, height=24, font=(FONT_FAMILY, 10.5),
                fg_color="#1E293B", hover_color="#334155", command=cmd, cursor="hand2",
            ).pack(side="right", padx=(6, 0))

        filter_row = ctk.CTkFrame(inner, fg_color="transparent")
        filter_row.pack(fill="x", pady=(10, 8))

        self.log_search = ctk.CTkEntry(filter_row, placeholder_text="Search log...", height=28)
        self.log_search.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.log_search.bind("<KeyRelease>", lambda e: self._render_log_rows())

        self.log_status_filter = ctk.CTkComboBox(
            filter_row, values=["All", "OK", "Warning", "Error"], width=104, height=28,
            command=lambda v: self._render_log_rows(),
        )
        self.log_status_filter.set("All")
        self.log_status_filter.pack(side="left")

        # Table header -- fixed widths for Time/Module/Status, Details fills
        # whatever's left. Row widgets below use the same widths so columns
        # actually line up.
        self._log_col_widths = {"time": 62, "module": 84, "status": 64}

        table_header = ctk.CTkFrame(inner, fg_color="#0F1626", corner_radius=6)
        table_header.pack(fill="x", pady=(0, 4))

        ctk.CTkLabel(
            table_header, text="Time", font=(FONT_FAMILY, 10.5, "bold"), text_color=COLOR_MUTED,
            anchor="w", width=self._log_col_widths["time"],
        ).pack(side="left", padx=(10, 4), pady=6)
        ctk.CTkLabel(
            table_header, text="Module", font=(FONT_FAMILY, 10.5, "bold"), text_color=COLOR_MUTED,
            anchor="w", width=self._log_col_widths["module"],
        ).pack(side="left", padx=4, pady=6)
        ctk.CTkLabel(
            table_header, text="Status", font=(FONT_FAMILY, 10.5, "bold"), text_color=COLOR_MUTED,
            anchor="w", width=self._log_col_widths["status"],
        ).pack(side="left", padx=4, pady=6)
        ctk.CTkLabel(
            table_header, text="Details", font=(FONT_FAMILY, 10.5, "bold"), text_color=COLOR_MUTED,
            anchor="w",
        ).pack(side="left", fill="x", expand=True, padx=4, pady=6)

        self.log_table_body = ctk.CTkScrollableFrame(inner, fg_color="transparent")
        self.log_table_body.pack(fill="both", expand=True)

        self._render_log_rows()

    # ==========================================================
    # LOGGING
    # ==========================================================

    _LOG_LEVEL_COLOR = {"ok": COLOR_OK, "warn": COLOR_WARN, "error": COLOR_ERROR, "info": COLOR_INFO}
    _LOG_LEVEL_LABEL = {"ok": "OK", "warn": "Warning", "error": "Error", "info": "Info"}
    _MAX_RENDERED_LOG_ROWS = 100

    @staticmethod
    def _detect_log_level(line):

        upper = line.strip().upper()

        if upper.startswith("[ERROR]"):
            return "error"
        if upper.startswith("[WARN"):
            return "warn"
        if upper.startswith("[OK]") or line.strip().startswith("✓"):
            return "ok"
        if upper.startswith("[INFO]"):
            return "info"

        return None

    def log(self, text, level=None, module=None):

        now = datetime.now().strftime("%H:%M:%S")
        new_entries = []

        for line in text.splitlines():

            if not line.strip():
                continue

            tag = level or self._detect_log_level(line)
            entry = {"time": now, "module": module or "System", "level": tag, "text": line}

            self._log_entries.append(entry)
            new_entries.append(entry)

        if new_entries:
            self._log_pending_entries.extend(new_entries)
            self._schedule_log_append()

    def _schedule_log_append(self):
        """
        Rebuilding the whole table is not free (widget creation adds up
        fast), and during a bulk operation log() can fire many times in
        a couple of seconds. Debouncing collapses rapid successive calls
        into a single flush instead of one per line.
        """
        if self._log_render_job is not None:
            try:
                self.after_cancel(self._log_render_job)
            except Exception:
                pass
        self._log_render_job = self.after(120, self._flush_log_append)

    def _log_matches_filter(self, entry):

        status_filter = self.log_status_filter.get() if hasattr(self, "log_status_filter") else "All"

        if status_filter != "All":
            wanted = status_filter.lower()
            level = entry["level"] or "info"
            if wanted == "ok" and level != "ok":
                return False
            if wanted == "warning" and level != "warn":
                return False
            if wanted == "error" and level != "error":
                return False

        query = self.log_search.get().strip().lower() if hasattr(self, "log_search") else ""
        if query:
            haystack = f"{entry['time']} {entry['module']} {entry['text']}".lower()
            if query not in haystack:
                return False

        return True

    def _filter_active(self):
        query = self.log_search.get().strip() if hasattr(self, "log_search") else ""
        status_filter = self.log_status_filter.get() if hasattr(self, "log_status_filter") else "All"
        return bool(query) or status_filter != "All"

    def _flush_log_append(self):
        """
        The common case: new lines arrived and nothing about the visible
        filter changed. Just append the new matching rows and trim any
        excess off the top, rather than tearing down and rebuilding
        everything that was already correctly on screen.
        """
        self._log_render_job = None
        pending = self._log_pending_entries
        self._log_pending_entries = []

        if self._filter_active():
            # A search/filter is active -- whether newly-arrived entries
            # even belong in the current view depends on matching against
            # it, and a full, consistent re-render is simpler and safer
            # than reasoning about partial matches mid-filter.
            self._render_log_rows()
            return

        self._log_filtered_cache = list(self._log_entries)

        if len(pending) >= self._MAX_RENDERED_LOG_ROWS:
            # A single burst (e.g. a large bulk PO run logging many lines
            # before the debounce window closes) already exceeds the cap
            # on its own -- every existing row would be trimmed away
            # regardless, so build only the rows that will actually
            # survive instead of creating and immediately destroying them.
            for child in self.log_table_body.winfo_children():
                child.destroy()
            for entry in pending[-self._MAX_RENDERED_LOG_ROWS:]:
                self._append_log_row(entry)
        else:
            for entry in pending:
                self._append_log_row(entry)
            self._trim_log_rows()

        # update_idletasks() first -- otherwise the canvas's scroll
        # region still reflects the content height from BEFORE these new
        # rows were packed, so yview_moveto(1.0) would scroll to what it
        # thinks is the bottom while actually landing short of the real
        # latest entry.
        self.log_table_body.update_idletasks()
        self.log_table_body._parent_canvas.yview_moveto(1.0)

    def _build_log_row(self, parent, entry, bg):

        w = self._log_col_widths
        level = entry["level"] or "info"
        color = self._LOG_LEVEL_COLOR.get(level, COLOR_MUTED)
        label_text = self._LOG_LEVEL_LABEL.get(level, "Info")

        details = entry["text"]
        details_short = details if len(details) <= 70 else details[:67] + "..."

        row = ctk.CTkFrame(parent, fg_color=bg, corner_radius=6, cursor="hand2")

        cells = [
            ctk.CTkLabel(row, text=entry["time"], font=(FONT_FAMILY, 10.5), text_color=COLOR_MUTED, anchor="w", width=w["time"]),
            ctk.CTkLabel(row, text=entry["module"], font=(FONT_FAMILY, 10.5), text_color=COLOR_TEXT, anchor="w", width=w["module"]),
            ctk.CTkLabel(row, text=f"\u25cf {label_text}", font=(FONT_FAMILY, 10.5, "bold"), text_color=color, anchor="w", width=w["status"]),
            ctk.CTkLabel(row, text=details_short, font=(FONT_FAMILY, 10.5), text_color=COLOR_TEXT, anchor="w"),
        ]

        cells[0].pack(side="left", padx=(10, 4), pady=6)
        cells[1].pack(side="left", padx=4, pady=6)
        cells[2].pack(side="left", padx=4, pady=6)
        cells[3].pack(side="left", fill="x", expand=True, padx=4, pady=6)

        for widget in (row, *cells):
            widget.bind("<Double-Button-1>", lambda e, entry=entry: self._show_log_detail(entry))

        return row

    def _append_log_row(self, entry):

        # Clear the "no entries match" placeholder if this is the first
        # real row landing in an empty table.
        existing = self.log_table_body.winfo_children()
        if len(existing) == 1 and not isinstance(existing[0], ctk.CTkFrame):
            existing[0].destroy()
            existing = []

        row_bg = "#131b2c" if len(existing) % 2 == 0 else "transparent"
        row = self._build_log_row(self.log_table_body, entry, row_bg)
        row.pack(fill="x", pady=1)

    def _trim_log_rows(self):

        children = self.log_table_body.winfo_children()
        excess = len(children) - self._MAX_RENDERED_LOG_ROWS

        if excess > 0:
            for child in children[:excess]:
                child.destroy()

    def _render_log_rows(self):
        """Full rebuild -- used when the filter/search changes, or on clear."""

        self._log_render_job = None
        self._log_pending_entries = []

        for child in self.log_table_body.winfo_children():
            child.destroy()

        filtered = [e for e in self._log_entries if self._log_matches_filter(e)]
        self._log_filtered_cache = filtered

        if not filtered:
            ctk.CTkLabel(
                self.log_table_body, text="No log entries match.",
                font=(FONT_FAMILY, 11.5), text_color=COLOR_MUTED,
            ).pack(pady=24)
            return

        to_render = filtered[-self._MAX_RENDERED_LOG_ROWS:]

        for i, entry in enumerate(to_render):
            row_bg = "#131b2c" if i % 2 == 0 else "transparent"
            row = self._build_log_row(self.log_table_body, entry, row_bg)
            row.pack(fill="x", pady=1)

        # update_idletasks() first -- otherwise the canvas's scroll
        # region still reflects the content height from BEFORE these new
        # rows were packed, so yview_moveto(1.0) would scroll to what it
        # thinks is the bottom while actually landing short of the real
        # latest entry.
        self.log_table_body.update_idletasks()
        self.log_table_body._parent_canvas.yview_moveto(1.0)

    # ==========================================================
    # SMART ERROR
    # ==========================================================

    _ERROR_PATTERNS = (
        {
            "match": ("timeout", "timed out"),
            "title": "SAP Session Lost",
            "reasons": ["SAP timeout", "User logout", "Connection closed"],
            "reconnect": True,
        },
        {
            "match": ("logon", "log on", "logged on", "logged off", "session", "rfc_abap"),
            "title": "SAP Session Lost",
            "reasons": ["SAP timeout", "User logout", "Connection closed"],
            "reconnect": True,
        },
        {
            "match": ("not connected", "connection refused", "rpc server", "sapgui"),
            "title": "Cannot Reach SAP",
            "reasons": ["SAP GUI is not open", "Network connection is down", "VPN disconnected"],
            "reconnect": True,
        },
        {
            "match": ("excel", "workbook"),
            "title": "Excel File Issue",
            "reasons": ["Excel file is not open", "File is locked by another user", "File path has changed"],
            "reconnect": True,
        },
        {
            "match": ("blocked",),
            "title": "Material Blocked in SAP",
            "reasons": ["Material has a quality block", "Material master is incomplete", "Plant-specific block active"],
            "reconnect": False,
        },
    )

    def _classify_error(self, raw_error):

        haystack = raw_error.lower()

        for pattern in self._ERROR_PATTERNS:
            if any(needle in haystack for needle in pattern["match"]):
                return pattern

        return {
            "title": "Something Went Wrong",
            "reasons": [],
            "reconnect": False,
        }

    def _show_smart_error(self, raw_error):

        info = self._classify_error(raw_error)

        popup = ctk.CTkToplevel(self)
        popup.title("Error")
        popup.configure(fg_color="#0B1220")
        popup.geometry("460x360")
        popup.resizable(True, True)
        popup.grab_set()

        content = ctk.CTkFrame(popup, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=22, pady=20)

        header = ctk.CTkFrame(content, fg_color="transparent")
        header.pack(fill="x", pady=(0, 4))

        ctk.CTkLabel(
            header, text="\u26a0", font=(FONT_FAMILY, 20), text_color=COLOR_ERROR,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkLabel(
            header, text=info["title"], font=(FONT_FAMILY, 16, "bold"), text_color=COLOR_TEXT,
        ).pack(side="left")

        if info["reasons"]:

            ctk.CTkLabel(
                content, text="Possible reason", font=(FONT_FAMILY, 11, "bold"), text_color=COLOR_MUTED,
            ).pack(anchor="w", pady=(14, 6))

            for reason in info["reasons"]:
                ctk.CTkLabel(
                    content, text=f"\u2022  {reason}", font=(FONT_FAMILY, 12), text_color=COLOR_TEXT, anchor="w",
                ).pack(anchor="w", pady=2)

        ctk.CTkFrame(content, height=1, fg_color=COLOR_CARD_BORDER).pack(fill="x", pady=(14, 10))

        ctk.CTkLabel(
            content, text="Details", font=(FONT_FAMILY, 10.5, "bold"), text_color=COLOR_MUTED,
        ).pack(anchor="w", pady=(0, 4))

        # Error text length varies wildly (a short "connection refused"
        # vs. a full COM exception dump) -- give it more room than a
        # cramped fixed box, while still capping it so one very long
        # error can't push the buttons off-screen; the textbox itself
        # scrolls internally if the text still exceeds this.
        details_box = ctk.CTkTextbox(
            content, height=110, font=(FONT_FAMILY, 10.5), fg_color="#131b2c",
            text_color=COLOR_MUTED, wrap="word",
        )
        details_box.pack(fill="x")
        details_box.insert("1.0", raw_error)
        details_box.configure(state="disabled")

        btn_row = ctk.CTkFrame(content, fg_color="transparent")
        btn_row.pack(fill="x", pady=(16, 0), side="bottom")

        ctk.CTkButton(
            btn_row, text="Close", command=popup.destroy, width=100,
            fg_color="#334155", hover_color="#475569", cursor="hand2",
        ).pack(side="right")

        if info["reconnect"]:
            ctk.CTkButton(
                btn_row,
                text="\U0001f504 Reconnect",
                command=lambda: [popup.destroy(), self.check_connection()],
                fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER, cursor="hand2",
            ).pack(side="left")

        # Size the window to fit what actually got built, instead of a
        # fixed guess -- that guess was consistently too tight (the
        # button row was getting clipped off the bottom), and a fixed
        # number would only ever be "tight enough" for one specific
        # error length. Capped so a very long message doesn't grow the
        # window past a reasonable share of the screen.
        popup.update_idletasks()
        needed_height = content.winfo_reqheight() + 40
        max_height = int(popup.winfo_screenheight() * 0.85)
        final_height = min(max(needed_height, 360), max_height)

        screen_w = popup.winfo_screenwidth()
        screen_h = popup.winfo_screenheight()
        x = (screen_w - 460) // 2
        y = (screen_h - final_height) // 2
        popup.geometry(f"460x{final_height}+{x}+{y}")

    def _show_log_detail(self, entry):

        popup = ctk.CTkToplevel(self)
        popup.title("Log Entry")
        popup.configure(fg_color="#0B1220")
        popup.geometry("440x280")
        popup.resizable(False, False)
        popup.grab_set()

        content = ctk.CTkFrame(popup, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=20, pady=20)

        level = entry["level"] or "info"
        color = self._LOG_LEVEL_COLOR.get(level, COLOR_MUTED)
        label_text = self._LOG_LEVEL_LABEL.get(level, "Info")

        def info_row(label, value, value_color=COLOR_TEXT):
            r = ctk.CTkFrame(content, fg_color="transparent")
            r.pack(fill="x", pady=3)
            ctk.CTkLabel(r, text=label, font=(FONT_FAMILY, 11), text_color=COLOR_MUTED, width=64, anchor="w").pack(side="left")
            ctk.CTkLabel(r, text=value, font=(FONT_FAMILY, 11.5, "bold"), text_color=value_color, anchor="w").pack(side="left")

        info_row("Time", entry["time"])
        info_row("Module", entry["module"])
        info_row("Status", label_text, color)

        ctk.CTkFrame(content, height=1, fg_color=COLOR_CARD_BORDER).pack(fill="x", pady=10)

        details_box = ctk.CTkTextbox(
            content, font=(FONT_FAMILY, 11.5), fg_color="#0B1220", text_color=COLOR_TEXT, wrap="word",
        )
        details_box.pack(fill="both", expand=True)
        details_box.insert("1.0", entry["text"])
        details_box.configure(state="disabled")

        ctk.CTkButton(content, text="Close", command=popup.destroy, width=100).pack(pady=(12, 0), anchor="e")

    def _copy_log(self):

        rows = getattr(self, "_log_filtered_cache", self._log_entries)

        if not rows:
            messagebox.showinfo("Copy", "No log entries to copy.")
            return

        lines = ["Time\tModule\tStatus\tDetails"]
        for e in rows:
            level_label = self._LOG_LEVEL_LABEL.get(e["level"] or "info", "Info")
            lines.append(f"{e['time']}\t{e['module']}\t{level_label}\t{e['text']}")

        try:
            self.clipboard_clear()
            self.clipboard_append("\n".join(lines))
            messagebox.showinfo("Copied", f"Copied {len(rows)} log entries to clipboard.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not copy to clipboard: {e}")

    def _export_log(self):

        rows = getattr(self, "_log_filtered_cache", self._log_entries)

        if not rows:
            messagebox.showinfo("Export", "No log entries to export.")
            return

        logs_dir = os.path.join(BASE_DIR, "Logs")
        os.makedirs(logs_dir, exist_ok=True)
        filename = os.path.join(logs_dir, f"activity_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")

        try:
            with open(filename, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Time", "Module", "Status", "Details"])
                for e in rows:
                    level_label = self._LOG_LEVEL_LABEL.get(e["level"] or "info", "Info")
                    writer.writerow([e["time"], e["module"], level_label, e["text"]])
            messagebox.showinfo("Export Complete", f"Exported {len(rows)} entries to:\n{filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not export log: {e}")

    def clear_log(self):
        self._log_entries = []
        self._render_log_rows()

    # ==========================================================
    # STATUS / PROGRESS HELPERS
    # ==========================================================

    def _set_running(self, text):

        self.status_value.configure(text=f"Status : {text}", text_color=COLOR_WARN)
        self._start_pulse([self.status_value])

        self.progress_percent.configure(text="")

        try:
            self.progress.configure(mode="indeterminate")
            self.progress.start()
        except Exception:
            self.progress.set(0.4)

        # Force the status bar to re-layout right now rather than
        # whenever Tk's next idle cycle happens to land -- this and
        # _set_connection_indicators() often fire back-to-back (see
        # update_shipment_number() etc.), and a longer "busy" string
        # replacing a short default one needs its neighbors pushed over
        # immediately, not on a delay.
        self.update_idletasks()

    def _set_finished(self, text, color, progress_value):

        self._stop_pulse([self.status_value])

        try:
            self.progress.stop()
            self.progress.configure(mode="determinate")
        except Exception:
            pass

        self.progress.set(progress_value)
        self.progress_percent.configure(text=f"{int(progress_value * 100)}%")

        self.status_value.configure(text=f"Status : {text}", text_color=color)

        # Every operation that finishes may have created, reused, or lost
        # a dedicated SAP session -- reflect that on the Dashboard
        # regardless of which specific operation just ran.
        self._refresh_sap_sessions()

    def _set_connection_indicators(self, ok, busy_text=None, targets=("sap", "excel")):

        indicators = []
        if "sap" in targets:
            indicators.append(self.sap_status)
        if "excel" in targets:
            indicators.append(self.excel_status)

        if busy_text:
            for name, indicator in (("SAP", self.sap_status), ("Excel", self.excel_status)):
                if indicator in indicators:
                    indicator.configure(text=f"{name} : {busy_text}", text_color=COLOR_WARN)
            self._start_pulse(indicators)
            self.update_idletasks()
            return

        self._stop_pulse(indicators)

        color = COLOR_OK if ok else COLOR_ERROR
        label = "Connected" if ok else "Error"

        for name, indicator in (("SAP", self.sap_status), ("Excel", self.excel_status)):
            if indicator in indicators:
                indicator.configure(text=f"{name} : {label}", text_color=color)

        self.update_idletasks()

    # ----------------------------------------------------------
    # Pulse animation: alternates a busy indicator's dot between two
    # shades every ~450ms, so an active SAP/Excel read or a running task
    # reads as "live" rather than a static label. Indicators are tracked
    # in a shared set so overlapping pulses (e.g. the status card *and*
    # the SAP/Excel dots at once during Check MB52) don't stomp on each
    # other -- starting one doesn't cancel another that's already running.
    # ----------------------------------------------------------

    _PULSE_SHADES = (COLOR_WARN, "#8A6212")

    def _start_pulse(self, indicators):
        if not hasattr(self, "_pulsing"):
            self._pulsing = set()
        self._pulsing.update(indicators)
        if getattr(self, "_pulse_job", None) is None:
            self._pulse_tick(0)

    def _stop_pulse(self, indicators):
        if hasattr(self, "_pulsing"):
            self._pulsing.difference_update(indicators)

    def _pulse_tick(self, tick):
        if not self.winfo_exists():
            self._pulse_job = None
            return

        if getattr(self, "_pulsing", None):
            shade = self._PULSE_SHADES[tick % 2]
            for indicator in list(self._pulsing):
                try:
                    indicator.set_dot_color(shade)
                except Exception:
                    pass
            self._pulse_job = self.after(450, lambda: self._pulse_tick(tick + 1))
        else:
            self._pulse_job = None

    def _refresh_po_dashboard(self, excel=None):
        """
        Updates the PO dashboard strip. Created Today / Success / Failed
        come from today's log file (cheap, local, no SAP/Excel dependency,
        so this can be called anytime). Pending PO requires an open Excel
        connection -- pass one in (reusing whatever the calling operation
        already has open) to refresh it; otherwise it's left as-is.
        """

        success, failed = logger.summarize_today()
        created_today = success + failed

        self.created_today_badge.configure(text=f"Created Today : {created_today}")
        self.po_success_badge.configure(
            text=f"Success : {success}",
            text_color=COLOR_OK if success else COLOR_MUTED,
        )
        self.po_failed_badge.configure(
            text=f"Failed : {failed}",
            text_color=COLOR_ERROR if failed else COLOR_MUTED,
        )

        if excel is not None:
            try:
                pending = len(excel.get_po_batches())
                self.pending_po_badge.configure(
                    text=f"Pending PO : {pending}",
                    text_color=COLOR_WARN if pending else COLOR_MUTED,
                )
            except Exception:
                pass

    def _set_buttons_state(self, state):

        self.check_btn.configure(state=state)
        self.update_shipment_btn.configure(state=state)
        self.update_gr_date_btn.configure(state=state)
        self.doi_button.configure(state=state)
        self.create_po_btn.configure(state=state)
        self.co02_btn.configure(state=state)
        self.coois_btn.configure(state=state)
        self.po_tally_btn.configure(state=state)
        self.product.configure(state=("disabled" if state == "disabled" else "normal"))

        if state == "disabled":
            self.update_batch_btn.configure(state="disabled")
        else:
            self.update_batch_btn.configure(
                state="normal" if self.new_batches else "disabled"
            )

    # ==========================================================
    # GENERIC BACKGROUND TASK RUNNER
    # ==========================================================

    def _run_task(self, task_fn, on_finish):
        """
        Runs task_fn() in a background thread, capturing stdout and any
        exception, then calls on_finish(output_text, result, error) back
        on the main thread.
        """

        def worker():

            output = io.StringIO()
            result = None
            error = None

            sys.stdout.register(output)
            try:
                result = task_fn()
            except Exception as e:
                error = e
            finally:
                sys.stdout.unregister()

            self.after(0, lambda: on_finish(output.getvalue(), result, error))

        threading.Thread(target=worker, daemon=True).start()

    def _doi_history_path(self):
        return os.path.join(BASE_DIR, "doi_history.json")


    def _load_doi_history(self):
        path = self._doi_history_path()

        if not os.path.exists(path):
            return {"M6": [], "G12": []}

        try:
            with open(path, "r") as f:
                return json.load(f)
        except Exception:
            return {"M6": [], "G12": []}


    def _save_doi_history(self, history):
        try:
            with open(self._doi_history_path(), "w") as f:
                json.dump(history, f, indent=4)
        except Exception:
            pass


    def _is_pointer_within(self, widget):
        """
        Checks whether the mouse is currently within widget's own
        screen rectangle -- a direct geometric comparison using the
        widget's own winfo_rootx/rooty/width/height, rather than asking
        Tk's window hierarchy "which widget owns this position"
        (winfo_containing). That distinction matters: winfo_containing
        can be fooled by a widget nested inside a CTkScrollableFrame,
        since CTkScrollableFrame implements scrolling internally via a
        Canvas, and winfo_containing resolved to that canvas rather
        than the actual widget visually sitting on top of it -- which
        is exactly what caused the reported "stuck border glow" bug on
        pages built as CTkScrollableFrame (every page in this app).
        This check sidesteps that resolution entirely.
        """
        try:
            x, y = widget.winfo_pointerxy()
            wx = widget.winfo_rootx()
            wy = widget.winfo_rooty()
            return wx <= x <= wx + widget.winfo_width() and wy <= y <= wy + widget.winfo_height()
        except Exception:
            return False

    def _lighten(self, hex_color, amount=0.15):
        """Blends hex_color toward white by amount (0-1)."""
        hex_color = hex_color.lstrip("#")
        r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        r = int(r + (255 - r) * amount)
        g = int(g + (255 - g) * amount)
        b = int(b + (255 - b) * amount)
        return f"#{r:02x}{g:02x}{b:02x}"

    def _darken(self, hex_color, amount=0.15):
        """Blends hex_color toward black by amount (0-1)."""
        hex_color = hex_color.lstrip("#")
        r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        r = int(r * (1 - amount))
        g = int(g * (1 - amount))
        b = int(b * (1 - amount))
        return f"#{r:02x}{g:02x}{b:02x}"

    def _lerp_color(self, color_a, color_b, t):
        """Linearly interpolates between two hex colors at t (0-1) --
        used to step through a short sequence of in-between colors for
        an eased hover transition, rather than an instant color snap."""
        a = color_a.lstrip("#")
        b = color_b.lstrip("#")
        ar, ag, ab = int(a[0:2], 16), int(a[2:4], 16), int(a[4:6], 16)
        br, bg, bb = int(b[0:2], 16), int(b[2:4], 16), int(b[4:6], 16)
        r = int(ar + (br - ar) * t)
        g = int(ag + (bg - ag) * t)
        bl = int(ab + (bb - ab) * t)
        return f"#{r:02x}{g:02x}{bl:02x}"

    def _tint(self, hex_color, bg="#162033", alpha=0.16):
        """
        Blends hex_color toward the card background by alpha. Canvas
        polygons don't support real alpha compositing, so a translucent
        fill under the sparkline is simulated by blending toward the
        known background color instead.
        """
        hex_color = hex_color.lstrip("#")
        bg = bg.lstrip("#")
        r1, g1, b1 = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        r2, g2, b2 = int(bg[0:2], 16), int(bg[2:4], 16), int(bg[4:6], 16)
        r = int(r1 * alpha + r2 * (1 - alpha))
        g = int(g1 * alpha + g2 * (1 - alpha))
        b = int(b1 * alpha + b2 * (1 - alpha))
        return f"#{r:02x}{g:02x}{b:02x}"

    def _draw_sidebar_chart(self, values):

        self.sidebar_chart.delete("all")

        w, h, pad_x, pad_y = 160, 46, 4, 8

        if not values:
            self.sidebar_chart.create_text(
                w / 2, h / 2, text="No data yet", fill=COLOR_MUTED, font=(FONT_FAMILY, 10)
            )
            return

        vals = [float(v) for v in values[-7:]]
        color = self._get_doi_status_color(vals[-1])

        if len(vals) == 1:
            # A single sample isn't a trend -- show a flat dashed guide
            # rather than implying a shape that isn't really there.
            y = h / 2
            self.sidebar_chart.create_line(pad_x, y, w - pad_x, y, fill=COLOR_MUTED, width=1, dash=(3, 2))
            self.sidebar_chart.create_oval(w - pad_x - 4, y - 4, w - pad_x + 4, y + 4, fill=color, outline="")
            return

        lo, hi = min(vals), max(vals)
        if hi == lo:
            hi = lo + 1

        step = (w - 2 * pad_x) / (len(vals) - 1)
        pts = []
        for i, v in enumerate(vals):
            x = pad_x + i * step
            y = pad_y + (1 - (v - lo) / (hi - lo)) * (h - 2 * pad_y)
            pts.extend((x, y))

        # subtle tinted area under the line, drawn first so the line sits on top
        fill_pts = pts + [pts[-2], h - 1, pts[0], h - 1]
        self.sidebar_chart.create_polygon(fill_pts, fill=self._tint(color), outline="")

        self.sidebar_chart.create_line(*pts, fill=color, width=2, smooth=True, capstyle="round")

        x, y = pts[-2], pts[-1]
        self.sidebar_chart.create_oval(x - 3.5, y - 3.5, x + 3.5, y + 3.5, fill=color, outline="")

    def _update_sidebar_chart(self):

        history = self._load_doi_history()
        product = self.product.get() if hasattr(self, "product") else config.PRODUCT_NAME
        raw = history.get(product, [])

        values = []
        for item in raw[-7:]:
            if isinstance(item, dict):
                values.append(item.get("value", 0))
            else:
                values.append(item)

        self._draw_sidebar_chart(values)

        if not values:
            self.sidebar_value_label.configure(text="--", text_color=COLOR_MUTED)
            self.sidebar_delta_label.configure(text="")
            self.sidebar_last_update.configure(text="No history yet")
            return

        current = values[-1]
        color = self._get_doi_status_color(current)
        self.sidebar_value_label.configure(text=f"{current:.1f}", text_color=color)

        if len(values) >= 2:
            delta = values[-1] - values[-2]
            if delta > 0.05:
                self.sidebar_delta_label.configure(text=f"\u25b2 {delta:.1f}", text_color=COLOR_MUTED)
            elif delta < -0.05:
                self.sidebar_delta_label.configure(text=f"\u25bc {abs(delta):.1f}", text_color=COLOR_MUTED)
            else:
                self.sidebar_delta_label.configure(text="\u2013 flat", text_color=COLOR_MUTED)
        else:
            self.sidebar_delta_label.configure(text="")

        reading_word = "reading" if len(values) == 1 else "readings"
        self.sidebar_last_update.configure(text=f"{len(values)} {reading_word} \u00b7 {product}")

    def _record_doi_history(self, doi_m6, doi_g12):
        """
        Appends today's reading for both products to the small local
        history file that powers this chart. Keeps the most recent 60
        entries per product -- comfortably more than a 7-day view needs
        even with several checks a day, without the file growing
        without bound.
        """
        history = self._load_doi_history()
        now_iso = datetime.now().isoformat(timespec="seconds")

        history.setdefault("M6", []).append({"value": doi_m6, "time": now_iso})
        history.setdefault("G12", []).append({"value": doi_g12, "time": now_iso})

        for product in ("M6", "G12"):
            history[product] = history[product][-60:]

        self._save_doi_history(history)

    # ==========================================================
    # PRODUCT CHANGE
    # ==========================================================

    def change_product(self, product):

        config.set_product(product)

        self.material_value.configure(text=config.SAP_MATERIAL)
        self.sheet_value.configure(text=config.SHEET_NAME)
        self.description_label.configure(
            text=f"Description : {config.PRODUCT_DESCRIPTION}"
        )

        self._update_sidebar_chart()
        self.log(f"Product Changed -> {config.PRODUCT_NAME}", level="info", module="System")

    def check_connection(self):

        self.log("[INFO] Checking Connection...", level="info", module="Connection")

        # Routed through sap_manager now instead of a second, separate
        # win32com implementation living here -- one place that knows
        # how to answer "is SAP actually running."
        try:
            from SAP import sap_manager
            sap_ok = sap_manager.is_sap_running()
        except Exception:
            sap_ok = False

        excel_ok = False

        # Check Excel
        try:

            import win32com.client

            win32com.client.GetActiveObject("Excel.Application")

            excel_ok = True

        except:
            pass

        if sap_ok and excel_ok:

            self._set_connection_indicators(ok=True)

            self.log("[OK] Connection Verified.", level="ok", module="Connection")

        else:

            self.sap_status.configure(
                text=f"SAP : {'Connected' if sap_ok else 'Disconnected'}",
                text_color=COLOR_OK if sap_ok else COLOR_ERROR
            )

            self.excel_status.configure(
                text=f"Excel : {'Connected' if excel_ok else 'Disconnected'}",
                text_color=COLOR_OK if excel_ok else COLOR_ERROR
            )

            self.log("[WARN] One or more connections are unavailable.", level="warn", module="Connection")

        self._refresh_sap_sessions()



    def open_create_po(self):

        if self._po_dialog_open or self._po_running:
            return

        self._po_dialog_open = True
        self.create_po_btn.configure(state="disabled")

        try:

            config.set_product(self.product.get())

            excel = ExcelManager()
            excel.connect()
            excel.open_workbook()

            batches = excel.get_po_batches()

            dialog = CreatePODialog(self, batches, config.PRODUCT_NAME)

            self.wait_window(dialog)

            if dialog.result is None:
                return

            if dialog.result["mode"] == "all":

                pending = [b for b in batches if b.get("po", "") == ""]
                self._run_bulk_po(excel, pending, dialog.result)

            else:

                self._create_single_po(excel, dialog.result)

        finally:

            self._po_dialog_open = False

            # If a creation flow actually started, it owns re-enabling the
            # button when it finishes. If the dialog was just cancelled (or
            # never got that far), nothing else will re-enable it, so do
            # it here.
            if not self._po_running:
                self.create_po_btn.configure(state="normal")

    # ----------------------------------------------------------
    # SINGLE PO
    # ----------------------------------------------------------

    def _create_single_po(self, excel, data):

        if self._po_running:
            return
        self._po_running = True

        self._set_buttons_state("disabled")
        self.create_po_btn.configure(text="Creating...")
        self.current_batch_label.configure(text=f"Batch: {data['batch']}")

        self._set_running("Creating Production Order...")

        def task():

            po = ProductionOrder()
            po_number = po.create_production_order(data)
            excel.update_po(data["row"], po_number, data["start_date"])

            return po_number

        def on_done(output, po_number, error):

            self.log(output, module="Create PO")
            self.current_batch_label.configure(text="")

            if error:

                self.log(f"[ERROR] {error}", level="error", module="Create PO")
                logger.log_failure(data["batch"], str(error))

                self._set_finished("Error", COLOR_ERROR, 0)
                self._show_smart_error(str(error))

            else:

                self.last_created_order = {
                    "batch": data["batch"],
                    "po": po_number,
                    "row": data["row"]
                }

                self.log(f"[OK] Production Order Created : {po_number}", level="ok", module="Create PO")
                logger.log_success(data["batch"], po_number, data["qty"])

                self._set_finished("Completed", COLOR_OK, 1)
                messagebox.showinfo("Completed", f"Production Order Created\n\nPO : {po_number}")

            self._po_running = False
            self.create_po_btn.configure(text="  CREATE PRODUCTION ORDER")
            self._set_buttons_state("normal")
            self._refresh_po_dashboard(excel)

        self._run_task(task, on_done)

    # ----------------------------------------------------------
    # BULK PO ("Create ALL Pending PO")
    # ----------------------------------------------------------

    def _run_bulk_po(self, excel, pending, shared_data):

        if not pending:
            messagebox.showinfo("Information", "No pending batch to create.")
            return

        if self._po_running:
            return
        self._po_running = True

        self._set_buttons_state("disabled")
        self.create_po_btn.configure(text="Creating...")

        total = len(pending)

        self.current_batch_label.configure(text=f"Batch: {pending[0]['batch']}")
        self.status_value.configure(text=f"Creating PO (0/{total})...", text_color=COLOR_WARN)
        self._start_pulse([self.status_value])

        self.progress.configure(mode="determinate")
        self.progress.set(0)
        self.progress_percent.configure(text=f"0/{total}")

        def worker():

            po = ProductionOrder()
            results = []
            start_time = time.time()

            for i, item in enumerate(pending, start=1):

                self.after(0, lambda i=i, item=item: self._update_bulk_po_progress(i, total, item["batch"]))

                output = io.StringIO()
                sys.stdout.register(output)

                try:

                    data = {
                        "batch": item["batch"],
                        "qty": item["qty"],
                        "row": item["row"],
                        "plant": shared_data["plant"],
                        "material": shared_data["material"],
                        "start_date": shared_data["start_date"],
                        "finish_date": shared_data["finish_date"],
                    }

                    po_number = po.create_production_order(data)
                    excel.update_po(item["row"], po_number, shared_data["start_date"])

                    logger.log_success(item["batch"], po_number, item["qty"])
                    results.append({"batch": item["batch"], "status": "success", "po": po_number})

                except Exception as e:

                    logger.log_failure(item["batch"], str(e))
                    results.append({"batch": item["batch"], "status": "failed", "reason": str(e)})

                    batch_name = item["batch"]
                    error_text = str(e)
                    self.after(
                        0,
                        lambda batch_name=batch_name, error_text=error_text: self.log(
                            f"[ERROR] {batch_name} failed: {error_text}", level="error", module="Create PO"
                        ),
                    )

                finally:

                    sys.stdout.unregister()
                    text = output.getvalue()
                    self.after(0, lambda text=text: self.log(text, module="Create PO"))

            elapsed = time.time() - start_time

            self.after(0, lambda: self._on_bulk_po_done(results, elapsed, excel, shared_data))

        threading.Thread(target=worker, daemon=True).start()

    def _update_bulk_po_progress(self, i, total, batch):

        self.current_batch_label.configure(text=f"Batch: {batch}")
        self.status_value.configure(text=f"Creating PO ({i}/{total})...", text_color=COLOR_WARN)
        self.progress.set(i / total)
        self.progress_percent.configure(text=f"{i}/{total}")

    def _on_bulk_po_done(self, results, elapsed, excel, shared_data):

        self._stop_pulse([self.status_value])

        success_count = sum(1 for r in results if r["status"] == "success")
        failed_count = sum(1 for r in results if r["status"] == "failed")

        self.current_batch_label.configure(text="")

        if failed_count == 0:
            self._set_finished("Completed", COLOR_OK, 1)
        else:
            self._set_finished("Completed with errors", COLOR_WARN, 1)

        self._po_running = False
        self.create_po_btn.configure(text="  CREATE PRODUCTION ORDER")
        self._set_buttons_state("normal")

        self._refresh_po_dashboard(excel)

        failed_batches = {r["batch"] for r in results if r["status"] == "failed"}

        def retry():

            fresh = excel.get_po_batches()
            retry_pending = [b for b in fresh if b["batch"] in failed_batches]

            if not retry_pending:
                messagebox.showinfo("Information", "No failed batch left to retry.")
                return

            self._run_bulk_po(excel, retry_pending, shared_data)

        POSummaryDialog(
            self,
            success_count=success_count,
            failed_count=failed_count,
            elapsed_seconds=elapsed,
            log_path=logger.today_log_path(),
            on_retry=retry if failed_batches else None,
        )



    def update_production_order(self):

        if self.last_created_order is None:
            messagebox.showwarning(
                "Information",
                "No newly created Production Order found."
            )
            return

        self._set_buttons_state("disabled")
        self._set_running("Updating Production Order (CO02)...")

        def task():

            from SAP.co02 import run_co02

            run_co02(self.last_created_order["po"])

            return self.last_created_order["po"]

        def on_done(output, result, error):

            self.log(output, module="Update PO")

            if error:

                self.log(f"[ERROR] {error}", level="error", module="Update PO")

                self._set_finished("Error", COLOR_ERROR, 0)

                self._show_smart_error(str(error))

            else:

                self.log(
                    f"[OK] Production Order {result} Updated.",
                    level="ok",
                    module="Update PO",
                )

                self._set_finished("Completed", COLOR_OK, 1)

                self.last_created_order = None

                messagebox.showinfo(
                    "Completed",
                    f"Production Order Updated Successfully\n\nPO : {result}"
                )

            self._set_buttons_state("normal")

        self._run_task(task, on_done)


    def production_order_list(self):

        self._set_buttons_state("disabled")

        self._set_running("Opening Production Order List (COOIS)...")

        def task():

            from SAP.coois import run_coois

            return run_coois()

        def on_done(output, result, error):

            self.log(output, module="PO List")

            if error:

                self.log(f"[ERROR] {error}", level="error", module="PO List")

                self._set_finished("Error", COLOR_ERROR, 0)

                self._show_smart_error(str(error))

            else:

                self.log("[OK] Production Order List Opened.", level="ok", module="PO List")

                self._set_finished("Completed", COLOR_OK, 1)

            self._set_buttons_state("normal")

        self._run_task(task, on_done)

    # ==========================================================
    # PO TALLY
    # ==========================================================

    def open_po_tally(self):
        """
        Reads Production Orders from the currently-open COOIS session
        (never re-executes COOIS itself), then shows the selection
        dialog. COOIS must already be open -- SAP.coois.read_production_orders()
        raises a clear error otherwise, surfaced via the Smart Error
        dialog rather than a cryptic SAP-level failure.
        """

        self._set_buttons_state("disabled")

        self._set_running("Reading COOIS...")

        def task():
            from SAP.coois import read_production_orders
            return read_production_orders()

        def on_read_done(output, orders, error):

            self.log(output, module="Variance Check")

            self._set_buttons_state("normal")

            if error:
                self.log(f"[ERROR] {error}", level="error", module="Variance Check")
                self._set_finished("Error", COLOR_ERROR, 0)
                self._show_smart_error(str(error))
                return

            self._set_finished("Completed", COLOR_OK, 1)

            dlg = POTallySelectionDialog(self, orders or [])
            self.wait_window(dlg)

            if dlg.result:
                self._run_po_tally_batch(dlg.result)

        self._run_task(task, on_read_done)

    def _run_po_tally_batch(self, selected_orders):
        """
        Runs ZPPMYR0520 reconciliation for each selected order, on one
        dedicated SAP window reused across the whole batch (opened
        once, not reopened per order), with live progress shown in a
        dedicated dialog -- updated via self.after(0, ...) from the
        background worker thread, the same thread-safe scheduling
        _run_task() itself already relies on elsewhere in this app.
        """

        progress_dlg = POTallyProgressDialog(self, total=len(selected_orders))

        self._set_buttons_state("disabled")

        def task():

            from SAP.zppmyr0520 import POTallyReader
            from SAP import mb52_bib
            from Excel import po_tally_excel

            # Look up each PO's COOIS-sourced fields by PO number, not
            # by position -- tally_batch() reads the actual order
            # number back from SAP for each row rather than assuming
            # row order matches input order, so results need to be
            # matched back the same way.
            coois_fields_by_po = {
                order["po"]: {
                    "material": order.get("material", ""),
                    "batch": order.get("batch", ""),
                }
                for order in selected_orders
            }

            # BIB: one MB52 query covering both products, read once for
            # the whole batch -- not re-run per order. Each order's BIB
            # is then a lookup against this same warehouse snapshot,
            # matched by its own batch's last 4 characters (see
            # mb52_bib.sum_unrestricted_for_batch()). Wrapped so a
            # failure here (SAP hiccup, unexpected screen state) still
            # lets the rest of the tally complete -- BIB just comes
            # back as 0.0 for this run rather than blocking Target /
            # Yield / Scrap / Posting Date / Status for every order.
            self.after(0, lambda: progress_dlg.update_progress(0, "", "Reading warehouse stock (BIB)..."))
            try:
                mb52_bib.run_mb52_for_bib()
                warehouse_rows = mb52_bib.read_all_batches()
            except Exception as e:
                print(f"[ERROR] BIB warehouse read failed, continuing without it: {e}")
                warehouse_rows = []

            reader = POTallyReader()
            reader.open_transaction()

            excel = po_tally_excel.connect()
            workbook = po_tally_excel.open_workbook(excel)
            sheet = po_tally_excel.get_or_create_sheet(workbook)

            def on_row_start(row_index, row_count, po):
                self.after(0, lambda: progress_dlg.update_progress(row_index, po, "Reading SAP..."))

            def on_row_complete(row_index, row_count, result):
                result.update(coois_fields_by_po.get(result["po"], {}))
                result["bib"] = mb52_bib.sum_unrestricted_for_batch(warehouse_rows, result.get("batch", ""))
                po_tally_excel.upsert_result(sheet, result)
                self.after(0, lambda: progress_dlg.update_progress(row_index + 1, result["po"], "Updating Excel..."))

            po_numbers = [order["po"] for order in selected_orders]

            results = reader.tally_batch(
                po_numbers,
                plant=config.SAP_PLANT,
                on_row_start=on_row_start,
                on_row_complete=on_row_complete,
                should_cancel=lambda: progress_dlg.cancel_requested,
            )

            workbook.Save()

            failed = len(selected_orders) - len(results)

            return {"results": results, "failed": failed}

        def on_done(output, data, error):

            self.log(output, module="Variance Check")

            progress_dlg.destroy()

            self._set_buttons_state("normal")

            if error:
                self.log(f"[ERROR] {error}", level="error", module="Variance Check")
                self._set_finished("Error", COLOR_ERROR, 0)
                self._show_smart_error(str(error))
                return

            results = data["results"]
            failed = data["failed"]

            match_count = sum(1 for r in results if r["status"] == config.PO_TALLY_STATUS_MATCH)
            less_count = sum(1 for r in results if r["status"] == config.PO_TALLY_STATUS_LESS)
            over_count = sum(1 for r in results if r["status"] == config.PO_TALLY_STATUS_OVER)

            self._set_finished("Completed", COLOR_OK, 1)
            self._refresh_sap_sessions()

            POTallyCompletionDialog(
                self,
                checked_count=len(results),
                match_count=match_count,
                less_count=less_count,
                over_count=over_count,
                failed_count=failed,
            )

        self._run_task(task, on_done)

    # ==========================================================
    # CHECK MB52
    # ==========================================================

    def check_mb52(self):

        self._set_buttons_state("disabled")
        self.clear_log()

        config.set_product(self.product.get())

        self._set_running("Checking SAP (MB52)...")
        self._set_connection_indicators(ok=True, busy_text="Reading...")

        self._run_task(run_automation, self._on_check_mb52_done)

    def _on_check_mb52_done(self, output, new_batches, error):

        self.log(output, module="MB52")

        if error:

            self.log(f"[ERROR] {error}", level="error", module="MB52")
            self._set_finished("Error", COLOR_ERROR, 0)
            self._set_connection_indicators(ok=False)

        else:

            self.new_batches = new_batches or []
            count = len(self.new_batches)

            self.batch_count.configure(
                text=f"New Batch Found : {count}",
                text_color=(COLOR_OK if count else COLOR_MUTED),
            )

            self._set_connection_indicators(ok=True)
            self._set_finished("Completed", COLOR_OK, 1)

        self._set_buttons_state("normal")

    # ==========================================================
    # UPDATE BATCH & QUANTITY
    # ==========================================================

    def update_batch_quantity(self):

        if not self.new_batches:
            messagebox.showwarning("Information", "No New Batch.")
            return

        self._set_buttons_state("disabled")
        self._set_running("Updating Excel...")

        pending = list(self.new_batches)

        def task():
            excel = ExcelManager()
            excel.connect()
            excel.open_workbook()
            excel.write_batches(pending)
            return len(pending)

        self._run_task(task, self._on_update_batch_done)

    def _on_update_batch_done(self, output, count, error):

        self.log(output, module="MB52")

        if error:

            self.log(f"[ERROR] {error}", level="error", module="MB52")
            self._set_finished("Error", COLOR_ERROR, 0)
            self._show_smart_error(str(error))

        else:

            self.log(f"[OK] {count} Batch Updated.", level="ok", module="MB52")
            messagebox.showinfo("Completed", f"{count} Batch Updated.")

            self.new_batches = []
            self.batch_count.configure(
                text="New Batch Found : 0", text_color=COLOR_MUTED
            )
            self._set_finished("Completed", COLOR_OK, 1)

        self._set_buttons_state("normal")

    # ==========================================================
    # UPDATE SHIPMENT NUMBER
    # ==========================================================

    def update_shipment_number(self):

        # Make sure the active product is applied even if the user
        # switched products without re-running CHECK MB52 first.
        config.set_product(self.product.get())

        self._set_buttons_state("disabled")
        self._set_running("Updating Shipment (SAP)...")
        self._set_connection_indicators(ok=True, busy_text="Reading...")

        self._run_task(run_shipment_automation, self._on_shipment_done)

    def _on_shipment_done(self, output, shipment, error):

        self.log(output, module="Shipment")

        if error:

            self.log(f"[ERROR] {error}", level="error", module="Shipment")
            self._set_finished("Error", COLOR_ERROR, 0)
            self._set_connection_indicators(ok=False)
            self._show_smart_error(str(error))

        else:

            count = len(shipment) if shipment else 0

            if count == 0:
                self.log("[INFO] No batch pending shipment update.", level="info", module="Shipment")
            else:
                self.log(f"[OK] {count} Shipment Updated.", level="ok", module="Shipment")

            self._set_connection_indicators(ok=True)
            self._set_finished("Completed", COLOR_OK, 1)

        self._set_buttons_state("normal")

    # ==========================================================
    # UPDATE GR DATE
    # ==========================================================

    def update_gr_date(self):

        # Make sure the active product is applied even if the user
        # switched products without re-running CHECK MB52 first.
        config.set_product(self.product.get())

        self._set_buttons_state("disabled")
        self._set_running("Updating GR Date (SAP)...")
        self._set_connection_indicators(ok=True, busy_text="Reading...")

        self._run_task(run_mb51_automation, self._on_gr_date_done)





    def _on_gr_date_done(self, output, gr_date, error):

        self.log(output, module="GR Date")

        if error:

            self.log(f"[ERROR] {error}", level="error", module="GR Date")
            self._set_finished("Error", COLOR_ERROR, 0)
            self._set_connection_indicators(ok=False)
            self._show_smart_error(str(error))

        else:

            count = len(gr_date) if gr_date else 0

            if count == 0:
                self.log("[INFO] No batch pending GR date update.", level="info", module="GR Date")
            else:
                self.log(f"[OK] {count} GR Date Updated.", level="ok", module="GR Date")

            self._set_connection_indicators(ok=True)
            self._set_finished("Completed", COLOR_OK, 1)

        self._set_buttons_state("normal")

    def check_doi(self):

        self.doi_button.configure(text="Loading...")

        self._set_buttons_state("disabled")

        self._set_running("Checking DOI...")
        self._set_connection_indicators(ok=True, busy_text="Reading...", targets=("sap",))

        self._run_task(run_doi, self._on_doi_done)

    def _get_doi_status_color(self, doi):

        if doi > 5:
            return COLOR_OK
        elif doi >= 3:
            return COLOR_WARN
        else:
            return COLOR_ERROR

    def _on_doi_done(self, output, result, error):

        self.log(output, module="DOI")

        if error:

            self.doi_button.configure(text=" Refresh DOI")

            self.doi_m6_indicator.configure(text="M6 : Error", text_color=COLOR_ERROR)
            self.doi_g12_indicator.configure(text="G12 : Error", text_color=COLOR_ERROR)

            self.log(f"[ERROR] {error}", level="error", module="DOI")

            self._set_finished("Error", COLOR_ERROR, 0)
            self._set_connection_indicators(ok=False, targets=("sap",))

        else:

            data = result
            doi_m6 = data['m6']['doi']
            doi_g12 = data['g12']['doi']
            m6_until = data['m6']['until']
            g12_until = data['g12']['until']

            # Use doi.py's own last_updated timestamp -- the actual
            # moment the calculation finished -- rather than calling
            # datetime.now() again here, which would reflect whenever
            # this callback happens to run instead of when the SAP data
            # was actually read.
            last_updated = data.get('last_updated') or datetime.now()
            now = last_updated.strftime("%d %b %Y %I:%M %p")

            # Each product gets its own dot -- M6 running low shouldn't be
            # masked by G12 being fine, or vice versa.
            self.doi_m6_indicator.configure(
                text=f"M6 : {doi_m6:.1f} Days",
                text_color=self._get_doi_status_color(doi_m6),
            )
            self.doi_g12_indicator.configure(
                text=f"G12 : {doi_g12:.1f} Days",
                text_color=self._get_doi_status_color(doi_g12),
            )
            self.doi_updated_label.configure(text=f"Last Updated : {now}")
            self.m6_until_label.configure(text=f" Until {m6_until.strftime('%d %b %Y (%a)')}")
            self.m6_time_label.configure(text=f" {m6_until.strftime('%I:%M %p')}")
            self.g12_until_label.configure(text=f" Until {g12_until.strftime('%d %b %Y (%a)')}")
            self.g12_time_label.configure(text=f" {g12_until.strftime('%I:%M %p')}")

            self.doi_button.configure(text=" Refresh DOI")

            self._record_doi_history(doi_m6, doi_g12)
            self._update_sidebar_chart()

            self.log("[OK] DOI Updated.", level="ok", module="DOI")

            self._set_finished("Completed", COLOR_OK, 1)
            self._set_connection_indicators(ok=True, targets=("sap",))

        self._set_buttons_state("normal")






# ==============================================================
# ENTRY POINT
# ==============================================================

def main():
    welcome = WelcomeScreen()
    welcome.mainloop()  # returns once the welcome screen destroys itself

    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
