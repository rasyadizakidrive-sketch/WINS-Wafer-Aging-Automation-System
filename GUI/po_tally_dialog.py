import os

import customtkinter as ctk
from tkinter import messagebox

import config
import icons

# ==============================================================
# DESIGN TOKENS -- matching create_po_dialog.py exactly, so this
# feels like part of the same app, not a bolted-on extra window.
# ==============================================================

COLOR_BG = "#0B1220"
COLOR_BG_HEADER = "#101826"
COLOR_CARD = "#1B2434"
COLOR_CARD_BORDER = "#2A3547"
COLOR_TEXT = "#F1F5F9"
COLOR_MUTED = "#94A3B8"

COLOR_PRIMARY = "#2563EB"
COLOR_PRIMARY_HOVER = "#1D4ED8"

COLOR_SUCCESS = "#22C55E"
COLOR_WARN = "#F59E0B"
COLOR_ERROR = "#EF4444"

FONT_FAMILY = "Segoe UI"


def _center_dialog(win, natural_w, natural_h, min_w=420, min_h=340):
    """Screen-aware sizing/centering, matching the same pattern already
    proven in create_po_dialog.py -- a fixed size can't assume every
    screen has enough vertical room above the taskbar."""

    win.update_idletasks()
    screen_w = win.winfo_screenwidth()
    screen_h = win.winfo_screenheight()
    max_h = max(min_h, screen_h - 140)
    dialog_w = min(natural_w, screen_w - 40)
    dialog_h = min(natural_h, max_h)
    x = max(0, (screen_w - dialog_w) // 2)
    y = max(0, (screen_h - dialog_h) // 2)
    win.geometry(f"{dialog_w}x{dialog_h}+{x}+{y}")
    win.minsize(min_w, min_h)
    win.resizable(True, True)


# ==============================================================
# SELECTION DIALOG
# ==============================================================

class POTallySelectionDialog(ctk.CTkToplevel):
    """
    Shows every Production Order read from the currently-open COOIS
    session, with a live search box and multi-select checkboxes.
    self.result is the list of selected order dicts if the user pressed
    "Run Tally", or None if they cancelled.
    """

    def __init__(self, parent, orders):

        super().__init__(parent)

        self.configure(fg_color=COLOR_BG)
        self.title("Production Order Variance Check")
        self.grab_set()

        self.result = None
        self.orders = orders

        # po -> (CTkCheckBox, BooleanVar, order_dict, row_frame)
        self._rows = {}

        _center_dialog(self, 480, 560, min_w=380, min_h=340)

        self._build_ui()

    def _build_ui(self):

        header = ctk.CTkFrame(self, fg_color=COLOR_BG_HEADER, corner_radius=0)
        header.pack(fill="x")

        ctk.CTkLabel(
            header, text="Production Order Variance Check",
            font=(FONT_FAMILY, 16, "bold"), text_color=COLOR_TEXT,
        ).pack(anchor="w", padx=22, pady=(11, 1))

        ctk.CTkLabel(
            header, text=f"SAP TCODE : {config.PO_TALLY_TCODE}  \u00b7  from current COOIS list",
            font=(FONT_FAMILY, 11), text_color=COLOR_MUTED,
        ).pack(anchor="w", padx=22, pady=(0, 10))

        # ---- Footer built and packed first, same reasoning as
        # create_po_dialog.py: the buttons must never lose the fight
        # for space to the scrollable list above them. ----
        footer = ctk.CTkFrame(self, fg_color=COLOR_BG_HEADER, corner_radius=0)
        footer.pack(fill="x", side="bottom")

        self.selected_label = ctk.CTkLabel(
            footer, text="Selected : 0", font=(FONT_FAMILY, 11.5), text_color=COLOR_MUTED,
        )
        self.selected_label.pack(anchor="w", padx=16, pady=(10, 0))

        btn_row = ctk.CTkFrame(footer, fg_color="transparent")
        btn_row.pack(fill="x", padx=16, pady=(6, 14))

        ctk.CTkButton(
            btn_row, text="Cancel", command=self.destroy, width=100,
            fg_color="#334155", hover_color="#475569",
            font=(FONT_FAMILY, 12), cursor="hand2",
        ).pack(side="right")

        ctk.CTkButton(
            btn_row, text="  Run Tally", image=icons.make_icon("check", color="#FFFFFF", size=14),
            compound="left", command=self._run_tally, width=140,
            fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER,
            font=(FONT_FAMILY, 12, "bold"), cursor="hand2",
        ).pack(side="right", padx=(0, 10))

        self.select_all_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            btn_row, text="Select All", variable=self.select_all_var,
            command=self._toggle_select_all,
            font=(FONT_FAMILY, 11.5), text_color=COLOR_TEXT,
        ).pack(side="left")

        # ---- Search ----
        search_wrap = ctk.CTkFrame(self, fg_color="transparent")
        search_wrap.pack(fill="x", padx=16, pady=(12, 8))

        ctk.CTkLabel(
            search_wrap, text="Search", font=(FONT_FAMILY, 11), text_color=COLOR_MUTED,
        ).pack(anchor="w", pady=(0, 4))

        self.search_entry = ctk.CTkEntry(
            search_wrap, placeholder_text="Order number or batch...",
        )
        self.search_entry.pack(fill="x")
        self.search_entry.bind("<KeyRelease>", self._on_search)

        # ---- Scrollable order list ----
        self.list_frame = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.list_frame.pack(fill="both", expand=True, padx=16, pady=(0, 4))

        if not self.orders:
            ctk.CTkLabel(
                self.list_frame, text="No Production Orders found in the current COOIS list.",
                font=(FONT_FAMILY, 12), text_color=COLOR_MUTED,
            ).pack(pady=30)
            return

        for order in self.orders:
            self._add_row(order)

    def _add_row(self, order):

        row = ctk.CTkFrame(
            self.list_frame, fg_color=COLOR_CARD, border_color=COLOR_CARD_BORDER,
            border_width=1, corner_radius=8,
        )
        row.pack(fill="x", pady=3)

        var = ctk.BooleanVar(value=False)

        batch_text = f"Batch {order['batch']}" if order.get("batch") else "No batch yet"

        chk = ctk.CTkCheckBox(
            row, text="", variable=var, width=20, command=self._update_selected_count,
        )
        chk.pack(side="left", padx=(12, 4), pady=10)

        ctk.CTkLabel(
            row, text=order["po"], font=(FONT_FAMILY, 13, "bold"), text_color=COLOR_TEXT,
        ).pack(side="left", padx=(4, 12), pady=10)

        ctk.CTkLabel(
            row, text=batch_text, font=(FONT_FAMILY, 11.5), text_color=COLOR_MUTED,
        ).pack(side="left", pady=10)

        self._rows[order["po"]] = (chk, var, order, row)

    # ----------------------------------------------------------
    def _on_search(self, event=None):

        term = self.search_entry.get().strip().lower()

        for po, (chk, var, order, row) in self._rows.items():

            haystack = f"{order['po']} {order.get('batch', '')} {order.get('material', '')}".lower()

            if term in haystack:
                row.pack(fill="x", pady=3)
            else:
                row.pack_forget()

    def _toggle_select_all(self):

        select = self.select_all_var.get()

        for po, (chk, var, order, row) in self._rows.items():
            # Only affects rows currently visible (matching the
            # filtered search results), not everything hidden by search.
            if row.winfo_ismapped():
                var.set(select)

        self._update_selected_count()

    def _update_selected_count(self):

        count = sum(1 for chk, var, order, row in self._rows.values() if var.get())

        self.selected_label.configure(text=f"Selected : {count}")

    def _run_tally(self):

        selected = [order for chk, var, order, row in self._rows.values() if var.get()]

        if not selected:
            messagebox.showwarning("No Selection", "Select at least one Production Order.")
            return

        self.result = selected

        self.destroy()


# ==============================================================
# PROGRESS DIALOG
# ==============================================================

class POTallyProgressDialog(ctk.CTkToplevel):
    """
    Live progress feedback while the batch runs -- update_progress() is
    called from the background worker thread via self.after(0, ...),
    the same thread-safe scheduling pattern _run_task() itself already
    uses elsewhere in this app.
    """

    def __init__(self, parent, total):

        super().__init__(parent)

        self.configure(fg_color=COLOR_BG)
        self.title("Running Variance Check")
        self.protocol("WM_DELETE_WINDOW", self._request_cancel)
        self.grab_set()

        self.total = max(total, 1)
        self.cancel_requested = False

        _center_dialog(self, 380, 300, min_w=320, min_h=260)

        header = ctk.CTkFrame(self, fg_color=COLOR_BG_HEADER, corner_radius=0)
        header.pack(fill="x")

        ctk.CTkLabel(
            header, text="Running Variance Check", font=(FONT_FAMILY, 15, "bold"), text_color=COLOR_TEXT,
        ).pack(anchor="w", padx=20, pady=14)

        card = ctk.CTkFrame(
            self, fg_color=COLOR_CARD, border_color=COLOR_CARD_BORDER, border_width=1, corner_radius=12
        )
        card.pack(fill="both", expand=True, padx=18, pady=(18, 8))

        self.progress_bar = ctk.CTkProgressBar(card, height=10)
        self.progress_bar.pack(fill="x", padx=18, pady=(20, 6))
        self.progress_bar.set(0)

        self.count_label = ctk.CTkLabel(
            card, text=f"0 / {self.total}", font=(FONT_FAMILY, 13, "bold"), text_color=COLOR_TEXT,
        )
        self.count_label.pack(anchor="w", padx=18, pady=(0, 14))

        ctk.CTkLabel(
            card, text="Current PO", font=(FONT_FAMILY, 10.5), text_color=COLOR_MUTED,
        ).pack(anchor="w", padx=18)
        self.po_label = ctk.CTkLabel(
            card, text="-", font=(FONT_FAMILY, 13, "bold"), text_color=COLOR_TEXT,
        )
        self.po_label.pack(anchor="w", padx=18, pady=(0, 12))

        self.stage_label = ctk.CTkLabel(
            card, text="Starting...", font=(FONT_FAMILY, 11.5), text_color=COLOR_MUTED,
        )
        self.stage_label.pack(anchor="w", padx=18, pady=(0, 16))

        # A way out if something hangs -- SAP-side calls all have their
        # own 10-second wait_until_ready() timeout, but the clipboard
        # read in zppmyr0520.py doesn't, and could genuinely hang (a
        # locked clipboard on Windows, for instance). Cooperative, not
        # forced: the batch loop checks this flag between orders and
        # stops there, keeping whatever results were already collected
        # rather than discarding them.
        self.cancel_btn = ctk.CTkButton(
            self, text="Cancel", command=self._request_cancel, width=110,
            fg_color="#334155", hover_color="#475569",
            font=(FONT_FAMILY, 12), cursor="hand2",
        )
        self.cancel_btn.pack(pady=(0, 14))

    def _request_cancel(self):

        self.cancel_requested = True
        self.cancel_btn.configure(state="disabled", text="Cancelling...")
        self.stage_label.configure(text="Stopping after the current order finishes...")

    def update_progress(self, current, po_number, stage_text):

        try:
            self.progress_bar.set(current / self.total)
            self.count_label.configure(text=f"{current} / {self.total}")
            self.po_label.configure(text=po_number)
            self.stage_label.configure(text=stage_text)
        except Exception:
            pass  # window may have already been closed


# ==============================================================
# COMPLETION DIALOG
# ==============================================================

class POTallyCompletionDialog(ctk.CTkToplevel):

    def __init__(self, parent, checked_count, match_count, less_count, over_count, failed_count=0):

        super().__init__(parent)

        self.configure(fg_color=COLOR_BG)
        self.title("Variance Check Completed")
        self.grab_set()

        _center_dialog(self, 380, 420, min_w=320, min_h=340)

        header = ctk.CTkFrame(self, fg_color=COLOR_BG_HEADER, corner_radius=0)
        header.pack(fill="x")

        icon_kind = "check" if failed_count == 0 else "warning"
        icon_color = COLOR_SUCCESS if failed_count == 0 else COLOR_WARN
        ctk.CTkLabel(
            header, text="  Variance Check Completed",
            image=icons.make_icon(icon_kind, color=icon_color, size=18), compound="left",
            font=(FONT_FAMILY, 16, "bold"), text_color=COLOR_TEXT,
        ).pack(anchor="w", padx=22, pady=16)

        card = ctk.CTkFrame(
            self, fg_color=COLOR_CARD, border_color=COLOR_CARD_BORDER, border_width=1, corner_radius=12
        )
        card.pack(fill="both", expand=True, padx=20, pady=20)

        def stat_row(label, value, color=COLOR_TEXT):
            row = ctk.CTkFrame(card, fg_color="transparent")
            row.pack(fill="x", padx=20, pady=7)
            ctk.CTkLabel(row, text=label, font=(FONT_FAMILY, 12.5), text_color=COLOR_MUTED).pack(side="left")
            ctk.CTkLabel(
                row, text=str(value), font=(FONT_FAMILY, 13.5, "bold"), text_color=color
            ).pack(side="right")

        stat_row("PO Checked", checked_count)
        stat_row("Tally (Match)", match_count, COLOR_SUCCESS)
        stat_row("Less Posting", less_count, COLOR_WARN if less_count else COLOR_MUTED)
        stat_row("Over Posting", over_count, COLOR_ERROR if over_count else COLOR_MUTED)

        if failed_count:
            stat_row("Failed", failed_count, COLOR_ERROR)

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(0, 18))

        ctk.CTkButton(
            btn_row, text="Close", command=self.destroy, width=110,
            fg_color="#334155", hover_color="#475569",
            font=(FONT_FAMILY, 12), cursor="hand2",
        ).pack(side="right")

        ctk.CTkButton(
            btn_row, text="  Open Excel", image=icons.make_icon("box", color="#FFFFFF", size=14),
            compound="left", command=self._open_excel, width=150,
            fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER,
            font=(FONT_FAMILY, 12, "bold"), cursor="hand2",
        ).pack(side="right", padx=(0, 10))

    def _open_excel(self):

        try:
            os.startfile(os.path.abspath(config.EXCEL_FILE))
        except Exception as e:
            messagebox.showwarning("Could Not Open Excel", str(e))
