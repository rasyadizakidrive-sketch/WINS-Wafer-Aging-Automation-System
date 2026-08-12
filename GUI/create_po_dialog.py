import customtkinter as ctk
from tkinter import messagebox
from tkcalendar import DateEntry

import config
import icons

# ==============================================================
# DESIGN TOKENS (matching the main window's visual language --
# this dialog should look like part of the same app, not a
# separate, unstyled window bolted onto it)
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


class CreatePODialog(ctk.CTkToplevel):

    def __init__(self, parent, batches, product=None):

        super().__init__(parent)

        self.configure(fg_color=COLOR_BG)
        self.title("Create Production Order")
        self.grab_set()

        self.result = None
        self.product_name = product or "-"

        self.batch_data = {
            item["batch"]: item
            for item in batches
            if item.get("po", "") == ""
        }

        self.batch_list = [
            f'{item["batch"]} ({item["qty"]:,} pcs)'
            for item in batches
            if item.get("po", "") == ""
        ]

        # Material choices are scoped to the product currently selected on
        # the main window (M6 / G12) -- e.g. "M6 Longi" -> "CF02-0119".
        # config.MATERIAL_OPTIONS is a dict of {product: {display_name: material_code}}.
        self.material_map = dict(getattr(config, "MATERIAL_OPTIONS", {}).get(product, {}))
        self.material_display_list = list(self.material_map.keys())

        # Plant never actually varies run to run -- it's read from config
        # rather than asked for every time, same as everywhere else the
        # app talks to SAP.
        self.plant = getattr(config, "SAP_PLANT", "P200")

        self.auto_qty = ctk.BooleanVar(value=True)
        self.create_all = ctk.BooleanVar(value=False)

        self._build_ui()
        self._update_summary()

    # ==========================================================
    # UI BUILDERS
    # ==========================================================

    def _build_ui(self):

        # A fixed 560x700 assumes every screen has at least 700px of usable
        # vertical space above the taskbar. It doesn't always -- this is
        # the same lesson the main window learned the hard way. Rather
        # than guess a constant, size against the actual screen and
        # center it, with resizing left on as a fallback in case the
        # guess is still wrong for some particular machine.
        self.update_idletasks()

        natural_w, natural_h = 560, 620
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()

        max_h = max(420, screen_h - 140)  # leave room for taskbar + title bar
        dialog_w = min(natural_w, screen_w - 40)
        dialog_h = min(natural_h, max_h)

        x = max(0, (screen_w - dialog_w) // 2)
        y = max(0, (screen_h - dialog_h) // 2)

        self.geometry(f"{dialog_w}x{dialog_h}+{x}+{y}")
        self.minsize(460, 380)
        self.resizable(True, True)

        # ---- Header ----
        header = ctk.CTkFrame(self, fg_color=COLOR_BG_HEADER, corner_radius=0)
        header.pack(fill="x")

        ctk.CTkLabel(
            header, text="Create Production Order",
            font=(FONT_FAMILY, 16, "bold"), text_color=COLOR_TEXT,
        ).pack(anchor="w", padx=22, pady=(11, 1))

        ctk.CTkLabel(
            header, text=f"SAP TCODE : CO01  \u00b7  Product {self.product_name}",
            font=(FONT_FAMILY, 11), text_color=COLOR_MUTED,
        ).pack(anchor="w", padx=22, pady=(0, 10))

        # ---- Footer (built and packed before the scrollable content, so
        # it reserves its space first -- pack() gives any shortfall to
        # whichever section is packed *last*, and the buttons in here are
        # the one thing that must never lose that fight. The scrollable
        # area below is specifically the part designed to give way and
        # scroll if the dialog's fixed height isn't enough for everything.) ----
        footer = ctk.CTkFrame(self, fg_color=COLOR_BG_HEADER, corner_radius=0)
        footer.pack(fill="x", side="bottom")

        btn_row = ctk.CTkFrame(footer, fg_color="transparent")
        btn_row.pack(fill="x", padx=16, pady=14)

        ctk.CTkButton(
            btn_row, text="Cancel", command=self.destroy, width=120,
            fg_color="#334155", hover_color="#475569",
            font=(FONT_FAMILY, 12), cursor="hand2",
        ).pack(side="right")

        self.submit_btn = ctk.CTkButton(
            btn_row, text="  Create PO", image=icons.make_icon("doc_plus", color="#FFFFFF", size=14),
            compound="left", command=self.create_po, width=160,
            fg_color=COLOR_PRIMARY, hover_color=COLOR_PRIMARY_HOVER,
            font=(FONT_FAMILY, 12, "bold"), cursor="hand2",
        )
        self.submit_btn.pack(side="right", padx=(0, 10))

        # ---- Scrollable content ----
        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=16, pady=(10, 0))

        # -- Section: batch selection --
        batch_card = ctk.CTkFrame(
            scroll, fg_color=COLOR_CARD, border_color=COLOR_CARD_BORDER, border_width=1, corner_radius=12
        )
        batch_card.pack(fill="x", pady=(0, 12))

        self.all_chk = ctk.CTkCheckBox(
            batch_card,
            text=f"Create ALL Pending PO  \u00b7  {len(self.batch_list)} batch waiting",
            variable=self.create_all,
            command=self._toggle_all_mode,
            font=(FONT_FAMILY, 12.5, "bold"),
            text_color=COLOR_TEXT,
        )
        self.all_chk.pack(anchor="w", padx=20, pady=(14, 12))
        if not self.batch_list:
            self.all_chk.configure(state="disabled")

        ctk.CTkFrame(batch_card, height=1, fg_color=COLOR_CARD_BORDER).pack(fill="x", padx=20)

        self._field_label(batch_card, "Batch")
        self.batch_combo = ctk.CTkComboBox(
            batch_card, values=self.batch_list, command=self.batch_changed, width=460
        )
        self.batch_combo.pack(padx=20)

        self.auto_chk = ctk.CTkCheckBox(
            batch_card, text="Auto-fill quantity from batch", variable=self.auto_qty,
            font=(FONT_FAMILY, 11.5), text_color=COLOR_MUTED,
        )
        self.auto_chk.pack(anchor="w", padx=20, pady=(10, 4))

        self._field_label(batch_card, "Quantity")
        self.qty_entry = ctk.CTkEntry(batch_card, width=460)
        self.qty_entry.pack(padx=20, pady=(0, 16))
        self.qty_entry.bind("<KeyRelease>", lambda e: self._update_summary())

        # -- Section: production details --
        details_card = ctk.CTkFrame(
            scroll, fg_color=COLOR_CARD, border_color=COLOR_CARD_BORDER, border_width=1, corner_radius=12
        )
        details_card.pack(fill="x", pady=(0, 12))

        self._section_label(details_card, "PRODUCTION DETAILS")

        dates_row = ctk.CTkFrame(details_card, fg_color="transparent")
        dates_row.pack(fill="x", padx=20, pady=(0, 4))
        dates_row.grid_columnconfigure(0, weight=1, uniform="d")
        dates_row.grid_columnconfigure(1, weight=1, uniform="d")

        start_col = ctk.CTkFrame(dates_row, fg_color="transparent")
        start_col.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkLabel(
            start_col, text="Start Date", font=(FONT_FAMILY, 12), text_color=COLOR_TEXT
        ).pack(anchor="w", pady=(0, 5))
        self.start_entry = DateEntry(start_col, width=16, date_pattern="dd.mm.yyyy")
        self.start_entry.pack(fill="x")

        finish_col = ctk.CTkFrame(dates_row, fg_color="transparent")
        finish_col.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        ctk.CTkLabel(
            finish_col, text="Finish Date", font=(FONT_FAMILY, 12), text_color=COLOR_TEXT
        ).pack(anchor="w", pady=(0, 5))
        self.finish_entry = DateEntry(finish_col, width=16, date_pattern="dd.mm.yyyy")
        self.finish_entry.pack(fill="x")

        self._field_label(details_card, "Material")
        self.material_combo = ctk.CTkComboBox(
            details_card, values=self.material_display_list, width=460,
            command=lambda v: self._update_summary(),
        )
        self.material_combo.pack(padx=20, pady=(0, 18))

        if self.material_display_list:
            self.material_combo.set(self.material_display_list[0])
        else:
            self.material_combo.set("")
            self.material_combo.configure(state="disabled")

        # -- Live summary --
        self.summary_label = ctk.CTkLabel(
            scroll, text="", font=(FONT_FAMILY, 11.5), text_color=COLOR_MUTED,
            wraplength=500, justify="left",
        )
        self.summary_label.pack(anchor="w", padx=6, pady=(0, 10))

        self.bind("<Return>", lambda e: self.create_po())
        self.bind("<Escape>", lambda e: self.destroy())

        if self.batch_list:
            self.batch_combo.set(self.batch_list[0])
            self.batch_changed(self.batch_list[0])
        else:
            self.batch_combo.configure(state="disabled")
            self.qty_entry.configure(state="disabled")
            self.auto_chk.configure(state="disabled")

    def _section_label(self, parent, text):
        ctk.CTkLabel(
            parent, text=text, font=(FONT_FAMILY, 11, "bold"), text_color=COLOR_MUTED,
        ).pack(anchor="w", padx=20, pady=(13, 6))

    def _field_label(self, parent, text):
        ctk.CTkLabel(
            parent, text=text, font=(FONT_FAMILY, 12), text_color=COLOR_TEXT,
        ).pack(anchor="w", padx=20, pady=(8, 4))

    def _toggle_all_mode(self):

        if self.create_all.get():
            state = "disabled"
            self.submit_btn.configure(text="  Create ALL PO")
        else:
            state = "normal"
            self.submit_btn.configure(text="  Create PO")

        self.batch_combo.configure(state=state)
        self.auto_chk.configure(state=state)
        self.qty_entry.configure(state=state)

        self._update_summary()

    def batch_changed(self, value):

        batch = value.split(" ")[0]

        data = self.batch_data.get(batch)

        if data is None:
            return

        if self.auto_qty.get():

            self.qty_entry.delete(0, "end")
            self.qty_entry.insert(0, str(data["qty"]))

        self._update_summary()

    # ==========================================================
    # LIVE SUMMARY
    # ==========================================================

    def _update_summary(self):

        if not self.batch_list:
            self.summary_label.configure(text="No pending batch to create.", text_color=COLOR_MUTED)
            return

        if self.create_all.get():
            self.summary_label.configure(
                text=f"\u25b8 Will create production orders for all {len(self.batch_list)} pending batches, "
                     f"using the dates and material below for each.",
                text_color=COLOR_WARN,
            )
            return

        batch = self.batch_combo.get().split(" ")[0]
        data = self.batch_data.get(batch)

        if data is None:
            self.summary_label.configure(text="Select a batch to continue.", text_color=COLOR_MUTED)
            return

        qty_text = self.qty_entry.get().replace(",", "").strip()

        try:
            qty = int(qty_text)
            qty_display = f"{qty:,} pcs"
        except ValueError:
            qty_display = "an invalid quantity"

        self.summary_label.configure(
            text=f"\u25b8 Will create one production order for batch {batch} ({qty_display}).",
            text_color=COLOR_MUTED,
        )

    # ==========================================================
    # VALIDATION
    # ==========================================================

    def _selected_material_code(self):
        """Resolves the chosen display name (e.g. 'M6 Longi') to its material code (e.g. 'CF02-0119')."""
        return self.material_map.get(self.material_combo.get(), "")

    def _shared_fields_valid(self):

        if not self.material_display_list:
            messagebox.showwarning(
                "Missing Information",
                "No material options configured for this product.\n"
                "Add it to config.MATERIAL_OPTIONS."
            )
            return False

        if not self._selected_material_code():
            messagebox.showwarning("Missing Information", "Please select a Material.")
            return False

        if not self.start_entry.get().strip() or not self.finish_entry.get().strip():
            messagebox.showwarning("Missing Information", "Start Date and Finish Date are required.")
            return False

        return True

    def create_po(self):

        if not self.batch_list:
            messagebox.showinfo("Information", "No pending batch to create.")
            return

        if not self._shared_fields_valid():
            return

        if self.create_all.get():

            self.result = {
                "mode": "all",
                "plant": self.plant,
                "material": self._selected_material_code(),
                "start_date": self.start_entry.get(),
                "finish_date": self.finish_entry.get(),
            }

            self.destroy()
            return

        batch = self.batch_combo.get().split(" ")[0]
        data = self.batch_data.get(batch)

        if data is None:
            messagebox.showwarning("Missing Information", "Please select a valid batch.")
            return

        qty_text = self.qty_entry.get().replace(",", "").strip()

        try:
            qty = int(qty_text)
            if qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Missing Information", "Quantity must be a positive number.")
            return

        self.result = {
            "mode": "single",
            "batch": data["batch"],
            "qty": qty,
            "row": data["row"],
            "plant": self.plant,
            "material": self._selected_material_code(),
            "start_date": self.start_entry.get(),
            "finish_date": self.finish_entry.get(),
        }

        self.destroy()


class POSummaryDialog(ctk.CTkToplevel):
    """
    Shown after a bulk (Create ALL) run completes. Reports success/failed
    counts and elapsed time, and offers to retry just the failed batches
    without re-running the ones that already succeeded.
    """

    def __init__(self, parent, success_count, failed_count, elapsed_seconds, log_path, on_retry=None):

        super().__init__(parent)

        self.configure(fg_color=COLOR_BG)
        self.title("Production Order Summary")
        self.grab_set()

        self.update_idletasks()
        natural_w, natural_h = 400, 340
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        max_h = max(300, screen_h - 140)
        dialog_w = min(natural_w, screen_w - 40)
        dialog_h = min(natural_h, max_h)
        x = max(0, (screen_w - dialog_w) // 2)
        y = max(0, (screen_h - dialog_h) // 2)
        self.geometry(f"{dialog_w}x{dialog_h}+{x}+{y}")
        self.resizable(True, True)

        minutes, seconds = divmod(int(elapsed_seconds), 60)
        time_text = f"{minutes} min {seconds} sec" if minutes else f"{seconds} sec"

        header = ctk.CTkFrame(self, fg_color=COLOR_BG_HEADER, corner_radius=0)
        header.pack(fill="x")

        icon_kind = "check" if failed_count == 0 else "warning"
        icon_color = COLOR_SUCCESS if failed_count == 0 else COLOR_WARN
        ctk.CTkLabel(
            header, text="  Production Order Completed",
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

        stat_row("Success", success_count, COLOR_SUCCESS)
        stat_row("Failed", failed_count, COLOR_ERROR if failed_count else COLOR_MUTED)
        stat_row("Time elapsed", time_text)

        ctk.CTkFrame(card, height=1, fg_color=COLOR_CARD_BORDER).pack(fill="x", padx=20, pady=(4, 12))

        ctk.CTkLabel(
            card,
            text=f"Log :  {log_path}",
            text_color=COLOR_MUTED,
            font=(FONT_FAMILY, 10.5),
            wraplength=320,
            justify="left",
        ).pack(anchor="w", padx=20, pady=(0, 16))

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(0, 18))

        ctk.CTkButton(
            btn_row,
            text="Close",
            command=self.destroy,
            fg_color="#334155",
            hover_color="#475569",
            font=(FONT_FAMILY, 12),
            cursor="hand2",
        ).pack(side="right")

        if failed_count and on_retry:
            ctk.CTkButton(
                btn_row,
                text=f"  Retry {failed_count} Failed",
                image=icons.make_icon("refresh", color="#FFFFFF", size=14), compound="left",
                command=lambda: [self.destroy(), on_retry()],
                fg_color=COLOR_WARN,
                hover_color="#B45309",
                font=(FONT_FAMILY, 12, "bold"),
                cursor="hand2",
            ).pack(side="left")
