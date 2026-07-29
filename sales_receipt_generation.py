# ── Portable path bootstrap (must be FIRST) ─────────────────────────────────
import os, sys

def app_dir() -> str:
    """Folder that contains auto.exe (or this .py file while developing)."""
    if getattr(sys, "frozen", False):  # running as PyInstaller exe
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def p(*parts: str) -> str:
    """Join path parts relative to app_dir()."""
    return os.path.join(app_dir(), *parts)

# Tell Playwright where to find browsers BEFORE importing Playwright (directly or indirectly)
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", p("pw-browsers"))
# ────────────────────────────────────────────────────────────────────────────

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import json
import os
import math
from pathlib import Path
import shutil
import subprocess
import sys
import threading
import urllib.error
import urllib.request
from typing import Optional

try:
    from openpyxl import Workbook
except ImportError:
    Workbook = None

from clipboard_export import (
    DEFAULT_CLIPBOARD_COL_DEFS,
    PRESET_FIELD_KEYS,
    PRESET_FIELD_LABELS,
    TRANSACTION_FIELD_LABELS,
    build_excel_clipboard_text,
    load_clipboard_col_defs,
    normalize_col_def,
    normalize_col_defs,
    normalize_date,
    save_clipboard_col_defs,
    slug_custom_field_key,
)

try:
    import automation as automation_module
    AUTOMATION_IMPORT_ERROR = None
except Exception as err:
    automation_module = None
    AUTOMATION_IMPORT_ERROR = err

def _calc_scroll_delta(event) -> int:
    """Return scroll units from a MouseWheel / Button-4/5 event.

    Tk 9.0 on macOS: delta = int(dy * 65536), a fixed-point representation.
    Values like 65475 ≈ +1 pixel, 131069 ≈ +2 pixels, -131115 ≈ -2 pixels.
    """
    if getattr(event, "num", None) == 4:
        return -1
    elif getattr(event, "num", None) == 5:
        return 1

    d = getattr(event, "delta", 0)
    if d == 0:
        return 0

    if sys.platform == "darwin" and _IS_AQUA_TK9:
        # Tk 9.0: delta = int(dy * 65536), fixed-point 16.16 format
        pixels = d / 65536.0
        units = int(round(pixels))
        # Ensure at least 1 unit of movement for any non-zero scroll
        if units == 0:
            units = 1 if pixels > 0 else -1
        # Clamp for smooth scrolling
        units = max(-5, min(5, units))
        return -units   # negate: positive dy (up in Cocoa) → scroll up (negative yview)
    elif sys.platform == "darwin":
        # Tk 8.6 macOS: delta is ±1..±N directly
        return -1 if d > 0 else 1
    else:
        # Windows / Linux: delta is ±120 per notch
        return int(d / -120) or (-1 if d > 0 else 1)

# Detect Tk 9.0 + macOS Aqua — tk.Button(relief="flat", bg=...) is non-functional
_IS_AQUA_TK9 = False
try:
    _probe = tk.Tk()
    _probe.withdraw()
    _IS_AQUA_TK9 = (
        _probe.tk.call("tk", "windowingsystem") == "aqua"
        and float(tk.TkVersion) >= 9.0
    )
    _probe.destroy()
    del _probe
except Exception:
    pass


def _make_button(
    parent, *, text="", command=None, bg="#e5e7eb", fg="#333333",
    activebackground=None, activeforeground=None,
    font=("Segoe UI", 10), padx=12, pady=6, cursor="hand2",
    relief="flat", borderwidth=0, width=None, state="normal",
    highlightthickness=0, **_extra
):
    """Create a styled button that works on ALL platforms including macOS Tk 9.0.

    On macOS Tk 9.0 Aqua, tk.Button with relief='flat' and custom bg/fg
    silently ignores clicks.  We work around this by building buttons from
    tk.Frame + tk.Label + <Button-1> bindings.
    """
    if activebackground is None:
        activebackground = bg
    if activeforeground is None:
        activeforeground = fg

    if not _IS_AQUA_TK9:
        # Standard tk.Button on Linux / Windows / older Tk
        btn_kwargs = dict(
            text=text, command=command, bg=bg, fg=fg, relief=relief,
            activebackground=activebackground, activeforeground=activeforeground,
            font=font, padx=padx, pady=pady, cursor=cursor,
            borderwidth=borderwidth, highlightthickness=highlightthickness,
            state=state,
        )
        if width is not None:
            btn_kwargs["width"] = width
        return tk.Button(parent, **btn_kwargs)

    # ── macOS Tk 9.0 workaround: Frame + Label acting as a button ──
    frame = tk.Frame(parent, bg=bg, cursor=cursor, padx=padx, pady=pady,
                     highlightthickness=1, highlightbackground=bg)
    lbl = tk.Label(frame, text=text, bg=bg, fg=fg, font=font, cursor=cursor)
    lbl.pack()

    # Store state so we can disable / enable / reconfig later
    frame._btn_command = command
    frame._btn_label = lbl
    frame._btn_bg = bg
    frame._btn_fg = fg
    frame._btn_abg = activebackground
    frame._btn_afg = activeforeground
    frame._btn_state = state

    def _on_click(event=None):
        if frame._btn_state == "disabled":
            return
        cb = frame._btn_command
        if cb:
            cb()

    # Capture the REAL frame.config before we monkey-patch it
    _original_frame_config = frame.configure

    def _on_enter(event=None):
        if frame._btn_state == "disabled":
            return
        _original_frame_config(bg=frame._btn_abg, highlightbackground=frame._btn_abg)
        lbl.config(bg=frame._btn_abg, fg=frame._btn_afg)

    def _on_leave(event=None):
        cur_bg = frame._btn_bg
        _original_frame_config(bg=cur_bg, highlightbackground=cur_bg)
        lbl.config(bg=cur_bg, fg=frame._btn_fg)

    for w in (frame, lbl):
        w.bind("<Button-1>", _on_click)
        w.bind("<Enter>", _on_enter)
        w.bind("<Leave>", _on_leave)

    # Monkey-patch .config() and .cget() so calling code can use btn.config(text=...)

    def _patched_config(cnf=None, **kw):
        # Intercept text, bg, fg, state, command changes
        if cnf:
            kw.update(cnf)
        text_val = kw.pop("text", None)
        bg_val = kw.pop("bg", None)
        fg_val = kw.pop("fg", None)
        state_val = kw.pop("state", None)
        cmd_val = kw.pop("command", None)
        abg_val = kw.pop("activebackground", None)
        afg_val = kw.pop("activeforeground", None)
        cursor_val = kw.pop("cursor", None)
        # Ignore font kwarg changes for simplicity, accept gracefully
        kw.pop("font", None)
        kw.pop("relief", None)
        kw.pop("padx", None)
        kw.pop("pady", None)
        kw.pop("borderwidth", None)
        kw.pop("highlightthickness", None)
        kw.pop("width", None)

        if text_val is not None:
            lbl.config(text=text_val)
        if bg_val is not None:
            frame._btn_bg = bg_val
            _original_frame_config(bg=bg_val, highlightbackground=bg_val)
            lbl.config(bg=bg_val)
        if fg_val is not None:
            frame._btn_fg = fg_val
            lbl.config(fg=fg_val)
        if abg_val is not None:
            frame._btn_abg = abg_val
        if afg_val is not None:
            frame._btn_afg = afg_val
        if state_val is not None:
            frame._btn_state = state_val
            if state_val == "disabled":
                _original_frame_config(cursor="")
                lbl.config(cursor="")
            else:
                _original_frame_config(cursor="hand2")
                lbl.config(cursor="hand2")
        if cmd_val is not None:
            frame._btn_command = cmd_val
        if cursor_val is not None and frame._btn_state != "disabled":
            _original_frame_config(cursor=cursor_val)
            lbl.config(cursor=cursor_val)
        if kw:
            _original_frame_config(**kw)

    frame.config = _patched_config
    frame.configure = _patched_config

    return frame

# --- Constants & Configuration ---
bg_color = "#F8F9FA"
sidebar_color = "#343A40"
header_color = "#FFFFFF"
accent_color = "#5C5252"
text_color = "#333333"
table_header_bg = "#E9ECEF"

API_TRANSACTIONS_URL = "https://uat-sobha.docuxray.ai/api/prePost/getAllPrePosted"
API_TOKEN = os.environ.get(
    "SOBHA_API_TOKEN",
    "hDmflUMicbk9oB8jRFBxOnlaYUkzP4jSYQwJm1weZWK",
)
# Keep automation.py and UI fetch in sync for bearer token usage.
os.environ.setdefault("SOBHA_API_TOKEN", API_TOKEN)

# Legacy dialog inputs (kept to avoid breaking optional dialog flow)
projects = []
units = {}

# Key map: maps column index to dict key
KEY_MAP = [
    "check",
    "batch_id",
    "value_date",
    "account",
    "credit",
    "offset_account",
    "method_of_payment",
    "reference_date",
    "payment_reference",
]

HEADERS = [
    "✔",
    "Batch",
    "Value Date",
    "Account",
    "Credit",
    "Offset Account",
    "Method of Payment",
    "Reference Date",
    "Payment Reference",
]

COL_WIDTHS = [50, 220, 140, 160, 140, 280, 180, 140, 360]

DATE_FIELD_KEYS = frozenset({"date", "value_date", "reference_date", "account_date"})


def format_d365_date(value) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    normalized = normalize_date(raw)
    return normalized or raw

# Columns shown in the batch transaction table.
DISPLAY_COL_DEFS = [
    ("value_date", "Value Date"),
    ("account", "Account"),
    ("credit", "Credit"),
    ("offset_account", "Offset Account"),
    ("method_of_payment", "Method of Payment"),
    ("reference_date", "Reference Date"),
    ("payment_reference", "Payment Reference"),
]

# ---------------------------------------------------------------------------
# Clipboard Column Settings Dialog
# ---------------------------------------------------------------------------
class ClipboardColumnSettingsDialog(tk.Toplevel):
    def __init__(self, parent, col_defs: list, on_save=None):
        super().__init__(parent)
        self.title("Copy Column Settings")
        self.geometry("900x680")
        self.minsize(780, 560)
        self.configure(bg="white")
        self.on_save = on_save
        self.row_items = []
        self._colors = getattr(parent, "colors", {})

        self.transient(parent)
        self.grab_set()

        main = tk.Frame(self, bg="white", padx=20, pady=16)
        main.pack(fill="both", expand=True)

        tk.Label(
            main,
            text="Copy Column Settings",
            bg="white",
            fg=self._colors.get("title", text_color),
            font=("Segoe UI", 13, "bold"),
        ).pack(anchor="w")
        tk.Label(
            main,
            text="Reorder columns, rename headers, add or remove fields. Copy uses this layout.",
            bg="white",
            fg=self._colors.get("muted", "#6b7280"),
            font=("Segoe UI", 9),
            wraplength=700,
            justify="left",
        ).pack(anchor="w", pady=(4, 12))

        table_shell = tk.Frame(main, bg="#eef2f8", bd=0, highlightthickness=1, highlightbackground="#dfe4ee")
        table_shell.pack(fill="both", expand=True)

        header = tk.Frame(table_shell, bg="#f9fafd", pady=6)
        header.pack(fill="x", padx=1, pady=(1, 0))
        for col_idx, (text, width, anchor) in enumerate(
            (
                ("#", 36, "center"),
                ("Field", 170, "w"),
                ("Excel header", 220, "w"),
                ("Default value", 160, "w"),
                ("Order", 120, "center"),
            )
        ):
            header.grid_columnconfigure(col_idx, weight=1 if col_idx in {2, 3} else 0)
            tk.Label(
                header,
                text=text,
                bg="#f9fafd",
                fg="#6b7280",
                font=("Segoe UI", 8, "bold"),
                width=width // 8 if width else None,
                anchor=anchor,
            ).grid(row=0, column=col_idx, sticky="ew", padx=(10 if col_idx == 0 else 6, 6))

        list_host = tk.Frame(table_shell, bg="white")
        list_host.pack(fill="both", expand=True, padx=1, pady=(0, 1))

        self.list_canvas = tk.Canvas(list_host, bg="white", highlightthickness=0, bd=0)
        list_scroll = ttk.Scrollbar(list_host, orient="vertical", command=self.list_canvas.yview)
        self.list_canvas.configure(yscrollcommand=list_scroll.set)
        list_scroll.pack(side="right", fill="y")
        self.list_canvas.pack(side="left", fill="both", expand=True)

        self.rows_frame = tk.Frame(self.list_canvas, bg="white")
        self.list_canvas_window = self.list_canvas.create_window((0, 0), window=self.rows_frame, anchor="nw")
        self.rows_frame.bind("<Configure>", self._on_rows_configure)
        self.list_canvas.bind("<Configure>", self._on_canvas_configure)
        self.list_canvas.bind("<Enter>", self._bind_mousewheel)
        self.list_canvas.bind("<Leave>", self._unbind_mousewheel)

        add_row = tk.Frame(main, bg="white")
        add_row.pack(fill="x", pady=(10, 0))
        tk.Label(add_row, text="Add preset field:", bg="white", fg=text_color, font=("Segoe UI", 9, "bold")).pack(side="left")
        self.add_field_var = tk.StringVar()
        self.add_field_cb = ttk.Combobox(add_row, textvariable=self.add_field_var, state="readonly", width=30)
        self.add_field_cb.pack(side="left", padx=(8, 8))
        _make_button(
            add_row,
            text="Add",
            command=self._add_selected_field,
            bg="#eef2ff",
            fg="#1d4ed8",
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padx=12,
            pady=4,
            cursor="hand2",
        ).pack(side="left")

        custom_row = tk.Frame(main, bg="white")
        custom_row.pack(fill="x", pady=(8, 0))
        tk.Label(custom_row, text="Add custom field:", bg="white", fg=text_color, font=("Segoe UI", 9, "bold")).pack(side="left")
        tk.Label(custom_row, text="Header", bg="white", fg="#6b7280", font=("Segoe UI", 8)).pack(side="left", padx=(10, 4))
        self.custom_header_var = tk.StringVar()
        tk.Entry(custom_row, textvariable=self.custom_header_var, width=24, font=("Segoe UI", 9)).pack(side="left", padx=(0, 8))
        tk.Label(custom_row, text="Default value", bg="white", fg="#6b7280", font=("Segoe UI", 8)).pack(side="left", padx=(0, 4))
        self.custom_default_var = tk.StringVar()
        tk.Entry(custom_row, textvariable=self.custom_default_var, width=20, font=("Segoe UI", 9)).pack(side="left", padx=(0, 8))
        _make_button(
            custom_row,
            text="Add custom",
            command=self._add_custom_field,
            bg="#ecfdf5",
            fg="#047857",
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padx=12,
            pady=4,
            cursor="hand2",
        ).pack(side="left")

        footer = tk.Frame(main, bg="white")
        footer.pack(fill="x", pady=(14, 0))
        _make_button(
            footer,
            text="Reset defaults",
            command=self._reset_defaults,
            bg="#f3f4f6",
            fg="#374151",
            relief="flat",
            font=("Segoe UI", 9),
            padx=12,
            pady=6,
            cursor="hand2",
        ).pack(side="left")
        _make_button(
            footer,
            text="Cancel",
            command=self.destroy,
            bg="white",
            fg="#555",
            relief="solid",
            borderwidth=1,
            font=("Segoe UI", 9),
            padx=14,
            pady=5,
            cursor="hand2",
        ).pack(side="right", padx=(8, 0))
        _make_button(
            footer,
            text="Save",
            command=self._save,
            bg=accent_color,
            fg="white",
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padx=16,
            pady=6,
            cursor="hand2",
        ).pack(side="right")

        self._set_rows(col_defs)
        self._center_window()

    def _center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _on_rows_configure(self, _event):
        self.list_canvas.configure(scrollregion=self.list_canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.list_canvas.itemconfigure(self.list_canvas_window, width=event.width)

    def _bind_mousewheel(self, _event):
        self.list_canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.list_canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.list_canvas.bind_all("<Button-5>", self._on_mousewheel)

    def _unbind_mousewheel(self, _event):
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.list_canvas.unbind_all(seq)

    def _on_mousewheel(self, event):
        delta = _calc_scroll_delta(event)
        if delta != 0:
            self.list_canvas.yview_scroll(delta, "units")

    def _set_rows(self, col_defs: list):
        for widget in self.rows_frame.winfo_children():
            widget.destroy()
        self.row_items = []
        normalized = normalize_col_defs(col_defs)
        for index, col in enumerate(normalized):
            self._append_row(col, index + 1)
        self._refresh_add_field_options()
        self._on_rows_configure(None)

    def _append_row(self, col: dict, index: int):
        key = col["key"]
        label = col["label"]
        default_value = col.get("default_value", "")
        is_custom = key.startswith("custom_") or key not in PRESET_FIELD_KEYS
        row_bg = "#ffffff" if index % 2 else "#fcfdff"
        row = tk.Frame(self.rows_frame, bg=row_bg, pady=4)
        row.pack(fill="x")

        tk.Label(
            row,
            text=str(index),
            bg=row_bg,
            fg="#6b7280",
            font=("Segoe UI", 9),
            width=4,
            anchor="center",
        ).grid(row=0, column=0, sticky="w", padx=(10, 6))

        field_text = f"{key} (custom)" if is_custom else key
        tk.Label(
            row,
            text=field_text,
            bg=row_bg,
            fg="#7c3aed" if is_custom else "#2e3b57",
            font=("Segoe UI", 9),
            width=22,
            anchor="w",
        ).grid(row=0, column=1, sticky="w", padx=(0, 8))

        label_var = tk.StringVar(value=label)
        tk.Entry(
            row,
            textvariable=label_var,
            relief="solid",
            borderwidth=1,
            font=("Segoe UI", 9),
        ).grid(row=0, column=2, sticky="ew", padx=(0, 8))

        default_var = tk.StringVar(value=default_value)
        tk.Entry(
            row,
            textvariable=default_var,
            relief="solid",
            borderwidth=1,
            font=("Segoe UI", 9),
        ).grid(row=0, column=3, sticky="ew", padx=(0, 8))

        actions = tk.Frame(row, bg=row_bg)
        actions.grid(row=0, column=4, sticky="e", padx=(0, 8))
        _make_button(
            actions,
            text="↑",
            command=lambda: self._move_row(key, -1),
            width=2,
            relief="flat",
            bg="#eef2ff",
            fg="#1d4ed8",
            cursor="hand2",
        ).pack(side="left", padx=(0, 4))
        _make_button(
            actions,
            text="↓",
            command=lambda: self._move_row(key, 1),
            width=2,
            relief="flat",
            bg="#eef2ff",
            fg="#1d4ed8",
            cursor="hand2",
        ).pack(side="left", padx=(0, 4))
        _make_button(
            actions,
            text="✕",
            command=lambda: self._remove_row(key),
            width=2,
            relief="flat",
            bg="#fff0f0",
            fg="#c0392b",
            cursor="hand2",
        ).pack(side="left")

        row.grid_columnconfigure(2, weight=1)
        row.grid_columnconfigure(3, weight=1)
        self.row_items.append({"key": key, "label_var": label_var, "default_var": default_var})

    def _current_col_defs(self) -> list:
        return [
            {
                "key": item["key"],
                "label": item["label_var"].get().strip() or item["key"],
                "default_value": item["default_var"].get().strip(),
            }
            for item in self.row_items
        ]

    def _refresh_add_field_options(self):
        used = {item["key"] for item in self.row_items}
        available = [key for key in PRESET_FIELD_KEYS if key not in used]
        labels = [f"{key} ({PRESET_FIELD_LABELS.get(key, key)})" for key in available]
        self.add_field_cb["values"] = labels
        if labels:
            self.add_field_var.set(labels[0])
        else:
            self.add_field_var.set("")

    def _add_selected_field(self):
        selected = self.add_field_var.get().strip()
        if not selected:
            return
        key = selected.split(" (", 1)[0].strip()
        if not key or any(item["key"] == key for item in self.row_items):
            return
        default_label = PRESET_FIELD_LABELS.get(key, key)
        col_defs = self._current_col_defs()
        col_defs.append({"key": key, "label": default_label, "default_value": ""})
        self._set_rows(col_defs)

    def _add_custom_field(self):
        header = self.custom_header_var.get().strip()
        if not header:
            messagebox.showwarning("Copy Column Settings", "Enter a header name for the custom field.")
            return
        default_value = self.custom_default_var.get().strip()
        used = {item["key"] for item in self.row_items}
        key = slug_custom_field_key(header, used)
        col_defs = self._current_col_defs()
        col_defs.append({"key": key, "label": header, "default_value": default_value})
        self._set_rows(col_defs)
        self.custom_header_var.set("")
        self.custom_default_var.set("")

    def _move_row(self, key: str, direction: int):
        keys = [item["key"] for item in self.row_items]
        if key not in keys:
            return
        idx = keys.index(key)
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(keys):
            return
        col_defs = self._current_col_defs()
        col_defs[idx], col_defs[new_idx] = col_defs[new_idx], col_defs[idx]
        self._set_rows(col_defs)

    def _remove_row(self, key: str):
        if len(self.row_items) <= 1:
            messagebox.showwarning("Copy Column Settings", "At least one column is required.")
            return
        col_defs = [col for col in self._current_col_defs() if col["key"] != key]
        self._set_rows(col_defs)

    def _reset_defaults(self):
        if messagebox.askyesno("Reset defaults", "Restore default copy columns?"):
            self._set_rows([normalize_col_def(item) for item in DEFAULT_CLIPBOARD_COL_DEFS])

    def _save(self):
        col_defs = self._current_col_defs()
        if not col_defs:
            messagebox.showwarning("Copy Column Settings", "Add at least one column before saving.")
            return
        if any(not col["label"] for col in col_defs):
            messagebox.showwarning("Copy Column Settings", "Every column needs a header name.")
            return
        save_clipboard_col_defs(col_defs)
        if callable(self.on_save):
            self.on_save(col_defs)
        self.destroy()


# ---------------------------------------------------------------------------
# Sales Acc Receipt Gen Dialog
# ---------------------------------------------------------------------------
class SalesAccReceiptGenDialog(tk.Toplevel):
    def __init__(self, parent, row_data, callback=None):
        super().__init__(parent)
        self.title("Sales Acc Receipt Gen")
        self.geometry("480x540")
        self.resizable(False, False)
        self.configure(bg="white")
        self.row_data = row_data
        self.callback = callback
        
        # Initialize variables
        self.project_var = tk.StringVar(value=self.row_data.get("project", ""))
        self.unit_var = tk.StringVar(value=self.row_data.get("unit", ""))
        self.amount_type_var = tk.StringVar(value="gross")
        self.amount_var = tk.StringVar()
        
        self.amount_var = tk.StringVar()
        
        self.transient(parent)
        self.grab_set()

        # --- UI Construction ---
        main = tk.Frame(self, bg="white", padx=24, pady=20)
        main.pack(fill="both", expand=True)

        # Title bar
        title_bar = tk.Frame(main, bg="#F0F0F0", padx=12, pady=8)
        title_bar.pack(fill="x", pady=(0, 20))
        tk.Label(title_bar, text="Sales Acc Receipt Gen", font=("Arial", 11, "bold"),
                 bg="#F0F0F0", fg=text_color).pack(anchor="w")

        def label(text, required=False):
            row = tk.Frame(main, bg="white")
            row.pack(fill="x", pady=(0, 2))
            lbl_text = f"{text} *" if required else text
            tk.Label(row, text=lbl_text, font=("Arial", 9, "bold"),
                     bg="white", fg=text_color).pack(side="left")

        def badge(value):
            f = tk.Frame(main, bg="white")
            f.pack(fill="x", pady=(0, 4))
            text_val = value if value else "—"
            tk.Label(f, text=text_val, bg="#FFF2D9", fg="#6B4F00",
                     font=("Arial", 8), padx=8, pady=3).pack(anchor="w")

        # --- Project Name ---
        label("Project Name", required=True)
        badge(self.row_data.get("project", ""))
        # self.project_var inited above
        self.project_cb = ttk.Combobox(main, textvariable=self.project_var,
                                        values=projects, state="readonly")
        self.project_cb.pack(fill="x", pady=(0, 14))
        self.project_cb.bind("<<ComboboxSelected>>", self._update_units)

        # --- Unit Number ---
        label("Unit Number", required=True)
        badge(self.row_data.get("unit", ""))
        # self.unit_var inited above
        self.unit_cb = ttk.Combobox(main, textvariable=self.unit_var, state="readonly")
        self.unit_cb.pack(fill="x", pady=(0, 14))
        self._update_units()

        # --- Amount Type ---
        label("Amount Type", required=True)
        # self.amount_type_var inited above
        radio_row = tk.Frame(main, bg="white")
        radio_row.pack(fill="x", pady=(4, 14))
        for val, txt in (("net", "Net"), ("gross", "Gross")):
            tk.Radiobutton(radio_row, text=txt, variable=self.amount_type_var,
                           value=val, bg="white", activebackground="white",
                           command=self._update_amount).pack(side="left", padx=(0, 20))

        # --- Amount ---
        label("Amount")
        # self.amount_var inited above
        self.amount_entry = ttk.Entry(main, textvariable=self.amount_var)
        self.amount_entry.pack(fill="x", pady=(4, 14))
        self._update_amount() 

        # --- Remarks ---
        label("Remarks")
        self.remarks_text = tk.Text(main, height=3, relief="solid", borderwidth=1,
                                     font=("Arial", 9))
        self.remarks_text.insert("1.0", self.row_data.get("remarks", ""))
        self.remarks_text.pack(fill="x", pady=(4, 20))

        # --- Footer ---
        sep = tk.Frame(main, bg="#E0E0E0", height=1)
        sep.pack(fill="x", pady=(0, 12))

        footer = tk.Frame(main, bg="white")
        footer.pack(fill="x")

        _make_button(footer, text="Mark as Non-SA", command=self._mark_non_sa,
                  bg="#FFF0F0", fg="#C0392B", relief="flat",
                  font=("Arial", 9), padx=10, pady=6).pack(side="left")

        _make_button(footer, text="Cancel", command=self.destroy,
                  bg="white", fg="#555", relief="solid", borderwidth=1,
                  font=("Arial", 9), padx=12, pady=5).pack(side="right", padx=(8, 0))

        _make_button(footer, text="  Save  ", command=self._save,
                  bg=accent_color, fg="white", relief="flat",
                  font=("Arial", 9, "bold"), padx=14, pady=6).pack(side="right")

        self._center_window()

    def _center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _update_units(self, event=None):
        project = self.project_var.get()
        self.unit_cb["values"] = units.get(project, [])

    def _update_amount(self):
        raw = self.row_data.get("credit", "0").replace(",", "").replace("₹", "").strip()
        try:
            val = float(raw)
        except ValueError:
            val = 0.0
        if self.amount_type_var.get() == "net":
            self.amount_var.set(f"{val * 0.9:.2f}")
        else:
            self.amount_var.set(f"{val:.2f}")

    def _save(self):
        if not self.project_var.get():
            messagebox.showwarning("Validation", "Please select a Project Name.", parent=self)
            return
        if not self.unit_var.get():
            messagebox.showwarning("Validation", "Please select a Unit Number.", parent=self)
            return
        data = {
            "project": self.project_var.get(),
            "unit": self.unit_var.get(),
            "amount_type": self.amount_type_var.get(),
            "amount": self.amount_var.get(),
            "remarks": self.remarks_text.get("1.0", "end-1c"),
        }
        print("Saved:", data)
        self.destroy()
        if self.callback:
            self.callback("Saved successfully!")

    def _mark_non_sa(self):
        if messagebox.askyesno("Confirm", "Mark this record as Non-SA?", parent=self):
            print("Marked as Non-SA")
            self.destroy()
            if self.callback:
                self.callback("Marked as Non-SA")


# ---------------------------------------------------------------------------
# Professional Scrollable Table
# ---------------------------------------------------------------------------
class ProfessionalTable(tk.Frame):
    """
    A two-canvas table with a frozen header and synchronised horizontal
    scrolling.  Mouse-wheel bindings are attached only to the body canvas
    widget (not bind_all) to avoid hijacking scroll events in dialogs.
    """

    def __init__(self, parent, headers, col_widths, *args, **kwargs):
        super().__init__(parent, bg="white", *args, **kwargs)
        self.headers = headers
        self.col_widths = col_widths

        # Scrollbars
        self.h_scroll = ttk.Scrollbar(self, orient="horizontal",
                                       command=self._scroll_x_both)
        self.v_scroll = ttk.Scrollbar(self, orient="vertical",
                                       command=self._scroll_y_body)

        # Canvases
        self.header_canvas = tk.Canvas(
            self, height=44, bg=table_header_bg, highlightthickness=0,
            xscrollcommand=self.h_scroll.set)
        self.body_canvas = tk.Canvas(
            self, bg="white", highlightthickness=0,
            xscrollcommand=self.h_scroll.set,
            yscrollcommand=self.v_scroll.set)

        # Inner frames
        self.header_frame = tk.Frame(self.header_canvas, bg=table_header_bg)
        self.body_frame = tk.Frame(self.body_canvas, bg="white")

        self.header_canvas.create_window((0, 0), window=self.header_frame, anchor="nw")
        self.body_canvas.create_window((0, 0), window=self.body_frame, anchor="nw")

        # Layout
        self.header_canvas.grid(row=0, column=0, sticky="ew")
        self.v_scroll.grid(row=0, column=1, rowspan=2, sticky="ns")
        self.body_canvas.grid(row=1, column=0, sticky="nsew")
        self.h_scroll.grid(row=2, column=0, sticky="ew")

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Configure scroll regions on resize
        self.body_frame.bind("<Configure>", self._on_body_configure)
        self.header_frame.bind("<Configure>", self._on_header_configure)

        # Mouse-wheel: bind only to body_canvas (not bind_all)
        self.body_canvas.bind("<Enter>", self._bind_mousewheel)
        self.body_canvas.bind("<Leave>", self._unbind_mousewheel)

        self._build_headers()

    # ---- Header construction ----
    def _build_headers(self):
        for i, text in enumerate(self.headers):
            cell = tk.Frame(self.header_frame, width=self.col_widths[i],
                            height=40, bg=table_header_bg)
            cell.pack_propagate(False)
            cell.grid(row=0, column=i, padx=1, pady=2, sticky="nsew")
            tk.Label(cell, text=text, bg=table_header_bg, fg="#495057",
                     font=("Arial", 9, "bold"), anchor="w", padx=6).pack(
                fill="both", expand=True)

    # ---- Scroll callbacks ----
    def _scroll_x_both(self, *args):
        self.header_canvas.xview(*args)
        self.body_canvas.xview(*args)

    def _scroll_y_body(self, *args):
        self.body_canvas.yview(*args)

    def _on_body_configure(self, _event):
        self.body_canvas.configure(scrollregion=self.body_canvas.bbox("all"))

    def _on_header_configure(self, _event):
        self.header_canvas.configure(scrollregion=self.header_canvas.bbox("all"))

    # ---- Mouse wheel (scoped to body_canvas hover) ----
    def _bind_mousewheel(self, _event):
        self.body_canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.body_canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.body_canvas.bind_all("<Button-5>", self._on_mousewheel)
        self.body_canvas.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel)
        self.body_canvas.bind_all("<Shift-Button-4>", self._on_shift_mousewheel)
        self.body_canvas.bind_all("<Shift-Button-5>", self._on_shift_mousewheel)

    def _unbind_mousewheel(self, _event):
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>",
                    "<Shift-MouseWheel>", "<Shift-Button-4>", "<Shift-Button-5>"):
            self.body_canvas.unbind_all(seq)

    def _on_mousewheel(self, event):
        delta = _calc_scroll_delta(event)
        if delta != 0:
            self.body_canvas.yview_scroll(delta, "units")

    def _on_shift_mousewheel(self, event):
        delta = _calc_scroll_delta(event)
        if delta != 0:
            self.header_canvas.xview_scroll(delta, "units")
            self.body_canvas.xview_scroll(delta, "units")


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------
class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sobha Reconciliation")
        self.geometry("1480x920")
        self.minsize(1160, 760)
        self.configure(bg="#d7dbe2")
        self.row_vars = []

        # Initialize variables
        self.selected_batch_var = tk.StringVar(value="")
        self.match_filter_var = tk.StringVar(value="ALL")
        self.all_rows = []
        self.batch_groups = []
        self.current_batch_count = 0
        self.current_sub_batch_count = 0

        # Initialize widgets to None to avoid AttributeErrors
        self.cards_frame: Optional[tk.Frame] = None
        self.cards_canvas: Optional[tk.Canvas] = None
        self.cards_canvas_window_id: Optional[int] = None
        self.row_count_label: Optional[tk.Label] = None
        self.section_count_label: Optional[tk.Label] = None
        self.status_bar: Optional[tk.Label] = None
        self.login_button: Optional[tk.Button] = None
        self._auth_probe_serial = 0
        self._session_anim_serial = 0
        self._login_display_name: Optional[str] = None
        self.clipboard_col_defs = load_clipboard_col_defs()
        self.bulk_paste_mode_var = tk.BooleanVar(value=self._load_bulk_paste_mode_default())

        # SOBHA Premium Theme Palette
        self.colors = {
            "primary_bg": "#5E5453",          # SOBHA Primary Header / Sidebar
            "white": "#FFFFFF",
            "page_bg": "#EAE5E4",             # Light gray page background
            "table_header_bg": "#D2D5DB",     # Table header gray
            "card_bg": "#F4F4F4",             # Light gray container card
            "dark_text": "#313232",           # Dark text
            "muted_text": "#9B9A9B",          # Medium gray text
            "border_gray": "#BCBABA",         # Border gray
            "secondary_text": "#777676",      # Secondary text gray
            "near_black": "#0F0F0F",          # Near-black text
            "frame_bg": "#EAE5E4",
            "card_border": "#BCBABA",
            "card_selected_border": "#5E5453",
            "card_header_bg": "#FFFFFF",
            "table_shell_bg": "#FFFFFF",
            "table_border": "#BCBABA",
            "row_bg_even": "#FFFFFF",
            "row_bg_odd": "#F8F9FA",
            "row_sep": "#E2E8F0",
            "title": "#313232",
            "text": "#313232",
            "muted": "#777676",
            "accent": "#2563EB",
            "success": "#16A34A",
            "pill_bg": "#D2D5DB",
            "selector_border": "#BCBABA",
            "selector_bg": "#FFFFFF",
            "selector_active": "#5E5453",
        }

        # Initialize Forest Theme (fallback to default if missing)
        style = ttk.Style()
        try:
            self.tk.call("source", "forest-light.tcl")
            style.theme_use("forest-light")
        except Exception:
            pass

        # Main Root Frame
        root_frame = tk.Frame(self, bg=self.colors["page_bg"])
        root_frame.pack(fill="both", expand=True)

        # ===================================================================
        # 1. TOP HEADER BAR (#5E5453) - Clean Brand & Functional Toolbar
        # ===================================================================
        header_bar = tk.Frame(root_frame, bg=self.colors["primary_bg"], padx=18, pady=10)
        header_bar.pack(fill="x")

        # Brand Title Left
        brand_left = tk.Frame(header_bar, bg=self.colors["primary_bg"])
        brand_left.pack(side="left")

        try:
            self.sobha_brand_img = tk.PhotoImage(file=p("sobha_logo_brand.png"))
            logo_lbl = tk.Label(brand_left, image=self.sobha_brand_img, bg=self.colors["primary_bg"])
            logo_lbl.pack(side="left", padx=(0, 10))
        except Exception:
            pass

        brand_text_wrap = tk.Frame(brand_left, bg=self.colors["primary_bg"])
        brand_text_wrap.pack(side="left")

        tk.Label(
            brand_text_wrap,
            text="Sobha Reconciliation",
            bg=self.colors["primary_bg"],
            fg=self.colors["white"],
            font=("Segoe UI", 13, "bold"),
            anchor="w",
        ).pack(anchor="w")

        tk.Label(
            brand_text_wrap,
            text="Sales Acc Receipt Gen",
            bg=self.colors["primary_bg"],
            fg="#D2D5DB",
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(anchor="w")

        # Functional Toolbar Actions Right
        header_actions = tk.Frame(header_bar, bg=self.colors["primary_bg"])
        header_actions.pack(side="right")

        _make_button(
            header_actions,
            text="Make Automation",
            command=self._submit_selection,
            relief="flat",
            cursor="hand2",
            bg=self.colors["accent"],
            fg="white",
            activebackground="#1D4ED8",
            font=("Segoe UI", 10, "bold"),
            padx=16,
            pady=6,
        ).pack(side="right", padx=(8, 0))

        export_btn = _make_button(
            header_actions,
            text="📥",
            command=self._export_to_excel,
            relief="flat",
            cursor="hand2",
            bg=self.colors["white"],
            fg=self.colors["dark_text"],
            activebackground="#F4F4F4",
            font=("Segoe UI Emoji", 11),
            padx=10,
            pady=4,
        )
        export_btn.pack(side="right", padx=(8, 0))
        self._bind_copy_button_tooltip(export_btn, "Export to Excel (.xlsx)")

        refresh_btn = _make_button(
            header_actions,
            text="🔄",
            command=self._load_transactions,
            cursor="hand2",
            relief="flat",
            bg="#0D9488",
            fg="white",
            activebackground="#0F766E",
            font=("Segoe UI Emoji", 11),
            padx=10,
            pady=4,
        )
        refresh_btn.pack(side="right", padx=(8, 0))
        self._bind_copy_button_tooltip(refresh_btn, "Refresh transactions")

        self.login_button = _make_button(
            header_actions,
            text="Login",
            command=self._run_login_automation,
            cursor="hand2",
            relief="flat",
            bg=self.colors["success"],
            fg="white",
            activebackground="#15803D",
            font=("Segoe UI", 9, "bold"),
            padx=14,
            pady=6,
        )
        self.login_button.pack(side="right", padx=(8, 0))

        _make_button(
            header_actions,
            text="⚙ Settings",
            command=self._open_clipboard_column_settings,
            cursor="hand2",
            relief="flat",
            bg=self.colors["pill_bg"],
            fg=self.colors["dark_text"],
            activebackground="#BCBABA",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=6,
        ).pack(side="right", padx=(8, 0))

        # ===================================================================
        # 2. CONTROLS & FILTER CARD (#F4F4F4)
        # ===================================================================
        body_content = tk.Frame(root_frame, bg=self.colors["page_bg"], padx=18, pady=14)
        body_content.pack(fill="both", expand=True)

        filter_card = tk.Frame(
            body_content,
            bg=self.colors["card_bg"],
            highlightbackground=self.colors["border_gray"],
            highlightthickness=1,
            padx=16,
            pady=12,
        )
        filter_card.pack(fill="x", pady=(0, 14))

        f_grid = tk.Frame(filter_card, bg=self.colors["card_bg"])
        f_grid.pack(fill="x")

        # Status Filter
        f_status_col = tk.Frame(f_grid, bg=self.colors["card_bg"])
        f_status_col.pack(side="left", padx=(0, 18))
        tk.Label(
            f_status_col,
            text="FILTER STATUS",
            bg=self.colors["card_bg"],
            fg=self.colors["secondary_text"],
            font=("Segoe UI", 8, "bold"),
        ).pack(anchor="w", pady=(0, 4))

        self.match_filter_var = tk.StringVar(value="All")
        status_cb = ttk.Combobox(
            f_status_col,
            textvariable=self.match_filter_var,
            state="readonly",
            values=["All", "Unmatched", "Posted", "Matched", "All Splits Completed"],
            width=20,
            font=("Segoe UI", 9),
        )
        status_cb.pack(anchor="w")
        status_cb.bind("<<ComboboxSelected>>", lambda _e: self._apply_filter())

        # Search Here Entry
        f_search_col = tk.Frame(f_grid, bg=self.colors["card_bg"])
        f_search_col.pack(side="left", fill="x", expand=True, padx=(0, 18))
        tk.Label(
            f_search_col,
            text="SEARCH TRANSACTIONS",
            bg=self.colors["card_bg"],
            fg=self.colors["secondary_text"],
            font=("Segoe UI", 8, "bold"),
        ).pack(anchor="w", pady=(0, 4))

        search_wrap = tk.Frame(
            f_search_col,
            bg=self.colors["white"],
            highlightbackground=self.colors["border_gray"],
            highlightthickness=1,
        )
        search_wrap.pack(fill="x")
        self.search_var = tk.StringVar()
        search_entry = tk.Entry(
            search_wrap,
            textvariable=self.search_var,
            bg=self.colors["white"],
            fg=self.colors["dark_text"],
            font=("Segoe UI", 9),
            bd=0,
            highlightthickness=0,
        )
        search_entry.pack(side="left", fill="x", expand=True, padx=8, pady=4)
        tk.Label(search_wrap, text="🔍", bg=self.colors["white"], fg=self.colors["muted_text"]).pack(side="right", padx=6)
        self.search_var.trace_add("write", lambda *args: self._apply_filter())

        # Bulk Paste Mode Checkbox
        f_chk_col = tk.Frame(f_grid, bg=self.colors["card_bg"])
        f_chk_col.pack(side="right")
        tk.Checkbutton(
            f_chk_col,
            text="Bulk paste mode",
            variable=self.bulk_paste_mode_var,
            command=self._persist_bulk_paste_mode,
            bg=self.colors["card_bg"],
            activebackground=self.colors["card_bg"],
            fg=self.colors["dark_text"],
            selectcolor="white",
            font=("Segoe UI", 9),
            cursor="hand2",
        ).pack(anchor="e", pady=(14, 0))

        # ===================================================================
        # 3. TRANSACTIONS VIEW CANVAS AREA (#EAE5E4)
        # ===================================================================
        body_split = tk.Frame(body_content, bg=self.colors["page_bg"])
        body_split.pack(fill="both", expand=True)
        body_split.grid_rowconfigure(0, weight=1)
        body_split.grid_columnconfigure(0, weight=1)

        cards_host = tk.Frame(body_split, bg=self.colors["page_bg"])
        cards_host.grid(row=0, column=0, sticky="nsew")
        cards_host.grid_rowconfigure(0, weight=1)
        cards_host.grid_columnconfigure(0, weight=1)

        self.cards_canvas = tk.Canvas(
            cards_host,
            bg=self.colors["page_bg"],
            highlightthickness=0,
            bd=0,
        )
        cards_scroll = tk.Scrollbar(cards_host, orient="vertical", command=self.cards_canvas.yview)
        self.cards_canvas.configure(yscrollcommand=cards_scroll.set)

        self.cards_canvas.grid(row=0, column=0, sticky="nsew")
        cards_scroll.grid(row=0, column=1, sticky="ns")

        self.cards_frame = tk.Frame(self.cards_canvas, bg=self.colors["page_bg"])
        self.cards_canvas_window_id = self.cards_canvas.create_window((0, 0), window=self.cards_frame, anchor="nw")
        self.cards_frame.bind("<Configure>", self._on_cards_frame_configure)
        self.cards_canvas.bind("<Configure>", self._on_cards_canvas_configure)
        self.bind_all("<MouseWheel>", self._on_cards_mousewheel)
        self.bind_all("<Button-4>", self._on_cards_mousewheel)
        self.bind_all("<Button-5>", self._on_cards_mousewheel)

        # Footer Status Bar
        footer = tk.Frame(root_frame, bg=self.colors["page_bg"], padx=18, pady=8)
        footer.pack(fill="x")

        self.section_count_label = tk.Label(
            footer,
            text="Sales Acc Receipt Gen (0 batches / 0 sub-batches)",
            font=("Segoe UI", 9),
            fg=self.colors["muted"],
            bg=self.colors["page_bg"],
        )
        self.section_count_label.pack(side="left")

        self.row_count_label = tk.Label(
            footer,
            text="0 batches | 0 sub-batches",
            font=("Segoe UI", 9),
            fg=self.colors["muted"],
            bg=self.colors["page_bg"],
        )
        self.row_count_label.pack(side="left", padx=(12, 0))

        self.status_bar = tk.Label(
            footer,
            text="Ready",
            fg=self.colors["muted"],
            bg=self.colors["page_bg"],
            font=("Segoe UI", 9),
        )
        self.status_bar.pack(side="right")
        self.row_count_label.pack(side="left")

        self.status_bar = tk.Label(
            footer,
            text="Ready",
            fg=self.colors["muted"],
            bg=self.colors["page_bg"],
            font=("Segoe UI", 9),
        )
        self.status_bar.pack(side="right")
        self.status_bar.pack(side="left", padx=(10, 0))

        self._browser_check_prompted = False
        self._set_login_button_state(False)
        self.after(250, self._refresh_login_button_async)
        self.after(600, self._check_browser_ready_on_launch)
        self.after(200, self._load_transactions)

    def _set_login_button_state(self, valid: bool, display_name: Optional[str] = None):
        if self.login_button is None:
            return

        normalized_name = str(display_name or "").strip() or None
        self._login_display_name = normalized_name if valid else None
        if valid and normalized_name and len(normalized_name) > 22:
            button_text = f"{normalized_name[:19]}..."
        else:
            button_text = (normalized_name or "Logged In") if valid else "Login"
        button_bg = self.colors["accent"] if valid else self.colors["success"]
        active_bg = "#264fdf" if valid else "#14913f"
        self.login_button.config(
            state="normal",
            cursor="hand2",
            text=button_text,
            bg=button_bg,
            activebackground=active_bg,
        )

    def _apply_auth_result(self, auth_result: Optional[dict], *, invalidate_pending: bool = False):
        if invalidate_pending:
            self._auth_probe_serial += 1

        valid = bool(isinstance(auth_result, dict) and auth_result.get("valid"))
        display_name = None
        if valid and isinstance(auth_result, dict):
            display_name = str(auth_result.get("display_name") or "").strip() or None
        self._set_login_button_state(valid, display_name)

    def _apply_auth_probe_result(self, probe_id: int, auth_result: Optional[dict]):
        if probe_id != self._auth_probe_serial:
            return
        self._apply_auth_result(auth_result)

    def _refresh_login_button_async(self):
        if automation_module is None or not hasattr(automation_module, "probe_saved_session"):
            self._set_login_button_state(False)
            return

        self._auth_probe_serial += 1
        probe_id = self._auth_probe_serial
        self._session_anim_serial += 1
        anim_id = self._session_anim_serial

        self._is_checking_session = True
        self._session_check_step = 3
        if self.login_button:
            self.login_button.config(state="disabled", bg="#6c757d", cursor="watch", text="Checking session...")

        def update_button_animation():
            if anim_id != self._session_anim_serial or not getattr(self, "_is_checking_session", False):
                return
            if not self.login_button:
                return
            if self._session_check_step > 0:
                self.login_button.config(text=f"Session checking {self._session_check_step}")
                self._session_check_step -= 1
                self.after(1000, update_button_animation)
            else:
                self.login_button.config(text="Opening...")

        update_button_animation()

        def probe_task():
            try:
                auth_result = automation_module.probe_saved_session(headless=True)
            except Exception as err:
                print(f"Auth status probe failed: {err}")
                auth_result = {"valid": False, "display_name": None, "reason": str(err)}

            def on_complete():
                self._session_anim_serial += 1
                self._is_checking_session = False
                if self.login_button:
                    self.login_button.config(state="normal", cursor="hand2")
                self._apply_auth_probe_result(probe_id, auth_result)

            self.after(0, on_complete)

        threading.Thread(target=probe_task, daemon=True).start()

    def _handle_session_expired(self, err: Exception):
        self._apply_auth_result({"valid": False}, invalidate_pending=True)
        detail = str(err).strip() or "Saved D365 session expired."
        if "click login" not in detail.lower():
            detail = f"{detail}\n\nClick Login and sign in again."
        messagebox.showerror("Session Expired", detail)

    def _refresh_cards_canvas(self, _event=None):
        if not self.cards_canvas or not self.cards_frame:
            return
        self.cards_frame.update_idletasks()
        canvas_width = max(self.cards_canvas.winfo_width(), 1)
        if canvas_width > 1 and self.cards_canvas_window_id is not None:
            self.cards_canvas.itemconfigure(self.cards_canvas_window_id, width=canvas_width)
        bbox = self.cards_canvas.bbox("all")
        if bbox:
            self.cards_canvas.configure(scrollregion=bbox)

    def _on_cards_frame_configure(self, _event=None):
        self._refresh_cards_canvas()

    def _on_cards_canvas_configure(self, event):
        if self.cards_canvas and self.cards_canvas_window_id is not None:
            self.cards_canvas.itemconfigure(self.cards_canvas_window_id, width=event.width)
        self._refresh_cards_canvas()

    # Scroll is now permanently bound at init — no Enter/Leave needed

    def _on_cards_mousewheel(self, event):
        if not self.cards_canvas:
            return
        delta = _calc_scroll_delta(event)
        if delta != 0:
            self.cards_canvas.yview_scroll(delta, "units")

    def _on_batch_canvas_configure(self, event):
        pass

    def _set_match_filter(self, filter_name: str):
        self.match_filter_var.set(filter_name)
        self._apply_filter()

    def _update_filter_tab_styles(self):
        pass

    def _row_is_matched(self, row: dict) -> bool:
        required = (
            "value_date",
            "account",
            "credit",
            "offset_account",
            "method_of_payment",
            "reference_date",
            "payment_reference",
        )
        return all(str(row.get(key, "")).strip() for key in required)

    def _row_key(self, row: dict) -> str:
        uuid = str(row.get("uuid", "")).strip()
        if uuid:
            return uuid
        return "|".join(
            [
                str(row.get("batch_id", "")).strip(),
                str(row.get("sub_batch_id", "")).strip(),
                str(row.get("value_date", "")).strip(),
                str(row.get("account", "")).strip(),
                str(row.get("credit", "")).strip(),
                str(row.get("payment_reference", "")).strip(),
            ]
        )

    def _sanitize_amount(self, value: str) -> str:
        if value is None:
            return ""
        raw = str(value).replace("₹", "").replace(" ", "").strip()
        return "".join(ch for ch in raw if ch.isdigit() or ch in {",", ".", "-"})

    def _extract_records(self, payload):
        if isinstance(payload, list):
            return payload
        if isinstance(payload, dict):
            data = payload.get("data")
            if isinstance(data, list):
                return data
            for key in ("data", "result", "rows", "records"):
                if isinstance(payload.get(key), list):
                    return payload[key]
        return []

    def _map_transaction(self, txn: dict, batch_id: str = "", sub_batch_id: str = "") -> dict:
        mapped_batch_id = str(txn.get("batch_id", "")).strip() or batch_id or "UNASSIGNED"
        mapped_sub_batch_id = str(txn.get("sub_batch_id", "")).strip() or sub_batch_id or mapped_batch_id
        account_date = format_d365_date(txn.get("account_date", ""))
        payment_reference = str(txn.get("transaction_description", "")).strip()
        mapped = {
            "uuid": str(txn.get("uuid", "")).strip(),
            "batch_id": mapped_batch_id,
            "sub_batch_id": mapped_sub_batch_id,
            "date": account_date,
            "value_date": account_date,
            "voucher": str(txn.get("receipt_number", "")).strip(),
            "company": (
                automation_module.d365_company_from_config()
                if automation_module and hasattr(automation_module, "d365_company_from_config")
                else ""
            ),
            "account": str(txn.get("account_number", "")).strip(),
            "account_name": "",
            "payee_name": "",
            "invoice": "",
            "description": payment_reference,
            "debit": "",
            "credit": self._sanitize_amount(txn.get("transaction_amount", "")),
            "currency": "",
            "sales_order_id": "",
            "bank_account": "",
            "offset_account_type": "",
            "offset_account": str(txn.get("offset_account", "")).strip(),
            "method_of_payment": str(txn.get("mode_of_transaction", "")).strip(),
            "payment_status": "",
            "demand_number": "",
            "reference_date": account_date,
            "payment_reference": payment_reference,
            "use_deposit_slip": "",
            "crm_transaction_type": str(txn.get("txn_source", "")).strip(),
            "original_payment_voucher": "",
            "reversal_payment_voucher": "",
        }
        for raw_key, raw_value in txn.items():
            if raw_key in mapped:
                continue
            mapped[raw_key] = "" if raw_value is None else str(raw_value).strip()
        return mapped

    def _extract_batch_groups(self, payload):
        if isinstance(payload, dict):
            data = payload.get("data")
            if isinstance(data, dict) and isinstance(data.get("batches"), list):
                batch_groups = []
                for batch_index, batch in enumerate(data.get("batches", []), start=1):
                    batch_id = str(batch.get("batch_id", "")).strip() or f"UNASSIGNED_{batch_index}"
                    sub_batch_groups = []
                    for sub_index, sub_batch in enumerate(batch.get("sub_batches", []), start=1):
                        sub_batch_id = str(sub_batch.get("sub_batch_id", "")).strip() or f"{batch_id}_{sub_index}"
                        transactions = [
                            self._map_transaction(txn, batch_id, sub_batch_id)
                            for txn in sub_batch.get("transactions", [])
                        ]
                        sub_batch_groups.append(
                            {
                                "batch_id": batch_id,
                                "sub_batch_id": sub_batch_id,
                                "transactions": transactions,
                            }
                        )
                    batch_groups.append(
                        {
                            "batch_id": batch_id,
                            "sub_batches": sub_batch_groups,
                        }
                    )
                return batch_groups

        grouped = {}
        for txn in self._extract_records(payload):
            mapped = self._map_transaction(txn)
            batch_id = mapped["batch_id"]
            sub_batch_id = mapped["sub_batch_id"]
            batch_group = grouped.setdefault(batch_id, {})
            batch_group.setdefault(sub_batch_id, []).append(mapped)

        batch_groups = []
        for batch_id, sub_batches in grouped.items():
            batch_groups.append(
                {
                    "batch_id": batch_id,
                    "sub_batches": [
                        {
                            "batch_id": batch_id,
                            "sub_batch_id": sub_batch_id,
                            "transactions": transactions,
                        }
                        for sub_batch_id, transactions in sub_batches.items()
                    ],
                }
            )
        return batch_groups

    def _batch_transaction_count(self, batch_info: dict) -> int:
        return sum(len(sub_batch.get("transactions", [])) for sub_batch in batch_info.get("sub_batches", []))

    def _selected_batch_group(self):
        selected_batch_id = self.selected_batch_var.get().strip()
        for batch_info in self.batch_groups:
            if str(batch_info.get("batch_id", "")).strip() == selected_batch_id:
                return batch_info
        return None

    def _transactions_for_batch(self, batch_info: dict) -> list:
        return [
            dict(record)
            for sub_batch in batch_info.get("sub_batches", [])
            for record in sub_batch.get("transactions", [])
        ]

    def _copy_batch_transactions_to_excel_clipboard(self, batch_info: dict) -> int:
        transactions = self._transactions_for_batch(batch_info)
        if not transactions:
            return 0

        clipboard_text = build_excel_clipboard_text(transactions, self.clipboard_col_defs)
        self.clipboard_clear()
        self.clipboard_append(clipboard_text)
        self.update_idletasks()
        return len(transactions)

    def _copy_selected_batch_to_excel_clipboard(self) -> int:
        self._sync_current_edits()
        selected_group = self._selected_batch_group()
        if selected_group is None:
            return 0
        return self._copy_batch_transactions_to_excel_clipboard(selected_group)

    def _copy_batch_to_excel_clipboard(self, batch_info: dict) -> None:
        self._sync_current_edits()
        copied_count = self._copy_batch_transactions_to_excel_clipboard(batch_info)
        if copied_count <= 0:
            messagebox.showwarning("Copy", "No transactions to copy for this batch.")
            return
        if self.status_bar:
            self.status_bar.config(
                text=f"{self._status_text()} | Copied {copied_count} rows to clipboard (Excel)"
            )

    def _open_clipboard_column_settings(self):
        def on_save(col_defs):
            self.clipboard_col_defs = normalize_col_defs(col_defs)
            if self.status_bar:
                self.status_bar.config(
                    text=f"{self._status_text()} | Copy columns saved ({len(col_defs)} columns)"
                )

        dlg = ClipboardColumnSettingsDialog(self, self.clipboard_col_defs, on_save=on_save)
        self.wait_window(dlg)

    def _bind_copy_button_tooltip(self, widget: tk.Widget, text: str) -> None:
        tooltip = {"window": None}

        def show_tooltip(_event):
            if tooltip["window"] is not None:
                return
            tip = tk.Toplevel(widget)
            tip.wm_overrideredirect(True)
            tip.wm_geometry(f"+{widget.winfo_rootx() + 24}+{widget.winfo_rooty() - 28}")
            tk.Label(
                tip,
                text=text,
                bg="#1f2937",
                fg="white",
                font=("Segoe UI", 9),
                padx=8,
                pady=4,
            ).pack()
            tooltip["window"] = tip

        def hide_tooltip(_event):
            if tooltip["window"] is not None:
                tooltip["window"].destroy()
                tooltip["window"] = None

        widget.bind("<Enter>", show_tooltip)
        widget.bind("<Leave>", hide_tooltip)

    def _refresh_match_counts(self):
        batch_count = len(self.batch_groups)
        sub_batch_count = sum(len(batch.get("sub_batches", [])) for batch in self.batch_groups)
        if self.section_count_label:
            self.section_count_label.configure(
                text=f"Sales Acc Receipt Gen ({batch_count} batches / {sub_batch_count} sub-batches)"
            )

    def _status_text(self) -> str:
        selected_batch_id = self.selected_batch_var.get().strip()
        base = f"Showing {self.current_batch_count} batches / {self.current_sub_batch_count} sub-batches"
        if selected_batch_id:
            return f"{base} | Selected batch: {selected_batch_id}"
        return f"{base} | Selected batch: None"

    def _sync_current_edits(self):
        for row_widgets in self.row_vars:
            data_ref = row_widgets.get("data")
            if not isinstance(data_ref, dict):
                continue
            for key in KEY_MAP[1:]:
                if key in row_widgets:
                    value = row_widgets[key].get()
                    if key in DATE_FIELD_KEYS:
                        value = format_d365_date(value)
                    data_ref[key] = value
            if "sub_batch_id" in row_widgets:
                data_ref["sub_batch_id"] = row_widgets["sub_batch_id"].get()

    def _render_rows(self, batch_groups):
        if not self.cards_frame:
            return

        for widget in self.cards_frame.winfo_children():
            widget.destroy()
        self.row_vars.clear()

        self.current_batch_count = len(batch_groups)
        self.current_sub_batch_count = sum(len(batch.get("sub_batches", [])) for batch in batch_groups)

        if not batch_groups:
            empty = tk.Label(
                self.cards_frame,
                text="No records for current filter",
                bg=self.colors["frame_bg"],
                fg=self.colors["muted"],
                font=("Segoe UI", 11),
                pady=24,
            )
            empty.pack(fill="x")
            if self.row_count_label:
                self.row_count_label.config(text="0 batches | 0 sub-batches")
            return

        col_weights = {
            "value_date": 1,
            "account": 1,
            "credit": 1,
            "offset_account": 2,
            "method_of_payment": 1,
            "reference_date": 1,
            "payment_reference": 3,
        }
        col_defs = [
            (key, label, col_weights.get(key, 1))
            for key, label in DISPLAY_COL_DEFS
        ]

        selected_batch_id = self.selected_batch_var.get().strip()
        for batch_info in batch_groups:
            batch_id = str(batch_info.get("batch_id", "")).strip() or "UNASSIGNED"
            batch_selected = batch_id == selected_batch_id
            sub_batches = batch_info.get("sub_batches", [])
            total_transactions = self._batch_transaction_count(batch_info)

            card = tk.Frame(
                self.cards_frame,
                bg=self.colors["card_bg"],
                highlightbackground=self.colors["card_selected_border"] if batch_selected else self.colors["card_border"],
                highlightthickness=2 if batch_selected else 1,
                bd=0,
                padx=10,
                pady=10,
            )
            card.pack(fill="x", pady=(0, 12))

            header = tk.Frame(card, bg=self.colors["card_header_bg"], height=42)
            header.pack(fill="x")
            header.pack_propagate(False)

            radio_btn = _make_button(
                header,
                text="",
                cursor="hand2",
                relief="solid",
                bg=self.colors["selector_bg"],
                activebackground=self.colors["selector_bg"],
                font=("Segoe UI", 10, "bold"),
                width=2,
                padx=0,
                pady=1,
                borderwidth=1,
                highlightthickness=0,
                command=lambda bid=batch_id: self._toggle_batch_selection(bid),
            )
            radio_btn.pack(side="left", padx=(0, 8))
            self._style_batch_radio_button(radio_btn, batch_selected)

            title_wrap = tk.Frame(header, bg=self.colors["card_header_bg"], cursor="hand2")
            title_wrap.pack(side="left", fill="x", expand=True, padx=(6, 0))
            lbl1 = tk.Label(
                title_wrap,
                text=batch_id,
                bg=self.colors["card_header_bg"],
                fg=self.colors["title"],
                font=("Segoe UI", 9, "bold"),
                anchor="w",
                cursor="hand2",
            )
            lbl1.pack(anchor="w")
            lbl2 = tk.Label(
                title_wrap,
                text=f"{len(sub_batches)} sub-batches",
                bg=self.colors["card_header_bg"],
                fg=self.colors["muted"],
                font=("Segoe UI", 8),
                anchor="w",
                cursor="hand2",
            )
            lbl2.pack(anchor="w")

            for widget in (header, title_wrap, lbl1, lbl2):
                widget.bind("<Button-1>", lambda _e, bid=batch_id: self._toggle_batch_selection(bid))

            if batch_selected:
                copy_btn = _make_button(
                    header,
                    text="\U0001f4cb",
                    cursor="hand2",
                    relief="flat",
                    bg=self.colors["card_header_bg"],
                    fg=self.colors["accent"],
                    activebackground=self.colors["card_header_bg"],
                    activeforeground="#1d4ed8",
                    font=("Segoe UI Emoji", 13),
                    padx=6,
                    pady=0,
                    borderwidth=0,
                    highlightthickness=0,
                    command=lambda bi=batch_info: self._copy_batch_to_excel_clipboard(bi),
                )
                copy_btn.pack(side="right", padx=(0, 4))
                self._bind_copy_button_tooltip(copy_btn, "Copy all transactions to Excel (clipboard)")

            tk.Label(
                header,
                text=f"{total_transactions} transactions",
                bg=self.colors["card_header_bg"],
                fg=self.colors["text"],
                font=("Segoe UI", 10, "bold"),
            ).pack(side="right", padx=(0, 4))

            for sub_batch in sub_batches:
                sub_batch_id = str(sub_batch.get("sub_batch_id", "")).strip() or batch_id
                sub_rows = sub_batch.get("transactions", [])

                sub_card = tk.Frame(
                    card,
                    bg=self.colors["table_shell_bg"],
                    highlightbackground=self.colors["table_border"],
                    highlightthickness=1,
                    bd=0,
                    padx=8,
                    pady=8,
                )
                sub_card.pack(fill="x", pady=(8, 2))

                sub_header = tk.Frame(sub_card, bg=self.colors["table_header_bg"], height=34)
                sub_header.pack(fill="x")
                sub_header.pack_propagate(False)

                sub_title = tk.Frame(sub_header, bg=self.colors["table_header_bg"])
                sub_title.pack(side="left", fill="x", expand=True, padx=(4, 0))
                tk.Label(
                    sub_title,
                    text=sub_batch_id,
                    bg=self.colors["table_header_bg"],
                    fg=self.colors["title"],
                    font=("Segoe UI", 9, "bold"),
                    anchor="w",
                ).pack(anchor="w")

                tk.Label(
                    sub_header,
                    text=f"{len(sub_rows)} transactions",
                    bg=self.colors["table_header_bg"],
                    fg=self.colors["text"],
                    font=("Segoe UI", 9, "bold"),
                ).pack(side="right", padx=(0, 4))

                table = tk.Frame(sub_card, bg=self.colors["table_shell_bg"], bd=0)
                table.pack(fill="x", padx=4, pady=(8, 2))

                header_row = tk.Frame(table, bg=self.colors["table_header_bg"], pady=4)
                header_row.pack(fill="x")
                for col_idx, (_key, label, weight) in enumerate(col_defs):
                    header_row.grid_columnconfigure(col_idx, weight=weight, uniform="batch_cols")
                    tk.Label(
                        header_row,
                        text=label,
                        bg=self.colors["table_header_bg"],
                        fg=self.colors["muted"],
                        font=("Segoe UI", 8, "bold"),
                        anchor="w",
                    ).grid(row=0, column=col_idx, sticky="ew", padx=(6, 8), pady=(2, 2))

                for row_index, record in enumerate(sub_rows):
                    row_key = self._row_key(record)
                    row_widgets = {
                        "data": record,
                        "row_key": row_key,
                        "uuid": tk.StringVar(value=str(record.get("uuid", ""))),
                        "batch_id": tk.StringVar(value=str(record.get("batch_id", ""))),
                        "sub_batch_id": tk.StringVar(value=str(record.get("sub_batch_id", ""))),
                    }
                    for key in KEY_MAP[2:]:
                        value = str(record.get(key, ""))
                        if key in DATE_FIELD_KEYS:
                            value = format_d365_date(value)
                            record[key] = value
                        row_widgets[key] = tk.StringVar(value=value)
                    self.row_vars.append(row_widgets)

                    row_bg = self.colors["row_bg_even"] if row_index % 2 == 0 else self.colors["row_bg_odd"]
                    row_frame = tk.Frame(table, bg=row_bg)
                    row_frame.pack(fill="x")
                    for col_idx, (key, _label, weight) in enumerate(col_defs):
                        row_frame.grid_columnconfigure(col_idx, weight=weight, uniform="batch_cols")
                        if key in {"value_date", "reference_date", "account", "credit", "offset_account", "method_of_payment", "payment_reference"}:
                            entry = tk.Entry(
                                row_frame,
                                textvariable=row_widgets[key],
                                relief="flat",
                                bd=0,
                                highlightthickness=0,
                                bg=row_bg,
                                fg=self.colors["text"],
                                insertbackground=self.colors["text"],
                                font=("Segoe UI", 11),
                            )
                            entry.grid(row=0, column=col_idx, sticky="ew", padx=(6, 8), pady=(6, 6))
                        else:
                            tk.Label(
                                row_frame,
                                textvariable=row_widgets[key],
                                bg=row_bg,
                                fg=self.colors["text"],
                                font=("Segoe UI", 11),
                                anchor="w",
                            ).grid(row=0, column=col_idx, sticky="ew", padx=(6, 8), pady=(7, 7))

                    if row_index < len(sub_rows) - 1:
                        sep = tk.Frame(table, height=1, bg=self.colors["row_sep"])
                        sep.pack(fill="x")

        if self.row_count_label:
            self.row_count_label.config(
                text=f"{self.current_batch_count} batches | {self.current_sub_batch_count} sub-batches"
            )
        self._refresh_cards_canvas()
        self.after_idle(self._refresh_cards_canvas)

    def _style_batch_radio_button(self, button: tk.Button, selected: bool):
        if selected:
            button.configure(
                text="\u2713",
                bg=self.colors["selector_active"],
                fg="white",
                activebackground=self.colors["selector_active"],
                activeforeground="white",
                highlightbackground=self.colors["selector_active"],
                highlightcolor=self.colors["selector_active"],
                relief="solid",
                borderwidth=1,
            )
        else:
            button.configure(
                text="",
                bg=self.colors["selector_bg"],
                fg=self.colors["selector_bg"],
                activebackground=self.colors["selector_bg"],
                activeforeground=self.colors["selector_bg"],
                highlightbackground=self.colors["selector_border"],
                highlightcolor=self.colors["selector_border"],
                relief="solid",
                borderwidth=1,
            )

    def _toggle_batch_selection(self, batch_id: str):
        batch_key = str(batch_id).strip()
        if not batch_key:
            return
        # Toggle: deselect if already selected, otherwise select
        if self.selected_batch_var.get().strip() == batch_key:
            self.selected_batch_var.set("")
        else:
            self.selected_batch_var.set(batch_key)
        self._apply_filter()

    def _confirm_automation_dialog(self, batch_id: str, sub_batch_count: int, transaction_count: int) -> bool:
        dlg = tk.Toplevel(self)
        dlg.title("Confirm Automation")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()
        dlg.configure(bg="#eef2f8")

        card = tk.Frame(
            dlg,
            bg="white",
            highlightbackground="#cfd8ea",
            highlightthickness=1,
            bd=0,
            padx=24,
            pady=22,
        )
        card.pack(fill="both", expand=True, padx=14, pady=14)

        content = tk.Frame(card, bg="white")
        content.pack(fill="both", expand=True)

        tk.Label(
            content,
            text="Are you sure to Proceed Automation?",
            bg="white",
            fg=self.colors["title"],
            font=("Segoe UI", 15, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(4, 8))

        summary_rows = [
            ("Main Batch", batch_id),
            ("Sub-Batches", str(sub_batch_count)),
            ("Transactions", str(transaction_count)),
            ("Bulk paste mode", "ON" if self.bulk_paste_mode_var.get() else "OFF"),
        ]
        if self.bulk_paste_mode_var.get():
            chunk_count = max(1, math.ceil(transaction_count / 20)) if transaction_count else 0
            summary_rows.append(("Paste batches (20 rows)", str(chunk_count)))
        for label, value in summary_rows:
            row = tk.Frame(content, bg="white")
            row.pack(fill="x", pady=(0, 6))
            tk.Label(
                row,
                text=f"{label}:",
                bg="white",
                fg=self.colors["muted"],
                font=("Segoe UI", 10, "bold"),
                width=20,
                anchor="w",
            ).pack(side="left")
            tk.Label(
                row,
                text=value,
                bg="white",
                fg=self.colors["title"],
                font=("Segoe UI", 10),
                anchor="w",
                wraplength=360,
                justify="left",
            ).pack(side="left", fill="x", expand=True)

        tk.Label(
            content,
            text="Please confirm to continue.",
            bg="white",
            fg=self.colors["muted"],
            font=("Segoe UI", 10),
            anchor="w",
        ).pack(fill="x", pady=(10, 12))

        result = {"ok": False}
        btn_row = tk.Frame(card, bg="white", pady=6)
        btn_row.pack(fill="x", side="bottom")

        def on_confirm():
            result["ok"] = True
            dlg.destroy()

        def on_cancel():
            result["ok"] = False
            dlg.destroy()

        _make_button(
            btn_row,
            text="Continue",
            command=on_confirm,
            cursor="hand2",
            relief="flat",
            bg="#16a34a",
            fg="white",
            activebackground="#15803d",
            activeforeground="white",
            font=("Segoe UI", 11, "bold"),
            padx=22,
            pady=9,
        ).pack(side="right")

        _make_button(
            btn_row,
            text="Cancel",
            command=on_cancel,
            cursor="hand2",
            relief="flat",
            bg="#e5e7eb",
            fg=self.colors["title"],
            activebackground="#d1d5db",
            activeforeground=self.colors["title"],
            font=("Segoe UI", 11, "bold"),
            padx=22,
            pady=9,
        ).pack(side="right", padx=(0, 10))

        dlg.update_idletasks()
        dlg_w = max(620, card.winfo_reqwidth() + 28)
        dlg_h = max(380, card.winfo_reqheight() + 28)
        x = self.winfo_rootx() + (self.winfo_width() - dlg_w) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dlg_h) // 2
        dlg.geometry(f"{dlg_w}x{dlg_h}+{x}+{y}")

        dlg.wait_window()
        return result["ok"]

    def _load_transactions(self):
        def worker():
            try:
                req = urllib.request.Request(
                    API_TRANSACTIONS_URL,
                    headers={"Authorization": f"Bearer {API_TOKEN}"},
                    method="GET",
                )
                with urllib.request.urlopen(req, timeout=30) as resp:
                    raw = resp.read().decode("utf-8")
                payload = json.loads(raw)
                batch_groups = self._extract_batch_groups(payload)
                flat_rows = [
                    transaction
                    for batch in batch_groups
                    for sub_batch in batch.get("sub_batches", [])
                    for transaction in sub_batch.get("transactions", [])
                ]
                sub_batch_count = sum(len(batch.get("sub_batches", [])) for batch in batch_groups)
                print(
                    f"API loaded {len(batch_groups)} batches, "
                    f"{sub_batch_count} sub-batches, {len(flat_rows)} transactions"
                )
                self.after(
                    0,
                    lambda groups=batch_groups, rows=flat_rows: self._apply_loaded_transactions(groups, rows),
                )
            except urllib.error.HTTPError as err:
                print(f"API HTTP error: {err.code} {err.reason}")
                self.after(0, lambda e=err: messagebox.showerror("API Error", f"HTTP {e.code}: {e.reason}"))
            except urllib.error.URLError as err:
                print(f"API network error: {err.reason}")
                self.after(0, lambda e=err: messagebox.showerror("API Error", f"Network error: {e.reason}"))
            except Exception as err:
                print(f"API fetch failed: {err}")
                self.after(0, lambda e=err: messagebox.showerror("API Error", f"Failed to fetch transactions:\n{e}"))

        threading.Thread(target=worker, daemon=True).start()

    def _apply_loaded_transactions(self, batch_groups, data):
        self.batch_groups = batch_groups
        self.all_rows = data
        valid_batch_ids = {
            str(batch.get("batch_id", "")).strip()
            for batch in batch_groups
        }
        current_selected = self.selected_batch_var.get().strip()
        if batch_groups and (not current_selected or current_selected not in valid_batch_ids):
            first_id = str(batch_groups[0].get("batch_id", "")).strip()
            self.selected_batch_var.set(first_id)
        self._refresh_match_counts()
        self._apply_filter()

    def _select_all_visible_rows(self):
        """Select the first batch if none is selected."""
        selected_group = self._selected_batch_group()
        if selected_group is None and self.batch_groups:
            self.selected_batch_var.set(str(self.batch_groups[0].get("batch_id", "")).strip())
            self._apply_filter()

    def _apply_filter(self):
        if not self.cards_frame:
            return

        self._sync_current_edits()
        self._render_rows(self.batch_groups)
    def _export_to_excel(self):
        try:
            from tkinter import filedialog
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if not path:
                return
            if Workbook is None:
                messagebox.showerror("Export Error", "openpyxl is not installed.")
                return
            wb = Workbook()
            ws = wb.active
            ws.title = "RERA Transactions"
            ws.append(["Batch ID", "Value Date", "Account", "Credit", "Offset Account", "Method of Payment", "Reference Date", "Payment Reference"])
            for row in self.all_rows:
                ws.append([
                    row.get("batch_id", ""),
                    row.get("value_date", ""),
                    row.get("account", ""),
                    row.get("credit", ""),
                    row.get("offset_account", ""),
                    row.get("method_of_payment", ""),
                    row.get("reference_date", ""),
                    row.get("payment_reference", ""),
                ])
            wb.save(path)
            messagebox.showinfo("Export Success", f"Successfully exported transactions to:\n{path}")
        except Exception as err:
            messagebox.showerror("Export Error", f"Failed to export: {err}")

    def _submit_selection(self):
        self._sync_current_edits()
        selected_group = self._selected_batch_group()
        if selected_group is None and self.batch_groups:
            first_id = str(self.batch_groups[0].get("batch_id", "")).strip()
            self.selected_batch_var.set(first_id)
            self._apply_filter()
            selected_group = self._selected_batch_group()

        if selected_group is None:
            messagebox.showwarning("No Selection", "Please select one main batch to submit.")
            return

        batch_id = str(selected_group.get("batch_id", "")).strip() or "UNASSIGNED"
        selected = [
            dict(record)
            for sub_batch in selected_group.get("sub_batches", [])
            for record in sub_batch.get("transactions", [])
        ]
        for record in selected:
            for key in DATE_FIELD_KEYS:
                if key in record:
                    record[key] = format_d365_date(record.get(key, ""))
        sub_batch_count = len(selected_group.get("sub_batches", []))
        transaction_count = len(selected)

        if not selected:
            messagebox.showwarning("No Transactions", "The selected batch does not contain any transactions.")
            return

        if self._confirm_automation_dialog(batch_id, sub_batch_count, transaction_count):
            if not self._validate_config_for_action(require_auth_state=True):
                return
            bulk_mode = self.bulk_paste_mode_var.get()
            col_defs = list(self.clipboard_col_defs)
            threading.Thread(
                target=self._run_automation,
                args=(selected, bulk_mode, col_defs),
                daemon=True,
            ).start()

    def _load_bulk_paste_mode_default(self) -> bool:
        try:
            config_path = self._resolve_config_path()
            if config_path.exists():
                with open(config_path, encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict) and "bulk_paste_mode" in data:
                    return bool(data["bulk_paste_mode"])
        except Exception:
            pass
        return True

    def _persist_bulk_paste_mode(self):
        try:
            config_path = self._resolve_config_path()
            self._ensure_config_file_exists(config_path)
            with open(config_path, encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                data = {}
            data["bulk_paste_mode"] = bool(self.bulk_paste_mode_var.get())
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
        except Exception as err:
            print(f"Could not save bulk_paste_mode to config: {err}")

    def _open_receipt_dialog(self, row_data):
        dlg = SalesAccReceiptGenDialog(self, row_data,
                                        callback=lambda msg: self._show_toast(msg))
        self.wait_window(dlg)

    def _show_toast(self, message):
        messagebox.showinfo("Result", message)

    def _resolve_config_path(self) -> Path:
        env_path = os.environ.get("SOBHA_CONFIG_PATH")
        if env_path:
            return Path(env_path).expanduser()

        if automation_module is not None:
            config_path = getattr(automation_module, "CONFIG_PATH", None)
            if config_path:
                return Path(str(config_path)).expanduser()

        return Path.home() / ".config" / "sobha-reconciliation" / "config.json"

    def _ensure_config_file_exists(self, config_path: Path):
        config_path.parent.mkdir(parents=True, exist_ok=True)
        if config_path.exists():
            return

        packaged_example = Path("/usr/share/sobha-reconciliation/config.example.json")
        local_example = Path(__file__).resolve().parent / "config.example.json"
        if packaged_example.exists():
            shutil.copy2(packaged_example, config_path)
            return
        if local_example.exists():
            shutil.copy2(local_example, config_path)
            return

        fallback = {
            "d365_url": "https://<your-tenant>.sandbox.operations.dynamics.com/?cmp=COMPANY&mi=LedgerJournalTable_CustPaym",
            "auth_json_path": "~/.config/sobha-reconciliation/auth.json",
            "journal_name": "ARBR Customers Receipt",
            "browser_headless": False,
            "browser_slow_mo_ms": 0,
            "page_load_timeout_ms": 60000,
            "page_load_wait_seconds": 1,
            "post_click_timeout_ms": 300000,
            "manual_login_button_timeout_ms": 1800000,
        }
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(fallback, f, indent=4)

    def _open_config_file(self):
        try:
            config_path = self._resolve_config_path()
            self._ensure_config_file_exists(config_path)
            config_path_str = str(config_path)

            if sys.platform.startswith("win"):
                os.startfile(config_path_str)  # type: ignore[attr-defined]
                return
            if sys.platform == "darwin":
                subprocess.Popen(["open", config_path_str])
                return
            subprocess.Popen(["xdg-open", config_path_str])
        except Exception as err:
            messagebox.showerror(
                "Open Config Failed",
                f"Could not open config automatically.\n\n"
                f"Path: {self._resolve_config_path()}\n\n"
                f"Error: {err}",
            )

    def _validate_config_for_action(self, require_auth_state: bool) -> bool:
        if automation_module is None:
            messagebox.showerror("Error", f"automation module import failed:\n{AUTOMATION_IMPORT_ERROR}")
            return False
        try:
            issues = automation_module.get_config_issues(require_auth_state=require_auth_state)
        except Exception as err:
            messagebox.showerror("Error", f"Unable to validate config:\n{err}")
            return False
        if not issues:
            return True
        messagebox.showerror(
            "Configuration Required",
            "Please update ~/.config/sobha-reconciliation/config.json:\n\n- "
            + "\n- ".join(issues),
        )
        return False

    def _bootstrap_login_config_if_needed(self) -> bool:
        if automation_module is None:
            messagebox.showerror("Error", f"automation module import failed:\n{AUTOMATION_IMPORT_ERROR}")
            return False

        try:
            issues = automation_module.get_config_issues(require_auth_state=False)
        except Exception as err:
            messagebox.showerror("Error", f"Unable to validate config:\n{err}")
            return False

        if not issues:
            return True

        d365_issues = [item for item in issues if "`d365_url`" in item]
        non_d365_issues = [item for item in issues if "`d365_url`" not in item]
        if non_d365_issues:
            messagebox.showerror(
                "Configuration Required",
                "Please update ~/.config/sobha-reconciliation/config.json:\n\n- "
                + "\n- ".join(issues),
            )
            return False

        current_journal = "ARBR Customers Receipt"
        try:
            current_journal = str(automation_module.CONFIG.get("journal_name", current_journal))
        except Exception:
            pass

        d365_url = simpledialog.askstring(
            "First Login Setup",
            "Enter your D365 URL (https://...)\n"
            "This will be saved to your user config for future runs.",
            parent=self,
        )
        if not d365_url:
            return False
        if not d365_url.strip().startswith("https://"):
            messagebox.showerror("Invalid URL", "D365 URL must start with https://")
            return False

        journal_name = simpledialog.askstring(
            "Journal Name",
            "Enter journal name:",
            initialvalue=current_journal,
            parent=self,
        )
        if journal_name is None:
            return False

        ok, msg = automation_module.update_user_runtime_config(
            d365_url=d365_url.strip(),
            journal_name=journal_name.strip(),
        )
        if not ok:
            messagebox.showerror("Setup Failed", msg)
            return False

        if d365_issues:
            messagebox.showinfo("Setup Saved", "Config saved for this user. Continuing login.")
        return True

    def _check_browser_ready_on_launch(self):
        if self._browser_check_prompted:
            return
        self._browser_check_prompted = True

        def check_task():
            try:
                if automation_module is None:
                    return
                ok, detail = automation_module.is_playwright_chromium_available()
                if ok:
                    return
                self.after(
                    0,
                    lambda d=detail: self._offer_browser_download(
                        Exception(d),
                        "App Startup Check",
                    ),
                )
            except Exception as err:
                print(f"Browser precheck skipped due to error: {err}")

        threading.Thread(target=check_task, daemon=True).start()

    def _is_missing_playwright_browser_error(self, err: Exception) -> bool:
        msg = str(err)
        return "Executable doesn't exist" in msg and "playwright" in msg.lower()

    def _offer_browser_download(self, err: Exception, action_name: str):
        detail = str(err)
        ask = messagebox.askyesno(
            "Browser Download Required",
            "Playwright Chromium browser is missing on this machine.\n\n"
            f"Action failed: {action_name}\n\n"
            "Do you want to download Chromium now?",
        )
        if not ask:
            messagebox.showerror("Error", f"{action_name} failed:\n{detail}")
            return

        def install_task():
            try:
                if automation_module is None:
                    raise ImportError(f"automation module import failed: {AUTOMATION_IMPORT_ERROR}")
                ok, output = automation_module.install_playwright_chromium()
                tail = "\n".join((output or "").splitlines()[-8:])
                if ok:
                    self.after(
                        0,
                        lambda: messagebox.showinfo(
                            "Download Complete",
                            "Chromium browser downloaded successfully.\n"
                            f"Please retry: {action_name}.\n\n{tail}",
                        ),
                    )
                else:
                    self.after(
                        0,
                        lambda: messagebox.showerror(
                            "Download Failed",
                            "Could not download Chromium automatically.\n\n"
                            "Please ask admin to run on this machine:\n"
                            "playwright install chromium\n\n"
                            f"Details:\n{tail}",
                        ),
                    )
            except Exception as install_err:
                self.after(
                    0,
                    lambda: messagebox.showerror(
                        "Download Failed",
                        "Automatic browser download failed.\n\n"
                        f"Details:\n{install_err}",
                    ),
                )

        threading.Thread(target=install_task, daemon=True).start()

    # ---- Automation thread (kept for integration) ----
    def _load_line_entry_mode_default(self) -> str:
        try:
            config_path = self._resolve_config_path()
            if config_path.exists():
                with open(config_path, encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    mode = str(data.get("line_entry_mode", "paste")).strip().lower()
                    if mode in {"paste", "dmf"}:
                        return mode
        except Exception:
            pass
        return "paste"

    def _run_automation(self, data, bulk_paste_mode=True, clipboard_col_defs=None):
        try:
            if automation_module is None:
                raise ImportError(f"automation module import failed: {AUTOMATION_IMPORT_ERROR}")
            print("--- Automation Started ---")
            line_entry_mode = self._load_line_entry_mode_default()
            automation_module.test_final8(
                data,
                bulk_paste_mode=bulk_paste_mode,
                clipboard_col_defs=clipboard_col_defs,
                line_entry_mode=line_entry_mode,
            )
            print("--- Automation Finished ---")
            if bulk_paste_mode or line_entry_mode == "dmf":
                self.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Success",
                        "All transactions completed.\nBrowser left open — close it when done.",
                    ),
                )
            else:
                self.after(
                    0,
                    lambda: messagebox.showinfo("Success", "Automation completed successfully."),
                )
        except ImportError:
            self.after(0, lambda: messagebox.showerror(
                "Error", "automation module not found."))
        except Exception as e:
            if e.__class__.__name__ == "AutomationStoppedByUser":
                print(f"Automation stopped by user: {e}")
                self.after(0, lambda err=e: messagebox.showinfo(
                    "Automation Stopped", f"{err}"))
                return
            session_expired_type = getattr(automation_module, "SessionExpiredError", None)
            if session_expired_type and isinstance(e, session_expired_type):
                print(f"Automation session expired: {e}")
                self.after(0, lambda err=e: self._handle_session_expired(err))
                return
            print(f"Automation error: {e}")
            if self._is_missing_playwright_browser_error(e):
                self.after(0, lambda err=e: self._offer_browser_download(err, "Automation"))
            else:
                self.after(0, lambda err=e: messagebox.showerror(
                    "Error", f"Automation failed:\n{err}"))

    def _run_login_automation(self):
        if not self._bootstrap_login_config_if_needed():
            return

        def run_task():
            try:
                if automation_module is None:
                    raise ImportError(f"automation module import failed: {AUTOMATION_IMPORT_ERROR}")
                # Use after to show info on main thread
                # self.after(0, lambda: messagebox.showinfo("Info", "Starting Login Automation..."))
                print("--- Login Automation Started ---")
                auth_result = automation_module.test_loginfunctionality()
                print("--- Login Automation Finished ---")
                self.after(0, lambda result=auth_result: self._apply_auth_result(result, invalidate_pending=True))
                self.after(0, lambda: messagebox.showinfo("Success", "Login automation completed."))
            except Exception as e:
                print(f"Login error: {e}")
                self.after(0, self._refresh_login_button_async)
                if self._is_missing_playwright_browser_error(e):
                    self.after(0, lambda err=e: self._offer_browser_download(err, "Login"))
                else:
                    self.after(0, lambda err=e: messagebox.showerror("Error", f"Login failed: {err}"))
        
        threading.Thread(target=run_task, daemon=True).start()


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    app = Application()
    app.mainloop()
