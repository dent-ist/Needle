# -*- coding: utf-8 -*-
"""
Needle - find the rows you need in huge CSV files.

Imports large CSV files into SQLite or DuckDB, searches them quickly,
and saves only the rows you need.
"""
import os
import queue
import shutil
import sys
import threading
import time
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

import csv_search_core as core

APP_NAME = "Needle"
APP_TAGLINE = "Find the rows you need in huge CSV files"
APP_ID = "Needle.CsvSearch"
PAGE_SIZE = 500
CELL_MAX = 300

SCOPE_QUERY = "[Query]"
SCOPE_ALL = "[All columns]"
SCOPE_SQL = "[SQL WHERE]"
MATCH_LABELS = {"Contains": "contains", "Equals": "equals", "Starts with": "startswith", "Regex": "regex"}
MATCH_NAMES = {v: k for k, v in MATCH_LABELS.items()}
EXPORT_ENCODINGS = {"UTF-8 (Excel-friendly)": "utf-8-sig", "UTF-8": "utf-8", "CP949 (ANSI Korean)": "cp949"}
CSV_TYPES = [("CSV files", "*.csv *.tsv *.txt"), ("All files", "*.*")]
DB_TYPES = [("Search databases", "*.sqlite *.db *.duckdb"), ("All files", "*.*")]
BLUE, RED, GREY = "#1f5f8b", "#b00020", "#555555"
INK = "#12324A"

# Needle icon (PNG, base64) so the exe needs no separate image files
ICON_64 = (
    "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"
)

ICON_32 = (
    "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"
)

ICON_22 = (
    "iVBORw0KGgoAAAANSUhEUgAAABYAAAAWCAYAAADEtGw7AAAD2ElEQVR42pWVXYiUZRTHf+d53pnZmd3ZdfbDFWvVitTVsJLScCET"
    "CTYNg4zSiy5aClTI0iwIybX0Ktjo07qIJcxSzCLDRaMMhKz1g9IMszYp8qL82Nb9mp15P57TxTrrLK42nqsX3off+Z//c/g/wnAJ"
    "oAA185akNUuD2sACqFNhjBIjCiA2FonHme7vvugvZknho37W/eVBLN4K8iiqE1C1l/5fo0QRiRDOouySMrvxUgMRaDV1M06mwmT2"
    "K+vF79EwQNVRSokM91XAeHFc4B/Fixb+e3jugIUDWnbj5M02VvaY8/N5UFs00lWAYI0QhCGAGhFcGPg2npjkApfK/fP+XqmZtyTt"
    "cuEpETOBYaXm2iohFxr8fJ6G+koGhwIGslnKEgmngKpecFFsmucG85PE2Dpw5n9HB4LI0FiXJVN7AytXvIDmuln9yhsMZoeMtRbQ"
    "2jjBFKNGLKgpxVNrhcFcwPSGNPvWGL7/cgtTbprK5InjyQcBw5aLqGdMScDCRQWhoyopNNz5OC/uqeXVllv4cMc2Oo//RjqVxDkd"
    "OX9d4KFcnpZlD7P+qcXct3gVr+0+x/zUp5Qly4nc6E0q0QJLb/8ASx9YyMZnnuDk73+xfvMmnnurk7qacSyZFdCbBWuuA+xZy8W+"
    "fu6dO5stm9ax/9vDPNiyjq4/z2DLx/HJEWXRbY4wklH7aUpROvPWm9na1sr23XtZvvolcr5PeSqJIeSPbkN1uRLzFFcK2BrDYDbL"
    "xPo6PmjbwJatO1nxfBviJUnEPaLIETmhtkIZ8uUKxd5YUGOEXN4nXVHB6xvW8Gb7x7Rv7WDN0kou9EfsOGKorQAXQfNMx8HTV+oz"
    "Y69VRDwe49mWZby7bRft2zuYMb2aJ5t8fvlbSJdB9yAsaIxonOD46JClKqUUbduVilUVI8LCprvZuedrTvzaRaauBucC8gGcPi/0"
    "9AmZSuWd5QEb98ToyQqZciUqMtkTax2RjvSKnKMmU0Xnjz9zvruHTFUluJCuc4YDXYbdq3ze+8ZjbXPA58csO49axqeV8DJUxTgn"
    "1XOaKzUwp8SY+kIIOVVQxbMWd7kn+RBWzo9omhqx7ydL+0FLVZKCBYqIotodBTJNAKpnL2ozscRa5+fzCPFLIYaOEUI9WXAqWKNk"
    "UkVQxTfxRCIKht7u+WHf0wKtpnrOoQoN7X4bi9/lQh9Ur77bRdc94qlIIeiPRZG3oPf4HX0jT1P1nOZKjczLYB4R5+oR7LXCvuAn"
    "4ETMWVQ/C6OBDb3HD1wsTDzqMa26/aFxYtwkz6hXSo6ETkIN3ZneEx09xaz/AKicr3SiE/aNAAAAAElFTkSuQmCC"
)


QUERY_HELP = """\
Query syntax

  word   or   "some text"        any column contains the text
  [Column]:"text"                that column contains the text
  [Column]."text"                (same as above)
  [Column]="text"                that column is exactly the text
  [Column]~"regex"               that column matches a regular expression

  "abc*"                         starts with abc
  "*abc"                         ends with abc
  "a*c"                          a ... c  (wildcard pattern)

  A AND B      or   A B   or   A && B      both
  A OR B       or   A || B                 either
  NOT A                                    exclude
  ( ... )                                  group

  AND is evaluated before OR:
      a OR b AND c   means   a OR (b AND c)

  Quotes are only needed for text with spaces, parentheses
  or the words AND / OR / NOT. Inside quotes, write \\" for a quote.
  The "Match case" checkbox applies to the whole query.

Examples

  ([상품명]."무선" AND "배송완료") OR "user1*"
  [배송상태]="반품" AND NOT [결제수단]:"카드"
  [주문일시]:"2025-12*" [지역]:서울
  keyboard OR speaker
"""


def cell(v):
    if v is None:
        return ""
    s = str(v).replace("\r", " ").replace("\n", " ⏎ ")
    return s if len(s) <= CELL_MAX else s[:CELL_MAX] + "…"


def spec_scope_label(entry):
    mode = entry.get("mode")
    if mode == "query":
        return "Query"
    if mode == "sql":
        return "SQL WHERE"
    if mode == "all":
        return f"All columns · {MATCH_NAMES.get(entry.get('match'), '')}"
    return f"[{entry.get('column')}] · {MATCH_NAMES.get(entry.get('match'), '')}"


# ---------------------------------------------------------------------------
# Import settings dialog
# ---------------------------------------------------------------------------
class ImportDialog(tk.Toplevel):
    def __init__(self, master, info, on_ok):
        super().__init__(master)
        self.info, self.on_ok = info, on_ok
        self.title("Import CSV")
        self.transient(master)
        self.geometry("840x660")
        self.minsize(720, 580)
        pad = {"padx": 10, "pady": 4}

        top = ttk.Frame(self)
        top.pack(fill="x", **pad)
        ttk.Label(top, text=os.path.basename(info.path), font=master.font_title).pack(anchor="w")
        ttk.Label(top, text=f"{core.human_size(info.size)}   ·   {info.path}",
                  foreground=GREY).pack(anchor="w")

        fmt = ttk.LabelFrame(self, text="File format (detected automatically; change it if the text looks garbled)")
        fmt.pack(fill="x", **pad)
        ttk.Label(fmt, text="Encoding").grid(row=0, column=0, padx=6, pady=6, sticky="w")
        self.enc_var = tk.StringVar(value=info.encoding)
        enc = ttk.Combobox(fmt, textvariable=self.enc_var, values=core.ENCODINGS, width=12)
        enc.grid(row=0, column=1, padx=6)
        enc.bind("<<ComboboxSelected>>", lambda e: self.refresh_preview())
        enc.bind("<Return>", lambda e: self.refresh_preview())
        ttk.Label(fmt, text="Delimiter").grid(row=0, column=2, padx=(20, 6), sticky="w")
        self.delim_labels = {v: k for k, v in core.DELIMITERS.items()}
        self.delim_var = tk.StringVar(value=core.DELIMITERS.get(info.delimiter, "Comma (,)"))
        dl = ttk.Combobox(fmt, textvariable=self.delim_var, values=list(self.delim_labels),
                          state="readonly", width=14)
        dl.grid(row=0, column=3, padx=6)
        dl.bind("<<ComboboxSelected>>", lambda e: self.refresh_preview())

        pv = ttk.Frame(fmt)
        pv.grid(row=1, column=0, columnspan=6, sticky="nsew", padx=6, pady=(0, 6))
        fmt.columnconfigure(5, weight=1)
        self.preview = ttk.Treeview(pv, show="headings", height=6)
        xs = ttk.Scrollbar(pv, orient="horizontal", command=self.preview.xview)
        self.preview.configure(xscrollcommand=xs.set)
        self.preview.pack(fill="x")
        xs.pack(fill="x")

        use = ttk.LabelFrame(self, text="How will you mostly search?")
        use.pack(fill="x", **pad)
        self.purpose = tk.StringVar(value="keyword")
        ttk.Radiobutton(use, text="Keywords  (find words wherever they appear, in any column)",
                        variable=self.purpose, value="keyword",
                        command=self.apply_recommendation).pack(anchor="w", padx=6, pady=2)
        ttk.Radiobutton(use, text="Conditions and exports  (filter by column values, save large results)",
                        variable=self.purpose, value="filter",
                        command=self.apply_recommendation).pack(anchor="w", padx=6, pady=2)

        eng = ttk.LabelFrame(self, text="Search engine")
        eng.pack(fill="x", **pad)
        self.engine = tk.StringVar()
        self.fts = tk.BooleanVar()
        self.rb_sqlite = ttk.Radiobutton(eng, variable=self.engine, value="sqlite",
                                         command=self.update_engine)
        self.rb_sqlite.grid(row=0, column=0, sticky="w", padx=6, pady=2)
        self.cb_fts = ttk.Checkbutton(eng, text="Build a full-text (FTS) index  (much faster keyword search, larger file)",
                                      variable=self.fts, command=self.update_engine)
        self.cb_fts.grid(row=1, column=0, sticky="w", padx=28)
        self.rb_duck = ttk.Radiobutton(eng, variable=self.engine, value="duckdb",
                                       command=self.update_engine)
        self.rb_duck.grid(row=2, column=0, sticky="w", padx=6, pady=2)
        if core.duckdb is None:
            self.rb_duck.state(["disabled"])
        self.reason = ttk.Label(eng, wraplength=780, foreground=BLUE, justify="left")
        self.reason.grid(row=3, column=0, sticky="w", padx=6, pady=(6, 2))
        self.estimate = ttk.Label(eng, foreground=GREY)
        self.estimate.grid(row=4, column=0, sticky="w", padx=6, pady=(0, 6))

        out = ttk.Frame(self)
        out.pack(fill="x", **pad)
        ttk.Label(out, text="Save database as").pack(side="left")
        self.db_var = tk.StringVar()
        ttk.Entry(out, textvariable=self.db_var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(out, text="Change…", command=self.browse).pack(side="left")

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=10, pady=10)
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Start import", style="Accent.TButton",
                   command=self.ok).pack(side="right", padx=6)

        self.refresh_preview()
        self.apply_recommendation()
        self.grab_set()
        self.focus_set()

    def refresh_preview(self):
        self.info.encoding = self.enc_var.get().strip() or "utf-8"
        self.info.delimiter = self.delim_labels.get(self.delim_var.get(), ",")
        try:
            self.info.refresh()
        except LookupError:
            messagebox.showerror(APP_NAME, f"Unknown encoding: {self.info.encoding}", parent=self)
            return
        cols = [f"k{i}" for i in range(len(self.info.headers))]
        self.preview.delete(*self.preview.get_children())
        self.preview["columns"] = cols
        for c, h in zip(cols, self.info.headers):
            self.preview.heading(c, text=h)
            self.preview.column(c, width=110, minwidth=60, stretch=False)
        for row in self.info.preview[1:]:
            self.preview.insert("", "end", values=[cell(v) for v in row])

    def apply_recommendation(self):
        eng, fts, why = core.recommend(self.info.size, self.purpose.get())
        self.rec = eng
        self.engine.set(eng)
        self.fts.set(fts)
        self.reason.configure(text="Why: " + why)
        self.update_engine()

    def update_engine(self):
        rec = lambda k: "      ← recommended" if k == self.rec else ""
        self.rb_sqlite.configure(
            text="SQLite   · instant lookups with column indexes; keyword search needs FTS" + rec("sqlite"))
        self.rb_duck.configure(
            text="DuckDB   · fast import, small file, parallel scans make every search quick" + rec("duckdb"))
        sqlite = self.engine.get() == "sqlite"
        self.cb_fts.state(["!disabled"] if sqlite else ["disabled"])
        use_fts = sqlite and self.fts.get()
        est = core.estimate_db_size(self.info.size, self.engine.get(), use_fts)
        base = os.path.splitext(self.info.path)[0]
        cur = self.db_var.get()
        ext = ".sqlite" if sqlite else ".duckdb"
        if not cur or os.path.splitext(cur)[1].lower() in (".sqlite", ".duckdb", ".db"):
            self.db_var.set((os.path.splitext(cur)[0] if cur else base) + ext)
        try:
            free = shutil.disk_usage(os.path.dirname(os.path.abspath(self.db_var.get()))).free
            free_txt = f"   ·   free disk space {core.human_size(free)}"
        except OSError:
            free_txt = ""
        self.estimate.configure(text=f"Estimated database size about {core.human_size(est)}{free_txt}")

    def browse(self):
        sqlite = self.engine.get() == "sqlite"
        p = filedialog.asksaveasfilename(
            parent=self, title="Save database as",
            initialfile=os.path.basename(self.db_var.get()),
            initialdir=os.path.dirname(self.db_var.get()),
            defaultextension=".sqlite" if sqlite else ".duckdb", filetypes=DB_TYPES)
        if p:
            self.db_var.set(p)
            self.update_engine()

    def ok(self):
        db = self.db_var.get().strip()
        if not db:
            return
        if os.path.abspath(db) == os.path.abspath(self.info.path):
            messagebox.showerror(APP_NAME, "The database cannot have the same name as the CSV file.", parent=self)
            return
        if os.path.exists(db) and not messagebox.askyesno(
                APP_NAME, f"{os.path.basename(db)} already exists.\nReplace it?", parent=self):
            return
        use_fts = self.engine.get() == "sqlite" and self.fts.get()
        est = core.estimate_db_size(self.info.size, self.engine.get(), use_fts)
        try:
            free = shutil.disk_usage(os.path.dirname(os.path.abspath(db))).free
            if free < est and not messagebox.askyesno(
                    APP_NAME, f"There may not be enough disk space.\nNeeded: about {core.human_size(est)}\n"
                              f"Free: {core.human_size(free)}\n\nContinue anyway?", parent=self):
                return
        except OSError:
            pass
        kind = self.engine.get()
        self.destroy()
        self.on_ok(self.info, kind, use_fts, db)


# ---------------------------------------------------------------------------
# History window
# ---------------------------------------------------------------------------
class HistoryWindow(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Search history")
        self.geometry("980x480")
        self.minsize(640, 300)

        top = ttk.Frame(self, padding=(10, 8))
        top.pack(fill="x")
        ttk.Label(top, text="Filter").pack(side="left")
        self.filter_var = tk.StringVar()
        fe = ttk.Entry(top, textvariable=self.filter_var)
        fe.pack(side="left", fill="x", expand=True, padx=6)
        self.only_db = tk.BooleanVar(value=app.engine is not None)
        cb = ttk.Checkbutton(top, text="Only the open database", variable=self.only_db, command=self.refresh)
        cb.pack(side="left", padx=6)
        if app.engine is None:
            cb.state(["disabled"])
        self.filter_var.trace_add("write", lambda *a: self.refresh())

        tf = ttk.Frame(self, padding=(10, 0))
        tf.pack(fill="both", expand=True)
        cols = ("time", "search", "scope", "results", "db")
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="extended")
        for c, title, w, st in (("time", "When", 160, False), ("search", "Search", 360, True),
                                ("scope", "Scope", 150, False), ("results", "Results", 80, False),
                                ("db", "Database", 200, False)):
            self.tree.heading(c, text=title)
            self.tree.column(c, width=w, stretch=st, anchor="e" if c == "results" else "w")
        ys = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ys.set)
        self.tree.pack(side="left", fill="both", expand=True)
        ys.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", lambda e: self.run())
        self.tree.bind("<Return>", lambda e: self.run())
        self.tree.bind("<Delete>", lambda e: self.delete())

        bf = ttk.Frame(self, padding=10)
        bf.pack(fill="x")
        ttk.Button(bf, text="Run search", style="Accent.TButton", command=self.run).pack(side="left")
        ttk.Button(bf, text="Put in search box", command=lambda: self.run(execute=False)).pack(side="left", padx=6)
        ttk.Button(bf, text="Close", command=self.destroy).pack(side="right")
        ttk.Button(bf, text="Clear all", command=self.clear).pack(side="right", padx=6)
        ttk.Button(bf, text="Delete", command=self.delete).pack(side="right")
        self.count_lbl = ttk.Label(bf, foreground=GREY)
        self.count_lbl.pack(side="left", padx=12)

        self.refresh()
        fe.focus_set()

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        self.rows = {}
        f = self.filter_var.get().lower()
        cur = os.path.normcase(os.path.abspath(self.app.engine.db_path)) if self.app.engine else None
        n = 0
        for i, e in enumerate(reversed(self.app.history.items)):
            db = e.get("db", "")
            if self.only_db.get() and cur and os.path.normcase(db) != cur:
                continue
            text = e.get("text", "")
            if f and f not in text.lower() and f not in os.path.basename(db).lower():
                continue
            iid = str(i)
            self.rows[iid] = e
            self.tree.insert("", "end", iid=iid, values=(
                e.get("time", ""), text.replace("\n", " "), spec_scope_label(e),
                e.get("results", ""), os.path.basename(db)))
            n += 1
        self.count_lbl.configure(text=f"{n:,} of {len(self.app.history.items):,} searches")

    def selected(self):
        return [self.rows[i] for i in self.tree.selection() if i in self.rows]

    def run(self, execute=True):
        sel = self.selected()
        if not sel:
            return
        if self.app.run_history_entry(sel[0], execute):
            self.destroy()

    def delete(self):
        sel = self.selected()
        if sel:
            self.app.history.remove(sel)
            self.refresh()

    def clear(self):
        if messagebox.askyesno(APP_NAME, "Delete the entire search history?", parent=self):
            self.app.history.clear()
            self.refresh()


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1240x780")
        self.minsize(940, 580)
        self._setup_fonts()
        self._setup_icon()

        self.engine = None
        self.job = None
        self.cancel_event = threading.Event()
        self.job_started = 0.0
        self.q = queue.Queue()
        self.active_spec = None
        self.offset = 0
        self.page_rows = []
        self.has_next = False
        self.history = core.History()
        self.hist_pos = None

        self._build()
        self._set_enabled(False)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(100, self._poll)

    # --- setup ---
    def _setup_fonts(self):
        win = sys.platform.startswith("win")
        ui_family = "Segoe UI" if win else None
        data_family = "Malgun Gothic" if win else None  # renders Korean data well
        for name, fam in (("TkDefaultFont", ui_family), ("TkMenuFont", ui_family),
                          ("TkTextFont", data_family), ("TkHeadingFont", data_family)):
            try:
                f = tkfont.nametofont(name)
                if fam:
                    f.configure(family=fam)
                f.configure(size=10)
            except tk.TclError:
                pass
        ui = tkfont.nametofont("TkDefaultFont").actual("family")
        data = tkfont.nametofont("TkTextFont").actual("family")
        self.font_title = (ui, 13, "bold")
        self.font_data = (data, 10)
        style = ttk.Style(self)
        style.configure("Treeview", rowheight=24, font=self.font_data)
        style.configure("Treeview.Heading", font=(data, 10, "bold"))
        style.configure("Accent.TButton", font=(ui, 10, "bold"))

    def _setup_icon(self):
        try:
            self.icons = {s: tk.PhotoImage(data="".join(globals()[f"ICON_{s}"])) for s in (64, 32, 22)}
            self.iconphoto(True, self.icons[64], self.icons[32])  # also used by every child window
        except tk.TclError:
            self.icons = {}

    def _build(self):
        bar = ttk.Frame(self, padding=(10, 8, 10, 4))
        bar.pack(fill="x")
        brand = ttk.Frame(bar)
        brand.pack(side="left", padx=(0, 14))
        if self.icons:
            ttk.Label(brand, image=self.icons[22]).pack(side="left", padx=(0, 6))
        ttk.Label(brand, text=APP_NAME, font=(self.font_title[0], 13, "bold"),
                  foreground=INK).pack(side="left")
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=(0, 12))
        self.btn_import = ttk.Button(bar, text="Import CSV…", command=self.open_csv)
        self.btn_import.pack(side="left")
        self.btn_open = ttk.Button(bar, text="Open database…", command=self.open_db)
        self.btn_open.pack(side="left", padx=6)
        self.btn_history = ttk.Button(bar, text="History…", command=self.show_history)
        self.btn_history.pack(side="left")
        self.info_lbl = ttk.Label(bar, text=f"{APP_TAGLINE}. Import a CSV or open a database to start.",
                                  foreground=GREY, font=self.font_data)
        self.info_lbl.pack(side="left", padx=12)

        sf = ttk.Frame(self, padding=(10, 4))
        sf.pack(fill="x")
        ttk.Label(sf, text="Search in").pack(side="left")
        self.scope_var = tk.StringVar(value=SCOPE_QUERY)
        self.scope = ttk.Combobox(sf, textvariable=self.scope_var, state="readonly", width=22)
        self.scope.pack(side="left", padx=(4, 10))
        self.match_lbl = ttk.Label(sf, text="Match")
        self.match_lbl.pack(side="left")
        self.match_var = tk.StringVar(value="Contains")
        self.match = ttk.Combobox(sf, textvariable=self.match_var, values=list(MATCH_LABELS),
                                  state="readonly", width=11)
        self.match.pack(side="left", padx=(4, 10))
        self.text_var = tk.StringVar()
        self.entry = ttk.Entry(sf, textvariable=self.text_var, font=(self.font_data[0], 11))
        self.entry.pack(side="left", fill="x", expand=True)
        self.entry.bind("<Return>", lambda e: self.do_search())
        self.entry.bind("<Up>", lambda e: self.history_step(-1))
        self.entry.bind("<Down>", lambda e: self.history_step(1))
        self.case_var = tk.BooleanVar(value=False)
        self.case = ttk.Checkbutton(sf, text="Match case", variable=self.case_var)
        self.case.pack(side="left", padx=8)
        self.btn_search = ttk.Button(sf, text="Search", style="Accent.TButton", command=self.do_search)
        self.btn_search.pack(side="left")
        self.btn_help = ttk.Button(sf, text="?", width=3, command=self.show_help)
        self.btn_help.pack(side="left", padx=(6, 0))

        hf = ttk.Frame(self, padding=(10, 0, 10, 4))
        hf.pack(fill="x")
        self.btn_index = ttk.Button(hf, text="Index this column", command=self.create_index)
        self.btn_index.pack(side="right", anchor="n")
        self.strategy_lbl = ttk.Label(hf, text="", foreground=BLUE, font=self.font_data,
                                      justify="left", wraplength=980)
        self.strategy_lbl.pack(side="left", fill="x", expand=True)
        hf.bind("<Configure>", lambda e: self.strategy_lbl.configure(wraplength=max(300, e.width - 170)))

        for v in (self.scope_var, self.match_var, self.text_var, self.case_var):
            v.trace_add("write", lambda *a: self.update_strategy())

        tf = ttk.Frame(self, padding=(10, 0))
        tf.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tf, show="headings", selectmode="extended")
        ys = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(tf, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", lambda e: self.show_record())
        self.tree.bind("<Return>", lambda e: self.show_record())
        self.tree.bind("<Control-c>", lambda e: self.copy_selection())
        self.tree.bind("<Control-a>", lambda e: self.tree.selection_set(self.tree.get_children()))

        bf = ttk.Frame(self, padding=(10, 6))
        bf.pack(fill="x")
        self.btn_prev = ttk.Button(bf, text="◀ Prev", width=8, command=lambda: self.page(-1))
        self.btn_prev.pack(side="left")
        self.page_lbl = ttk.Label(bf, text="", width=22, anchor="center")
        self.page_lbl.pack(side="left", padx=4)
        self.btn_next = ttk.Button(bf, text="Next ▶", width=8, command=lambda: self.page(1))
        self.btn_next.pack(side="left")
        self.btn_count = ttk.Button(bf, text="Count results", command=self.do_count)
        self.btn_count.pack(side="left", padx=(16, 0))

        self.btn_export = ttk.Button(bf, text="Save all results as CSV…", command=self.do_export)
        self.btn_export.pack(side="right")
        self.exp_enc = tk.StringVar(value=list(EXPORT_ENCODINGS)[0])
        ttk.Combobox(bf, textvariable=self.exp_enc, values=list(EXPORT_ENCODINGS),
                     state="readonly", width=20).pack(side="right", padx=6)
        ttk.Label(bf, text="Encoding").pack(side="right")

        st = ttk.Frame(self, padding=(10, 0, 10, 8))
        st.pack(fill="x")
        self.progress = ttk.Progressbar(st, length=220, mode="determinate")
        self.progress.pack(side="left")
        self.btn_cancel = ttk.Button(st, text="Cancel", width=7, command=self.cancel_job)
        self.btn_cancel.pack(side="left", padx=6)
        self.btn_cancel.state(["disabled"])
        self.status = ttk.Label(st, text="Ready", font=self.font_data)
        self.status.pack(side="left", padx=6, fill="x", expand=True)
        self.timer = ttk.Label(st, text="", foreground=GREY)
        self.timer.pack(side="right")

    # --- background jobs ---
    def run_job(self, label, work, on_done):
        if self.job:
            return
        self.job = label
        self.cancel_event = threading.Event()
        self.job_started = time.time()
        self._set_busy(True, label)

        def progress(frac, text=None):
            self.q.put(("progress", frac, text))

        def target():
            try:
                res, err = work(progress, self.cancel_event), None
            except BaseException as e:  # noqa
                res, err = None, e
            self.q.put(("done", on_done, res, err, time.time() - self.job_started))

        threading.Thread(target=target, daemon=True).start()

    def _poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                if msg[0] == "progress":
                    _, frac, text = msg
                    if frac is None:
                        if str(self.progress["mode"]) != "indeterminate":
                            self.progress.configure(mode="indeterminate")
                            self.progress.start(12)
                    else:
                        if str(self.progress["mode"]) != "determinate":
                            self.progress.stop()
                            self.progress.configure(mode="determinate")
                        self.progress["value"] = frac * 100
                    if text:
                        self.status.configure(text=text)
                else:
                    _, cb, res, err, el = msg
                    self.job = None
                    self._set_busy(False)
                    if err is not None and (self.cancel_event.is_set() or core.is_cancel_error(err)):
                        self.status.configure(text="Cancelled.")
                        cb(None, core.Cancelled(), el)
                    elif err is not None:
                        self.status.configure(text="Something went wrong.")
                        cb(None, err, el)
                    else:
                        cb(res, None, el)
        except queue.Empty:
            pass
        if self.job:
            self.timer.configure(text=f"{time.time() - self.job_started:,.1f} s")
        self.after(100, self._poll)

    def cancel_job(self):
        if self.job:
            self.cancel_event.set()
            if self.engine:
                self.engine.interrupt()
            self.status.configure(text="Cancelling…")

    def _set_busy(self, busy, label=""):
        if busy:
            self.progress.configure(mode="indeterminate")
            self.progress.start(12)
            self.status.configure(text=label)
            self.btn_cancel.state(["!disabled"])
            for w in (self.btn_import, self.btn_open, self.btn_history):
                w.state(["disabled"])
            self._set_enabled(False)
            self.configure(cursor="watch")
        else:
            self.progress.stop()
            self.progress.configure(mode="determinate")
            self.progress["value"] = 0
            self.btn_cancel.state(["disabled"])
            for w in (self.btn_import, self.btn_open, self.btn_history):
                w.state(["!disabled"])
            self._set_enabled(self.engine is not None)
            self.configure(cursor="")

    def _set_enabled(self, on):
        flag = ["!disabled"] if on else ["disabled"]
        for w in (self.btn_search, self.btn_count, self.btn_export, self.entry, self.case):
            w.state(flag)
        for w in (self.scope, self.match):
            w.state(["readonly", "!disabled"] if on else ["disabled"])
        self._update_paging()
        self.update_strategy()

    def show_error(self, err, title="Error"):
        messagebox.showerror(APP_NAME, f"{title}\n\n{err}")

    # --- import / open ---
    def open_csv(self, path=None):
        if self.job:
            return
        path = path or filedialog.askopenfilename(title="Choose a CSV file", filetypes=CSV_TYPES)
        if not path:
            return
        try:
            info = core.CsvInfo(path)
        except Exception as e:
            self.show_error(e, "Could not read the CSV file.")
            return
        ImportDialog(self, info, self.start_import)

    def start_import(self, info, kind, use_fts, db_path):
        self.close_engine()
        eng = core.new_engine(kind, db_path)

        def work(progress, cancel):
            return eng.import_csv(info, use_fts, progress, cancel)

        def done(res, err, el):
            if err is not None:
                eng.close()
                core.remove_db_files(db_path)
                self.engine = None
                self._set_enabled(False)
                if not isinstance(err, core.Cancelled):
                    self.show_error(err, "Import failed.")
                return
            msg = (f"Imported {res['rows']:,} rows in {el:,.1f} s · "
                   f"database {core.human_size(os.path.getsize(db_path))}")
            if res.get("fixed_rows"):
                msg += f" · fixed {res['fixed_rows']:,} rows with the wrong number of fields"
            if kind == "sqlite" and use_fts and not res.get("fts"):
                msg += " · full-text index skipped (not supported here)"
            self.set_engine(eng, msg)

        self.engine = eng  # so Cancel can interrupt it
        self.run_job("Importing…", work, done)

    def open_db(self, path=None, spec=None):
        if self.job:
            return
        path = path or filedialog.askopenfilename(title="Open a database", filetypes=DB_TYPES)
        if not path:
            return
        self.close_engine()

        def done(eng, err, el):
            if err is not None:
                self.engine = None
                self._set_enabled(False)
                self.show_error(err, "Could not open the database.")
                return
            self.set_engine(eng, f"Opened {os.path.basename(path)}", spec)

        self.run_job("Opening database…", lambda p, c: core.open_database(path), done)

    def close_engine(self):
        if self.engine:
            self.engine.close()
        self.engine = None
        self.tree.delete(*self.tree.get_children())
        self.page_rows = []
        self.active_spec = None

    def set_engine(self, eng, msg=None, spec=None):
        self.engine = eng
        fts = ""
        if eng.kind == "sqlite":
            fts = " · full-text index" if eng.fts else " · no full-text index"
        self.info_lbl.configure(
            text=f"{os.path.basename(eng.db_path)}   ({eng.label} · {eng.row_count:,} rows · "
                 f"{eng.ncols} columns{fts})", foreground="#222")
        self.title(f"{os.path.basename(eng.db_path)} - {APP_NAME}")
        self.scope.configure(values=[SCOPE_QUERY, SCOPE_ALL] + eng.headers + [SCOPE_SQL])
        cols = ["_row"] + eng.cols
        self.tree.configure(columns=cols)
        self.tree.heading("_row", text="Row")
        self.tree.column("_row", width=80, minwidth=60, stretch=False, anchor="e")
        for c, h in zip(eng.cols, eng.headers):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=140, minwidth=60, stretch=False)
        self.hist_pos = None
        if spec:
            self.apply_entry(spec)
        else:
            self.scope_var.set(SCOPE_QUERY)
            self.text_var.set("")
        self._set_enabled(True)
        self.do_search(after_msg=None if spec else msg)
        self.entry.focus_set()

    # --- search ---
    def current_spec(self):
        scope = self.scope_var.get()
        if scope == SCOPE_SQL:
            mode, col = "sql", None
        elif scope == SCOPE_QUERY:
            mode, col = "query", None
        elif scope == SCOPE_ALL or not self.engine or scope not in self.engine.headers:
            mode, col = "all", None
        else:
            mode, col = "column", self.engine.headers.index(scope)
        return {"mode": mode, "col": col, "match": MATCH_LABELS.get(self.match_var.get(), "contains"),
                "text": self.text_var.get(), "case": self.case_var.get()}

    def update_strategy(self):
        if not hasattr(self, "strategy_lbl"):
            return
        spec = self.current_spec()
        simple = spec["mode"] in ("all", "column")
        if self.engine and not self.job:
            self.match.state(["readonly", "!disabled"] if simple else ["disabled"])
            self.case.state(["disabled"] if spec["mode"] == "sql" else ["!disabled"])
        if not self.engine or self.job:
            self.btn_index.state(["disabled"])
            if not self.engine:
                self.strategy_lbl.configure(text="")
            return
        try:
            if spec["mode"] == "sql":
                cast = "TRY_CAST" if self.engine.kind == "duckdb" else "CAST"
                how = (f"Example: {cast}(\"수량\" AS INTEGER) >= 5 AND \"지역\" = '서울'   "
                       "(column names in double quotes, values in single quotes)")
            elif spec["mode"] == "query" and not spec["text"].strip():
                how = ('Examples:  keyboard   ·   [Column]:"text" AND NOT "word"   ·   '
                       '("a" OR "b*") [Column]="exact"      Click ? for the full syntax.')
            else:
                how = self.engine.build_where(spec)[3]
            self.strategy_lbl.configure(text=how, foreground=BLUE)
        except Exception as e:
            self.strategy_lbl.configure(text=str(e), foreground=RED)
        can_idx = (self.engine.can_index() and spec["mode"] == "column"
                   and spec["col"] not in getattr(self.engine, "indexed", set()))
        self.btn_index.state(["!disabled"] if can_idx else ["disabled"])

    def do_search(self, offset=0, spec=None, after_msg=None):
        if not self.engine or self.job:
            return
        spec = spec or self.current_spec()
        try:
            how = self.engine.build_where(spec)[3]
        except Exception as e:
            self.show_error(e, "Please check the search.")
            return
        if spec["mode"] == "query":
            how = how.split("   ·   ")[-1]
        eng = self.engine
        is_new = offset == 0 and (spec["text"].strip() != "" or spec["mode"] == "sql")

        def done(rows, err, el):
            if err is not None:
                if not isinstance(err, core.Cancelled):
                    self.show_error(err, "The search failed.")
                return
            self.active_spec, self.offset = spec, offset
            self.has_next = len(rows) > PAGE_SIZE
            self.page_rows = rows[:PAGE_SIZE]
            self.tree.delete(*self.tree.get_children())
            for i, r in enumerate(self.page_rows):
                self.tree.insert("", "end", iid=str(i), values=[cell(v) for v in r])
            self._update_paging()
            n = len(self.page_rows)
            what = "No results" if n == 0 else f"Showing results {offset + 1:,}–{offset + n:,}"
            self.status.configure(text=after_msg or f"{what} · {el:,.2f} s · {how}")
            if is_new:
                self.save_history(spec, f"{PAGE_SIZE}+" if self.has_next else str(n), el)

        self.run_job("Searching…", lambda p, c: eng.search(spec, PAGE_SIZE + 1, offset), done)

    def page(self, step):
        if self.active_spec is None:
            return
        self.do_search(max(0, self.offset + step * PAGE_SIZE), self.active_spec)

    def _update_paging(self):
        if not hasattr(self, "btn_prev"):
            return
        idle = self.engine is not None and not self.job and self.active_spec is not None
        self.btn_prev.state(["!disabled"] if idle and self.offset > 0 else ["disabled"])
        self.btn_next.state(["!disabled"] if idle and self.has_next else ["disabled"])
        if self.active_spec is None or not self.page_rows:
            self.page_lbl.configure(text="")
        else:
            self.page_lbl.configure(text=f"Page {self.offset // PAGE_SIZE + 1:,} ({PAGE_SIZE} per page)")

    def do_count(self):
        if not self.engine or self.job:
            return
        spec = self.active_spec or self.current_spec()
        eng = self.engine
        if spec["mode"] != "sql" and spec["text"].strip() == "":
            self.status.configure(text=f"{eng.row_count:,} rows in total")
            return

        def done(n, err, el):
            if err is not None:
                if not isinstance(err, core.Cancelled):
                    self.show_error(err, "Could not count the results.")
                return
            self.status.configure(text=f"{n:,} matching rows · {el:,.2f} s")
            self.save_history(spec, str(n), el)

        self.run_job("Counting results…", lambda p, c: eng.count(spec), done)

    def create_index(self):
        spec = self.current_spec()
        if not self.engine or self.job or spec["mode"] != "column":
            return
        name = self.engine.headers[spec["col"]]
        if not messagebox.askyesno(
                APP_NAME, f"Create an index on the '{name}' column?\n\n"
                          "On a large file this can take a few minutes and makes the database bigger.\n"
                          "Afterwards, 'Equals' and 'Starts with' searches with Match case on are instant."):
            return
        eng, col = self.engine, spec["col"]

        def done(_, err, el):
            if err is not None:
                if not isinstance(err, core.Cancelled):
                    self.show_error(err, "Could not create the index.")
                return
            self.case_var.set(True)
            self.status.configure(text=f"Created the index on '{name}' in {el:,.1f} s")
            self.update_strategy()

        self.run_job(f"Indexing '{name}'…", lambda p, c: eng.create_index(col), done)

    def do_export(self):
        if not self.engine or self.job:
            return
        spec = self.active_spec or self.current_spec()
        if spec["mode"] != "sql" and spec["text"].strip() == "":
            if not messagebox.askyesno(APP_NAME, "There is no search, so every row will be saved. Continue?"):
                return
        src = self.engine.meta.get("source", "results.csv")
        base = os.path.splitext(os.path.basename(src))[0]
        out = filedialog.asksaveasfilename(title="Save results", initialfile=f"{base}_results.csv",
                                           defaultextension=".csv", filetypes=CSV_TYPES)
        if not out:
            return
        enc = EXPORT_ENCODINGS[self.exp_enc.get()]
        eng = self.engine

        def done(n, err, el):
            if err is not None:
                try:
                    os.remove(out)
                except OSError:
                    pass
                if not isinstance(err, core.Cancelled):
                    self.show_error(err, "Could not save the file.")
                return
            self.status.configure(text=f"Saved {n:,} rows in {el:,.1f} s · {out}")

        self.run_job("Saving…", lambda p, c: eng.export(spec, out, enc, p, c), done)

    # --- history ---
    def save_history(self, spec, results, seconds):
        if not self.engine:
            return
        col = self.engine.headers[spec["col"]] if spec["mode"] == "column" else None
        self.history.add({
            "db": os.path.abspath(self.engine.db_path), "mode": spec["mode"], "column": col,
            "match": spec["match"], "text": spec["text"], "case": bool(spec["case"]),
            "results": results, "seconds": round(seconds, 3)})
        self.hist_pos = None

    def apply_entry(self, e):
        mode = e.get("mode", "query")
        if mode == "query":
            self.scope_var.set(SCOPE_QUERY)
        elif mode == "sql":
            self.scope_var.set(SCOPE_SQL)
        elif mode == "column" and self.engine and e.get("column") in self.engine.headers:
            self.scope_var.set(e["column"])
        else:
            self.scope_var.set(SCOPE_ALL)
        self.match_var.set(MATCH_NAMES.get(e.get("match"), "Contains"))
        self.case_var.set(bool(e.get("case")))
        self.text_var.set(e.get("text", ""))
        self.entry.icursor("end")

    def run_history_entry(self, e, execute=True):
        """returns True when the history window can close"""
        if self.job:
            return False
        db = e.get("db", "")
        same = self.engine and os.path.normcase(os.path.abspath(self.engine.db_path)) == os.path.normcase(db)
        if same:
            self.apply_entry(e)
            if execute:
                self.do_search()
            return True
        if not os.path.exists(db):
            messagebox.showerror(APP_NAME, f"The database for this search no longer exists:\n{db}")
            return False
        if not messagebox.askyesno(APP_NAME, f"This search was made in another database.\n\n"
                                             f"Open {os.path.basename(db)} and run it?"):
            return False
        self.open_db(db, spec=e)
        return True

    def history_step(self, step):
        """Up / Down in the search box walk through this database's history"""
        if not self.engine or self.job:
            return "break"
        cur = os.path.normcase(os.path.abspath(self.engine.db_path))
        items = [e for e in self.history.items if os.path.normcase(e.get("db", "")) == cur]
        if not items:
            return "break"
        if self.hist_pos is None:
            if step > 0:
                return "break"
            self.hist_pos = len(items)
        self.hist_pos = max(0, min(len(items), self.hist_pos + step))
        if self.hist_pos == len(items):
            self.text_var.set("")
            self.hist_pos = None
        else:
            self.apply_entry(items[self.hist_pos])
        return "break"

    def show_history(self):
        if self.job:
            return
        HistoryWindow(self)

    def show_help(self):
        win = tk.Toplevel(self)
        win.title("Query syntax")
        win.geometry("640x620")
        txt = tk.Text(win, wrap="none", padx=14, pady=12, font=("Consolas" if sys.platform.startswith("win")
                                                               else "TkFixedFont", 10))
        txt.insert("1.0", QUERY_HELP)
        txt.configure(state="disabled")
        ttk.Button(win, text="Close", command=win.destroy).pack(side="bottom", anchor="e", padx=10, pady=8)
        txt.pack(fill="both", expand=True)

    # --- viewing results ---
    def selected_rows(self):
        return [self.page_rows[int(i)] for i in self.tree.selection() if int(i) < len(self.page_rows)]

    def show_record(self):
        rows = self.selected_rows()
        if not rows or not self.engine:
            return
        row = rows[0]
        win = tk.Toplevel(self)
        win.title(f"Row {row[0]}")
        win.geometry("640x520")
        txt = tk.Text(win, wrap="word", padx=10, pady=10, font=self.font_data, spacing1=2, spacing3=4)
        ys = ttk.Scrollbar(win, command=txt.yview)
        txt.configure(yscrollcommand=ys.set)
        txt.tag_configure("k", font=(self.font_data[0], 10, "bold"), foreground=BLUE)
        for h, v in zip(self.engine.headers, row[1:]):
            txt.insert("end", h + "\n", "k")
            txt.insert("end", ("" if v is None else str(v)) + "\n\n")
        txt.configure(state="disabled")
        bar = ttk.Frame(win, padding=6)
        bar.pack(side="bottom", fill="x")

        def copy():
            self.clipboard_clear()
            self.clipboard_append("\n".join(f"{h}\t{'' if v is None else v}"
                                            for h, v in zip(self.engine.headers, row[1:])))
            self.status.configure(text="Copied.")

        ttk.Button(bar, text="Copy", command=copy).pack(side="right")
        ttk.Button(bar, text="Close", command=win.destroy).pack(side="right", padx=6)
        ys.pack(side="right", fill="y")
        txt.pack(fill="both", expand=True)

    def copy_selection(self):
        rows = self.selected_rows()
        if not rows:
            return
        clean = lambda v: "" if v is None else str(v).replace("\t", " ").replace("\r", " ").replace("\n", " ")
        lines = ["\t".join(self.engine.headers)] + ["\t".join(clean(v) for v in r[1:]) for r in rows]
        self.clipboard_clear()
        self.clipboard_append("\n".join(lines))
        self.status.configure(text=f"Copied {len(rows):,} rows (ready to paste into Excel).")

    def on_close(self):
        if self.job:
            if not messagebox.askyesno(APP_NAME, "A task is still running. Stop it and quit?"):
                return
            self.cancel_event.set()
            if self.engine:
                self.engine.interrupt()
        try:
            if self.engine:
                self.engine.close()
        except Exception:
            pass
        self.destroy()


def main():
    if sys.platform.startswith("win"):
        import ctypes
        try:  # keeps text sharp on high-DPI screens
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
        try:  # shows the Needle icon on the taskbar instead of Python's
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
        except Exception:
            pass
    app = App()
    # dropping a file onto the exe opens it
    if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
        p = sys.argv[1]
        if os.path.splitext(p)[1].lower() in (".sqlite", ".db", ".duckdb"):
            app.after(200, lambda: app.open_db(p))
        else:
            app.after(200, lambda: app.open_csv(p))
    app.mainloop()


if __name__ == "__main__":
    main()
