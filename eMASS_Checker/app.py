"""Tkinter desktop GUI for the CMMC eMASS Results Checker (Windows)."""
from __future__ import annotations

import json
import os
import queue
import shutil
import sys
import tempfile
import threading
import traceback
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import openpyxl

import reports
import xlsx_patch
from engine import DOMAINS, EDITABLE, Engine, Results, Scope, as_text
from textcheck import (APP_DIR, LanguageToolClient, TextChecker, bundled_rules_path, download_languagetool, find_lt_jar, load_rules,
                       user_dictionary_path, user_dir, user_rules_path)

APP_TITLE = "CMMC eMASS Results Checker"
APP_VERSION = "1.0.0"
SEV_COLORS = {"Error": "#f8d7da", "Warning": "#fff3cd", "Info": "#e3eef9"}


def settings_path():
    return os.path.join(user_dir(), "settings.json")


def load_settings():
    d = {"lt_folder": os.path.join(user_dir(), "LanguageTool"), "lt_port": 8081, "java": "java", "use_lt": False, "spelling": True, "last_dir": ""}
    try:
        with open(settings_path(), encoding="utf-8") as fh:
            d.update(json.load(fh))
    except (OSError, ValueError):
        pass
    return d


def save_settings(s):
    with open(settings_path(), "w", encoding="utf-8") as fh:
        json.dump(s, fh, indent=2)


def open_path(p):
    if os.name == "nt":
        os.startfile(p)  # noqa: S606
    else:
        webbrowser.open("file://" + os.path.abspath(p))


class App(tk.Tk):
    def __init__(self, initial_file=None):
        super().__init__()
        self.title(f"{APP_TITLE} {APP_VERSION}")
        self.geometry("1400x880")
        self.minsize(1050, 650)
        self.settings = load_settings()
        self.q = queue.Queue()
        self.res: Results | None = None
        self.engine: Engine | None = None
        self.checker: TextChecker | None = None
        self.lt: LanguageToolClient | None = None
        self.visible = []
        self.current = None
        self.busy = False
        self._style()
        self._menu()
        self._build()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(100, self._poll)
        self._bg("Loading dictionary", self._load_checker, lambda _r: self.status("Ready. Open an eMASS results workbook to begin (Ctrl+O)."))
        if initial_file:
            self.after(600, lambda: self.open_file(initial_file))

    # ------------------------------------------------------------------ UI
    def _style(self):
        st = ttk.Style(self)
        try:
            st.theme_use("vista" if os.name == "nt" else "clam")
        except tk.TclError:
            pass
        st.configure("Treeview", rowheight=24)
        st.configure("Accent.TButton", font=("Segoe UI", 9, "bold"))

    def _menu(self):
        m = tk.Menu(self)
        f = tk.Menu(m, tearoff=0)
        f.add_command(label="Open Results File...", accelerator="Ctrl+O", command=self.open_dialog)
        f.add_command(label="Save Corrected Copy...", accelerator="Ctrl+S", command=self.save_copy)
        f.add_separator()
        f.add_command(label="Export Report (HTML)...", command=lambda: self.export("html"))
        f.add_command(label="Export Report (Excel)...", command=lambda: self.export("xlsx"))
        f.add_command(label="Export Report (CSV)...", command=lambda: self.export("csv"))
        f.add_command(label="Print Report", accelerator="Ctrl+P", command=self.print_report)
        f.add_separator()
        f.add_command(label="Exit", command=self.on_close)
        m.add_cascade(label="File", menu=f)
        t = tk.Menu(m, tearoff=0)
        t.add_command(label="Run Checks", accelerator="F5", command=self.run_checks)
        t.add_command(label="Apply All Safe Fixes", command=self.apply_safe)
        t.add_separator()
        t.add_command(label="Settings...", command=self.settings_dialog)
        t.add_command(label="Edit Rules (rules.json)", command=self.edit_rules)
        t.add_command(label="Reload Rules", command=self.reload_rules)
        t.add_command(label="Edit Custom Dictionary", command=lambda: open_path(user_dictionary_path()))
        m.add_cascade(label="Tools", menu=t)
        h = tk.Menu(m, tearoff=0)
        h.add_command(label="How to Use", command=self.show_help)
        h.add_command(label="About", command=lambda: messagebox.showinfo("About", f"{APP_TITLE} {APP_VERSION}\nOffline checker for CMMC Level 2 eMASS Assessment Results (template v3.9).\nNo assessment data leaves this computer."))
        m.add_cascade(label="Help", menu=h)
        self.config(menu=m)
        self.bind("<Control-o>", lambda e: self.open_dialog())
        self.bind("<Control-s>", lambda e: self.save_copy())
        self.bind("<Control-p>", lambda e: self.print_report())
        self.bind("<F5>", lambda e: self.run_checks())

    def _build(self):
        top = ttk.Frame(self, padding=(8, 6))
        top.pack(fill="x")
        for text, cmd, style in (("Open Results File", self.open_dialog, "Accent.TButton"), ("Run Checks (F5)", self.run_checks, "Accent.TButton"),
                                 ("Apply Safe Fixes", self.apply_safe, None), ("Save Corrected Copy", self.save_copy, "Accent.TButton"),
                                 ("Export Report", self.export_menu, None), ("Print Report", self.print_report, None), ("Settings", self.settings_dialog, None)):
            b = ttk.Button(top, text=text, command=cmd, style=style) if style else ttk.Button(top, text=text, command=cmd)
            b.pack(side="left", padx=(0, 6))
            if text == "Export Report":
                self.export_btn = b
        self.file_lbl = ttk.Label(top, text="No file loaded", foreground="#555")
        self.file_lbl.pack(side="left", padx=10)

        # scope
        sc = ttk.LabelFrame(self, text="Scope (process only the assigned domains / objective range)", padding=(8, 4))
        sc.pack(fill="x", padx=8)
        row1 = ttk.Frame(sc)
        row1.pack(fill="x")
        self.dom_vars = {}
        self.all_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row1, text="All", variable=self.all_var, command=self._toggle_all).pack(side="left", padx=(0, 8))
        for d in DOMAINS:
            v = tk.BooleanVar(value=True)
            self.dom_vars[d] = v
            ttk.Checkbutton(row1, text=d, variable=v).pack(side="left", padx=2)
        row2 = ttk.Frame(sc)
        row2.pack(fill="x", pady=(4, 0))
        with open(os.path.join(APP_DIR, "data", "template_reference.json"), encoding="utf-8") as fh:
            ref_objs = [o["obj"] for o in json.load(fh)["objectives"]]
        ttk.Label(row2, text="Objective from").pack(side="left")
        self.obj_from = ttk.Combobox(row2, values=[""] + ref_objs, width=20)
        self.obj_from.pack(side="left", padx=4)
        ttk.Label(row2, text="to").pack(side="left")
        self.obj_to = ttk.Combobox(row2, values=[""] + ref_objs, width=20)
        self.obj_to.pack(side="left", padx=4)
        self.inc_asm = tk.BooleanVar(value=True)
        self.inc_req = tk.BooleanVar(value=True)
        self.inc_ssp = tk.BooleanVar(value=True)
        self.use_spell = tk.BooleanVar(value=self.settings.get("spelling", True))
        self.use_lt = tk.BooleanVar(value=self.settings.get("use_lt", False))
        for text, var in (("Assessment sheet", self.inc_asm), ("Requirements sheet", self.inc_req), ("OSC SSP(s) sheet", self.inc_ssp),
                          ("Spelling", self.use_spell), ("Grammar: LanguageTool (local)", self.use_lt)):
            ttk.Checkbutton(row2, text=text, variable=var).pack(side="left", padx=(12, 0))

        # filters
        fl = ttk.Frame(self, padding=(8, 6))
        fl.pack(fill="x")
        ttk.Label(fl, text="Severity").pack(side="left")
        self.f_sev = ttk.Combobox(fl, values=["All", "Error", "Warning", "Info"], width=9, state="readonly")
        self.f_sev.set("All")
        self.f_sev.pack(side="left", padx=(4, 10))
        ttk.Label(fl, text="Category").pack(side="left")
        self.f_cat = ttk.Combobox(fl, values=["All"], width=20, state="readonly")
        self.f_cat.set("All")
        self.f_cat.pack(side="left", padx=(4, 10))
        ttk.Label(fl, text="Sheet").pack(side="left")
        self.f_sheet = ttk.Combobox(fl, values=["All", "Assessment", "Requirements", "Requirement Objectives", "OSC SSP(s)"], width=20, state="readonly")
        self.f_sheet.set("All")
        self.f_sheet.pack(side="left", padx=(4, 10))
        ttk.Label(fl, text="Status").pack(side="left")
        self.f_status = ttk.Combobox(fl, values=["Open", "All", "Fixed", "Ignored"], width=8, state="readonly")
        self.f_status.set("Open")
        self.f_status.pack(side="left", padx=(4, 10))
        self.f_fix = tk.BooleanVar(value=False)
        ttk.Checkbutton(fl, text="Has proposed fix", variable=self.f_fix, command=self.refresh).pack(side="left", padx=(0, 10))
        ttk.Label(fl, text="Search").pack(side="left")
        self.f_search = ttk.Entry(fl, width=28)
        self.f_search.pack(side="left", padx=4)
        self.count_lbl = ttk.Label(fl, text="")
        self.count_lbl.pack(side="right")
        for w in (self.f_sev, self.f_cat, self.f_sheet, self.f_status):
            w.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        self.f_search.bind("<KeyRelease>", lambda e: self.refresh())

        pw = ttk.PanedWindow(self, orient="vertical")
        pw.pack(fill="both", expand=True, padx=8, pady=(0, 4))
        tf = ttk.Frame(pw)
        cols = ("n", "sev", "status", "sheet", "cell", "field", "obj", "rule", "msg")
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")
        for c, h, w, st in (("n", "#", 45, False), ("sev", "Severity", 70, False), ("status", "Status", 65, False), ("sheet", "Sheet", 150, False),
                            ("cell", "Cell", 55, False), ("field", "Field", 130, False), ("obj", "Objective", 125, False), ("rule", "Rule", 170, False), ("msg", "Issue", 600, True)):
            self.tree.heading(c, text=h, command=lambda c=c: self._sort(c))
            self.tree.column(c, width=w, stretch=st, anchor="w")
        for s, col in SEV_COLORS.items():
            self.tree.tag_configure(s, background=col)
        self.tree.tag_configure("closed", foreground="#888")
        ys = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ys.set)
        self.tree.pack(side="left", fill="both", expand=True)
        ys.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Delete>", lambda e: self.ignore())
        pw.add(tf, weight=3)

        # detail panel
        df = ttk.Frame(pw, padding=(4, 4))
        self.d_head = ttk.Label(df, text="Select an issue to see details.", font=("Segoe UI", 10, "bold"))
        self.d_head.grid(row=0, column=0, columnspan=2, sticky="w")
        self.d_msg = ttk.Label(df, text="", wraplength=1300, justify="left")
        self.d_msg.grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))
        self.d_why = ttk.Label(df, text="", wraplength=1300, justify="left", foreground="#444")
        self.d_why.grid(row=2, column=0, columnspan=2, sticky="w", pady=(2, 4))
        ttk.Label(df, text="Current cell text (flagged text highlighted)").grid(row=3, column=0, sticky="w")
        self.d_cur_lbl = ttk.Label(df, text="Proposed text (edit before accepting if needed)")
        self.d_cur_lbl.grid(row=3, column=1, sticky="w")
        self.cur_txt = tk.Text(df, height=8, wrap="word", font=("Segoe UI", 10), background="#fafafa")
        self.cur_txt.grid(row=4, column=0, sticky="nsew", padx=(0, 6))
        self.cur_txt.tag_configure("hl", background="#ffe08a")
        self.prop_txt = tk.Text(df, height=8, wrap="word", font=("Segoe UI", 10), undo=True)
        self.prop_txt.grid(row=4, column=1, sticky="nsew")
        self.prop_txt.bind("<KeyRelease>", lambda e: self._count())
        self.prop_txt.bind("<Control-Return>", lambda e: (self.accept(), "break"))
        df.columnconfigure(0, weight=1)
        df.columnconfigure(1, weight=1)
        df.rowconfigure(4, weight=1)
        bf = ttk.Frame(df)
        bf.grid(row=5, column=0, columnspan=2, sticky="we", pady=(6, 0))
        self.b_accept = ttk.Button(bf, text="Accept Change (Ctrl+Enter)", style="Accent.TButton", command=self.accept)
        self.b_accept.pack(side="left")
        ttk.Button(bf, text="Reset to Proposed", command=self._reset_prop).pack(side="left", padx=6)
        ttk.Button(bf, text="Ignore Issue (Del)", command=self.ignore).pack(side="left", padx=6)
        self.b_word = ttk.Button(bf, text="Add Word to Dictionary", command=self.add_word)
        self.b_word.pack(side="left", padx=6)
        ttk.Button(bf, text="Revert Cell to Original", command=self.revert_cell).pack(side="left", padx=6)
        ttk.Button(bf, text="Copy Location", command=self.copy_loc).pack(side="left", padx=6)
        self.char_lbl = ttk.Label(bf, text="")
        self.char_lbl.pack(side="right")
        pw.add(df, weight=2)

        sb = ttk.Frame(self, padding=(8, 2))
        sb.pack(fill="x")
        self.status_lbl = ttk.Label(sb, text="")
        self.status_lbl.pack(side="left")
        self.pbar = ttk.Progressbar(sb, length=220, mode="determinate")
        self.pbar.pack(side="right")

    # ------------------------------------------------------------------ helpers
    def status(self, text):
        self.status_lbl.config(text=text)

    def _toggle_all(self):
        for v in self.dom_vars.values():
            v.set(self.all_var.get())

    def _bg(self, label, fn, done=None, *args):
        """Run fn in a worker thread; fn may call self._progress(i, n)."""
        if self.busy:
            messagebox.showinfo(APP_TITLE, "Please wait for the current task to finish.")
            return
        self.busy = True
        self.status(label + "...")
        self.pbar.config(value=0)

        def work():
            try:
                r = fn(*args)
                self.q.put(("done", done, r))
            except Exception as e:  # noqa: BLE001
                self.q.put(("error", label, f"{e}\n\n{traceback.format_exc()}"))
        threading.Thread(target=work, daemon=True).start()

    def _progress(self, i, n):
        self.q.put(("progress", i, n))

    def _poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                if msg[0] == "progress":
                    _, i, n = msg
                    self.pbar.config(maximum=max(n, 1), value=i)
                elif msg[0] == "done":
                    self.busy = False
                    self.pbar.config(value=0)
                    if msg[1]:
                        msg[1](msg[2])
                elif msg[0] == "error":
                    self.busy = False
                    self.pbar.config(value=0)
                    self.status(f"{msg[1]} failed.")
                    messagebox.showerror(APP_TITLE, f"{msg[1]} failed:\n{msg[2][:1500]}")
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _load_checker(self):
        self.checker = TextChecker(load_rules())
        return True

    # ------------------------------------------------------------------ file
    def open_dialog(self):
        p = filedialog.askopenfilename(title="Open eMASS Assessment Results workbook", initialdir=self.settings.get("last_dir") or None,
                                       filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")])
        if p:
            self.open_file(p)

    def open_file(self, path):
        if self.res and self.res.changes and not messagebox.askyesno(APP_TITLE, "You have unsaved changes. Discard them and open another file?"):
            return
        self.settings["last_dir"] = os.path.dirname(path)
        save_settings(self.settings)

        def load():
            return Results(path)

        def done(res):
            self.res = res
            self.engine = None
            self.file_lbl.config(text=f"{os.path.basename(path)}   (template v{res.template_version})")
            self.tree.delete(*self.tree.get_children())
            self._clear_detail()
            self.status("File loaded. Set the scope, then click Run Checks (F5).")
            self.run_checks()
        self._bg("Opening file", load, done)

    # ------------------------------------------------------------------ checks
    def _scope(self):
        doms = [d for d, v in self.dom_vars.items() if v.get()]
        return Scope(doms, self.obj_from.get().strip() or None, self.obj_to.get().strip() or None, self.inc_asm.get(), self.inc_req.get(), self.inc_ssp.get())

    def run_checks(self):
        if not self.res:
            return self.open_dialog()
        if self.busy:
            return
        if self.checker is None:
            return self.after(500, self.run_checks)
        scope = self._scope()
        if not scope.domains:
            return messagebox.showwarning(APP_TITLE, "Select at least one domain.")
        self.settings["spelling"] = self.use_spell.get()
        self.settings["use_lt"] = self.use_lt.get()
        save_settings(self.settings)
        use_lt = self.use_lt.get()
        use_spell = self.use_spell.get()
        prev_ignored = self.engine.ignored if self.engine else set()

        def work():
            if use_lt:
                self._ensure_lt()
            self.checker.lt = self.lt if use_lt else None
            eng = Engine(self.res, {"spelling": use_spell, "languagetool": use_lt}, self.checker)
            eng.ignored = prev_ignored
            eng.run(scope, self._progress)
            return eng

        def done(eng):
            self.engine = eng
            cats = sorted({i.category for i in eng.issues})
            self.f_cat.config(values=["All"] + cats)
            self.refresh()
            s = eng.summary()["by_severity"]
            self.status(f"Checks complete: {s.get('Error', 0)} errors, {s.get('Warning', 0)} warnings, {s.get('Info', 0)} info. Scope: {scope.describe()}")
        self._bg("Running checks", work, done)

    def _ensure_lt(self):
        if self.lt and self.lt._alive():
            return
        rules = load_rules()
        self.lt = LanguageToolClient(self.settings.get("lt_folder"), self.settings.get("lt_port", 8081), self.settings.get("java", "java"),
                                     rules.get("languagetool_disabled_rules"))
        self.q.put(("progress", 0, 1))
        self.lt.start()

    def reload_rules(self):
        def work():
            self.checker = TextChecker(load_rules())
            return True
        self._bg("Reloading rules", work, lambda _r: (self.status("Rules reloaded."), self.res and self.run_checks()))

    def apply_safe(self):
        if not self.engine:
            return
        n_safe = sum(1 for i in self.engine.issues if i.safe and i.status == "Open")
        if not n_safe:
            return messagebox.showinfo(APP_TITLE, "No open issues have safe automatic fixes.")
        if not messagebox.askyesno(APP_TITLE, f"Apply {n_safe} safe mechanical fixes?\n\nSafe fixes are limited to spacing, hidden characters, doubled punctuation, repeated words, "
                                   "terminal periods, sentence capitalization, contractions, product-name capitalization, and replacing \"the assigned certified assessor\". "
                                   "Protected fields and Findings results are never changed. Every change is logged and can be reverted per cell."):
            return

        def done(n):
            self.refresh()
            self.status(f"Applied {n} safe fixes. Review remaining issues, then Save Corrected Copy.")
        self._bg("Applying safe fixes", lambda: self.engine.apply_safe(progress=self._progress), done)

    # ------------------------------------------------------------------ list
    def _filtered(self):
        if not self.engine:
            return []
        sev, cat, sheet, st = self.f_sev.get(), self.f_cat.get(), self.f_sheet.get(), self.f_status.get()
        q = self.f_search.get().strip().lower()
        out = []
        for i in self.engine.issues:
            if sev != "All" and i.severity != sev:
                continue
            if cat != "All" and i.category != cat:
                continue
            if sheet != "All" and i.sheet != sheet:
                continue
            if st != "All" and i.status != st:
                continue
            if self.f_fix.get() and i.proposed is None:
                continue
            if q and q not in " ".join([i.location, i.message, i.rule, i.objective, i.field, i.current]).lower():
                continue
            out.append(i)
        return out

    def refresh(self, keep=None):
        sel = keep or (self.current.uid if self.current else None)
        self.tree.delete(*self.tree.get_children())
        self.visible = self._filtered()
        for n, i in enumerate(self.visible, start=1):
            tags = (i.severity,) if i.status == "Open" else ("closed",)
            self.tree.insert("", "end", iid=str(i.uid), values=(n, i.severity, i.status, i.sheet, i.cell, i.field, i.objective, i.rule, i.message), tags=tags)
        if self.engine:
            s = self.engine.summary()
            self.count_lbl.config(text=f"Showing {len(self.visible)} | Open {s['open']} | Fixed {s['fixed']} | Ignored {s['ignored']} | Cells changed {len(self.res.changes)}")
        if sel and self.tree.exists(str(sel)):
            self.tree.selection_set(str(sel))
            self.tree.see(str(sel))

    def _sort(self, col):
        key = {"n": lambda i: i.uid, "sev": lambda i: {"Error": 0, "Warning": 1, "Info": 2}[i.severity], "status": lambda i: i.status, "sheet": lambda i: (i.sheet, i.row, i.col),
               "cell": lambda i: (i.sheet, i.row, i.col), "field": lambda i: i.field, "obj": lambda i: i.objective, "rule": lambda i: i.rule, "msg": lambda i: i.message}[col]
        if self.engine:
            rev = getattr(self, "_sort_rev", False) and getattr(self, "_sort_col", None) == col
            self.engine.issues.sort(key=key, reverse=rev)
            self._sort_col, self._sort_rev = col, not rev
            self.refresh()

    def _issue(self, uid):
        return next((i for i in self.engine.issues if i.uid == int(uid)), None) if self.engine else None

    def on_select(self, _e=None):
        sel = self.tree.selection()
        if not sel:
            return
        i = self._issue(sel[0])
        if not i:
            return
        self.current = i
        self.d_head.config(text=f"{i.severity.upper()}  |  {i.location}  (row {i.row or '-'}, column {i.col_letter or '-'} {i.field})  |  {i.objective}  |  {i.rule}  |  {i.status}")
        self.d_msg.config(text="Issue: " + i.message)
        self.d_why.config(text="Why: " + i.why)
        live = as_text(self.res.get(i.sheet, i.row, i.col)) if i.col else i.current
        self.cur_txt.config(state="normal")
        self.cur_txt.delete("1.0", "end")
        self.cur_txt.insert("1.0", live)
        a, b = i.span
        if live == i.current and (a, b) != (0, 0):
            self.cur_txt.tag_add("hl", f"1.0+{a}c", f"1.0+{max(b, a + 1)}c")
            self.cur_txt.see(f"1.0+{a}c")
        self.cur_txt.config(state="disabled")
        editable = (i.sheet, i.col) in EDITABLE and i.status == "Open"
        self.prop_txt.config(state="normal")
        self.prop_txt.delete("1.0", "end")
        if editable:
            self.prop_txt.insert("1.0", i.proposed if i.proposed is not None else live)
            self.d_cur_lbl.config(text="Proposed text (edit before accepting if needed)" if i.proposed is not None else "No automatic fix. Edit the text manually, then Accept Change.")
            self.b_accept.state(["!disabled"])
        else:
            self.prop_txt.insert("1.0", "This field is protected or the issue is closed. Flagged for assessor review; it is not modified by this tool." if i.status == "Open" else f"Issue {i.status.lower()}.")
            self.prop_txt.config(state="disabled")
            self.d_cur_lbl.config(text="Protected field")
            self.b_accept.state(["disabled"])
        self.b_word.state(["!disabled"] if i.word else ["disabled"])
        self._count()

    def _count(self):
        if not self.current:
            return
        t = self.prop_txt.get("1.0", "end-1c")
        lim = 2000 if self.current.field in ("Executive Summary", "C3PAO Executive Summary", "ESP Name") else 4000
        self.char_lbl.config(text=f"{len(t):,} / {lim:,} characters", foreground="red" if len(t) > lim else "black")

    def _clear_detail(self):
        self.current = None
        for w in (self.cur_txt, self.prop_txt):
            w.config(state="normal")
            w.delete("1.0", "end")
        self.d_head.config(text="Select an issue to see details.")
        self.d_msg.config(text="")
        self.d_why.config(text="")

    def _reset_prop(self):
        if self.current:
            self.on_select()

    def _next_after(self, issue):
        """Select the next open issue after a change."""
        vis = self._filtered()
        nxt = next((x for x in vis if x.status == "Open" and (x.sheet, x.row, x.col) == (issue.sheet, issue.row, issue.col)), None)
        if not nxt:
            uids = [v.uid for v in self.visible]
            idx = uids.index(issue.uid) if issue.uid in uids else 0
            later = [x for x in self.visible[idx + 1:] if x.status == "Open"]
            nxt = later[0] if later else None
        self.refresh(keep=nxt.uid if nxt else None)
        if nxt:
            self.on_select()
        else:
            self._clear_detail()

    # ------------------------------------------------------------------ actions
    def accept(self):
        i = self.current
        if not i or i.status != "Open" or (i.sheet, i.col) not in EDITABLE:
            return
        text = self.prop_txt.get("1.0", "end-1c")
        live = as_text(self.res.get(i.sheet, i.row, i.col))
        if text == live:
            return messagebox.showinfo(APP_TITLE, "The text is unchanged. Edit it or use Ignore Issue.")
        try:
            self.engine.accept(i, text)
        except ValueError as e:
            return messagebox.showwarning(APP_TITLE, str(e))
        self.status(f"Updated {i.location}. {len(self.res.changes)} cell(s) changed; use Save Corrected Copy to write a new file.")
        self._next_after(i)

    def ignore(self):
        i = self.current
        if i and i.status == "Open":
            self.engine.ignore(i)
            self._next_after(i)

    def add_word(self):
        i = self.current
        if not (i and i.word):
            return
        self.checker.speller.add_user_word(i.word)
        for x in self.engine.issues:
            if x.rule == "TX-SPELL" and x.word and x.word.lower() == i.word.lower() and x.status == "Open":
                x.status = "Ignored"
        self.status(f'Added "{i.word}" to your custom dictionary.')
        self._next_after(i)

    def revert_cell(self):
        i = self.current
        if not i or (i.sheet, i.row, i.col) not in self.res.changes:
            return messagebox.showinfo(APP_TITLE, "This cell has no changes to revert.")
        self.res.revert(i.sheet, i.row, i.col)
        self.engine.recheck_cell(i.sheet, i.row, i.col)
        self.refresh()
        self.status(f"Reverted {i.location} to the original text.")

    def copy_loc(self):
        if self.current:
            self.clipboard_clear()
            self.clipboard_append(self.current.location)

    def save_copy(self):
        if not self.res:
            return
        if not self.res.changes:
            return messagebox.showinfo(APP_TITLE, "No changes have been accepted yet. Nothing to save.")
        default = xlsx_patch.next_version_path(self.res.path)
        dst = filedialog.asksaveasfilename(title="Save corrected copy", initialdir=os.path.dirname(default), initialfile=os.path.basename(default),
                                           defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")])
        if not dst:
            return
        if os.path.abspath(dst) == os.path.abspath(self.res.path):
            return messagebox.showwarning(APP_TITLE, "Choose a new file name. The original results file is never overwritten.")
        changes = dict(self.res.changes)

        def work():
            n = xlsx_patch.write_copy(self.res.path, dst, changes)
            wb = openpyxl.load_workbook(dst)  # verify every written cell reads back exactly
            bad = [f"{s}!{r},{c}" for (s, r, c), t in changes.items() if as_text(wb[s].cell(r, c).value) != t.replace("\r\n", "\n").replace("\r", "\n")]
            log = os.path.splitext(dst)[0] + "_change_log.html"
            reports.write_html(log, self.engine, include_closed=True)
            return n, bad, log

        def done(r):
            n, bad, log = r
            if bad:
                messagebox.showerror(APP_TITLE, "Saved, but these cells did not read back as expected:\n" + "\n".join(bad[:20]))
            else:
                if messagebox.askyesno(APP_TITLE, f"Saved {n} changed cell(s) to:\n{dst}\n\nThe original file was not modified. Template dropdowns, formulas, and formatting are preserved.\n"
                                       f"A change log report was saved to:\n{log}\n\nOpen the folder?"):
                    open_path(os.path.dirname(dst))
            self.status(f"Saved corrected copy: {os.path.basename(dst)}")
        self._bg("Saving corrected copy", work, done)

    def export_menu(self):
        m = tk.Menu(self, tearoff=0)
        m.add_command(label="HTML report (printable)", command=lambda: self.export("html"))
        m.add_command(label="Excel report (.xlsx)", command=lambda: self.export("xlsx"))
        m.add_command(label="CSV report", command=lambda: self.export("csv"))
        x, y = self.export_btn.winfo_rootx(), self.export_btn.winfo_rooty() + self.export_btn.winfo_height()
        m.tk_popup(x, y)

    def export(self, kind):
        if not self.engine:
            return messagebox.showinfo(APP_TITLE, "Run checks first.")
        include_closed = messagebox.askyesno(APP_TITLE, "Include fixed and ignored issues in the report?\n\nYes = full history, No = open issues only.")
        stem = os.path.splitext(os.path.basename(self.res.path))[0] + "_check_report"
        p = filedialog.asksaveasfilename(title="Save report", initialdir=os.path.dirname(self.res.path), initialfile=f"{stem}.{kind}",
                                         defaultextension=f".{kind}", filetypes=[(kind.upper(), f"*.{kind}")])
        if not p:
            return
        try:
            if kind == "html":
                reports.write_html(p, self.engine, include_closed)
            elif kind == "xlsx":
                reports.write_xlsx(p, self.engine, include_closed)
            else:
                reports.write_csv(p, self.engine.issues, include_closed)
        except PermissionError:
            return messagebox.showerror(APP_TITLE, "Could not write the file. Close it in Excel or your browser and try again.")
        self.status(f"Report saved: {p}")
        if messagebox.askyesno(APP_TITLE, f"Report saved:\n{p}\n\nOpen it now?"):
            open_path(p)

    def print_report(self):
        if not self.engine:
            return messagebox.showinfo(APP_TITLE, "Run checks first.")
        p = os.path.join(tempfile.gettempdir(), "eMASS_check_report_print.html")
        reports.write_html(p, self.engine, include_closed=False, auto_print=True)
        webbrowser.open("file:///" + p.replace("\\", "/"))
        self.status("Report opened in your browser with the print dialog. The temporary print file contains CUI excerpts; it is overwritten on each print.")

    # ------------------------------------------------------------------ settings
    def edit_rules(self):
        up = user_rules_path()
        if not os.path.exists(up):
            shutil.copy(bundled_rules_path(), up)
        open_path(up)
        messagebox.showinfo(APP_TITLE, f"Editing your rules file:\n{up}\n\nSave it, then choose Tools > Reload Rules. Delete the file to return to the defaults.")

    def settings_dialog(self):
        d = tk.Toplevel(self)
        d.title("Settings")
        d.transient(self)
        d.grab_set()
        frm = ttk.Frame(d, padding=12)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="LanguageTool (local grammar engine)", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(frm, text="Runs on this computer only (127.0.0.1). Requires Java 17 or newer. Download is about 250 MB, one time.", foreground="#555").grid(row=1, column=0, columnspan=3, sticky="w")
        v_folder = tk.StringVar(value=self.settings.get("lt_folder", ""))
        v_java = tk.StringVar(value=self.settings.get("java", "java"))
        v_port = tk.StringVar(value=str(self.settings.get("lt_port", 8081)))
        ttk.Label(frm, text="LanguageTool folder").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=v_folder, width=60).grid(row=2, column=1, sticky="we")
        ttk.Button(frm, text="Browse", command=lambda: v_folder.set(filedialog.askdirectory() or v_folder.get())).grid(row=2, column=2, padx=4)
        ttk.Label(frm, text="Java executable").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=v_java, width=60).grid(row=3, column=1, sticky="we")
        ttk.Button(frm, text="Browse", command=lambda: v_java.set(filedialog.askopenfilename(filetypes=[("java.exe", "java.exe"), ("All", "*.*")]) or v_java.get())).grid(row=3, column=2, padx=4)
        ttk.Label(frm, text="Local port").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=v_port, width=8).grid(row=4, column=1, sticky="w")
        st = ttk.Label(frm, text="Status: " + ("server jar found" if find_lt_jar(v_folder.get()) else "not installed"))
        st.grid(row=5, column=0, columnspan=3, sticky="w", pady=4)
        pb = ttk.Progressbar(frm, length=300)
        pb.grid(row=6, column=0, columnspan=2, sticky="w")

        def do_download():
            if not messagebox.askyesno(APP_TITLE, "Download LanguageTool from languagetool.org now (about 250 MB)?\n\nOnly the LanguageTool software is downloaded. No assessment text is sent.", parent=d):
                return
            dest = os.path.join(user_dir(), "LanguageTool")

            def prog(got, total):
                self.after(0, lambda: (pb.config(maximum=total or 1, value=got), st.config(text=f"Downloading... {got // (1 << 20)} MB")))

            def work():
                try:
                    folder = download_languagetool(dest, prog)
                    self.after(0, lambda: (v_folder.set(folder), st.config(text="Status: installed")))
                except Exception as e:  # noqa: BLE001
                    self.after(0, lambda e=e: (st.config(text="Status: download failed"),
                                               messagebox.showerror(APP_TITLE, f"Download failed: {e}\n\nYou can download LanguageTool-stable.zip manually from https://languagetool.org/download/, unzip it, and set the folder here.", parent=d)))
            threading.Thread(target=work, daemon=True).start()

        def do_test():
            try:
                c = LanguageToolClient(v_folder.get(), int(v_port.get()), v_java.get(), load_rules().get("languagetool_disabled_rules"))
                st.config(text="Starting LanguageTool...")
                d.update()
                c.start()
                hits = c.check("The assessment team have reviewed the the policy.")
                self.lt = c
                st.config(text=f"Status: running. Test sentence returned {len(hits)} grammar hit(s).")
            except Exception as e:  # noqa: BLE001
                st.config(text=f"Status: {e}")

        bf = ttk.Frame(frm)
        bf.grid(row=7, column=0, columnspan=3, sticky="w", pady=6)
        ttk.Button(bf, text="Download LanguageTool", command=do_download).pack(side="left")
        ttk.Button(bf, text="Test LanguageTool", command=do_test).pack(side="left", padx=6)
        ttk.Separator(frm).grid(row=8, column=0, columnspan=3, sticky="we", pady=8)
        ttk.Label(frm, text="Rules and dictionary", font=("Segoe UI", 10, "bold")).grid(row=9, column=0, columnspan=3, sticky="w")
        ttk.Label(frm, text="Openers, prohibited phrases, product terms, Findings label format, and thresholds live in rules.json.", foreground="#555").grid(row=10, column=0, columnspan=3, sticky="w")
        bf2 = ttk.Frame(frm)
        bf2.grid(row=11, column=0, columnspan=3, sticky="w", pady=6)
        ttk.Button(bf2, text="Edit Rules", command=self.edit_rules).pack(side="left")
        ttk.Button(bf2, text="Edit Custom Dictionary", command=lambda: open_path(user_dictionary_path())).pack(side="left", padx=6)
        ttk.Button(bf2, text="Open Settings Folder", command=lambda: open_path(user_dir())).pack(side="left")

        def ok():
            self.settings.update(lt_folder=v_folder.get(), java=v_java.get(), lt_port=int(v_port.get() or 8081))
            save_settings(self.settings)
            d.destroy()
        ttk.Button(frm, text="Save", style="Accent.TButton", command=ok).grid(row=12, column=2, sticky="e", pady=(10, 0))

    def show_help(self):
        p = os.path.join(APP_DIR, "README.md")
        if os.path.exists(p):
            open_path(p)
        else:
            messagebox.showinfo(APP_TITLE, "1. Open the results workbook.\n2. Set scope and Run Checks.\n3. Review each issue; accept, edit, or ignore.\n4. Save Corrected Copy (new versioned file).\n5. Export or print the report.")

    def on_close(self):
        if self.res and self.res.changes and not messagebox.askyesno(APP_TITLE, "You have unsaved changes. Exit anyway?"):
            return
        if self.lt:
            self.lt.stop()
        self.destroy()


def main(initial_file=None):
    App(initial_file).mainloop()


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
