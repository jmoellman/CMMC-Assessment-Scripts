"""
CMMC Pre-Assessment Package Checker (Windows desktop GUI, Python + Tkinter)

Run:   python app.py            (GUI)
       python app.py --cli ...  (command line, see --help)
"""
from __future__ import annotations

import argparse
import datetime as _dt
import os
import re
import subprocess
import sys
import threading
import warnings
import webbrowser

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import engine
import reports

NAVY, ORANGE, CHARCOAL, LIGHT, MID, TINT, DARKTINT = "#292C35", "#F7941D", "#53555F", "#F4F4F5", "#9A9BA3", "#FEF0DC", "#3D404C"
SLOT_ORDER = ["pa", "ai", "drl", "sap"]
SLOT_HINT = {
    "pa": "eMASS CMMC_Level2_PreAssessment_Form (.xlsx)",
    "ai": "Ignyte CMMC L2 Asset Inventory Scoping (.xlsx)",
    "drl": "Ignyte CMMC L2 Assessment Document Request List (.xlsx)",
    "sap": "Completed CMMC L2 Security Assessment Plan (.docx)",
}


def open_path(p):
    try:
        if sys.platform.startswith("win"):
            os.startfile(p)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", p])
        else:
            webbrowser.open("file://" + os.path.abspath(p))
    except Exception:
        webbrowser.open("file://" + os.path.abspath(p))


def report_dir_for(paths):
    """Keep reports beside the source files (they may contain CUI)."""
    first = next((paths[k] for k in SLOT_ORDER if paths.get(k)), None)
    base = os.path.dirname(os.path.abspath(first)) if first else os.path.expanduser("~")
    d = os.path.join(base, "Checker_Reports")
    os.makedirs(d, exist_ok=True)
    return d


def stem_for(res):
    osc = ""
    for lab, val, _ in res.data.get("pa_fields", []):
        if lab == "OSC Name":
            osc = val
    osc = osc or res.data.get("sap_osc_name", "") or "OSC"
    osc = re.sub(r"[^A-Za-z0-9]+", "_", osc).strip("_")[:40] or "OSC"
    return f"PreAssessment_Check_{osc}_{_dt.datetime.now():%Y%m%d_%H%M}"


# --------------------------------------------------------------------------- #
# GUI
# --------------------------------------------------------------------------- #
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    try:
        rules = engine.load_rules()
    except Exception as e:
        tk.Tk().withdraw()
        messagebox.showerror("Rules file", f"config/rules.json could not be loaded:\n{e}")
        return

    root = tk.Tk()
    root.title(f"{engine.APP_NAME} v{engine.APP_VERSION}")
    root.geometry("1280x820")
    root.minsize(1000, 640)
    root.configure(bg="white")
    try:
        root.iconbitmap(engine.resource_path("assets", "app.ico"))
    except Exception:
        pass

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    base_font = ("Segoe UI", 10)
    root.option_add("*Font", base_font)
    style.configure(".", background="white", foreground=CHARCOAL, font=base_font)
    style.configure("TFrame", background="white")
    style.configure("Card.TLabelframe", background="white", bordercolor="#E5E5E7")
    style.configure("Card.TLabelframe.Label", background="white", foreground=NAVY, font=("Segoe UI Semibold", 11))
    style.configure("TLabel", background="white", foreground=CHARCOAL)
    style.configure("Muted.TLabel", foreground=MID)
    style.configure("Count.TLabel", font=("Segoe UI Semibold", 11))
    style.configure("TButton", padding=(10, 5))
    style.configure("Accent.TButton", background=ORANGE, foreground=NAVY, font=("Segoe UI Semibold", 11), padding=(18, 8))
    style.map("Accent.TButton", background=[("active", "#ffa940"), ("disabled", LIGHT)])
    style.configure("Treeview", rowheight=24, fieldbackground="white")
    style.configure("Treeview.Heading", background=NAVY, foreground="white", font=("Segoe UI Semibold", 10))
    style.map("Treeview.Heading", background=[("active", DARKTINT)])
    style.map("Treeview", background=[("selected", "#ffd9a6")], foreground=[("selected", NAVY)])

    state = {"paths": {k: None for k in SLOT_ORDER}, "res": None}

    # Header
    hdr = tk.Frame(root, bg=NAVY, height=64)
    hdr.pack(fill="x")
    tk.Label(hdr, text="IGNYTE ASSURANCE PLATFORM", bg=NAVY, fg=ORANGE, font=("Segoe UI Semibold", 9)).pack(anchor="w", padx=18, pady=(8, 0))
    tk.Label(hdr, text="CMMC Level 2 Pre-Assessment Package Checker", bg=NAVY, fg="white", font=("Segoe UI Semibold", 15)).pack(anchor="w", padx=18)
    tk.Label(hdr, text="Load the OSC files, run the checks, then view, print, or export the issues. Source files are never changed.",
             bg=NAVY, fg="#d9dade", font=("Segoe UI", 9)).pack(anchor="w", padx=18, pady=(0, 8))
    tk.Frame(root, bg=ORANGE, height=4).pack(fill="x")

    body = ttk.Frame(root, padding=(14, 10))
    body.pack(fill="both", expand=True)

    # Step 1: files
    f1 = ttk.LabelFrame(body, text="  1.  Load files  ", style="Card.TLabelframe", padding=(10, 6))
    f1.pack(fill="x")
    bar = ttk.Frame(f1)
    bar.grid(row=0, column=0, columnspan=5, sticky="w", pady=(0, 6))
    path_vars, status_lbls = {}, {}

    def set_slot(key, p):
        state["paths"][key] = p
        path_vars[key].set(p or "")
        status_lbls[key].configure(text="Loaded" if p else "Not loaded", foreground=NAVY if p else MID)

    def auto_add(files):
        unknown = []
        for p in files:
            k = engine.detect_file(p)
            if k:
                set_slot(k, p)
            else:
                unknown.append(os.path.basename(p))
        if unknown:
            messagebox.showwarning("Not recognized",
                                   "These files were not recognized as one of the four package files:\n\n" + "\n".join(unknown) +
                                   "\n\nUse the Browse button on the right row to assign a file manually.")

    def add_files():
        fs = filedialog.askopenfilenames(title="Select the OSC package files",
                                         filetypes=[("Excel and Word", "*.xlsx *.xlsm *.docx"), ("All files", "*.*")])
        if fs:
            auto_add(list(fs))

    def add_folder():
        d = filedialog.askdirectory(title="Select the folder with the OSC files")
        if not d:
            return
        fs = []
        for name in sorted(os.listdir(d)):
            if name.startswith("~$") or "_corrected_v" in name:
                continue
            if name.lower().endswith((".xlsx", ".xlsm", ".docx")):
                fs.append(os.path.join(d, name))
        auto_add(fs)

    def clear_all():
        for k in SLOT_ORDER:
            set_slot(k, None)
        state["res"] = None
        refresh_results()

    ttk.Button(bar, text="Add files...", command=add_files).pack(side="left")
    ttk.Button(bar, text="Load a folder...", command=add_folder).pack(side="left", padx=6)
    ttk.Button(bar, text="Clear all", command=clear_all).pack(side="left")
    ttk.Label(bar, text="Files are detected automatically. Any file can be left out; checks run on what is loaded.",
              style="Muted.TLabel").pack(side="left", padx=12)

    for r, key in enumerate(SLOT_ORDER, start=1):
        ttk.Label(f1, text=engine.FILE_LABELS[key], font=("Segoe UI Semibold", 10)).grid(row=r, column=0, sticky="w", padx=(0, 10), pady=2)
        v = tk.StringVar()
        path_vars[key] = v
        e = ttk.Entry(f1, textvariable=v, state="readonly")
        e.grid(row=r, column=1, sticky="ew", pady=2)

        def browse(k=key):
            ft = [("Word document", "*.docx")] if k == "sap" else [("Excel workbook", "*.xlsx *.xlsm")]
            p = filedialog.askopenfilename(title=SLOT_HINT[k], filetypes=ft + [("All files", "*.*")])
            if p:
                set_slot(k, p)

        ttk.Button(f1, text="Browse...", command=browse).grid(row=r, column=2, padx=4)
        ttk.Button(f1, text="Remove", command=lambda k=key: set_slot(k, None)).grid(row=r, column=3, padx=(0, 6))
        s = ttk.Label(f1, text="Not loaded", style="Muted.TLabel", width=11)
        s.grid(row=r, column=4, sticky="w")
        status_lbls[key] = s
    f1.columnconfigure(1, weight=1)

    # Step 2: run
    f2 = ttk.Frame(body, padding=(0, 8))
    f2.pack(fill="x")
    run_btn = ttk.Button(f2, text="2.  Run checks", style="Accent.TButton")
    run_btn.pack(side="left")
    prog = ttk.Progressbar(f2, mode="indeterminate", length=160)
    counts_var = tk.StringVar(value="No results yet.")
    ttk.Label(f2, textvariable=counts_var, style="Count.TLabel").pack(side="left", padx=16)

    # Step 3: results
    f3 = ttk.LabelFrame(body, text="  3.  Review issues  ", style="Card.TLabelframe", padding=(10, 6))
    f3.pack(fill="both", expand=True)
    flt = ttk.Frame(f3)
    flt.pack(fill="x", pady=(0, 6))
    sev_var, file_var, own_var, q_var = tk.StringVar(value="All"), tk.StringVar(value="All"), tk.StringVar(value="All"), tk.StringVar()
    ttk.Label(flt, text="Severity").pack(side="left")
    ttk.Combobox(flt, textvariable=sev_var, values=["All"] + engine.SEVERITIES, width=9, state="readonly").pack(side="left", padx=(4, 12))
    ttk.Label(flt, text="File").pack(side="left")
    ttk.Combobox(flt, textvariable=file_var, values=["All"] + list(engine.FILE_LABELS.values()), width=24, state="readonly").pack(side="left", padx=(4, 12))
    ttk.Label(flt, text="Owner").pack(side="left")
    ttk.Combobox(flt, textvariable=own_var, values=["All", "OSC", "Ignyte"], width=8, state="readonly").pack(side="left", padx=(4, 12))
    ttk.Label(flt, text="Search").pack(side="left")
    ttk.Entry(flt, textvariable=q_var, width=28).pack(side="left", padx=4)

    pane = ttk.PanedWindow(f3, orient="vertical")
    pane.pack(fill="both", expand=True)
    tv_frame = ttk.Frame(pane)
    cols = ("sev", "owner", "file", "loc", "field", "issue")
    tv = ttk.Treeview(tv_frame, columns=cols, show="headings", selectmode="browse")
    heads = {"sev": ("Severity", 80), "owner": ("Owner", 70), "file": ("File", 170), "loc": ("Location", 210),
             "field": ("Field", 200), "issue": ("Issue", 560)}
    for c in cols:
        tv.heading(c, text=heads[c][0])
        tv.column(c, width=heads[c][1], stretch=(c == "issue"), anchor="w")
    tv.tag_configure("Error", background="#FDE3C2")
    tv.tag_configure("Warning", background=TINT)
    tv.tag_configure("Info", background="white")
    ys = ttk.Scrollbar(tv_frame, orient="vertical", command=tv.yview)
    tv.configure(yscrollcommand=ys.set)
    tv.pack(side="left", fill="both", expand=True)
    ys.pack(side="right", fill="y")
    pane.add(tv_frame, weight=4)

    det = tk.Text(pane, height=8, wrap="word", bg=LIGHT, fg=NAVY, relief="flat", padx=10, pady=8, font=("Segoe UI", 10))
    det.tag_configure("h", font=("Segoe UI Semibold", 10), foreground=NAVY)
    det.tag_configure("v", font=("Consolas", 10))
    det.configure(state="disabled")
    pane.add(det, weight=1)

    shown = []

    def refresh_results(*_):
        tv.delete(*tv.get_children())
        shown.clear()
        res = state["res"]
        if not res:
            counts_var.set("No results yet.")
            return
        c = res.counts()
        counts_var.set(f"Errors: {c['Error']}    Warnings: {c['Warning']}    Info: {c['Info']}    Auto-fixable (Pre-Assessment): {c['Auto-fixable']}")
        q = q_var.get().strip().lower()
        for i in res.sorted_issues():
            if sev_var.get() != "All" and i.severity != sev_var.get():
                continue
            if file_var.get() != "All" and engine.FILE_LABELS[i.file_key] != file_var.get():
                continue
            if own_var.get() != "All" and i.owner != own_var.get():
                continue
            if q and q not in " ".join([i.location, i.field, i.message, i.value, i.fix]).lower():
                continue
            shown.append(i)
            tv.insert("", "end", iid=str(len(shown) - 1), tags=(i.severity,),
                      values=(i.severity, i.owner, engine.FILE_LABELS[i.file_key], i.location, i.field, i.message))

    for v in (sev_var, file_var, own_var, q_var):
        v.trace_add("write", refresh_results)

    def on_select(_=None):
        sel = tv.selection()
        det.configure(state="normal")
        det.delete("1.0", "end")
        if sel:
            i = shown[int(sel[0])]
            det.insert("end", f"{i.severity}  |  Owner: {i.owner}  |  {engine.FILE_LABELS[i.file_key]}  |  {i.location}\n", "h")
            det.insert("end", f"{i.field}: {i.message}\n")
            if i.value:
                det.insert("end", "Found: ", "h")
                det.insert("end", i.value + "\n", "v")
            if i.expected:
                det.insert("end", "Expected: ", "h")
                det.insert("end", i.expected + "\n")
            if i.fix:
                det.insert("end", "How to resolve: ", "h")
                det.insert("end", i.fix + "\n")
            if i.autofix:
                det.insert("end", "Auto-fix available: ", "h")
                det.insert("end", i.autofix["desc"] + " (use 'Save corrected Pre-Assessment')\n")
        det.configure(state="disabled")

    tv.bind("<<TreeviewSelect>>", on_select)

    # Step 4: outputs
    f4 = ttk.Frame(body, padding=(0, 8, 0, 0))
    f4.pack(fill="x")
    rep_scope = tk.StringVar(value="All items")
    ttk.Label(f4, text="Report contents").pack(side="left")
    ttk.Combobox(f4, textvariable=rep_scope, values=["All items", "OSC items only", "Ignyte items only"], width=16,
                 state="readonly").pack(side="left", padx=(4, 12))

    def need_res():
        if not state["res"]:
            messagebox.showinfo("Run checks first", "Load files and click Run checks first.")
            return False
        return True

    def owner_filter():
        return {"OSC items only": "OSC", "Ignyte items only": "Ignyte"}.get(rep_scope.get())

    def make_html(auto_print):
        d = report_dir_for(state["paths"])
        suffix = {"OSC": "_OSC", "Ignyte": "_Ignyte"}.get(owner_filter(), "")
        p = os.path.join(d, stem_for(state["res"]) + suffix + ".html")
        reports.html_report(state["res"], p, auto_print=auto_print, owner_filter=owner_filter())
        return p

    def view_report():
        if need_res():
            open_path(make_html(False))

    def print_report():
        if need_res():
            p = make_html(True)
            open_path(p)
            messagebox.showinfo("Print", "The report opened in your browser with the print dialog.\n"
                                         "Choose a printer, or choose 'Save as PDF'.\n\nSaved copy:\n" + p)

    def export_xlsx():
        if not need_res():
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", initialdir=report_dir_for(state["paths"]),
                                         initialfile=stem_for(state["res"]) + ".xlsx", filetypes=[("Excel", "*.xlsx")])
        if p:
            reports.excel_report(state["res"], p)
            if messagebox.askyesno("Exported", f"Saved:\n{p}\n\nOpen it now?"):
                open_path(p)

    def export_csv():
        if not need_res():
            return
        p = filedialog.asksaveasfilename(defaultextension=".csv", initialdir=report_dir_for(state["paths"]),
                                         initialfile=stem_for(state["res"]) + ".csv", filetypes=[("CSV", "*.csv")])
        if p:
            reports.csv_report(state["res"], p)
            messagebox.showinfo("Exported", f"Saved:\n{p}")

    def fix_dialog():
        if not need_res():
            return
        import xlsx_patch
        src = state["paths"].get("pa")
        fixes = [i.autofix for i in state["res"].sorted_issues() if i.autofix and i.file_key == "pa"]
        if not src or not fixes:
            messagebox.showinfo("Nothing to correct", "There are no automatic corrections for the Pre-Assessment form.")
            return
        win = tk.Toplevel(root)
        win.title("Save corrected Pre-Assessment form")
        win.geometry("820x520")
        win.configure(bg="white")
        win.transient(root)
        win.grab_set()
        ttk.Label(win, text="Select the corrections to apply. The original file is not changed; a new copy is saved.",
                  padding=(12, 10, 12, 4)).pack(anchor="w")
        ttk.Label(win, text="Items not listed here (for example missing required values) must be corrected by the OSC or Ignyte.",
                  style="Muted.TLabel", padding=(12, 0, 12, 6)).pack(anchor="w")
        outer = ttk.Frame(win)
        outer.pack(fill="both", expand=True, padx=12)
        cv = tk.Canvas(outer, bg="white", highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=cv.yview)
        inner = ttk.Frame(cv)
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.create_window((0, 0), window=inner, anchor="nw")
        cv.configure(yscrollcommand=sb.set)
        cv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        vars_ = []
        for fx in fixes:
            v = tk.BooleanVar(value=fx.get("default", True))
            vars_.append(v)
            ttk.Checkbutton(inner, text=fx["desc"], variable=v).pack(anchor="w", pady=1)
        btns = ttk.Frame(win, padding=12)
        btns.pack(fill="x")
        ttk.Button(btns, text="Select all", command=lambda: [v.set(True) for v in vars_]).pack(side="left")
        ttk.Button(btns, text="Select none", command=lambda: [v.set(False) for v in vars_]).pack(side="left", padx=6)

        def do_save():
            chosen = [fx for fx, v in zip(fixes, vars_) if v.get()]
            if not chosen:
                messagebox.showinfo("Nothing selected", "Select at least one correction.", parent=win)
                return
            default = xlsx_patch.next_version_path(src)
            dst = filedialog.asksaveasfilename(parent=win, defaultextension=".xlsx", initialdir=os.path.dirname(default),
                                               initialfile=os.path.basename(default), filetypes=[("Excel", "*.xlsx")])
            if not dst:
                return
            if os.path.abspath(dst) == os.path.abspath(src):
                messagebox.showerror("Not allowed", "Choose a new file name. The original is never overwritten.", parent=win)
                return
            try:
                applied = xlsx_patch.apply_fixes(src, dst, chosen)
            except Exception as e:
                messagebox.showerror("Save failed", f"The corrected copy could not be saved:\n{e}\n\nIf the file is open in Excel, close it and try again.", parent=win)
                return
            win.destroy()
            if messagebox.askyesno("Saved", f"Applied {len(applied)} correction(s).\n\nSaved:\n{dst}\n\n"
                                            "Load the corrected copy and run the checks again?"):
                set_slot("pa", dst)
                start_run()

        ttk.Button(btns, text="Save corrected copy...", style="Accent.TButton", command=do_save).pack(side="right")
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="right", padx=6)

    ttk.Button(f4, text="View report", command=view_report).pack(side="left")
    ttk.Button(f4, text="Print report", command=print_report).pack(side="left", padx=6)
    ttk.Button(f4, text="Export to Excel", command=export_xlsx).pack(side="left")
    ttk.Button(f4, text="Export to CSV", command=export_csv).pack(side="left", padx=6)
    ttk.Button(f4, text="Save corrected Pre-Assessment...", command=fix_dialog).pack(side="right")
    ttk.Button(f4, text="Open reports folder",
               command=lambda: open_path(report_dir_for(state["paths"])) if any(state["paths"].values()) else None).pack(side="right", padx=6)

    # Run
    def start_run():
        if not any(state["paths"].values()):
            messagebox.showinfo("No files", "Load at least one file first.")
            return
        for k, p in state["paths"].items():
            if p and not os.path.exists(p):
                messagebox.showerror("File missing", f"This file no longer exists:\n{p}")
                return
        run_btn.configure(state="disabled")
        prog.pack(side="left", padx=10)
        prog.start(12)
        counts_var.set("Checking...")
        out = {}

        def work():
            try:
                out["res"] = engine.run_all(dict(state["paths"]), rules)
            except Exception as e:
                out["err"] = e

        th = threading.Thread(target=work, daemon=True)
        th.start()

        def poll():
            if th.is_alive():
                root.after(120, poll)
                return
            prog.stop()
            prog.pack_forget()
            run_btn.configure(state="normal")
            if "err" in out:
                messagebox.showerror("Error", f"The checks could not complete:\n{out['err']}")
                counts_var.set("Check failed.")
                return
            state["res"] = out["res"]
            refresh_results()
            if shown:
                tv.selection_set("0")
                tv.see("0")

        poll()

    run_btn.configure(command=start_run)

    # Files passed on the command line (e.g. dragged onto the .bat/.exe)
    if len(sys.argv) > 1:
        auto_add([a for a in sys.argv[1:] if os.path.isfile(a)])

    root.mainloop()


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def run_cli(argv):
    ap = argparse.ArgumentParser(prog="app.py --cli", description=engine.APP_NAME)
    ap.add_argument("files", nargs="*", help="Files to auto-detect")
    ap.add_argument("--pa"), ap.add_argument("--ai"), ap.add_argument("--drl"), ap.add_argument("--sap")
    ap.add_argument("--html", help="Write HTML report to this path")
    ap.add_argument("--xlsx", help="Write Excel issue log to this path")
    ap.add_argument("--csv", help="Write CSV issue log to this path")
    ap.add_argument("--owner", choices=["OSC", "Ignyte"], help="Limit the HTML report to one owner")
    ap.add_argument("--fix-out", help="Save a corrected Pre-Assessment copy with all default fixes to this path")
    a = ap.parse_args(argv)
    paths = {k: getattr(a, k) for k in SLOT_ORDER}
    for f in a.files:
        k = engine.detect_file(f)
        if k and not paths.get(k):
            paths[k] = f
        elif not k:
            print(f"Not recognized: {f}")
    res = engine.run_all(paths)
    c = res.counts()
    print(f"Errors: {c['Error']}  Warnings: {c['Warning']}  Info: {c['Info']}  Auto-fixable: {c['Auto-fixable']}")
    for i in res.sorted_issues():
        print(f"[{i.severity}] {engine.FILE_LABELS[i.file_key]} | {i.location} | {i.field} | {i.message}")
    if a.html:
        reports.html_report(res, a.html, owner_filter=a.owner)
    if a.xlsx:
        reports.excel_report(res, a.xlsx)
    if a.csv:
        reports.csv_report(res, a.csv)
    if a.fix_out and paths.get("pa"):
        import xlsx_patch
        fx = [i.autofix for i in res.sorted_issues() if i.autofix and i.autofix.get("default", True)]
        print(f"Applied {len(xlsx_patch.apply_fixes(paths['pa'], a.fix_out, fx))} fix(es) -> {a.fix_out}")
    return 1 if c["Error"] else 0


if __name__ == "__main__":
    if "--cli" in sys.argv:
        argv = [x for x in sys.argv[1:] if x != "--cli"]
        sys.exit(run_cli(argv))
    run_gui()
