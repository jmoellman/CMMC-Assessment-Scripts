"""CMMC eMASS Results Checker: entry point.

    python emass_checker.py                         -> opens the desktop app
    python emass_checker.py results.xlsx            -> opens the app with a file loaded
    python emass_checker.py check results.xlsx [options]   -> command-line check, no GUI
    python emass_checker.py --selftest              -> verifies the install
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def cli(argv):
    from engine import DOMAINS, Engine, Results, Scope
    from textcheck import LanguageToolClient, TextChecker, load_rules
    import reports
    import xlsx_patch

    ap = argparse.ArgumentParser(prog="emass_checker check", description="Check a CMMC eMASS Assessment Results workbook.")
    ap.add_argument("file")
    ap.add_argument("--domains", help="Comma list, for example AC,AU (default: all)")
    ap.add_argument("--from", dest="obj_from", help="First objective, for example AC.L2-3.1.1[a]")
    ap.add_argument("--to", dest="obj_to", help="Last objective, for example AC.L2-3.1.22[e]")
    ap.add_argument("--no-assessment", action="store_true")
    ap.add_argument("--no-requirements", action="store_true")
    ap.add_argument("--no-ssp", action="store_true")
    ap.add_argument("--no-spelling", action="store_true")
    ap.add_argument("--languagetool", metavar="FOLDER", help="Use local LanguageTool from this folder")
    ap.add_argument("--html", help="Write HTML report")
    ap.add_argument("--xlsx", help="Write Excel report")
    ap.add_argument("--csv", help="Write CSV report")
    ap.add_argument("--apply-safe", action="store_true", help="Apply safe mechanical fixes and save a new versioned copy")
    ap.add_argument("--output", help="Path for the corrected copy (default: <name>_corrected_vN.xlsx)")
    a = ap.parse_args(argv)

    res = Results(a.file)
    rules = load_rules()
    tc = TextChecker(rules)
    lt = None
    if a.languagetool:
        lt = LanguageToolClient(a.languagetool, disabled_rules=rules.get("languagetool_disabled_rules"))
        lt.start()
        tc.lt = lt
    eng = Engine(res, {"spelling": not a.no_spelling, "languagetool": bool(lt)}, tc)
    doms = [d.strip().upper() for d in a.domains.split(",")] if a.domains else DOMAINS
    eng.run(Scope(doms, a.obj_from, a.obj_to, not a.no_assessment, not a.no_requirements, not a.no_ssp))
    s = eng.summary()
    print(f"{os.path.basename(a.file)}: {s['by_severity'].get('Error', 0)} errors, {s['by_severity'].get('Warning', 0)} warnings, {s['by_severity'].get('Info', 0)} info")
    if a.apply_safe:
        n = eng.apply_safe()
        out = a.output or xlsx_patch.next_version_path(a.file)
        xlsx_patch.write_copy(a.file, out, res.changes)
        print(f"Applied {n} safe fixes to {len(res.changes)} cells -> {out}")
    if a.html:
        reports.write_html(a.html, eng, include_closed=True)
    if a.xlsx:
        reports.write_xlsx(a.xlsx, eng, include_closed=True)
    if a.csv:
        reports.write_csv(a.csv, eng.issues, include_closed=True)
    if not (a.html or a.xlsx or a.csv):
        for i in eng.issues:
            if i.status == "Open":
                print(f"[{i.severity}] {i.location} {i.rule}: {i.message}")
    if lt:
        lt.stop()
    return 1 if s["by_severity"].get("Error") else 0


def selftest():
    print("Python", sys.version.split()[0])
    try:
        import openpyxl
        print("openpyxl", openpyxl.__version__, "OK")
    except ImportError:
        print("openpyxl MISSING: run  py -m pip install openpyxl")
        return 1
    from textcheck import TextChecker
    tc = TextChecker()
    hits = tc.check("The assessor reviewd the policy , which is a best practice.", "Examine")
    print("Text checks OK:", sorted({h.rule for h in hits}))
    try:
        import tkinter
        print("tkinter", tkinter.TkVersion, "OK")
    except ImportError:
        print("tkinter MISSING: reinstall Python from python.org with 'tcl/tk and IDLE' checked")
        return 1
    print("Self-test passed.")
    return 0


if __name__ == "__main__":
    args = sys.argv[1:]
    if args and args[0] == "check":
        sys.exit(cli(args[1:]))
    if args and args[0] == "--selftest":
        sys.exit(selftest())
    from app import main
    main(args[0] if args else None)
