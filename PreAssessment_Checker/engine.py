"""
CMMC Pre-Assessment Package Checker - validation engine.

Checks four OSC-supplied files against the Ignyte Security Assessment Plan (SAP)
instructions and the eMASS Pre-Assessment template rules:
  1. eMASS CMMC Level 2 Pre-Assessment Form (first tab)
  2. Ignyte CMMC L2 Asset Inventory Scoping workbook
  3. Ignyte CMMC L2 Assessment Document Request List (DRL)
  4. Completed CMMC L2 Security Assessment Plan (.docx)
plus cross-file consistency checks.

Nothing here writes to the source files. Proposed corrections for the
Pre-Assessment form are attached to issues as `fix` dicts and applied by
xlsx_patch.py only after the user approves them.
"""
from __future__ import annotations

import datetime as _dt
import hashlib
import json
import os
import re
from collections import OrderedDict, defaultdict
from dataclasses import dataclass, field, asdict
from typing import Any

import openpyxl

try:
    import docx  # python-docx
except Exception:  # pragma: no cover
    docx = None

APP_NAME = "CMMC Pre-Assessment Package Checker"
APP_VERSION = "1.0.0"

FILE_LABELS = OrderedDict([
    ("pa", "Pre-Assessment Form"),
    ("ai", "Asset Inventory Scoping"),
    ("drl", "Document Request List"),
    ("sap", "Security Assessment Plan"),
    ("x", "Cross-File"),
])

SEVERITIES = ["Error", "Warning", "Info"]
SEV_RANK = {s: i for i, s in enumerate(SEVERITIES)}
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# --------------------------------------------------------------------------- #
# Data model
# --------------------------------------------------------------------------- #
@dataclass
class Issue:
    severity: str            # Error | Warning | Info
    file_key: str            # pa | ai | drl | sap | x
    location: str            # sheet!cell, paragraph, etc.
    field: str               # short field / topic name
    message: str             # what is wrong
    value: str = ""          # what was found
    expected: str = ""       # what is expected
    fix: str = ""            # how to resolve
    owner: str = "OSC"       # OSC | Ignyte
    rule: str = ""           # rule id
    autofix: dict | None = None   # proposed correction (Pre-Assessment form only)

    def to_dict(self):
        d = asdict(self)
        d["file"] = FILE_LABELS.get(self.file_key, self.file_key)
        d["autofixable"] = bool(self.autofix)
        return d


@dataclass
class Result:
    files: dict = field(default_factory=dict)      # key -> {"path","name","sha256","detected"}
    issues: list = field(default_factory=list)
    data: dict = field(default_factory=dict)       # captured values for the report
    generated: str = ""

    def add(self, *a, **k):
        self.issues.append(Issue(*a, **k))

    def counts(self):
        c = {s: 0 for s in SEVERITIES}
        for i in self.issues:
            c[i.severity] = c.get(i.severity, 0) + 1
        c["Auto-fixable"] = sum(1 for i in self.issues if i.autofix)
        return c

    def sorted_issues(self):
        order = list(FILE_LABELS.keys())
        return sorted(self.issues, key=lambda i: (order.index(i.file_key) if i.file_key in order else 99,
                                                  SEV_RANK.get(i.severity, 9)))


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def app_dir():
    import sys
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resource_path(*parts):
    """Prefer an editable copy beside the app/exe, then the bundled copy."""
    import sys
    p = os.path.join(app_dir(), *parts)
    if os.path.exists(p):
        return p
    return os.path.join(getattr(sys, "_MEIPASS", app_dir()), *parts)


def load_rules(path: str | None = None) -> dict:
    path = path or resource_path("config", "rules.json")
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def as_text(v) -> str:
    """Render a cell value as the text a reviewer would see."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (_dt.datetime, _dt.date)):
        return fmt_date(v)
    return str(v)


def fmt_date(d) -> str:
    return f"{d.day:02d}-{MONTHS[d.month - 1]}-{d.year:04d}"


def short(s, n=160):
    s = as_text(s).replace("\n", " / ")
    return s if len(s) <= n else s[: n - 3] + "..."


EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")
URL_RE = re.compile(r"^https?://[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+(:\d+)?(/\S*)?$", re.I)
DATE_DMY_RE = re.compile(r"^(\d{1,2})-([A-Za-z]{3})-(\d{4})$")
PLACEHOLDER_RE = re.compile(r"^(tbd|tba|n/?a|none|x+|\?+|-+|pending|<.*>|\[.*\])$", re.I)


def parse_date(v):
    """Return (date or None, is_text_in_required_format)."""
    if isinstance(v, (_dt.datetime, _dt.date)):
        return (v.date() if isinstance(v, _dt.datetime) else v), False
    s = as_text(v).strip()
    if not s:
        return None, False
    m = DATE_DMY_RE.match(s)
    if m:
        mon = m.group(2).capitalize()
        if mon in MONTHS:
            try:
                return _dt.date(int(m.group(3)), MONTHS.index(mon) + 1, int(m.group(1))), m.group(2) == mon
            except ValueError:
                return None, False
    for f in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%d %b %Y", "%B %d, %Y", "%b %d, %Y", "%d-%B-%Y", "%m-%d-%Y"):
        try:
            return _dt.datetime.strptime(s, f).date(), False
        except ValueError:
            pass
    return None, False


def find_sheet(wb, name_prefix):
    """Sheet names are truncated to 31 chars, so match by normalized prefix."""
    n = norm(name_prefix)
    for ws in wb.worksheets:
        if norm(ws.title) == n:
            return ws
    for ws in wb.worksheets:
        if norm(ws.title).startswith(n[:25]) or n.startswith(norm(ws.title)):
            return ws
    return None


def header_map(ws, row, keywords: dict):
    """Map logical names to column indexes by header keyword (case-insensitive contains)."""
    out = {}
    hdr = {c.column: norm(c.value) for c in ws[row] if c.value is not None}
    for key, kw in keywords.items():
        kws = kw if isinstance(kw, (list, tuple)) else [kw]
        for col, text in hdr.items():
            if any(norm(k) in text for k in kws):
                out[key] = col
                break
    return out


def row_values(ws, r, cols):
    return {k: ws.cell(row=r, column=c).value for k, c in cols.items()}


def is_blank(v):
    return as_text(v).strip() == ""


def split_multi(s):
    return [p.strip() for p in re.split(r"[;\n]+", as_text(s)) if p.strip()]


def family_of(obj_id: str):
    m = re.match(r"3\.(\d+)\.", str(obj_id))
    fams = {1: "AC", 2: "AT", 3: "AU", 4: "CM", 5: "IA", 6: "IR", 7: "MA", 8: "MP", 9: "PS",
            10: "PE", 11: "RA", 12: "CA", 13: "SC", 14: "SI"}
    if m:
        return fams.get(int(m.group(1)), "??")
    m = re.match(r"([A-Z]{2})\.", str(obj_id))
    return m.group(1) if m else "??"


def compress_ids(ids, limit=40):
    ids = list(ids)
    s = ", ".join(ids[:limit])
    if len(ids) > limit:
        s += f", and {len(ids) - limit} more"
    return s


# --------------------------------------------------------------------------- #
# File detection
# --------------------------------------------------------------------------- #
def detect_file(path):
    """Return file key (pa, ai, drl, sap) or None."""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in (".xlsx", ".xlsm"):
            wb = openpyxl.load_workbook(path, read_only=True)
            names = [norm(n) for n in wb.sheetnames]
            wb.close()
            if any(n.startswith("pre-assessment") for n in names):
                return "pa"
            if any(n.startswith("main scoped assets") for n in names) or any(n.startswith("consolidated external") for n in names):
                return "ai"
            if "evidence plan" in names or "osc assessment team" in names:
                return "drl"
        elif ext == ".docx" and docx is not None:
            d = docx.Document(path)
            text = " ".join(p.text for p in d.paragraphs[:80])
            if "Security Assessment Plan" in text or "Assessment Plan" in text:
                return "sap"
    except Exception:
        return None
    return None


# --------------------------------------------------------------------------- #
# 1. Pre-Assessment form
# --------------------------------------------------------------------------- #
class PreAssessmentChecker:
    def __init__(self, path, rules, res: Result):
        self.path, self.rules, self.res = path, rules, res
        self.cfg = rules["pre_assessment"]
        self.values = {}      # field key -> effective value (raw)
        self.cells = {}       # field key -> dict(row, input_cell, guide_cell, required)
        self.fields = {}      # captured, display

    def add(self, sev, loc, fld, msg, **k):
        self.res.add(sev, "pa", loc, fld, msg, **k)

    def run(self):
        wb = openpyxl.load_workbook(self.path)  # formulas as written
        ws = find_sheet(wb, "Pre-Assessment")
        if ws is None:
            self.add("Error", "workbook", "Sheet", "The 'Pre-Assessment' worksheet was not found.",
                     fix="Return the eMASS CMMC_Level2_PreAssessment_Form with the Pre-Assessment tab intact.")
            return
        self.ws = ws
        # Template version
        ver = None
        for row in ws.iter_rows(min_row=1, max_row=6):
            for c in row:
                m = re.search(r"Template Version\s*([\d.]+)", as_text(c.value))
                if m:
                    ver = m.group(1)
        self.res.data["pa_template_version"] = ver or "Not found"
        if not ver:
            self.add("Warning", f"{ws.title}!A1:E6", "Template version",
                     "The eMASS template version label was not found.",
                     expected="'Provided by eMASS / Template Version x.x' in the header",
                     fix="Confirm the OSC used the current eMASS Pre-Assessment template and did not alter the header.",
                     owner="Ignyte", rule="PA-VER")
        elif ver not in self.cfg["expected_template_versions"]:
            self.add("Warning", f"{ws.title}!header", "Template version",
                     f"Template version {ver} differs from the expected version list.",
                     value=ver, expected=", ".join(self.cfg["expected_template_versions"]),
                     fix="Confirm this is the current eMASS template version, or update expected_template_versions in config/rules.json.",
                     owner="Ignyte", rule="PA-VER")
        if find_sheet(wb, "Lookup Values") is None:
            self.add("Warning", "workbook", "Lookup Values",
                     "The hidden 'Lookup Values' worksheet is missing, so the dropdowns may not work.",
                     fix="Use an unmodified copy of the eMASS template.", owner="Ignyte", rule="PA-LKP")

        # Locate header row and columns
        hdr_row, cols = None, {}
        for r in range(1, 12):
            vals = {c.column: norm(c.value) for c in ws[r] if c.value is not None}
            if any(v.startswith(norm(self.cfg["label_header"])) for v in vals.values()):
                hdr_row = r
                for col, v in vals.items():
                    if v.startswith(norm(self.cfg["label_header"])):
                        cols["label"] = col
                    elif v.startswith("required"):
                        cols["req"] = col
                    elif v.startswith(norm(self.cfg["guidance_header"])) or v.startswith("values"):
                        cols["guide"] = col
                    elif v == norm(self.cfg["input_header"]):
                        cols["input"] = col
                break
        if not hdr_row or "input" not in cols or "label" not in cols:
            self.add("Error", ws.title, "Layout",
                     "The header row (Data Field / Required/Optional / Acceptable Values / Input) was not found.",
                     fix="Use an unmodified copy of the eMASS Pre-Assessment template.", rule="PA-HDR")
            return
        self.hdr_row, self.cols = hdr_row, cols
        L = openpyxl.utils.get_column_letter
        label_rows = {}
        for r in range(hdr_row + 1, ws.max_row + 1):
            lab = norm(ws.cell(r, cols["label"]).value)
            if lab and lab not in label_rows:
                label_rows[lab] = r

        for key, label in self.cfg["fields"].items():
            r = label_rows.get(norm(label))
            if r is None:
                self.add("Error", ws.title, label, f"Row '{label}' was not found in the form.",
                         fix="Do not delete or rename rows in the eMASS template.", rule="PA-ROW",
                         owner="Ignyte" if key in self.cfg["ignyte_owned_fields"] else "OSC")
                continue
            req_txt = as_text(ws.cell(r, cols.get("req", 2)).value)
            info = dict(row=r, label=label,
                        input_cell=f"{L(cols['input'])}{r}",
                        guide_cell=f"{L(cols['guide'])}{r}" if "guide" in cols else None,
                        required=req_txt.lower().startswith("required"),
                        owner="Ignyte" if key in self.cfg["ignyte_owned_fields"] else "OSC")
            self.cells[key] = info
            raw_in = ws.cell(r, cols["input"]).value
            raw_g = ws.cell(r, cols["guide"]).value if "guide" in cols else None
            eff = raw_in
            if is_blank(raw_in) and not is_blank(raw_g) and not self._is_guidance(key, raw_g):
                val = self._value_from_guidance(key, raw_g)
                eff = val
                cleaned = self._clean(key, val)
                note = "" if as_text(val).strip() == cleaned else f" (cleaned to '{short(cleaned, 50)}')"
                restore = self._restore_text(key, raw_g)
                self.add("Error" if info["owner"] == "OSC" else "Warning", f"{ws.title}!{info['guide_cell']}", label,
                         "Value was entered in the 'Acceptable Values' column instead of the 'Input' column. eMASS reads the Input column only, so this field would import as blank.",
                         value=short(raw_g), expected=f"Value in {info['input_cell']}",
                         fix=f"Move the value to {info['input_cell']}.", owner=info["owner"], rule="PA-COL",
                         autofix={"type": "move", "field": key, "label": label, "from": info["guide_cell"],
                                  "to": info["input_cell"], "value": self._clean(key, val),
                                  "restore": restore, "default": True,
                                  "desc": f"Move {label} from {info['guide_cell']} to {info['input_cell']}{note}"})
            self.values[key] = eff
            self.fields[key] = as_text(eff).strip()

        self.res.data["pa_fields"] = [(self.cfg["fields"][k], self.fields.get(k, ""), self.cells[k]["input_cell"] if k in self.cells else "")
                                      for k in self.cfg["fields"]]
        self._validate()

    # -- guidance detection ------------------------------------------------- #
    def _is_guidance(self, key, g):
        t = as_text(g).strip()
        if not t:
            return True
        tl = t.lower()
        if tl.startswith("use a fully complete"):
            return True
        if key in ("contract_date", "plan_start", "plan_end"):
            rest = re.sub(r"^format:\s*", "", t, flags=re.I).strip()
            return parse_date(rest)[0] is None
        opts = {"sector": self.rules["sectors"], "scope": self.cfg["scope_values"],
                "standard": self.cfg["standard_values"]}.get(key)
        if opts:
            hits = sum(1 for o in opts if o.lower() in tl)
            return hits >= 2
        return False

    def _value_from_guidance(self, key, g):
        if key in ("contract_date", "plan_start", "plan_end"):
            return re.sub(r"^format:\s*", "", as_text(g).strip(), flags=re.I).strip()
        return g

    def _restore_text(self, key, g):
        if key in ("contract_date", "plan_start", "plan_end"):
            return None  # leave the format column as-is
        rt = self.cfg.get("guidance_restore", {}).get(key)
        if rt == "__SECTOR_LIST__":
            rt = "\n".join(self.rules["sectors"])
        return rt if rt is not None else ""

    def _clean(self, key, v):
        """Best normalized form of a value, used by proposed fixes."""
        if v is None:
            return ""
        if key in ("contract_date", "plan_start", "plan_end"):
            d, _ = parse_date(v)
            return fmt_date(d) if d else as_text(v).strip()
        s = as_text(v).strip()
        s = re.sub(r"[ \t]*\n[ \t]*", "\n", s).strip()
        if key == "uei":
            return s.upper().replace(" ", "")
        if key in ("hlo_cage", "cage_scope"):
            parts = [p for p in re.split(r"[;,\s]+", s.upper()) if p]
            return "; ".join(parts)
        if key == "sector":
            out = []
            for p in split_multi(s):
                m = next((o for o in self.rules["sectors"] if o.lower() == p.lower()), p)
                out.append(m)
            return "; ".join(out)
        if key == "scope":
            return next((o for o in self.cfg["scope_values"] if o.lower() == s.lower()), s)
        if key == "standard":
            return next((o for o in self.cfg["standard_values"] if o.lower() == s.lower()), s)
        if key == "url":
            return re.sub(r"^(https?)//", r"\1://", s, flags=re.I)
        if key == "phone":
            d = re.sub(r"\D", "", s)
            if len(d) == 11 and d.startswith("1"):
                d = d[1:]
            if len(d) == 10:
                return f"{d[:3]}-{d[3:6]}-{d[6:]}"
            return s
        if key == "employees":
            return s
        return s

    # -- validation ---------------------------------------------------------- #
    def _fix(self, key, new, desc, default=True):
        # A pending "move" fix for this field already writes the cleaned value.
        if any(i.autofix and i.autofix.get("field") == key and i.autofix["type"] == "move"
               for i in self.res.issues if i.file_key == "pa"):
            return None
        c = self.cells[key]
        return {"type": "set", "field": key, "label": c["label"], "to": c["input_cell"], "value": new,
                "default": default, "desc": desc}

    def _validate(self):
        ws = self.ws
        f = self.fields
        cells = self.cells

        def loc(k):
            return f"{ws.title}!{cells[k]['input_cell']}" if k in cells else ws.title

        def has(k):
            return k in cells

        # Required fields and whitespace hygiene
        for k, c in cells.items():
            raw = self.values.get(k)
            txt = as_text(raw)
            moved = any(i.autofix and i.autofix.get("field") == k and i.autofix["type"] == "move"
                        for i in self.res.issues if i.file_key == "pa")
            if c["required"] and not txt.strip():
                if c["owner"] == "Ignyte":
                    self.add("Warning", loc(k), c["label"], "Required eMASS field is blank. Per SAP section 5.1, the Ignyte team completes this row.",
                             expected="Value entered by Ignyte before eMASS upload",
                             fix="Ignyte to complete before eMASS submission.", owner="Ignyte", rule="PA-REQ")
                else:
                    self.add("Error", loc(k), c["label"], "Required field is blank. eMASS denies a template import when a required field is blank.",
                             fix=f"Enter the {c['label']} in {c['input_cell']}.", rule="PA-REQ")
                continue
            if txt.strip() and c["required"] and PLACEHOLDER_RE.match(txt.strip()):
                self.add("Error", loc(k), c["label"], "Required field contains placeholder text instead of a real value.",
                         value=short(txt), fix="Replace the placeholder with the actual value.", owner=c["owner"], rule="PA-PH")
            if txt and txt != txt.strip() and not moved and not isinstance(raw, (int, float)):
                self.add("Warning", loc(k), c["label"], "Value has leading or trailing spaces or line breaks.",
                         value=repr(short(txt, 60)), fix="Remove the extra spaces or line breaks.", owner=c["owner"], rule="PA-WS",
                         autofix=self._fix(k, self._clean(k, txt), f"Trim spaces/line breaks in {c['input_cell']}"))

        # UEI
        if has("uei") and f.get("uei"):
            u = f["uei"]
            if not re.fullmatch(r"[A-Za-z0-9]{12}", u.replace(" ", "")):
                self.add("Error", loc("uei"), "UEI", "UEI must be exactly 12 letters and numbers (template instruction 7).",
                         value=u, expected="12-character SAM.gov Unique Entity ID", fix="Copy the UEI exactly from the SAM.gov entity registration.", rule="PA-UEI")
            else:
                if u != u.upper() or " " in u:
                    self.add("Warning", loc("uei"), "UEI", "UEI should be uppercase with no spaces.", value=u,
                             fix="Use uppercase letters.", rule="PA-UEI",
                             autofix=self._fix("uei", self._clean("uei", u), "Uppercase the UEI"))
                if re.search(r"[IOio]", u):
                    self.add("Warning", loc("uei"), "UEI", "SAM.gov UEIs do not use the letters I or O. This UEI may contain a typo (1 or 0 intended).",
                             value=u, fix="Verify the UEI against SAM.gov.", rule="PA-UEI")
                if u.startswith("0"):
                    self.add("Warning", loc("uei"), "UEI", "SAM.gov UEIs do not start with zero.", value=u,
                             fix="Verify the UEI against SAM.gov.", rule="PA-UEI")

        # Names
        if f.get("hq_name") and f.get("osc_name") and norm(f["hq_name"]) != norm(f["osc_name"]):
            self.add("Info", loc("osc_name"), "OSC Name", "HQ Organization Name and OSC Name differ. This is valid for a subsidiary or division, but confirm both match SAM.gov.",
                     value=f"HQ: {f['hq_name']} | OSC: {f['osc_name']}", fix="Confirm against the SAM.gov registration and CAGE hierarchy.", rule="PA-NAME")

        # Zip
        if has("zip") and f.get("zip"):
            raw = self.values.get("zip")
            z = f["zip"]
            if isinstance(raw, (int, float)):
                new = z if len(z) >= 5 else z.zfill(5)
                self.add("Warning", loc("zip"), "OSC Zip Code", "Zip code is stored as an Excel number, which drops leading zeros and can import incorrectly.",
                         value=z, expected="Text value such as 02134 or 28374", fix="Store the zip code as text.", rule="PA-ZIP",
                         autofix=self._fix("zip", new, f"Store zip code as text '{new}'", default=len(z) >= 5))
                if len(z) < 5:
                    self.add("Warning", loc("zip"), "OSC Zip Code", "Zip code has fewer than 5 digits. A leading zero was likely lost.",
                             value=z, fix="Confirm the correct 5-digit zip code.", rule="PA-ZIP")
            elif not re.fullmatch(r"\d{5}(-\d{4})?", z) and norm(f.get("country")) in [norm(x) for x in ["United States"] + self.cfg["us_country_aliases"]]:
                self.add("Warning", loc("zip"), "OSC Zip Code", "US zip code is not in 12345 or 12345-6789 format.", value=z,
                         fix="Correct the zip code.", rule="PA-ZIP")

        # Country
        if f.get("country") and norm(f["country"]) in [norm(a) for a in self.cfg["us_country_aliases"]]:
            self.add("Info", loc("country"), "OSC Country", "Country is entered as an abbreviation. The eMASS template example uses 'United States'.",
                     value=f["country"], expected="United States (per the eMASS Example tab)",
                     fix="Confirm the country value eMASS accepts and update if needed.", rule="PA-CTRY")

        # Phone
        if has("phone") and f.get("phone"):
            raw = self.values.get("phone")
            digits = re.sub(r"\D", "", f["phone"])
            new = self._clean("phone", f["phone"])
            if isinstance(raw, (int, float)):
                self.add("Warning", loc("phone"), "OSC Business Phone", "Phone number is stored as an Excel number.",
                         value=f["phone"], expected="Text such as 999-888-7767 (eMASS Example tab format)",
                         fix="Enter the phone number as text.", rule="PA-PHONE",
                         autofix=self._fix("phone", new, f"Store phone as text '{new}'") if new != f["phone"] else None)
            if len(digits) not in (10, 11):
                self.add("Warning", loc("phone"), "OSC Business Phone", "Phone number does not have 10 digits.", value=f["phone"],
                         fix="Verify the phone number.", rule="PA-PHONE")

        # URL
        if has("url") and f.get("url"):
            u = f["url"]
            if not URL_RE.match(u):
                fixed = self._clean("url", u)
                self.add("Error", loc("url"), "OSC Business Web URL",
                         "URL is not a complete, valid URL. The template requires 'http://' or 'https://' (template instruction 8).",
                         value=u, expected="Example: https://www.disa.mil/", fix="Enter the full URL including https://.", rule="PA-URL",
                         autofix=self._fix("url", fixed, f"Correct URL to '{fixed}'") if URL_RE.match(fixed) and fixed != u else None)

        # Sector
        if has("sector") and f.get("sector"):
            parts = split_multi(f["sector"])
            if len(parts) == 1 and "," in parts[0] and parts[0] not in self.rules["sectors"]:
                parts = [p.strip() for p in parts[0].split(",")]
                self.add("Error", loc("sector"), "Sector", "Multiple sectors must be separated by semicolons, not commas (template instruction 9).",
                         value=f["sector"], fix="Separate each sector with a semicolon.", rule="PA-SECT")
            bad = [p for p in parts if p not in self.rules["sectors"]]
            case_only = [p for p in bad if p.lower() in [s.lower() for s in self.rules["sectors"]]]
            hard = [p for p in bad if p not in case_only]
            if hard:
                self.add("Error", loc("sector"), "Sector", "Sector value is not one of the eMASS lookup values.",
                         value=", ".join(hard), expected="One or more of: " + "; ".join(self.rules["sectors"]),
                         fix="Select a value from the dropdown list.", rule="PA-SECT")
            if case_only:
                self.add("Warning", loc("sector"), "Sector", "Sector does not match the lookup value capitalization exactly.",
                         value=", ".join(case_only), fix="Use the exact lookup value.", rule="PA-SECT",
                         autofix=self._fix("sector", self._clean("sector", f["sector"]), "Match sector to exact lookup value"))
            if "Other" in [p for p in parts] or "other" in [p.lower() for p in parts]:
                if not f.get("sector_other"):
                    self.add("Error", loc("sector_other"), "Sector (Other)", "Sector is 'Other', so Sector (Other) is required.",
                             fix="Describe the sector in Sector (Other).", rule="PA-SECT")
            elif f.get("sector_other"):
                self.add("Warning", loc("sector_other"), "Sector (Other)", "Sector (Other) has a value but Sector does not include 'Other'.",
                         value=f["sector_other"], fix="Clear Sector (Other) or add 'Other' to Sector.", rule="PA-SECT")

        # Employees
        if has("employees") and f.get("employees"):
            e = f["employees"].replace(",", "")
            if not re.fullmatch(r"\d+", e) or int(e) <= 0:
                self.add("Error", loc("employees"), "Number of Employees", "Number of Employees must be a whole number greater than zero.",
                         value=f["employees"], fix="Enter a whole number.", rule="PA-EMP")

        # CAGE codes
        hlo = f.get("hlo_cage", "")
        scope_codes = []
        if has("hlo_cage") and hlo:
            if re.search(r"[;,\s]", hlo.strip()):
                self.add("Error", loc("hlo_cage"), "HLO CAGE Code", "Only the single highest-level owner CAGE code is allowed (template instruction 10).",
                         value=hlo, fix="Enter one CAGE code only.", rule="PA-CAGE")
            else:
                self._cage_format(hlo, loc("hlo_cage"), "HLO CAGE Code", "hlo_cage")
        if has("cage_scope") and f.get("cage_scope"):
            raw = f["cage_scope"]
            parts = [p for p in re.split(r"[;,\s]+", raw.strip()) if p]
            scope_codes = [p.upper() for p in parts]
            if len(parts) > 1 and (re.search(r",", raw) or not re.search(r";", raw)):
                self.add("Error", loc("cage_scope"), "CAGE Code(s) in Scope", "Multiple CAGE codes must be separated by semicolons (SAP 5.1 and template instruction 11).",
                         value=raw, expected="Example: 12B45; 45P78", fix="Separate the codes with semicolons.", rule="PA-CAGE",
                         autofix=self._fix("cage_scope", self._clean("cage_scope", raw), "Separate CAGE codes with semicolons"))
            elif raw.strip() != self._clean("cage_scope", raw) and raw.upper() != raw:
                self.add("Warning", loc("cage_scope"), "CAGE Code(s) in Scope", "CAGE codes should be uppercase.",
                         value=raw, fix="Use uppercase.", rule="PA-CAGE",
                         autofix=self._fix("cage_scope", self._clean("cage_scope", raw), "Uppercase CAGE codes"))
            for p in parts:
                self._cage_format(p, loc("cage_scope"), "CAGE Code(s) in Scope", "cage_scope")
            dups = sorted({p for p in scope_codes if scope_codes.count(p) > 1})
            if dups:
                self.add("Warning", loc("cage_scope"), "CAGE Code(s) in Scope", "Duplicate CAGE code listed.", value=", ".join(dups),
                         fix="Remove duplicates.", rule="PA-CAGE")
            if hlo and hlo.upper() not in scope_codes:
                self.add("Info", loc("cage_scope"), "CAGE Code(s) in Scope", "The HLO CAGE code is not listed in scope. Include it only if the HLO is part of this assessment's scope.",
                         value=f"HLO {hlo}; in scope: {raw}", fix="Confirm the HLO is outside the assessment scope.", rule="PA-CAGE")
        self.res.data["pa_cage_codes"] = scope_codes

        # Scope
        sc = f.get("scope", "")
        if has("scope") and sc:
            if sc not in self.cfg["scope_values"]:
                ci = sc.lower() in [v.lower() for v in self.cfg["scope_values"]]
                self.add("Warning" if ci else "Error", loc("scope"), "Scope",
                         "Scope must be exactly 'Enterprise' or 'Enclave'.", value=sc, expected="Enterprise or Enclave",
                         fix="Select from the dropdown.", rule="PA-SCOPE",
                         autofix=self._fix("scope", self._clean("scope", sc), "Match scope to lookup value") if ci else None)
            if sc.lower() == "enclave":
                d = f.get("scope_desc", "")
                if not d:
                    self.add("Error", loc("scope_desc"), "Scope Description", "Scope is 'Enclave', so Scope Description is required (SAP 5.1, template instruction 12).",
                             fix="Enter a meaningful description of the enclave (1000 characters max).", rule="PA-SDESC")
                else:
                    if len(d) > self.cfg["scope_description_max_chars"]:
                        self.add("Error", loc("scope_desc"), "Scope Description", f"Scope Description is {len(d)} characters. The limit is {self.cfg['scope_description_max_chars']}.",
                                 fix="Shorten the description.", rule="PA-SDESC")
                    if len(d) < 25:
                        self.add("Info", loc("scope_desc"), "Scope Description", "Scope Description is very short. The template asks for a meaningful description of the assessment scope.",
                                 value=d, fix="Confirm the description identifies the enclave and what it covers.", rule="PA-SDESC")
            elif sc.lower() == "enterprise" and f.get("scope_desc"):
                self.add("Info", loc("scope_desc"), "Scope Description", "Scope Description is filled in but Scope is 'Enterprise'. It is only required for Enclave.",
                         value=short(f["scope_desc"]), fix="Confirm the scope selection is correct.", rule="PA-SDESC")

        # Dates
        dates = {}
        for k in ("contract_date", "plan_start", "plan_end"):
            if not has(k) or not f.get(k):
                continue
            raw = self.values.get(k)
            d, ok_text = parse_date(raw)
            if d is None:
                self.add("Error", loc(k), cells[k]["label"], "Date could not be read.", value=as_text(raw),
                         expected=self.cfg["date_format_label"], fix="Enter the date as DD-MMM-YYYY, for example 05-Oct-2026.",
                         owner="Ignyte", rule="PA-DATE")
                continue
            dates[k] = d
            if not ok_text:
                already = any(i.autofix and i.autofix.get("field") == k for i in self.res.issues if i.file_key == "pa")
                self.add("Warning", loc(k), cells[k]["label"], "Date is not stored as text in DD-MMM-YYYY format (template instruction 6).",
                         value=as_text(raw) if not isinstance(raw, (_dt.date, _dt.datetime)) else f"Excel date ({raw:%Y-%m-%d})",
                         expected=fmt_date(d), fix=f"Enter as {fmt_date(d)}.", owner="Ignyte", rule="PA-DATE",
                         autofix=None if already else self._fix(k, fmt_date(d), f"Store date as '{fmt_date(d)}'"))
        if "contract_date" in dates and "plan_start" in dates and dates["plan_start"] < dates["contract_date"]:
            self.add("Error", loc("plan_start"), "Assessment Planning Start Date", "Planning start date is before the C3PAO contract date.",
                     value=f"Contract {fmt_date(dates['contract_date'])}; start {fmt_date(dates['plan_start'])}", fix="Correct the dates.", owner="Ignyte", rule="PA-DATE")
        if "plan_start" in dates and "plan_end" in dates and dates["plan_end"] < dates["plan_start"]:
            self.add("Error", loc("plan_end"), "Assessment Planning Completion Date", "Planning completion date is before the planning start date.",
                     value=f"Start {fmt_date(dates['plan_start'])}; end {fmt_date(dates['plan_end'])}", fix="Correct the dates.", owner="Ignyte", rule="PA-DATE")

        # Standard
        st = f.get("standard", "")
        if has("standard") and st:
            if st not in self.cfg["standard_values"]:
                ci = st.lower() in [v.lower() for v in self.cfg["standard_values"]]
                self.add("Error", loc("standard"), "Assessment Standard", "Assessment Standard is not one of the lookup values.", value=st,
                         expected=" or ".join(self.cfg["standard_values"]), fix="Select from the dropdown.", owner="Ignyte", rule="PA-STD",
                         autofix=self._fix("standard", self._clean("standard", st), "Match standard to lookup value") if ci else None)
            elif st != self.cfg["cmmc_l2_expected_standard"]:
                self.add("Warning", loc("standard"), "Assessment Standard",
                         f"Assessment Standard is '{st}'. CMMC Level 2 under 32 CFR 170 is assessed against {self.cfg['cmmc_l2_expected_standard']}.",
                         value=st, fix="Confirm the selection.", owner="Ignyte", rule="PA-STD")

    def _cage_format(self, code, loc, fld, key):
        c = code.strip().upper()
        if not re.fullmatch(r"[A-Z0-9]{5}", c):
            self.add("Error", loc, fld, "CAGE code must be exactly 5 letters and numbers.", value=code,
                     expected="5 characters, such as 1A2B3", fix="Copy the CAGE code from SAM.gov.", rule="PA-CAGE")
            return
        if re.search(r"[IO]", c):
            self.add("Warning", loc, fld, "CAGE codes do not use the letters I or O. This code may contain a typo.", value=code,
                     fix="Verify the code in SAM.gov.", rule="PA-CAGE")
        if not (c[0].isdigit() and c[4].isdigit()):
            self.add("Info", loc, fld, "Code does not follow the US CAGE pattern (numeric first and last character). This is expected only for a foreign NCAGE code.",
                     value=code, fix="Verify the code in SAM.gov.", rule="PA-CAGE")


# --------------------------------------------------------------------------- #
# 2. Asset Inventory Scoping
# --------------------------------------------------------------------------- #
class AssetInventoryChecker:
    def __init__(self, path, rules, res: Result):
        self.path, self.rules, self.res = path, rules, res
        self.cfg = rules["asset_inventory"]

    def add(self, sev, loc, fld, msg, **k):
        self.res.add(sev, "ai", loc, fld, msg, **k)

    def run(self):
        wb = openpyxl.load_workbook(self.path, data_only=True)
        cfg = self.cfg
        m = re.search(r"Version\s*([\d.]+)", " ".join(as_text(c.value) for r in wb.worksheets[0].iter_rows(max_row=40) for c in r))
        self.res.data["ai_version"] = m.group(1) if m else "Not found"

        # --- Main Scoped Assets
        ws = find_sheet(wb, cfg["sheet_main"])
        assets = []
        if ws is None:
            self.add("Error", "workbook", "Main Scoped Assets", "The 'Main Scoped Assets' tab is missing.",
                     fix="Use the Ignyte CMMC L2 Asset Inventory Scoping template.", rule="AI-TAB")
        else:
            hr = cfg["main_header_row"]
            cols = header_map(ws, hr, {"cat": "Asset Category", "asset": ["Asset"], "qty": "Asset Quantity",
                                       "type": "Asset Type", "desc": "Asset Description", "host": "Hosting Location",
                                       "comment": "Comment"})
            # 'Asset' keyword also matches other headers; force column B if present
            if norm(ws.cell(hr, 2).value) == "asset":
                cols["asset"] = 2
            need = ["cat", "asset", "qty", "type", "desc", "host"]
            if any(k not in cols for k in need):
                self.add("Error", ws.title, "Layout", "Expected column headers were not found on row 3.",
                         fix="Do not move or rename the template columns.", rule="AI-HDR")
            else:
                names = defaultdict(list)
                for r in range(hr + 1, ws.max_row + 1):
                    v = row_values(ws, r, cols)
                    if all(is_blank(x) for x in v.values()):
                        continue
                    a = {k: as_text(x).strip() for k, x in v.items()}
                    a["row"] = r
                    assets.append(a)
                    L = openpyxl.utils.get_column_letter
                    labels = {"cat": "Asset Category", "asset": "Asset", "qty": "Asset Quantity", "type": "Asset Type",
                              "desc": "Asset Description", "host": "Hosting Location"}
                    missing = [labels[k] for k in need if not a.get(k)]
                    if missing:
                        self.add("Error", f"{ws.title}!row {r}", a.get("asset") or "Asset row",
                                 "Required field(s) blank: " + ", ".join(missing) + ".",
                                 fix="Complete all required columns A through F (SAP section 4).", rule="AI-REQ")
                    if a.get("cat") and a["cat"] not in cfg["asset_categories"]:
                        self.add("Error", f"{ws.title}!{L(cols['cat'])}{r}", "Asset Category", "Asset Category is not one of the five scoping categories.",
                                 value=a["cat"], expected="; ".join(cfg["asset_categories"]), fix="Select from the dropdown.", rule="AI-CAT")
                    if a.get("type") and a["type"] not in cfg["asset_types"]:
                        self.add("Error", f"{ws.title}!{L(cols['type'])}{r}", "Asset Type", "Asset Type is not one of the dropdown values.",
                                 value=a["type"], expected="; ".join(cfg["asset_types"]), fix="Select from the dropdown.", rule="AI-TYPE")
                    if a.get("qty"):
                        q = a["qty"].replace(",", "")
                        if not re.fullmatch(r"\d+(\.0+)?", q) or float(q) <= 0:
                            self.add("Error", f"{ws.title}!{L(cols['qty'])}{r}", "Asset Quantity", "Asset Quantity must be a whole number greater than zero.",
                                     value=a["qty"], fix="Enter a numeric quantity.", rule="AI-QTY")
                    host = a.get("host", "")
                    if "fedramp" in host.lower() and not re.search(cfg["fedramp_id_regex"], host):
                        self.add("Warning", f"{ws.title}!{L(cols['host'])}{r}", a.get("asset") or "Hosting Location",
                                 "Hosting location says FedRAMP but no FedRAMP package ID is listed. SAP section 4 asks for the FedRAMP ID if applicable.",
                                 value=short(host), expected="Example: FedRAMP Moderate SaaS F1512167750", fix="Add the FedRAMP Marketplace package ID.", rule="AI-FR")
                    for mk in cfg["example_asset_markers"]:
                        if mk.lower() in " ".join(str(x) for x in a.values()).lower():
                            self.add("Warning", f"{ws.title}!row {r}", a.get("asset") or "Asset row",
                                     "Row appears to be copied from the Main Scope Assets_Example tab.", value=mk,
                                     fix="Replace example content with the OSC's actual asset.", rule="AI-EX")
                            break
                    if a.get("asset"):
                        names[norm(a["asset"])].append(r)
                for n, rows in names.items():
                    if len(rows) > 1:
                        self.add("Info", f"{ws.title}!rows {', '.join(map(str, rows))}", "Asset", "Same asset name listed more than once.",
                                 value=n, fix="Confirm these are distinct entries or combine them with a total quantity.", rule="AI-DUP")
                if not assets:
                    self.add("Error", ws.title, "Main Scoped Assets", "No assets are listed.",
                             fix="List every asset mapped to one of the five categories (SAP section 4 and 5.4).", rule="AI-EMPTY")
                else:
                    cats = {a.get("cat") for a in assets}
                    if "Controlled Unclassified Information" not in cats:
                        self.add("Warning", ws.title, "Asset Category", "No assets are categorized as Controlled Unclassified Information. A Level 2 scope is built around CUI assets.",
                                 fix="Confirm the categorization of assets that process, store, or transmit CUI.", rule="AI-CUI")
                    if "Security Protection" not in cats:
                        self.add("Warning", ws.title, "Asset Category", "No Security Protection Assets are listed (for example identity, MFA, logging, endpoint protection, firewalls).",
                                 fix="Confirm all assets that provide security functions are listed.", rule="AI-SPA")
        self.res.data["assets"] = assets

        # --- ESPs
        ws = find_sheet(wb, cfg["sheet_esp"])
        esps = []
        if ws is None:
            self.add("Error", "workbook", "Consolidated ESP", "The 'Consolidated External Service Providers' tab is missing.",
                     fix="Use the Ignyte CMMC L2 Asset Inventory Scoping template.", rule="AI-TAB")
        else:
            hr = cfg["esp_header_row"]
            cols = header_map(ws, hr, {"name": "External Service Provider Name", "last": "POC Last Name", "first": "POC First Name",
                                       "phone": "POC Phone", "email": "POC Email", "desc": "Description",
                                       "status": "CMMC Status", "sector": "Sector (Required", "other": "Sector (Other)"})
            if "sector" not in cols:
                cols.update(header_map(ws, hr, {"sector": "External Service Provider Sector"}))
                if cols.get("sector") == cols.get("other"):
                    cols["sector"] = 8
            need = ["name", "last", "first", "phone", "email", "desc", "status", "sector"]
            labels = {"name": "ESP Name", "last": "POC Last Name", "first": "POC First Name", "phone": "POC Phone",
                      "email": "POC Email", "desc": "Description", "status": "CMMC Status", "sector": "Sector", "other": "Sector (Other)"}
            if any(k not in cols for k in need):
                self.add("Error", ws.title, "Layout", "Expected ESP column headers were not found on row 5.",
                         fix="Do not move or rename the template columns.", rule="AI-HDR")
            else:
                L = openpyxl.utils.get_column_letter
                for r in range(hr + 1, ws.max_row + 1):
                    v = row_values(ws, r, cols)
                    rowtxt = " ".join(as_text(ws.cell(r, c).value) for c in range(1, 12))
                    if all(is_blank(x) for x in v.values()):
                        continue
                    e = {k: as_text(x).strip() for k, x in v.items()}
                    e["row"] = r
                    if any(mk.lower() in rowtxt.lower() for mk in cfg["esp_sample_markers"]):
                        self.add("Error", f"{ws.title}!row {r}", e.get("name") or "ESP row",
                                 "The template sample record is still present. SAP section 4 says to remove the example once your ESPs are added.",
                                 value=short(rowtxt, 120), fix="Delete the sample record.", rule="AI-ESPSAMPLE")
                        e["sample"] = True
                        continue
                    esps.append(e)
                    missing = [labels[k] for k in need if not e.get(k)]
                    if missing:
                        self.add("Error", f"{ws.title}!row {r}", e.get("name") or "ESP row", "Required ESP field(s) blank: " + ", ".join(missing) + ".",
                                 fix="All columns except I are required (SAP section 4). eMASS requires these when an ESP is used.", rule="AI-ESPREQ")
                    if e.get("email") and not EMAIL_RE.match(e["email"]):
                        self.add("Error", f"{ws.title}!{L(cols['email'])}{r}", "POC Email", "Email address is not valid.", value=e["email"],
                                 fix="Correct the email address.", rule="AI-ESPEMAIL")
                    if e.get("phone") and len(re.sub(r"\D", "", e["phone"])) not in (10, 11):
                        self.add("Warning", f"{ws.title}!{L(cols['phone'])}{r}", "POC Phone", "Phone number does not have 10 digits.", value=e["phone"],
                                 fix="Verify the phone number.", rule="AI-ESPPHONE")
                    if e.get("status") and e["status"] not in self.rules["esp_cmmc_status_values"]:
                        self.add("Error", f"{ws.title}!{L(cols['status'])}{r}", "CMMC Status", "CMMC Status must be None, Level 2, or Level 3 (eMASS instruction 23).",
                                 value=e["status"], fix="Select from the dropdown.", rule="AI-ESPSTAT")
                    if e.get("sector"):
                        bad = [p for p in split_multi(e["sector"]) if p not in self.rules["sectors"]]
                        if bad:
                            self.add("Error", f"{ws.title}!{L(cols['sector'])}{r}", "ESP Sector", "Sector is not one of the eMASS lookup values.",
                                     value=", ".join(bad), fix="Select from the dropdown.", rule="AI-ESPSECT")
                        if "Other" in split_multi(e["sector"]) and not e.get("other"):
                            self.add("Error", f"{ws.title}!row {r}", "ESP Sector (Other)", "Sector is 'Other', so Sector (Other) is required.",
                                     fix="Describe the sector.", rule="AI-ESPSECT")
                names = defaultdict(list)
                for e in esps:
                    names[norm(e.get("name"))].append(e["row"])
                for n, rows in names.items():
                    if n and len(rows) > 1:
                        self.add("Warning", f"{ws.title}!rows {', '.join(map(str, rows))}", "ESP Name", "Same ESP listed more than once.",
                                 value=n, fix="Combine duplicates. eMASS expects one record per ESP.", rule="AI-ESPDUP")
        self.res.data["esps"] = esps

        # --- Out of scope areas
        ws = find_sheet(wb, cfg["sheet_oos"])
        oos = []
        if ws is None:
            self.add("Error", "workbook", "Out of Scope Areas", "The 'Out of Scope Areas of Ops and BUs' tab is missing.",
                     fix="Use the Ignyte CMMC L2 Asset Inventory Scoping template.", rule="AI-TAB")
        else:
            hr = cfg["oos_header_row"]
            for r in range(hr + 1, ws.max_row + 1):
                a, b, c = (as_text(ws.cell(r, i).value).strip() for i in (1, 2, 3))
                if not (a or b or c):
                    continue
                oos.append({"row": r, "name": a, "desc": b, "notes": c})
                if not a or not b:
                    self.add("Error", f"{ws.title}!row {r}", a or "Out of scope row", "Area name and description are both required.",
                             fix="Enter the area or BU name and a short description (SAP section 4).", rule="AI-OOS")
            if not oos:
                self.add("Warning", ws.title, "Out of Scope Areas",
                         "No out-of-scope areas of operation or business units are listed. SAP section 4 asks the OSC to list them and affirm it.",
                         fix="List each out-of-scope area or BU, or document in writing that none exist.", rule="AI-OOSEMPTY")
        self.res.data["oos"] = oos


# --------------------------------------------------------------------------- #
# 3. Document Request List
# --------------------------------------------------------------------------- #
class DRLChecker:
    def __init__(self, path, rules, res: Result):
        self.path, self.rules, self.res = path, rules, res
        self.cfg = rules["drl"]

    def add(self, sev, loc, fld, msg, **k):
        self.res.add(sev, "drl", loc, fld, msg, **k)

    def run(self):
        wb = openpyxl.load_workbook(self.path, data_only=True)
        missing = [s for s in self.cfg["required_sheets"] if find_sheet(wb, s) is None]
        for s in missing:
            self.add("Error", "workbook", s, f"The '{s}' tab is missing from the Document Request List.",
                     fix="Use the full Ignyte CMMC L2 Assessment Document Request List template (SAP section 4).", rule="DRL-TAB")
        m = re.search(r"Version\s*([\d.]+)", " ".join(as_text(c.value) for c in (find_sheet(wb, "Instructions") or wb.worksheets[0])["C"][:40]))
        self.res.data["drl_version"] = m.group(1) if m else "Not found"
        self.team = self._team(wb)
        self._evidence_plan(wb)
        self._ownership(wb)
        self._facilities(wb)
        self._physical_security(wb)

    # Evidence Plan
    def _evidence_plan(self, wb):
        ws = find_sheet(wb, "Evidence Plan")
        if ws is None:
            return
        cols = header_map(ws, 1, {"obj": "OBJECTIVE", "type": "EVIDENCE TYPE", "erl": "ERL", "map": "DOCUMENT MAPPING"})
        if any(k not in cols for k in ("obj", "type", "erl", "map")):
            self.add("Error", ws.title, "Layout", "Evidence Plan headers (Objective, Evidence Type, ERL #, Document Mapping) were not found.",
                     fix="Do not rename the template columns.", rule="DRL-HDR")
            return
        L = openpyxl.utils.get_column_letter
        objs, reqs = [], 0
        for r in range(2, ws.max_row + 1):
            o = as_text(ws.cell(r, cols["obj"]).value).strip()
            if not o:
                continue
            if "[" in o:
                objs.append((r, o, as_text(ws.cell(r, cols["type"]).value).strip(),
                             as_text(ws.cell(r, cols["erl"]).value).strip(), as_text(ws.cell(r, cols["map"]).value).strip()))
            elif re.fullmatch(r"3\.\d+\.\d+", o):
                reqs += 1
        if len(objs) != self.cfg["expected_objectives"] or reqs != self.cfg["expected_requirements"]:
            self.add("Warning", ws.title, "Template integrity",
                     f"Evidence Plan lists {reqs} requirements and {len(objs)} objectives. Expected {self.cfg['expected_requirements']} and {self.cfg['expected_objectives']}.",
                     fix="Confirm no rows were deleted from the Evidence Plan.", rule="DRL-COUNT")
        # Template ERL integrity (Ignyte)
        erls = [x[3] for x in objs if x[3]]
        dups = sorted({e for e in erls if erls.count(e) > 1})
        bad = sorted({e for e in erls if not re.fullmatch(r"E-[A-Z]{2}-\d{2,3}", e)})
        if dups:
            self.add("Info", ws.title, "ERL #", "Duplicate ERL numbers in the template. File-name prefixes for these cannot be matched to one objective.",
                     value=", ".join(dups), fix="Ignyte to correct the DRL template.", owner="Ignyte", rule="DRL-ERLDUP")
        if bad:
            self.add("Info", ws.title, "ERL #", "Malformed ERL number in the template.", value=", ".join(bad),
                     fix="Ignyte to correct the DRL template.", owner="Ignyte", rule="DRL-ERLBAD")
        # Mapping
        mtypes = [t.lower() for t in self.cfg["mapping_evidence_types"]]
        unmapped = defaultdict(list)
        mapped_count, need_count = 0, 0
        for r, o, t, erl, mp in objs:
            if not any(mt in t.lower() for mt in mtypes):
                continue
            need_count += 1
            if not mp:
                unmapped[family_of(o)].append(o)
                continue
            mapped_count += 1
            if erl and erl not in dups and erl not in bad:
                files = [x.strip() for x in re.split(r"[\n;|]+", mp) if x.strip()]
                wrong = [x for x in files if not x.upper().startswith(erl.upper())]
                if wrong:
                    self.add("Warning", f"{ws.title}!{L(cols['map'])}{r}", o, f"File name in Document Mapping does not start with the ERL # ({erl}). SAP section 8 requires the ERL # as the file name prefix.",
                             value=short("; ".join(wrong)), expected=f"{erl} <file name>", fix=f"Rename the file with the prefix {erl} and update column {L(cols['map'])}.", rule="DRL-PREFIX")
        self.res.data["drl_mapping"] = (mapped_count, need_count)
        if need_count and mapped_count == 0:
            self.add("Warning", f"{ws.title}!{L(cols['map'])}", "Document Mapping",
                     f"No Document or Artifact objectives are mapped to files (0 of {need_count}). SAP section 4 notes unmapped documentation increases auditor time and may delay Phase 2.",
                     fix="Add the uploaded file names (ERL # prefix) in the Document Mapping column.", rule="DRL-MAP")
        else:
            for fam, ids in sorted(unmapped.items()):
                self.add("Warning", f"{ws.title}!{L(cols['map'])}", f"{fam} Document Mapping",
                         f"{len(ids)} Document/Artifact objective(s) in {fam} have no file mapped.",
                         value=compress_ids(ids), fix="Add the uploaded file names (ERL # prefix) in the Document Mapping column.", rule="DRL-MAP")

    # Control ownership
    def _ownership(self, wb):
        ws = find_sheet(wb, "171 R2 Control Ownership")
        if ws is None:
            return
        cols = header_map(ws, 1, {"id": "CMMC v2.0 L2", "owner": "Control Owner", "op": "Control Operator"})
        if any(k not in cols for k in ("id", "owner", "op")):
            self.add("Error", ws.title, "Layout", "Control Owner (G) and Control Operator (H) headers were not found.",
                     fix="Do not rename the template columns.", rule="DRL-HDR")
            return
        miss_o, miss_p, total = defaultdict(list), defaultdict(list), 0
        people = set()
        for r in range(2, ws.max_row + 1):
            oid = as_text(ws.cell(r, cols["id"]).value).strip()
            if not oid:
                continue
            total += 1
            ow = as_text(ws.cell(r, cols["owner"]).value).strip()
            op = as_text(ws.cell(r, cols["op"]).value).strip()
            if not ow:
                miss_o[family_of(oid)].append(oid)
            if not op:
                miss_p[family_of(oid)].append(oid)
            for p in (ow, op):
                for n in re.split(r"[;,/\n]+| and ", p):
                    if n.strip():
                        people.add(n.strip())
        nmiss = sum(len(v) for v in miss_o.values()) + sum(len(v) for v in miss_p.values())
        if total and nmiss == total * 2:
            self.add("Error", ws.title, "Control Owner / Operator",
                     f"Control Owner (G) and Control Operator (H) are blank for all {total} objectives. SAP section 4 says completing columns G and H is required.",
                     fix="Enter the control owner and operator for each objective.", rule="DRL-OWN")
        else:
            for fam in sorted(set(miss_o) | set(miss_p)):
                ids_o, ids_p = miss_o.get(fam, []), miss_p.get(fam, [])
                parts = []
                if ids_o:
                    parts.append(f"Owner blank: {compress_ids(ids_o, 25)}")
                if ids_p:
                    parts.append(f"Operator blank: {compress_ids(ids_p, 25)}")
                self.add("Error", f"{ws.title}", f"{fam} Owner / Operator",
                         f"{len(ids_o)} objective(s) missing Control Owner and {len(ids_p)} missing Control Operator in {fam}.",
                         value=" | ".join(parts), fix="Complete columns G and H (required per SAP section 4).", rule="DRL-OWN")
        # People named who are not on the OSC Assessment Team
        if self.team and people:
            team_names = [norm(t["name"]) for t in self.team]
            team_tokens = set(" ".join(team_names).split())
            unknown = sorted(p for p in people if norm(p) not in team_names and not (set(norm(p).split()) & team_tokens))
            if unknown:
                self.add("Info", ws.title, "Owner / Operator names", "Names used as control owner or operator do not appear on the OSC Assessment Team tab.",
                         value=compress_ids(unknown, 15), fix="Add these people to the OSC Assessment Team tab, or confirm they are roles rather than names.", rule="DRL-OWNTEAM")

    # Facilities
    def _facilities(self, wb):
        ws = find_sheet(wb, "Physical Facilities")
        facs = []
        if ws is None:
            self.res.data["facilities"] = facs
            return
        cols = header_map(ws, 1, {"name": "Location Name", "addr": "Address", "desc": "Description", "own": "Rented",
                                  "cage": "CAGE", "cleared": "Cleared", "onsite": "Onsite"})
        L = openpyxl.utils.get_column_letter
        for r in range(2, ws.max_row + 1):
            v = {k: as_text(ws.cell(r, c).value).strip() for k, c in cols.items()}
            if not any(v.values()):
                continue
            v["row"] = r
            facs.append(v)
            if "example" in " ".join(str(x) for x in v.values()).lower():
                self.add("Warning", f"{ws.title}!row {r}", v.get("name") or "Facility", "Row appears to be a template example record.",
                         fix="Remove example records once your facilities are added (SAP section 5.6).", rule="DRL-FACEX")
            req = {"name": "Location Name", "addr": "Address", "desc": "Description", "own": "Ownership status", "cleared": "Cleared facility (Yes/No)"}
            miss = [lab for k, lab in req.items() if k in cols and not v.get(k)]
            if miss:
                self.add("Error", f"{ws.title}!row {r}", v.get("name") or "Facility", "Required facility field(s) blank: " + ", ".join(miss) + ".",
                         fix="Complete columns A through F for each in-scope facility (SAP section 5.6).", rule="DRL-FACREQ")
            if v.get("own"):
                first = re.split(r"[\s,:;(-]+", v["own"].strip())[0].capitalize()
                if first not in self.cfg["facility_ownership_values"]:
                    self.add("Error", f"{ws.title}!{L(cols['own'])}{r}", "Ownership status", "Ownership must be Rented, Leased, Own, or Other.",
                             value=v["own"], fix="Enter one of the listed values.", rule="DRL-FACOWN")
                elif first == "Other" and len(v["own"].strip()) <= 6:
                    self.add("Error", f"{ws.title}!{L(cols['own'])}{r}", "Ownership status", "'Other' requires a description.",
                             value=v["own"], fix="Describe the arrangement, for example 'Other - shared space provided by parent company'.", rule="DRL-FACOWN")
            for k in ("cleared", "onsite"):
                if k in cols and v.get(k) and v[k].capitalize() not in self.cfg["yes_no"]:
                    self.add("Error", f"{ws.title}!{L(cols[k])}{r}", "Yes/No field", "Value must be Yes or No.", value=v[k],
                             fix="Enter Yes or No.", owner="OSC" if k == "cleared" else "Ignyte", rule="DRL-FACYN")
            if "onsite" in cols and not v.get("onsite"):
                self.add("Info", f"{ws.title}!{L(cols['onsite'])}{r}", "Selected for Onsite", "Onsite selection is blank. Per SAP section 4, this value is determined from the Physical Security tab.",
                         fix="Ignyte to determine onsite selection.", owner="Ignyte", rule="DRL-FACONSITE")
            if v.get("cage"):
                for code in re.split(r"[;,\s]+", v["cage"]):
                    if code and not re.fullmatch(r"[A-Za-z0-9]{5}", code) and code.lower() not in ("n/a", "na", "none"):
                        self.add("Error", f"{ws.title}!{L(cols['cage'])}{r}", "CAGE Code", "CAGE code must be 5 letters and numbers.", value=code,
                                 fix="Correct the CAGE code.", rule="DRL-FACCAGE")
        if not facs:
            self.add("Error", ws.title, "Physical Facilities", "No in-scope facilities are listed. SAP section 5.6 requires at least the primary facility housing in-scope assets and personnel.",
                     fix="Add each in-scope facility.", rule="DRL-FACEMPTY")
        self.res.data["facilities"] = facs

    def _physical_security(self, wb):
        ws = find_sheet(wb, "Physical Security")
        if ws is None:
            return
        marker = self.cfg["unanswered_physical_security_marker"].lower()
        open_rows = []
        for r in range(2, ws.max_row + 1):
            oid = as_text(ws.cell(r, 2).value).strip()
            e = as_text(ws.cell(r, 5).value).strip()
            if oid and (not e or e.lower().startswith(marker)):
                open_rows.append(oid)
        if open_rows:
            self.add("Info", f"{ws.title}!E", "Physical Security selections",
                     f"{len(open_rows)} of the Physical Security rows still show the '[SELECT FROM' options. The DRL Instructions tab asks the OSC to select options in column E, while SAP section 5.6 says no OSC action is required on this tab.",
                     value=compress_ids(open_rows, 20), fix="Confirm who completes this tab for this engagement before onsite determination.", owner="Ignyte", rule="DRL-PHYSEC")

    # Team
    def _team(self, wb):
        ws = find_sheet(wb, "OSC Assessment Team")
        team = []
        if ws is None:
            self.res.data["team"] = team
            return team
        cols = header_map(ws, 1, {"name": "Name", "email": "Email", "title": "Title", "resp": "Responsibility"})
        L = openpyxl.utils.get_column_letter
        emails = defaultdict(list)
        for r in range(2, ws.max_row + 1):
            v = {k: as_text(ws.cell(r, c).value).strip() for k, c in cols.items()}
            if not any(v.values()):
                continue
            v["row"] = r
            team.append(v)
            miss = [lab for k, lab in (("name", "Name"), ("email", "Email"), ("title", "Title"), ("resp", "Responsibility")) if not v.get(k)]
            if miss:
                self.add("Error", f"{ws.title}!row {r}", v.get("name") or "Team member", "Required field(s) blank: " + ", ".join(miss) + ".",
                         fix="Provide Name, Email, Title, and Responsibility (SAP section 6).", rule="DRL-TEAMREQ")
            if v.get("name") and len(v["name"].split()) < 2:
                self.add("Warning", f"{ws.title}!{L(cols['name'])}{r}", "Name", "Name should include first and last name (SAP section 6).",
                         value=v["name"], fix="Enter first and last name.", rule="DRL-TEAMNAME")
            if v.get("email"):
                if not EMAIL_RE.match(v["email"]):
                    self.add("Error", f"{ws.title}!{L(cols['email'])}{r}", "Email", "Email address is not valid.", value=v["email"],
                             fix="Correct the email address.", rule="DRL-TEAMEMAIL")
                emails[v["email"].lower()].append(r)
        for e, rows in emails.items():
            if len(rows) > 1:
                self.add("Warning", f"{ws.title}!rows {', '.join(map(str, rows))}", "Email", "Same email listed more than once.", value=e,
                         fix="Remove duplicates.", rule="DRL-TEAMDUP")
        if not team:
            self.add("Error", ws.title, "OSC Assessment Team", "No OSC key personnel are listed. SAP section 6 requires them.",
                     fix="Add each key person with Name, Email, Title, and Responsibility.", rule="DRL-TEAMEMPTY")
        else:
            self.add("Info", ws.title, "eMASS contacts",
                     "eMASS requires an OSC Assessment Official (signature authority) and an OSC Technical POC, each with last name, first name, title, email, and phone. This tab has no phone column.",
                     value=f"{len(team)} team member(s) listed", fix="Confirm which listed person is the Assessment Official and which is the Technical POC, and collect their phone numbers.",
                     owner="Ignyte", rule="DRL-EMASSPOC")
        self.res.data["team"] = team
        return team


# --------------------------------------------------------------------------- #
# 4. Security Assessment Plan
# --------------------------------------------------------------------------- #
class SAPChecker:
    def __init__(self, path, rules, res: Result):
        self.path, self.rules, self.res = path, rules, res
        self.cfg = rules["sap"]

    def add(self, sev, loc, fld, msg, **k):
        self.res.add(sev, "sap", loc, fld, msg, **k)

    def run(self):
        if docx is None:
            self.add("Error", "file", "python-docx", "python-docx is not installed, so the SAP cannot be read.",
                     fix="Run Run_PreAssessment_Checker.bat to install requirements.", owner="Ignyte", rule="SAP-LIB")
            return
        d = docx.Document(self.path)
        paras = [p.text for p in d.paragraphs]
        covered = set()
        sec = "Front matter"
        sections = []
        for i, p in enumerate(d.paragraphs):
            if p.style is not None and p.style.name.lower().startswith("heading") and p.text.strip():
                sec = p.text.strip()
            sections.append(sec)

        def ploc(i):
            return f"Paragraph {i + 1} ({short(sections[i], 50)})"

        # OSC name from cover
        osc = ""
        for i, t in enumerate(paras):
            if t.strip().lower().startswith("prepared for"):
                for t2 in paras[i + 1:i + 5]:
                    if t2.strip():
                        osc = t2.strip()
                        break
                break
        if not osc:
            osc = next((t.strip() for t in paras if t.strip()), "")
        self.res.data["sap_osc_name"] = osc
        dv = next((re.search(r"Document Version:\s*([\d.]+)", t).group(1) for t in paras if re.search(r"Document Version:\s*([\d.]+)", t)), "")
        self.res.data["sap_version"] = dv or "Not found"

        # Section 3 fields
        def after(prefix):
            for i, t in enumerate(paras):
                ts = t.strip()
                if ts.lower().startswith(prefix.lower()):
                    rest = ts[len(prefix):]
                    if rest[:1] in (":", "?") or rest == "":
                        return i, rest.strip(" :?\t")
            return None, None

        answers = {}
        i, v = after("Score submitted as")
        if i is not None:
            covered.add(i)
            answers["score_type"] = v
            if not v or "<" in v or "Select One" in v:
                self.add("Error", ploc(i), "Score submitted as", "SPRS score submission type is not selected.", value=short(v),
                         expected="Enterprise, Enclave, or Contract", fix="Select one value (SAP section 3).", rule="SAP-SPRS")
            elif v.split()[0].strip(".,").capitalize() not in self.cfg["score_types"]:
                self.add("Warning", ploc(i), "Score submitted as", "Value is not Enterprise, Enclave, or Contract.", value=v, fix="Select one listed value.", rule="SAP-SPRS")
        i, v = after("Current SPRS Score")
        if i is not None:
            covered.add(i)
            answers["sprs"] = v
            m = re.search(r"-?\d+", v or "")
            if not v or not m or "__" in v:
                self.add("Error", ploc(i), "Current SPRS Score", "Current SPRS score is blank.", value=short(v),
                         fix="Enter the current SPRS score (SAP section 3).", rule="SAP-SPRS")
            elif not (self.cfg["sprs_min"] <= int(m.group()) <= self.cfg["sprs_max"]):
                self.add("Error", ploc(i), "Current SPRS Score", f"SPRS score must be between {self.cfg['sprs_min']} and {self.cfg['sprs_max']}.",
                         value=v, fix="Verify the SPRS score.", rule="SAP-SPRS")
        i, v = after("Has your organization been previously assessed by DIBCAC")
        dib = None
        if i is not None:
            covered.add(i)
            dib = self._yn(v)
            answers["dibcac"] = v
            if dib is None:
                self.add("Error", ploc(i), "Prior DIBCAC assessment", "Yes or No answer is missing.", value=short(v),
                         fix="Answer Yes or No (SAP section 3).", rule="SAP-DIB")
        dib_items = [("Date assessed by DIBCAC", "date"), ("Score post DIBCAC assessment", "score"),
                     ("Have all POA&Ms been addressed since the DIBCAC assessment", "yn"),
                     ("Has the scope changed within your organizational boundary since the last DIBCAC Assessment", "yn"),
                     ("Please describe change in scope since last DIBCAC assessment (if yes)", "text")]
        scope_changed = None
        for prefix, kind in dib_items:
            i, v = after(prefix)
            if i is None:
                continue
            covered.add(i)
            blank = (not v) or "<" in v or "__" in v
            if dib is True:
                if kind == "text":
                    if scope_changed and blank:
                        self.add("Error", ploc(i), "DIBCAC scope change", "Scope change since DIBCAC is 'Yes' but no description is given.",
                                 fix="Describe the scope change.", rule="SAP-DIB")
                    continue
                if blank or (kind == "yn" and self._yn(v) is None):
                    self.add("Error", ploc(i), short(prefix, 50), "Prior DIBCAC assessment is 'Yes', so this item must be answered.",
                             value=short(v), fix="Complete the DIBCAC follow-up items (SAP section 3).", rule="SAP-DIB")
                if "scope changed" in prefix:
                    scope_changed = self._yn(v)
                if "POA&Ms" in prefix and self._yn(v) is False:
                    self.add("Warning", ploc(i), "DIBCAC POA&Ms", "OSC indicates not all DIBCAC POA&Ms have been addressed.",
                             value=v, fix="Assessor awareness item for planning.", owner="Ignyte", rule="SAP-DIB")
            elif dib is False and not blank and kind != "text":
                self.add("Info", ploc(i), short(prefix, 50), "Prior DIBCAC assessment is 'No' but this follow-up item has a value.",
                         value=short(v), fix="Confirm the DIBCAC answer.", rule="SAP-DIB")
        self.res.data["sap_answers"] = answers

        # Yes/No items
        yn_results = {}
        for item in self.cfg["yes_no_items"]:
            idx = next((j for j, t in enumerate(paras) if item["key"].lower() in t.lower()), None)
            if idx is None:
                self.add("Warning", "Document", item["label"], "Affirmation item was not found. The SAP text may have been altered.",
                         fix="Use the unmodified Ignyte SAP template.", owner="Ignyte", rule="SAP-ITEM")
                continue
            covered.add(idx)
            t = paras[idx]
            pre = t.split("\t", 1)[0] if "\t" in t else t[: t.lower().find(item["key"].lower())]
            ans = self._yn(pre)
            yn_results[item["id"]] = ans
            item["_idx"] = idx
            if ans is None:
                self.add("Error", ploc(idx), item["label"], "Yes or No answer is missing.", value=short(pre.strip() or "(blank)", 40),
                         expected="Yes or No", fix=f"Replace '<Yes or No>' with Yes or No (section {item['id'].split('.')[0]}).", rule="SAP-YN")
            elif ans is False and item.get("group") != "coi":
                self.add("Warning", ploc(idx), item["label"], "OSC answered 'No'. This affirmation is part of Phase 1 readiness before Phase 2 starts.",
                         value="No", fix="Resolve the item with the OSC before the assessment start, or document the reason.", rule="SAP-NO")
            # Inline placeholders in the same paragraph (e.g. [insert SSP name])
            for pat in (r"\[insert[^\]]*\]",):
                for mm in re.finditer(pat, t, flags=re.I):
                    self.add("Error", ploc(idx), item["label"], "Placeholder text was not replaced.", value=mm.group(0),
                             fix="Enter the requested detail.", rule="SAP-PH")
        self.res.data["sap_yn"] = {it["label"]: yn_results.get(it["id"]) for it in self.cfg["yes_no_items"]}
        coi = [yn_results.get(it["id"]) for it in self.cfg["yes_no_items"] if it.get("group") == "coi"]
        if coi and all(a is not None for a in coi) and sum(1 for a in coi if a) != 1:
            self.add("Error", "Section 9", "Conflict of Interest", "Exactly one COI statement should be 'Yes' (the SAP says choose one based on the actual status).",
                     value=", ".join("Yes" if a else "No" for a in coi), fix="Mark one COI statement Yes.", owner="Ignyte", rule="SAP-COI")
        self.yn = yn_results

        # Signature blocks
        for who, owner in (("On behalf of the Organization Seeking Certification", "OSC"), ("On behalf of Ignyte", "Ignyte")):
            j = next((k for k, t in enumerate(paras) if t.strip().lower().startswith(who.lower())), None)
            if j is None:
                continue
            for k in range(j + 1, min(j + 8, len(paras))):
                t = paras[k].strip()
                if t.lower().startswith("on behalf of"):
                    break
                m = re.match(r"(Name|Title|Date):\s*(.*)$", t)
                if not m:
                    continue
                covered.add(k)
                fld, val = m.group(1), m.group(2).strip()
                if not val or val in ("Name", "Position", "Name/Affirming Official") or re.fullmatch(r"_+", val):
                    self.add("Error" if owner == "OSC" else "Warning", f"Paragraph {k + 1} (Plan Acceptance signature block)", f"Plan Acceptance {fld} ({owner})",
                             f"Plan Acceptance {fld.lower()} is not completed for the {owner} signer.", value=short(val or "(blank)"),
                             fix="Complete and sign the Plan Acceptance block (SAP section 15).", owner=owner, rule="SAP-SIG")
                elif owner == "OSC" and fld == "Name":
                    self.res.data["sap_osc_signer"] = val

        # Version references in acceptance text (template consistency)
        for k, t in enumerate(paras):
            if "By signing below" in t:
                m = re.search(r"version\s*([\d.]+)", t, flags=re.I)
                if m and dv and m.group(1) != dv:
                    self.add("Info", ploc(k), "Plan Acceptance text", f"Acceptance text cites version {m.group(1)}, but the document version is {dv}.",
                             value=short(t, 120), fix="Ignyte to align the version and date in the SAP template.", owner="Ignyte", rule="SAP-VER")

        # Generic placeholders
        pats = [re.compile(p, re.I) for p in self.cfg["placeholder_patterns"]]
        for k, t in enumerate(paras):
            if k in covered or not t.strip():
                continue
            found = []
            for p in pats:
                found += [m.group(0) for m in p.finditer(t)]
            if found:
                self.add("Error" if "System" in sections[k] or "Data Flow" in sections[k] or "Phase 1" in sections[k] else "Warning",
                         ploc(k), short(sections[k], 40), "Placeholder text was not replaced.",
                         value=", ".join(sorted(set(found)))[:150], expected=short(t, 140),
                         fix="Replace the placeholder with the engagement-specific detail.", rule="SAP-PH")
        # Tables (signature or extra fill-ins inside tables)
        for ti, tb in enumerate(d.tables):
            seen = set()
            for row in tb.rows:
                for c in row.cells:
                    if id(c._tc) in seen:
                        continue
                    seen.add(id(c._tc))
                    for p in pats:
                        for m in p.finditer(c.text):
                            self.add("Warning", f"Table {ti + 1}", "Table placeholder", "Placeholder text was not replaced.",
                                     value=m.group(0), fix="Replace the placeholder.", rule="SAP-PH")

    @staticmethod
    def _yn(s):
        t = norm(s).strip(" :.-–—☐☒☑[]()")
        if not t or "yes or no" in t or "<" in t:
            return None
        if re.match(r"^(☒|☑|x|\[x\])?\s*yes\b", t):
            return True
        if re.match(r"^(☒|☑|x|\[x\])?\s*no\b", t):
            return False
        return None


# --------------------------------------------------------------------------- #
# Cross-file checks
# --------------------------------------------------------------------------- #
def cross_checks(res: Result, rules, sap: SAPChecker | None):
    d = res.data
    pa = dict((lab, val) for lab, val, _ in d.get("pa_fields", []))
    fields = rules["pre_assessment"]["fields"]
    osc_name = pa.get(fields["osc_name"], "")
    hq_name = pa.get(fields["hq_name"], "")
    scope = pa.get(fields["scope"], "")

    def add(sev, loc, fld, msg, **k):
        res.add(sev, "x", loc, fld, msg, **k)

    # SAP OSC name vs PA
    sap_name = d.get("sap_osc_name", "")
    if sap_name and (osc_name or hq_name):
        def key(s):
            return re.sub(r"[^a-z0-9]", "", re.sub(r"\b(inc|llc|corp|corporation|co|ltd|company)\b", "", s.lower()))
        if key(sap_name) not in (key(osc_name), key(hq_name)):
            add("Warning", "SAP cover / Pre-Assessment", "OSC name", "OSC name on the SAP does not match the Pre-Assessment OSC Name or HQ Organization Name.",
                value=f"SAP: {sap_name} | PA OSC: {osc_name} | PA HQ: {hq_name}", fix="Confirm the legal entity name matches SAM.gov across documents.", rule="X-NAME")

    # SAP SPRS submission type vs PA scope
    st = (d.get("sap_answers") or {}).get("score_type", "")
    if st and scope and "<" not in st:
        first = st.split()[0].strip(".,").capitalize()
        if first in ("Enterprise", "Enclave") and first != scope:
            add("Warning", "SAP section 3 / Pre-Assessment Scope", "Scope", "SPRS score submission type in the SAP differs from the Pre-Assessment Scope.",
                value=f"SAP: {st} | PA: {scope}", fix="Confirm the assessment scope type.", rule="X-SCOPE")

    # Facility CAGE codes vs PA scope codes
    pa_codes = set(d.get("pa_cage_codes") or [])
    fac_codes = set()
    for f in d.get("facilities", []):
        for c in re.split(r"[;,\s]+", f.get("cage", "")):
            if re.fullmatch(r"[A-Za-z0-9]{5}", c or ""):
                fac_codes.add(c.upper())
                if pa_codes and c.upper() not in pa_codes:
                    add("Warning", f"DRL Physical Facilities!row {f['row']}", "CAGE Code", "Facility CAGE code is not listed in the Pre-Assessment 'CAGE Code(s) in Scope'.",
                        value=c, expected="; ".join(sorted(pa_codes)), fix="Add the code to the Pre-Assessment form or correct the facility record.", rule="X-CAGE")
    if pa_codes and d.get("facilities") and fac_codes:
        for c in sorted(pa_codes - fac_codes):
            add("Info", "Pre-Assessment / DRL Physical Facilities", "CAGE Code", "In-scope CAGE code is not tied to any listed facility.",
                value=c, fix="Confirm the facility for this CAGE code is listed if it is in scope.", rule="X-CAGE")

    # ESPs vs cloud-hosted assets
    esps = d.get("esps", [])
    assets = d.get("assets", [])
    kws = rules["asset_inventory"]["cloud_keywords"]
    cloud_assets = [a for a in assets if any(k in a.get("host", "").lower() for k in kws) and a.get("cat") != "Out-of-Scope"]
    if "esps" in d and "assets" in d:
        if cloud_assets and not esps:
            add("Warning", "Asset Inventory", "ESPs", "Assets are hosted by cloud or service providers, but the Consolidated ESP tab has no ESP records. SAP section 5.5 says missing ESP information postpones the audit start.",
                value=compress_ids([a.get("asset", "") for a in cloud_assets], 10), fix="List each in-scope ESP on the Consolidated ESP tab.", rule="X-ESP")
        elif esps:
            esp_tokens = []
            for e in esps:
                toks = [t for t in re.split(r"[^a-z0-9]+", e.get("name", "").lower()) if len(t) > 2 and t not in ("inc", "llc", "the", "corp", "cloud", "services", "service", "gov")]
                esp_tokens.append(toks)
            for a in cloud_assets:
                text = (a.get("host", "") + " " + a.get("asset", "")).lower()
                if not any(any(t in text for t in toks) for toks in esp_tokens if toks):
                    add("Info", f"Asset Inventory Main Scoped Assets!row {a['row']}", a.get("asset") or "Asset",
                        "Asset is hosted by a cloud or service provider that does not match any ESP name on the Consolidated ESP tab.",
                        value=short(a.get("host", ""), 100), fix="Confirm whether this provider is an in-scope ESP and list it if so.", rule="X-ESPMATCH")

    # SAP Yes answers vs actual content
    if sap is not None and getattr(sap, "yn", None) is not None:
        items = {it["id"]: it for it in rules["sap"]["yes_no_items"]}
        checks = []
        if "assets" in d:
            checks.append(("4.07", not d.get("assets"), "SAP says the Asset Inventory Scoping file is complete, but no assets are listed."))
        if "oos" in d:
            checks.append(("4.15", not d.get("oos"), "SAP affirms out-of-scope areas were added, but the Out of Scope Areas tab is empty."))
        if "team" in d:
            checks.append(("4.16", not d.get("team"), "SAP affirms key personnel were added, but the OSC Assessment Team tab is empty."))
        if "drl_mapping" in d:
            mc, nc = d["drl_mapping"]
            checks.append(("4.08", nc and mc == 0, "SAP says the Document Request List is complete, but no documents are mapped on the Evidence Plan."))
        for iid, cond, msg in checks:
            if cond and sap.yn.get(iid) is True:
                add("Error", f"SAP item {iid}", items[iid]["label"], msg + " The affirmation conflicts with the submitted file.",
                    value="SAP answer: Yes", fix="Reconcile the file content with the SAP affirmation before Phase 2.", rule="X-AFFIRM")

    # SAP signer on team
    signer = d.get("sap_osc_signer")
    if signer and d.get("team"):
        if norm(signer) not in [norm(t.get("name")) for t in d["team"]]:
            add("Info", "SAP Plan Acceptance / DRL OSC Assessment Team", "Affirming official", "The OSC signer on the SAP is not listed on the OSC Assessment Team tab.",
                value=signer, fix="Add the affirming official to the team list (needed as the eMASS OSC Assessment Official).", rule="X-SIGNER")


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #
def run_all(paths: dict, rules: dict | None = None) -> Result:
    """paths: {"pa": path|None, "ai": ..., "drl": ..., "sap": ...}"""
    rules = rules or load_rules()
    res = Result(generated=_dt.datetime.now().strftime("%d-%b-%Y %H:%M"))
    sap = None
    for key in ("pa", "ai", "drl", "sap"):
        p = paths.get(key)
        if not p:
            continue
        res.files[key] = {"path": p, "name": os.path.basename(p), "sha256": sha256(p)}
        try:
            if key == "pa":
                PreAssessmentChecker(p, rules, res).run()
            elif key == "ai":
                AssetInventoryChecker(p, rules, res).run()
            elif key == "drl":
                DRLChecker(p, rules, res).run()
            elif key == "sap":
                sap = SAPChecker(p, rules, res)
                sap.run()
        except Exception as e:  # keep going on other files
            import traceback
            res.add("Error", key, "file", "File could not be read", f"The checker could not process this file: {e}",
                    value=traceback.format_exc(limit=2)[-300:], fix="Confirm the file is the correct template, is closed in Excel/Word, and is not password protected.",
                    owner="Ignyte", rule="READ")
    if len([k for k in res.files]) >= 2:
        cross_checks(res, rules, sap)
    return res
