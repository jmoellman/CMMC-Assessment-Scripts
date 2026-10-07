"""
Safe writer for the eMASS Pre-Assessment form.

Saving with openpyxl can drop template features (for example x14 data
validation dropdowns). This module copies the original .xlsx and edits only
the XML of the cells being corrected, leaving every other part byte-for-byte
unchanged. The original file is never overwritten.
"""
from __future__ import annotations

import os
import re
import shutil
import tempfile
import zipfile

from lxml import etree

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
M = "{%s}" % NS["m"]


def next_version_path(src):
    base, ext = os.path.splitext(src)
    base = re.sub(r"_corrected_v\d+$", "", base)
    n = 1
    while os.path.exists(f"{base}_corrected_v{n}{ext}"):
        n += 1
    return f"{base}_corrected_v{n}{ext}"


def _col_num(ref):
    letters = re.match(r"([A-Z]+)", ref).group(1)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _sheet_part(z: zipfile.ZipFile, sheet_name_prefix: str):
    wb = etree.fromstring(z.read("xl/workbook.xml"))
    rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid = None
    for s in wb.iter(M + "sheet"):
        if s.get("name", "").strip().lower().startswith(sheet_name_prefix.lower()):
            rid = s.get("{%s}id" % NS["r"])
            break
    if rid is None:
        raise ValueError(f"Worksheet '{sheet_name_prefix}' not found")
    for r in rels:
        if r.get("Id") == rid:
            t = r.get("Target")
            t = t.lstrip("/")
            return t if t.startswith("xl/") else "xl/" + t
    raise ValueError("Worksheet relationship not found")


def _get_cell(sheet_data, ref, create=True):
    rownum = int(re.search(r"(\d+)$", ref).group(1))
    row = None
    for r in sheet_data.findall(M + "row"):
        rn = int(r.get("r"))
        if rn == rownum:
            row = r
            break
        if rn > rownum:
            if not create:
                return None
            row = etree.Element(M + "row", r=str(rownum))
            r.addprevious(row)
            break
    if row is None:
        if not create:
            return None
        row = etree.SubElement(sheet_data, M + "row", r=str(rownum))
    col = _col_num(ref)
    for c in row.findall(M + "c"):
        cc = _col_num(c.get("r"))
        if cc == col:
            return c
        if cc > col:
            if not create:
                return None
            new = etree.Element(M + "c", r=ref)
            c.addprevious(new)
            return new
    if not create:
        return None
    return etree.SubElement(row, M + "c", r=ref)


def _clear_value(c):
    for child in list(c):
        if child.tag in (M + "v", M + "is", M + "f"):
            c.remove(child)
    if "t" in c.attrib:
        del c.attrib["t"]


def _set(c, value):
    _clear_value(c)
    if value is None or value == "":
        return
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        v = etree.SubElement(c, M + "v")
        v.text = str(value)
        return
    c.set("t", "inlineStr")
    is_ = etree.SubElement(c, M + "is")
    t = etree.SubElement(is_, M + "t")
    t.text = str(value)
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def apply_fixes(src, dst, fixes, sheet="Pre-Assessment"):
    """fixes: list of autofix dicts from engine (type move|set). Returns list of applied descriptions."""
    if os.path.abspath(src) == os.path.abspath(dst):
        raise ValueError("Refusing to overwrite the original file")
    applied = []
    with zipfile.ZipFile(src) as z:
        part = _sheet_part(z, sheet)
        root = etree.fromstring(z.read(part))
        sd = root.find(M + "sheetData")
        for fx in fixes:
            val = fx.get("value")
            if fx.get("field") == "employees" and re.fullmatch(r"\d+", str(val or "")):
                val = int(val)
            if fx["type"] == "move":
                _set(_get_cell(sd, fx["to"]), val)
                restore = fx.get("restore")
                if restore is not None:
                    _set(_get_cell(sd, fx["from"]), restore)
            elif fx["type"] == "set":
                _set(_get_cell(sd, fx["to"]), val)
            applied.append(fx["desc"])
        new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(os.path.abspath(dst)) or None)
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
                for item in z.infolist():
                    data = new_xml if item.filename == part else z.read(item.filename)
                    if item.filename == "xl/calcChain.xml":
                        pass
                    out.writestr(item, data)
            shutil.move(tmp, dst)
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)
    return applied
