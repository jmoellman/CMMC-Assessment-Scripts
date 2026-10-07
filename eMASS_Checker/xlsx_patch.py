"""Write text changes into a copy of an .xlsx file by editing only the affected cell XML.

Why not openpyxl.save()? The eMASS template stores its dropdown lists as Excel 2010
"x14" data validations (Score, Inherited, Requirement in POA&M, Standards Acceptance).
openpyxl drops those on save, which would quietly strip the lookups from the results
file. This module copies every part of the package byte-for-byte and only rewrites the
worksheet XML for cells that changed, plus appends the new strings to sharedStrings.xml.
"""
from __future__ import annotations

import os
import re
import shutil
import tempfile
import zipfile
from xml.sax.saxutils import escape

from openpyxl.utils import column_index_from_string, get_column_letter

NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")


def _xml_text(s: str) -> str:
    s = _ILLEGAL.sub("", s)
    s = re.sub(r"_(x[0-9A-Fa-f]{4})_", r"_x005F_\1_", s)  # escape literal _xHHHH_ the way Excel does
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    return escape(s)


def _sheet_paths(z: zipfile.ZipFile) -> dict:
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid_target = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", rels):
        tag = m.group(0)
        rid = re.search(r'\bId="([^"]+)"', tag)
        tgt = re.search(r'\bTarget="([^"]+)"', tag)
        if rid and tgt:
            t = tgt.group(1)
            t = t.lstrip("/") if t.startswith("/") else "xl/" + t
            rid_target[rid.group(1)] = t
    out = {}
    for m in re.finditer(r"<sheet\b[^>]*/>", wb):
        tag = m.group(0)
        name = re.search(r'\bname="([^"]+)"', tag).group(1)
        rid = re.search(r'\br:id="([^"]+)"', tag) or re.search(r'\bid="([^"]+)"', tag)
        name = name.replace("&amp;", "&").replace("&apos;", "'").replace("&quot;", '"').replace("&lt;", "<").replace("&gt;", ">")
        out[name] = rid_target.get(rid.group(1))
    return out


def _cell_regex(coord: str):
    # attribute order independent: <c r="F6" s="3" t="s"> or <c s="3" r="F6">
    return re.compile(r'<c(?P<attrs>\s[^>]*?\br="' + coord + r'"[^>]*?)(?:/>|>(?P<body>.*?)</c>)', re.S)


def _set_cell(xml: str, coord: str, ss_index: int | None, text: str) -> str:
    rx = _cell_regex(coord)
    m = rx.search(xml)
    if m:
        if m.group("body") and "<f" in m.group("body"):
            raise ValueError(f"{coord} contains a formula; refusing to overwrite it.")
        attrs = m.group("attrs") or ""
        attrs = re.sub(r'\s(?:r|t|cm|vm)="[^"]*"', "", attrs).rstrip("/").rstrip()
        new = _cell_xml(coord, attrs, ss_index, text)
        return xml[: m.start()] + new + xml[m.end():]
    # cell not present: insert into its row in column order
    col_letters = re.match(r"[A-Z]+", coord).group(0)
    row_num = int(coord[len(col_letters):])
    col_idx = column_index_from_string(col_letters)
    new = _cell_xml(coord, "", ss_index, text)
    rrx = re.compile(r'<row r="' + str(row_num) + r'"(?P<attrs>[^>]*?)(?:/>|>(?P<body>.*?)</row>)', re.S)
    rm = rrx.search(xml)
    if rm:
        if rm.group("body") is None:  # self-closing row
            attrs = rm.group("attrs").rstrip("/").rstrip()
            return xml[: rm.start()] + f'<row r="{row_num}"{attrs}>{new}</row>' + xml[rm.end():]
        body = rm.group("body")
        insert_at = len(body)
        for cm in re.finditer(r'<c r="([A-Z]+)\d+"', body):
            if column_index_from_string(cm.group(1)) > col_idx:
                insert_at = cm.start()
                break
        body = body[:insert_at] + new + body[insert_at:]
        attrs = re.sub(r'\sspans="[^"]*"', "", rm.group("attrs"))
        return xml[: rm.start()] + f'<row r="{row_num}"{attrs}>{body}</row>' + xml[rm.end():]
    # row not present: insert row in order inside sheetData
    sd = re.search(r"<sheetData\s*/>|<sheetData>(?P<body>.*?)</sheetData>", xml, re.S)
    row_xml = f'<row r="{row_num}">{new}</row>'
    if sd.group(0).endswith("/>"):
        return xml[: sd.start()] + f"<sheetData>{row_xml}</sheetData>" + xml[sd.end():]
    body = sd.group("body")
    insert_at = len(body)
    for rr in re.finditer(r'<row r="(\d+)"', body):
        if int(rr.group(1)) > row_num:
            insert_at = rr.start()
            break
    body = body[:insert_at] + row_xml + body[insert_at:]
    return xml[: sd.start()] + f"<sheetData>{body}</sheetData>" + xml[sd.end():]


def _cell_xml(coord, attrs, ss_index, text):
    if ss_index is None:
        return f'<c r="{coord}"{attrs} t="inlineStr"><is><t xml:space="preserve">{_xml_text(text)}</t></is></c>'
    return f'<c r="{coord}"{attrs} t="s"><v>{ss_index}</v></c>'


def write_copy(src: str, dst: str, changes: dict) -> int:
    """changes: {(sheet_name, row, col): new_text}. Returns number of cells written.
    Writes to a temporary file first, then moves it into place; never touches src."""
    if os.path.abspath(src) == os.path.abspath(dst):
        raise ValueError("Output must be a new file; the original is never overwritten.")
    with zipfile.ZipFile(src) as zin:
        sheets = _sheet_paths(zin)
        names = zin.namelist()
        ss_path = "xl/sharedStrings.xml" if "xl/sharedStrings.xml" in names else None
        ss_xml = zin.read(ss_path).decode("utf-8") if ss_path else None
        by_sheet: dict = {}
        for (sheet, row, col), text in changes.items():
            if sheet not in sheets or not sheets[sheet]:
                raise ValueError(f'Worksheet "{sheet}" not found in workbook.')
            by_sheet.setdefault(sheets[sheet], []).append((f"{get_column_letter(col)}{row}", text))
        new_parts = {}
        if ss_xml is not None:
            existing = len(re.findall(r"<si>|<si\s", ss_xml))
            additions = []
        for part, cells in by_sheet.items():
            xml = zin.read(part).decode("utf-8")
            for coord, text in cells:
                idx = None
                if ss_xml is not None:
                    idx = existing + len(additions)
                    additions.append(f'<si><t xml:space="preserve">{_xml_text(text)}</t></si>')
                xml = _set_cell(xml, coord, idx, text)
            new_parts[part] = xml
        if ss_xml is not None and additions:
            ss_xml = ss_xml.replace("</sst>", "".join(additions) + "</sst>")
            total = existing + len(additions)
            ss_xml = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{total}"', ss_xml, count=1)
            cm = re.search(r'\bcount="(\d+)"', ss_xml)
            if cm:
                ss_xml = ss_xml[: cm.start()] + f'count="{int(cm.group(1)) + len(additions)}"' + ss_xml[cm.end():]
            new_parts[ss_path] = ss_xml
        fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(os.path.abspath(dst)) or None)
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp, "w") as zout:
                for info in zin.infolist():
                    data = new_parts[info.filename].encode("utf-8") if info.filename in new_parts else zin.read(info.filename)
                    zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                    zi.compress_type = info.compress_type
                    zi.external_attr = info.external_attr
                    zout.writestr(zi, data)
            shutil.move(tmp, dst)
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)
    return sum(len(v) for v in by_sheet.values())


def next_version_path(src: str) -> str:
    folder, base = os.path.split(os.path.abspath(src))
    stem, ext = os.path.splitext(base)
    stem = re.sub(r"_corrected_v\d+$", "", stem)
    n = 1
    while True:
        p = os.path.join(folder, f"{stem}_corrected_v{n}{ext}")
        if not os.path.exists(p):
            return p
        n += 1
