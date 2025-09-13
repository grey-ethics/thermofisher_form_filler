#!/usr/bin/env python3
"""
analyze_docx_template.py

Scans a .docx template and reports:
 - structured Content Controls (w:sdt) with type, tag, alias, text and table coordinates when inside a table
 - all table cells that contain ballot/check glyphs (☐, ☑, ☒, ✓, etc.)
 - saves a JSON report and prints summary DataFrames

Usage:
  python analyze_docx_template.py
"""

import zipfile
from lxml import etree as ET
import re
import json
import os
import pandas as pd
from collections import defaultdict


# ---------- CONFIG ----------
DOCX_PATH = r"C:\Users\K Santosh Kumar\Desktop\HEALTHARK\04_thermofisher\reference_template.docx"
OUTPUT_JSON = r"C:\Users\K Santosh Kumar\Desktop\HEALTHARK\04_thermofisher\02_template_analysis_report.json"
# ---------- END CONFIG ----------


# Namespaces used by WordprocessingML
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'xml': 'http://www.w3.org/XML/1998/namespace'
}

# Regex for ballot / check glyphs commonly used
GLYPH_REGEX = re.compile(r'[\u2610\u2611\u2612\u25A1\u25A2\u2713\u2714\u2717]')

def parse_part(zipf, partname):
    """Read and parse a .xml part from the docx zip. Returns an lxml root or None."""
    try:
        raw = zipf.read(partname)
    except KeyError:
        return None
    parser = ET.XMLParser(ns_clean=True, recover=True)
    root = ET.fromstring(raw, parser=parser)
    return root

def extract_text(elem):
    """Concatenate all w:t text nodes found under elem."""
    texts = elem.xpath('.//w:t/text()', namespaces=NS)
    return ''.join(texts).strip()

def get_table_coordinates(tc):
    """
    Given a <w:tc> element, return (row_index, col_index) within its immediate table.
    Indexes are 1-based. Returns (None, None) if unable.
    """
    # find the ancestor tr and tbl
    tr = tc.getparent()
    while tr is not None and ET.QName(tr.tag).localname != 'tr':
        tr = tr.getparent()
    if tr is None:
        return (None, None)

    tbl = tr.getparent()
    while tbl is not None and ET.QName(tbl.tag).localname != 'tbl':
        tbl = tbl.getparent()
    if tbl is None:
        return (None, None)

    # find direct 'tr' children of tbl (in order)
    direct_trs = [r for r in tbl if ET.QName(r.tag).localname == 'tr']
    try:
        row_index = direct_trs.index(tr) + 1
    except ValueError:
        # fallback
        row_index = None

    # within the tr, find direct 'tc' children
    tcs = [c for c in tr if ET.QName(c.tag).localname == 'tc']
    try:
        col_index = tcs.index(tc) + 1
    except ValueError:
        col_index = None

    return (row_index, col_index)

def analyze_docx(docx_path):
    report = {
        "file": os.path.basename(docx_path),
        "parts_scanned": [],
        "content_controls": [],
        "tables": [],
        "glyph_cells": []
    }

    with zipfile.ZipFile(docx_path, 'r') as z:
        # Part list to scan: main document + any headers/footers/footnotes/endnotes
        parts_to_scan = ['/word/document.xml']
        for name in z.namelist():
            if name.startswith('word/header') or name.startswith('word/footer') or name.startswith('word/footnotes') or name.startswith('word/endnotes'):
                parts_to_scan.append(name)

        for part in parts_to_scan:
            root = parse_part(z, part)
            if root is None:
                continue
            report["parts_scanned"].append(part)

            # Find tables and sdts in this part
            tables = root.findall('.//w:tbl', namespaces=NS)
            sdts = root.findall('.//w:sdt', namespaces=NS)

            # Build mapping of table element id to local index
            tbl_to_index = {}
            for i, tbl in enumerate(tables, start=1):
                tbl_to_index[id(tbl)] = i

            # ---- content controls (w:sdt) ----
            for idx, sdt in enumerate(sdts, start=1):
                sdtPr = sdt.find('.//w:sdtPr', namespaces=NS)
                tag = None
                alias = None
                ctype = None
                choices = []
                checked_default = None

                if sdtPr is not None:
                    tag_elem = sdtPr.find('.//w:tag', namespaces=NS)
                    if tag_elem is not None:
                        # typically attribute is w:val - fetch safely
                        tag = tag_elem.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') or tag_elem.attrib.get('val')

                    alias_elem = sdtPr.find('.//w:alias', namespaces=NS)
                    if alias_elem is not None:
                        alias = alias_elem.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') or alias_elem.attrib.get('val')

                    # checkbox (w14:checkbox)
                    cb = sdtPr.find('.//w14:checkbox', namespaces=NS)
                    if cb is not None:
                        ctype = 'checkbox'
                        checked_elem = cb.find('.//w14:checked', namespaces=NS)
                        if checked_elem is not None:
                            checked_default = checked_elem.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') or checked_elem.attrib.get('val')

                    # dropdown / combo (w:dropDownList, w:comboBox)
                    dd = sdtPr.find('.//w:dropDownList', namespaces=NS)
                    combo = sdtPr.find('.//w:comboBox', namespaces=NS)
                    if dd is not None:
                        ctype = 'dropdown'
                        entries = dd.findall('.//w:listEntry', namespaces=NS) or dd.findall('.//w:entry', namespaces=NS)
                        for e in entries:
                            v = e.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') or e.attrib.get('val') or (e.text or '').strip()
                            if v:
                                choices.append(v)
                    elif combo is not None:
                        ctype = 'combo'
                        entries = combo.findall('.//w:listEntry', namespaces=NS) or combo.findall('.//w:entry', namespaces=NS)
                        for e in entries:
                            v = e.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') or e.attrib.get('val') or (e.text or '').strip()
                            if v:
                                choices.append(v)

                    # legacy formfield detection
                    ff = sdtPr.find('.//w:ffData', namespaces=NS)
                    if ff is not None and ctype is None:
                        ctype = 'legacy_formfield'
                        if ff.find('.//w:checkBox', namespaces=NS) is not None:
                            ctype = 'legacy_checkbox'
                        if ff.find('.//w:listEntry', namespaces=NS) is not None:
                            ctype = 'legacy_dropdown'

                sdt_text = extract_text(sdt)

                # determine if sdt is in a table cell
                tcs = sdt.xpath('ancestor::w:tc', namespaces=NS)
                table_index = None
                row_idx = None
                col_idx = None
                if tcs:
                    tc = tcs[0]
                    coords = get_table_coordinates(tc)
                    row_idx, col_idx = coords
                    # find parent table to get its index among tables in part
                    ancestor_tbl = tc.getparent()
                    while ancestor_tbl is not None and ET.QName(ancestor_tbl.tag).localname != 'tbl':
                        ancestor_tbl = ancestor_tbl.getparent()
                    if ancestor_tbl is not None:
                        table_index = tbl_to_index.get(id(ancestor_tbl))

                report["content_controls"].append({
                    "part": part,
                    "sdt_index": idx,
                    "tag": tag,
                    "alias": alias,
                    "type": ctype or 'unknown',
                    "choices": choices,
                    "checked_default": checked_default,
                    "text": sdt_text,
                    "table_index_in_part": table_index,
                    "table_row": row_idx,
                    "table_col": col_idx
                })

            # ---- scan tables for glyphs ----
            for t_i, tbl in enumerate(tables, start=1):
                # iterate rows (direct child 'tr' elements)
                direct_rows = [r for r in tbl if ET.QName(r.tag).localname == 'tr']
                for r_idx, tr in enumerate(direct_rows, start=1):
                    tcs = [c for c in tr if ET.QName(c.tag).localname == 'tc']
                    for c_idx, tc in enumerate(tcs, start=1):
                        text = extract_text(tc)
                        glyphs = GLYPH_REGEX.findall(text)
                        if glyphs:
                            report["glyph_cells"].append({
                                "part": part,
                                "table_index_in_part": t_i,
                                "row": r_idx,
                                "col": c_idx,
                                "text": text,
                                "glyphs": glyphs,
                                "glyph_count": len(glyphs)
                            })
                        # always add to a minimal table summary
                        report.setdefault("tables", []).append({
                            "part": part,
                            "table_index_in_part": t_i,
                            "row": r_idx,
                            "col": c_idx,
                            "text_preview": (text[:120] + '...') if len(text) > 120 else text,
                            "glyph_count": len(glyphs)
                        })

    # Save JSON report:
    with open(OUTPUT_JSON, 'w', encoding='utf-8') as outf:
        json.dump(report, outf, indent=2, ensure_ascii=False)

    # Prepare small pandas previews (for console)
    cc_df = pd.DataFrame(report["content_controls"])
    glyph_df = pd.DataFrame(report["glyph_cells"])

    # Print summary
    print("==== Template analysis summary ====")
    print(f"File: {report['file']}")
    print(f"Parts scanned: {', '.join(report['parts_scanned'])}")
    print(f"Structured content controls found: {len(report['content_controls'])}")
    print(f"Table cells with ballot/check glyphs found: {len(report['glyph_cells'])}")
    print(f"JSON report saved to: {OUTPUT_JSON}")
    print("===================================\n")

    if not cc_df.empty:
        print("Sample content-controls (first 10):")
        print(cc_df.head(10).to_string(index=False))
    else:
        print("No structured content controls found (w:sdt) in scanned parts.")

    if not glyph_df.empty:
        print("\nSample glyph-containing table cells (first 10):")
        print(glyph_df.head(10).to_string(index=False))
    else:
        print("No table cells containing ballot/check glyphs were found by the glyph regex.")

    return report

if __name__ == "__main__":
    if not os.path.exists(DOCX_PATH):
        print("ERROR: DOCX not found at", DOCX_PATH)
    else:
        analyze_docx(DOCX_PATH)
