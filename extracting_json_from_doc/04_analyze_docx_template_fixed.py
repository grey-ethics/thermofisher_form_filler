#!/usr/bin/env python3
"""
analyze_docx_template_fixed.py

Scans a .docx template and reports content-controls and ballot glyphs.
DOCX_PATH and OUTPUT_JSON are defined inside this file.

Usage:
  python analyze_docx_template_fixed.py
"""

import zipfile
from lxml import etree as ET
import re
import json
import os
import sys

# ---------- CONFIG (edit these paths if needed) ----------
DOCX_PATH = r"C:\Users\K Santosh Kumar\Desktop\HEALTHARK\04_thermofisher\reference_template.docx"
OUTPUT_JSON = r"C:\Users\K Santosh Kumar\Desktop\HEALTHARK\04_thermofisher\04_template_analysis_report_fixed.json"
# ---------- END CONFIG ----------

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}

# Matches common ballot/check glyphs: ☑ ☐ ☒ □ ▢ ✓ ✔ ✗
GLYPH_REGEX = re.compile(r'[\u2610\u2611\u2612\u25A1\u25A2\u2713\u2714\u2717]')

def parse_part(zipf, partname):
    try:
        raw = zipf.read(partname)
    except KeyError:
        return None
    parser = ET.XMLParser(ns_clean=True, recover=True)
    root = ET.fromstring(raw, parser=parser)
    return root

def extract_text(elem):
    texts = elem.xpath('.//w:t/text()', namespaces=NS)
    return ''.join(texts).strip()

def get_table_coordinates(tc):
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
    direct_trs = [r for r in tbl if ET.QName(r.tag).localname == 'tr']
    try:
        row_index = direct_trs.index(tr) + 1
    except ValueError:
        row_index = None
    tcs = [c for c in tr if ET.QName(c.tag).localname == 'tc']
    try:
        col_index = tcs.index(tc) + 1
    except ValueError:
        col_index = None
    return (row_index, col_index)

def analyze_docx(docx_path):
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"DOCX not found: {docx_path}")

    report = {"file": os.path.basename(docx_path), "abs_path": os.path.abspath(docx_path),
              "parts_scanned": [], "content_controls": [], "glyph_cells": []}

    with zipfile.ZipFile(docx_path, 'r') as z:
        all_parts = z.namelist()
        # print a short sample of word/ parts for debugging
        word_parts = [p for p in all_parts if p.startswith('word/')]
        print("DOCX word parts present (sample):")
        for p in word_parts[:40]:
            print("  ", p)
        # Ensure main document is scanned (no leading slash)
        parts_to_scan = ['word/document.xml']
        for name in all_parts:
            if name.startswith('word/header') or name.startswith('word/footer') or name.startswith('word/footnotes') or name.startswith('word/endnotes'):
                parts_to_scan.append(name)

        for part in parts_to_scan:
            root = parse_part(z, part)
            if root is None:
                print("Part not found in docx zip, skipping:", part)
                continue
            report["parts_scanned"].append(part)

            # content controls (w:sdt)
            sdts = root.findall('.//w:sdt', namespaces=NS)
            for idx, sdt in enumerate(sdts, start=1):
                sdtPr = sdt.find('.//w:sdtPr', namespaces=NS)
                tag = None
                alias = None
                ctype = None
                choices = []
                if sdtPr is not None:
                    tag_elem = sdtPr.find('.//w:tag', namespaces=NS)
                    if tag_elem is not None:
                        tag = tag_elem.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') or tag_elem.attrib.get('val')
                    alias_elem = sdtPr.find('.//w:alias', namespaces=NS)
                    if alias_elem is not None:
                        alias = alias_elem.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') or alias_elem.attrib.get('val')
                    if sdtPr.find('.//w14:checkbox', namespaces=NS) is not None:
                        ctype = 'checkbox'
                    dd = sdtPr.find('.//w:dropDownList', namespaces=NS)
                    combo = sdtPr.find('.//w:comboBox', namespaces=NS)
                    if dd is not None:
                        ctype = 'dropdown'
                        for e in dd.findall('.//w:listEntry', namespaces=NS) + dd.findall('.//w:entry', namespaces=NS):
                            v = e.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') or e.attrib.get('val') or (e.text or '').strip()
                            if v:
                                choices.append(v)
                    if combo is not None:
                        ctype = 'combo'
                        for e in combo.findall('.//w:listEntry', namespaces=NS) + combo.findall('.//w:entry', namespaces=NS):
                            v = e.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') or e.attrib.get('val') or (e.text or '').strip()
                            if v:
                                choices.append(v)

                text = extract_text(sdt)
                # coordinates if inside table
                tcs = sdt.xpath('ancestor::w:tc', namespaces=NS)
                tbl_index = None; row=None; col=None
                if tcs:
                    tc = tcs[0]
                    row, col = get_table_coordinates(tc)
                    # identify the table index
                    tables = root.findall('.//w:tbl', namespaces=NS)
                    for i, tbl in enumerate(tables, start=1):
                        # if the tc is a descendant of this tbl element, it's the table
                        if any(True for _ in tbl.iter() if _ is tc):
                            tbl_index = i
                            break

                report["content_controls"].append({
                    "part": part,
                    "sdt_index": idx,
                    "tag": tag,
                    "alias": alias,
                    "type": ctype or "unknown",
                    "choices": choices,
                    "text": text,
                    "table_index": tbl_index,
                    "table_row": row,
                    "table_col": col
                })

            # scan tables for glyphs
            tables = root.findall('.//w:tbl', namespaces=NS)
            for t_i, tbl in enumerate(tables, start=1):
                rows = [r for r in tbl if ET.QName(r.tag).localname == 'tr']
                for r_idx, tr in enumerate(rows, start=1):
                    tcs = [c for c in tr if ET.QName(c.tag).localname == 'tc']
                    for c_idx, tc in enumerate(tcs, start=1):
                        txt = extract_text(tc)
                        glyphs = GLYPH_REGEX.findall(txt)
                        if glyphs:
                            report["glyph_cells"].append({
                                "part": part,
                                "table_index": t_i,
                                "row": r_idx,
                                "col": c_idx,
                                "text": txt,
                                "glyphs": glyphs
                            })

    return report

def main():
    try:
        print("Analyzing:", DOCX_PATH)
        if not os.path.exists(DOCX_PATH):
            print("ERROR: DOCX not found at", DOCX_PATH)
            sys.exit(1)

        report = analyze_docx(DOCX_PATH)
        out = OUTPUT_JSON
        with open(out, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        print("\nAnalysis complete.")
        print("Input:", os.path.abspath(DOCX_PATH))
        print("Report:", os.path.abspath(out))
        print("Content controls found:", len(report.get("content_controls", [])))
        print("Table cells with ballot/check glyphs:", len(report.get("glyph_cells", [])))
    except Exception as e:
        print("Error during analysis:", e)
        raise

if __name__ == "__main__":
    main()
