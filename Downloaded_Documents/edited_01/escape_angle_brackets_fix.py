#!/usr/bin/env python3
# escape_angle_brackets_fix.py
import zipfile, os, io, shutil
from lxml import etree as ET

# -- EDIT this path to point to your problematic docx --
IN_PATH = r"refernce_template_unlocked.docx"
# Backup and output names (auto)
BACKUP = IN_PATH.replace(".docx", "_bak.docx")
OUT_PATH = IN_PATH.replace(".docx", "_escaped.docx")

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

def escape_text_nodes(xml_bytes):
    parser = ET.XMLParser(ns_clean=True, recover=False, encoding='utf-8')
    root = ET.fromstring(xml_bytes, parser=parser)
    changed = False
    for t in root.xpath('.//w:t', namespaces=NS):
        if t.text:
            new = t.text.replace("<", "&lt;").replace(">", "&gt;")
            if new != t.text:
                t.text = new
                changed = True
    return ET.tostring(root, encoding='utf-8', xml_declaration=True), changed

def main():
    if not os.path.exists(IN_PATH):
        print("Input not found:", IN_PATH); return
    shutil.copy2(IN_PATH, BACKUP)
    print("Backup created:", BACKUP)

    with zipfile.ZipFile(IN_PATH, 'r') as zin:
        names = zin.namelist()

        # we'll write to a new zip
        with zipfile.ZipFile(OUT_PATH, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            made_changes = False
            for name in names:
                data = zin.read(name)
                # Only attempt to parse / change Word XML parts (document, headers, footers, glossary)
                if name in ("word/document.xml",) or name.startswith("word/header") or name.startswith("word/footer") or name.startswith("word/glossary"):
                    try:
                        newdata, changed = escape_text_nodes(data)
                        if changed:
                            made_changes = True
                            zout.writestr(name, newdata)
                        else:
                            zout.writestr(name, data)
                    except Exception as e:
                        # If parse fails, write original and note it
                        print("Warning: could not parse part", name, " — leaving unchanged. Error:", e)
                        zout.writestr(name, data)
                else:
                    zout.writestr(name, data)
    print("Wrote copy at:", OUT_PATH)
    if not made_changes:
        print("No text nodes required escaping (no changes).")
    # quick validation: try to parse document.xml
    try:
        with zipfile.ZipFile(OUT_PATH, 'r') as z2:
            _ = z2.read("word/document.xml")
            ET.fromstring(_, parser=ET.XMLParser(ns_clean=True, recover=False))
        print("Validation: word/document.xml parsed OK.")
    except Exception as e:
        print("Validation: ERROR parsing word/document.xml:", e)
        print("Try opening the file in Word (Open & Repair).")

if __name__ == "__main__":
    main()
