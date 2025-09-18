#!/usr/bin/env python3
# force_escape_text_nodes.py
# Raw-text-based fixer: escapes literal < and > inside <w:t>...</w:t> sequences
# so the document XML becomes parseable.

import zipfile, os, shutil, re, sys
from lxml import etree as ET

INPUT_DOCX = "refernce_template_unlocked.docx"   # <<--- change if needed
BACKUP = INPUT_DOCX.replace(".docx", "_rawfix_bak.docx")
OUT_DOCX = INPUT_DOCX.replace(".docx", "_forced_escaped.docx")

# Parts we will attempt to fix (main doc + headers/footers + glossary)
TARGET_PARTS = set([
    "word/document.xml",
    "word/footer1.xml", "word/footer2.xml", "word/footer3.xml",
    "word/header1.xml", "word/header2.xml", "word/header3.xml",
    "word/glossary/document.xml", "word/footnotes.xml", "word/endnotes.xml"
])

# regex finds <w:t ...> ... </w:t> with non-greedy match for inner text
RE_W_T = re.compile(r'(<w:t\b[^>]*>)(.*?)(</w:t>)', flags=re.DOTALL | re.IGNORECASE)

def escape_inner_text(m):
    open_tag, inner, close_tag = m.group(1), m.group(2), m.group(3)
    # protect already-escaped sequences
    inner = inner.replace('&lt;', '\0LT_ESC\0').replace('&gt;', '\0GT_ESC\0')
    # replace raw < and > in inner text
    inner = inner.replace('<', '&lt;').replace('>', '&gt;')
    # restore previously escaped placeholders
    inner = inner.replace('\0LT_ESC\0', '&lt;').replace('\0GT_ESC\0', '&gt;')
    return open_tag + inner + close_tag

def process_part_bytes(b):
    try:
        s = b.decode('utf-8')
    except UnicodeDecodeError:
        # try with windows-1252 fallback (sometimes Word uses different encoding declaration)
        s = b.decode('cp1252', errors='replace')
    new_s, n = RE_W_T.subn(escape_inner_text, s)
    return new_s.encode('utf-8'), n

def main():
    if not os.path.exists(INPUT_DOCX):
        print("Input DOCX not found:", INPUT_DOCX); sys.exit(1)

    # Backup original
    shutil.copy2(INPUT_DOCX, BACKUP)
    print("Backup created:", BACKUP)

    changed_any = False
    with zipfile.ZipFile(INPUT_DOCX, 'r') as zin:
        names = zin.namelist()
        with zipfile.ZipFile(OUT_DOCX, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for nm in names:
                data = zin.read(nm)
                if nm in TARGET_PARTS or (nm.startswith("word/header") or nm.startswith("word/footer") or nm.startswith("word/")):
                    # We will attempt to process only common Word xml parts (but be conservative)
                    if nm in TARGET_PARTS or nm.startswith("word/"):
                        try:
                            newdata, nrepl = process_part_bytes(data)
                            if nrepl > 0:
                                changed_any = True
                                print(f"Escaped {nrepl} <w:t> inner occurrences in {nm}")
                                zout.writestr(nm, newdata)
                                continue
                        except Exception as e:
                            print(f"Warning: processing {nm} failed with: {e}; writing original")
                # default: copy original
                zout.writestr(nm, data)

    print("Wrote fixed docx to:", OUT_DOCX)
    if not changed_any:
        print("Note: no replacements were made. The file may not contain raw < or > inside w:t nodes.")
    # Quick validation: try to parse document.xml inside out file
    try:
        with zipfile.ZipFile(OUT_DOCX, 'r') as z2:
            content = z2.read("word/document.xml")
            ET.fromstring(content)   # will raise if still broken
        print("Validation: word/document.xml parsed OK.")
    except Exception as e:
        print("Validation: ERROR parsing word/document.xml:", e)
        print("Try opening the output in Word with Open & Repair. If it still fails, paste the first snippet+error here.")
    print("Done.")

if __name__ == '__main__':
    main()
