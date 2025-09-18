# escape_angle_in_text_nodes.py
# Usage:
#   pip install lxml
#   python escape_angle_in_text_nodes.py
#
# This script:
# - makes a backup of DOCX_PATH (if not already present)
# - for each word/*.xml part, finds <w:t>...</w:t> and <w:instrText>...</w:instrText>
#   and replaces any literal '<' or '>' chars inside those text contents with &lt; / &gt;
# - writes a new DOCX with suffix _fixed.docx and validates with lxml.
#
import os, sys, shutil, zipfile, re
from lxml import etree

DOCX_PATH = r"C:\Users\K Santosh Kumar\Desktop\HEALTHARK\04_thermofisher\Downloaded_Documents\edited\refernce_template_unlocked_angles_fixed.docx"
# adjust path above to match the file you want to repair

if not os.path.exists(DOCX_PATH):
    print("DOCX not found:", DOCX_PATH); sys.exit(1)

BASE = os.path.splitext(DOCX_PATH)[0]
BACKUP = BASE + "_escape_bak.docx"
OUTPATH = BASE + "_escaped_fixed.docx"

def backup_if_needed():
    if not os.path.exists(BACKUP):
        shutil.copy2(DOCX_PATH, BACKUP)
        print("Backup created:", BACKUP)
    else:
        print("Backup already exists:", BACKUP)

# regex to find <w:t ...> ... </w:t> (non-greedy)
# xml parts are typically one long line, so use DOTALL
T_TAG_RE = re.compile(r'(<w:t\b[^>]*>)(.*?)(</w:t>)', flags=re.DOTALL)
INSTR_RE = re.compile(r'(<w:instrText\b[^>]*>)(.*?)(</w:instrText>)', flags=re.DOTALL)
# fldSimple encloses text as attribute/value or content; handle common <w:fldSimple instr="...">text</w:fldSimple>
FLDS_RE = re.compile(r'(<w:fldSimple\b[^>]*>)(.*?)(</w:fldSimple>)', flags=re.DOTALL)

def escape_text_content(s: str) -> str:
    # escape literal < and > inside the text node
    # (ampersands are left alone - assume they are already &amp; or valid entities)
    # but ensure we don't double-escape existing entities &lt; &gt;
    # safe approach: replace literal < and > characters
    s = s.replace('<', '&lt;')
    s = s.replace('>', '&gt;')
    return s

def process_xml_text(xml_bytes: bytes) -> (bytes, bool):
    """
    Returns (possibly_modified_bytes, changed_flag)
    Operates on decoded utf-8 text. If decoding fails, tries replace errors.
    """
    try:
        txt = xml_bytes.decode('utf-8')
    except Exception:
        txt = xml_bytes.decode('utf-8', errors='replace')

    changed = False

    def repl_t(m):
        nonlocal changed
        head, body, tail = m.group(1), m.group(2), m.group(3)
        # if body contains any literal < or >, escape them
        if '<' in body or '>' in body:
            newbody = escape_text_content(body)
            changed = True
            return head + newbody + tail
        return m.group(0)

    txt2 = T_TAG_RE.sub(repl_t, txt)
    # instrText nodes (fields/field codes)
    def repl_instr(m):
        nonlocal changed
        head, body, tail = m.group(1), m.group(2), m.group(3)
        if '<' in body or '>' in body:
            newbody = escape_text_content(body)
            changed = True
            return head + newbody + tail
        return m.group(0)

    txt2 = INSTR_RE.sub(repl_instr, txt2)

    def repl_fld(m):
        nonlocal changed
        head, body, tail = m.group(1), m.group(2), m.group(3)
        if '<' in body or '>' in body:
            newbody = escape_text_content(body)
            changed = True
            return head + newbody + tail
        return m.group(0)

    txt2 = FLDS_RE.sub(repl_fld, txt2)

    return txt2.encode('utf-8'), changed

def build_fixed_docx(in_path, out_path):
    names = None
    modified_any = []
    with zipfile.ZipFile(in_path, 'r') as zin:
        names = zin.namelist()
        # prepare new zip
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                data = zin.read(name)
                if name.startswith('word/') and name.endswith('.xml'):
                    newbytes, changed = process_xml_text(data)
                    if changed:
                        modified_any.append(name)
                        zout.writestr(name, newbytes)
                        print("Escaped angle-brackets in text nodes of:", name)
                    else:
                        zout.writestr(name, data)
                else:
                    # copy other parts unchanged
                    zout.writestr(name, data)
    return modified_any

def validate_docx(path):
    errs = []
    with zipfile.ZipFile(path, 'r') as z:
        for name in z.namelist():
            if name.endswith('.xml'):
                raw = z.read(name)
                try:
                    etree.fromstring(raw)
                except etree.XMLSyntaxError as e:
                    errs.append({
                        "part": name,
                        "error": str(e),
                        "position": getattr(e, 'position', None)
                    })
    return errs

def main():
    backup_if_needed()
    print("Processing and escaping angle brackets inside <w:t> and similar text nodes...")
    modified = build_fixed_docx(DOCX_PATH, OUTPATH)
    if not modified:
        print("No text nodes required escaping (no changes). Wrote copy at:", OUTPATH)
    else:
        print("Wrote fixed candidate to:", OUTPATH)
    print("Validating fixed package...")
    errors = validate_docx(OUTPATH)
    if not errors:
        print("Validation OK. Try opening the file in Word (Open and Repair if prompted):", OUTPATH)
    else:
        print("Validation still finds XML errors. First few errors:")
        for e in errors[:6]:
            print("Part:", e['part'])
            print("Error:", e['error'])
            print("Position:", e['position'])
            print("-" * 60)
        print("If errors persist, paste the first part+error here and I'll analyze further.")

if __name__ == '__main__':
    main()
