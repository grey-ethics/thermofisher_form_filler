# validate_and_fix_docx.py
# Run this inside your venv (where lxml is available).
#
# Usage:
#   python validate_and_fix_docx.py
#
# It will create a backup DOCX and write a fixed file with suffix _fixed.docx.
#
import zipfile, os, shutil, re, sys
from lxml import etree

# --- EDIT this to point at your problematic docx file ---
DOCX_PATH = r"C:\Users\K Santosh Kumar\Desktop\HEALTHARK\04_thermofisher\Downloaded_Documents\edited\refernce_template_unlocked_angles_fixed.docx"
# ----------------------------------------------------------------

if not os.path.exists(DOCX_PATH):
    print("DOCX not found:", DOCX_PATH); sys.exit(1)

BASE = os.path.splitext(DOCX_PATH)[0]
BACKUP = BASE + "_repair_bak.docx"
OUTPATH = BASE + "_fixed.docx"
TMPDIR = BASE + "_tmp"

def safe_backup():
    if not os.path.exists(BACKUP):
        shutil.copy2(DOCX_PATH, BACKUP)
        print("Backup created:", BACKUP)
    else:
        print("Backup already exists:", BACKUP)

# remove illegal XML control chars
ILLEGAL_RE = re.compile(
    r'[\x00-\x08\x0B\x0C\x0E-\x1F]'
)

# We know the original problem had literal <<CHK>> inside <w:t>:
# replace both literal and escaped forms with a checkbox glyph
MARKER_REPLACEMENTS = [
    (b'<<CHK>>', '\u2610'),          # literal
    (b'&lt;&lt;CHK&gt;&gt;', '\u2610') # already-escaped
]

def process_parts(in_path, out_path):
    modified = {}
    with zipfile.ZipFile(in_path, 'r') as zin:
        names = zin.namelist()
        # iterate and pre-fix parts that are XML (word/*.xml)
        for name in names:
            data = zin.read(name)
            if name.startswith('word/') and name.endswith('.xml'):
                try:
                    txt = data.decode('utf-8')
                except Exception:
                    # try latin-1 fallback but note this is rare for Word XML
                    txt = data.decode('utf-8', errors='replace')

                orig_txt = txt

                # 1) remove illegal control chars
                txt2 = ILLEGAL_RE.sub('', txt)

                # 2) replace marker forms with checkbox glyph (operate on the decoded text)
                # handle both literal and already-escaped
                txt2 = txt2.replace('<<CHK>>', '\u2610')
                txt2 = txt2.replace('&lt;&lt;CHK&gt;&gt;', '\u2610')

                if txt2 != orig_txt:
                    modified[name] = txt2.encode('utf-8')
                    print(f"Prepared fixes for: {name} (changed)")
                else:
                    # keep original bytes if unchanged
                    modified[name] = None

            else:
                # non-word xml or binary part: keep as-is
                modified[name] = None

        # Build a candidate fixed zip in memory (write changed bytes or original bytes)
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                if modified.get(name) is not None:
                    zout.writestr(name, modified[name])
                else:
                    zout.writestr(name, zin.read(name))

    print("Wrote candidate fixed docx to:", out_path)

def validate_docx(path):
    errors = []
    with zipfile.ZipFile(path, 'r') as z:
        for name in z.namelist():
            if name.endswith('.xml'):
                raw = z.read(name)
                try:
                    # try parse
                    etree.fromstring(raw)
                except etree.XMLSyntaxError as e:
                    # collect error info
                    msg = str(e)
                    # try to show snippet around error line/col
                    ln = getattr(e, 'position', (None, None))[0]
                    col = getattr(e, 'position', (None, None))[1]
                    snippet = None
                    try:
                        text = raw.decode('utf-8', errors='replace')
                        if ln is not None and col is not None:
                            lines = text.splitlines()
                            idx = max(0, ln-2)
                            snippet = "\n".join(f"{i+1:5d}: {lines[i]}" for i in range(idx, min(len(lines), ln+1)))
                        else:
                            snippet = text[:2000]
                    except Exception:
                        snippet = "<could not decode snippet>"
                    errors.append({
                        "part": name,
                        "error": msg,
                        "line": ln,
                        "col": col,
                        "snippet": snippet
                    })
    return errors

def main():
    safe_backup()
    process_parts(DOCX_PATH, OUTPATH)

    print("Validating fixed package...")
    errs = validate_docx(OUTPATH)
    if not errs:
        print("Validation passed: no XML syntax errors detected in any .xml parts.")
        print("Try opening", OUTPATH, "in Word (use Open and Repair if Word prompts).")
    else:
        print("Validation found XML errors in the fixed package. Details:")
        for e in errs:
            print("-" * 80)
            print("Part:", e['part'])
            print("Error:", e['error'])
            print("Line:", e['line'], "Col:", e['col'])
            print("Snippet around error (if available):")
            print(e['snippet'][:2000])
            print("-" * 80)
        print("If errors persist, copy & paste the first part+error snippet here and I'll analyze further.")

if __name__ == "__main__":
    main()
