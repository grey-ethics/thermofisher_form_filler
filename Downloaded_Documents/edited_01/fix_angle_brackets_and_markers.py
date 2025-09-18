# fix_angle_brackets_and_markers.py
import zipfile, re, os, shutil

# << EDIT THIS >>
DOCX_PATH = r"C:\Users\K Santosh Kumar\Desktop\HEALTHARK\04_thermofisher\Downloaded_Documents\edited\refernce_template_unlocked.docx"
# --------------------------------

OUT_PATH = os.path.splitext(DOCX_PATH)[0] + "_angles_fixed.docx"
BACKUP_PATH = os.path.splitext(DOCX_PATH)[0] + "_bak.docx"

def safe_backup(src, bak):
    if not os.path.exists(bak):
        shutil.copy2(src, bak)
        print("Backup created:", bak)
    else:
        print("Backup already exists:", bak)

# pattern to find all <w:t ...>...</w:t> (including tags with attributes)
WT_RE = re.compile(r'(<w:t[^>]*>)(.*?)(</w:t>)', flags=re.DOTALL)

def fix_wt_text(xml_text):
    changed_total = 0

    def repl(m):
        nonlocal changed_total
        open_tag, inner, close_tag = m.group(1), m.group(2), m.group(3)
        orig_inner = inner

        # 1) escape lone '<' and '>' inside text (leave existing entities like &lt; intact)
        # We do this by replacing literal < and > characters
        # (we avoid touching &xxx; sequences)
        # Because inner is text content, any '<' or '>' present are invalid and must be escaped
        inner2 = inner.replace('<', '&lt;').replace('>', '&gt;')

        # 2) optionally replace the marker &lt;&lt;CHK&gt;&gt; (or literal <<CHK>>) with checkbox glyph
        # handle both cases (in case markers were already escaped or not)
        if '&lt;&lt;CHK&gt;&gt;' in inner2:
            inner2 = inner2.replace('&lt;&lt;CHK&gt;&gt;', '\u2610')  # ☐
        if '<<CHK>>' in inner2:
            inner2 = inner2.replace('<<CHK>>', '\u2610')

        if inner2 != orig_inner:
            changed_total += 1
        return open_tag + inner2 + close_tag

    new_xml, n = WT_RE.subn(repl, xml_text)
    return new_xml, changed_total

def process_docx(in_path, out_path):
    with zipfile.ZipFile(in_path, 'r') as zin:
        names = zin.namelist()
        if 'word/document.xml' not in names:
            raise RuntimeError("No word/document.xml in docx")
        fixed_parts = {}
        # process all word/*.xml parts (document + headers/footers/footnotes/endnotes)
        for part in names:
            if part.startswith('word/') and part.endswith('.xml'):
                data = zin.read(part)
                try:
                    txt = data.decode('utf-8')
                except Exception:
                    txt = data.decode('utf-8', errors='replace')
                new_txt, changed = fix_wt_text(txt)
                if changed:
                    fixed_parts[part] = new_txt.encode('utf-8')
                    print(f"Fixed {changed} <w:t> entries in {part}")
            # else: leave as-is

        # write a new zip copying unchanged files, replacing modified parts
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in names:
                if item in fixed_parts:
                    zout.writestr(item, fixed_parts[item])
                else:
                    zout.writestr(item, zin.read(item))

def main():
    if not os.path.exists(DOCX_PATH):
        print("DOCX not found:", DOCX_PATH); return
    safe_backup(DOCX_PATH, BACKUP_PATH)
    process_docx(DOCX_PATH, OUT_PATH)
    print("Wrote fixed docx to:", OUT_PATH)
    print("Open the fixed file in Word (try Open and Repair if Word warns).")

if __name__ == "__main__":
    main()
