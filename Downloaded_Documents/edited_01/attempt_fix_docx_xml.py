# attempt_fix_docx_xml.py
import zipfile, re, os, shutil

DOCX_PATH = r"C:\Users\K Santosh Kumar\Desktop\HEALTHARK\04_thermofisher\Downloaded_Documents\edited\refernce_template_unlocked.docx"
OUT_PATH = os.path.splitext(DOCX_PATH)[0] + "_fixed.docx"
BACKUP_PATH = os.path.splitext(DOCX_PATH)[0] + "_bak.docx"

def safe_backup(src, bak):
    if not os.path.exists(bak):
        shutil.copy2(src, bak)
        print("Backup created:", bak)
    else:
        print("Backup already exists:", bak)

def escape_ampersands(xml):
    # Replace ampersands that are not part of entities (e.g. &amp; &lt; &#123;)
    # We use a regex negative lookahead for valid entity pattern: &[A-Za-z0-9#]+;
    pattern = re.compile(r'&(?![A-Za-z0-9#]+;)')
    new_xml, n = pattern.subn('&amp;', xml)
    return new_xml, n

def remove_illegal_chars(xml):
    # Remove ASCII control chars except tab(9), LF(10), CR(13)
    # Valid XML 1.0 chars: #x9 | #xA | #xD | [#x20-#xD7FF] | ...
    # We'll strip characters with ord < 32 except 9,10,13
    out = []
    removed = 0
    for ch in xml:
        o = ord(ch)
        if o == 9 or o == 10 or o == 13 or o >= 32:
            out.append(ch)
        else:
            removed += 1
    return ''.join(out), removed

def fix_document_xml_bytes(data_bytes):
    try:
        xml = data_bytes.decode('utf-8')
    except Exception:
        xml = data_bytes.decode('utf-8', errors='replace')
    xml2, amp_changes = escape_ampersands(xml)
    xml3, removed_cnt = remove_illegal_chars(xml2)
    print(f"escape_ampersands made {amp_changes} replacements; removed {removed_cnt} illegal control chars.")
    return xml3.encode('utf-8')

def main():
    if not os.path.exists(DOCX_PATH):
        print("DOCX not found:", DOCX_PATH); return
    safe_backup(DOCX_PATH, BACKUP_PATH)

    with zipfile.ZipFile(DOCX_PATH, 'r') as zin:
        namelist = zin.namelist()
        if 'word/document.xml' not in namelist:
            print("No word/document.xml found; aborting."); return
        orig_doc_xml = zin.read('word/document.xml')

        fixed_doc_xml = fix_document_xml_bytes(orig_doc_xml)

        # create new zip with replaced file
        with zipfile.ZipFile(OUT_PATH, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in namelist:
                if item == 'word/document.xml':
                    zout.writestr(item, fixed_doc_xml)
                else:
                    zout.writestr(item, zin.read(item))
    print("Wrote fixed docx to:", OUT_PATH)
    print("Try opening the fixed file in Word (Open and Repair if needed).")

if __name__ == "__main__":
    main()
