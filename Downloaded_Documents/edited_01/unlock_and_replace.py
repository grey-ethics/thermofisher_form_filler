import zipfile
import io
import re
from pathlib import Path

INPUT = Path("refernce_template.docx")   # change if needed
OUTPUT = Path("refernce_template_unlocked.docx")

DOC_XML = "word/document.xml"

def read_docx_xml(z, name):
    try:
        data = z.read(name)
        return data.decode("utf-8")
    except KeyError:
        return None

def write_docx_xml(z_out, name, xml_text):
    z_out.writestr(name, xml_text.encode("utf-8"))

def process_docx(in_path: Path, out_path: Path):
    # read input zip
    with zipfile.ZipFile(in_path, 'r') as zin:
        # read document.xml
        doc_xml = read_docx_xml(zin, DOC_XML)
        if doc_xml is None:
            raise RuntimeError("document.xml not found in docx")

        # 1) remove any <w:documentProtection .../> element
        #    (handles both empty-element and start/end pair)
        doc_xml_new = re.sub(r'<w:documentProtection\b[^/>]*(?:/>|>.*?</w:documentProtection>)', '', doc_xml, flags=re.DOTALL)

        # 2) replace all checkbox glyphs with token <<CHK>>
        doc_xml_new = doc_xml_new.replace("☐", "<<CHK>>")

        # copy everything into new zip, but use modified document.xml
        with zipfile.ZipFile(out_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == DOC_XML:
                    # write modified document.xml
                    write_docx_xml(zout, DOC_XML, doc_xml_new)
                else:
                    # copy as-is
                    zout.writestr(item, zin.read(item.filename))

if __name__ == "__main__":
    if not INPUT.exists():
        print("Input not found:", INPUT)
    else:
        process_docx(INPUT, OUTPUT)
        print("Wrote", OUTPUT)
