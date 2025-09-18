# services/pdf_replace.py
from pypdf import PdfReader, PdfWriter

def replace_pdf_page(src_pdf_path: str, replacement_page_pdf_path: str, out_pdf_path: str, replace_index: int = 2):
    """
    Replace page at 'replace_index' (0-based) in src_pdf with the single page from replacement_page_pdf_path.
    """
    src = PdfReader(src_pdf_path)
    rep = PdfReader(replacement_page_pdf_path)
    rep_page = rep.pages[0]

    writer = PdfWriter()
    total = len(src.pages)

    for i in range(total):
        if i == replace_index:
            writer.add_page(rep_page)
        else:
            writer.add_page(src.pages[i])

    with open(out_pdf_path, "wb") as f:
        writer.write(f)
