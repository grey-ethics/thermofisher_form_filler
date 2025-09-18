import aspose.words as aw

def docx_to_pdf_with_forms(src_docx: str, out_pdf: str):
    doc = aw.Document(src_docx)
    options = aw.saving.PdfSaveOptions()
    # Critical flag: tells Aspose to map form-ish things to AcroForm fields where possible
    options.preserve_form_fields = True
    doc.save(out_pdf, options)

if __name__ == "__main__":
    docx_to_pdf_with_forms(
        "refernce_template_tagged.docx",   # or your latest source DOCX
        "form_preview.pdf"
    )
    print("Wrote form_preview.pdf")
