# services/word_com_replace.py
import os
import pythoncom
import win32com.client as win32

WD_GO_TO_PAGE = 1
WD_GO_TO_ABSOLUTE = 1
WD_FORMAT_DOCX = 16   # wdFormatXMLDocument
WD_FORMAT_PDF  = 17   # wdFormatPDF

def _start_word():
    pythoncom.CoInitialize()
    app = win32.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    return app

def _stop_word(app):
    try:
        app.Quit()
    finally:
        pythoncom.CoUninitialize()

def replace_docx_page3_with_file(target_docx: str, page3_docx: str, out_docx: str):
    """
    Open target_docx, copy entire content of page3_docx (1 page),
    select page 3 in target, paste, save to out_docx.
    """
    app = _start_word()
    try:
        tgt = app.Documents.Open(os.path.abspath(target_docx))
        src = app.Documents.Open(os.path.abspath(page3_docx))

        # copy all content from the 1-page page
        src.Content.WholeStory()
        src.Content.Copy()

        sel = app.Selection
        # Go to page 3
        sel.GoTo(What=WD_GO_TO_PAGE, Which=WD_GO_TO_ABSOLUTE, Count=3)
        # Select only that page's range
        sel.Bookmarks("\\Page").Range.Select()
        # Paste
        sel.Range.Paste()

        # Save as DOCX
        tgt.SaveAs2(os.path.abspath(out_docx), FileFormat=WD_FORMAT_DOCX)

        src.Close(False)
        tgt.Close(False)
    finally:
        _stop_word(app)

def docx_to_pdf(in_docx: str, out_pdf: str):
    """
    Export a DOCX to PDF using Word's fixed-format export.
    """
    app = _start_word()
    try:
        doc = app.Documents.Open(os.path.abspath(in_docx))
        # ExportAsFixedFormat(OutputFileName, ExportFormat=17 (wdExportFormatPDF))
        doc.ExportAsFixedFormat(os.path.abspath(out_pdf), WD_FORMAT_PDF)
        doc.Close(False)
    finally:
        _stop_word(app)
